// Copied from VBASA/VbaTestHost/VbeErrorRecord.cpp for the AccessMCP VBE add-in.
// Keep the x86 VBE7 offsets in sync with that source. AccessMCP adds a best-effort
// read of the record's message pointer at +0x08.
#include "stdafx.h"
#include "VbeErrorRecord.h"

#include <cstdio>
#include <cstring>
#include <cwctype>

namespace vt {

namespace {

#if defined(_M_IX86)

// --- x86 VBE7 offsets ------------------------------------------------------
// VBAErr (the object rtcErrObj returns) + 0x18 holds _g_EbThread.
constexpr size_t kErrObjToEbThread = 0x18;
// BTSRC_ERR lives inside the _g_EbThread structure.
constexpr size_t kEbThreadToBtsrcErr = 0x17C;
// BTSRC_ERR (24 bytes): module root, flags, number, message, line/column.
constexpr size_t kRecordRoot = 0x00;
constexpr size_t kRecordFlags = 0x04;
constexpr size_t kRecordNumber = 0x06;
constexpr size_t kRecordLine = 0x0C;
constexpr size_t kRecordColumn = 0x0E;
constexpr unsigned kFlagPresent = 0x0002;
constexpr unsigned kFlagColumns = 0x0004;
// Object Info (epiModule) -> Object Table and Public Object Descriptor.
constexpr size_t kObjectInfoObjectTable = 0x04;
constexpr size_t kObjectInfoPublicObject = 0x18;
// Public Object Descriptor (epiModuleDesc) -> object name.
constexpr size_t kPublicObjectName = 0x18;
// Object Table -> object array and project name.
constexpr size_t kObjectTableArray = 0x30;
constexpr size_t kObjectTableProjectName = 0x40;

bool readable(const void* address, size_t size) {
    if (!address || size == 0) {
        return false;
    }
    MEMORY_BASIC_INFORMATION info = { 0 };
    if (VirtualQuery(address, &info, sizeof(info)) == 0) {
        return false;
    }
    if (info.State != MEM_COMMIT) {
        return false;
    }
    DWORD protect = info.Protect & 0xFF;
    switch (protect) {
        case PAGE_READONLY:
        case PAGE_READWRITE:
        case PAGE_WRITECOPY:
        case PAGE_EXECUTE_READ:
        case PAGE_EXECUTE_READWRITE:
        case PAGE_EXECUTE_WRITECOPY:
            break;
        default:
            return false;
    }
    if (info.Protect & (PAGE_NOACCESS | PAGE_GUARD)) {
        return false;
    }
    const uint8_t* start = static_cast<const uint8_t*>(address);
    const uint8_t* regionEnd = static_cast<const uint8_t*>(info.BaseAddress) + info.RegionSize;
    return start + size <= regionEnd;
}

template <typename T>
bool read_at(const void* base, size_t offset, T& out) {
    if (!base) {
        return false;
    }
    const uint8_t* address = static_cast<const uint8_t*>(base) + offset;
    if (!readable(address, sizeof(T))) {
        return false;
    }
    memcpy(&out, address, sizeof(T));
    return true;
}

bool read_pointer(const void* base, size_t offset, void*& out) {
    return read_at<void*>(base, offset, out);
}

// An ANSI C string in the VBE's structures; bounded so a bad pointer cannot
// walk the address space.
std::wstring read_ansi(const void* address) {
    std::wstring text;
    if (!address) {
        return text;
    }
    const uint8_t* cursor = static_cast<const uint8_t*>(address);
    for (int i = 0; i < 256; ++i) {
        if (!readable(cursor + i, 1)) {
            return std::wstring();
        }
        uint8_t ch = cursor[i];
        if (ch == 0) {
            return text;
        }
        if (ch < 0x20 || ch > 0x7E) {
            return std::wstring();
        }
        text.push_back(static_cast<wchar_t>(ch));
    }
    return std::wstring();
}

// A BSTR at address (length prefix four bytes before the data pointer).
std::wstring read_bstr(const void* address) {
    std::wstring text;
    if (!address) {
        return text;
    }
    uint32_t byteLength = 0;
    if (!read_at<uint32_t>(static_cast<const uint8_t*>(address) - 4, 0, byteLength)) {
        return text;
    }
    if (byteLength == 0 || byteLength > 2048 || (byteLength & 1) != 0) {
        return text;
    }
    for (uint32_t i = 0; i < byteLength / 2; ++i) {
        wchar_t ch = 0;
        if (!read_at<wchar_t>(address, static_cast<size_t>(i) * sizeof(wchar_t), ch)) {
            return std::wstring();
        }
        if (ch == 0) {
            break;
        }
        if (ch < 0x20) {
            return std::wstring();
        }
        text.push_back(ch);
    }
    return text;
}

bool looks_like_identifier(const std::wstring& text) {
    if (text.empty() || text.size() > 128) {
        return false;
    }
    for (wchar_t ch : text) {
        if (!(iswalnum(ch) || ch == L'_' || ch == L' ')) {
            return false;
        }
    }
    return true;
}

// Ask a COM object for its Name property. No C++ objects here, so SEH is legal.
bool dispatch_name_raw(void* object, BSTR* out) {
    *out = nullptr;
    if (!object) {
        return false;
    }
    __try {
        IUnknown* unknown = reinterpret_cast<IUnknown*>(object);
        IDispatch* dispatch = nullptr;
        if (FAILED(unknown->QueryInterface(IID_IDispatch, reinterpret_cast<void**>(&dispatch))) || !dispatch) {
            return false;
        }
        LPOLESTR name = const_cast<LPOLESTR>(L"Name");
        DISPID dispid = 0;
        bool ok = false;
        if (SUCCEEDED(dispatch->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid))) {
            DISPPARAMS params = { nullptr, nullptr, 0, 0 };
            VARIANT value;
            VariantInit(&value);
            if (SUCCEEDED(dispatch->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                    DISPATCH_PROPERTYGET, &params, &value, nullptr, nullptr))) {
                if (value.vt == VT_BSTR && value.bstrVal != nullptr) {
                    *out = SysAllocString(value.bstrVal);
                    ok = (*out != nullptr);
                }
                VariantClear(&value);
            }
        }
        dispatch->Release();
        return ok;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

std::wstring dispatch_name(void* object) {
    BSTR value = nullptr;
    if (!dispatch_name_raw(object, &value) || value == nullptr) {
        return std::wstring();
    }
    std::wstring text(value);
    SysFreeString(value);
    return text;
}

// Replicates the main branch of VBE7's _TipGetModuleName: the module's
// BASIC_TYPEINFO interface (the object GetCompileErrorInfo returns) has the
// name getter at [[object+0x08]+0x10], called with this = object+0x08.
bool tip_module_name_raw(void* moduleObject, BSTR* out) {
    *out = nullptr;
    if (!moduleObject) {
        return false;
    }
    __try {
        void* vtable = nullptr;
        if (!read_pointer(moduleObject, 0x08, vtable) || !readable(vtable, 0x14)) {
            return false;
        }
        void* function = nullptr;
        if (!read_pointer(vtable, 0x10, function) || !function) {
            return false;
        }
        typedef unsigned (__stdcall * NameFn)(void*, wchar_t**);
        wchar_t* name = nullptr;
        unsigned hr = reinterpret_cast<NameFn>(function)(
            reinterpret_cast<uint8_t*>(moduleObject) + 0x08, &name);
        if (hr != 0 || name == nullptr) {
            return false;
        }
        *out = SysAllocString(name);
        SysFreeString(name);
        return *out != nullptr;
    } __except (EXCEPTION_EXECUTE_HANDLER) {
        return false;
    }
}

std::wstring tip_module_name(void* moduleObject) {
    BSTR value = nullptr;
    if (!tip_module_name_raw(moduleObject, &value) || value == nullptr) {
        return std::wstring();
    }
    std::wstring text(value);
    SysFreeString(value);
    return text;
}

std::wstring hex_value(size_t value) {
    wchar_t buffer[32];
    swprintf_s(buffer, L"0x%llx", static_cast<unsigned long long>(value));
    return buffer;
}

// Public Object Descriptor.lpszObjectName, one hop from the Object Info.
std::wstring object_info_module_name(void* objectInfo) {
    void* publicObject = nullptr;
    if (!read_pointer(objectInfo, kObjectInfoPublicObject, publicObject)) {
        return std::wstring();
    }
    void* name = nullptr;
    if (!read_pointer(publicObject, kPublicObjectName, name)) {
        return std::wstring();
    }
    return read_ansi(name);
}

// Object Table.lpszProjectName.
std::wstring object_table_project_name(void* objectTable) {
    void* name = nullptr;
    if (!read_pointer(objectTable, kObjectTableProjectName, name)) {
        return std::wstring();
    }
    return read_ansi(name);
}

void dump_dwords(VbeCompileError& out, const wchar_t* label, void* base, int count) {
    if (!base) {
        return;
    }
    std::wstring line = std::wstring(label) + L":";
    for (int i = 0; i < count; ++i) {
        uint32_t value = 0;
        if (!read_at<uint32_t>(base, static_cast<size_t>(i) * sizeof(uint32_t), value)) {
            break;
        }
        wchar_t buffer[16];
        swprintf_s(buffer, L" %08x", value);
        line += buffer;
    }
    out.probe.push_back(line);
}

void probe_object_info(VbeCompileError& out, const wchar_t* label, void* candidate) {    if (!candidate) {
        return;
    }
    std::wstring name = object_info_module_name(candidate);
    out.probe.push_back(std::wstring(label) + L"=" + hex_value(reinterpret_cast<size_t>(candidate))
        + L" -> module=\"" + name + L"\"");
    if (out.module.empty() && looks_like_identifier(name)) {
        out.module = name;
    }
}

// A bounded walk of the pointer graph from the error record's root, looking for
// a node that behaves like an Object Info (name at [[node+0x18]+0x18]) or a
// Public Object Descriptor (name at [node+0x18]). Used only while probing, to
// pin the record -> module link for this VBE7 build.
void probe_pointer_graph(VbeCompileError& out, void* start) {
    if (!start) {
        return;
    }
    std::vector<void*> queue;
    std::vector<int> depths;
    std::vector<std::wstring> paths;
    std::vector<void*> visited;
    queue.push_back(start);
    depths.push_back(0);
    paths.push_back(L"root");
    visited.push_back(start);

    size_t head = 0;
    int budget = 4000;
    int reported = 0;
    auto seen = [&visited](void* p) {
        for (void* v : visited) {
            if (v == p) {
                return true;
            }
        }
        return false;
    };
    auto plausible_com = [](void* node) {
        void* vtable = nullptr;
        if (!read_pointer(node, 0, vtable)) {
            return false;
        }
        size_t address = reinterpret_cast<size_t>(vtable);
        return address > 0x10000 && readable(vtable, 4 * sizeof(void*));
    };
    while (head < queue.size() && budget-- > 0) {
        void* node = queue[head];
        int depth = depths[head];
        std::wstring path = paths[head];
        ++head;

        // A COM object we can reach: ask it for Name (the VBComponent name).
        if (plausible_com(node)) {
            std::wstring name = dispatch_name(node);
            if (looks_like_identifier(name) && reported < 60) {
                out.probe.push_back(L"graph[COM] " + path + L" -> \"" + name + L"\"");
                ++reported;
                if (out.module.empty()) {
                    out.module = name;
                }
            }
        }
        // Object Info: [node+0x18] -> POD, then POD.lpszObjectName.
        {
            std::wstring name = object_info_module_name(node);
            if (looks_like_identifier(name) && reported < 60) {
                out.probe.push_back(L"graph[OI] " + path + L" -> \"" + name + L"\"");
                ++reported;
                if (out.module.empty()) {
                    out.module = name;
                }
            }
        }
        // Public Object Descriptor directly.
        {
            void* namePtr = nullptr;
            if (read_pointer(node, kPublicObjectName, namePtr)) {
                std::wstring name = read_ansi(namePtr);
                if (looks_like_identifier(name) && reported < 60) {
                    out.probe.push_back(L"graph[POD] " + path + L" -> \"" + name + L"\"");
                    ++reported;
                }
            }
        }
        if (depth >= 3) {
            continue;
        }
        for (size_t offset = 0; offset <= 0x40; offset += sizeof(void*)) {
            void* child = nullptr;
            if (!read_pointer(node, offset, child) || !child || seen(child)) {
                continue;
            }
            visited.push_back(child);
            queue.push_back(child);
            depths.push_back(depth + 1);
            wchar_t step[24];
            swprintf_s(step, L"+0x%x", static_cast<unsigned>(offset));
            paths.push_back(path + step);
        }
    }
    out.probe.push_back(L"graph: visited " + std::to_wstring(visited.size()) + L" nodes");
}

#endif // _M_IX86

} // namespace

VbeCompileError read_compile_error_record() {
    VbeCompileError out;
#if !defined(_M_IX86)
    return out; // x64 offsets are not mapped yet; fall back to the description
#else
    HMODULE vbe = GetModuleHandleW(L"VBE7.DLL");
    if (!vbe) {
        out.probe.push_back(L"VBE7.DLL is not loaded");
        return out;
    }
    typedef void* (__stdcall * RtcErrObjFn)();
    RtcErrObjFn rtcErrObj = reinterpret_cast<RtcErrObjFn>(GetProcAddress(vbe, "rtcErrObj"));
    if (!rtcErrObj) {
        out.probe.push_back(L"VBE7.DLL does not export rtcErrObj");
        return out;
    }

    void* errObj = rtcErrObj();
    out.probe.push_back(L"Err=" + hex_value(reinterpret_cast<size_t>(errObj)));
    void* ebThread = nullptr;
    if (!read_pointer(errObj, kErrObjToEbThread, ebThread)) {
        out.probe.push_back(L"the Err object's _g_EbThread pointer is not readable");
        return out;
    }
    out.probe.push_back(L"_g_EbThread=" + hex_value(reinterpret_cast<size_t>(ebThread)));

    const uint8_t* record = static_cast<const uint8_t*>(ebThread) + kEbThreadToBtsrcErr;
    uint16_t flags = 0;
    if (!read_at<uint16_t>(record, kRecordFlags, flags)) {
        out.probe.push_back(L"BTSRC_ERR is not readable");
        return out;
    }
    if (!(flags & kFlagPresent)) {
        out.probe.push_back(L"no compile error is recorded");
        return out;
    }
    out.present = true;

    uint16_t number = 0;
    uint16_t lineUnits = 0;
    uint16_t column = 0;
    read_at<uint16_t>(record, kRecordNumber, number);
    read_at<uint16_t>(record, kRecordLine, lineUnits);
    if ((flags & kFlagColumns) && read_at<uint16_t>(record, kRecordColumn, column)) {
        out.column = column;
    }
    out.errorNumber = number;
    out.lineUnits = lineUnits;
    // The record holds the compiler's line counter. Measured on this VBE7
    // build, it advances one unit per physical source line (line N sits at
    // N-1), so the source line is value + 1. A line continuation makes the
    // compiler count logical lines and can shift it.
    out.line = static_cast<unsigned>(lineUnits) + 1;
    {
        wchar_t buffer[160];
        swprintf_s(buffer, L"flags=0x%x number=0x%x lineUnits=%u line=%u column=%u",
            flags, number, static_cast<unsigned>(lineUnits), out.line, static_cast<unsigned>(column));
        out.probe.push_back(buffer);
    }

    // BTSRC_ERR + 0x08 is the compiler's message pointer. Depending on the VBE
    // build it is either an ANSI string or a BSTR; accept whichever reads as a
    // longer, printable string.
    {
        void* messagePtr = nullptr;
        if (read_pointer(record, 0x08, messagePtr) && messagePtr) {
            std::wstring ansi = read_ansi(messagePtr);
            std::wstring bstr = read_bstr(messagePtr);
            out.message = bstr.size() > ansi.size() ? bstr : ansi;
            if (!out.message.empty()) {
                wchar_t buffer[512];
                swprintf_s(buffer, L"message=\"%s\"", out.message.c_str());
                out.probe.push_back(buffer);
            }
        }
    }

    // The compiler records the module's type root here. Walk every plausible
    // link to an Object Info and report all of them; the first that yields a
    // printable identifier becomes the answer.
    void* root = nullptr;
    read_pointer(record, kRecordRoot, root);
    out.probe.push_back(L"root=" + hex_value(reinterpret_cast<size_t>(root)));

    probe_object_info(out, L"root", root);

    void* docProvider = nullptr;
    if (read_pointer(root, 0x24, docProvider)) {
        probe_object_info(out, L"[root+0x24]", docProvider);
    }
    void* cvbaExtModule = nullptr;
    if (read_pointer(root, 0x18, cvbaExtModule)) {
        probe_object_info(out, L"[root+0x18]", cvbaExtModule);
        void* inner = nullptr;
        if (read_pointer(cvbaExtModule, 0x04, inner)) {
            probe_object_info(out, L"[[root+0x18]+0x04]", inner);
        }
    }

    // Object Table routes: the root may carry the table at +0x04, or be one.
    void* objectTable = nullptr;
    if (read_pointer(root, 0x04, objectTable)) {
        probe_object_info(out, L"[root+0x04]-as-objectinfo", objectTable);
        std::wstring project = object_table_project_name(objectTable);
        out.probe.push_back(L"[root+0x04] -> project=\"" + project + L"\"");
        if (out.project.empty() && looks_like_identifier(project)) {
            out.project = project;
        }
    }
    {
        std::wstring project = object_table_project_name(root);
        out.probe.push_back(L"root-as-table -> project=\"" + project + L"\"");
        if (out.project.empty() && looks_like_identifier(project)) {
            out.project = project;
        }
    }

    {
        void* table = objectTable ? objectTable : root;
        void* array = nullptr;
        if (read_pointer(table, kObjectTableArray, array)) {
            out.probe.push_back(L"objectArray=" + hex_value(reinterpret_cast<size_t>(array)));
            for (int i = 0; i < 8; ++i) {
                void* descriptor = nullptr;
                if (!read_pointer(array, static_cast<size_t>(i) * sizeof(void*), descriptor) || !descriptor) {
                    continue;
                }
                void* name = nullptr;
                read_pointer(descriptor, kPublicObjectName, name);
                out.probe.push_back(L"array[" + std::to_wstring(i) + L"]="
                    + hex_value(reinterpret_cast<size_t>(descriptor)) + L" name=\"" + read_ansi(name) + L"\"");
            }
        }
    }

    // The type root points at the module's extensibility object (CVbaExtModule)
    // at +0x18; its Name is the module name. This is the supported COM route,
    // and it is what makes the walk resilient to VBE layout changes.
    void* extensibilityModule = nullptr;
    if (read_pointer(root, 0x18, extensibilityModule)) {
        std::wstring name = dispatch_name(extensibilityModule);
        out.probe.push_back(L"[root+0x18] IDispatch(Name)=\"" + name + L"\"");
        if (out.module.empty() && looks_like_identifier(name)) {
            out.module = name;
        }
    }
    void* docProviderCom = nullptr;
    if (read_pointer(root, 0x24, docProviderCom)) {
        std::wstring name = dispatch_name(docProviderCom);
        out.probe.push_back(L"[root+0x24] IDispatch(Name)=\"" + name + L"\"");
        if (out.module.empty() && looks_like_identifier(name)) {
            out.module = name;
        }
        std::wstring tipName = tip_module_name(docProviderCom);
        out.probe.push_back(L"[root+0x24] TipGetModuleName=\"" + tipName + L"\"");
        if (out.module.empty() && looks_like_identifier(tipName)) {
            out.module = tipName;
        }
    }

    bool probing = false;
    {
        wchar_t value[8] = { 0 };
        probing = GetEnvironmentVariableW(L"VBATEST_PROBE", value, _countof(value)) > 0;
    }
    if (probing) {
        dump_dwords(out, L"record", const_cast<uint8_t*>(record), 6);
        dump_dwords(out, L"root", root, 28);
        void* p18 = nullptr;
        if (read_pointer(root, 0x18, p18)) {
            dump_dwords(out, L"[root+0x18]", p18, 16);
        }
        void* p24 = nullptr;
        if (read_pointer(root, 0x24, p24)) {
            dump_dwords(out, L"[root+0x24]", p24, 16);
        }
        void* p28 = nullptr;
        if (read_pointer(root, 0x28, p28)) {
            dump_dwords(out, L"[root+0x28]", p28, 16);
        }
        probe_pointer_graph(out, root);
    }
    return out;
#endif
}

} // namespace vt
