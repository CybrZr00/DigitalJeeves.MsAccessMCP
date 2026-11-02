// AccessMCP VBE add-in.
//
// A minimal x86 in-process VBE add-in. It exists for one reason: the VBE's
// compiler interface (IVbaProject) is not exposed through the out-of-process
// Access automation proxy, but it is available in-process to a VBE add-in.
//
// The add-in is loaded with Application.VBE.AddIns("AccessMCP.VbeProbe").Object
// and exposes CompileProject(), which calls IVbaProject::Compile() and reads
// the VBE's BTSRC_ERR compile-error record (VbeErrorRecord.cpp, shared with
// VbaTestHost) before anything clears it.
//
// The supported method is ReadErrorRecord(). AccessMCP calls DoCmd.RunCommand(126)
// and then reads the record; that path returns the exact module and line.
// CompileProject()/IVBEProject are experimental and unused: the project type
// library only sees saved modules, so it can return S_OK for an unsaved broken
// module.

#include "stdafx.h"
#include "VbeErrorRecord.h"

#include <new>
#include <string>

namespace {

const IID IID_IVbaProject =
    { 0xCD9B1AC5, 0xFAA3, 0x11D1, { 0x99, 0x22, 0x00, 0xA0, 0xC9, 0x70, 0x24, 0x42 } };

const IID IID_IVba =
    { 0xCD9B1AC8, 0xFAA3, 0x11D1, { 0x99, 0x22, 0x00, 0xA0, 0xC9, 0x70, 0x24, 0x42 } };

const IID IID_IVBEProject =
    { 0xDDD557E0, 0xD96F, 0x11CD, { 0x95, 0x70, 0x00, 0xAA, 0x00, 0x51, 0xE5, 0xD4 } };

const IID IID_ITypeLibStd =
    { 0x00020402, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

const IID IID_ITypeInfoStd =
    { 0x00020401, 0x0000, 0x0000, { 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46 } };

const IID IID_IDTExtensibility2 =
    { 0xB65AD801, 0xABAF, 0x11D0, { 0xBB, 0x8B, 0x00, 0xA0, 0xC9, 0x0F, 0x27, 0x44 } };

const CLSID CLSID_AccessMcpVbeProbe =
    { 0xEFAA3F97, 0x4E2F, 0x410E, { 0x85, 0x45, 0x1F, 0x44, 0x64, 0xE9, 0xE0, 0x60 } };

enum ext_ConnectMode {
    ext_cm_AfterStartup = 0,
    ext_cm_Startup = 1,
    ext_cm_External = 2,
    ext_cm_CommandLine = 3,
    ext_cm_Solution = 4,
    ext_cm_UISetup = 5
};

enum ext_DisconnectMode {
    ext_dm_HostShutdown = 0,
    ext_dm_UserClosed = 1,
    ext_dm_UISetupComplete = 2,
    ext_dm_SolutionClosed = 3
};

// Only the first 15 methods are needed to reach Compile(). The remaining
// IVbaProject members are irrelevant here, and their signatures do not affect
// the vtable slot.
struct IVbaProject : public IUnknown {
    virtual HRESULT __stdcall GetName(BSTR* name) = 0;
    virtual HRESULT __stdcall SetName(LPCOLESTR name) = 0;
    virtual HRESULT __stdcall IsDirty(BOOL* dirty) = 0;
    virtual HRESULT __stdcall SetDirty(BOOL dirty) = 0;
    virtual HRESULT __stdcall GetVbaProcs(IUnknown** procs) = 0;
    virtual HRESULT __stdcall ParseLine(LPCOLESTR line) = 0;
    virtual HRESULT __stdcall ExecuteLine(LPCOLESTR line) = 0;
    virtual HRESULT __stdcall GetVbaObject(IUnknown** vba) = 0;
    virtual HRESULT __stdcall GetMode(int* mode) = 0;
    virtual HRESULT __stdcall SetMode(int mode) = 0;
    virtual HRESULT __stdcall GetProjItemCol(IUnknown** collection) = 0;
    virtual HRESULT __stdcall LoadComplete() = 0;
    virtual HRESULT __stdcall SetFileName(LPCOLESTR path) = 0;
    virtual HRESULT __stdcall GetFileName(BSTR* path) = 0;
    virtual HRESULT __stdcall Compile() = 0;
};

// Minimal IVba: GetExecutingProject is the 8th method. The seven methods before
// it are only placeholders for the vtable slot; none are called.
struct IVba : public IUnknown {
    virtual HRESULT __stdcall M1() = 0;
    virtual HRESULT __stdcall M2() = 0;
    virtual HRESULT __stdcall M3() = 0;
    virtual HRESULT __stdcall M4() = 0;
    virtual HRESULT __stdcall M5() = 0;
    virtual HRESULT __stdcall M6() = 0;
    virtual HRESULT __stdcall M7() = 0;
    virtual HRESULT __stdcall GetExecutingProject(IVbaProject** project) = 0;
};

// IVBEProject is implemented by the VBA project's ITypeLib. CompileProject is
// the 22nd method; the 21 methods before it are vtable placeholders.
struct IVBEProject : public IUnknown {
    virtual HRESULT __stdcall GetProjectName(BSTR* name) = 0;
    virtual HRESULT __stdcall M02() = 0;
    virtual HRESULT __stdcall M03() = 0;
    virtual HRESULT __stdcall M04() = 0;
    virtual HRESULT __stdcall M05() = 0;
    virtual HRESULT __stdcall M06() = 0;
    virtual HRESULT __stdcall M07() = 0;
    virtual HRESULT __stdcall M08() = 0;
    virtual HRESULT __stdcall M09() = 0;
    virtual HRESULT __stdcall M10() = 0;
    virtual HRESULT __stdcall M11() = 0;
    virtual HRESULT __stdcall M12() = 0;
    virtual HRESULT __stdcall M13() = 0;
    virtual HRESULT __stdcall M14() = 0;
    virtual HRESULT __stdcall M15() = 0;
    virtual HRESULT __stdcall M16() = 0;
    virtual HRESULT __stdcall M17() = 0;
    virtual HRESULT __stdcall M18() = 0;
    virtual HRESULT __stdcall M19() = 0;
    virtual HRESULT __stdcall M20() = 0;
    virtual HRESULT __stdcall M21() = 0;
    virtual HRESULT __stdcall CompileProject() = 0;
};

struct IDTExtensibility2 : public IDispatch {
    virtual HRESULT __stdcall OnConnection(IDispatch* application, ext_ConnectMode connectMode, IDispatch* addInInst, SAFEARRAY** custom) = 0;
    virtual HRESULT __stdcall OnDisconnection(ext_DisconnectMode removeMode, SAFEARRAY** custom) = 0;
    virtual HRESULT __stdcall OnAddInsUpdate(SAFEARRAY** custom) = 0;
    virtual HRESULT __stdcall OnStartupComplete(SAFEARRAY** custom) = 0;
    virtual HRESULT __stdcall OnBeginShutdown(SAFEARRAY** custom) = 0;
};

LONG g_objectCount = 0;
LONG g_lockCount = 0;

std::wstring escape_json(const std::wstring& text) {
    std::wstring out;
    out.reserve(text.size() + 8);
    for (wchar_t ch : text) {
        switch (ch) {
            case L'"': out += L"\\\""; break;
            case L'\\': out += L"\\\\"; break;
            case L'\n': out += L"\\n"; break;
            case L'\r': out += L"\\r"; break;
            case L'\t': out += L"\\t"; break;
            default:
                if (ch < 0x20) {
                    wchar_t buffer[8];
                    swprintf_s(buffer, L"\\u%04x", static_cast<unsigned>(ch));
                    out += buffer;
                } else {
                    out += ch;
                }
                break;
        }
    }
    return out;
}

std::wstring quote(const std::wstring& text) {
    return L"\"" + escape_json(text) + L"\"";
}

std::wstring hex_hr(HRESULT hr) {
    wchar_t buffer[16];
    swprintf_s(buffer, L"0x%08X", static_cast<unsigned>(hr));
    return buffer;
}

void log_line(const wchar_t* message) {
    wchar_t enabled[8] = {};
    if (GetEnvironmentVariableW(L"ACCESSMCP_VBEADDIN_LOG", enabled, _countof(enabled)) == 0) {
        return;
    }

    wchar_t temp[MAX_PATH] = {};
    if (GetTempPathW(MAX_PATH, temp) == 0) {
        return;
    }

    std::wstring path(temp);
    path += L"AccessMCP.VbeAddin.log";

    FILE* file = nullptr;
    if (_wfopen_s(&file, path.c_str(), L"a, ccs=UTF-8") != 0 || !file) {
        return;
    }

    SYSTEMTIME now;
    GetLocalTime(&now);
    fwprintf(file, L"%04d-%02d-%02d %02d:%02d:%02d.%03d %s\n",
             now.wYear, now.wMonth, now.wDay, now.wHour, now.wMinute, now.wSecond, now.wMilliseconds,
             message);
    fclose(file);
}

std::wstring error_json(const std::wstring& message, HRESULT hr) {
    return L"{\"ok\":false,\"error\":" + quote(message) + L",\"hresult\":" + quote(hex_hr(hr)) + L"}";
}

HRESULT get_property(IDispatch* object, const wchar_t* name, VARIANT* result) {
    if (!object || !result) {
        return E_POINTER;
    }

    DISPID id = 0;
    OLECHAR* member = const_cast<OLECHAR*>(name);
    HRESULT hr = object->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) {
        return hr;
    }

    DISPPARAMS params = {};
    VariantInit(result);
    return object->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, result, nullptr, nullptr);
}

HRESULT get_item(IDispatch* collection, LONG index, VARIANT* result) {
    if (!collection || !result) {
        return E_POINTER;
    }

    DISPID id = 0;
    OLECHAR* member = const_cast<OLECHAR*>(L"Item");
    HRESULT hr = collection->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) {
        return hr;
    }

    VARIANT argument;
    VariantInit(&argument);
    argument.vt = VT_I4;
    argument.lVal = index;

    DISPPARAMS params = {};
    params.rgvarg = &argument;
    params.cArgs = 1;
    VariantInit(result);

    hr = collection->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, result, nullptr, nullptr);
    if (FAILED(hr)) {
        hr = collection->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &params, result, nullptr, nullptr);
    }

    VariantClear(&argument);
    return hr;
}

void log_qi(const wchar_t* label, HRESULT hr) {
    wchar_t buffer[128];
    swprintf_s(buffer, L"%s hr=0x%08X", label, static_cast<unsigned>(hr));
    log_line(buffer);
}

// Project.References wraps a structure whose sixth pointer (x86) is the project
// ITypeLib. This mirrors Rubberduck's VBEReferencesObj layout:
//   vtable1, vtable2, vtable3, object1, object2, _typeLib, ...
#if defined(_M_IX86)
constexpr size_t kReferencesTypeLibOffset = 0x14;
#elif defined(_M_AMD64)
constexpr size_t kReferencesTypeLibOffset = 0x28;
#else
#error Unsupported architecture
#endif

bool read_pointer_at(const void* base, size_t offset, void** out) {
    if (!base || !out) {
        return false;
    }

    const auto* address = static_cast<const unsigned char*>(base) + offset;
    MEMORY_BASIC_INFORMATION info = {};
    if (VirtualQuery(address, &info, sizeof(info)) == 0) {
        return false;
    }
    if (info.State != MEM_COMMIT || (info.Protect & (PAGE_NOACCESS | PAGE_GUARD)) != 0) {
        return false;
    }

    const auto* end = static_cast<const unsigned char*>(info.BaseAddress) + info.RegionSize;
    if (address + sizeof(void*) > end) {
        return false;
    }

    memcpy(out, address, sizeof(void*));
    return true;
}

// The public VBProject wrapper does not expose IVbaProject. Rubberduck reaches
// the compiler through the project's ITypeLib, which implements IVBEProject.
IVBEProject* find_vbe_project(IDispatch* project) {
    if (!project) {
        return nullptr;
    }

    IVBEProject* result = nullptr;

    // Preferred route: Project.References -> internal _typeLib -> IVBEProject.
    VARIANT references;
    VariantInit(&references);
    HRESULT referencesHr = get_property(project, L"References", &references);
    if (SUCCEEDED(referencesHr) && references.vt == VT_DISPATCH && references.pdispVal) {
        void* rawTypeLib = nullptr;
        if (read_pointer_at(references.pdispVal, kReferencesTypeLibOffset, &rawTypeLib) && rawTypeLib) {
            void* vtable = nullptr;
            if (read_pointer_at(rawTypeLib, 0, &vtable) && vtable) {
                auto* typeLib = static_cast<ITypeLib*>(rawTypeLib);

                BSTR libName = nullptr;
                HRESULT doc = typeLib->GetDocumentation(-1, &libName, nullptr, nullptr, nullptr);
                if (SUCCEEDED(doc) && libName) {
                    std::wstring label(L"References._typeLib name=");
                    label += libName;
                    log_line(label.c_str());
                    SysFreeString(libName);
                } else {
                    log_qi(L"References._typeLib GetDocumentation", doc);
                }

                HRESULT q = typeLib->QueryInterface(IID_IVBEProject, reinterpret_cast<void**>(&result));
                log_qi(L"References._typeLib QI IVBEProject", q);
                if (SUCCEEDED(q) && result) {
                    BSTR projectName = nullptr;
                    HRESULT nameHr = result->GetProjectName(&projectName);
                    if (SUCCEEDED(nameHr) && projectName) {
                        std::wstring label(L"IVBEProject project name=");
                        label += projectName;
                        log_line(label.c_str());
                        SysFreeString(projectName);
                    } else {
                        log_qi(L"IVBEProject GetProjectName", nameHr);
                    }
                    VariantClear(&references);
                    return result;
                }
            } else {
                log_line(L"References._typeLib is not a readable COM object");
            }
        } else {
            log_line(L"could not read References._typeLib");
        }
    } else {
        log_qi(L"project References property", referencesHr);
    }
    VariantClear(&references);

    HRESULT hr = project->QueryInterface(IID_IVBEProject, reinterpret_cast<void**>(&result));
    log_qi(L"project QI IVBEProject", hr);
    if (SUCCEEDED(hr) && result) {
        return result;
    }

    ITypeLib* typeLib = nullptr;
    hr = project->QueryInterface(IID_ITypeLibStd, reinterpret_cast<void**>(&typeLib));
    log_qi(L"project QI ITypeLib", hr);
    if (SUCCEEDED(hr) && typeLib) {
        hr = typeLib->QueryInterface(IID_IVBEProject, reinterpret_cast<void**>(&result));
        log_qi(L"project ITypeLib QI IVBEProject", hr);
        typeLib->Release();
        if (SUCCEEDED(hr) && result) {
            return result;
        }
    }

    VARIANT components;
    VariantInit(&components);
    hr = get_property(project, L"VBComponents", &components);
    if (SUCCEEDED(hr) && components.vt == VT_DISPATCH && components.pdispVal) {
        VARIANT component;
        VariantInit(&component);
        hr = get_item(components.pdispVal, 1, &component);
        if (SUCCEEDED(hr) && component.vt == VT_DISPATCH && component.pdispVal) {
            ITypeInfo* typeInfo = nullptr;
            hr = component.pdispVal->QueryInterface(IID_ITypeInfoStd, reinterpret_cast<void**>(&typeInfo));
            log_qi(L"VBComponent QI ITypeInfo", hr);
            if (SUCCEEDED(hr) && typeInfo) {
                ITypeLib* containing = nullptr;
                UINT index = 0;
                hr = typeInfo->GetContainingTypeLib(&containing, &index);
                log_qi(L"ITypeInfo GetContainingTypeLib", hr);
                if (SUCCEEDED(hr) && containing) {
                    hr = containing->QueryInterface(IID_IVBEProject, reinterpret_cast<void**>(&result));
                    log_qi(L"containing ITypeLib QI IVBEProject", hr);
                    containing->Release();
                }
                typeInfo->Release();
                if (SUCCEEDED(hr) && result) {
                    VariantClear(&component);
                    VariantClear(&components);
                    return result;
                }
            }

            VARIANT codeModule;
            VariantInit(&codeModule);
            hr = get_property(component.pdispVal, L"CodeModule", &codeModule);
            if (SUCCEEDED(hr) && codeModule.vt == VT_DISPATCH && codeModule.pdispVal) {
                ITypeInfo* moduleTypeInfo = nullptr;
                hr = codeModule.pdispVal->QueryInterface(IID_ITypeInfoStd, reinterpret_cast<void**>(&moduleTypeInfo));
                log_qi(L"CodeModule QI ITypeInfo", hr);
                if (SUCCEEDED(hr) && moduleTypeInfo) {
                    ITypeLib* moduleContaining = nullptr;
                    UINT moduleIndex = 0;
                    hr = moduleTypeInfo->GetContainingTypeLib(&moduleContaining, &moduleIndex);
                    log_qi(L"CodeModule ITypeInfo GetContainingTypeLib", hr);
                    if (SUCCEEDED(hr) && moduleContaining) {
                        hr = moduleContaining->QueryInterface(IID_IVBEProject, reinterpret_cast<void**>(&result));
                        log_qi(L"CodeModule containing ITypeLib QI IVBEProject", hr);
                        moduleContaining->Release();
                    }
                    moduleTypeInfo->Release();
                }
                VariantClear(&codeModule);
            }
        }
        VariantClear(&component);
    }
    VariantClear(&components);

    return (SUCCEEDED(hr) && result) ? result : nullptr;
}

class CVbeProbe : public IDTExtensibility2 {
public:
    CVbeProbe() : _refCount(1), _vbe(nullptr), _addin(nullptr) {
        InterlockedIncrement(&g_objectCount);
    }

    ~CVbeProbe() {
        if (_vbe) {
            _vbe->Release();
            _vbe = nullptr;
        }
        if (_addin) {
            _addin->Release();
            _addin = nullptr;
        }
        InterlockedDecrement(&g_objectCount);
    }

    // IUnknown
    HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) {
            return E_POINTER;
        }

        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDTExtensibility2)) {
            *ppv = static_cast<IDTExtensibility2*>(this);
        } else if (IsEqualIID(riid, IID_IDispatch)) {
            *ppv = static_cast<IDispatch*>(this);
        } else {
            *ppv = nullptr;
            return E_NOINTERFACE;
        }

        AddRef();
        return S_OK;
    }

    ULONG __stdcall AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&_refCount));
    }

    ULONG __stdcall Release() override {
        LONG value = InterlockedDecrement(&_refCount);
        if (value == 0) {
            delete this;
        }
        return static_cast<ULONG>(value);
    }

    // IDispatch
    HRESULT __stdcall GetTypeInfoCount(UINT* count) override {
        if (!count) {
            return E_POINTER;
        }
        *count = 0;
        return S_OK;
    }

    HRESULT __stdcall GetTypeInfo(UINT, LCID, ITypeInfo** info) override {
        if (info) {
            *info = nullptr;
        }
        return E_NOTIMPL;
    }

    HRESULT __stdcall GetIDsOfNames(REFIID, LPOLESTR* names, UINT nameCount, LCID, DISPID* ids) override {
        if (!names || !ids || nameCount != 1) {
            return DISP_E_UNKNOWNNAME;
        }

        if (_wcsicmp(names[0], L"CompileProject") == 0) {
            ids[0] = 10;
            return S_OK;
        }
        if (_wcsicmp(names[0], L"GetLastCompileError") == 0) {
            ids[0] = 11;
            return S_OK;
        }
        if (_wcsicmp(names[0], L"ReadErrorRecord") == 0) {
            ids[0] = 12;
            return S_OK;
        }

        // IDTExtensibility2 dispatch IDs are 1-5.
        static const wchar_t* extensibilityNames[] = {
            L"OnConnection", L"OnDisconnection", L"OnAddInsUpdate", L"OnStartupComplete", L"OnBeginShutdown"
        };
        for (int i = 0; i < 5; ++i) {
            if (_wcsicmp(names[0], extensibilityNames[i]) == 0) {
                ids[0] = 1 + i;
                return S_OK;
            }
        }

        return DISP_E_UNKNOWNNAME;
    }

    HRESULT __stdcall Invoke(DISPID id, REFIID, LCID, WORD flags, DISPPARAMS* params, VARIANT* result, EXCEPINFO*, UINT*) override {
        if (id == 10 || id == 11 || id == 12) {
            if (!(flags & DISPATCH_METHOD)) {
                return DISP_E_MEMBERNOTFOUND;
            }

            std::wstring json;
            HRESULT hr = (id == 10) ? CompileProject(json)
                       : (id == 11) ? GetLastCompileError(json)
                       : ReadErrorRecord(json);
            if (FAILED(hr)) {
                return hr;
            }

            if (result) {
                VariantInit(result);
                result->vt = VT_BSTR;
                result->bstrVal = SysAllocString(json.c_str());
                if (!result->bstrVal) {
                    return E_OUTOFMEMORY;
                }
            }
            return S_OK;
        }

        if (id >= 1 && id <= 5) {
            return InvokeExtensibility(static_cast<int>(id) - 1, params);
        }

        return DISP_E_MEMBERNOTFOUND;
    }

    // IDTExtensibility2
    HRESULT __stdcall OnConnection(IDispatch* application, ext_ConnectMode, IDispatch* addInInst, SAFEARRAY**) override {
        log_line(L"OnConnection enter");
        if (_vbe) {
            _vbe->Release();
        }
        _vbe = application;
        if (_vbe) {
            _vbe->AddRef();
        }

        if (_addin) {
            _addin->Release();
        }
        _addin = addInInst;
        if (_addin) {
            _addin->AddRef();
            SetAddInObject(_addin);
        }

        log_line(L"OnConnection exit");
        return S_OK;
    }

    HRESULT __stdcall OnDisconnection(ext_DisconnectMode, SAFEARRAY**) override {
        log_line(L"OnDisconnection enter");
        if (_vbe) {
            _vbe->Release();
            _vbe = nullptr;
        }
        if (_addin) {
            _addin->Release();
            _addin = nullptr;
        }
        log_line(L"OnDisconnection exit");
        return S_OK;
    }

    HRESULT __stdcall OnAddInsUpdate(SAFEARRAY**) override { return S_OK; }
    HRESULT __stdcall OnStartupComplete(SAFEARRAY**) override { return S_OK; }
    HRESULT __stdcall OnBeginShutdown(SAFEARRAY**) override { return S_OK; }

private:
    HRESULT CompileProject(std::wstring& json) {
        json = L"{\"ok\":false,\"error\":\"The VBE is not connected.\"}";
        if (!_vbe) {
            return S_OK;
        }

        VARIANT projects;
        VariantInit(&projects);
        HRESULT hr = get_property(_vbe, L"VBProjects", &projects);
        if (FAILED(hr) || projects.vt != VT_DISPATCH || !projects.pdispVal) {
            json = error_json(L"Could not access VBE.VBProjects.", hr);
            VariantClear(&projects);
            return S_OK;
        }

        VARIANT project;
        VariantInit(&project);
        hr = get_item(projects.pdispVal, 1, &project);
        if (FAILED(hr) || project.vt != VT_DISPATCH || !project.pdispVal) {
            json = error_json(L"Could not access VBProjects.Item(1).", hr);
            VariantClear(&project);
            VariantClear(&projects);
            return S_OK;
        }

        HRESULT compile = E_NOINTERFACE;
        bool hasCompiler = false;

        // Preferred: the project's ITypeLib implements IVBEProject, whose
        // CompileProject is silent (no VBE dialog) and sets BTSRC_ERR.
        IVBEProject* vbeProject = find_vbe_project(project.pdispVal);
        if (vbeProject) {
            log_line(L"using IVBEProject::CompileProject");
            compile = vbeProject->CompileProject();
            vbeProject->Release();
            hasCompiler = true;
        } else {
            // Fallback: the native IVbaProject / IVba interfaces.
            IVbaProject* vbaProject = nullptr;
            HRESULT query = project.pdispVal->QueryInterface(IID_IVbaProject, reinterpret_cast<void**>(&vbaProject));

            if (FAILED(query) || !vbaProject) {
                log_line(L"public VBProject QI IVbaProject failed; trying IVba");
                IVba* engine = nullptr;
                HRESULT qiEngine = _vbe->QueryInterface(IID_IVba, reinterpret_cast<void**>(&engine));
                if (SUCCEEDED(qiEngine) && engine) {
                    query = engine->GetExecutingProject(&vbaProject);
                    wchar_t buffer[96];
                    swprintf_s(buffer, L"IVba::GetExecutingProject hr=0x%08X project=%p",
                               static_cast<unsigned>(query), vbaProject);
                    log_line(buffer);
                    engine->Release();
                } else {
                    wchar_t buffer[96];
                    swprintf_s(buffer, L"VBE QI IVba failed 0x%08X", static_cast<unsigned>(qiEngine));
                    log_line(buffer);
                }
            }

            if (SUCCEEDED(query) && vbaProject) {
                compile = vbaProject->Compile();
                vbaProject->Release();
                hasCompiler = true;
            }
        }

        if (!hasCompiler) {
            json = error_json(L"No compiler interface is available in-process from this VBProject.", compile);
            VariantClear(&project);
            VariantClear(&projects);
            return S_OK;
        }

        if (SUCCEEDED(compile)) {
            json = L"{\"ok\":true,\"compiled\":true}";
        } else {
            // Read the record before anything can call BuildErrorInfo, which clears it.
            vt::VbeCompileError record = vt::read_compile_error_record();

            _com_error error(compile);
            _bstr_t errorDescription = error.Description();
            std::wstring description;
            if (errorDescription.length() > 0) {
                description = static_cast<const wchar_t*>(errorDescription);
            }

            json = L"{\"ok\":false,\"compiled\":false,\"hresult\":" + quote(hex_hr(compile)) +
                   L",\"description\":" + quote(description);

            if (record.present) {
                json += L",\"errorNumber\":" + std::to_wstring(record.errorNumber);
                json += L",\"line\":" + std::to_wstring(record.line);
                if (record.column != 0) {
                    json += L",\"column\":" + std::to_wstring(record.column);
                }
                if (!record.module.empty()) {
                    json += L",\"module\":" + quote(record.module);
                }
                if (!record.project.empty()) {
                    json += L",\"project\":" + quote(record.project);
                }
            }

            json += L"}";
        }

        _lastJson = json;
        VariantClear(&project);
        VariantClear(&projects);
        return S_OK;
    }

    HRESULT GetLastCompileError(std::wstring& json) {
        json = _lastJson.empty() ? L"{\"ok\":false,\"error\":\"No compile has been run.\"}" : _lastJson;
        return S_OK;
    }

    HRESULT ReadErrorRecord(std::wstring& json) {
        vt::VbeCompileError record = vt::read_compile_error_record();
        json = L"{\"present\":" + std::wstring(record.present ? L"true" : L"false");
        if (record.present) {
            json += L",\"errorNumber\":" + std::to_wstring(record.errorNumber);
            json += L",\"line\":" + std::to_wstring(record.line);
            if (record.column != 0) {
                json += L",\"column\":" + std::to_wstring(record.column);
            }
            if (!record.module.empty()) {
                json += L",\"module\":" + quote(record.module);
            }
            if (!record.project.empty()) {
                json += L",\"project\":" + quote(record.project);
            }
            if (!record.message.empty()) {
                json += L",\"message\":" + quote(record.message);
            }
        }
        json += L"}";
        _lastJson = json;
        return S_OK;
    }

    HRESULT InvokeExtensibility(int index, DISPPARAMS* params) {
        if (!params) {
            return E_INVALIDARG;
        }

        VARIANT* args = params->rgvarg;
        UINT count = params->cArgs;

        auto dispatchAt = [&](UINT i) -> IDispatch* {
            if (i >= count) {
                return nullptr;
            }
            return args[i].vt == VT_DISPATCH ? args[i].pdispVal : nullptr;
        };
        auto longAt = [&](UINT i) -> LONG {
            if (i >= count) {
                return 0;
            }
            if (args[i].vt == VT_I4) {
                return args[i].lVal;
            }
            if (args[i].vt == VT_I2) {
                return args[i].iVal;
            }
            return 0;
        };

        switch (index) {
            case 0: // OnConnection(custom, AddInInst, ConnectMode, Application)
                return OnConnection(dispatchAt(3), static_cast<ext_ConnectMode>(longAt(2)), dispatchAt(1), nullptr);
            case 1: // OnDisconnection(custom, RemoveMode)
                return OnDisconnection(static_cast<ext_DisconnectMode>(longAt(1)), nullptr);
            case 2:
                return OnAddInsUpdate(nullptr);
            case 3:
                return OnStartupComplete(nullptr);
            case 4:
                return OnBeginShutdown(nullptr);
            default:
                return DISP_E_MEMBERNOTFOUND;
        }
    }

    void SetAddInObject(IDispatch* addIn) {
        DISPID id = 0;
        OLECHAR* name = const_cast<OLECHAR*>(L"Object");
        HRESULT hr = addIn->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &id);
        if (FAILED(hr)) {
            wchar_t buffer[96];
            swprintf_s(buffer, L"SetAddInObject GetIDsOfNames(Object) failed 0x%08X", static_cast<unsigned>(hr));
            log_line(buffer);
            return;
        }

        VARIANT value;
        VariantInit(&value);
        value.vt = VT_DISPATCH;
        value.pdispVal = static_cast<IDispatch*>(this);
        AddRef();

        DISPID put = DISPID_PROPERTYPUT;
        DISPPARAMS params = {};
        params.rgvarg = &value;
        params.rgdispidNamedArgs = &put;
        params.cArgs = 1;
        params.cNamedArgs = 1;

        hr = addIn->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &params, nullptr, nullptr, nullptr);
        VariantClear(&value);

        wchar_t buffer[96];
        swprintf_s(buffer, L"SetAddInObject Invoke(put Object) hr=0x%08X", static_cast<unsigned>(hr));
        log_line(buffer);
    }

    LONG _refCount;
    IDispatch* _vbe;
    IDispatch* _addin;
    std::wstring _lastJson;
};

class CClassFactory : public IClassFactory {
public:
    CClassFactory() : _refCount(1) {}

    HRESULT __stdcall QueryInterface(REFIID riid, void** ppv) override {
        if (!ppv) {
            return E_POINTER;
        }
        if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory)) {
            *ppv = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }
        *ppv = nullptr;
        return E_NOINTERFACE;
    }

    ULONG __stdcall AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&_refCount));
    }

    ULONG __stdcall Release() override {
        LONG value = InterlockedDecrement(&_refCount);
        if (value == 0) {
            delete this;
        }
        return static_cast<ULONG>(value);
    }

    HRESULT __stdcall CreateInstance(IUnknown* outer, REFIID riid, void** ppv) override {
        if (outer) {
            return CLASS_E_NOAGGREGATION;
        }
        if (!ppv) {
            return E_POINTER;
        }

        CVbeProbe* probe = new (std::nothrow) CVbeProbe();
        if (!probe) {
            return E_OUTOFMEMORY;
        }

        HRESULT hr = probe->QueryInterface(riid, ppv);
        probe->Release();
        return hr;
    }

    HRESULT __stdcall LockServer(BOOL lock) override {
        if (lock) {
            InterlockedIncrement(&g_lockCount);
        } else {
            InterlockedDecrement(&g_lockCount);
        }
        return S_OK;
    }

private:
    LONG _refCount;
};

} // namespace

BOOL APIENTRY DllMain(HMODULE module, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(module);
    }
    return TRUE;
}

extern "C" HRESULT __stdcall DllGetClassObject(REFCLSID clsid, REFIID riid, void** ppv) {
    if (!IsEqualCLSID(clsid, CLSID_AccessMcpVbeProbe)) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    CClassFactory* factory = new (std::nothrow) CClassFactory();
    if (!factory) {
        return E_OUTOFMEMORY;
    }

    HRESULT hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}

extern "C" HRESULT __stdcall DllCanUnloadNow() {
    return (g_objectCount == 0 && g_lockCount == 0) ? S_OK : S_FALSE;
}

extern "C" HRESULT __stdcall DllRegisterServer() {
    // Registration is per-user and performed by scripts/register-vbe-addin.ps1.
    return S_OK;
}

extern "C" HRESULT __stdcall DllUnregisterServer() {
    return S_OK;
}
