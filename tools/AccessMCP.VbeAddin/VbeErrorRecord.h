// Copied from VBASA/VbaTestHost/VbeErrorRecord.h for the AccessMCP VBE add-in.
#pragma once

// Reads the VBE engine's global compile-error record directly, the way the
// VBE's own GetCompileErrorLocation does, by reaching it through the exported
// rtcErrObj -> _g_EbThread -> BTSRC_ERR chain. This is x86 (32-bit Office)
// only: the offsets below are for that build. On any other target the reader
// answers "nothing" and the host falls back to the description alone.
//
// The offsets come from reverse engineering the x86 VBE7.DLL; see
// mcp/README.md for the chain and the structure layouts.

#include <string>
#include <vector>

namespace vt {

struct VbeCompileError {
    bool present = false;
    unsigned errorNumber = 0;   // the VBE's error code (e.g. 0x9c64)
    unsigned lineUnits = 0;     // the record's raw line counter
    unsigned line = 0;          // logical line of the erroring statement
    unsigned column = 0;        // best-effort; only valid when columns are set
    std::wstring module;        // module the compiler blamed, if resolved
    std::wstring project;       // project name, if resolved
    std::wstring message;       // compiler message, when readable
    std::vector<std::wstring> probe;  // diagnostics from the pointer walk
};

// Returns an empty record (present == false) when there is no compile error to
// report or the chain does not look the way this build expects it.
VbeCompileError read_compile_error_record();

} // namespace vt
