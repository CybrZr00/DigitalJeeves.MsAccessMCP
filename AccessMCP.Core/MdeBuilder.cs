using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Loader;
using Microsoft.Office.Interop.Access;
using AccessApplication = Microsoft.Office.Interop.Access.Application;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>Outcome of an MDE/ACCDE build.</summary>
internal sealed record MdeBuildResult(
    bool Ok,
    string Strategy,
    string SourcePath,
    string? OutputPath,
    bool Verified,
    bool ManualRequired,
    string? Error);

/// <summary>
/// Builds a compiled Access database (.mde/.accde).
/// <para>
/// <c>SysCmd(603, source, target)</c> is undocumented and has one hard requirement discovered on
/// Access 365 16.0.20430 x86: the call must be made on the <b>same Access instance</b> that opened
/// the source, compiled it and then closed it. A fresh instance silently returns 0 without creating
/// a file. The call must also be early-bound: a late-bound call with the same integer no-ops.
/// </para>
/// </summary>
internal static class MdeBuilder
{
    private const int MakeMdeAction = 603;

    static MdeBuilder()
    {
        // The Office PIAs (office, DAO, VBE) are deployed next to the host but are not part of the
        // app's .deps.json, so the default probe cannot find them by simple name. Resolve them from
        // the application directory on demand.
        AssemblyLoadContext.Default.Resolving += (context, name) =>
        {
            var candidate = Path.Combine(AppContext.BaseDirectory, name.Name + ".dll");
            return File.Exists(candidate) ? context.LoadFromAssemblyPath(candidate) : null;
        };
    }

    /// <summary>
    /// Builds <paramref name="outputPath"/> from <paramref name="sourcePath"/> using the live Access
    /// instance in <paramref name="appObject"/>. The caller must have opened the source, compiled it
    /// (IsCompiled true) and closed it on that instance immediately before calling this.
    /// </summary>
    public static MdeBuildResult BuildWithSysCmd(object appObject, string sourcePath, string outputPath)
    {
        if (appObject is not AccessApplication app)
        {
            // An RCW created through IDispatch (Activator.CreateInstance + dynamic) sometimes fails
            // the direct interface cast even though the COM object supports it. Query the interface
            // through IUnknown explicitly before giving up.
            try
            {
                var unknown = Marshal.GetIUnknownForObject(appObject);
                try
                {
                    app = (AccessApplication)Marshal.GetTypedObjectForIUnknown(unknown, typeof(AccessApplication));
                }
                finally
                {
                    Marshal.Release(unknown);
                }
            }
            catch (Exception ex)
            {
                return new MdeBuildResult(
                    false,
                    "syscmd",
                    sourcePath,
                    outputPath,
                    false,
                    true,
                    "The Access automation object cannot be queried for the early-bound " +
                    "Microsoft.Office.Interop.Access.Application interface: " + ex.Message + ". The Office PIAs " +
                    "(office, Microsoft.Office.Interop.Access.Dao, Microsoft.Vbe.Interop) must be deployed next to " +
                    "the MCP host.");
            }
        }

        try
        {
            if (File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            // Early-bound by construction: the PIA types the first parameter as AcSysCmdAction.
            // The enum does not define 603; the cast is deliberate.
            var result = app.SysCmd((AcSysCmdAction)MakeMdeAction, sourcePath, outputPath);

            if (!File.Exists(outputPath))
            {
                return new MdeBuildResult(
                    false,
                    "syscmd",
                    sourcePath,
                    outputPath,
                    false,
                    true,
                    $"SysCmd({MakeMdeAction}) returned [{result}] but did not create '{outputPath}'. " +
                    "Use the 'prebuilt' strategy or build the MDE manually.");
            }

            var size = WaitForStableSize(outputPath);
            if (size <= 0)
            {
                return new MdeBuildResult(false, "syscmd", sourcePath, outputPath, false, true,
                    "The generated file is empty.");
            }

            return new MdeBuildResult(true, "syscmd", sourcePath, outputPath, true, false, null);
        }
        catch (Exception ex)
        {
            return new MdeBuildResult(
                false,
                "syscmd",
                sourcePath,
                outputPath,
                false,
                true,
                $"SysCmd({MakeMdeAction}) failed: {ex.Message} (0x{ex.HResult:X8})");
        }
    }

    /// <summary>Copies a caller-supplied compiled database into the workspace.</summary>
    public static MdeBuildResult UsePrebuilt(string prebuiltPath, string outputPath)
    {
        if (!File.Exists(prebuiltPath))
        {
            return new MdeBuildResult(false, "prebuilt", prebuiltPath, outputPath, false, true,
                $"Prebuilt database not found: {prebuiltPath}");
        }

        if (!AccessDbPathPolicy.IsCompiledDatabase(prebuiltPath))
        {
            return new MdeBuildResult(false, "prebuilt", prebuiltPath, outputPath, false, true,
                "prebuiltPath must be an .mde or .accde file.");
        }

        Directory.CreateDirectory(Path.GetDirectoryName(outputPath) ?? ".");

        if (Path.GetFullPath(prebuiltPath).Equals(Path.GetFullPath(outputPath), StringComparison.OrdinalIgnoreCase))
        {
            return new MdeBuildResult(true, "prebuilt", prebuiltPath, outputPath, true, false, null);
        }

        File.Copy(prebuiltPath, outputPath, overwrite: true);

        return new MdeBuildResult(true, "prebuilt", prebuiltPath, outputPath, true, false, null);
    }

    /// <summary>Default output extension for a source database.</summary>
    public static string DefaultExtension(string sourcePath)
        => AccessDbPathPolicy.IsCompiledDatabase(sourcePath)
            ? Path.GetExtension(sourcePath)
            : Path.GetExtension(sourcePath).Equals(".accdb", StringComparison.OrdinalIgnoreCase) ? ".accde" : ".mde";

    private static long WaitForStableSize(string path, int attempts = 10)
    {
        long previous = -1;
        for (var i = 0; i < attempts; i++)
        {
            try
            {
                var size = new FileInfo(path).Length;
                if (size > 0 && size == previous)
                {
                    return size;
                }

                previous = size;
            }
            catch
            {
                return 0;
            }

            Thread.Sleep(150);
        }

        return previous > 0 ? previous : 0;
    }
}
