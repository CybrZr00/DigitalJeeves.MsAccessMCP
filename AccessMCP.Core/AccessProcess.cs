using System.Diagnostics;
using System.Globalization;
using System.IO;
using Microsoft.Win32;

namespace DigitalJeeves.MsAccessMCP.Core;

/// <summary>
/// Locates and describes the installed Microsoft Access executable. Discovery is registry-first so
/// it works for Click-to-Run installs and does not require COM.
/// </summary>
internal static class AccessProcess
{
    private static readonly string[] FallbackPaths =
    [
        @"C:\Program Files (x86)\Microsoft Office\root\Office16\MSACCESS.EXE",
        @"C:\Program Files\Microsoft Office\root\Office16\MSACCESS.EXE",
        @"C:\Program Files (x86)\Microsoft Office\Office16\MSACCESS.EXE",
        @"C:\Program Files\Microsoft Office\Office16\MSACCESS.EXE",
    ];

    /// <summary>Full path to MSACCESS.EXE, or null when Access is not installed.</summary>
    public static string? FindExecutable()
    {
        foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                using var key = baseKey.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE");
                if (key?.GetValue(null) is string path && File.Exists(path))
                {
                    return Path.GetFullPath(path);
                }
            }
            catch
            {
                // Registry view unavailable; try the next one.
            }
        }

        return FallbackPaths.FirstOrDefault(File.Exists);
    }

    /// <summary>Product version and build, for example 16.0.20430.20182.</summary>
    public static string? GetVersion(string executablePath)
    {
        try
        {
            return FileVersionInfo.GetVersionInfo(executablePath).FileVersion;
        }
        catch
        {
            return null;
        }
    }

    /// <summary>PE machine architecture of an executable: x86, x64, arm64 or unknown.</summary>
    public static string GetBitness(string executablePath)
    {
        try
        {
            using var stream = File.OpenRead(executablePath);
            using var reader = new BinaryReader(stream);
            if (stream.Length < 0x40)
            {
                return "unknown";
            }

            stream.Position = 0x3C;
            var peOffset = reader.ReadInt32();
            if (peOffset <= 0 || peOffset + 6 > stream.Length)
            {
                return "unknown";
            }

            stream.Position = peOffset;
            if (reader.ReadUInt32() != 0x00004550)
            {
                return "unknown";
            }

            return reader.ReadUInt16() switch
            {
                0x014C => "x86",
                0x8664 => "x64",
                0xAA64 => "arm64",
                _ => "unknown"
            };
        }
        catch
        {
            return "unknown";
        }
    }

    /// <summary>
    /// True when the Office install looks like Access Runtime rather than full Access. The runtime
    /// has no VBE, so it cannot compile, edit modules or build MDEs.
    /// </summary>
    public static bool LooksLikeRuntime(string executablePath)
    {
        var directory = Path.GetDirectoryName(executablePath);
        if (directory is null)
        {
            return false;
        }

        // Full Access ships MSACC.OLB next to MSACCESS.EXE; the shared VBE lives in Common Files.
        // The runtime has neither a type library nor a VBE. Checking both avoids a false positive on
        // Click-to-Run installs, where VBE7.DLL is not in the Office root.
        var typeLibrary = Path.Combine(directory, "MSACC.OLB");
        var vbe = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.CommonProgramFilesX86),
            "Microsoft Shared",
            "VBA",
            "VBA7.1",
            "VBE7.DLL");

        return !File.Exists(typeLibrary) && !File.Exists(vbe);
    }

    /// <summary>
    /// Runs MSACCESS.EXE /compact /decompile against a database copy. Returns the exit code, or null
    /// when the process had to be killed after <paramref name="timeout"/>.
    /// </summary>
    public static int? Decompile(string executablePath, string databasePath, TimeSpan timeout)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = executablePath,
                UseShellExecute = false,
                CreateNoWindow = false
            }
        };

        process.StartInfo.ArgumentList.Add("/compact");
        process.StartInfo.ArgumentList.Add("/decompile");
        process.StartInfo.ArgumentList.Add(databasePath);

        process.Start();
        if (!process.WaitForExit(timeout))
        {
            try { process.Kill(entireProcessTree: true); } catch { }
            return null;
        }

        return process.ExitCode;
    }

    /// <summary>
    /// True when an Access lock file sits beside the database (.ldb or .laccdb). Copying a database
    /// while another process has it open can produce a torn copy, so callers refuse to proceed.
    /// </summary>
    public static bool HasLockFile(string databasePath)
    {
        var ldb = Path.ChangeExtension(databasePath, ".ldb");
        var laccdb = Path.ChangeExtension(databasePath, ".laccdb");
        return File.Exists(ldb) || File.Exists(laccdb);
    }

    /// <summary>Best-effort description of the installed Access, for diagnostics.</summary>
    public static AccessInstallation Describe()
    {
        var path = FindExecutable();
        if (path is null)
        {
            return new AccessInstallation(false, null, null, null, null, false);
        }

        var version = GetVersion(path);
        var bitness = GetBitness(path);
        var isRuntime = LooksLikeRuntime(path);
        var build = version;

        return new AccessInstallation(true, path, version, build, bitness, isRuntime);
    }
}

/// <summary>Immutable description of the installed Access executable.</summary>
internal sealed record AccessInstallation(
    bool Installed,
    string? ExecutablePath,
    string? Version,
    string? Build,
    string? Bitness,
    bool IsRuntime)
{
    public static AccessInstallation None { get; } = new(false, null, null, null, null, false);

    public string? VersionText => Version is null
        ? null
        : Version.Split('.') is { Length: >= 2 } parts
            ? $"{parts[0]}.{parts[1]}"
            : Version;

    public override string ToString()
        => Installed
            ? $"{VersionText ?? "unknown"} {Bitness}{(IsRuntime ? " (runtime)" : string.Empty)}"
            : "not installed";
}
