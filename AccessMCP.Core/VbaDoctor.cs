using System.IO;
using Microsoft.Win32;

namespace DigitalJeeves.MsAccessMCP.Core;

internal sealed record VbaDoctorAccess(bool Installed, string? ExecutablePath, string? Version, string? Build, string? Bitness, bool IsRuntime);
internal sealed record VbaDoctorHost(bool Is64BitProcess, string Runtime);
internal sealed record VbaDoctorDatabase(string Path, string Extension, string Kind, bool Exists);
internal sealed record VbaDoctorVbe(bool Accessible, string? TrustMessage, string? ProjectName, int? ModuleCount, bool AddinLoadable);
internal sealed record VbaDoctorAce(string Provider, bool Available, string? Bitness);
internal sealed record VbaDoctorStorage(string Root, bool Writable, long FreeBytes, string? Error);
internal sealed record VbaDoctorAddin(string ProgId, bool Registered, bool Loadable);
internal sealed record VbaDoctorErrorTable(int Codes, string? SourceHash);
internal sealed record VbaDoctorCapabilities(bool CanEditModules, bool CanCompile, bool CanRun, bool CanBuildMde, bool PiaAvailable, string MdeStrategy);
internal sealed record VbaDoctorInfo(
    VbaDoctorAccess Access,
    VbaDoctorHost Host,
    VbaDoctorDatabase? Database,
    VbaDoctorVbe Vbe,
    VbaDoctorAddin VbeAddin,
    VbaDoctorErrorTable ErrorTable,
    VbaDoctorAce Ace,
    VbaDoctorStorage Scratch,
    VbaDoctorStorage Cache,
    VbaDoctorCapabilities Capabilities);

/// <summary>
/// Environment probe behind <c>access_vba_doctor</c>. Everything is best-effort: the doctor reports
/// what will and will not work rather than throwing.
/// </summary>
internal static class VbaDoctor
{
    public static async Task<VbaDoctorInfo> ProbeAsync(
        AccessMcpOptions options,
        string? databasePath,
        CancellationToken cancellationToken)
    {
        var installation = AccessProcess.Describe();

        VbaDoctorDatabase? database = null;
        var probePath = databasePath;
        if (string.IsNullOrWhiteSpace(probePath) && !string.IsNullOrWhiteSpace(options.DefaultDatabasePath))
        {
            probePath = options.DefaultDatabasePath;
        }

        if (!string.IsNullOrWhiteSpace(probePath))
        {
            try
            {
                var full = AccessDbPathPolicy.NormalizeComDatabasePath(probePath, options);
                database = new VbaDoctorDatabase(
                    full,
                    Path.GetExtension(full),
                    AccessDbPathPolicy.GetDatabaseKind(full),
                    File.Exists(full));
            }
            catch (Exception ex)
            {
                database = new VbaDoctorDatabase(
                    Path.GetFullPath(Environment.ExpandEnvironmentVariables(probePath)),
                    Path.GetExtension(probePath),
                    "unknown",
                    false);
                _ = ex;
            }
        }

        var vbe = options.EnableComAutomation
            ? await ProbeVbeAsync(options, database?.Exists == true ? database.Path : null, cancellationToken).ConfigureAwait(false)
            : new VbaDoctorVbe(false, "EnableComAutomation is false.", null, null, false);

        var ace = ProbeAce(options.OleDbProvider);
        var scratchRoot = AccessDbPathPolicy.NormalizeDirectory(
            options.VbaScratchDirectory,
            Path.Combine(Path.GetTempPath(), "AccessMCP.VbaTest"));
        var cacheRoot = AccessDbPathPolicy.NormalizeDirectory(
            options.VbaCacheDirectory,
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "AccessMCP", "VbaTest", "cache"));

        var piaAvailable = IsPiaAvailable();
        var isRuntime = installation.IsRuntime && !vbe.Accessible;
        var canEdit = installation.Installed && !isRuntime && vbe.Accessible && options.EnableComAutomation;
        var canRun = installation.Installed && options.EnableVbaTestTools;

        return new VbaDoctorInfo(
            new VbaDoctorAccess(
                installation.Installed,
                installation.ExecutablePath,
                installation.Version,
                installation.Build,
                installation.Bitness,
                isRuntime),
            new VbaDoctorHost(Environment.Is64BitProcess, System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription),
            database,
            vbe,
            new VbaDoctorAddin(VbeAddinClient.ProgId, IsAddinRegistered(), vbe.AddinLoadable),
            new VbaDoctorErrorTable(VbeCompileErrorTable.Count, VbeCompileErrorTable.SourceHash),
            ace,
            ProbeStorage(scratchRoot),
            ProbeStorage(cacheRoot),
            new VbaDoctorCapabilities(
                canEdit,
                canEdit,
                canRun,
                installation.Installed && !isRuntime && piaAvailable,
                piaAvailable,
                options.VbaMdeBuildStrategy));
    }

    private static async Task<VbaDoctorVbe> ProbeVbeAsync(AccessMcpOptions options, string? databasePath, CancellationToken cancellationToken)
    {
        var com = new AccessComAutomation(options);
        try
        {
            var outcome = await com.RunInAccessWithWatchdogAsync(
                databasePath,
                app =>
                {
                    dynamic dynamicApp = app;
                    object? vbe = null;
                    object? projects = null;
                    try
                    {
                        vbe = dynamicApp.VBE;

                        var addinLoadable = false;
                        try
                        {
                            using var addin = VbeAddinClient.TryConnect(vbe);
                            addinLoadable = addin is not null;
                        }
                        catch
                        {
                        }

                        projects = AccessComAutomation.GetProperty(vbe!, "VBProjects");
                        var count = AccessComAutomation.GetCount(projects!);
                        if (count == 0)
                        {
                            return ((string?)null, (int?)null, addinLoadable);
                        }

                        object? project = null;
                        object? components = null;
                        try
                        {
                            project = AccessComAutomation.GetItem(projects!, 1);
                            var name = Convert.ToString(AccessComAutomation.GetProperty(project!, "Name")) ?? string.Empty;
                            components = AccessComAutomation.GetProperty(project!, "VBComponents");
                            return (name, (int?)AccessComAutomation.GetCount(components!), addinLoadable);
                        }
                        finally
                        {
                            AccessComAutomation.ReleaseComObject(components);
                            AccessComAutomation.ReleaseComObject(project);
                        }
                    }
                    finally
                    {
                        AccessComAutomation.ReleaseComObject(projects);
                        AccessComAutomation.ReleaseComObject(vbe);
                    }
                },
                new AccessRunOptions
                {
                    Timeout = TimeSpan.FromSeconds(20),
                    OpenDatabase = databasePath is not null,
                    SuppressStartupForm = true
                },
                cancellationToken).ConfigureAwait(false);

            if (outcome.Ok)
            {
                return new VbaDoctorVbe(true, null, outcome.Value.Item1, outcome.Value.Item2, outcome.Value.Item3);
            }

            return new VbaDoctorVbe(false, outcome.Exception?.Message ?? "VBE probing failed.", null, null, false);
        }
        catch (Exception ex)
        {
            return new VbaDoctorVbe(false, ex.Message, null, null, false);
        }
    }

    private static bool IsAddinRegistered()
    {
        foreach (var view in new[] { RegistryView.Registry32, RegistryView.Registry64 })
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, view);
                using var key = baseKey.OpenSubKey($@"Software\Microsoft\VBA\VBE\6.0\Addins\{VbeAddinClient.ProgId}");
                if (key is not null)
                {
                    return true;
                }
            }
            catch
            {
            }
        }

        return false;
    }

    private static VbaDoctorAce ProbeAce(string provider)
    {
        var (x64, x86) = (false, false);

        foreach (var view in new[] { RegistryView.Registry64, RegistryView.Registry32 })
        {
            try
            {
                using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view);
                using var key = baseKey.OpenSubKey($@"SOFTWARE\Classes\{provider}");
                var found = key is not null;
                if (view == RegistryView.Registry64)
                {
                    x64 = found;
                }
                else
                {
                    x86 = found;
                }
            }
            catch
            {
            }
        }

        var hostBitness = Environment.Is64BitProcess ? "x64" : "x86";
        var available = Environment.Is64BitProcess ? x64 : x86;

        return new VbaDoctorAce(
            provider,
            available,
            x64 && x86 ? "both" : x64 ? "x64" : x86 ? "x86" : null);
    }

    private static VbaDoctorStorage ProbeStorage(string root)
    {
        try
        {
            Directory.CreateDirectory(root);
            var probe = Path.Combine(root, $".doctor-{Guid.NewGuid():N}.tmp");
            File.WriteAllText(probe, "probe");
            File.Delete(probe);

            var drive = new DriveInfo(Path.GetPathRoot(root) ?? root);
            return new VbaDoctorStorage(root, true, drive.AvailableFreeSpace, null);
        }
        catch (Exception ex)
        {
            return new VbaDoctorStorage(root, false, 0, ex.Message);
        }
    }

    private static bool IsPiaAvailable()
    {
        try
        {
            var directory = AppContext.BaseDirectory;
            if (File.Exists(Path.Combine(directory, "office.dll")) &&
                File.Exists(Path.Combine(directory, "Microsoft.Office.Interop.Access.dll")))
            {
                return true;
            }

            var access = Type.GetType("Microsoft.Office.Interop.Access.Application, Microsoft.Office.Interop.Access");
            var office = Type.GetType("Microsoft.Office.Core.MsoAutomationSecurity, office");
            return access is not null && office is not null;
        }
        catch
        {
            return false;
        }
    }
}
