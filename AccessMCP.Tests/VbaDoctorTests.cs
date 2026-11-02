using DigitalJeeves.MsAccessMCP.Core;

namespace DigitalJeeves.MsAccessMCP.Tests;

/// <summary>
/// The doctor's capability flags must match the gates the tools check, whatever this machine has
/// installed, so these drive <see cref="VbaDoctor.Capabilities"/> directly.
/// </summary>
public sealed class VbaDoctorTests
{
    private static readonly AccessMcpOptions Everything = new()
    {
        EnableComAutomation = true,
        EnableVbaTestTools = true
    };

    [Fact]
    public void FullyEnabledMachineCanDoEverything()
    {
        var capabilities = Probe(Everything);

        Assert.True(capabilities.CanEditModules);
        Assert.True(capabilities.CanCompile);
        Assert.True(capabilities.CanRun);
        Assert.True(capabilities.CanBuildMde);
    }

    [Fact]
    public void ComAutomationOffDisablesEverything()
    {
        var capabilities = Probe(Everything with { EnableComAutomation = false });

        Assert.False(capabilities.CanEditModules);
        Assert.False(capabilities.CanCompile);
        Assert.False(capabilities.CanRun);
        Assert.False(capabilities.CanBuildMde);
    }

    [Fact]
    public void VbaTestToolsOffStillAllowsEditingButNotRunning()
    {
        var capabilities = Probe(Everything with { EnableVbaTestTools = false });

        Assert.True(capabilities.CanEditModules);
        Assert.True(capabilities.CanCompile);
        Assert.False(capabilities.CanRun);
        Assert.False(capabilities.CanBuildMde);
    }

    [Fact]
    public void ReadOnlyBlocksEditsButNotScratchRuns()
    {
        var capabilities = Probe(Everything with { ReadOnly = true });

        Assert.False(capabilities.CanEditModules);
        Assert.False(capabilities.CanCompile);
        Assert.True(capabilities.CanRun);
        Assert.True(capabilities.CanBuildMde);
    }

    [Fact]
    public void RuntimeOrNoVbeTrustBlocksEditing()
    {
        Assert.False(Probe(Everything, isRuntime: true).CanEditModules);
        Assert.False(Probe(Everything, isRuntime: true).CanBuildMde);
        Assert.True(Probe(Everything, isRuntime: true).CanRun);
        Assert.False(Probe(Everything, vbeAccessible: false).CanEditModules);
    }

    [Fact]
    public void NoAccessOrNoPiaIsReported()
    {
        var noAccess = Probe(Everything, accessInstalled: false);
        var noPia = Probe(Everything, piaAvailable: false);

        Assert.False(noAccess.CanRun);
        Assert.False(noAccess.CanEditModules);
        Assert.False(noPia.CanBuildMde);
        Assert.False(noPia.PiaAvailable);
        Assert.True(noPia.CanRun);
    }

    [Fact]
    public void ReportsTheConfiguredMdeStrategy()
    {
        Assert.Equal("prebuilt", Probe(Everything with { VbaMdeBuildStrategy = "prebuilt" }).MdeStrategy);
    }

    private static VbaDoctorCapabilities Probe(
        AccessMcpOptions options,
        bool accessInstalled = true,
        bool isRuntime = false,
        bool vbeAccessible = true,
        bool piaAvailable = true)
        => VbaDoctor.Capabilities(options, accessInstalled, isRuntime, vbeAccessible, piaAvailable);
}
