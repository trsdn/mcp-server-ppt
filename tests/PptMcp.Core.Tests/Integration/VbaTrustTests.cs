// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.InteropServices;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Commands.Vba;
using PptMcp.Core.Diagnostics;
using PptMcp.Core.Tests.Helpers;
using Xunit;

namespace PptMcp.Core.Tests.Integration;

/// <summary>
/// Covers the VBA trust precondition (issue #130).
///
/// "Trust access to the VBA project object model" is disabled on a default Windows
/// install, so every VBA operation fails on first use with a raw COMException whose
/// message names no cause and no fix. These tests cover the diagnostic path that
/// turns that into actionable text.
///
/// Deliberately does NOT toggle the registry value. Doing so would mutate machine
/// state shared with the developer's own PowerPoint, and - as issue #130 documents -
/// the value is silently reverted to 0 if PowerPoint is running when it is written.
/// The disabled branch is covered by asserting on the remediation text itself, which
/// is a pure function of the probe result.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "Core")]
[Trait("Feature", "VbaTrust")]
[Trait("RequiresPowerPoint", "true")]
public sealed class VbaTrustTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;
    private readonly VbaCommands _commands = new();

    public VbaTrustTests(TempDirectoryFixture fixture)
    {
        _fixture = fixture;
    }

    [Fact]
    public void CheckTrust_ReportsTheLiveRegistryState()
    {
        var filePath = _fixture.CreateTestFile(extension: ".pptm");

        using var batch = PptSession.BeginBatch(filePath);
        var result = _commands.CheckTrust(batch);

        Assert.True(result.Success);

        // The probe must agree with the registry it claims to read, whichever way the
        // machine is configured. Asserting a fixed value would make this test pass or
        // fail on the developer's Trust Center setting rather than on the code.
        Assert.Equal(VbaTrustProbe.IsTrustEnabled(), result.TrustEnabled);

        Assert.False(string.IsNullOrWhiteSpace(result.RegistryPath));
        Assert.Contains("AccessVBOM", result.RegistryPath, StringComparison.Ordinal);
    }

    [Fact]
    public void CheckTrust_WhenTrustDisabled_CarriesRemediation()
    {
        var filePath = _fixture.CreateTestFile(extension: ".pptm");

        using var batch = PptSession.BeginBatch(filePath);
        var result = _commands.CheckTrust(batch);

        if (result.TrustEnabled)
        {
            // Trust is on, so there is nothing to remediate and the field must be empty
            // rather than carrying stale advice.
            Assert.True(string.IsNullOrEmpty(result.Remediation));
            return;
        }

        Assert.False(string.IsNullOrWhiteSpace(result.Remediation));
    }

    [Fact]
    public void VbaOperation_WhenTrustDisabled_SurfacesRemediationNotRawComException()
    {
        if (VbaTrustProbe.IsTrustEnabled())
        {
            // Cannot exercise the failure path without disabling trust machine-wide,
            // which this suite refuses to do. The remediation text itself is asserted
            // below in RemediationText_NamesRegistryPathValueAndRestart.
            return;
        }

        var filePath = _fixture.CreateTestFile(extension: ".pptm");

        using var batch = PptSession.BeginBatch(filePath);

        var ex = Assert.ThrowsAny<Exception>(() => _commands.List(batch));

        // The whole point of #130: the caller must not receive a bare COMException.
        Assert.IsNotType<COMException>(ex);
        Assert.Contains("AccessVBOM", ex.Message, StringComparison.Ordinal);
        Assert.Contains("restart", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RemediationText_NamesRegistryPathValueAndRestart()
    {
        var text = VbaTrustProbe.RemediationText;

        // An agent that hits this error has only this string to act on. Each element
        // below is a step it cannot infer: the key, the value, and the fact that
        // PowerPoint reads the setting only at process start.
        Assert.Contains("AccessVBOM", text, StringComparison.Ordinal);
        Assert.Contains("PowerPoint\\Security", text, StringComparison.Ordinal);
        Assert.Contains("1", text, StringComparison.Ordinal);
        Assert.Contains("restart", text, StringComparison.OrdinalIgnoreCase);

        // Setting the value while PowerPoint is running is silently reverted on exit,
        // so the order of operations has to be stated or the fix appears not to work.
        Assert.Contains("running", text, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void IsTrustEnabled_DoesNotThrow_RegardlessOfOfficeLayout()
    {
        // Office version keys vary by install, and a stray key such as "8.0" with no
        // PowerPoint\Security subkey exists on real machines. The probe must tolerate
        // that rather than throw, since it runs inside an error path.
        var ex = Record.Exception(() => VbaTrustProbe.IsTrustEnabled());
        Assert.Null(ex);
    }
}
