// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace PptMcp.Core.Diagnostics;

/// <summary>
/// Raised when a VBA operation fails because "Trust access to the VBA project object
/// model" is disabled.
///
/// Carries the remediation in its message so the text reaches the caller through the
/// normal failure path. Both entry points surface an exception's message: the CLI
/// prints it, and <c>batch.Execute()</c> turns it into
/// <c>OperationResult { Success = false, ErrorMessage }</c> for MCP. Nothing extra is
/// needed at either layer for the advice to arrive.
///
/// Note deliberately NOT thrown from a bare <c>catch (Exception)</c>. It is raised
/// only after <see cref="VbaTrustProbe"/> confirms the setting is actually disabled,
/// so an unrelated COM failure is never mislabelled as a trust problem.
/// </summary>
[SupportedOSPlatform("windows")]
public sealed class VbaTrustException : InvalidOperationException
{
    /// <summary>Creates the exception with the standard remediation text.</summary>
    /// <param name="inner">The original COM failure, preserved for diagnostics.</param>
    public VbaTrustException(Exception? inner)
        : base(BuildMessage(inner), inner)
    {
    }

    /// <summary>Creates the exception with the standard remediation text.</summary>
    public VbaTrustException()
        : base(BuildMessage(null))
    {
    }

    /// <summary>Creates the exception with a caller-supplied message.</summary>
    /// <param name="message">Message to use instead of the standard remediation.</param>
    public VbaTrustException(string message)
        : base(message)
    {
    }

    /// <summary>Creates the exception with a caller-supplied message and inner exception.</summary>
    /// <param name="message">Message to use instead of the standard remediation.</param>
    /// <param name="innerException">The original failure.</param>
    public VbaTrustException(string message, Exception innerException)
        : base(message, innerException)
    {
    }

    /// <summary>The registry state observed when the failure was diagnosed.</summary>
    public string RegistryPath { get; } = VbaTrustProbe.DescribeRegistryPath();

    private static string BuildMessage(Exception? inner)
    {
        var original = inner is COMException com
            ? $" (original COM error 0x{com.HResult:X8}: {com.Message})"
            : inner is not null ? $" (original error: {inner.Message})" : string.Empty;

        return $"{VbaTrustProbe.RemediationText} Observed: {VbaTrustProbe.DescribeRegistryPath()}.{original}";
    }
}
