// Copyright (c) 2026 Torsten Mahr. All rights reserved.
// Licensed under the MIT License.

using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using Microsoft.Win32;

namespace PptMcp.Core.Diagnostics;

/// <summary>
/// Determines whether "Trust access to the VBA project object model" is enabled, and
/// supplies the remediation text for when it is not.
///
/// This exists because the precondition is disabled on a default Windows install, so
/// every VBA operation fails on first use with a raw <see cref="COMException"/> whose
/// message names neither the cause nor the fix (issue #130). An agent that receives
/// that message has no path forward and will either retry indefinitely or conclude
/// that VBA is unsupported.
///
/// The probe reads the registry rather than matching on the exception's HRESULT or
/// message. Both alternatives are unreliable: the message is localized, and matching
/// a single HRESULT risks reporting unrelated COM failures as a trust problem. The
/// registry value is the setting itself, so consulting it answers the question
/// directly and in a locale-independent way.
/// </summary>
[SupportedOSPlatform("windows")]
public static class VbaTrustProbe
{
    private const string OfficeRoot = @"Software\Microsoft\Office";
    private const string SecuritySuffix = @"PowerPoint\Security";
    private const string ValueName = "AccessVBOM";

    /// <summary>
    /// Human-readable remediation naming every step the caller cannot infer: the key,
    /// the value, the restart requirement, and the ordering trap.
    ///
    /// The ordering matters and is easy to get wrong. Writing the value while
    /// PowerPoint is running appears to succeed, but PowerPoint rewrites its Security
    /// key on exit and silently reverts it to 0 - so the obvious sequence (set the
    /// value, then restart) leaves the setting exactly as it was.
    /// </summary>
    public static string RemediationText { get; } =
        "VBA operations require 'Trust access to the VBA project object model', which is " +
        "disabled by default on Windows. To enable it: " +
        "(1) close PowerPoint completely - setting the value while PowerPoint is running " +
        "is silently reverted to 0 when it exits; " +
        @"(2) set HKCU\Software\Microsoft\Office\<version>\PowerPoint\Security\AccessVBOM " +
        "(DWORD) to 1, where <version> is your Office version such as 16.0; " +
        "(3) restart PowerPoint - the setting is read only at process start. " +
        "The equivalent UI path is File > Options > Trust Center > Trust Center Settings > " +
        "Macro Settings > Trust access to the VBA project object model.";

    /// <summary>
    /// Returns true when any installed Office version has <c>AccessVBOM</c> set to 1.
    ///
    /// Any version is sufficient because a machine with several Office versions
    /// registered will drive whichever one PowerPoint actually launched, and this
    /// probe has no reliable way to know which that is. Reporting "enabled" when at
    /// least one is enabled is the conservative direction: it never invents a trust
    /// problem that would send the caller to change a setting that is already correct.
    /// </summary>
    public static bool IsTrustEnabled()
    {
        foreach (var entry in ReadAllTrustValues())
        {
            if (entry.Value == 1)
            {
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Returns the registry path this probe consulted, for inclusion in results so the
    /// caller can verify the diagnosis rather than take it on trust.
    /// </summary>
    public static string DescribeRegistryPath()
    {
        var found = ReadAllTrustValues();

        if (found.Count == 0)
        {
            return $@"HKCU\{OfficeRoot}\<version>\{SecuritySuffix}\{ValueName} (no Office version found)";
        }

        var parts = found.Select(f =>
        {
            var shown = f.Value.HasValue
                ? f.Value.Value.ToString(CultureInfo.InvariantCulture)
                : "(not set)";
            return $@"HKCU\{OfficeRoot}\{f.Version}\{SecuritySuffix}\{ValueName}={shown}";
        });

        return string.Join("; ", parts);
    }

    /// <summary>
    /// Enumerates every installed Office version key and reads its AccessVBOM value.
    ///
    /// Tolerates a partial or unusual Office layout by design. Real machines carry
    /// stray version keys - an "8.0" with no PowerPoint subkey is common - and this
    /// method runs inside an error path, where throwing would replace a useful
    /// diagnosis with a second, less informative failure.
    /// </summary>
    private static List<(string Version, int? Value)> ReadAllTrustValues()
    {
        var results = new List<(string Version, int? Value)>();

        try
        {
            using var office = Registry.CurrentUser.OpenSubKey(OfficeRoot);

            if (office is null)
            {
                return results;
            }

            foreach (var versionName in office.GetSubKeyNames())
            {
                // Office version keys are numeric ("16.0"). Skip named subkeys such as
                // "Common" or "ClickToRun", which have no PowerPoint security settings.
                if (!IsVersionKey(versionName))
                {
                    continue;
                }

                using var security = office.OpenSubKey($@"{versionName}\{SecuritySuffix}");

                if (security is null)
                {
                    continue;
                }

                var raw = security.GetValue(ValueName);
                results.Add((versionName, raw is int i ? i : null));
            }
        }
        catch (Exception ex) when (ex is System.Security.SecurityException or UnauthorizedAccessException or IOException)
        {
            // A caller without read access to HKCU cannot be diagnosed, but must still
            // receive the original COM failure rather than a registry exception layered
            // on top of it. Returning empty means "unknown", which callers treat as
            // "not proven enabled".
        }

        return results;
    }

    private static bool IsVersionKey(string name)
    {
        return name.Length > 0 && char.IsDigit(name[0]);
    }
}
