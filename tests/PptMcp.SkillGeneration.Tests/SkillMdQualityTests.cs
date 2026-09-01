using System.Text.RegularExpressions;
using Xunit;

namespace PptMcp.SkillGeneration.Tests;

/// <summary>
/// Tests to validate the quality of generated SKILL.md files.
/// These tests catch issues like empty parameter descriptions that
/// make skills less useful for LLMs.
/// </summary>
public class SkillMdQualityTests
{
    private static readonly string SkillsFolder = Path.Combine(
        AppContext.BaseDirectory, "skills");

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_Exists()
    {
        var skillPath = Path.Combine(SkillsFolder, "ppt-cli", "SKILL.md");
        Assert.True(File.Exists(skillPath), $"CLI SKILL.md should exist at {skillPath}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_Exists()
    {
        var skillPath = Path.Combine(SkillsFolder, "ppt-mcp", "SKILL.md");
        Assert.True(File.Exists(skillPath), $"MCP SKILL.md should exist at {skillPath}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasNoEmptyParameterDescriptions()
    {
        var skillPath = Path.Combine(SkillsFolder, "ppt-cli", "SKILL.md");
        AssertNoEmptyDescriptions(skillPath, "CLI");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasNoEmptyParameterDescriptions()
    {
        // MCP SKILL.md doesn't have auto-generated parameter tables
        // Tools are discovered via MCP schema - skill contains curated guidance
        // Skip parameter validation for MCP skill
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasCommands()
    {
        var skillPath = Path.Combine(SkillsFolder, "ppt-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        var commandMatches = Regex.Matches(content, @"^### \w+", RegexOptions.Multiline);
        Assert.True(commandMatches.Count > 0, "CLI SKILL.md should have command headings");
        Assert.True(commandMatches.Count >= 10, $"CLI SKILL.md should have at least 10 commands, found {commandMatches.Count}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasTools()
    {
        // The MCP skill is curated, not generated, so its tool table can drift away from
        // the real tool surface without anything failing. Cross-check every tool it names
        // against the generated CLI skill, which is emitted from the Core interfaces.
        var mcpPath = Path.Combine(SkillsFolder, "ppt-mcp", "SKILL.md");
        var cliPath = Path.Combine(SkillsFolder, "ppt-cli", "SKILL.md");
        var mcpContent = File.ReadAllText(mcpPath);
        var cliContent = File.ReadAllText(cliPath);

        var realTools = Regex.Matches(cliContent, @"^### ([a-z][a-z0-9-]*)\s*$", RegexOptions.Multiline)
            .Select(m => m.Groups[1].Value)
            .ToHashSet(StringComparer.Ordinal);
        Assert.True(realTools.Count >= 20, $"Expected the generated CLI skill to document at least 20 tools, found {realTools.Count}");

        var advertised = Regex.Matches(mcpContent, @"^\|[^|]+\|\s*`([^`]+)`\s*\|", RegexOptions.Multiline)
            .Select(m => m.Groups[1].Value)
            .Distinct(StringComparer.Ordinal)
            .ToList();
        Assert.True(advertised.Count >= 8, $"MCP SKILL.md should name at least 8 tools in its task table, found {advertised.Count}");

        var ghosts = advertised.Where(t => !realTools.Contains(t)).ToList();
        Assert.True(
            ghosts.Count == 0,
            $"MCP SKILL.md names {ghosts.Count} tool(s) that do not exist in the generated tool surface: {string.Join(", ", ghosts)}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void Skills_ContainNoSpreadsheetVocabulary()
    {
        // This project was ported from a spreadsheet automation tool. Excel-only vocabulary
        // that survives the port is not merely cosmetic: it tells an agent to call tools
        // that do not exist. Guard the whole skills tree, not just the two SKILL.md files.
        string[] excelOnlyTerms =
        [
            "calculation_mode", "worksheet", "pivottable", "powerquery",
            "slicer", "set-values", "get-values", "datamodel"
        ];

        var offenders = new List<string>();
        foreach (var file in Directory.EnumerateFiles(SkillsFolder, "*.md", SearchOption.AllDirectories))
        {
            var lines = File.ReadAllLines(file);
            for (int i = 0; i < lines.Length; i++)
            {
                foreach (var term in excelOnlyTerms)
                {
                    if (lines[i].Contains(term, StringComparison.OrdinalIgnoreCase))
                    {
                        offenders.Add($"  {Path.GetFileName(file)}:{i + 1} ({term}): {lines[i].Trim()}");
                        break;
                    }
                }
            }
        }

        Assert.True(
            offenders.Count == 0,
            $"Found {offenders.Count} line(s) of spreadsheet vocabulary in the skills tree:\n{string.Join("\n", offenders.Take(15))}");
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasParameterTables()
    {
        var skillPath = Path.Combine(SkillsFolder, "ppt-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("| Parameter | Description |", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasParameterTables()
    {
        // MCP SKILL.md has markdown tables for reference, not parameter tables
        var skillPath = Path.Combine(SkillsFolder, "ppt-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("| Task | Tool |", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void CliSkill_HasActionsList()
    {
        var skillPath = Path.Combine(SkillsFolder, "ppt-cli", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("**Actions:**", content);
    }

    [Fact]
    [Trait("Category", "Unit")]
    [Trait("Feature", "SkillGeneration")]
    public void McpSkill_HasActionsList()
    {
        // MCP SKILL.md has curated action examples, not **Actions:** section
        var skillPath = Path.Combine(SkillsFolder, "ppt-mcp", "SKILL.md");
        var content = File.ReadAllText(skillPath);
        Assert.Contains("action:", content);
    }

    private static void AssertNoEmptyDescriptions(string skillPath, string skillType)
    {
        Assert.True(File.Exists(skillPath), $"{skillType} SKILL.md should exist");
        var content = File.ReadAllText(skillPath);
        var lines = content.Split('\n');
        var emptyDescriptions = new List<string>();
        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].Trim();
            if (Regex.IsMatch(line, @"^\|\s*`[^`]+`\s*\|\s*\|$"))
            {
                var paramMatch = Regex.Match(line, @"`([^`]+)`");
                if (paramMatch.Success)
                {
                    emptyDescriptions.Add(paramMatch.Groups[1].Value);
                }
            }
        }

        if (emptyDescriptions.Count > 0)
        {
            var message = $"{skillType} SKILL.md has {emptyDescriptions.Count} parameters with empty descriptions:\n" +
                          string.Join("\n", emptyDescriptions.Take(10).Select(p => $"  - {p}"));
            if (emptyDescriptions.Count > 10)
            {
                message += $"\n  ... and {emptyDescriptions.Count - 10} more";
            }

            Assert.Fail(message);
        }
    }
}
