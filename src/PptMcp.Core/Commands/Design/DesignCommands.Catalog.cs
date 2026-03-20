using PptMcp.ComInterop.Session;
using PptMcp.Core.Data;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Design;

public partial class DesignCommands
{
    // ── Archetypes ─────────────────────────────────────

    public ArchetypeListResult ListArchetypes(IPptBatch batch)
    {
        var result = new ArchetypeListResult { Success = true };
        var coreArchetypes = DesignCatalogProvider.GetArchetypes();
        var countsByArchetype = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        ReferenceSubArchetypeCatalog? referenceCatalog = null;

        if (TryLoadReferenceCatalog(out var manifest, out var catalog, out _))
        {
            countsByArchetype = GetCountsByArchetype(manifest);
            referenceCatalog = catalog;
        }

        var emittedIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var archetype in coreArchetypes)
        {
            emittedIds.Add(archetype.Id);
            var topLevel = referenceCatalog is null ? null : GetReferenceTopLevel(referenceCatalog, archetype.Id);
            result.Archetypes.Add(new ArchetypeListItem
            {
                Id = archetype.Id,
                Name = archetype.Name,
                When = archetype.When,
                BestDensity = archetype.BestDensity,
                Variants = archetype.Variants,
                ExampleTitle = archetype.ExampleTitle,
                HasCuratedLayoutGuidance = true,
                ObservedSlideCount = countsByArchetype.GetValueOrDefault(archetype.Id),
                ObservedSubtypeCount = topLevel?.Subtypes.Count(subtype => subtype.Count > 0) ?? 0,
                ObservedExampleSlides = topLevel is null ? [] : GetExampleSlides(topLevel, 5)
            });
        }

        if (referenceCatalog != null)
        {
            foreach (var topLevel in referenceCatalog.TopLevels.Where(entry => !emittedIds.Contains(entry.ArchetypeId)))
            {
                var metadata = GetLearnedArchetypeMetadata(topLevel.ArchetypeId);
                result.Archetypes.Add(new ArchetypeListItem
                {
                    Id = topLevel.ArchetypeId,
                    Name = metadata.Name,
                    When = metadata.Summary,
                    HasCuratedLayoutGuidance = false,
                    ObservedSlideCount = countsByArchetype.GetValueOrDefault(topLevel.ArchetypeId),
                    ObservedSubtypeCount = topLevel.Subtypes.Count(subtype => subtype.Count > 0),
                    ObservedExampleSlides = GetExampleSlides(topLevel, 5)
                });
            }
        }

        return result;
    }

    public ArchetypeDetailResult GetArchetype(IPptBatch batch, string archetypeId)
    {
        List<ReferenceManifestEntry> manifest = [];
        var coreEntry = DesignCatalogProvider.GetArchetypes()
            .Find(archetype => string.Equals(archetype.Id, archetypeId, StringComparison.OrdinalIgnoreCase));
        var countsByArchetype = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        ReferenceSubArchetypeCatalog? referenceCatalog = null;
        ReferenceTopLevelEntry? topLevel = null;

        if (TryLoadReferenceCatalog(out manifest, out var catalog, out _))
        {
            countsByArchetype = GetCountsByArchetype(manifest);
            referenceCatalog = catalog;
            topLevel = GetReferenceTopLevel(catalog, archetypeId);
        }

        if (coreEntry == null && topLevel == null)
        {
            return new ArchetypeDetailResult
            {
                Success = false,
                ErrorMessage = $"Archetype '{archetypeId}' not found. Use 'list-archetypes' to see available archetypes."
            };
        }

        var metadata = coreEntry == null
            ? GetLearnedArchetypeMetadata(archetypeId)
            : (coreEntry.Name, coreEntry.When);
        var detail = BuildUnifiedArchetypeDetail(archetypeId, coreEntry != null ? DesignCatalogProvider.GetArchetypeDetail(archetypeId) : null);
        return new ArchetypeDetailResult
        {
            Success = true,
            Id = archetypeId,
            Name = metadata.Item1,
            When = metadata.Item2,
            BestDensity = coreEntry?.BestDensity ?? [],
            Variants = coreEntry?.Variants ?? [],
            HasCuratedLayoutGuidance = coreEntry != null,
            ObservedSlideCount = countsByArchetype.GetValueOrDefault(archetypeId),
            ObservedExampleSlides = topLevel is null ? [] : GetExampleSlides(topLevel, 10),
            ObservedExamples = topLevel is null ? [] : GetObservedExamples(manifest, topLevel, 10),
            ObservedSubtypes = topLevel is null ? [] : BuildSubtypeInfos(topLevel, manifest),
            AuditSamples = referenceCatalog is null ? [] : GetMisbucketedSamples(referenceCatalog, archetypeId),
            Detail = detail
        };
    }

    // ── Color Palettes ─────────────────────────────────

    public PaletteListResult ListPalettes(IPptBatch batch)
    {
        var palettes = DesignCatalogProvider.GetPalettes();
        var result = new PaletteListResult { Success = true };
        foreach (var p in palettes)
        {
            result.Palettes.Add(new PaletteListItem
            {
                Id = p.Id,
                Name = p.Name,
                BestFor = p.BestFor
            });
        }
        return result;
    }

    public PaletteDetailResult GetPalette(IPptBatch batch, string paletteId)
    {
        var entry = DesignCatalogProvider.GetPalette(paletteId);
        if (entry == null)
        {
            return new PaletteDetailResult
            {
                Success = false,
                ErrorMessage = $"Palette '{paletteId}' not found. Use 'list-palettes' to see available palettes."
            };
        }

        return new PaletteDetailResult
        {
            Success = true,
            Id = entry.Id,
            Name = entry.Name,
            BestFor = entry.BestFor,
            Colors = entry.Colors
        };
    }

    // ── Style Profiles ─────────────────────────────────

    public StyleProfileListResult ListStyleProfiles(IPptBatch batch)
    {
        var profiles = DesignCatalogProvider.GetStyleProfiles();
        var result = new StyleProfileListResult { Success = true };
        foreach (var p in profiles)
        {
            result.Profiles.Add(new StyleProfileListItem
            {
                Id = p.Id,
                Name = p.Name,
                BestFor = p.BestFor,
                ColorScheme = p.ColorScheme
            });
        }
        return result;
    }

    public StyleProfileDetailResult GetStyleProfile(IPptBatch batch, string profileId)
    {
        var entry = DesignCatalogProvider.GetStyleProfile(profileId);
        if (entry == null)
        {
            return new StyleProfileDetailResult
            {
                Success = false,
                ErrorMessage = $"Style profile '{profileId}' not found. Use 'list-style-profiles' to see available profiles."
            };
        }

        return new StyleProfileDetailResult
        {
            Success = true,
            Id = entry.Id,
            Name = entry.Name,
            Description = entry.Description,
            BestFor = entry.BestFor,
            ColorScheme = entry.ColorScheme,
            Font = entry.Font,
            TitleStyle = entry.TitleStyle,
            TitleSize = entry.TitleSize,
            BodySize = entry.BodySize,
            FootnoteSize = entry.FootnoteSize,
            BulletsPerSlide = entry.BulletsPerSlide,
            WordsPerBullet = entry.WordsPerBullet,
            ContentDensity = entry.ContentDensity,
            PreferredArchetypes = entry.PreferredArchetypes,
            Whitespace = entry.Whitespace,
            Background = entry.Background,
            ChartStyle = entry.ChartStyle,
            SpecialRules = entry.SpecialRules
        };
    }

    // ── Layout Grids ───────────────────────────────────

    public LayoutGridListResult ListLayoutGrids(IPptBatch batch)
    {
        var grids = DesignCatalogProvider.GetLayoutGridData();
        var result = new LayoutGridListResult { Success = true };
        foreach (var g in grids.Grids)
        {
            result.Grids.Add(new LayoutGridListItem
            {
                Id = g.Id,
                Name = g.Name,
                BestFor = g.BestFor
            });
        }
        return result;
    }

    public LayoutGridResult GetLayoutGrid(IPptBatch batch, string gridId)
    {
        var entry = DesignCatalogProvider.GetLayoutGrid(gridId);
        if (entry == null)
        {
            return new LayoutGridResult
            {
                Success = false,
                ErrorMessage = $"Layout grid '{gridId}' not found. Use 'list-layout-grids' to see available grids."
            };
        }

        var result = new LayoutGridResult
        {
            Success = true,
            Id = entry.Id,
            Name = entry.Name,
            BestFor = entry.BestFor
        };

        foreach (var z in entry.Zones)
        {
            result.Zones.Add(new LayoutZone
            {
                Name = z.Name,
                X = z.X,
                Y = z.Y,
                W = z.W,
                H = z.H,
                Description = z.Description
            });
        }

        return result;
    }

    // ── Density Profiles ───────────────────────────────

    public DensityProfileListResult ListDensityProfiles(IPptBatch batch)
    {
        var profiles = DesignCatalogProvider.GetDensityProfiles();
        var result = new DensityProfileListResult { Success = true };
        foreach (var p in profiles)
        {
            result.Profiles.Add(new DensityProfileListItem
            {
                Id = p.Id,
                Name = p.Name,
                UsedFor = p.UsedFor
            });
        }
        return result;
    }

    public DensityProfileResult GetDensityProfile(IPptBatch batch, string densityId)
    {
        var entry = DesignCatalogProvider.GetDensityProfile(densityId);
        if (entry == null)
        {
            return new DensityProfileResult
            {
                Success = false,
                ErrorMessage = $"Density profile '{densityId}' not found. Use 'list-density-profiles' to see available profiles."
            };
        }

        return new DensityProfileResult
        {
            Success = true,
            Id = entry.Id,
            Name = entry.Name,
            UsedFor = entry.UsedFor,
            Audience = entry.Audience,
            Mode = entry.Mode,
            TextVolume = entry.TextVolume,
            ElementCount = entry.ElementCount,
            DataGranularity = entry.DataGranularity,
            AnnotationDepth = entry.AnnotationDepth,
            SourceCompleteness = entry.SourceCompleteness,
            WhiteSpaceRatio = entry.WhiteSpaceRatio,
            Character = entry.Character,
            BestArchetypes = entry.BestArchetypes
        };
    }

    // ── Context Model ──────────────────────────────────

    public ContextModelResult GetContextModel(IPptBatch batch)
    {
        var model = DesignCatalogProvider.GetContextModel();
        var result = new ContextModelResult
        {
            Success = true,
            DefaultDensity = model.DensitySelection.Default
        };

        foreach (var m in model.MeetingTypes)
        {
            result.MeetingTypes.Add(new MeetingTypeInfo
            {
                Id = m.Id,
                Name = m.Name,
                Audience = m.Audience,
                TimePerSlide = m.TimePerSlide,
                Goal = m.Goal,
                DecisionPressure = m.DecisionPressure,
                PrimaryMode = m.PrimaryMode,
                SecondaryMode = m.SecondaryMode,
                DefaultDensity = m.DefaultDensity
            });
        }

        foreach (var a in model.AudienceLevels)
        {
            result.AudienceLevels.Add(new AudienceLevelInfo
            {
                Id = a.Id,
                Label = a.Label,
                Roles = a.Roles,
                PreferredDensity = a.PreferredDensity,
                WantsToSee = a.WantsToSee
            });
        }

        foreach (var c in model.ConsumptionModes)
        {
            result.ConsumptionModes.Add(new ConsumptionModeInfo
            {
                Id = c.Id,
                Name = c.Name,
                SpeakerPresent = c.SpeakerPresent,
                SelfContained = c.SelfContained,
                TextDensity = c.TextDensity
            });
        }

        return result;
    }

    // ── Deck Sequences ─────────────────────────────────

    public DeckSequenceDetailResult GetDeckSequence(IPptBatch batch, string sequenceId)
    {
        var entry = DesignCatalogProvider.GetDeckSequence(sequenceId);
        if (entry == null)
        {
            var sequences = DesignCatalogProvider.GetDeckSequences();
            var available = string.Join(", ", sequences.Select(s => $"{s.Id} ({s.Name})"));
            return new DeckSequenceDetailResult
            {
                Success = false,
                ErrorMessage = $"Deck sequence '{sequenceId}' not found. Available: {available}"
            };
        }

        var result = new DeckSequenceDetailResult
        {
            Success = true,
            Id = entry.Id,
            Name = entry.Name,
            UsedFor = entry.UsedFor,
            Intent = entry.Intent
        };

        foreach (var s in entry.Slides)
        {
            result.Slides.Add(new DeckSlideInfo
            {
                Position = s.Position?.ToString() ?? "",
                Purpose = s.Purpose,
                Archetype = s.Archetype,
                Density = s.Density
            });
        }

        return result;
    }

    // ── Slide Patterns & Icon Shapes ───────────────────

    public SlidePatternListResult GetSlidePatterns(IPptBatch batch)
    {
        return new SlidePatternListResult
        {
            Success = true,
            Content = DesignCatalogProvider.GetSlidePatterns()
        };
    }

    public IconShapeListResult GetIconShapes(IPptBatch batch)
    {
        return new IconShapeListResult
        {
            Success = true,
            Content = DesignCatalogProvider.GetIconShapes()
        };
    }
}
