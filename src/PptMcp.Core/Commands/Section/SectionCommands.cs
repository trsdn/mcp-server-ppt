using PptMcp.ComInterop;
using PptMcp.ComInterop.Session;
using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.Section;

public class SectionCommands : ISectionCommands
{
    public SectionListResult List(IPptBatch batch)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic sections = pres.SectionProperties;
            try
            {
                int count = Convert.ToInt32(sections.Count);

                var result = new SectionListResult
                {
                    Success = true,
                    FilePath = ctx.PresentationPath
                };

                for (int i = 1; i <= count; i++)
                {
                    string name = sections.Name(i)?.ToString() ?? $"Section {i}";
                    int firstSlide = Convert.ToInt32(sections.FirstSlide(i));
                    int slideCount = Convert.ToInt32(sections.SlidesCount(i));

                    result.Sections.Add(new SectionInfo
                    {
                        Index = i,
                        Name = name,
                        FirstSlideIndex = firstSlide,
                        SlideCount = slideCount
                    });
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref sections!);
            }
        });
    }

    public OperationResult Add(IPptBatch batch, string sectionName, int slideIndex)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sectionName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic sections = pres.SectionProperties;
            try
            {
                // Find the section that contains slideIndex (or use count+1 to append)
                int count = Convert.ToInt32(sections.Count);
                int insertAtSection = count + 1; // default: append after last section

                // Find which section slideIndex falls into, so we insert before that section
                for (int i = 1; i <= count; i++)
                {
                    int firstSlide = Convert.ToInt32(sections.FirstSlide(i));
                    if (firstSlide >= slideIndex)
                    {
                        insertAtSection = i;
                        break;
                    }
                }

                // AddSection(sectionIndex, name) — inserts a new section at sectionIndex
                int newIndex = Convert.ToInt32(sections.AddSection(insertAtSection, sectionName));

                return new OperationResult
                {
                    Success = true,
                    Action = "add",
                    Message = $"Added section '{sectionName}' at index {newIndex}",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref sections!);
            }
        });
    }

    public OperationResult Rename(IPptBatch batch, int sectionIndex, string newName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(newName);

        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic sections = pres.SectionProperties;
            try
            {
                sections.Rename(sectionIndex, newName);
                return new OperationResult
                {
                    Success = true,
                    Action = "rename",
                    Message = $"Renamed section {sectionIndex} to '{newName}'",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref sections!);
            }
        });
    }

    public OperationResult Delete(IPptBatch batch, int sectionIndex)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic pres = (dynamic)ctx.Presentation;
            dynamic sections = pres.SectionProperties;
            try
            {
                // Delete(sectionIndex, deleteSlides=false)
                sections.Delete(sectionIndex, false);
                return new OperationResult
                {
                    Success = true,
                    Action = "delete",
                    Message = $"Deleted section {sectionIndex} (slides preserved)",
                    FilePath = ctx.PresentationPath
                };
            }
            finally
            {
                ComUtilities.Release(ref sections!);
            }
        });
    }
}
