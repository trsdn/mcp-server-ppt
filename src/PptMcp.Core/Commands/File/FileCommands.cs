using PptMcp.Core.Models;

namespace PptMcp.Core.Commands.File;

public class FileCommands : IFileCommands
{
    public FileValidationInfo Test(string filePath)
    {
        string fullPath = Path.GetFullPath(filePath);

        if (!System.IO.File.Exists(fullPath))
        {
            return new FileValidationInfo
            {
                Success = false,
                Exists = false,
                FilePath = fullPath
            };
        }

        var fileInfo = new FileInfo(fullPath);
        string ext = fileInfo.Extension.ToLowerInvariant();

        return new FileValidationInfo
        {
            Success = true,
            Exists = true,
            FilePath = fullPath,
            FileName = fileInfo.Name,
            FileSizeBytes = fileInfo.Length,
            IsReadOnly = fileInfo.IsReadOnly,
            IsMacroEnabled = ext == ".pptm",
            SlideCount = -1 // Requires opening the file; set by caller if needed
        };
    }
}
