# PptMcp.ComInterop

**Low-level COM interop utilities for PowerPoint automation.**

## Overview

This library provides PowerPoint-specific COM object lifecycle management and OLE message filtering. It's the foundation layer for PptMcp projects, handling STA threading, session management, and batch operations specifically for PowerPoint COM automation.

## Features

- **STA Threading Management** - Ensures proper single-threaded apartment model for PowerPoint COM objects
- **COM Object Lifecycle** - Automatic COM object cleanup and garbage collection
- **OLE Message Filtering** - Handles busy/rejected COM calls with retry logic using Polly
- **PowerPoint Session Management** - Manages PowerPoint.Application lifecycle safely
- **Batch Operations** - Efficient handling of multiple PowerPoint operations in a single session

## Usage Example

```csharp
using PptMcp.ComInterop;

// Use PptSession for safe PowerPoint automation
await using var session = await PptSession.BeginAsync("path/to/presentation.pptx");
await using var batch = await session.BeginBatchAsync();

await batch.ExecuteAsync<int>(async (ctx, ct) => 
{
    // Access PowerPoint presentation through ctx.Book
    dynamic slides = ctx.Book.Slides;
    dynamic slide = slides.Item(1);
    
    // Perform PowerPoint operations
    slide.Name = "UpdatedSlide";
    
    return 0;
});

await batch.Save();
```

## Key Classes

- **PptSession** - Manages PowerPoint.Application lifecycle and presentation operations
- **PptBatch** - Groups multiple operations for efficient execution
- **ComUtilities** - Helper methods for COM object cleanup and safe property access
- **OleMessageFilter** - Implements retry logic for busy PowerPoint instances

## Requirements

- Windows OS
- .NET 10.0 or later
- Microsoft PowerPoint 2016+ installed

## Platform Support

- ✅ Windows x64
- ✅ Windows ARM64
- ❌ Linux (PowerPoint COM not available)
- ❌ macOS (PowerPoint COM not available)

