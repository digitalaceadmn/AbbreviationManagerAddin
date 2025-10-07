# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Microsoft Word VSTO (Visual Studio Tools for Office) add-in written in C# that provides intelligent abbreviation management and expansion. The add-in loads abbreviations from an embedded Excel file and offers real-time suggestions as users type.

**Target Framework**: .NET Framework 4.8
**Office Application**: Microsoft Word (Office 15.0+)

## Build and Development Commands

### Building the Project
```bash
# Build in Debug configuration (default)
msbuild AbbreviationWordAddin.sln /p:Configuration=Debug

# Build in Release configuration
msbuild AbbreviationWordAddin.sln /p:Configuration=Release
```

### Debugging
- Open the solution in Visual Studio
- Press F5 to launch Word with the add-in attached
- The add-in loads automatically when Word starts

### Publishing
The project is configured for ClickOnce deployment:
- Default publish location: `D:\add-in\`
- Auto-increment revision: enabled
- Application version format: `1.0.0.x`

## Architecture Overview

### Core Components

**ThisAddIn.cs** - Main entry point and orchestrator
- Manages add-in lifecycle (startup, shutdown)
- Handles Word application events (NewDocument, WindowActivate, DocumentChange)
- Coordinates between AbbreviationManager, suggestion pane, and Word UI
- Implements Trie-based prefix matching for fast lookups
- Manages per-window CustomTaskPanes with state tracking

**AbbreviationManager.cs** - Data layer
- Loads abbreviations from embedded `Abbreviations.xlsx` using EPPlus
- Implements two-tier caching: JSON file cache and Word AutoCorrect cache
- Cache location: `%APPDATA%\AbbreviationWordAddin\abbreviations.json`
- Version-aware caching (clears on version mismatch)
- Provides lookup methods for both forward (abbrev → full) and reverse (full → abbrev) mapping

**AbbreviationRibbon.cs** - UI controls
- Custom ribbon in Word with Enable/Disable toggle
- "Replace All" button - batch replaces all abbreviations in document
- "Highlight All" button - highlights all phrases that have abbreviations
- "Highlight Like" button - highlights partial matches using progressive phrase matching
- "Show Suggestions" button - reopens task pane if user closed it

**SuggestionPaneControl.cs** - Task pane interface
- Three-tab design:
  1. **Abbreviations tab**: Type phrases, see abbreviation suggestions
  2. **Reverse tab**: Type abbreviations, see full form suggestions
  3. **Dictionary tab**: Browse all loaded abbreviation pairs
- Real-time suggestion updates as user types
- Double-click to accept suggestions (currently disabled in code)
- Replace/Ignore controls for batch processing

### Key Architectural Patterns

**Event-Driven Typing Detection**
- `typingTimer` (300ms interval) triggers debounce logic
- `debounceTimer` (300ms delay) prevents excessive lookups
- Looks back up to 12 words (`maxPhraseLength`) from cursor position
- Stops at paragraph breaks or very short words

**Per-Window Task Pane Management**
- Dictionary of `taskPanes` keyed by `Word.Window`
- Tracks user-closed panes in `userClosedTaskPanes` HashSet
- **Dual-mode auto-reopen behavior**: Task pane automatically reopens when user types either:
  1. A phrase from dictionary (e.g., "station" → shows "stn")
  2. An abbreviation from dictionary (e.g., "stn" → shows "station")
- System searches **both directions simultaneously** to never miss a match
- The pane is removed from `userClosedTaskPanes` when matches are found in either mode

**Progressive Phrase Matching** (HighlightLike feature)
- Splits multi-word phrases into progressive partials
- Example: "Accounting Manager Assistant" → ["Accounting", "Accounting Manager", "Accounting Manager Assistant"]
- Filters out stop words ("a", "the", "of", etc.)
- Matches phrases even if user hasn't typed the full phrase

**Batch Replacement with Progress Tracking**
- Uses `ProgressForm` with background threads
- `SynchronizationContext` for thread-safe Word COM interop
- Disables screen updating during batch operations for performance
- Builds reverse maps for "undo abbreviation" operations

### Data Flow

1. **Startup**: Load abbreviations from cache/Excel → Build Trie index → Initialize AutoCorrect cache
2. **Typing**: DebounceTimer → Check last N words → Trie lookup → Update suggestion pane
3. **Replace All**: Collect all matches → Show in task pane → User reviews each → Replace in document
4. **Highlight**: Regex-based paragraph scanning → Apply formatting to matching ranges

## Important Implementation Notes

### Cache Versioning
The add-in uses `Properties.Settings.Default.AbbreviationDataVersion` and `LastLoadedAbbreviationVersion` to detect when abbreviation data has changed. On version mismatch:
- JSON cache file is deleted
- AutoCorrect cache is cleared
- Fresh load from Excel is triggered

### COM Object Lifecycle
Always release COM objects after use:
```csharp
System.Runtime.InteropServices.Marshal.ReleaseComObject(range);
```

### AutoCorrect Integration
The add-in manipulates Word's native AutoCorrect entries. On startup, it clears all existing entries and repopulates from Excel. The `autoCorrect.ReplaceText` flag is used as a proxy for "is the add-in enabled?"

### Thread Safety
Word COM interop must occur on the UI thread. Use `syncContext.Send()` for synchronous calls from background threads. Never call Word APIs directly from `System.Threading.Thread`.

### Trie Structure
Custom Trie implementation stores all phrases case-insensitively. Each node maintains a list of words passing through it, enabling prefix-based lookups in O(m) time where m is prefix length.

## File Structure

- **Core Logic**: `ThisAddIn.cs`, `AbbreviationManager.cs`
- **UI**: `AbbreviationRibbon.cs`, `SuggestionPaneControl.cs`
- **Dialogs**: `Form1.cs`, `ProgressForm.cs`, `ReplaceDialog.cs`
- **Utilities**: `KeyboardHook.cs` (low-level keyboard hook, currently unused)
- **Data**: `Abbreviations.xlsx` (embedded resource)
- **Templates**: `Templates/*.docx` (embedded document templates)
- **Ribbon XML**: `Ribbon.xml`, `Ribbon1.xml`, etc. (ribbon customization)

## Common Development Tasks

### Adding New Abbreviations
1. Edit `Abbreviations.xlsx` (two columns: phrase, abbreviation)
2. Increment `AbbreviationDataVersion` in `Properties/Settings.settings`
3. Rebuild project (Excel file is embedded as resource)
4. On next Word launch, cache will refresh automatically

### Modifying Suggestion Behavior
- Adjust `maxPhraseLength` in `ThisAddIn.cs` to look back further
- Modify `DebounceDelayMs` constant to change typing sensitivity
- Edit `DebounceTimer_Tick()` method for lookup logic changes

### Adding Ribbon Buttons
1. Open `AbbreviationRibbon.Designer.cs` in Visual Studio Designer
2. Add button via UI designer
3. Implement click handler in `AbbreviationRibbon.cs`
4. Follow async/await pattern for long-running operations (see existing buttons)

## Known Patterns and Conventions

- **Debug Mode**: Set `debug = true` in `ThisAddIn.cs` to enable MessageBox debugging
- **Case Insensitivity**: All phrase lookups use `ToLowerInvariant()` and `InvariantCultureIgnoreCase`
- **AutoCorrect Cache**: Checked first before dictionary lookup for performance
- **Screen Updating**: Disabled during batch operations to prevent flicker
- **Find/Replace**: Always use `MatchWholeWord = true` to avoid partial word matches
