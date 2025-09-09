# WordPicturePlaceholderAddin

![Picture Add In](https://github.com/bigboybamo/WordPicturePlaceholderAddin/blob/master/Pictures/Picture_Add_In.png)

A lightweight VSTO add-in for Microsoft Word that helps you mark, find, and clean up image spots while drafting. It adds a small ribbon with three buttons:
- **Insert placeholder** – Inserts a numbered picture placeholder like `(picture 1)`, `(picture 2)`, …

- **Show placeholders** – Lists all placeholders in the active document so you can jump to them.

- **Remove placeholder** – Removes a specific placeholder by number.


## Why it exists

I built this after too many drafts juggling screenshots and losing track of “(picture n)” notes. It gives me a dead-simple way to drop clear markers, see them all at once, and clean them up when the real images arrive.

## Features

- 🔢 Auto-numbered placeholders – Always inserts the next number after the highest (picture N) found.
- 🔎 Document-wide listing – Quickly see all placeholders, their numbers, and (if supported by Word) page context.
- ❌ Selective removal – Remove by number to remove the placeholder. Automatic placeholder readjusting

## Requirements

- Windows: Word on Windows only (VSTO)
- Microsoft Word: 2016 or later / Microsoft 365
- .NET: .NET Framework 4.8 (or the framework your VSTO project targets)
- Visual Studio: 2019/2022 with Office developer tools (for development/build)

Usage

1. **Insert placeholder**
   
  ![Insert Placeholder](https://github.com/bigboybamo/WordPicturePlaceholderAddin/blob/master/Pictures/Tool_In_Use.png)
 - Click Insert placeholder.
 - The add-in scans the document, finds the highest existing number, and inserts the next one, e.g. `(picture 3)`.

2. **Show placeholders**
   
   ![Show Placeholders](https://github.com/bigboybamo/WordPicturePlaceholderAddin/blob/master/Pictures/Placeholder_List.png)
   - Click Show placeholders.
   - A small dialog lists all placeholders found in the active document.
  
4. **Remove placeholder**
   
    ![Remove Placeholder](https://github.com/bigboybamo/WordPicturePlaceholderAddin/blob/master/Pictures/Remove_PlaceHolder.png)
   - Click Remove placeholder.
   - Choose By number, and enter e.g. 3 to remove `(picture 3)`.
  
## Build & debug
1. Open the solution in Visual Studio.
2. Set WordPicturePlaceholderAddin as startup project.
3. Press F5. Visual Studio will launch Word with the add-in loaded.
4. Use the ribbon buttons to test.
