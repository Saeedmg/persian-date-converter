================================================================
  PERSIAN DATE CONVERTER  —  Quick Start Guide
  Version 2.0
  Translation Office 1324
  Saeed Majidi  |  rassamtranslation@gmail.com
================================================================

  This guide walks you through:
    1. Unzipping the tool to C:\
    2. Understanding the file layout
    3. Adding the macro to Word
    4. Binding keyboard shortcuts (F2, F3, ...)
    5. Using the quick shortcuts
    6. A brief overview of the batch converter

  Read top to bottom. Follow the steps in order.


================================================================
STEP 1 — UNZIP TO DRIVE C
================================================================

You will receive a file called one of these:

    PersianDateTool.zip
    PersianDateTool.rar

Extract it so the final folder is at:

    C:\PersianDateTool\

  ▸ On Windows 10/11, right-click the zip → Extract All…
  ▸ In the destination box, type:  C:\
  ▸ Click Extract.

After extraction, the folder MUST be at:

    C:\PersianDateTool\

  ⚠ Do NOT put it in Downloads, Desktop, or Documents.
    The Word macro has C:\PersianDateTool\ hard-coded.

  If you must use a different path, see the note at the bottom
  of this guide under "MOVING TO A DIFFERENT FOLDER".


================================================================
STEP 2 — FILE LAYOUT (WHAT YOU'LL SEE)
================================================================

After unzipping, C:\PersianDateTool\ contains:

    C:\PersianDateTool\
    │
    ├── PersianDateConverter.exe          ← Double-click to open the GUI
    │
    ├── dist\
    │   ├── PersianDateConverter.exe               (GUI - same as above)
    │   ├── PersianDateConverterCLI\               (CLI folder)
    │   │   └── PersianDateConverterCLI.exe            (Word calls this)
    │   ├── RunHidden.vbs                          (hides the console)
    │   ├── settings.ini                           (remembers your format)
    │   └── log.txt                                (conversion history)
    │
    └── word\
        └── ConvertPersianDate.bas                 (macro backup)


  The two files that matter for Word:

    C:\PersianDateTool\dist\PersianDateConverterCLI\PersianDateConverterCLI.exe
    C:\PersianDateTool\dist\RunHidden.vbs

  The Word macro talks to these two files. Keep them where they are.


================================================================
STEP 3 — ADD THE MACRO TO WORD
================================================================

  1) Open Microsoft Word.

  2) Press Alt + F11 on your keyboard.
     The VBA editor opens in a new window.

  3) In the left panel, look for "Normal" → "Modules".
     If a module called ConvertPersianDate already exists,
     right-click it and choose  Remove  →  No.

  4) Menu:  File  →  Import File…

  5) Navigate to:
        C:\PersianDateTool\word\ConvertPersianDate.bas
     Click Open.

  6) Press Ctrl + S.
     A dialog asks: "Save changes to the Normal template?"
     Click YES.

  7) Press Alt + Q to return to Word.

  Done. The macro is now part of Word.
  You only do this ONCE.


================================================================
STEP 4 — BIND KEYBOARD SHORTCUTS (F2, F3, ...)
================================================================

Choose one or more shortcuts for the macros you'll use most.
I recommend:

    F2       →  ConvertPersianDate            (uses GUI format)
    F3       →  ConvertPersianDate_Full       (September 17, 2024)
    F4       →  ConvertPersianDate_Abbrev     (Sep. 17, 2024)
    F6       →  ConvertPersianDate_ISO        (2024-09-17)

  F2, F3, F4, F6 are free in a default Word install (F5 is
  "Go To" and F7 is spell-check, so avoid those).

  To bind them:

  1) In Word:  File  →  Options  →  Customize Ribbon.

  2) At the bottom, click  "Keyboard shortcuts: Customize…"
     A new dialog opens.

  3) In the left list "Categories", scroll to the bottom
     and select  Macros.

  4) In the right list "Macros", click  ConvertPersianDate.

  5) Click inside  "Press new shortcut key".
     Press  F2  on your keyboard.

  6) Click  Assign.

  7) Repeat steps 4–6 for the other macros:
        ConvertPersianDate_Full     →  F3
        ConvertPersianDate_Abbrev   →  F4
        ConvertPersianDate_ISO      →  F6

  8) Click Close.

  If a key says "Currently assigned to", you can still assign
  it — Word will overwrite the old binding. Or pick another key.


================================================================
STEP 5 — USING THE QUICK SHORTCUTS
================================================================

In any Word document:

    Press  F2   (or your chosen key)
    Type the Persian date:   1403/06/27
    Click  OK

    The converted date pastes at the cursor.

  That's it. No black window, no dialog after the input box.


  WHICH SHORTCUT TO USE
  ---------------------

    F2  (ConvertPersianDate)
         Uses whatever format is currently set in the GUI.
         Best for general work — change the format in the
         GUI, and F2 follows automatically.

    F3  (ConvertPersianDate_Full)
         Always:  September 17, 2024

    F4  (ConvertPersianDate_Abbrev)
         Always:  Sep. 17, 2024

    F6  (ConvertPersianDate_ISO)
         Always:  2024-09-17

  You can bind more formats to other keys if you want.
  There are 7 pre-defined macros, each with its own format:

    ConvertPersianDate                 (GUI format)
    ConvertPersianDate_Full            (Sept 17, 2024)
    ConvertPersianDate_Abbrev          (Sep. 17, 2024)
    ConvertPersianDate_British         (17 September 2024)
    ConvertPersianDate_BritishAbbrev   (17 Sep. 2024)
    ConvertPersianDate_ISO             (2024-09-17)
    ConvertPersianDate_DotISO          (2024.09.17)
    ConvertPersianDate_SlashUS         (09/17/2024)


  INPUT FORMAT
  ------------
  Always type the Persian date as:

        YYYY/MM/DD

  Examples:
        1403/06/27
        1364/01/20
        1403/1/1       (leading zeros are optional)

  Use the English keyboard. Forward slashes only.


  WORKS WITH PROTECTED DOCUMENTS
  ------------------------------
  If the document has form fields (like an application form),
  the macro fills the field directly. If the document is
  password-protected, you'll be asked for the password once.


================================================================
STEP 6 — THE BATCH CONVERTER (BRIEF OVERVIEW)
================================================================

The GUI includes a batch feature that converts whole folders
of Word files at once. Useful when you have 50 documents and
don't want to press F2 for each one.

  HOW TO USE

  1) Double-click:
        C:\PersianDateTool\dist\PersianDateConverter.exe

  2) Click the  "Word Files (Batch)"  tab.

  3) In "Input folder", choose the folder containing your
     .docx files. (Click Browse…)

  4) In "Output folder", choose where the converted files
     should go. (New folder recommended.)

  5) From the "Format" dropdown, pick a date format.

  6) Click  "Convert Word Files".

  7) Wait. A progress log appears at the bottom of the tab.

  WHAT IT DOES
  ------------
  It scans every .docx in the input folder, finds every date
  written as  YYYY/MM/DD, converts each one, and saves a new
  file called:

        OriginalName_updated.docx

  in the output folder. The originals are untouched.

  It works on:
    ✔ Paragraphs
    ✔ Tables
    ✔ Headers and footers

  It does NOT:
    ✘ Change the original files
    ✘ Touch .doc files (only .docx)
    ✘ Convert dates written in other formats (only YYYY/MM/DD)


================================================================
GUI QUICK REFERENCE
================================================================

Open the GUI:

    C:\PersianDateTool\dist\PersianDateConverter.exe

  Tabs:

    Single Date          — Type one date, get the English result,
                           Copy to Clipboard button.

    Word Files (Batch)   — Convert whole folders of .docx files.

    Word Macro           — Shows the macro code and copies it
                           to the clipboard (useful when setting
                           up Word on a new PC).

    Log                  — Shows every conversion performed,
                           with a timestamp.

    About                — Version and contact info.

  To change the DEFAULT format that the F2 shortcut uses:
    1. Open the GUI.
    2. Single Date tab.
    3. Click a format radio button.
    4. Close the GUI.
  Next time you press F2, the new format is used.


================================================================
WHEN YOU GET A NEW PC
================================================================

  1) Copy the whole  C:\PersianDateTool\  folder to the new PC
     at the same path.

  2) Add the macro to Word (Step 3 in this guide).

  3) Bind the shortcuts (Step 4).

  That's it. No Python needed.


================================================================
MOVING TO A DIFFERENT FOLDER
================================================================

If you must install somewhere other than C:\PersianDateTool\:

  1) Put the folder wherever you want, e.g. D:\Tools\PersianDateTool\

  2) In Word:  Alt + F11  →  click ConvertPersianDate  →  find:

        exePath = "C:\PersianDateTool\dist\PersianDateConverterCLI\PersianDateConverterCLI.exe"
        vbsPath = "C:\PersianDateTool\dist\RunHidden.vbs"

  3) Change them to your actual location.

  4) Ctrl + S → Yes.

  Do NOT move files individually. Always move the whole folder.


================================================================
TROUBLESHOOTING
================================================================

  Problem: "Nothing pastes, no error"
  Fix:     Confirm these two files exist:
             C:\PersianDateTool\dist\PersianDateConverterCLI\PersianDateConverterCLI.exe
             C:\PersianDateTool\dist\RunHidden.vbs

  Problem: "Error: Please enter a valid Persian date"
  Fix:     Use YYYY/MM/DD. English digits. Forward slashes.

  Problem: "A black window flashes"
  Fix:     The macro is the old version. Re-import
           ConvertPersianDate.bas (Step 3).

  Problem: "Arabic shows as ????? or boxes"
  Fix:     The macro is the old version. Re-import
           ConvertPersianDate.bas.

  Problem: "Shortcut does nothing"
  Fix:     The macro isn't saved to the Normal template.
           Alt + F11 → Ctrl + S → Yes.

  Problem: "The document is protected with a password"
  Fix:     The macro asks for the password once. Enter it, or
           unprotect the file manually (Review → Restrict Editing
           → Stop Protection) and save.

  Problem: "Macro is gone after restarting Word"
  Fix:     Always click YES when Word asks to save changes
           to the Normal template.


================================================================
CONTACT
================================================================

  Developed by:  Saeed Majidi
  Company:       Translation Office 1324
  Email:         rassamtranslation@gmail.com
  Version:       2.0

================================================================
  End of Quick Start Guide
================================================================