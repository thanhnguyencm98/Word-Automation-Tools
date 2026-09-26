# Word Automation Tools

Beginner-friendly installation and usage guide for Microsoft Word VBA `.bas` modules.

## 1. Purpose

This package provides reusable Microsoft Word macros for document cleanup and footnote-format review:

| Module file | Macro | Purpose |
|---|---|---|
| `RemoveAllHyperlinks.bas` | `RemoveAllHyperlinks` | Removes clickable hyperlinks while retaining the visible text. |
| `HighlightItalicFootnotes.bas` | `HighlightItalicFootnotes` | Highlights italic-formatted text in footnotes using yellow highlighting. |
| `HighlightSmallCapsFootnotes.bas` | `HighlightSmallCapsFootnotes` | Highlights text formatted with Word's **Small caps** property in footnotes using yellow highlighting. |
| `RunAllChecks.bas` | `RunAllChecks` | Runs the three main macros in sequence. This module is optional. |

> **Important:** These macros modify the active document. Always save a backup copy before running them.

---

## 2. Requirements

- Microsoft Word desktop application for Windows
- Permission to use VBA macros under your organization's security policy
- The supplied `.bas` files
- A document that you are permitted to modify

Word for the web cannot create, edit, or run VBA macros. Use **Open in Desktop App** when a document is stored in OneDrive or SharePoint.

---

## 3. What is a `.bas` file?

A `.bas` file is an exported VBA standard module. It contains macro source code but is not a Word document by itself.

The `.bas` files must be imported into one of these containers:

- A **Word Macro-Enabled Document**, using the `.docm` extension
- A **Word Macro-Enabled Template**, using the `.dotm` extension
- Word's global `Normal.dotm` template

For this package, the recommended container is:

```text
WordAutomationTools.docm
```

This keeps the tools separate from `Normal.dotm` and makes the package easier to review, back up, and share according to organizational policy.

---

## 4. Package layout

The repository and release package are structured as follows:

```text
WordAutomationTools/
├── README.md
├── Macros/
│   ├── CommonUtilities.bas              # Shared helpers and guards
│   ├── RemoveAllHyperlinks.bas          # Hyperlink removal
│   ├── HighlightItalicFootnotes.bas     # Footnote italic highlighting
│   ├── HighlightSmallCapsFootnotes.bas  # Footnote Small Caps highlighting
│   └── RunAllChecks.bas                 # Master review runner
├── TestDocuments/
│   └── SampleFootnoteTest.docx          # Test document fixture
└── Results/
    └── TestingReport.docx               # Verification test report
```

If `RunAllChecks.bas` is not supplied, the three main macros can still be run separately.

---

## 5. Create `WordAutomationTools.docm`

### Step 1: Create the Word document

1. Open Microsoft Word for Windows.
2. Select **Blank document**.
3. Select **File** > **Save As** > **Browse**.
4. Enter the file name:

   ```text
   WordAutomationTools.docm
   ```

5. In **Save as type**, select **Word Macro-Enabled Document (*.docm)**.
6. Select **Save**.

Do not save the file as `.docx`. A `.docx` file cannot retain VBA code.

### Step 2: Open the VBA editor

Press:

```text
Alt + F11
```

The **Microsoft Visual Basic for Applications** editor opens.

If the shortcut does not work, enable the **Developer** tab:

1. In Word, select **File** > **Options**.
2. Select **Customize Ribbon**.
3. Under **Main Tabs**, select **Developer**.
4. Select **OK**.
5. On the **Developer** tab, select **Visual Basic**.

### Step 3: Locate the correct VBA project

In the left-side **Project Explorer**, find a project similar to:

```text
Project (WordAutomationTools.docm)
```

If **Project Explorer** is not visible, press:

```text
Ctrl + R
```

Make sure the selected project belongs to `WordAutomationTools.docm`, not `Normal.dotm` or another open document.

---

## 6. Import the `.bas` files

Repeat the following procedure for every supplied `.bas` file:

1. In **Project Explorer**, select `Project (WordAutomationTools.docm)`.
2. In the VBA editor, select **File** > **Import File**.
3. Browse to the folder containing the `.bas` files.
4. Select one `.bas` file.
5. Select **Open**.
6. Confirm that the imported module appears under the project's **Modules** folder.

Import these required modules:

```text
RemoveAllHyperlinks.bas
HighlightItalicFootnotes.bas
HighlightSmallCapsFootnotes.bas
```

Import this module when included:

```text
RunAllChecks.bas
```

After importing, **Project Explorer** should look similar to:

```text
Project (WordAutomationTools.docm)
└── Modules
    ├── CommonUtilities
    ├── RemoveAllHyperlinks
    ├── HighlightItalicFootnotes
    ├── HighlightSmallCapsFootnotes
    └── RunAllChecks
```

The exact module display names may differ slightly from the file names. This is acceptable if the expected macro names are present.

---

## 7. Review the source code before enabling it

Macro files can contain executable instructions. Review the source before use, especially when files were received through email, chat, or an internet download.

For each imported module:

1. Double-select the module under **Modules**.
2. Read the code in the editor window.
3. Confirm that the code performs only the documented Word operations.
4. Confirm that the code does not contain unexpected operations such as:
   - Starting external applications
   - Running shell or command-line commands
   - Downloading content
   - Sending data over a network
   - Writing unrelated files
   - Automatically running when the document opens

These tools are expected to operate on Word document ranges, hyperlinks, footnotes, font formatting, and highlighting only.

---

## 8. Compile the VBA project

Compilation detects syntax errors and unresolved VBA references before normal use.

1. In the VBA editor, select **Debug** > **Compile Project** or **Compile VBAProject**.
2. If no message appears and the compile command becomes unavailable, compilation completed successfully.
3. If an error is displayed, note the highlighted line and verify that:
   - All required `.bas` files were imported.
   - No source lines were accidentally changed.
   - The document is open in Microsoft Word, not another Office application.

### Check for missing references

If compilation reports **Can't find project or library**:

1. In the VBA editor, select **Tools** > **References**.
2. Look for an entry beginning with `MISSING:`.
3. Do not select an arbitrary replacement library.
4. Clear only an unrelated missing optional reference, or restore the required component according to your organization's support process.
5. Compile the project again.

The supplied macros should rely on Word's built-in object model and should not require an added third-party reference.

---

## 9. Save and reopen the tool document

1. In the VBA editor, press Ctrl + S.
2. Close the VBA editor with Alt + Q.
3. In Word, press Ctrl + S again.
4. Close `WordAutomationTools.docm`.
5. Reopen `WordAutomationTools.docm` in desktop Word.

If Word warns that the document contains macros, this is expected for a `.docm` file.

---

## 10. Macro security

Only enable macros after reviewing the source and confirming that the package is trusted and permitted by your organization.

### If Word displays **Security Warning: Macros have been disabled**

Select **Enable Content** only when all of the following are true:

- The source of the files is trusted.
- The VBA code has been reviewed.
- The macros are allowed by organizational policy.
- The file is the expected `WordAutomationTools.docm` package.

### If Word says macros are blocked because the file came from the internet

On a personally managed Windows device, a trusted file may expose an **Unblock** option:

1. Close Word.
2. In File Explorer, right-select `WordAutomationTools.docm` and select **Properties**.
3. On the **General** tab, look for **Unblock**.
4. Select **Unblock**, then select **Apply** and **OK**.
5. Reopen the file.

If **Unblock** is unavailable, or company policy still blocks the macros, contact the organization's IT or security team. Do not weaken organization-managed Trust Center or Group Policy settings.

---

## 11. Confirm that the macros are installed

1. Open `WordAutomationTools.docm`.
2. Press Alt + F8.
3. In **Macros in**, select **All active templates and documents** if needed.
4. Confirm that the following macros are listed:

```text
RemoveAllHyperlinks
HighlightItalicFootnotes
HighlightSmallCapsFootnotes
RunAllChecks                 # when the optional module was imported
```

If the macros are not listed, see the troubleshooting section.

---

## 12. Recommended way to use the tool document

The macros safely target your open working documents without modifying `WordAutomationTools.docm` itself:

### Workflow:

1. Open `WordAutomationTools.docm`.
2. Open the document(s) you need to review and format.
3. Switch to the **`WordAutomationTools.docm`** window and press **Alt + F8** (Word only lists macros of an open `.docm` document when its own window is active).
4. Select the desired macro (e.g. `RunAllChecks`) and click **Run**.
5. **Automatic Document Selection**:
   - If 1 target document is open: A prompt asks for confirmation to process it.
   - If multiple target documents are open: A dialog lists all open documents with numbers (1, 2, 3...). Enter the number of the document you wish to process.
6. Review the summary report displayed after execution.
7. Save your modified document.

> **Note:** Because `WordAutomationTools.docm` is a standalone document rather than an installed Add-in (`.dotm`), you must trigger Alt + F8 from its window. `GetTargetDocument` will then redirect and apply all operations to your working document.

---

## 13. Macro usage

### 13.1 `RemoveAllHyperlinks`

**Purpose:** Removes clickable hyperlinks while keeping their visible display text.

**Recommended procedure:**

1. Save a backup of the target document.
2. Activate the target document.
3. Press Alt + F8.
4. Select `RemoveAllHyperlinks`.
5. Select **Run**.
6. Accept the confirmation only if hyperlink removal is intended.
7. Review hyperlinks in body text, footnotes, endnotes, headers, footers, and text boxes.

**Example:**

Before:

```text
Microsoft Support
```

The text is clickable and opens a webpage.

After:

```text
Microsoft Support
```

The visible text remains, but the hyperlink is removed.

**Important distinction:**

This macro should remove Word hyperlink objects. It should not be replaced with a general `Fields.Unlink` operation, because unlinking every field can affect tables of contents, cross-references, page-number fields, citation fields, and other dynamic Word content.

### 13.2 `HighlightItalicFootnotes`

**Purpose:** Highlights italic-formatted text in Word footnotes.

**Procedure:**

1. Activate the target document.
2. Press Alt + F8.
3. Select `HighlightItalicFootnotes`.
4. Select **Run**.
5. Review the yellow highlights in the footnote area.

**Expected result:**

- Italic text in footnotes is highlighted yellow.
- Bold italic text is also expected to match because it has italic formatting.
- Italic text in the main document body is not targeted.
- Endnotes are not targeted unless the VBA code is explicitly extended for endnotes.
- Text that merely looks slanted because of the chosen typeface may not be detected as Word italic formatting.

### 13.3 `HighlightSmallCapsFootnotes`

**Purpose:** Highlights text using Word's **Small caps** font property in footnotes.

**Procedure:**

1. Activate the target document.
2. Press Alt + F8.
3. Select `HighlightSmallCapsFootnotes`.
4. Select **Run**.
5. Review the yellow highlights in the footnote area.

**Expected result:**

- Text formatted with **Font > Small caps** is highlighted yellow.
- Plain text typed with uppercase letters is not necessarily Small Caps.
- Small Caps in the main document body is not targeted.
- Endnotes are not targeted unless the code is extended.

### 13.4 `RunAllChecks`

**Purpose:** Runs the three main operations in sequence.

Use this macro only when all three changes are intended. Since hyperlink removal changes document content behavior, the safer review workflow is normally:

1. Run `HighlightItalicFootnotes`.
2. Run `HighlightSmallCapsFootnotes`.
3. Review and correct citations.
4. Run `RemoveAllHyperlinks` only when hyperlink removal is required.

---

## 14. Validation before production use

Create a disposable test document containing all of the following:

### Body text

- One normal paragraph
- One clickable web hyperlink
- One italic phrase
- One Small Caps phrase

### Footnotes

- Normal footnote text
- Italic footnote text
- Bold italic footnote text
- Word Small Caps footnote text
- Plain `ALL-CAPS` footnote text without Small Caps formatting
- A clickable hyperlink in a footnote

### Optional document elements

- A hyperlink in a header
- A hyperlink in a footer
- A hyperlink in a text box
- A cross-reference field
- A table of contents

Run each macro separately and record the result:

| Check | Expected result |
|---|---|
| Body hyperlink | Removed by `RemoveAllHyperlinks`; visible text remains. |
| Footnote hyperlink | Removed by `RemoveAllHyperlinks`; visible text remains. |
| Footnote italic text | Highlighted yellow. |
| Footnote bold italic text | Highlighted yellow. |
| Footnote Small Caps text | Highlighted yellow. |
| Plain ALL-CAPS footnote text | Not highlighted unless Small Caps formatting is applied. |
| Body italic text | Not highlighted by the footnote macro. |
| Body Small Caps text | Not highlighted by the footnote macro. |
| Cross-reference and table of contents | Remain functional when the hyperlink macro deletes hyperlinks rather than unlinking all fields. |

Do not deploy the package for normal document processing until the test results match the intended behavior in the Word version used by the target users.

---

## 15. How to verify Italic and Small Caps formatting manually

### Verify Italic

1. Select the text.
2. Open the **Home** tab.
3. Check whether the **Italic** button is active.

### Verify Small Caps

1. Select the text.
2. On the **Home** tab, open the **Font** dialog launcher.
3. Under **Effects**, check whether **Small caps** is selected.

Typing uppercase letters is not the same as applying the **Small caps** effect.

---

## 16. Undo and recovery

### Immediate undo

After one macro finishes, press:

```text
Ctrl + Z
```

However, VBA operations and long editing sessions may not always provide the same undo behavior as ordinary manual edits. A backup copy is the reliable recovery method.

### Recommended backup naming

```text
OriginalDocument_BACKUP.docx
OriginalDocument_WORKING.docx
```

Run macros only on the working copy until validation is complete.

---

## 17. Troubleshooting

### Problem: Alt + F11 does not open the VBA editor

- Use **Developer** > **Visual Basic**.
- Confirm that desktop Microsoft Word is being used.
- If the command is disabled, VBA may be restricted by organizational policy.

### Problem: The `.bas` file opens as text instead of importing

Do not open the `.bas` file from File Explorer. Import it from the VBA editor using **File** > **Import File**.

### Problem: The macros disappear after saving

The document was probably saved as `.docx`.

1. Reopen the source document before closing it, if possible.
2. Select **File** > **Save As**.
3. Choose **Word Macro-Enabled Document (*.docm)**.
4. Reimport the `.bas` files if the macros have already been removed.

### Problem: The macro is not shown in Alt + F8

Check the following:

- The procedure is declared as `Public Sub` or `Sub`.
- The procedure has no required parameters.
- The module is located under **Modules**, not inside an unexpected object module.
- **Macros in** is set to **All active templates and documents**.
- `WordAutomationTools.docm` is open.
- The VBA project compiles without errors.

### Problem: The wrong document was modified

The macro acts on `ActiveDocument`.

1. Stop editing and avoid saving over the original.
2. Press Ctrl + Z when immediately available.
3. Close the incorrectly modified document without saving, if appropriate.
4. Restore the backup when necessary.
5. Reopen both files and activate the intended target before running the macro.

### Problem: No footnote text is highlighted

Verify that:

- The document contains footnotes rather than endnotes.
- The target text has actual Italic or Small Caps formatting.
- The correct target document was active.
- The target text is not inside a comment, text box, or another story type that is not part of the Footnotes collection.

### Problem: Plain uppercase text was not highlighted

This is expected. Apply Word's **Small caps** property if the text is intended to be Small Caps.

### Problem: Word reports a compile error

1. Note the exact error message and highlighted line.
2. Confirm that every `.bas` file is complete and unmodified.
3. Confirm that all required modules were imported.
4. Select **Tools** > **References** and check for `MISSING:` entries.
5. Reimport the original module if accidental edits are suspected.

### Problem: Macros remain blocked

The organization's security policy may prohibit unsigned macros or macros from downloaded files. Use the approved trusted-location or code-signing process provided by the organization's IT/security team. Do not select **Enable all macros** as a general workaround.

---

## 18. Export an updated module

If the VBA source is intentionally changed and needs to be saved back to a `.bas` file:

1. Open the VBA editor with Alt + F11.
2. In **Project Explorer**, select the module.
3. Select **File** > **Export File**.
4. Save the file using its existing `.bas` name.
5. Store the exported source under version control or in the approved project location.
6. Compile and retest the complete project.

Do not rename a module casually if another module calls one of its public macros.

---

## 19. Optional: Add macro buttons to the Quick Access Toolbar

After the package has been reviewed and tested:

1. Select **File** > **Options**.
2. Select **Quick Access Toolbar**.
3. In **Choose commands from**, select **Macros**.
4. Select the required macro.
5. Select **Add**.
6. Use **Modify** to choose a recognizable icon and display name.
7. Select **OK**.

Suggested display names:

```text
Remove Hyperlinks
Highlight Footnote Italics
Highlight Footnote Small Caps
Run All Word Checks
```

Keep destructive or content-changing actions visually distinct. `RemoveAllHyperlinks` should not be assigned to an easily pressed keyboard shortcut without a confirmation prompt.

---

## 20. Maintenance and change control

For professional or organizational use:

- Keep the `.bas` source files under version control.
- Record the package version in this README.
- Require source review for every change.
- Compile and test after every modification.
- Test against representative documents, including long documents and documents with many footnotes.
- Digitally sign the VBA project when required by organizational policy.
- Distribute the package only through an approved internal channel.
- Do not add automatic execution procedures such as `AutoOpen` or `Document_Open` unless explicitly reviewed and approved.

Suggested version record:

```text
Package: Word Automation Tools
Version: 1.0.0
Container: WordAutomationTools.docm
Modules: 4 required (including CommonUtilities), 1 optional
Highlight color: Yellow
Scope: Hyperlinks throughout available Word story ranges; formatting checks in footnotes
```

---

## 21. Known limitations

- The formatting macros target footnotes, not endnotes.
- The macros detect Word formatting properties, not visual appearance inferred from a font design.
- Plain uppercase text is not equivalent to Small Caps.
- Existing yellow highlighting is preserved without duplicate counting; non-yellow highlights (e.g., turquoise, pink) are preserved and reported as skipped per FR-5.
- Password-protected, read-only, restricted, or protected documents may prevent modifications.
- Tracked Changes will record hyperlink deletions and highlights as markup if enabled. Turn off Track Changes or accept prior revisions before running.
- Hyperlinks produced or refreshed by other dynamic fields may require separate review.
- Results should be validated in the organization's supported Word version before production use.

---

## 22. Final user checklist

Before running a macro:

- [ ] I am using desktop Microsoft Word for Windows.
- [ ] I reviewed or trust the imported VBA source.
- [ ] Organizational policy allows me to run the macros.
- [ ] I saved a backup of the target document.
- [ ] `WordAutomationTools.docm` is open.
- [ ] The intended target document is the active document.
- [ ] I selected the correct macro.

After running a macro:

- [ ] I read the completion or error message.
- [ ] I inspected the changed locations.
- [ ] I confirmed that unrelated fields and formatting still work.
- [ ] I saved the result under the intended file name.
- [ ] I retained the original backup until review is complete.

---

## 23. Security references

- [Create or run a macro in Word](https://support.microsoft.com/office/create-or-run-a-macro-c6b99036-905c-49a6-818a-dfb98b7c3c9c)
- [Enable or disable macros in Microsoft 365 files](https://support.microsoft.com/topic/12b036fd-d140-4e74-b45e-16fed1a7e5c6)
- [Macros from the internet are blocked by default in Office](https://learn.microsoft.com/deployoffice/security/internet-macros-blocked)

---

## Disclaimer

These macros are document-review assistance tools. Users remain responsible for reviewing the resulting document and confirming compliance with applicable citation, editorial, legal, security, and organizational requirements.
