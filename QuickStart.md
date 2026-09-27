# Quick Start Guide: Word Citation Automation Tools

A quick guide to using the macros in `WordAutomationTools.docm`.

---

## 1. Setup (first time only)

1. Download or open **`WordAutomationTools.docm`**.
2. If Word shows a yellow security warning bar, select **Enable Content**.

---

## 2. Daily workflow

### Step 1: Open the documents
1. Open the tool file **`WordAutomationTools.docm`**.
2. Open the Word document (`.docx` or `.docm`) whose citations you need to review.

### Step 2: Run a macro
1. **Switch to the `WordAutomationTools.docm` window** and press **Alt + F8**.
   *(Note: Word only lists the macros when Alt + F8 is pressed from the `WordAutomationTools.docm` window.)*
2. Select one of the macros:
   - **`RunAllChecks`**: Runs all 3 steps (remove links → highlight italics → highlight Small Caps in footnotes). *(Recommended)*
   - **`RemoveAllHyperlinks`**: Removes hyperlinks only.
   - **`HighlightItalicFootnotes`**: Highlights italic text in footnotes only.
   - **`HighlightSmallCapsFootnotes`**: Highlights Small Caps text in footnotes only.
3. Select **Run**.

### Step 3: Choose the document to process
- **If only 1 document is open:** A dialog asks for confirmation:
  `"Do you want to process the open target document: 'File_Name.docx'?"` → Select **Yes**.
- **If several documents are open:** A dialog shows a numbered list:
  ```text
  Multiple target documents are open.
  Enter the number corresponding to the document you want to process:

  1. LegalBrief_v1.docx
  2. ResearchNote.docx
  ```
  Type the number (for example `1`) and select **OK**.
- **If the document has Track Changes turned on:** The macro asks whether to continue, because every change will be recorded as a revision. To avoid this, select **No**, turn off Track Changes, and run the macro again.
- **For `RunAllChecks` and `RemoveAllHyperlinks`:** Confirm the final prompt with **Yes**.

### Step 4: Review the results
1. When the macro finishes, it shows a summary (number of hyperlinks removed, number of text segments highlighted). Select **OK**.
2. Check the changes in the document.
3. If the result is not what you wanted, press **Ctrl + Z** once to undo the whole macro run.
4. Save the modified document.

---

## 3. Safety notes
- Always save a backup copy of the document before removing hyperlinks.
- Table of contents links are kept, so the table of contents stays clickable.
- Italic or Small Caps text that already has a highlight color other than yellow keeps its original color.
- The macros never modify the tool file `WordAutomationTools.docm` itself.
