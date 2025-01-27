# DocxToPDF

A macro-based tool for automatically converting Word documents (.docx) to PDFs using VBA in Microsoft Word.

## How to Use

This macro allows you to quickly save your active Word document as a PDF with the same filename and location as the original `.docx` file. Follow the steps below to use this code.

### Requirements
- Microsoft Word (for VBA support).
- Basic knowledge of how to enable and run macros in Word.

### Instructions

1. **Open your Word document**: Make sure it's the document you want to save as a PDF.

2. **Enable Developer Tab**:
   If the Developer tab is not visible in Word, you can enable it by:
   - Going to `File` > `Options`.
   - In the Options window, select `Customize Ribbon`.
   - Check the box for `Developer` on the right side and click `OK`.

3. **Create the Macro**:
   - Open the Word document you want to convert to a PDF.
   - Go to the `Developer` tab and click on `Macros`.
   - Enter a name for your macro (e.g., `SaveAsPDF`), then click `Create`.
   - Copy and paste the following VBA code into the editor:

     ```vba
     Sub SaveAsPDF()
         Dim docName As String
         Dim pdfPath As String

         ' Get the name of the active document (without the extension)
         docName = Left(ActiveDocument.Name, InStrRev(ActiveDocument.Name, ".") - 1)

         ' Set the PDF file path (same folder as the Word document)
         pdfPath = ActiveDocument.Path & "\" & docName & ".pdf"

         ' Save the document as PDF
         ActiveDocument.ExportAsFixedFormat OutputFileName:=pdfPath, _
             ExportFormat:=wdExportFormatPDF

         MsgBox "PDF saved as: " & pdfPath
     End Sub
     ```

4. **Run the Macro**:
   - After pasting the code, close the VBA editor.
   - Go back to Word, and under the `Developer` tab, click on `Macros`.
   - Select `SaveAsPDF` from the list and click `Run`.

5. **Result**:
   - The macro will automatically save the active Word document as a PDF in the same location with the same name (but with a `.pdf` extension).
   - A confirmation message will pop up displaying the location of the saved PDF.

---

## Configuring Ctrl+Alt+P to Run the Macro

To make the macro run with the **Ctrl+Alt+P** keyboard shortcut, follow these steps:

1. **Assign the Macro to Ctrl+Alt+P**:
   - Open Word and go to the `File` tab.
   - Click on `Options` to open Word's settings.
   - In the `Word Options` window, select `Customize Ribbon`.
   - Click on the `Customize...` button next to `Keyboard shortcuts` at the bottom of the window.
   - In the `Customize Keyboard` window, under the "Categories" list, select `Macros`.
   - In the "Macros" list, select the `SaveAsPDF` macro.
   - Click in the "Press new shortcut key" field and press **Ctrl+Alt+P**.
   - If the shortcut is already assigned to something else, you can choose to overwrite it or pick another combination.
   - Click `Assign` and then `Close`.

2. **Use Ctrl+Alt+P**:
   - Now, you can simply press **Ctrl+Alt+P** anytime you want to save your active Word document as a PDF.

---

### Customization

You can customize the file-saving location or add more export settings based on your needs. For example, you can modify the `pdfPath` to save the PDF in a different folder.

