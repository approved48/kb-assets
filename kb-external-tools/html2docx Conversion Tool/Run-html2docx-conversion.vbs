Set kbFSO = CreateObject("Scripting.FileSystemObject")
Set kbFldr = kbFSO.GetFolder("C:\ADDHERE")

Set kbWord = CreateObject("Word.Application")
kbWord.Visible = False

For Each kbFile In kbFldr.Files
    If LCase(kbFSO.GetExtensionName(kbFile.Name)) = "html" Then
        Set kbDoc = kbWord.Documents.Open(kbFile.path)
        kbWord.ActiveDocument.SaveAs kbFile.path & ".docx", 12
        kbDoc.Close
    End If
Next

kbWord.Quit
