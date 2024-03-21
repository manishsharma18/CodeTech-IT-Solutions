' Define the folder path containing the text files
strFolderPath = "C:\Users\lavis\OneDrive\Desktop\STEMCO\pdf to text"

' Get a list of text files in the folder
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strFolderPath)
Set colFiles = objFolder.Files

' Iterate over each text file in the folder
For Each objF In colFiles
    If LCase(objFSO.GetExtensionName(objF)) = "txt" Then ' Check if it's a text file
        ' Open the text file for reading
        Set objFile = objFSO.OpenTextFile(objF.Path, 1)
        

    ' Set objFSO = CreateObject("Scripting.FileSystemObject") 
    ' Set objFile = objFSO.OpenTextFile("C:\Users\lavis\OneDrive\Desktop\STEMCO\pdf to text\Invoice_12300966.txt", 1) 

    ' filepath = C:\Users\lavis\OneDrive\Desktop\STEMCO\pdf to text
    ' Set objFSO = CreateObject("Scripting.FileSystemObject") 
    ' Set objFolder=objFSO.GetFolder(filepath)
    ' For Each objFile in objFolder.Files

arr=Array("Invoice #","Invoice Date","Order #","Delivery #")

Dim fetch
ReDim fetch(UBound(arr))
Do Until objFile.AtEndOfStream   
  strLine = objFile.ReadLine   

    for i=0 to UBound(arr)
        If InStr(strLine, arr(i)) > 0 Then 
            fetch(i) = Trim(Mid(strLine, InStr(strLine, arr(i)) + Len(arr(i))))   
            Exit For     
        End If 
    next
Loop   

' for i=0 to UBound(fetch)
'     msgbox fetch(i)  
' Next


Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True

' Open the workbook
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\lavis\OneDrive\Desktop\STEMCO\invoice excel.xlsx")

' Select the first worksheet
Set objWorksheet = objWorkbook.Worksheets(1)

' Find the first empty row in column A
intRow = objWorksheet.Cells(objWorksheet.Rows.Count, 1).End(-4162).Row + 1

' Write the fetched data to the worksheet
For i = 0 To UBound(fetch)
    objWorksheet.Cells(intRow, i + 1).Value = fetch(i)
Next

' Save and close the workbook
objWorkbook.Save
objWorkbook.Close

' Clean up
objExcel.Quit
Set objWorksheet = Nothing
Set objWorkbook = Nothing
Set objExcel = Nothing
End If
Next

' ' Release the object variables
Set objFile = Nothing
Set objFSO = Nothing
