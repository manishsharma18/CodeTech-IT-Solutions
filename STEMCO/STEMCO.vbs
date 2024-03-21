' strfolderpath="C:\Users\HardikSingh\Desktop\NEW_INVOICES(26-DEC)\STEMCO\Text|C:\Users\HardikSingh\Desktop\NEW_INVOICES(26-DEC)\SCRIPTS_26DEC\Hendrickson_Template (1).xlsx|C:\Users\HardikSingh\Desktop\NEW_INVOICES(26-DEC)\STEMCO|C:\Users\HardikSingh\Desktop\NEW_INVOICES(26-DEC)\BF|C:\Users\HardikSingh\Desktop\NEW_INVOICES(26-DEC)\log.csv|C:\Users\HardikSingh\Desktop\NEW_INVOICES(26-DEC)\Exception"
strfolderpath = "C:\Users\lavis\OneDrive\Desktop\STEMCO\pdf to text|C:\Users\lavis\OneDrive\Desktop\STEMCO\temp.xlsx|C:\Users\lavis\OneDrive\Desktop\STEMCO\pdf|C:\Users\lavis\OneDrive\Desktop\STEMCO\New folder|C:\Users\lavis\OneDrive\Desktop\STEMCO\log.csv|C:\Users\lavis\OneDrive\Desktop\STEMCO\New"

Function STEMCO(strfolderpath)
On Error Resume Next
filepath=Split(strfolderpath,"|")(0)
excelfile=Split(strfolderpath,"|")(1)
pdfFolder = Split(strfolderPath, "|")(2)
BusinessFolder = Split(strfolderPath, "|")(3)    ' & "\" & "STEMCO SAMPLES"
LogFile = Split(strfolderPath, "|")(4)
ExceptionFolder = Split(strfolderPath, "|")(5)

Set objFSO=CreateObject("Scripting.FileSystemObject")
Set objFolder=objFSO.GetFolder(filepath)

Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile(LogFile, 8, True, 0)

Set objexcel=CreateObject("Excel.Application")
Set destWorkbook = objExcel.Workbooks.Open(excelfile)
objExcel.visible=True
Set destWorksheet = destWorkbook.Worksheets(1)
rowCount = destworksheet.UsedRange.Rows.Count
targetRow = rowCount + 1
For Each objFile in objFolder.Files 
msgbox "hi"
    currentTime = "("&Now&")"
    filePath2 = objFile
    fileName = fso.GetFileName(filePath2)
    Info = "Info"
    scriptName="STEMCO.vbs"
    file.WriteLine currentTime & "," & scriptName & "," & "INFO" & "," & filePath2  & " - Start"
    Set objfile=objFSO.OpenTextFile(objFile.path,1)
    filecontents=objfile.ReadAll
    line=Split(filecontents,vblf)

    set reg_space=New RegExp
    reg_space.Pattern="\s{2,}"
    reg_space.Global=True

    V_company="STEMCO"
 
    '**************************************invoice number********************************************
    V_invoicenumber=""
    for i=0 To UBound(line)
        If Instr(1,line(i),"invoice",1) And Instr(1,line(i),"#",1) Then
            startpos=Instr(1,line(i),"#",1)
            startpos=startpos+len("#")
            V_invoicenumber=Trim(Replace(Mid(line(i),startpos),vbcr,""))
            Exit For
        End If
    Next
    '  MsgBox V_invoicenumber,vbInformation,"INVOICE DATE"

    '***************************************invoice date*********************************************
    V_invoicedate=""
    for i=0 To UBound(line)
        If Instr(1,line(i),"invoice",1) And Instr(1,line(i),"date",1) Then
            startpos=Instr(1,line(i),"date",1)
            startpos=startpos+len("date")
            V_invoicedate=Trim(Replace(Mid(line(i),startpos),vbcr,""))
            Exit For
        End If
    Next
    ' MsgBox V_invoicedate,vbInformation,"INVOICE DATE"
    
    '*******************************ORDER NUMBER*****************************************************
    V_ordernumber=""
    for i=0 To UBound(line)
        If Instr(1,line(i),"order",1) And Instr(1,line(i),"#",1) Then
            startpos=Instr(1,line(i),"#",1)
            startpos=startpos+len("#")
            V_ordernumber=Trim(Replace(Mid(line(i),startpos),vbcr,""))
            Exit For
        End If
    Next
    ' MsgBox V_ordernumber,vbInformation,"ORDER NUMBER"

    '*******************************DELIVERY NUMBER*****************************************************
    V_deliverynumber=""
    for i=0 To UBound(line)
        If Instr(1,line(i),"delivery",1) And Instr(1,line(i),"#",1) Then
            startpos=Instr(1,line(i),"#",1)
            startpos=startpos+len("#")
            V_deliverynumber=Trim(Replace(Mid(line(i),startpos),vbcr,""))
            Exit For
        End If
    Next
    ' MsgBox V_deliverynumber,vbInformation,"DELIVERY NUMBER"

    '***********************************PO NUMBER***************************************************
    V_ponumber=""
    for i=0 To UBound(line)
        If Instr(1,line(i),"Bill-To",1) And Instr(1,line(i),"po",1) Then
            V_ponumber=Split(Trim(line(i+1)),"  ")(0)
            Exit For
        End If
    Next
    ' MsgBox V_ponumber,vbInformation,"PO NUMBER"

    '***********************************part details***********************************************
    redim V_partdetails(-1)
    counter=-1

    set reg_dec=New RegExp
    reg_dec.Pattern="[\d\,]+\.[\d\,\-]+"
    reg_dec.Global=True

    For i=0 To Ubound(line)
        If Instr(1,line(i),"subtotal",1) Then
            Exit For
        End If

        set mtc=reg_dec.Execute(line(i))
        If mtc.Count>1 Then
            part_name=""
            uc=""
            qt=""
            tc=""
            Redim Preserve V_partdetails(UBound(V_partdetails)+1)
            counter=counter+1
            line1=reg_space.Replace(Trim(line(i)),"  ")
            list1=Split(line1,"  ")
            If Ubound(list1)>3 Then
                tc=list1(UBound(list1))
                qt=list1(UBound(list1)-1)
                uc=list1(UBound(list1)-2)
                part_name=list1(UBound(list1)-3)
                set tc1=reg_dec.Execute(tc)
                If tc1.Count>0 Then
                    tc=tc1(0).Value
                End If
                set uc1=reg_dec.Execute(uc)
                If uc1.Count>0 Then
                    uc=uc1(0).Value
                End If
            End If
            V_partdetails(counter)=part_name&"|"&qt&"|"&uc&"|"&tc
            'MsgBox V_partdetails(counter)
        End If
    Next

    '*********************writing data to excel***************************************************
    ' t=targetRow
    ' For j=0 To counter
    '     destWorksheet.Cells(targetRow,"A").value="abcde"
    '     destWorksheet.Cells(targetRow, "A").Value = V_company
    '     values=Split(V_partdetails(j),"|")
    '     destWorksheet.Cells(targetRow, "B").Value = Replace(V_invoicenumber,vbcr,"")
    '     destWorksheet.Cells(targetRow, "C").Value = Replace(V_invoicedate,vbcr,"")
    '     destWorksheet.Cells(targetRow, "D").Value = Replace(values(1),vbcr,"")
    '     destWorksheet.Cells(targetRow, "E").Value = Replace(values(2),vbcr,"")
    '     destWorksheet.Cells(targetRow, "F").Value =Replace(values(3),vbcr,"")
    '     destWorksheet.Cells(targetRow, "G").Value =Replace(V_ponumber,vbcr,"")
    '     destWorksheet.Cells(targetRow, "J").Value =Replace(V_deliverynumber,vbcr,"")
    '     destWorksheet.Cells(targetRow, "L").Value =Replace(V_ordernumber,vbcr,"")
    '     destWorksheet.Cells(targetRow, "N").Value =Replace(values(0),vbcr,"")
    '     targetRow=targetRow+1

    ' Next

    '****************************renaming*******************************************************************
    ' Dim filesys
    ' Set filesys = CreateObject("Scripting.FileSystemObject")
    ' FileName = filesys.GetFileName(filePath2)
    ' FileName = Replace(FileName,".txt",".pdf")
    ' sourcefile = pdfFolder & "\" & Filename                                                                                 
    ' destinationfile = BusinessFolder & "\" & Ford & "\" & values(0) & "\" &V_invoicedate& "\" & & "\" & &  ".pdf"
    ' sourcefile = Replace(sourcefile,vbLf,"")
    ' sourcefile = Replace(sourcefile,vbCr,"")
    ' If fso.FileExists(destinationfile) Then ' Generate a unique file name for the copy
    '     intCounter = 1
    '     Do Until Not fso.FileExists(destinationfile)
    '         destinationfile = BusinessFolder & "\" & V_invoicenumber & " " & V_company & "_" & intCounter & ".pdf"
    '         destinationfile = Replace(destinationfile,vbLf,"")
    '         destinationfile = Replace(destinationfile,vbCr,"")
    '         intCounter = intCounter + 1
    '     Loop
    ' End If
    ' destinationfile = Replace(destinationfile,vbLf,"")
    ' destinationfile = Replace(destinationfile,vbCr,"")
    currentTime = "("&Now&")"
    If err.number = 0 Then
        file.WriteLine currentTime & "," & scriptName & "," & "INFO" & "," & filePath2 & "  : END -pdf File Copy From : " & sourcefile & " to : " & destinationfile &  " - Success"
        For t=t To targetRow-1
            destWorksheet.Range("T" &t).value = destinationfile
        Next
    Else
        file.WriteLine currentTime & "," & scriptName & "," & "INFO" & "," & filePath2 & "  : END "  & " - Failed at line No. - " & Err.Line & " error message - " & Err.Description
        Set filesys = CreateObject("Scripting.FileSystemObject")
        FileName = filesys.GetFileName(filePath2)
        FileName = Replace(FileName,".txt",".pdf")
        sourcefile = pdfFolder & "\" & Filename
        destinationfile = ExceptionFolder & "\" & Filename
        For t=t To targetRow-1
            destWorksheet.Range("T" &t).value = destinationfile
        Next
        err.number = clear
    End If
    filesys.CopyFile sourcefile, destinationfile, True
    err.clear

Next   

destWorkbook.Save
' destWorkbook.Close
' objexcel.Quit
Set destWorksheet=Nothing
Set destWorkbook=Nothing
Set objexcel=Nothing

If Err.NUMBER=0 Then
    STEMCO="Success"
Else
    STEMCO="Failure"
End If


End Function
call STEMCO(strfolderpath)

