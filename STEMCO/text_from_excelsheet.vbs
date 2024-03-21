Set eo= CreateObject("Excel.Application")
eo.DisplayAlerts=False
Set os= CreateObject("Wscript.Shell")
efp=os.CurrentDirectory
Set wb= eo.Workbooks.Open(efp&"\A360 Training Status.xlsx")
Set ws=wb.Sheets(1)

MsgBox ws.Cells(6,3)

wb.Close False
eo.Quit

Set eo= Nothing
Set os= Nothing
Set wb= Nothing
Set ws= Nothing
