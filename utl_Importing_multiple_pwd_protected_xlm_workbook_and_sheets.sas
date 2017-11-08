Removing excel workbook and worksheet passwords

 Sheet 1 cell B1 has the worbook password
 Sheet 1 cell B2 has the sheet password (same for all sheets)

see
https://goo.gl/PHNwU2
https://communities.sas.com/t5/SAS-Enterprise-Guide/Importing-multiple-pwd-protected-xlfiles-from-a-URL-at-the-same/m-p/411446

Sub ChDirNet(szPath As String)
    SetCurrentDirectoryA szPath
End Sub

Sub RemovePasswords()

    Dim SaveDriveDir As String
    Dim FName As Variant
    Dim FNum As Long
    Dim mybook As Workbook, ws As Worksheet
    Dim workbookpass As String
    Dim sheetpass As String


    workbookpass = ThisWorkbook.Sheets(1).Range("B1").Value
    sheetpass = ThisWorkbook.Sheets(1).Range("B2").Value

    SaveDriveDir = CurDir
    ChDir ThisWorkbook.Path

    FName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xl*), *.xl*", _
                                        MultiSelect:=True)

    If IsArray(FName) Then
        For FNum = LBound(FName) To UBound(FName)
                Set mybook = Nothing
                Set mybook = Workbooks.Open(FName(FNum), UpdateLinks:=0, ReadOnly:=False, Password:=workbookpass)

                On Error GoTo 0
        If Not mybook Is Nothing Then
            On Error Resume Next
            With mybook
                mybook.Unprotect Password:=sheetpass
                For Each ws In Worksheets
                    ws.Unprotect Password:=sheetpass
                Next ws
            End With
    Application.DisplayAlerts = False
            mybook.SaveAs Password:=""
            mybook.Close
    Application.DisplayAlerts = True
        End If
        Next FNum
    End If

End Sub

