Sub Export_ALL_DistributionListToExcel()
 
    '--> Create some constants
    Const SCRIPT_NAME = "Export Lists to Excel"
    Const BASED_ON = "Export Distribution Lists by TechnicLee"
     
    '--> Create some variables
    Dim olkLst As Object, _
        olkRcp As Outlook.Recipient, _
        excApp As Object, _
        excWkb As Object, _
        excWks As Object, _
        intCount As Integer, _
        strFilename As String, _
    strAdr As String, _
    olkFol As Object, _
        lngRow As Long
        
    '--> Initialize variables
    lngRow = 2

    '--> Connect to Excel
    Set excApp = CreateObject("Excel.Application")
    Set excWkb = excApp.Workbooks.Add
    Set excWks = excWkb.Worksheets(1)
    'Write headers
        With excWks
            .Cells(1, 1) = "Client"
            .Cells(1, 2) = "User"
            .Cells(1, 3) = "Address"
        End With
           
     
    '--> Main routine
    Set olkFol = Application.ActiveExplorer.CurrentFolder
        If olkFol.DefaultItemType = olContactItem Then
            For Each olkLst In olkFol.Items
                If olkLst.Class = olDistributionList Then
                    excWks.Cells(lngRow, 1) = olkLst.DLName & "   " & olkLst.MemberCount
                    'Read the list members and write them to the spreadsheet
                    For intCount = 1 To olkLst.MemberCount
                        Set olkRcp = olkLst.GetMember(intCount)
                        excWks.Cells(lngRow, 2) = olkRcp.Name
                        excWks.Cells(lngRow, 3) = olkRcp.Address
                        lngRow = lngRow + 1
                    Next
                End If
            Next
        Else
            MsgBox "The selected folder is not a contacts folder.  Operation cancelled.", vbInformation + vbOKOnly, MACRO_NAME
        End If

    '--> Complete Excel Format and Save
    'Autofit the columns
    excWks.Columns("A:B").AutoFit
        'Get a file path/name to save the spreadsheet to
        strFilename = InputBox("Enter a path and file name for this export", SCRIPT_NAME, Environ("UserProfile") & "\Desktop\GPSupport.xlsx")
        'Did we get a file path/name?
        If strFilename = "" Then
            'No
            'Set the file path to your Documents folder and the file name to the name of the list.
            strFilename = Environ("UserProfile") & "\Desktop\" & olkLst.Subject & ".xlsx"
        Else
            'Yes
            'If the file extension isn't .xlsx
            If Right(LCase(strFilename), 5) <> ".xlsx" Then
                'Set the extention so .xlsx
                strFilename = strFilename & ".xlsx"
            End If
        End If
        'Close and save the spreadsheet
        excWkb.Close True, strFilename
        'Did the file save okay?
        If Err.Number = 0 Then
            'Yes
            MsgBox "Export complete.", vbInformation + vbOKOnly, SCRIPT_NAME
        Else
            'No
            'Make Excel visible so the user cansave the file
            excApp.Visible = True
        End If

    '--> Clean-up
    Set excWks = Nothing
    Set excWkb = Nothing
    Set excApp = Nothing
    Set olkRcp = Nothing
    Set olkLst = Nothing
    Set olkFol = Nothing

End Sub
