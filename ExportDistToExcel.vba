'From https://techniclee.wordpress.com/2013/05/13/exporting-an-outlook-distribution-list-to-excel/
'
Option Explicit
 
Sub ExportDistributionListToExcel()
 
    '--> Create some constants
    Const SCRIPT_NAME = "Export Distribution List to Excel"
     
    '--> Create some variables
    Dim olkLst As Object, _
        olkRcp As Outlook.RECIPIENT, _
        excApp As Object, _
        excWkb As Object, _
        excWks As Object, _
        intCount As Integer, _
        lngRow As Long, _
        strFilename As String
         
    '--> Initialize variables
    lngRow = 2
     
    '--> Main routine
    'Turn error handling off
    On Error Resume Next
    'What type of window is open?
    Select Case TypeName(Application.ActiveWindow)
        Case "Explorer"
            Set olkLst = Application.ActiveExplorer.Selection(1)
        Case "Inspector"
            Set olkLst = Application.ActiveInspector.CurrentItem
        Case Else
            Set olkLst = Nothing
    End Select
    'Was a list open or selected?
    If TypeName(olkLst) = "Nothing" Then
        'No
        MsgBox "You must select or open an item for this macro to work.", vbCritical + vbOKOnly, SCRIPT_NAME
    Else
        'Yes
        'Is the open/selected item a dist list?
        If olkLst.Class = olDistributionList Then
            'Yes
            'Connect to Excel
            Set excApp = CreateObject("Excel.Application")
            Set excWkb = excApp.Workbooks.Add
            Set excWks = excWkb.Worksheets(1)
            'Write headers
            With excWks
                .Cells(1, 1) = "Name"
                .Cells(1, 2) = "Address"
            End With
            'Read the list members and write them to the spreadsheet
            For intCount = 1 To olkLst.MemberCount
                Set olkRcp = olkLst.GetMember(intCount)
                excWks.Cells(lngRow, 1) = olkRcp.Name
                excWks.Cells(lngRow, 2) = olkRcp.Address
                lngRow = lngRow + 1
            Next
            'Autofit the columns
            excWks.Columns("A:B").AutoFit
            'Get a file path/name to save the spreadsheet to
            strFilename = InputBox("Enter a path and file name for this export", SCRIPT_NAME, Environ("UserProfile") & "\My Documents\" & olkLst.Subject & ".xlsx")
            'Did we get a file path/name?
            If strFilename = "" Then
                'No
                'Set the file path to your Documents folder and the file name to the name of the list.
                strFilename = Environ("UserProfile") & "\My Documents\" & olkLst.Subject & ".xlsx"
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
        Else
            'No
            MsgBox "The item you selected is not a distribution list.  Export cancelled.", vbCritical + vbOKOnly, SCRIPT_NAME
        End If
    End If
     
    '--> Clean-up
    Set excWks = Nothing
    Set excWkb = Nothing
    Set excApp = Nothing
    Set olkRcp = Nothing
    Set olkLst = Nothing
    'Turn error handling back on
    On Error GoTo 0
End Sub 
