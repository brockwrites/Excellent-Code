Sub FindContactInList()
    Const MACRO_NAME = "Find Contact in List"
    Dim olkFol As Object, _
        olkLst As Object, _
        strAdr As String, _
        strMat As String
    strAdr = InputBox("Enter the SMTP address to serch for.", "Enter Address")
    If strAdr = "" Then
        MsgBox "You did not enter an address.  Operation cancelled.", vbInformation + vbOKOnly, MACRO_NAME
    Else
        Set olkFol = Application.ActiveExplorer.CurrentFolder
        If olkFol.DefaultItemType = olContactItem Then
            For Each olkLst In olkFol.Items
                If olkLst.Class = olDistributionList Then
                    If IsMember(olkLst, strAdr) Then
                        strMat = strMat & olkLst.DLName & vbCrLf
                    End If
                End If
            Next
            If strMat = "" Then
                MsgBox "The address " & strAdr & " is not in an list in this folder.", vbInformation + vbOKOnly, MACRO_NAME
            Else
                MsgBox "The address " & strAdr & " is in the following lists " & vbCrLf & vbCrLf & strMat, vbInformation + vbOKOnly, MACRO_NAME
            End If
        Else
            MsgBox "The selected folder is not a contacts folder.  Operation cancelled.", vbInformation + vbOKOnly, MACRO_NAME
        End If
    End If
    Set olkFol = Nothing
    Set olkLst = Nothing
End Sub
