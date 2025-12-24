
<%

Private Function GetMcUser(UserId)

    Dim McUserRs
    Dim LocalName
    LocalName = ""

    Set McUserRs = Server.CreateObject("ADODB.Recordset")
	McUserRs.ActiveConnection = Session("ConnMcLogon")
	McUserRs.Source = "Select UserName From Users Where (Id = " & UserId & ")" 
	McUserRs.CursorType = Application("adOpenForwardOnly")
	McUserRs.CursorLocation = Application("adUseClient") 
	McUserRs.LockType = Application("adLockReadOnly")
	McUserRs.Open
	Set McUserRs.ActiveConnection = Nothing

    If McUserRs.BOF = True Or McUserRs.EOF = True Then
        LocalName = ""
    Else
        LocalName = McUserRs("UserName")
    End If
    
    McUserRs.Close
    Set McUserRs = Nothing

    GetMcUser = LocalName

End Function



%>

