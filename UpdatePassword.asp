<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("ConnMcLogon") = "" Then Response.Redirect("CloseWindow.asp")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count

Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl


If Request.Form("chkUpdate") = "on" Then
    '## Testing don't save data, will only be on if at home or on my pc at work
Else
   
'############################## Edit User Update ##############################
 
    If Request.Form("frmName") = "frmAddPW" Then

        UpdateSql = "SELECT * FROM Users Where (Id = " & Request.Form("txtUserID") & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMcLogon")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        UpdateRs("Active") = 1
        UpdateRs("ShowInList") = 0
        UpdateRs("Password") = Trim(Request.Form("txtPassword"))
        
        UpdateRs("ShowHS") = 1
        UpdateRs("ShowNC") = 1
        
        UpdateRs.Update
        UpdateRs.MoveLast
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing    
        
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Add password page - Update user"
            Response.Redirect ("SystemError.asp")       
        End If        
        
        
        '## Update user count on system
        UpdateSql = "SELECT Id, UserCount FROM System" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMcLogon")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        UpdateRs("UserCount") = UpdateRs("UserCount") +1
        
        UpdateRs.Update    
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing 
        
    Else
        Response.Redirect ("CloseWindow.asp")
    End If
End If

Response.Redirect ("CloseWindow.asp")
%>
