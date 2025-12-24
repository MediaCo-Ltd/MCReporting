<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("ConnMcLogon") = "" Then Response.Redirect("Login.asp")

Dim Error
Dim UserId
Dim Uname
Dim UpdateOk
Dim ResetCount


Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl

UserId = Cint(Request.QueryString("Uid"))

RedirectUrl = "Login.asp"
   
'############################## Reset User Password ##############################  

    UpdateSql = "SELECT * FROM Users Where (Id = " & UserId & ")" 

    Set UpdateRs = Server.CreateObject("ADODB.Recordset")
    UpdateRs.ActiveConnection = Session("ConnMcLogon")
    UpdateRs.Source = UpdateSql
    UpdateRs.CursorType = Application("adOpenStatic")
    UpdateRs.CursorLocation = Application("adUseClient")
    UpdateRs.LockType = Application("adLockOptimistic")
    UpdateRs.Open
    
    Uname = UpdateRs("UserName")
        
    UpdateRs("Active") = 1
    UpdateRs("ShowInList") = 1
    UpdateRs("Password") = ""
    ResetCount = UpdateRs("PwResets") + 1
    UpdateRs("PwResets") = ResetCount 
    

    UpdateRs.Update
    UpdateRs.MoveLast
    
    Set UpdateRs.ActiveConnection = Nothing
    UpdateRs.Close
    Set UpdateRs = Nothing    
    
    If Err = 0 Then
        UpdateOk = True
        RedirectUrl = "login.asp"
        UpdateResets
    Else
        Session("SystemError") = "Rest user page - reset Password"
        Response.Redirect ("SystemError.asp")       
    End If
        
    Response.Redirect RedirectUrl  
    
'##################################################################################    
    
Private Sub UpdateResets()

    Dim ResetSql
    Dim ResetRs
    Dim ResetWsNetwork
    Dim ResetComputerName
    
    ResetSql = "Select * From PwResets"

    Set ResetRs = Server.CreateObject("ADODB.Recordset")
    ResetRs.ActiveConnection = Session("ConnMcLogon")
    ResetRs.Source = ResetSql
    ResetRs.CursorType = Application("adOpenStatic")
    ResetRs.CursorLocation = Application("adUseClient")
    ResetRs.LockType = Application("adLockOptimistic")
    ResetRs.Open

    ResetRs.AddNew
    
    ResetRs("UserId") = UserId
    ResetRs("UserName") = Uname
    ResetRs("ResetDate") = Now
    
    Set ResetWsNetwork = createobject("wscript.network")
    ResetComputerName = ResetWsNetwork.computername
    Set ResetWsNetwork = Nothing
    ResetRs("Resetlocation") = ResetComputerName    
    
    ResetRs.Update
    ResetRs.MoveLast
    ResetRs.Close
    Set ResetRs = Nothing

End Sub               

%>
