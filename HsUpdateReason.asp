<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("ConnHSReports") = "" Then Response.Redirect("Admin.asp")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count

Dim NewId
Dim NewIdRs

Dim AddRs
Dim AddSql
Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl
Dim ReturnedTypeName
ReturnedTypeName = ""

RedirectUrl = "Admin.asp"

If Request.Form("chkUpdate") = "on" Then
    '## Testing don't save data, will only be on if at home or on my pc at work
Else
'############################## Add New User ##############################
    If Request.Form("frmName") = "frmAddReason" Then  

    '############################## Create client 

        AddSql = "SELECT * FROM Reasons" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnHSReports")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open
        AddRs.AddNew

        AddRs("Description") = Cstr(Request.Form("txtReasonID"))        
        AddRs("GroupId") = Cint(Request.Form("hReasonType")) 
        
        If Cint(Request.Form("hReasonType")) = 1 Then
            AddRs("Usage") = "Accident"
        ElseIf Cint(Request.Form("hReasonType")) = 2 Then
            AddRs("Usage") = "Incident"
        Else
            AddRs("Usage") = "Unsafe"
        End If        
        
        If Request.Form("chkActive") = "on" Then
            AddRs("Active") = 1
        Else
            AddRs("Active") = 0
        End If    
        
        AddRs.Update
        AddRs.MoveLast
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing         
               
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Add Reason page - Add Reason"
            Response.Redirect ("SystemError.asp")       
        End If
    
'############################## Edit Reason Update ##############################
 
    ElseIf Request.Form("frmName") = "frmEditReason" Then

        UpdateSql = "SELECT * FROM Reasons Where (Id = " & Request.Form("gId") & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnHSReports")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        '## Still to do any other changes
       
        If Request.Form("chkActive") = "on" Then
            UpdateRs("Active") = 1
        Else
            UpdateRs("Active") = 0
        End If        
        
        UpdateRs.Update
        UpdateRs.MoveLast 
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing   
        
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Edit Reason page - Update Reason"
            Response.Redirect ("SystemError.asp")       
        End If        
        
    Else
        Response.Redirect (RedirectUrl)
    End If
End If

Response.Redirect (RedirectUrl)

%>
