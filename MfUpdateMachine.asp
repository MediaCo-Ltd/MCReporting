<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("ConnMachinefaults") = "" Then Response.Redirect("Admin.asp")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count

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
    If Request.Form("frmName") = "frmAddMachine" Then  

    '############################## Create client 

        AddSql = "SELECT * FROM Machine" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnMachinefaults")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs.AddNew

        AddRs("MachineName") = Cstr(Request.Form("txtMachineID"))        
        AddRs("MachineTypeId") = Cint(Request.Form("hMachineType"))        
        ReturnedTypeName = GetTypeName(Cint(Request.Form("hMachineType")))        
        AddRs("MachineType") = ReturnedTypeName
        AddRs("FaultGroups") = Cstr(Request.Form("hGroupId"))  
        
        If Request.Form("chkActive") = "on" Then  
            AddRs("Active") = 1
        Else
            AddRs("Active") = 0
        End If
        
        AddRs.Update
        AddRs.MoveLast

               
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Add Machine page - Add Machine"
            Response.Redirect ("SystemError.asp")       
        End If
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing     
        
    
'############################## Edit Machine Update ##############################
 
    ElseIf Request.Form("frmName") = "frmEditMachine" Then

        UpdateSql = "SELECT * FROM Machine Where (Id = " & Request.Form("hId") & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMachinefaults")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        UpdateRs("MachineName") = Cstr(Request.Form("txtMachineID"))        
        UpdateRs("MachineTypeId") = Cint(Request.Form("hMachineType"))        
        ReturnedTypeName = GetTypeName(Cint(Request.Form("hMachineType")))        
        UpdateRs("MachineType") = ReturnedTypeName
        
        '## Only Update groups if changed
        If Request.Form("hGroupId") <> "" Then
            UpdateRs("FaultGroups") = Cstr(Request.Form("hGroupId"))
        End If
       
        If Request.Form("chkActive") = "on" Then
            UpdateRs("Active") = 1
        Else
            UpdateRs("Active") = 0
        End If        
        
        UpdateRs.Update
        UpdateRs.MoveLast    
        
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Edit Machine page - Update Machine"
            Response.Redirect ("SystemError.asp")       
        End If
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing

'############################## Edit Machine Status Update ##############################
        
    ElseIf Request.Form("frmName") = "frmMachineStatus" Then
   
        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMachinefaults")
        UpdateRs.Source = "Select Id, Active From Machine Order By Id"
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open
        
        UpdateRs.MoveFirst
        While Not UpdateRs.EOF
            If Request.Form(Cstr("Active" & UpdateRs.AbsolutePosition)) = "on" Then
                UpdateRs("Active") = 1
            Else
                UpdateRs("Active") = 0
            End If            
            UpdateRs.MoveNext
        Wend        
        
        UpdateRs.Close
        Set UpdateRs = Nothing
        
    Else
        Response.Redirect (RedirectUrl)
    End If
End If

Response.Redirect (RedirectUrl)

'##########################################################

Function GetTypeName(TypeId)

Dim MachineTypeRs
Dim MachineTypeSql
Dim LocalTypeName
LocalTypeName = ""

MachineTypeSql = "SELECT DISTINCT MachineTypeId, MachineType From Machine"
MachineTypeSql = MachineTypeSql & " Where (MachineTypeId = " & TypeId & ")"

Set MachineTypeRs = Server.CreateObject("ADODB.Recordset")
MachineTypeRs.ActiveConnection = Session("ConnMachinefaults")
MachineTypeRs.Source = MachineTypeSql
MachineTypeRs.CursorType = Application("adOpenForwardOnly")
MachineTypeRs.CursorLocation = Application("adUseClient")
MachineTypeRs.LockType = Application("adLockReadOnly")
MachineTypeRs.Open
Set MachineTypeRs.ActiveConnection = Nothing

If MachineTypeRs.BOF = True Or MachineTypeRs.EOF = True Then
    LocalTypeName = ""
Else
    LocalTypeName = MachineTypeRs("MachineType")
End If

MachineTypeRs.Close
Set MachineTypeRs = Nothing

GetTypeName = LocalTypeName

End function
%>
