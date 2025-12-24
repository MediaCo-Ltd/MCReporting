<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("ConnMcLogon") = "" Then Response.Redirect("Admin.asp")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count

Dim AddRs
Dim AddSql
Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl

RedirectUrl = "Admin.asp"

If Request.Form("chkUpdate") = "on" Then
    '## Testing don't save data, will only be on if at home or on my pc at work
Else
'############################## Add New User ##############################
    If Request.Form("frmName") = "frmAddUser" Then  

    '############################## Create client 

        AddSql = "SELECT * FROM Users" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnMcLogon")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs.AddNew

        AddRs("UserName") = Cstr(Request.Form("txtUserID"))
        
        If Cstr(Request.Form("txtPassword")) = "0" Then
            '## New user set so shows in list
            AddRs("Active") = 0
            AddRs("ShowInList") = 1
        Else
            AddRs("Password") = Cstr(Request.Form("txtPassword"))
            If Request.Form("chkActive") = "on" Then  
                AddRs("Active") = 1
                AddRs("ShowInList") = 0
            Else
                AddRs("Active") = 0
                AddRs("ShowInList") = 1
            End If
        End If      
                     
        If Request.Form("chkAdmin") = "on" Then  
            AddRs("AdminUser") = 1            
            AddRs("HsEdit") = 1
            AddRs("NcEdit") = 1
            AddRs("MfEdit") = 1
        Else
            AddRs("AdminUser") = 0
            
            If Request.Form("ChkHsEdit") = "on" Then  
                AddRs("HsEdit") = 1
            Else
                AddRs("HsEdit") = 0
            End If
            
            If Request.Form("ChkNcEdit") = "on" Then  
                AddRs("NcEdit") = 1
            Else
                AddRs("NcEdit") = 0
            End If
            
            If Request.Form("ChkMfEdit") = "on" Then  
                AddRs("MfEdit") = 1
            Else
                AddRs("MfEdit") = 0
            End If
            
        End If
        
        If Request.Form("chkHS") = "on" Then  
            AddRs("ShowHS") = 1
        Else
            AddRs("ShowHS") = 0
            AddRs("HsEdit") = 0
        End If
        
        If Request.Form("chkNC") = "on" Then  
            AddRs("ShowNC") = 1
        Else
            AddRs("ShowNC") = 0
            AddRs("NcEdit") = 0
        End If
        
        If Request.Form("chkMF") = "on" Then  
            AddRs("ShowMF") = 1
        Else
            AddRs("ShowMF") = 0
            AddRs("MfEdit") = 0
        End If
        
        AddRs.Update
        AddRs.MoveLast
               
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Add user page - Add user"
            Response.Redirect ("SystemError.asp")       
        End If
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing     
    
    
        '## Update user count on system
        AddSql = "SELECT Id, UserCount FROM System" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnMcLogon")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs("UserCount") = AddRs("UserCount") +1
        
        AddRs.Update    
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing        

        If Err = 0 Then
            UpdateOk = True
            RedirectUrl = "UsersOnline.asp"
        Else
            Session("SystemError") = "Add user page - Update User Count"
            Response.Redirect ("SystemError.asp")       
        End If
    
'############################## Edit User Update ##############################
 
    ElseIf Request.Form("frmName") = "frmEditUser" Then

        UpdateSql = "SELECT * FROM Users Where (Id = " & Request.Form("hUid") & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMcLogon")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        UpdateRs("UserName") = Cstr(Request.Form("txtUserID"))        
        
        If Request.Form("hDelete") = "1" Then        
            UpdateRs("Active") = 0            
            UpdateRs("AdminUser") = 0
            UpdateRs("ShowHS") = 0
            UpdateRs("ShowNC") = 0
            UpdateRs("ShowMF") = 0
            UpdateRs("ShowInList") = 0
            UpdateRs("Dormant") = 1 
            UpdateRs("Password") = ""
            UpdateRs("HsEdit") = 0
            UpdateRs("NcEdit") = 0
            UpdateRs("MfEdit") = 0       
        Else      
            If Cstr(Request.Form("txtPassword")) = "0" Then
                UpdateRs("Active") = 0
                UpdateRs("ShowInList") = 1
            Else
                UpdateRs("Password") = Cstr(Request.Form("txtPassword"))
                If Request.Form("chkActive") = "on" Then  
                    UpdateRs("Active") = 1
                    UpdateRs("ShowInList") = 0
                Else
                    UpdateRs("Active") = 0
                    UpdateRs("ShowInList") = 1
                End If
            End If            
            
            If Request.Form("chkAdmin") = "on" Then  
                UpdateRs("AdminUser") = 1
                UpdateRs("HsEdit") = 1
                UpdateRs("NcEdit") = 1
                UpdateRs("MfEdit") = 1
            Else
                UpdateRs("AdminUser") = 0
                If Request.Form("ChkHsEdit") = "on" Then  
                    UpdateRs("HsEdit") = 1
                Else
                    UpdateRs("HsEdit") = 0
                End If
                
                If Request.Form("ChkNcEdit") = "on" Then  
                    UpdateRs("NcEdit") = 1
                Else
                    UpdateRs("NcEdit") = 0
                End If
                
                If Request.Form("ChkMfEdit") = "on" Then  
                    UpdateRs("MfEdit") = 1
                Else
                    UpdateRs("MfEdit") = 0
                End If
            End If
            
            If Request.Form("chkHS") = "on" Then  
                UpdateRs("ShowHS") = 1
            Else
                UpdateRs("ShowHS") = 0
                UpdateRs("HsEdit") = 0
            End If
            
            If Request.Form("chkNC") = "on" Then  
                UpdateRs("ShowNC") = 1
            Else
                UpdateRs("ShowNC") = 0
                UpdateRs("NcEdit") = 0
            End If
            
            If Request.Form("chkMF") = "on" Then  
                UpdateRs("ShowMF") = 1
            Else
                UpdateRs("ShowMF") = 0
                UpdateRs("MfEdit") = 0
            End If
                
            UpdateRs("Dormant") = 0
                
        End If        
        
        UpdateRs.Update
        UpdateRs.MoveLast
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing    
        
        If Err = 0 Then
            UpdateOk = True
            RedirectUrl = "UsersOnline.asp"
        Else
            Session("SystemError") = "Edit user page - Update client"
            Response.Redirect ("SystemError.asp")       
        End If       
        

'############################## Add Email User Update ##############################
        
    ElseIf Request.Form("frmName") = "frmAddEmailUser" Then
        
        AddSql = "SELECT * FROM Email" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnMcLogon")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs.AddNew

        AddRs("EmailName") = Cstr(Request.Form("txtUserID"))
        AddRs("EmailAddress") = Cstr(Lcase(Request.Form("txtEmail")))
        AddRs("UserId") = Cint(Request.Form("hEmailUserId"))
        
        If Request.Form("chkHS") = "on" Then  
            AddRs("InHSGroup") = 1
        Else
            AddRs("InHSGroup") = 0
        End If
        
        If Request.Form("chkNC") = "on" Then  
            AddRs("InNCGroup") = 1
        Else
            AddRs("InNCGroup") = 0
        End If
        
        If Request.Form("chkMF") = "on" Then  
            AddRs("InMFGroup") = 1
        Else
            AddRs("InMFGroup") = 0
        End If      
        
        'If Request.Form("chkActive") = "on" Then  
        '    AddRs("EmailActive") = 1
        'Else
            AddRs("EmailActive") = 0
        'End If
        
        AddRs.Update
        AddRs.MoveLast
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing
               
        If Err = 0 Then
            UpdateOk = True
            UpdateEmailStatus (Cint(Request.Form("hEmailUserId")))
            RedirectUrl = "EmailUsers.asp"
        Else
            Session("SystemError") = "Add email user page - Add email user"
            Response.Redirect ("SystemError.asp")       
        End If
       

'############################## Edit Email User Update ##############################
    
    ElseIf Request.Form("frmName") = "frmEditEmailUser" Then
    
        UpdateSql = "SELECT * FROM Email Where (Id = " & Request.Form("hUid") & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnMcLogon")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open
        
        UpdateRs("EmailName") = Cstr(Request.Form("txtUserID"))
        UpdateRs("EmailAddress") = Cstr(Lcase(Request.Form("txtEmail")))              
                
        If Request.Form("chkHS") = "on" Then  
            UpdateRs("InHSGroup") = 1
        Else
            UpdateRs("InHSGroup") = 0
        End If
        
        If Request.Form("chkNC") = "on" Then  
            UpdateRs("InNCGroup") = 1
        Else
            UpdateRs("InNCGroup") = 0
        End If
        
        If Request.Form("chkMF") = "on" Then  
            UpdateRs("InMFGroup") = 1
        Else
            UpdateRs("InMFGroup") = 0
        End If   
        
        'If Request.Form("chkActive") = "on" Then  
        '    UpdateRs("EmailActive") = 1
        'Else
            UpdateRs("EmailActive") = 0
        'End If
        
        UpdateRs.Update
        UpdateRs.MoveLast
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing    
        
        If Err = 0 Then
            UpdateOk = True
            RedirectUrl = "EmailUsers.asp"
        Else
            Session("SystemError") = "Edit email user page - Update email user"
            Response.Redirect ("SystemError.asp")       
        End If 
        
    Else
        Response.Redirect (RedirectUrl)
    End If
End If

Response.Redirect (RedirectUrl)


Sub UpdateEmailStatus (Uid)

    Dim NewUserRs

    Set NewUserRs = Server.CreateObject("ADODB.Recordset")
    NewUserRs.ActiveConnection = Session("ConnMcLogon")
    NewUserRs.Source = "SELECT Id, ActiveEmail FROM Users Where (Id = " &  Uid & ")"
    NewUserRs.CursorType = Application("adOpenStatic")
    NewUserRs.CursorLocation = Application("adUseClient")
    NewUserRs.LockType = Application("adLockOptimistic")
    NewUserRs.Open

    NewUserRs("ActiveEmail") = 1
    NewUserRs.Update
    
    NewUserRs.Close

End Sub
%>
