<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\connMCReportsDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\connQualityControlDB.asp" -->

<%

Dim Count
Dim UserRs
Dim LogonErrMsg
Dim strUserID 
Dim strRedirectLoginFailed
Dim strRedirectLoginSuccess

'## Session("ConnQC") = IhRedoDB 
Session("ConnQC") = strConnQualityControlDB
Session("QcPath") = strQualityPath


LogonErrMsg = ""

strRedirectLoginFailed = "Login.asp"
strRedirectLoginSuccess = "SelectSite.asp" 

strUserID = Request.Form("txtUserID")		
If strUserID <> "" Then		
	 
    Set UserRs = Server.CreateObject("ADODB.Recordset")
    UserRs.ActiveConnection = Session("ConnMcLogon")    
    UserRs.Source = "SELECT * "
   
    UserRs.Source = UserRs.Source & " FROM Users WHERE Id=" & strUserID & " AND PassWord='" & Trim(Request.Form("txtPassword")) & "'"
    
    UserRs.CursorType = Application("adOpenForwardOnly")
    UserRs.CursorLocation = Application("adUseClient")
    UserRs.LockType = Application("adLockReadOnly")
    UserRs.Open
	Set UserRs.ActiveConnection = Nothing
	
    If Not UserRs.EOF Or Not UserRs.BOF Then 
	    '## username and password match - this is a valid user
        '## now check whether this account is active
    	
    	Session("UserId") = UserRs("Id") 
    	Session("UserName") = UserRs("UserName")
    	Session("AdminUser") = Cbool(UserRs("AdminUser"))
    	Session("ShowHS") = Cbool(UserRs("ShowHS"))
    	Session("ShowNC") = Cbool(UserRs("ShowNC"))
    	Session("ShowMF") = Cbool(UserRs("ShowMF"))
    	Session("EditHS") = Cbool(UserRs("HsEdit"))
    	Session("EditNC") = Cbool(UserRs("NcEdit"))
    	Session("EditMF") = Cbool(UserRs("MfEdit"))
    	Session("ViewHS") = Cbool(UserRs("HsView"))
    	
    	Session("AddRedo") = Cbool(UserRs("RedoAdd"))
    	Session("ViewRedo") = Cbool(UserRs("RedoView"))
    	Session("ShowRedoCost") = Cbool(UserRs("ShowRedoCost"))
    	Session("UserEditOwnMf") = Cbool(UserRs("UserEditOwnMf"))
    	
    	
    	If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Then 
    	    Session("ShowRiddor") = Cbool(True)
    	Else
    	    Session("ShowRiddor") = Cbool(False)
    	End If
    	
    	Dim BrowserType
        BrowserType = Ucase(Request.ServerVariables("http_user_agent"))

        If Instr(1,BrowserType,"FIREFOX",1) > 0 Then
            Session("Browser") = "Firefox" 
        ElseIf Instr(1,BrowserType,"EDGE",1) > 0 Then
            Session("Browser") = "Chrome"        
        ElseIf Instr(1,BrowserType,"CHROME",1) > 0 And Instr(1,BrowserType,"SAFARI",1) > 0 Then
            Session("Browser") = "Chrome"
        ElseIf Instr(1,BrowserType,"SAFARI",1) > 0 Then
            Session("Browser") = "Safari"
        Else
            Session("Browser") = "Unknown"
        End If
    	
    	If UserRs("Id") = 4 Then
    	    Session("ShowText") = Cbool(True)
    	Else
    	    Session("ShowText") = Cbool(False)
    	End If
    	
        UserRs.Close
        Set UserRs = Nothing	
       
        Response.Cookies("IpChecked") = "True"
        Session("LogonErrMsg") = ""
       
        dim fso
        dim txtfile 
        Set fso = Server.CreateObject("Scripting.FileSystemObject")    
        Set txtfile = fso.OpenTextFile(Session("LogFile") , 8, True)
	    txtfile.write "Session " & Session.SessionId & " begun. " & Session("UserName") & " logged on " & Now() & vbCrLf
	    txtfile.close
        Set txtfile = Nothing
        Set fso = Nothing 
        
        If SystemData = False Then
	        Session("SystemError") = "Can not get system data"
            Response.Redirect ("SystemError.asp")
        End If
        
        If Session("LocalLocked") = Cbool(True) Then 
            Session.Abandon
            Response.Redirect("SystemLocked.html")
        End If
        
        Session("LoggedOff") = False        
        Session("LogonErrMsg") = ""
        Application("UsersOnLine") = Cint(Application("UsersOnLine")) +1
        
        Application("ConnMcLogon") = Session("ConnMcLogon")
       
        For Count = 1 to Session("UserCount")
            If Session("UserId") = Cint(Count) Then Application("UsersOnLineName" & Cstr(Count)) = Session("UserName")
        Next
        
        '## Clear any locks set by this user
        ClearLocks Session("UserId")
                        
        Response.Redirect(strRedirectLoginSuccess)						
		    
    Else
	    UserRs.Close
        Set UserRs = Nothing
        			
        Session("LogonErrMsg") = "Invalid or missing Password"
	    
        Response.Redirect(strRedirectLoginFailed)
    End if
		
Else		
    '## nothing entered reload page
    Response.Redirect(strRedirectLoginFailed)
End If

'########################################### Get system settings 

Function SystemData()

    Dim SystemRs
    Dim LocalSystem
    Set SystemRs = Server.CreateObject("ADODB.Recordset")
    SystemRs.ActiveConnection = Session("ConnMcLogon") 
    SystemRs.Source = "SELECT UserCount, Locked From System"
    SystemRs.CursorType = Application("adOpenForwardOnly")
    SystemRs.CursorLocation = Application("adUseClient")
    SystemRs.LockType = Application("adLockReadOnly")    
    SystemRs.Open
    Set SystemRs.ActiveConnection = Nothing

    If Err <> 0 Then
        Err.Clear
        LocalSystem = False        
    Else 
        Session("LocalLocked") = SystemRs("Locked")   
        Session("UserCount") = SystemRs("UserCount")
        LocalSystem = True
    End If
    
    SystemRs.Close
    Set SystemRs = Nothing
    
    SystemData = LocalSystem

End Function


Private Sub ClearLocks(UserID)

    '## Only clears locks made by current user

    On Error Resume Next

    Dim LockSql
    Dim LockRs
    
    Set LockRs = Server.CreateObject("ADODB.Recordset")
    LockRs.ActiveConnection = Session("ConnMcLogon") 
    LockRs.Source = "Delete From RecordLocks"  
    LockRs.Source = LockRs.Source & " Where (LockedById = " & Clng(UserId) & ")"
    LockRs.CursorType = Application("adOpenStatic")
    LockRs.CursorLocation = Application("adUseClient")
    LockRs.LockType = Application("adLockOptimistic")
    LockRs.Open
    
    If LockRs.BOF = True Or LockRs.EOF = True Then
        '## No match so do nothing
        Err.Clear
    End If
    
    Err.Clear

    Set LockRs = Nothing

End Sub


%>

