<%@Language="VBScript" Codepage="1252"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim LocalMsg
LocalMsg = ""
If Session("OrderOk") <> "" Then LocalMsg = Session("OrderOk")
Session("OrderOk") = ""
Session("OrderUpdated") = ""
Session("Session_Id") = Session.SessionID

Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    If Session("UserId") = 4 Then
        InputType = "text"
    Else
        InputType = "hidden"
    End If
Else
    InputType = "hidden"
End If

Dim AccidentDisabled
Dim IncidentDisabled
Dim UnsafeDisabled
Dim CustomDisabled
Dim MyLogsRedirect
Dim Optionvisible
Dim MyLogsVisible

AccidentDisabled = ""
IncidentDisabled = ""
UnsafeDisabled = ""
CustomDisabled = ""

If UserHasLogs = True Then
    MyLogsVisible = ""
    MyLogsRedirect = "HsDisplayUserLogs.asp"
Else
    MyLogsVisible = " style='visibility: hidden' "
    MyLogsRedirect = ""
End If    

If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Or Session("ViewHS") = Cbool(True) Then 
    If TypeChk(1) = False Then AccidentDisabled = " disabled='disabled' "
    If TypeChk(2) = False Then IncidentDisabled = " disabled='disabled' "
    If TypeChk(3) = False Then UnsafeDisabled = " disabled='disabled' "
    Optionvisible = ""
Else
    AccidentDisabled = " disabled='disabled' "
    IncidentDisabled = " disabled='disabled' "
    UnsafeDisabled = " disabled='disabled' "
    CustomDisabled = " disabled='disabled' "
    Optionvisible = " style='visibility: hidden' "
End If

If Session("EditHS") <> Cbool(True) Then CustomDisabled = " disabled='disabled' "
If Session("ViewHS") = Cbool(True) Then CustomDisabled = ""

Dim OptionByNumber

If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Then    '##Or Session("ViewHS") = Cbool(True) 
    OptionByNumber = "&nbsp;&nbsp;View/Edit&nbsp;Record&nbsp;"
    OptionByNumber = OptionByNumber & "&nbsp;<input title='Enter number & press enter key' type='text' value='' id='JumpTxt' size='12' onkeyup='ShowOrder(event);'/> "
Else
    OptionByNumber = ""
End If

If Session("LockedId") <> "" Then
    '## Clear record lock, but only if made by this user
    
    Dim LockSql
    Dim LockRs

    Set LockRs = Server.CreateObject("ADODB.Recordset")
    LockRs.ActiveConnection = Session("ConnMcLogon")
    LockRs.Source = "Select * From RecordLocks"          
    LockRs.Source = LockRs.Source & " Where (RecordId = " & Session("LockedId") & ")"
    LockRs.Source = LockRs.Source & " And (LockedById = " & Session("UserId") & ")"
    LockRs.Source = LockRs.Source & " And (SystemName = 'HS')"
    LockRs.CursorType = Application("adOpenStatic")
    LockRs.CursorLocation = Application("adUseClient")
    LockRs.LockType = Application("adLockOptimistic")
    LockRs.Open

    If LockRs.BOF = True Or LockRs.EOF = True Then
        '## No match so do nothing
    Else
        If LockRs("LockedbyId") = Session("UserId") Then  'And LockRs("SystemName") = "HS"
            LockRs.Delete
            LockRs.MoveNext
        End If   
    End If

    Set LockRs.ActiveConnection = Nothing
    LockRs.Close
    Set LockRs = Nothing

    Session("LockedId") = ""
    Session("LockedBy") = ""
    Session("LockedByName") = ""
    
End If

Dim UnresolvedRs
Dim AccidentUnresolved
Dim IncidentUnresolved
Dim UnsafeUnresolved

AccidentUnresolved = 0
IncidentUnresolved = 0
UnsafeUnresolved = 0

Set UnresolvedRs = Server.CreateObject("ADODB.Recordset")
UnresolvedRs.ActiveConnection = Session("ConnHSReports")
UnresolvedRs.Source = "Select GroupId From Logs Where (Resolved = 0)"          
UnresolvedRs.CursorType = Application("adOpenForwardOnly") 
UnresolvedRs.CursorLocation = Application("adUseClient")
UnresolvedRs.LockType = Application("adLockReadOnly")
UnresolvedRs.Open

If UnresolvedRs.BOF = True Or UnresolvedRs.EOF = True Then
        '## No match so do nothing
Else
    While Not UnresolvedRs.EOF
        If UnresolvedRs("GroupId") = 1 Then AccidentUnresolved = AccidentUnresolved +1
        If UnresolvedRs("GroupId") = 2 Then IncidentUnresolved = IncidentUnresolved +1
        If UnresolvedRs("GroupId") = 3 Then UnsafeUnresolved = UnsafeUnresolved +1    
        UnresolvedRs.MoveNext
    Wend
End If

UnresolvedRs.Close
Set UnresolvedRs = Nothing

Dim TableTop
If Session("AdminUser") = Cbool(True) Or Session("EditHS") = Cbool(True) Or Session("ViewHS") = Cbool(True) Then 
    TableTop = "20%"
Else
    TableTop = "30%"
End If

Dim RiddorDisabled
Dim RiddorRs
Dim RiddorSql

RiddorSql = "Select Id, Riddor FROM Logs Where (Riddor = 1)"

Set RiddorRs = Server.CreateObject("ADODB.Recordset")
RiddorRs.ActiveConnection = Session("ConnHSReports") 
RiddorRs.Source = RiddorSql
RiddorRs.CursorType = Application("adOpenForwardOnly")
RiddorRs.CursorLocation = Application("adUseClient") 
RiddorRs.LockType = Application("adLockReadOnly")
RiddorRs.Open

If RiddorRs.BOF = True Or RiddorRs.EOF = True Then
    '## No match so do nothing
    RiddorDisabled = " disabled='disabled' "
Else
    RiddorDisabled = ""
End If

RiddorRs.Close
Set RiddorRs = Nothing


Session("AddFolder") = ""

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/HSReportsJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:HSOptionsLoad();">
    <table id="Logo" style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
        <tr>
		    <td align="left" valign="bottom" height="100" colspan="3">
                <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />                
            </td>            
        </tr>
	    <tr>
		    <td height="8" valign="top" colspan="3">
		        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		    </td>
	    </tr>
	    <tr>
	        <td width="33%">&nbsp;&nbsp;
	        <a id="logoffL" href ="javascript:window.location.replace('SelectSite.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return&nbsp;to&nbsp;select&nbsp;system</a></td>
            <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">HS&nbsp;Reporting&nbsp;(<%=Session("UserName")%>)</font></td>
            <td align="right" width="33%"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
        </tr>
        
    </table>
    
    <table id="Options" style="width: 100%; position: absolute; top: <%=TableTop%>; padding-right: 20px; padding-left: 20px;" >            
        <tr>   
            <td valign="middle" width="33%" >      
                <p>
                    <font size="3" color="<%=NewBlue%>">
                        <br  /><br  />                    
                        &nbsp;<input type="radio" id="Add1" onclick="javascript:RelocateOption('HsAddNew.asp?st=1');" /> 
                        <span onclick="javascript:RelocateOption('HsAddNew.asp?st=1');" >Add&nbsp;Accident&nbsp;Record</span>
                        <br  /><br  />
                        &nbsp;<input type="radio" id="Add2" onclick="javascript:RelocateOption('HsAddNew.asp?st=2');" /> 
                        <span onclick="javascript:RelocateOption('HsAddNew.asp?st=2');" >Add&nbsp;Near&nbsp;Miss&nbsp;Record</span>                    
                        <br  /><br  />     
                        &nbsp;<input type="radio" id="Add3" onclick="javascript:RelocateOption('HsAddNew.asp?st=3');" /> 
                        <span onclick="javascript:RelocateOption('HsAddNew.asp?st=3');" >Add&nbsp;Unsafe&nbsp;Condition&nbsp;or&nbsp;Damage&nbsp;Record</span> 
                        
                     
                        
                        <br  /><br  />
                        &nbsp;<input <%=MyLogsVisible%> type="radio" id="WiewMyLogs" onclick="javascript:RelocateOption('<%=MyLogsRedirect%>');"/> 
                        <span <%=MyLogsVisible%> onclick="javascript:RelocateOption('<%=MyLogsRedirect%>');" >View&nbsp;My&nbsp;Logs</span>                  
                        <br  /><br  />
                        
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Radio1" onclick="javascript:RelocateOption('HsDisplayRiddor.asp');"  <%=RiddorDisabled%> />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('HsDisplayRiddor.asp');"  <%=RiddorDisabled%> >View&nbsp;All&nbsp;Riddor</span>
                        <br  /><br  />
                        
                        &nbsp;<input <%=Optionvisible %> type="radio" id="Wiew1" onclick="javascript:RelocateOption('HsDisplay.asp?st=1&rf=1');" <%=AccidentDisabled%> /> 
                        <span <%=Optionvisible %> onclick="javascript:RelocateOption('HsDisplay.asp?st=1&rf=1');" <%=AccidentDisabled%> >View&nbsp;All&nbsp;Accidents</span>  
                        <br  /><br  />
                        &nbsp;<input <%=Optionvisible %> type="radio" id="view11" onclick="javascript:RelocateOption('HsDisplay.asp?st=1&rf=0');" <%If AccidentUnresolved = 0 Then Response.Write " disabled='disabled' "%> /> 
                        <span <%=Optionvisible %> onclick="<%If AccidentUnresolved > 0 Then Response.Write "javascript:RelocateOption('HsDisplay.asp?st=1&rf=0');" %>" >View&nbsp;Unresolved&nbsp;Accidents</span>                  
                        <br  /><br  />
                             
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Wiew2" onclick="javascript:RelocateOption('HsDisplay.asp?st=2&rf=1');" <%=IncidentDisabled%> />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('HsDisplay.asp?st=2&rf=1');" <%=IncidentDisabled%> >View&nbsp;All&nbsp;Near&nbsp;Miss</span> 
                        <br  /><br  />
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Wiew22" onclick="javascript:RelocateOption('HsDisplay.asp?st=2&rf=0');" <%If IncidentUnresolved = 0 Then Response.Write " disabled='disabled' "%> />
                        <span  <%=Optionvisible %> onclick="<%If IncidentUnresolved > 0 Then Response.Write "javascript:RelocateOption('HsDisplay.asp?st=2&rf=0');" %>" >View&nbsp;Unresolved&nbsp;Near&nbsp;Miss</span> 
                        <br  /><br  />
                        
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Wiew3" onclick="javascript:RelocateOption('HsDisplay.asp?st=3&rf=1');" <%=UnsafeDisabled%> />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('HsDisplay.asp?st=3&rf=1');" <%=UnsafeDisabled%> >View&nbsp;All&nbsp;Unsafe&nbsp;Condition&nbsp;or&nbsp;Damage&nbsp;Records</span>                   
                        <br  /><br  />
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Wiew33" onclick="javascript:RelocateOption('HsDisplay.asp?st=3&rf=0');" <%If UnsafeUnresolved = 0 Then Response.Write " disabled='disabled' "%> />
                        <span  <%=Optionvisible %> onclick="<%If UnsafeUnresolved > 0 Then Response.Write "javascript:RelocateOption('HsDisplay.asp?st=3&rf=0');" %>" >View&nbsp;Unresolved&nbsp;Unsafe&nbsp;Condition&nbsp;or&nbsp;Damage&nbsp;Records</span>                   
                        <br  /><br  /><br  />
                                        
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="WiewCustom" onclick="javascript:RelocateOption('HsCustomSelect.asp');"  <%=CustomDisabled%> />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('HsCustomSelect.asp');"  <%=CustomDisabled%> >Create&nbsp;Custom&nbsp;Report</span>
                        <br  /><br  />
                        
                        
                         
                        <%=OptionByNumber%>                       
                    </font>                
                </p>               
            </td>
            <td valign="middle" align="center" width="33%" id="loading" style="font-size: medium; color: <%=NewBlue%>; font-weight: bold;">
                <noscript style="color:Red" >Your Browser has Javascript disabled<br />Please enable, to allow full functionality</noscript>                       
            </td>
            
        </tr>
      
        <tr>           
            <td colspan="3">
                <%If InputType = "text" Then Response.Write "ErrBox"%>
                <input type="<%=InputType %>" name="ErrBox" id ="ErrBox" value="<%=LocalMsg%>"/>&nbsp;
                <%If InputType = "text" Then Response.Write "hUserId"%>
                <input type="<%=InputType %>" name="hUserId" id ="hUserId" value="<%=Session("UserId")%>" />&nbsp;
                <%If InputType = "text" Then Response.Write "hUserName"%>
                <input type="<%=InputType %>" name="hUserName" id ="hUserName" value="<%=Session("UserName")%>" />&nbsp;
                <%If InputType = "text" Then Response.Write "AdminUser"%>
                <input type="<%=InputType %>" name="hAdminUser" id ="hAdminUser" value="<%=Session("AdminUser")%>" />&nbsp;
                <%If InputType = "text" Then Response.Write "Locked"%>
                <input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" />&nbsp;
                
                
            </td>
        </tr>     
    </table>
                  
    <table id="Footer" style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
        <tr>  
            <td height="50" >
                <hr style="border-style: none; height: 1px; background-color: <%=NewCyan%>; display: block;" />
                <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
                Trafford Park, Manchester M17 1BS</font></p>
            </td>
        </tr>
    </table>
</body>  
</html>


<%
Private Function TypeChk(TypeToCheck)
    
    Dim TypeChkRs
    Dim LocalResult
    
    LocalResult = False
    
    Set TypeChkRs = Server.CreateObject("ADODB.Recordset")
    TypeChkRs.ActiveConnection = Session("ConnHSReports") 
    TypeChkRs.Source = "SELECT Id, GroupId From Logs Where (GroupId = " & TypeToCheck & ")"
    
    TypeChkRs.CursorType = Application("adOpenForwardOnly")
    TypeChkRs.CursorLocation = Application("adUseClient")
    TypeChkRs.LockType = Application("adLockReadOnly")
    TypeChkRs.Open
	Set TypeChkRs.ActiveConnection = Nothing
	
	If TypeChkRs.BOF = True Or TypeChkRs.EOF = True Then
	    LocalResult = False
	Else
	    LocalResult = True
	End If
	
	TypeChkRs.Close
	Set TypeChkRs = Nothing
	
	TypeChk = LocalResult


End Function



Private Function UserHasLogs
    
    Dim UserHasLogsRs
    Dim LocalResult
    LocalResult = False

    Set UserHasLogsRs = Server.CreateObject("ADODB.Recordset")
    UserHasLogsRs.ActiveConnection = Session("ConnHSReports") 
    UserHasLogsRs.Source = "SELECT Id From Logs Where (UserId = " & Clng(Session("UserId")) & ")"
    
    UserHasLogsRs.CursorType = Application("adOpenForwardOnly")
    UserHasLogsRs.CursorLocation = Application("adUseClient")
    UserHasLogsRs.LockType = Application("adLockReadOnly")
    UserHasLogsRs.Open
	Set UserHasLogsRs.ActiveConnection = Nothing
	
	If UserHasLogsRs.BOF = True Or UserHasLogsRs.EOF = True Then
	    LocalResult = False
	Else
	    LocalResult = True
	End If
	
	UserHasLogsRs.Close
	Set UserHasLogsRs = Nothing
	
	UserHasLogs = LocalResult

End Function
%>