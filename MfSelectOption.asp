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

Session("AddFolder") = ""

Dim MachineSql
Dim MachineRs

MachineSql = "SELECT DISTINCT Logs.MachineId, Machine.MachineName, Machine.MachineTypeId, Machine.FaultGroups" 
MachineSql = MachineSql & " FROM Logs INNER JOIN  Machine ON Logs.MachineId = Machine.Id"
MachineSql = MachineSql & " WHERE (NOT (Logs.LogDateSerial IS NULL))"
MachineSql = MachineSql & " Order By Machine.MachineTypeId, Machine.MachineName" 

Set MachineRs = Server.CreateObject("ADODB.Recordset")
MachineRs.ActiveConnection = Session("ConnMachinefaults") 
MachineRs.Source = MachineSql
MachineRs.CursorType = Application("adOpenForwardOnly")
MachineRs.CursorLocation = Application("adUseClient") 
MachineRs.LockType = Application("adLockReadOnly")
MachineRs.Open
Set MachineRs.ActiveConnection = Nothing

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

Dim OptionByNumber
Dim Optionvisible
Dim MyLogsRedirect
Dim MyLogsVisible


If UserHasLogs = True Then
    MyLogsVisible = ""
    MyLogsRedirect = "MfDisplayUserLogs.asp"
Else
    MyLogsVisible = " style='visibility: hidden' "
    MyLogsRedirect = ""
End If 

If Session("AdminUser") = Cbool(True) Or Session("EditMF") = Cbool(True) Then
    Optionvisible = ""    
    OptionByNumber = "&nbsp;&nbsp;View/Edit&nbsp;Record&nbsp;"
    OptionByNumber = OptionByNumber & "&nbsp;<input title='Enter number & press enter key' type='text' value='' id='JumpTxt' size='12' onkeyup='ShowOrder(event);'/> "
Else
    OptionByNumber = ""
    Optionvisible = " style='visibility: hidden' "
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
    LockRs.Source = LockRs.Source & " And (SystemName = 'MF')"
    LockRs.CursorType = Application("adOpenStatic")
    LockRs.CursorLocation = Application("adUseClient")
    LockRs.LockType = Application("adLockOptimistic")
    LockRs.Open

    If LockRs.BOF = True Or LockRs.EOF = True Then
        '## No match so do nothing
    Else
        If LockRs("LockedbyId") = Session("UserId") Then
            LockRs.Delete
            LockRs.MoveNext
        End If   
    End If

    Set LockRs.ActiveConnection = Nothing
    LockRs.Close
    Set LockRs = Nothing

    Session("LockedId") = ""
    Session("LockedSid") = ""
    Session("LockedBy") = ""
    Session("LockedByName") = ""
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>Machine Faults</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/MachineFaultJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:OptionsLoad();">
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
                <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
                <font style="color: #0069AA; font-weight: bold; font-size: 16px;">MF&nbsp;Reporting&nbsp;(<%=Session("UserName")%>)</font>
            </td>
            <td align="right" width="33%"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
        </tr>
        
    </table>
    
    <table id="Options" style="width: 100%; position: absolute; top: 32%; padding-right: 20px; padding-left: 20px;" >            
        <tr>   
            <td valign="middle" width="33%" >      
                <p>
                    <font size="3" color="<%=NewBlue%>">
                        <br  /><br  />                    
                        &nbsp;<input type="radio" id="Add" onclick="javascript:RelocateOption('MfAddNew.asp');" />
                        <span onclick="javascript:RelocateOption('AddNew.asp);" >Add&nbsp;New</span>
                        <br  /><br  />
                        &nbsp;<input <%=MyLogsVisible%> type="radio" id="WiewMyLogs" onclick="javascript:RelocateOption('<%=MyLogsRedirect%>');"/> 
                        <span  <%=MyLogsVisible%> onclick="javascript:RelocateOption('<%=MyLogsRedirect%>');" >View&nbsp;My&nbsp;Logs</span>                  
                        <br  /><br  />
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Vp" onclick="javascript:RelocateOption('MfDisplay.asp?st=0');" />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('MfDisplay.asp?st=0');">View&nbsp;Pending</span>                    
                        <br  /><br  />     
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Vc" onclick="javascript:RelocateOption('MfDisplay.asp?st=1');" />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('MfDisplay.asp?st=1');">View&nbsp;Complete</span> 
                        <br  /><br  />
                        &nbsp;<input  <%=Optionvisible %> type="radio" id="Va" onclick="javascript:RelocateOption('MfDisplay.asp?st=2');" />
                        <span  <%=Optionvisible %> onclick="javascript:RelocateOption('MfDisplay.asp?st=2');">View&nbsp;All</span>                   
                        <br  /><br  />                
                        <label  <%=Optionvisible %> >&nbsp;&nbsp;View&nbsp;By&nbsp;Machine&nbsp;</label>
                        <select name="cboMachine" id="cboMachine" onchange="javascript:RelocateMachine();"  <%=Optionvisible %> >
                            <option value="">Select Machine</option>
                            <% 
                            While Not MachineRs.EOF
                                Response.Write Response.Write ("<option value='" & MachineRs("MachineId") & "'>" & Trim(MachineRs("MachineName") & "</option>)"))  & VbCrLf
                            
                                MachineRs.MoveNext
                            Wend
                            MachineRs.Close
                            Set MachineRs = Nothing
                            %>
                        </select>
                        <br  /><br  /> 
                        <%=OptionByNumber%>                   
                    </font>                    
                </p>               
            </td>
            <td valign="middle" align="center" width="33%" id="loading" style="font-size: medium; color: <%=NewBlue%>; font-weight: bold;">
                <noscript style="color:Red" >Your Browser has Javascript disabled<br />Please enable, to allow full functionality</noscript>                       
            </td>
            <td width="33%">&nbsp;<input type = "<%=InputType%>" id="Machine" value = ""/></td>
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

Private Function UserHasLogs
    
    Dim UserHasLogsRs
    Dim LocalResult
    LocalResult = False

    Set UserHasLogsRs = Server.CreateObject("ADODB.Recordset")
    UserHasLogsRs.ActiveConnection = Session("ConnMachinefaults") 
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