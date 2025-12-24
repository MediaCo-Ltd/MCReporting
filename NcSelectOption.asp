<%@Language="VBScript" Codepage="1252"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\connIpAddressDB.asp" -->
<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<%

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp") 

response.expires = -1
response.AddHeader "Pragma", "no-cache"
response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"

'## blue line is #2D4291

Session("JobType") = ""
Session("JobTypeNo") = ""


Dim LocalMsg
LocalMsg = ""
If Session("OrderOk") <> "" Then LocalMsg = Session("OrderOk")
Session("OrderOk") = ""

Dim Optionvisible
Dim MyLogsVisible
Dim MyLogsRedirect

If Session("AdminUser") = Cbool(True) Or Session("EditNC") = Cbool(True) Then
    Optionvisible = ""
Else
    Optionvisible = " style='visibility: hidden' "
End If

If UserHasLogs = True Then
    MyLogsVisible = ""
    MyLogsRedirect = "NcDisplayUserLogs.asp"
Else
    MyLogsVisible = " style='visibility: hidden' "
    MyLogsRedirect = ""
End If



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

If Session("LockedId") <> "" Then
    '## Clear record lock, but only if made by this user
       
    Dim LockSql
    Dim LockRs

    Set LockRs = Server.CreateObject("ADODB.Recordset")
    LockRs.ActiveConnection = Session("ConnMcLogon")
    LockRs.Source = "Select * From RecordLocks"          
    LockRs.Source = LockRs.Source & " Where (RecordId = " & Session("LockedId") & ")"
    LockRs.Source = LockRs.Source & " And (LockedById = " & Session("UserId") & ")"
    LockRs.Source = LockRs.Source & " And (SystemName = 'NC')"
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
    <title>NC Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/NCReportsJSFunc.js"></script>
    <link rel="shortcut icon" href="Images/warning.png" type="image/x-icon" /> 
    <script type="text/javascript" >
        function Relocate(strUrl) 
        {
            document.getElementById("GoBack").style.color = "#ffffff";
            document.body.style.cursor = 'wait';
            document.getElementById('loading').innerHTML = "Loading Data, Please wait...";
            window.location.replace(strUrl);
        }   
        function RelocateNoMsg(strUrl)
        {
            window.location.replace(strUrl);
        }    
    </script>       
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:OptionsLoad();">
    <table  style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
        <tr>
		    <td align="left" valign="bottom" height="100" colspan="3">
                <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
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
                <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" />
                <font style="color: #0069AA; font-weight: bold; font-size: 16px;">NC&nbsp;Reporting&nbsp;(<%=Session("UserName")%>)</font>
            </td>
            <td align="right" width="33%"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>; ">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
        </tr>
    </table>
    <table  style="width: 100%; position: absolute; top: 40%; padding-right: 10px; padding-left: 10px;" >            
        <tr>   
            <td valign="middle" width="33%" >      
                <p><font size="3" color="<%=NewBlue%>">
                    <br  /><br  />                    
                    &nbsp;<input type="radio" onclick="javascript:RelocateNoMsg('NcJobNo.asp');" />
                    <span onclick="javascript:RelocateNoMsg('NcJobNo.asp');" >Add&nbsp;New.&nbsp;by&nbsp;Job&nbsp;No</span>                    
                    <br  /><br  /> 
                    &nbsp;<input type="radio" onclick="javascript:RelocateNoMsg('NcAddNewNJ.asp');" />
                    <span onclick="javascript:RelocateNoMsg('NcAddNewNJ.asp');" >Add&nbsp;New.&nbsp;Not&nbsp;Linked&nbsp;to&nbsp;Job</span>
                    <br  /><br  />
                    &nbsp;<input <%=MyLogsVisible%> type="radio" id="WiewMyLogs" onclick="javascript:RelocateOption('<%=MyLogsRedirect%>');"/> 
                    <span <%=MyLogsVisible%> onclick="javascript:RelocateOption('<%=MyLogsRedirect%>');" >View&nbsp;My&nbsp;Logs</span>                  
                    <br  /><br  />                   
                    &nbsp;<input <%=Optionvisible %> type="radio" onclick="javascript:RelocateNoMsg('NcDisplay.asp');" />
                    <span <%=Optionvisible %> onclick="javascript:RelocateNoMsg('NcDisplay.asp');" >View&nbsp;Records</span> 
                    <br  /><br  />
                                
                </font>
                </p>
               
            </td>
            <td valign="middle" align="center" width="33%" id="loading" style="font-size: medium; color: <%=NewBlue%>; font-weight: bold;">
                <noscript style="color:Red" >Your Browser has Javascript disabled<br />Please enable, to allow full functionality</noscript>                       
            </td>
            <td width="33%">
            
                <%If InputType = "text" Then Response.Write "ErrBox"%>
                <input type="<%=InputType %>" name="ErrBox" id ="ErrBox" value="<%=LocalMsg%>"/>&nbsp;
                <%If InputType = "text" Then Response.Write "Locked"%>
                <input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" />&nbsp;            
            </td>
        </tr>     
    </table>
                  
    <!--<table  style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
        <tr>  
            <td height="50" >
                <hr style="border-style: none; height: 1px; background-color: <%'=NewCyan%>; display: block;" />
                <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
                Trafford Park, Manchester M17 1BS</font></p>
            </td>
        </tr>
    </table>-->
</body>  
</html>

<%
Private Function UserHasLogs
    
    Dim UserHasLogsRs
    Dim LocalResult
    LocalResult = False

    Set UserHasLogsRs = Server.CreateObject("ADODB.Recordset")
    UserHasLogsRs.ActiveConnection = Session("ConnNCReports") 
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