<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\connMCReportsDB.asp" -->
<!--#include file="..\##GlobalFiles\connQualityControlDB.asp" -->

<%

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.Cookies("IpChecked") = "False" 

Session("ConnQC") = strConnQualityControlDB

Dim AutoLogin
AutoLogin = Request.QueryString("id")
If AutoLogin <> "" Then Response.Redirect ("Autologin.asp?id=" & AutoLogin)


Dim ErrMsg
If Session("LogonErrMsg") <> "" Then
    ErrMsg = Session("LogonErrMsg")
    Session("LogonErrMsg") = ""
Else
    ErrMsg = ""
End If

Dim UsersRs

Set UsersRs = Server.CreateObject("ADODB.Recordset")
UsersRs.ActiveConnection = Session("ConnMcLogon") 
UsersRs.Source = "SELECT Id, UserName, PwRequired FROM Users Where(Id > 1) AND (Active = 1) Order by UserName"

UsersRs.CursorType = Application("adOpenForwardOnly")
UsersRs.CursorLocation = Application("adUseClient")
UsersRs.LockType = Application("adLockReadOnly")
UsersRs.Open
Set UsersRs.ActiveConnection = Nothing

Dim InputType
Dim InputSize

If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
    InputSize = " size='8' "
Else
    InputType = "hidden"
    InputSize = ""
End If

Dim DemoMsg
Dim DemoMsg2
DemoMsg = ""
DemoMsg2 = ""

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/MCReportsJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" onload="LogonLoadChk();">

<form action="LoginPw.asp" method="post" name="frmSelectUser" id="frmSelectUser">

<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100">
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td>            
    </tr>
    <tr>
	    <td height="8" valign="top">
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	    </td>
    </tr>
    <tr>
    <td align="center">
        <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
        <font style="color: #0069AA; font-weight: bold; font-size: 16px;">MediaCo&nbsp;Reporting</font>
    </td>
    </tr>	
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 30%;" id="data">
    <tr>
        <td height="140px" valign="middle" colspan="2" >
            <p>&nbsp;<font size="3" style="color: #0069AA">Select&nbsp;Your&nbsp;User&nbsp;Name.<br /><br />
            <!--&nbsp;New&nbsp;users&nbsp;click
            <a href ="javascript:SetPassword();" style="color: #ff0000;">here</a>&nbsp;to&nbsp;&nbsp;to&nbsp;gain&nbsp;access-->
                &nbsp;Not&nbsp;in&nbsp;list&nbsp;ask&nbsp;Alan&nbsp;or&nbsp;Dave&nbsp;to&nbsp;add&nbsp;you
            </font>
            </p>
        </td>
    </tr>    
        
    <tr >
        <td valign="bottom" width="50px" class="inputlabel">&nbsp;<font size="3">User</font></td>
        <td valign="middle" >
            <select id="cboUsers" onchange="javascript:SelectName();javascript:SubmitLogin();">  <!--;javascript:SubmitLogin();-->
                <option value="" >Select User</option>
                <%
                While Not UsersRs.EOF
                    Response.Write ("<option value='" & Trim(UsersRs("Id")) & "'>" & Trim(UsersRs("UserName")) & "</option>)") & VbCrLf
                    UsersRs.MoveNext
                Wend
                
                UsersRs.Close
                Set UsersRs = Nothing
                
                %>
            </select>
        </td>
        
    </tr>
    <!--
    <tr>
        <td height="20px"  style="font-size: medium; color: blue; font-weight: bold; text-align:center;" colspan="2">
            &nbsp;
        </td>
    </tr>
    
   <tr >
        <td valign="middle" width="50px" class="inputlabel" >&nbsp;<font size="3">Enter&nbsp;Password</font></td>
        <td valign="middle" >
            <input name="txtPassword" type="password" id="txtPassword" class="inputboxes" value="" 
                  style="color: #000000" />
            &nbsp;<font size="3" style="color: #0069AA; ">&nbsp;Then&nbsp;click&nbsp;Submit</font>
         </td>
    </tr>-->
    
    <tr>
        <td height="20px" id="loading" style="font-size: medium; color: blue; font-weight: bold; text-align:center;" colspan="2">
            &nbsp;
        </td>
    </tr>
    
    <tr >        
        <td colspan="2" >
            <noscript style="color:Red; text-align:center; height: 25px; vertical-align:bottom;"><h3>Your Browser has Javascript disabled.&nbsp;Please enable, to allow full functionality</h3></noscript> 
        </td>
    </tr>
   
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 75%;" id="buttons">
    <tr>
		<td class="inputlabel" width="50px">&nbsp;</td>
		<td >
			<input name="btnSubmit" id="btnSubmit" type="submit" value="Submit" disabled="disabled" />
			<%
            If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                Response.Write ("&nbsp;&nbsp;")
                Response.Write ("<input type='button' id='Nologin' name='Nologin' value='Bypass' onclick='javascript:Bypass();'/>")
            End If
            Session.Abandon          
            %> 
			&nbsp;&nbsp;<input name="ErrBox" id="ErrBox" type="hidden" value="<%=ErrMsg%>" />
			&nbsp;&nbsp;<input name="txtUserID" id="txtUserID" type="<%=InputType%>" />
			&nbsp;&nbsp;<input name="frmName" id="frmName" type="<%=InputType%>" value="frmSelectUser"  />
		</td>
    </tr>  
</table>
      
<table style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
    <tr>  
        <td height="50" >
            <hr style="width: 98%; border-style: none;  height: 1px; background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, Trafford Park, Manchester M17 1BS</font></p>
        </td>
    </tr>
</table>

</form> 
</body>
</html>

