<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\connMCReportsDB.asp" -->

<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.Cookies("IpChecked") = "False"

Dim ErrMsg
If Session("LogonErrMsg") <> "" Then
    ErrMsg = Session("LogonErrMsg")
    Session("LogonErrMsg") = ""
Else
    ErrMsg = ""
End If 

Dim PasswordReq
Dim LocalUserName
Dim LocalUserId
Dim Password
Dim Email
LocalUserId = Request.Form("txtUserID")

Dim UsersRs

Set UsersRs = Server.CreateObject("ADODB.Recordset")
UsersRs.ActiveConnection = Session("ConnMcLogon")  
UsersRs.Source = "SELECT Id, UserName, AdminUser, HsEdit, NcEdit, MfEdit, PwRequired, Password, PwEmail FROM Users Where(Id = " & LocalUserId & ")"

UsersRs.CursorType = Application("adOpenForwardOnly")
UsersRs.CursorLocation = Application("adUseClient")
UsersRs.LockType = Application("adLockReadOnly")
UsersRs.Open
Set UsersRs.ActiveConnection = Nothing

LocalUserName = Trim(UsersRs("UserName"))
Dim NoReset



If UsersRs("AdminUser") = Cbool(True) Then 
    NoReset = Cbool(True)
ElseIf UsersRs("HsEdit") = Cbool(True) Then 
    NoReset = Cbool(True)
ElseIf UsersRs("NcEdit") = Cbool(True) Then 
    NoReset = Cbool(True)
ElseIf UsersRs("MfEdit") = Cbool(True) Then
    NoReset = Cbool(True)
Else
    '## NoReset = Cbool(False)
    NoReset = Cbool(True)
End If 

PasswordReq = UsersRs("PwRequired")
Password = UsersRs("Password")
Email = UsersRs("PwEmail")


UsersRs.Close
Set UsersRs = Nothing

If PasswordReq = Cbool(False) Then Response.Redirect ("AutoLogin.asp?id=" & LocalUserId)

Dim PassBoxDisabled
Dim Visbillity



If PasswordReq = Cbool(True) And Password = "" Then 
    PassBoxDisabled = " disabled='disabled' "
    Visbillity = " style='visibility: hidden' "
Else
    PassBoxDisabled = ""
    Visbillity = ""
End If    

Dim InputType
Dim InputSize

If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    If LocalUserId = "1" Or LocalUserId = "4" Then
        InputType = "text"
        InputSize = " size='8' "
    Else
        InputType = "hidden"
        InputSize = ""
    End If
Else
    InputType = "hidden"
    InputSize = ""
End If

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

<body style="padding: 0px; margin: 0px" onload="PwLoadChk('LoginPw');">

<form action="LoginCheck.asp" method="post" name="frmLoginPw" id="frmLoginPw" onsubmit="return ValidatePW();">

<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
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
        <td width="33%">&nbsp;&nbsp;<a href ="javascript:LogOff();" style="color: <%=NewBlue%>;">Return&nbsp;to&nbsp;select&nbsp;user&nbsp;page</a></td>
        <td align="center">
            <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">MediaCo&nbsp;Reporting</font></td>
        <td width="33%">&nbsp;</td>
    </tr>	
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 30%;" id="data">
    
    
    <tr>
        <td height="20px" id="loading" style="font-size: medium; color: blue; font-weight: bold; text-align:center;" colspan="2">&nbsp;</td>
    </tr>
    
    <tr>
        <td colspan="2" class="Userlabel" >
            <font size="3" style="color: #0069AA">&nbsp;&nbsp;<%=LocalUserName%>.&nbsp;
            <%  If PassBoxDisabled = " disabled='disabled' " Then
                    Response.Write "&nbsp;You need to create a password click <a href ='javascript:SetPassword(" & LocalUserId & ");' style='color: #ff0000;'>here</a> to set it"
                End If
            %>
            </font>
        </td>
    </tr>
    <tr>
        <td height="10px" >
            
         </td>
        <td height="10px" id="Show"></td>        
    </tr>
    <tr <%=Visbillity%>>
        <td valign="middle" width="50px" class="inputlabel" >&nbsp;<font size="3">Enter&nbsp;Your&nbsp;Password</font></td>
        <td valign="middle" >
            <input name="txtPassword" type="text" id="txtPassword" class="inputboxes" value="" style="color: #000000"
              onkeydown="return DisableEnterKey(event,'LoginPw');" onkeyup="javascript:EnableSubmit('LoginPw');" <%=PassBoxDisabled%> />
                &nbsp;<font size="3" style="color: #0069AA; ">&nbsp;Then&nbsp;click&nbsp;Submit</font>
         </td>
    </tr>
    <tr>
        <td height="10px" colspan="2">&nbsp;</td>
    </tr>
    <tr>
        <td colspan="2" class="Pwlabel" <%=Visbillity%> >
            <font style="color: #FF0000;"><br />Please Note.</font>
            <!--<br />You can't paste your password in, it must be typed and it will never be saved.-->
            <br />If you type your password in wrong, the backspace key will have no effect. Click the reset button.<br />
            <br />Forgotten your password ?&nbsp;
            <a href ="javascript:PwReminder();" style="color: <%=NewBlue%>;">Send reminder
            <!--SendEmailReminder.asp?Id="-->
            <!--<button onclick="javascript:PwReminder();" >Send reminder</button>-->
            
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
		<td <%=Visbillity %>>
			<input name="btnSubmit" id="btnSubmit" type="submit" value="Submit" disabled="disabled"  />
			&nbsp;&nbsp;<input name="btnReset" id="btnReset" type="button" value="Reset Page" onclick="javascript:ResetPwUser();"/>
			&nbsp;&nbsp;<input name="frmName" id="frmName" type="<%=InputType%>" value="frmLoginPw" <%=InputSize%> />
			&nbsp;&nbsp;<input name="ErrBox" id="ErrBox" type="<%=InputType%>" value="<%=ErrMsg%>" <%=InputSize%>/>
			&nbsp;&nbsp;<input name="txtUserID" id="txtUserID" type="<%=InputType%>" value="<%=LocalUserId%>" <%=InputSize%>/>
			&nbsp;&nbsp;<input name="txtUserPW" id="txtUserPW" type="<%=InputType%>" value="" <%=InputSize%>/>
			&nbsp;&nbsp;<input name="txtUserChar" id="txtUserChar" type="<%=InputType%>" value="" <%=InputSize%> />	
			&nbsp;&nbsp;<input name="LocalUserName" id="LocalUserName" type="<%=InputType%>" value="<%=LocalUserName%>" <%=InputSize%> />
			&nbsp;&nbsp;<input name="PasswordReq" id="PasswordReq" type="<%=InputType%>" value="<%=PasswordReq%>" <%=InputSize%> />
			&nbsp;&nbsp;<input name="Email" id="Email" type="<%=InputType%>" value="<%=Email%>" <%=InputSize%> />
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

