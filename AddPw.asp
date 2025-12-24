<%@Language="VBScript"  Codepage="1252" EnableSessionState=True%>
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

Dim UsersRs
Dim UserId

UserId = Request.QueryString("id")

Set UsersRs = Server.CreateObject("ADODB.Recordset")
UsersRs.ActiveConnection = Session("ConnMcLogon") 
UsersRs.Source = "SELECT Id, UserName, Password FROM Users Where(id = " & UserId & ")"

UsersRs.CursorType = Application("adOpenForwardOnly")
UsersRs.CursorLocation = Application("adUseClient")
UsersRs.LockType = Application("adLockReadOnly")
UsersRs.Open
Set UsersRs.ActiveConnection = Nothing
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/MCReportsJSFunc.js"></script>
</head>
                                           
<body style="padding: 0px; margin: 0px" onunload="window.opener.location.replace('login.asp');" onload="PwLoadChk();"> 
<form action="UpdatePassword.asp" method="post" name="frmAddPW" id="frmAddPW" onsubmit="return ValidatePW();" > 

    <table id="Logo" style="width: 100%;  padding-right: 10px; padding-left: 10px;" >
        <tr>
	        <td align="left" valign="bottom" height="80">
                <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" /> 
            </td> 
        </tr>
        
        <tr>
		    <td height="8" valign="top">
		        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		    </td>
	    </tr>
	    <tr>
	        <td align="center" valign="bottom" >&nbsp;&nbsp;
                <font style="font-weight: bold; color:<%=NewBlue%>;" size="2">
                Select&nbsp;your&nbsp;name&nbsp;&amp;&nbsp;and&nbsp;create&nbsp;your&nbsp;own&nbsp;password</font>
            </td>	        
        </tr>  
    </table>   
    
    <table id="Data" style="width: 98%;  padding-right: 10px; padding-left: 20px;" align="left" >         
               
        <tr>
            <td height="10px" colspan="2">&nbsp;</td>
        </tr>
                
        <tr>
            <td valign="bottom" width="150px" class="inputlabel" >&nbsp;<font size="3">Select&nbsp;your&nbsp;name</font></td>
            <td valign="middle" > 
                <select id="cboUsers"  onchange="javascript:SelectName('AddNew');javascript:TabToPw();">
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
        
        <tr>
            <td height="10px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td valign="middle" width="150px" class="inputlabel" >&nbsp;<font size="3">Enter&nbsp;password</font></td>
            <td valign="middle" >
                <input name="txtPassword" type="text" id="txtPassword" class="inputboxes"  value="" onclick="javascript:PwClick('AddNew');"
                onkeypress="return DisableEnterKey(event,'');" onfocus="javascript:PwClick('AddNew');" onkeyup="javascript:EnableSubmitAddPw();" 
                size="40" maxlength="38"
                />
            </td>
        </tr> 
       
        <tr>
            <td height="10px" colspan="2">&nbsp;</td>
        </tr>
        
        <tr>
            <td colspan="2" align="center" >
                &nbsp;Password&nbsp;is&nbsp;not&nbsp;related&nbsp;to&nbsp;any&nbsp;windows&nbsp;logon&nbsp;&amp;&nbsp;will&nbsp;not&nbsp;change&nbsp;once&nbsp;set
            </td>
        </tr>
              
        <tr>
            <td colspan="2" align="center" style="color: #FF0000">
                &nbsp;Password&nbsp;can&nbsp;only&nbsp;contain&nbsp;alphanumeric&nbsp;characters&nbsp;plus&nbsp;these&nbsp;£$#@*=+-&nbsp;symbols
            </td>
        </tr>
        
        <tr>
        <td class="NonDataTables" colspan="2">
            <br /><input name="btnSubmit" id="btnSubmit" type="submit" value="Update" disabled="disabled" />
            &nbsp;<input id="reset" type="button"  onclick="javascript:ResetAddPw();" value="Reset"/>            
            <%
            If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                Response.Write ("&nbsp;Don't Save Data&nbsp;&nbsp;")
                Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                Response.Write ("&nbsp;")
            End If          
            %>   
            &nbsp;&nbsp;<input name="ErrBox" id="ErrBox" type="hidden" value="<%=ErrMsg%>" />
			&nbsp;&nbsp;<input name="txtUserID" id="txtUserID" type="hidden" value="" />
			&nbsp;&nbsp;<input name="frmName" id="frmName" type="hidden" value="frmAddPW" />
        </td>
        </tr>
    </table>
    </form>
</body>  
</html>

