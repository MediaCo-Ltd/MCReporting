<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<!--#include file="..\##GlobalFiles\connMCReportsDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\connQualityControlDB.asp" -->

<% 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

Dim UserRedirect
Dim UserEmailRedirect
Dim HsAdminRedirect
Dim NcAdminRedirect
Dim MfAdminRedirect
Dim UserDormantRedirect
Dim AddUserRedirect

AddUserRedirect = "javascript:Relocate('AddUser.asp?rt=1');"
UserRedirect =  "javascript:Relocate('UsersOnLine.asp');"
UserDormantRedirect = "javascript:Relocate('Usersdormant.asp');"
UserEmailRedirect = "javascript:Relocate('EmailUsers.asp')"
HsAdminRedirect = "javascript:Relocate('HsAdmin.asp');" 
NcAdminRedirect = "javascript:Relocate('NcAdmin.asp');" 
MfAdminRedirect = "javascript:Relocate('MfAdmin.asp');" 


Dim HideRedoRs
Set HideRedoRs = Server.CreateObject("ADODB.Recordset")
HideRedoRs.ActiveConnection = strConStatus   
HideRedoRs.Source = "SELECT IndoorRedoClientOnly FROM Status"
HideRedoRs.CursorType = Application("adOpenForwardOnly")
HideRedoRs.CursorLocation = Application("adUseClient")
HideRedoRs.LockType = Application("adLockReadOnly")
HideRedoRs.Open
Set HideRedoRs.ActiveConnection = Nothing

Session("HideIndoorRedo") = Not(HideRedoRs("IndoorRedoClientOnly"))
'Session("HideIndoorRedo") = HideRedoRs("IndoorRedoClientOnly")

Session("ConnQC") = strConnQualityControlDB

HideRedoRs.Close
Set HideRedoRs = Nothing

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>MediaCo Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/PenPad64.png" type="image/x-icon" />
    <script type="text/javascript" >
        function Relocate(strUrl) 
        {
           window.location.replace(strUrl);
        }                
    </script>     
</head>

<body id="admin" style="padding: 0px; margin: 0px">

<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
        <td align="left" valign="bottom" height="100" colspan="3">
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td>               
    </tr>
	<tr>
		<td height="8" valign="top" colspan="3">
		    <hr style="border-style: none;  height: 4px; background-color: #07AFEE; display: block;" />
		</td>
	</tr>
	
	<tr>
	    <td align="left" valign="top" width="33%">
            &nbsp;<a href ="javascript:window.location.replace('Login.asp');" style="font-size:12px; color:<%=NewBlue%>">Go&nbsp;to&nbsp;Logon&nbsp;page</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="Reporting logo" src="Images/PenPad64.png" style="width: 20px; height: 20px;" />
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;">MediaCo&nbsp;Reporting&nbsp;Admin</font>
	    </td>
	    <td align="left" valign="top" width="33%" >&nbsp;</td>
	</tr>		
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 32%;" >
    <tr>
        <td >
            <label style="font-style: italic; font-size: 18px">Please&nbsp;Note.</label>
            <label style="font-size: 16px">If you have a current reporting session open.&nbsp;Logout of your reporting session &amp; refresh this page before continuing</label>
        </td>
    </tr>
    <tr>   
        <td valign="middle" >
            <!-- colour #0069AA New blue or cyan #07AFEE  -->     
            <p>
                <font size="3" color="<%=NewBlue%>"> 
                
                    <br  /><br  />                    
                    &nbsp;<input id="Add" type="radio" onclick="<%=AddUserRedirect%>"/>&nbsp;Add&nbsp;User               
                    <br  /><br  />                    
                    &nbsp;<input id="User" type="radio" onclick="<%=UserRedirect%>"/>&nbsp;Active&nbsp;Users
                    <br  /><br  />
                    &nbsp;<input id="UserDormant" type="radio" onclick="<%=UserDormantRedirect%>"/>&nbsp;Dormant&nbsp;Users
                    <br  /><br  />
                    &nbsp;<input id="Email" type="radio" onclick="<%=UserEmailRedirect%>"/>&nbsp;Email&nbsp;Settings
                    <br  /><br  />                 
                    &nbsp;<input id="HS" type="radio" onclick="<%=HsAdminRedirect%>"/>&nbsp;HS&nbsp;Admin 
                    <br  /><br  />                
                    &nbsp;<input id="NC" type="radio" onclick="<%=NcAdminRedirect%>" disabled="disabled" />&nbsp;NC&nbsp;Admin
                    <br  /><br  />  
                    &nbsp;<input id="MF" type="radio" onclick="<%=MfAdminRedirect%>" />&nbsp;MF&nbsp;Admin 
                </font>
            </p>
        </td>
    </tr>   
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; bottom: 5px;">
    <tr>  
        <td height="50" >
            <hr style="border-style: none; width: 98%; height: 1px;  background-color: #07AFEE; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
            Trafford Park, Manchester M17 1BS</font></p>
        </td>
    </tr>
</table>
</body>  
</html>