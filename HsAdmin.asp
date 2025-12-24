<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<!--#include file="..\##GlobalFiles\connMCReportsDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<% 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

If Session("ConnMcLogon") = "" Then Response.Redirect("Admin.asp")

Dim ReasonRedirectAdd
Dim ReasonRedirectEdit
Dim DormantRedirectEdit


ReasonRedirectAdd = "javascript:Relocate('HsAddReason.asp')"    
ReasonRedirectEdit = "javascript:Relocate('HsSelectReason.asp')"
DormantRedirectEdit = "javascript:Relocate('HsDormant.asp')" 


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/HSReportsJSFunc.js"></script>
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
            &nbsp;<a href ="javascript:window.location.replace('Admin.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;Mediaco&nbsp;Admin&nbsp;</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;">HS&nbsp;Reporting&nbsp;Admin</font>
	    </td>
	    <td  align="right" width="33%">
	        &nbsp;
	    </td>
	
	</tr>		
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 40%;" >
    <tr>   
        <td valign="middle" >
            <!-- colour #0069AA New blue or cyan #07AFEE  -->     
            <p>
            <font size="3" color="<%=NewBlue%>">                
                &nbsp;<input id="ReasonAdd" type="radio" onclick="<%=ReasonRedirectAdd%>" />&nbsp;Add&nbsp;Reason
                <br  /><br  />  
                &nbsp;<input id="ReasonEdit" type="radio" onclick="<%=ReasonRedirectEdit%>" disabled="disabled" />&nbsp;Edit&nbsp;Reason 
                <br  /><br  />  
                &nbsp;<input id="Dormant" type="radio" onclick="<%=DormantRedirectEdit%>" />&nbsp;Show/Hide&nbsp;Records                
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