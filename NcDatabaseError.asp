<%@Language="VBScript" Codepage="1252" %>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->  

<% 
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store" 

Session.Abandon
%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>Database Error</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/error.ico" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
</head>

<body style="padding: 0px; margin: 0px">

<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100">
            <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
        </td>            
    </tr>
    <tr>
	    <td height="8" valign="top">
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	        
	    </td>
    </tr>        
</table>

<table style="width: 100%; position: absolute; top: 45%; padding-right: 20px; padding-left: 20px;" > 
    <tr>
        <td>
            <h2 align="center" style="color: #FF0000">Database Error</h2>            
        </td>
    </tr>    
</table>

<table style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
    <tr>  
        <td height="50" >
            <hr style="width: 98%; border-style: none;  height: 1px; background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way,> 
            Trafford Park, Manchester M17 1BS</font></p>
        </td>
    </tr>
</table> 

</body>

</html>
