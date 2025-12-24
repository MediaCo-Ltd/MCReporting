<%@Language="VBScript" Codepage="1252" EnableSessionState=False%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store" 

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>Session Expired</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/X1.ico" type="image/x-icon" />
    <script type="text/javascript" >
        function GoBack()
        { window.location.replace("IhrJobNo.asp"); }
    </script>
</head>

<body style="padding: 0px; margin: 0px">

<table class="NonDataTables" style="width: 100%; position: absolute; top: 0px;">
    <tr>
		<td align="left" valign="bottom" height="100">
            <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
        </td>        
	</tr>
	<tr>
		<td height="8" valign="top" >
		    <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		</td>
	</tr>	
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 45%;" >
    <tr>   
        <td align="center">
            <h3>Your Session has timed out</h3>
        </td>
    </tr>
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 60%;" >
    <tr>   
        <td align="center">
            <h4><a href ="javascript:GoBack();" style="color: <%=NewBlue%>" >Return to main page</a></h4>
        </td>
    </tr>
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; bottom: 5px;">
    <tr>  
        <td height="50" >
            <hr style="border-style: none; width: 98%; height: 1px;  background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
            Trafford Park, Manchester M17 1BS <br />Tel:(+44)161 875 2020 Fax:(+44)161 873 7740</font></p>
        </td>
    </tr>
</table>




</body>  
</html>