<%@language="vbscript" codepage="1252"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<%
'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp") 

response.expires = -1
response.AddHeader "Pragma", "no-cache"
response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"

Dim ErrMsg
Dim ReturnAddress

ErrMsg = Session("JobNoError")

Session("JobNo") = ""
Session("JobNoError") = ""
Session("JobType") = ""
Session("JobTypeNo") = ""

ReturnAddress = "NcSelectOption.asp"

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>NC Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/NCReportsJSFunc.js"></script>
    <link rel="shortcut icon" href="Images/warning.png" type="image/x-icon" />    
</head>

<body style="padding: 0px; margin: 0px" onload="PageLoadChk();">

<form action="NcCheckJobNo.asp" method="post" name="frmJobNo" id="frmJobNo" >

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
	    <td align="left" width="33%">
            <a href ="javascript:window.location.replace('<%=ReturnAddress%>');" style="font-size:12px; color: <%=NewBlue%>;">Return to option page</a>
        </td>
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" /> 
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">NC&nbsp;Reporting&nbsp;(<%=Session("UserName")%>)</font>
        </td>
        <td align="right" width="33%"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>; font-size:12px; ">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
	</tr>	
</table>


<table  style="width: 100%; position: absolute; top: 42%; padding-right: 10px; padding-left: 10px;" >
    <tr>   
        <td >			
            <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font size="3" color="<%=NewBlue%>">Enter job No</font></p>
            <p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="txtJobNo" id="txtJobNo" style="color: black;" />
            &nbsp;<input name="ErrBox" id="ErrBox" type="hidden" value="<%=ErrMsg%>" />
            
           
            </p>
            <%
            If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then 
                If Session("UserId") = 4 Then Response.Write "Live test on 145995, home 145433 BIG JOB 146509" 
            End if           
            %>
            
            			
        </td>
    </tr>
</table>

<table  style="width: 100%; position: absolute; top: 72%; padding-right: 10px; padding-left: 10px;" >
    <tr>   
        <td >
            <input style="position: relative; left: 30px;" type="submit" value="Submit" />
        </td>
    </tr>
</table>

<table  style="width: 100%; position: absolute; bottom: 5px; padding-right: 10px; padding-left: 10px;">
    <tr>  
        <td height="50" >
            <hr style="border-style: none; height: 1px; background-color: <%=NewCyan%>; display: block;" />
            <p align="center"><font size="2.5"> MediaCo Ltd. Churchill Point, Churchill Way, 
            Trafford Park, Manchester M17 1BS</font></p>
        </td>
    </tr>
</table>
</form>
</body> 
</html>
 