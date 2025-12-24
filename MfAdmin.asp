<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<!--#include file="..\##GlobalFiles\connMCReportsDB.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<% 

If Session("ConnMcLogon") = "" Then Response.Redirect("Admin.asp")

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

Dim MachineRedirectAdd
Dim MachineRedirectEdit
Dim GroupRedirectAdd
Dim GroupRedirectEdit
Dim DormantRedirectEdit
Dim MachineRedirectEditDormant
Dim MachineStatusEdit

MachineRedirectAdd = "javascript:Relocate('MfAddMachine.asp')"    
MachineRedirectEdit = "javascript:Relocate('MfSelectMachine.asp')"
GroupRedirectAdd = "javascript:Relocate('MfAddGroup.asp')"
DormantRedirectEdit = "javascript:Relocate('MfDormant.asp')" 
MachineRedirectEditDormant = "javascript:Relocate('MfSelectDormantMachine.asp')" 
MachineStatusEdit = "javascript:Relocate('MfEditMachineStatus.asp')"    

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>
    <title>Machine Faults</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
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
	        <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;">MF&nbsp;Reporting&nbsp;Admin</font>
	    </td>
	    <td align="left" valign="top" width="33%" >&nbsp;</td>
	</tr>		
</table>

<table class="NonDataTables" style="width: 100%; position: absolute; top: 34%;" >
    <tr>   
        <td valign="middle" >
            <!-- colour #0069AA New blue or cyan #07AFEE  -->     
            <p>
            <font size="3" color="<%=NewBlue%>">                
                &nbsp;<input id="MachineAdd" type="radio" onclick="<%=MachineRedirectAdd%>" />&nbsp;Add&nbsp;Machine
                <br  /><br  />  
                &nbsp;<input id="MachineEdit" type="radio" onclick="<%=MachineRedirectEdit%>" />&nbsp;Edit&nbsp;Machine 
                <br  /><br  />                
                &nbsp;<input id="GroupAdd" type="radio" onclick="<%=GroupRedirectAdd%>" />&nbsp;Add&nbsp;Group
                <br  /><br  />  
                &nbsp;<input id="Dormant" type="radio" onclick="<%=DormantRedirectEdit%>" />&nbsp;Show/Hide&nbsp;Records
                <br  /><br  />  
                &nbsp;<input id="MachineDormant" type="radio" onclick="<%=MachineRedirectEditDormant%>" />&nbsp;Edit&nbsp;Dormant&nbsp;Machine
                <br  /><br  />  
                &nbsp;<input id="MachineStatus" type="radio" onclick="<%=MachineStatusEdit%>" />&nbsp;Show/Hide&nbsp;Machine
                
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