<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

If Session("ConnMachinefaults") = "" Then Response.Redirect("Admin.asp")

Dim CentreLogo
Dim CentreAlt

CentreLogo = "Images/" & Session("CentreLogo") 
If Session("CentreLogo") = "blank.jpg" Then
    CentreAlt = ""
Else
    CentreAlt = Session("Client") & " Logo"
End If

Dim MachineSql
Dim MachineRs

MachineSql = "SELECT Machine.Id, Machine.MachineName, Machine.HasLogs FROM Machine Where (HasLogs = 0) AND (Active=1) Order By MachineTypeId, MachineName" 

Set MachineRs = Server.CreateObject("ADODB.Recordset")
MachineRs.ActiveConnection = Session("ConnMachinefaults") 
MachineRs.Source = MachineSql
MachineRs.CursorType = Application("adOpenForwardOnly")
MachineRs.CursorLocation = Application("adUseClient") 
MachineRs.LockType = Application("adLockReadOnly")
MachineRs.Open
Set MachineRs.ActiveConnection = Nothing

Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>Machine Faults Admin</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/MachineFaultJSFunc.js"></script>
    <script type="text/javascript" >
        function Relocate() 
        {
            var Id = document.getElementById("hMachineId").value;
            window.location.replace('MfEditMachine.asp?Id=' + Id);
        }                
    </script>
</head>

<body style="padding: 0px; margin: 0px" >           


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
        <td align="left" valign="top" width="33%">
            &nbsp;<a href ="javascript:window.location.replace('MfAdmin.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;MF&nbsp;admin&nbsp;page</a>
        </td>
	    <td  align="center" width="34%">
	        <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
	        <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Machine&nbsp;Faults</font>
	    </td>
	    <td align="right" width="33%">&nbsp;</td>
	</tr>    
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: 40%;" >
    <tr >
        <td colspan="2"><font size="3" color="#FF0000">&nbsp;&nbsp;Only&nbsp;Machines&nbsp;without&nbsp;any&nbsp;records&nbsp;can&nbsp;be&nbsp;edited</font></td>
    </tr> 
    <tr>
        <td height="20px" colspan="3">&nbsp;</td>
    </tr>      
    <tr >
        <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Select&nbsp;Machine</font></td>
        <td  align="left" >
            <!--Machine-->
            <select name="cboMachine" id="cboMachine" onchange="javascript:SelectMachine();" >
            <option value="">Select Machine</option>
            <%
            While Not MachineRs.EOF
                Response.Write ("<option value='" & Trim(MachineRs("Id")) & "'>" & Trim(MachineRs("MachineName")) & "</option>)") & VbCrLf
                MachineRs.MoveNext
            Wend
            
            MachineRs.Close
            Set MachineRs = Nothing
            %>
            </select>   
        </td> 
            
    </tr>            
    <tr>
        <td height="20px" colspan="2">&nbsp;</td>
    </tr>                              
         
</table>
                                                        
<table style="width: 100%; position: absolute; bottom: 12%; padding-right: 20px; padding-left: 20px;">
        <tr>
            <td>&nbsp;&nbsp;</td>
            <td >
                &nbsp;&nbsp;<input id="btnsubmit" name="btnsubmit" type="button" value="Submit" onclick="Relocate();" disabled="disabled" />&nbsp;&nbsp;
		        <input id="btnReset" name="btnReset" onclick="javascript:ResetPage();" type="button" value=" Reset " />&nbsp;&nbsp;                   
                <input type="<%=InputType %>" name="hMachineId" id="hMachineId" value=""/>
            </td>
        </tr>          
    </table>    
    
<!--</form>-->   
</body>  
</html>

