<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

If Session("ConnMachinefaults") = "" Then Response.Redirect("Admin.asp")


Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

Dim MachineRs
Dim MachineList

Set MachineRs = Server.CreateObject("ADODB.Recordset")
MachineRs.ActiveConnection = Session("ConnMachinefaults")
MachineRs.Source = "SELECT Distinct * From Machine Order By Id"
MachineRs.CursorType = Application("adOpenForwardOnly")
MachineRs.CursorLocation = Application("adUseClient")
MachineRs.LockType = Application("adLockReadOnly")
MachineRs.Open
Set MachineRs.ActiveConnection = Nothing

While Not MachineRs.EOF
    If MachineList = "" Then    
        MachineList = Ucase(MachineRs("MachineName"))
    Else
        MachineList = MachineList & "#" & Ucase(MachineRs("MachineName"))
    End If
    MachineRs.MoveNext
Wend

MachineRs.Close
Set MachineRs = Nothing

Dim MachineTypeRs
Dim MachineTypeSql

MachineTypeSql = "SELECT DISTINCT MachineTypeId, MachineType From Machine"
MachineTypeSql = MachineTypeSql & " Order By MachineType"

Set MachineTypeRs = Server.CreateObject("ADODB.Recordset")
MachineTypeRs.ActiveConnection = Session("ConnMachinefaults")
MachineTypeRs.Source = MachineTypeSql
MachineTypeRs.CursorType = Application("adOpenForwardOnly")
MachineTypeRs.CursorLocation = Application("adUseClient")
MachineTypeRs.LockType = Application("adLockReadOnly")
MachineTypeRs.Open
Set MachineTypeRs.ActiveConnection = Nothing

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
    <script type="text/javascript" src="JsFiles/MachineFaultAjaxFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px"  onload="javascript:MachinePageLoad('Add');">           

<form action="MfUpdateMachine.asp" method="post" name="frmAdMachine" id="frmAddMachine" onsubmit="return ValidateMachine('Add');" >
<table style="width: 100%; position: absolute; top: 0px; padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100"> 
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
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Add&nbsp;Machine</font>
        </td>
        <td align="right" width="33%">&nbsp;</td>
     </tr>	
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: 40%;" >        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Machine&nbsp;Name</font></td>
            <td valign="middle" width="300px">
            <input name="txtMachineID" type="text" id="txtMachineID" class="inputboxes" onfocus='this.select();' onmouseup='return false;' 
                onkeypress='return DisableEnterKey(event);' onkeyup="javascript:AddMachineEnableType();"  value="" maxlength="48" />       
            </td>
            <td align="left">&nbsp;&nbsp</td>                
        </tr>            
                
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>          
        <tr >                                                           
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Machine&nbsp;Type</font></td>
            <td align="left" >
                <select id="cboMachineType" name="cboMachineType" onchange="javascript:MachineType();" disabled="disabled">
                    <option value="">Select&nbsp;Machine&nbsp;Type</option>
                    <% 
                    While Not MachineTypeRs.EOF
                        Response.Write ("<option value='" & MachineTypeRs("MachineTypeId") & "'>" & MachineTypeRs("MachineType") & "</option>") & vbcrlf 
                        MachineTypeRs.MoveNext
                    Wend 
                    MachineTypeRs.Close
                    Set MachineTypeRs = Nothing             
                    %>
                </select>
            </td>
            <td align="left" >&nbsp;                
            </td>                
        </tr>
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr>
        <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Fault&nbsp;Groups</font></td>
            <td colspan="1" align="left">
                <select id="cboFaultGroup" name="cboFaultGroup" onchange="javascript:FaultSelectNewMachine();" disabled="disabled" >
                    <option value="">Select&nbsp;Fault&nbsp;Group</option>             
                </select>                          
            </td>
        
            <td align="left" >
                <label id="Grouplist" style="text-align: left; font-size: medium; color: #009933;"></label>
            </td>
        </tr>
        
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
               
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Active</font></td>
            <td valign="middle" width="300px">
                <input type="checkbox" id="chkActive" name="chkActive" />       
            </td>
            <td align="right">&nbsp;</td>                
        </tr> 
</table>
                                                        
<table style="width: 100%; position: absolute; bottom: 12%; padding-right: 20px; padding-left: 20px;">
        <tr>
            <td width="10%">&nbsp;
                <%
                    If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                        Response.Write ("Don't Save Data&nbsp;&nbsp;")
                        Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                    End If          
                 %>            
            </td>
            <td width="10%">
                <input id="btnSubmit" name="btnSubmit" type="submit" value="Update" />                   
            </td>
            <td width="10%">
                <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            </td>
            <td >
                
                <%If InputType = "text" Then response.Write "frmName"%>
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmAddMachine" />
                <%If InputType = "text" Then response.Write "hMachineList"%>
                <input type="<%=InputType %>" name="hMachineList" id="hMachineList" value="<%=MachineList%>"/>
                <%If InputType = "text" Then response.Write "hMachineType"%>
                <input type="<%=InputType %>" name="hMachineType" id="hMachineType" value=""/>
                <%If InputType = "text" Then response.Write "hMachineChange"%>
                <input type="<%=InputType %>" name="hMachineChange" id="hMachineChange" value="1"/> 
                <%If InputType = "text" Then  Response.Write "hGroupId"%>
                <input type="<%=InputType%>" name="hGroupId" id="hGroupId" value=""/>
                <%If InputType = "text" Then  Response.Write "hSelectId"%>
                <input type="<%=InputType%>" name="hSelectId" id="hSelectId" value=""/>
            </td>            
        </tr>          
    </table>    
    
</form>   
</body>  
</html>


