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


Dim InputType
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
Else
    InputType = "hidden"
End If

Dim GroupRs
Dim GroupList

Set GroupRs = Server.CreateObject("ADODB.Recordset")
GroupRs.ActiveConnection = Session("ConnMachinefaults")
GroupRs.Source = "SELECT * From FaultGroups Order By Description"
GroupRs.CursorType = Application("adOpenForwardOnly")
GroupRs.CursorLocation = Application("adUseClient")
GroupRs.LockType = Application("adLockReadOnly")
GroupRs.Open
Set GroupRs.ActiveConnection = Nothing

While Not GroupRs.EOF
    If GroupList = "" Then    
        GroupList = Ucase(GroupRs("Description"))
    Else
        GroupList = GroupList & "#" & Ucase(GroupRs("Description"))
    End If
    GroupRs.MoveNext
Wend

GroupRs.Close
Set GroupRs = Nothing

Dim GroupTypeRs
Dim GroupTypeSql

GroupTypeSql = "SELECT DISTINCT MachineTypeId, MachineType From Machine"
GroupTypeSql = GroupTypeSql & " Order By MachineType"

Set GroupTypeRs = Server.CreateObject("ADODB.Recordset")
GroupTypeRs.ActiveConnection = Session("ConnMachinefaults")
GroupTypeRs.Source = GroupTypeSql
GroupTypeRs.CursorType = Application("adOpenForwardOnly")
GroupTypeRs.CursorLocation = Application("adUseClient")
GroupTypeRs.LockType = Application("adLockReadOnly")
GroupTypeRs.Open
Set GroupTypeRs.ActiveConnection = Nothing

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
</head>

<body style="padding: 0px; margin: 0px" >           

<form action="MfUpdateGroup.asp" method="post" name="frmAddGroup" id="frmAddGroup"  onsubmit="return ValidateGroup('Add');" >
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
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Add&nbsp;Group</font>
        </td>
        <td align="right" width="33%">&nbsp;</td>
     </tr>	
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: 40%;" >        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Group&nbsp;Name</font></td>
            <td valign="middle" width="300px">
            <input name="txtGroupID" type="text" id="txtGroupID" class="inputboxes" onfocus='this.select();' onmouseup='return false;' 
                onkeypress='return DisableEnterKey(event);' value="" maxlength="48" />       
            </td>
            <td align="left">&nbsp;&nbsp</td>                
        </tr>            
                
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>          
        <tr >                                                           
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Machine&nbsp;Type</font></td>
            <td align="left" ><!--style="visibility: hidden"-->
                <select id="cboGroupType" name="cboGroupType" onchange="javascript:GroupType();"><!--  -->
                    <option value="">Select&nbsp;Group&nbsp;Type</option>
                    <option value="0">Global</option>
                    <%
                     
                    While Not GroupTypeRs.EOF
                        Response.Write ("<option value='" & GroupTypeRs("MachineTypeId") & "'>" & GroupTypeRs("MachineType") & "</option>") & vbcrlf 
                        GroupTypeRs.MoveNext
                    Wend 
                    GroupTypeRs.Close
                    Set GroupTypeRs = Nothing             
                    %>
                </select>
            </td>
            <td align="left" >&nbsp;                
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
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmAddGroup" />
                <%If InputType = "text" Then response.Write "hGroupList"%>
                <input type="<%=InputType %>" name="hGroupList" id="hGroupList" value="<%=GroupList%>"/>
                <%If InputType = "text" Then response.Write "hGroupType"%>
                <input type="<%=InputType %>" name="hGroupType" id="hGroupType" value=""/>
                <%If InputType = "text" Then response.Write "hGroupChange"%>
                <input type="<%=InputType %>" name="hGroupChange" id="hGroupChange" value="1"/>
            </td>            
        </tr>          
    </table>    
    
</form>   
</body>  
</html>


