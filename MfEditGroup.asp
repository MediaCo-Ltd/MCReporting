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

Dim GroupToGet
GroupToGet = Request.QueryString("Id")

Dim GroupRs
Dim GroupSql
Dim GroupData(2)

GroupSql = "SELECT * FROM FaultGroups Where (Id =" & GroupToGet & ")"

Set GroupRs = Server.CreateObject("ADODB.Recordset")
GroupRs.ActiveConnection = Session("ConnMachinefaults")
GroupRs.Source = GroupSql
GroupRs.CursorType = Application("adOpenForwardOnly")
GroupRs.CursorLocation = Application("adUseClient")
GroupRs.LockType = Application("adLockReadOnly")
GroupRs.Open
Set GroupRs.ActiveConnection = Nothing

If GroupRs.BOF = True Or GroupRs.EOF = True Then
    GroupRs.Close
    Set GroupRs = Nothing
    '## redirect to add with GroupToGet
Else
  
    GroupData(0) = GroupRs("Usage")
    GroupData(1) = GroupRs("Description")
    GroupData(2) = GroupRs("MachineTypeId")
    '## Also add to list
    
    GroupRs.Close
    Set GroupRs = Nothing
End If


Dim GroupNameList

Set GroupRs = Server.CreateObject("ADODB.Recordset")
GroupRs.ActiveConnection = Session("ConnMachinefaults")
GroupRs.Source = "SELECT * From FaultGroups Order By Description"
GroupRs.CursorType = Application("adOpenForwardOnly")
GroupRs.CursorLocation = Application("adUseClient")
GroupRs.LockType = Application("adLockReadOnly")
GroupRs.Open
Set GroupRs.ActiveConnection = Nothing

While Not GroupRs.EOF
    If GroupNameList = "" Then    
        GroupNameList = Ucase(GroupRs("Description"))
    Else
        GroupNameList = GroupNameList & "#" & Ucase(GroupRs("Description"))
    End If
    GroupRs.MoveNext
Wend

GroupRs.Close
Set GroupRs = Nothing


Set GroupRs = Server.CreateObject("ADODB.Recordset")
GroupRs.ActiveConnection = Session("ConnMachinefaults")
GroupRs.Source = "SELECT Distinct MachineType, MachineTypeId From Machine"  ' Where(MachineTypeId <> " & GroupData(2) & ")"
GroupRs.CursorType = Application("adOpenForwardOnly")
GroupRs.CursorLocation = Application("adUseClient")
GroupRs.LockType = Application("adLockReadOnly")
GroupRs.Open
Set GroupRs.ActiveConnection = Nothing


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

<form action="MFUpdateGroup.asp" method="post" name="frmEditGroup" id="frmEditGroup"  onsubmit="return ValidateGroup('Edit');">
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
            &nbsp;<a href ="javascript:window.location.replace('MfSelectGroup.asp');" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;select&nbsp;group&nbsp;page</a>
        </td>
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >Edit&nbsp;<%=GroupData(1)%>&nbsp;Data</font>
        </td>
        <td align="right" width="33%">&nbsp;</td>
     </tr>	
</table>
                                                        
                                                        
<table class="NonDataTables" style="width: 100%; position: absolute; top: 35%;" >        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Group&nbsp;Name</font></td>
            <td valign="middle" width="300px">
            <input name="txtGroupID" type="text" id="txtGroupID" class="inputboxes" onfocus='this.select();' onmouseup='return false;' 
                onkeypress='return DisableEnterKey(event);' value="<%=GroupData(1)%>" readonly="readonly" />       
            </td>
            <td align="left">
            &nbsp;&nbsp
            
            </td>
                
        </tr>            
                
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>          
        <tr >                                                           
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">Current&nbsp;Usage</font></td>
            <td align="left" >
            <input name="txtMachineID" type="text" id="txtMachineID" class="inputboxes" onfocus='this.select();' onmouseup='return false;' 
                onkeypress='return DisableEnterKey(event);' value="<%=GroupData(0)%>" readonly="readonly" />
                
            </td>
            <td align="left" >&nbsp;
                
            </td>                
        </tr>
        <tr>
            <td height="20px" colspan="3">&nbsp;</td>
        </tr>
        
        <tr >
            <td valign="middle" width="100px" class="inputlabel" colspan="1"><font size="3">New&nbsp;Usage</font></td>
            
            <td align="left" >
            <select id='cboMachine' name='cboMachine'  ><!--onchange='javascript:AddMachine("True");'-->
                <option value="">Select&nbsp;Machine&nbsp;Type</option>
                <%
                While Not GroupRs.EOF                    
                    Response.Write ("<option value='" & GroupRs("MachineTypeId")  & "'>" & GroupRs("MachineType") & "</option>") & vbcrlf 
                    GroupRs.MoveNext                
                Wend
                
                GroupRs.Close
                Set GroupRs = Nothing
                %>                               
                </select>               
                       
            </td>                   
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
                <input id="btnSubmit" name="btnSubmit" type="submit" value="Update"  />                   
            </td>
            <td width="10%">
                <input type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            </td>
            <td >
                <%If InputType = "text" Then response.Write "gId"%>
                <input type="<%=InputType %>" name="gId" id="gId" value="<%=GroupToGet%>" />
                <%If InputType = "text" Then response.Write "frmName"%>
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmEditGroup" />
                <%If InputType = "text" Then response.Write "hGroupList"%>
                <input type="<%=InputType %>" name="hGroupList" id="hGroupList" value="<%=GroupNameList%>"/>
                <%If InputType = "text" Then response.Write "hGroupType"%>
                <input type="<%=InputType %>" name="hGroupType" id="hGroupType" value="<%=GroupData(2)%>"/>
                <%If InputType = "text" Then response.Write "hGroupChange"%>
                <input type="<%=InputType %>" name="hGroupChange" id="hGroupChange" value="0"/> 
            </td>            
        </tr>          
    </table>    
    
</form>   
</body>  
</html>

<%Erase GroupData%>


