<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim MachineSql
Dim MachineRs

MachineSql = "SELECT Machine.Id, Machine.MachineName, Machine.MachineTypeId, Machine.Active, Machine.FaultGroups" 
MachineSql = MachineSql & " FROM Machine Where (Active = 1) Order By Machine.MachineTypeId, Machine.MachineName"

Set MachineRs = Server.CreateObject("ADODB.Recordset")
MachineRs.ActiveConnection = Session("ConnMachinefaults") 
MachineRs.Source = MachineSql
MachineRs.CursorType = Application("adOpenForwardOnly")
MachineRs.CursorLocation = Application("adUseClient") 
MachineRs.LockType = Application("adLockReadOnly")
MachineRs.Open
Set MachineRs.ActiveConnection = Nothing

Dim InputType
Dim InputSize
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    InputType = "text"
    InputSize = "size = '4'"
Else
    InputType = "hidden"
    InputSize = ""
End If

Session("AddFolder") = ""
Session("AddFolder") = Session.SessionID
Session("DeleteFolder") = Session("MfImagePath") & Session.SessionID

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>Machine Faults</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/MachineFaultJSFunc.js"></script>
    <script type="text/javascript" src="JsFiles/MachineFaultAjaxFunc.js"></script>    
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:AddLogLoad();">
    <table style="width: 100%; padding-right: 10px; padding-left: 10px;" >
        <tr>
		    <td align="left" valign="bottom" height="100" colspan="3">
                <img align="left" alt="mediaco logo" src='Images/mediaco_logo.jpg' width="160" />
            </td>            
        </tr>
	    <tr>
		    <td height="8" valign="top" colspan="3">
		        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
		    </td>
	    </tr>	
        <tr>
            <td align="left" style="height: 20px" valign="bottom" width="33%" >&nbsp;&nbsp;
                <a href ="javascript:GoBackOption();" style="font-size:12px; color: <%=NewBlue%>;">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>
            </td>
             <td height="20px" width="34%" valign="bottom" align="center"> <!-- color #0069AA; Blue-->
                <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" /> 
                <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">New&nbsp;Fault&nbsp;Log</font>
            </td>
            <td valign="bottom" align="right" width="33%">
                <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;off</a>&nbsp;&nbsp;
            </td> 
        </tr>
        <tr>
            <td colspan="3" nowrap="nowrap">
                <p id="toptext">
                    <font size="3" style="color: #0069AA">                        
                    &nbsp;&nbsp;Fields&nbsp;marked&nbsp;</font><font size="3" style="color: red">*</font>
                    <font size="3" style="color: #0069AA">are&nbsp;mandatory.</font>                    
                </p>
            </td>
        </tr>                
    </table>
    <br />
    <br />  
    <form method="post" action="MfUpdateLog.asp"  name="frmAddLog" id="frmAddlog" onsubmit="return MFValidateLog('Add');">
    
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;" >          
        <tr>
            <td align="left" width="60%" >
                <label style=" color: <%=NewBlue%>">&nbsp;Fault&nbsp;Details<font style="color: red">*</font></label><br />
                <textarea  rows="15" id="txtError" name="txtError" style="text-align: left; " cols="110" ></textarea>
            </td>            
            <td align="left" valign="top" style="text-align: left; padding-left: 6px;">
                <label id="GroupTitle"  style=" color: <%=NewBlue%>"></label><br /><br />
                <label id="Grouplist" style="text-align: left; font-size: medium; color: #009933;"></label>
            </td>            
        </tr>           
    </table>
    
    <br />
    <br />
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;">
 
    <tr>   
            <td valign="middle" width="100px" ><label style=" color: <%=NewBlue%>">&nbsp;Select&nbsp;Machine&nbsp;<font style="color: red">*</font></label></td>                      
            <td valign="middle"  width="200px" >
            <select id='cboMachine' name='cboMachine' onchange='javascript:AddMachine();'>
                <option value="">Select&nbsp;Machine</option>
                <%
                While Not MachineRs.EOF                    
                    Response.Write ("<option value='" & MachineRs("Id") & "#" & MachineRs("MachineTypeId") & "#" & MachineRs("FaultGroups")  & "'>" & MachineRs("MachineName") & "</option>") & vbcrlf 
                    MachineRs.MoveNext                
                Wend
                
                MachineRs.Close
                Set MachineRs = Nothing
                %>                               
                </select>               
                       
            </td>             
            
            <td valign="middle" width="100px" ><label style=" color: <%=NewBlue%>">&nbsp;Select&nbsp;Severity&nbsp;<font style="color: red">*</font></label></td>         
            <td valign="middle" >
                
                <select id="cboSeverity" name="cboSeverity" onchange="javascript:Severity();" disabled="disabled">
                    <option value="">Select&nbsp;Severity</option>
                    <option value="1">Minor</option>
                    <option value="2">Medium</option>
                    <option value="3">Critical</option>
                </select>                
            </td>             
                      
        </tr>
        
        <tr>
            <td height="5px" colspan="4"></td>
        </tr>
        
        
        <tr>
            <td valign="middle"  ><label style=" color: <%=NewBlue%>">&nbsp;Recurring&nbsp;Fault</label></td>
            <td align="left"><input type="checkbox" id="chkRecurring" name="chkRecurring"  disabled="disabled"/></td><!--onclick="javascript:ShowGroup()"-->
            <td align="left"><label id="GroupLabel" style="color: <%=NewBlue%>">&nbsp;Select&nbsp;Fault&nbsp;Group&nbsp;<font style="color: red">*</font></label></td>
                                                            <!--visibility: hidden;--> 
            <td align="left" id="GroupCbo" ><!--style="visibility: hidden"-->
                <select title="Multiple options can be selected" id="cboFaultGroup" name="cboFaultGroup" onchange="javascript:FaultSelectAdd();"  disabled="disabled">
                    <option value="">Select&nbsp;Fault&nbsp;Group</option>             
                </select>
            </td>
        </tr>
       
       <tr>
            <td height="5px" colspan="4"></td>
        </tr>
        
        
        <tr>
            <td valign="middle"  ><label style=" color: <%=NewBlue%>">&nbsp;Fault&nbsp;Fixed</label></td>
            <td colspan="3" align="left"><input type="checkbox" id="chkFixed" name="chkFixed" disabled="disabled"/></td>
        </tr>
       
        <tr>
            <td height="20px" colspan="4"></td>
        </tr>
        
        <tr>
	        <td valign="top" height="30" colspan="4" >  <!--disabled="disabled"-->    
		        <input id="btnSubmit" name="btnSubmit" type="submit"  value="Update"  />&nbsp;&nbsp;
		        <input id="btnReset" name="btnReset" onclick="javascript:ResetPage();" type="button" value=" Reset " />&nbsp;&nbsp;
		        <input id="btnAdd" name="btnAdd" onclick="javascript:AddImage('A');" type="button" value="Add Image" title="Any images added will be viewable after record is saved"/>&nbsp;&nbsp;
		        <!--<font size="3" style="color: black">&nbsp;&nbsp;Any images added will be viewable after record is saved</font>-->
		        
		        <%
                If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                    Response.Write ("Don't Save Data&nbsp;&nbsp;")
                    Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                    Response.Write ("&nbsp;&nbsp;")
                End If          
                %>
                <br /><br />    
                <%If InputType = "text" Then  Response.Write "hMachine"%>
                <input type="<%=InputType%>" name="hMachine" id="hMachine" value="" <%=InputSize %> />&nbsp;&nbsp; 
                <%If InputType = "text" Then  Response.Write "hSeverity"%>
                <input type="<%=InputType%>" name="hSeverity" id="hSeverity" value="" <%=InputSize %> />&nbsp;&nbsp;
                <%If InputType = "text" Then Response.Write "hType"%>
                <input type="<%=InputType%>" name="hType" id="hType" value="" <%=InputSize %> />&nbsp;&nbsp;
                <%If InputType = "text" Then  Response.Write "hGroups"%>
                <input type="<%=InputType%>" name="hGroups" id="hGroups" value="" <%=InputSize %> />&nbsp;&nbsp;
                <%If InputType = "text" Then  Response.Write "hGroupId"%>
                <input type="<%=InputType%>" name="hGroupId" id="hGroupId" value="" <%=InputSize %> />&nbsp;&nbsp;
                <%If InputType = "text" Then  Response.Write "hSelectId"%>
                <input type="<%=InputType%>" name="hSelectId" id="hSelectId" value="" <%=InputSize %> />&nbsp;&nbsp;
                <%If InputType = "text" Then  Response.Write "Locked"%>
                <input type="<%=InputType %>" name="" id ="Locked" value="<%=Session("Locked")%>"  <%=InputSize %> />&nbsp;&nbsp;
                <input type="<%=InputType%>" name="frmName" id="frmName" value="frmAddLog"  />
                
                <input type="<%=InputType%>" name="SessionID" id="SessionID" value="<%=Session.SessionID%>" />
                
            </td>
        </tr>
        <tr>
            <td height="49px" colspan="4"></td>
        </tr> 
    </table>  
    
    </form>          
   
</body>  
</html>

<%
If Session("Locked") = Cbool(True) Then 
    If Application("UsersOnLine") > 0 Then
        Application("UsersOnLine") = Application("UsersOnLine") -1
    End If
    Session.Abandon  
End If 
%>