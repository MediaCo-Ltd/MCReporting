<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%> 
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
If Session("UserName") = "" Then Response.Redirect ("Login.asp")

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"

Dim CentreLogo
Dim CentreAlt

CentreLogo = "Images/" & Session("CentreLogo") 
If Session("CentreLogo") = "blank.jpg" Then
    CentreAlt = ""
Else
    CentreAlt = Session("Client") & " Logo"
End If

Dim InputType
Dim InputSize

If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    If Session("UserId") = 1 Or Session("UserId") = 4 Then
        InputType = "text"
        InputSize = "size = '6'"
    Else
        InputType = "hidden"
        InputSize = ""
    End If
Else
    InputType = "hidden"
    InputSize = ""
End If

Dim Count
Dim SubCount 

Dim ReturnLogType
Dim ReturnReason
Dim ReturnUser
Dim ReturnShift
Dim ReturnStartSerial
Dim ReturnEndSerial
Dim ReturnDateSource

'####################### All dropdowns are only populated if they exist in logs #####################

'###################### Log Types

Dim LogTypeRs
Dim LogTypeSql

LogTypeSql = "SELECT DISTINCT Usage, HasLogs, GroupId"
LogTypeSql = LogTypeSql & " FROM Reasons"
LogTypeSql = LogTypeSql & " WHERE (HasLogs = 1) AND (NOT (GroupId = 0)) ORDER BY Usage"

Set LogTypeRs = Server.CreateObject("ADODB.Recordset")
LogTypeRs.ActiveConnection = Session("ConnHSReports")
LogTypeRs.Source = LogTypeSql
LogTypeRs.CursorType = Application("adOpenForwardOnly")
LogTypeRs.CursorLocation = Application("adUseClient")
LogTypeRs.LockType = Application("adLockReadOnly")
LogTypeRs.Open
Set LogTypeRs.ActiveConnection = Nothing


'###################### Reasons

Dim ReasonRs
Dim ReasonSql

ReasonSql = "SELECT Id, Description, HasLogs, GroupId"
ReasonSql = ReasonSql & " FROM Reasons"
ReasonSql = ReasonSql & " WHERE (HasLogs = 1) ORDER BY Usage, Description"

Set ReasonRs = Server.CreateObject("ADODB.Recordset")
ReasonRs.ActiveConnection = Session("ConnHSReports")
ReasonRs.Source = ReasonSql
ReasonRs.CursorType = Application("adOpenForwardOnly")
ReasonRs.CursorLocation = Application("adUseClient")
ReasonRs.LockType = Application("adLockReadOnly")
ReasonRs.Open
Set ReasonRs.ActiveConnection = Nothing

'###################### Users


Dim UserRs
Dim UserSql

UserSql = "SELECT DISTINCT Logs.UserId FROM Logs"    
UserSql = UserSql & " ORDER BY Logs.UserId"

Set UserRs = Server.CreateObject("ADODB.Recordset")
UserRs.ActiveConnection = Session("ConnHSReports")
UserRs.Source = UserSql
UserRs.CursorType = Application("adOpenForwardOnly")
UserRs.CursorLocation = Application("adUseClient")
UserRs.LockType = Application("adLockReadOnly")
UserRs.Open
Set UserRs.ActiveConnection = Nothing

'###################### Shift

Dim ShiftRs
Dim ShiftSql

ShiftSql = "SELECT DISTINCT SelectedTimeSlot As TimeId," 
ShiftSql = ShiftSql & " CASE WHEN SelectedTimeSlot = 1 THEN 'Earlies' ELSE CASE WHEN SelectedTimeSlot = 2 THEN 'Lates' ELSE 'Nights' END END AS Shift"
ShiftSql = ShiftSql & " FROM Logs ORDER BY SelectedTimeSlot"

Set ShiftRs = Server.CreateObject("ADODB.Recordset")
ShiftRs.ActiveConnection = Session("ConnHSReports")
ShiftRs.Source = ShiftSql
ShiftRs.CursorType = Application("adOpenForwardOnly")
ShiftRs.CursorLocation = Application("adUseClient")
ShiftRs.LockType = Application("adLockReadOnly")
ShiftRs.Open
Set UserRs.ActiveConnection = Nothing

'###################### Location

Dim LocationRs
Dim LocationSql

LocationSql = "SELECT DISTINCT Location.LocationName, Logs.LocationId"
LocationSql = LocationSql & " FROM Logs INNER JOIN"
LocationSql = LocationSql & " Location ON Logs.LocationId = Location.Id"
                         
Set LocationRs = Server.CreateObject("ADODB.Recordset")
LocationRs.ActiveConnection = Session("ConnHSReports")
LocationRs.Source = LocationSql
LocationRs.CursorType = Application("adOpenForwardOnly")
LocationRs.CursorLocation = Application("adUseClient")
LocationRs.LockType = Application("adLockReadOnly")
LocationRs.Open
Set LocationRs.ActiveConnection = Nothing                         

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />    
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/HSReportsJSExportFunc.js"></script>
    <script type="text/javascript" src="JsFiles/HSReportsAjaxFunc.js"></script>    
    <script type="text/javascript" src="JsFiles/HSReportsDateTimePickerExportOnly_css.js"></script>
</head>

<body style="padding: 0px; margin: 0px" >            <!--onload="javascript:lockSubmit();"-->

<table style="width: 100%;  padding-right: 10px; padding-left: 10px;" >
    <tr>
        <td align="left" valign="bottom" colspan="3" >  
            <img align="left" alt="mediaco logo" src='Images/mediaco_logo.jpg' width="160" />
        </td>          
    </tr>
    <tr>
	    <td height="8" valign="top" colspan="3">
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	    </td>
    </tr>
    <tr>
        <td align="left" valign="bottom" width="33%">&nbsp;&nbsp;
            <a id="GoBack" href ="javascript:GoBackHSOption();" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;selection&nbsp;page</a>
        </td>        
        <td align="center" valign="bottom" width="33%">
            <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;">Custom&nbsp;Export</font>
        </td>
        <td align="right" valign="bottom" width="33%">
            <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;off</a>&nbsp;&nbsp;
        </td>   
     </tr>
     
     <tr>
        <td colspan="3" nowrap="nowrap">
            <p>
                <font size="3" style="color: #0069AA">                        
                &nbsp;&nbsp;Log&nbsp;Type&nbsp;must&nbsp;be&nbsp;selected.<br />
                &nbsp;&nbsp;Dropdowns&nbsp;are&nbsp;populated&nbsp;from&nbsp;existing&nbsp;logs&nbsp;and&nbsp;will&nbsp;only&nbsp;display&nbsp;data&nbsp;that&nbsp;exists.
                Making&nbsp;no&nbsp;selection&nbsp;will&nbsp;select&nbsp;all&nbsp;for&nbsp;that&nbsp;option.<br />
                &nbsp;&nbsp;Depending&nbsp;on&nbsp;combination&nbsp;of&nbsp;filters&nbsp;selected.&nbsp;You&nbsp;may&nbsp;end&nbsp;up&nbsp;with&nbsp;a&nbsp;empty&nbsp;report.
                Resulting&nbsp;display&nbsp;will&nbsp;allow&nbsp;you&nbsp;to&nbsp;download&nbsp;the&nbsp;report.
                </font>
            </p>
        </td>
    </tr>  	
</table>
<br />                                                       
<br />                                                        
<table class="NonDataTables" style="width: 100%" >
    <tr>
        <td class="supplierlabel" colspan="5" ><font size="2">Log&nbsp;Type</font></td>
    </tr>
    
    <tr>
        <td class="supplierinput" >    
            <select id="cboLogType" name="cboLogType" onchange="javascript:HSSelectLogType();">
            <option value="">Select&nbsp;Log&nbsp;Type</option>
            <%
            While Not LogTypeRs.EOF
                If Trim(LogTypeRs("Usage"))= "Incident" Then
                    Response.Write ("<option value='" & Trim(LogTypeRs("GroupId")) & "'>Near Miss</option>)") & VbCrLf
                Else
                    Response.Write ("<option value='" & Trim(LogTypeRs("GroupId")) & "'>" & Trim(LogTypeRs("Usage")) & "</option>)") & VbCrLf
                End If
                    LogTypeRs.MoveNext
                Wend
            
            LogTypeRs.Close
            Set LogTypeRs = Nothing
            %>
            <option value="4">All</option>
            </select>         
        </td>        
    </tr>
    
    <tr>
        <td height="15px" colspan="5"></td>
    </tr>
    
    <tr>
        <td class="supplierlabel" nowrap="nowrap" colspan="5" >
            <font size="2">&nbsp;</font>
        </td>
    </tr>
    
    <tr>
        <td class="supplierlabel" width="20%"><font size="2">Reasons</font></td>
        <td class="supplierlabel" width="20%"><font size="2">Logged By</font></td>
        <td class="supplierlabel" width="20%"><font size="2">Shift</font></td>
        <td class="supplierlabel" width="20%"><font size="2">Location&nbsp;(Only&nbsp;affected&nbsp;by&nbsp;Type)</font></td>    
        <td class="supplierlabel" width="20%"><font size="2">&nbsp;</font></td>    
    </tr>        
    <tr >
        <td class="supplierinput">        
        <!--Reasons populated by type selection-->
        <select name="cboReasons" id="cboReasons" onchange="javascript:HSSelectReasons();" style="width: 50%">
            <option value="">Select Reason</option>
            <%
            While Not ReasonRs.EOF
                Response.Write ("<option value='" & Trim(ReasonRs("Id")) & "'>" & Trim(ReasonRs("Description")) & "</option>)") & VbCrLf
                ReasonRs.MoveNext
            Wend
            
            ReasonRs.Close
            Set ReasonRs = Nothing
            %>            
        </select>   
        </td>
            
        <td class="supplierinput">
            <!--Log Created User-->                
            <select name="cboLogUser" id="cboLogUser" onchange="javascript:HSSelectLogUser();" style="width: 50%"> <!--onchange="javascript:AddPsg();"-->
            <option value="">Select User</option>
            <!--<option value="0">All</option>-->                         
            <%
            While Not UserRs.EOF
                Response.Write ("<option value='" & Trim(UserRs("UserId")) & "'>" & GetHsUser (UserRs("UserId")) & "</option>)") & VbCrLf
                UserRs.MoveNext
            Wend
            
            UserRs.Close
            Set UserRs = Nothing
            %>
            </select>          
        </td>
            
        <td class="supplierinput" >
            <!--Shift-->
            <select name="cboShift" id="cboShift" onchange="javascript:HSSelectShift();" style="width: 50%"> <!--style="width: 70%"-->
            <option value="" >Select Shift</option>
            <!--<option value="0" >All</option>-->
            <%
            While Not ShiftRs.EOF
                Response.Write ("<option value='" & Trim(ShiftRs("TimeId")) & "'>" & Trim(ShiftRs("Shift")) & "</option>)") & VbCrLf
                ShiftRs.MoveNext
            Wend
        
            ShiftRs.Close
            Set ShiftRs = Nothing
            %>
            </select>
        </td>  
            
        <td class="supplierinput" >
            <select name="cboLocation" id="cboLocation" onchange="javascript:HSSelectLocation();" style="width: 200px">
            <option value="">Select&nbsp;Location</option>
            <%
            While Not LocationRs.EOF                    
                Response.Write ("<option value='" & LocationRs("LocationId") & "'>" & LocationRs("LocationName") & "</option>") & vbcrlf 
                LocationRs.MoveNext                
            Wend
            
            LocationRs.Close
            Set LocationRs = Nothing
           
            %>   
        </select> 
        
        </td>        
        
        <td class="supplierinput">
        <!--Products-->
        
        </td> 
    </tr>            
    <tr>
        <td colspan="5" >&nbsp;</td>
    </tr>                                
    
    <tr>
        <td colspan="5" class="supplierlabel" nowrap="nowrap"><font size="2">
            Select&nbsp;Date&nbsp;source&nbsp;to&nbsp;enable&nbsp;date&nbsp;selection.&nbsp;Can&nbsp;be&nbsp;used&nbsp;on&nbsp;it's&nbsp;own&nbsp;to&nbsp;change&nbsp;grouping.
        </font>
        </td>
    </tr>
     <tr>
        <td colspan="5" class="supplierinput" >
            <select name="cboSource" id="cboSource" onchange="javascript:HSDateSource();"> 
            <option value="" >Select Source</option>
            <option value="0" >Created Date</option>
            <option value="1" >Last Updated</option>
        </select>
        </td>
    </tr>
</table>
    
<br /> 

<table class="NonDataTables" style="width: 100%;">
      
    <tr>
        <td class="supplierlabel" colspan="2" nowrap="nowrap">
            <font size="2">
            For&nbsp;exact&nbsp;date&nbsp;put&nbsp;end&nbsp;the&nbsp;same&nbsp;as&nbsp;start&nbsp;or&nbsp;click&nbsp;As&nbsp;Start&nbsp;button.
            Date&nbsp;Picker&nbsp;will&nbsp;allow&nbsp;any&nbsp;selection&nbsp;of&nbsp;dates.&nbsp;Date&nbsp;Picker&nbsp;overides&nbsp;dropdowns&nbsp;&amp;&nbsp;vice&nbsp;Versa.
            </font>
        </td>
    </tr>
   
    <tr>
        <td class="supplierlabel" width="25%"><font size="2">Start&nbsp;Date</font></td>
        <td class="supplierlabel"><font size="2">End&nbsp;Date   maybe have a time pick so can get earlies lates</font></td>
    </tr> 
    <tr>
        <td class="supplierinput" nowrap="nowrap">
            <!--Start Date-->        
            <select name="cboStart" id="cboStart" onchange="javascript:HSStartDate();" style="width: 120px">
                <option value="" >Select Date</option>
            </select>
            &nbsp;&nbsp;&nbsp;
            <img id="DtPickerStart"  src="images/cal.gif" width="15" height="15" alt="" 
                onclick="javascript: ShowDatePicker('txtStart','cboStart','hStartDate');" />&nbsp;
                <input type="text" id="txtStart" value= "" onclick="javascript: ShowDatePicker('txtStart','cboStart','hStartDate');"
                readonly="readonly" style="width: 120px"/> 
                           
        
        </td>
        <td class="supplierinput" nowrap="nowrap">
            <!--End Date-->        
            <select name="cboEnd" id="cboEnd" onchange="javascript:HSEndDate();" style="width: 120px">
                <option value="" >Select Date</option>
            </select>  
            &nbsp;&nbsp;&nbsp;            
            <img id="DtPickerEnd"  src="images/cal.gif" width="15" height="15" alt=""  
                onclick="javascript: ShowDatePicker('txtEnd','cboEnd','hEndDate');" />&nbsp;
                <input type="text" id="txtEnd" value= "" onclick="javascript: ShowDatePicker('txtEnd','cboEnd','hEndDate');"
                readonly="readonly" style="width: 120px"/>
                &nbsp;&nbsp;&nbsp;<input type="button" id="EndSame" onclick="javascript:HSEndSame();" value="As start" disabled="disabled"/>
        </td>
    </tr>
    
    <tr>
        <td colspan="2" class="supplierlabel">&nbsp;</td>
    </tr>                              
         
</table>
                                                        
<table style="width: 100%; position: absolute; bottom: 8%; padding-right: 20px; padding-left: 20px;">
        <tr>
            
            <td width="50px">&nbsp;</td>
            <td align="left" width="100px">            <!--this will be a proper submit ?-->
                <input id="btnView" name="btnView" type="button" value="View" onclick="javascript:Relocate('v');" disabled="disabled"/>                               
            </td>
            <td align="left" width="100px">
            <input name="btnReset" id="btnReset" type="reset" value="Reset" onclick="javascript:ResetPage('');"/>
                <!--<input name="btnExport" id="btnExport" type="button" value="Export" onclick="javascript:Relocate('e');" disabled="disabled"/>-->
            </td>
            <td align="left">
                <!--<input name="btnReset" id="btnReset" type="reset" value="Reset" onclick="javascript:ResetPage('1');"/>-->
                <!--<textarea id="txtDetails" cols="30"  rows="5"></textarea>-->
            </td>
            
         </tr>
                     
         <tr>
            <td colspan="4" >
                
                <%If InputType = "text" Then Response.Write "frmName"%>
                <input type="<%=InputType %>" name="frmName" id="frmName" value="frmCustomSelect" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hLogType"%> 
                <input type="<%=InputType %>" name="hLogType" id ="hLogType" value="" <%=InputSize %>/>&nbsp; <%'If ReturnOrderType >= Clng(0) Then Response.Write ReturnOrderType Else Response.Write ""%>
                <%If InputType = "text" Then Response.Write "hReasonId"%>
                <input type="<%=InputType %>" name="hReasonId" id ="hReasonId" value="0" <%=InputSize %>/>&nbsp; 
                <%If InputType = "text" Then Response.Write "hUserId"%>
                <input type="<%=InputType %>" name="hUserId" id ="hUserId" value="0" <%=InputSize %>/>&nbsp;                            
                <%If InputType = "text" Then Response.Write "hShift"%>
                <input type="<%=InputType %>" name="hShift" id ="hShift" value="0" <%=InputSize %>/>
                <%If InputType = "text" Then  Response.Write "hLocation"%>
                <input type="<%=InputType %>" name="hLocation" id ="hLocation" value="0" <%=InputSize%>/>&nbsp;&nbsp;
                
                 <br />
                 <%If InputType = "text" Then Response.Write "hResolvedUserId"%>
                <input type="<%=InputType %>" name="hResolvedUserId" id ="hResolvedUserId" value="0" <%=InputSize %>/> 
                <%If InputType = "text" Then Response.Write "hStartDate"%>
                <input type="<%=InputType %>" name="hStartDate" id ="hStartDate" value="0" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hEndDate"%>
                <input type="<%=InputType %>" name="hEndDate" id ="hEndDate" value="0" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hSource"%>
                <input type="<%=InputType %>" name="hSource" id ="hSource" value="0" <%=InputSize %>/>&nbsp;
               
                
                <%If InputType = "text" Then Response.Write "ErrBox"%>
                <input id="ErrBox" type="<%=InputType%>" value="" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "Locked"%> 
                <input type="<%=InputType %>" name="" id ="Locked" value="<%=Session("Locked")%>" <%=InputSize %>/>&nbsp;
            </td>            
        </tr>          
    </table>    
</body>  
</html>

<%
Private Function GetHsUser(UserId)

    Dim HsUserRs
    Dim LocalName
    LocalName = ""

    Set HsUserRs = Server.CreateObject("ADODB.Recordset")
	HsUserRs.ActiveConnection = Session("ConnMcLogon")
	HsUserRs.Source = "Select UserName From Users Where (Id = " & UserId & ")" 
	HsUserRs.CursorType = Application("adOpenForwardOnly")
	HsUserRs.CursorLocation = Application("adUseClient") 
	HsUserRs.LockType = Application("adLockReadOnly")
	HsUserRs.Open
	Set HsUserRs.ActiveConnection = Nothing

    If HsUserRs.BOF = True Or HsUserRs.EOF = True Then
        LocalName = ""
    Else
        LocalName = HsUserRs("UserName")
    End If
    
    HsUserRs.Close
    Set HsUserRs = Nothing

    GetHsUser = LocalName

End Function
%>

