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
If Session("Locked") = True Then Response.Redirect "SystemLocked.asp"

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim GroupType
Dim GroupTitleExtra
Dim ResolvedDisabled

GroupType = Cint(Request.QueryString("st"))

Select Case GroupType
    Case 1
        GroupTitleExtra = "Accident"
        ResolvedDisabled = " disabled='disabled' "
    Case 2
        GroupTitleExtra = "Near Miss"
        ResolvedDisabled = ""
    Case 3
        GroupTitleExtra = "Unsafe Condition or Damage"
        ResolvedDisabled = ""
End Select    

Dim GroupSql
Dim GroupRs

GroupSql = "SELECT Id, GroupId, Usage, Description, ExcludedLocations" 
GroupSql = GroupSql & " FROM Reasons Where (GroupId In(0, " & GroupType & ")) AND (Active = 1) Order By GroupId Desc, Description"

Set GroupRs = Server.CreateObject("ADODB.Recordset")
GroupRs.ActiveConnection = Session("ConnHSReports") 
GroupRs.Source = GroupSql
GroupRs.CursorType = Application("adOpenForwardOnly")
GroupRs.CursorLocation = Application("adUseClient") 
GroupRs.LockType = Application("adLockReadOnly")
GroupRs.Open
Set GroupRs.ActiveConnection = Nothing

Dim InputType
Dim InputSize
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    If Cint(Session("UserId")) = 1 Or Cint(Session("UserId")) = 4 Then
        InputType = "text"
        InputSize = " size='8' "
    Else
        InputType = "hidden"
        InputSize = ""
    End If
Else
    InputType = "hidden"
    InputSize = ""
End If

Dim DayNumber
DayNumber = Day(Date)
If Cint(DayNumber) < 10 Then DayNumber = "0" & DayNumber

Dim HourNumber
HourNumber = Hour(Now)
If Cint(HourNumber) < 10 Then HourNumber = "0" & HourNumber

Dim MinuteNumber
MinuteNumber = Minute(Now)
If Cint(MinuteNumber) < 10 Then MinuteNumber = "0" & MinuteNumber

Dim LocationSql
Dim LocationRs

LocationSql = "SELECT Id, LocationTypeId, LocationName" 
LocationSql = LocationSql & " FROM Location Where (Active = 1) Order By LocationGroup, LocationTypeId"

Set LocationRs = Server.CreateObject("ADODB.Recordset")
LocationRs.ActiveConnection = Session("ConnHSReports") 
LocationRs.Source = LocationSql
LocationRs.CursorType = Application("adOpenForwardOnly")
LocationRs.CursorLocation = Application("adUseClient") 
LocationRs.LockType = Application("adLockReadOnly")
LocationRs.Open
Set LocationRs.ActiveConnection = Nothing

Dim LocalRiddor
Dim LocalRiddorDays
Dim TodayDateSerial

TodayDateSerial = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))

Dim RiddorRs
Dim RiddorSql

RiddorSql = "Select * FROM Riddor"

Set RiddorRs = Server.CreateObject("ADODB.Recordset")
RiddorRs.ActiveConnection = Session("ConnHSReports") 
RiddorRs.Source = RiddorSql
RiddorRs.CursorType = Application("adOpenForwardOnly")
RiddorRs.CursorLocation = Application("adUseClient") 
RiddorRs.LockType = Application("adLockReadOnly")
RiddorRs.Open

LocalRiddorDays = RiddorRs("Days")


Session("AddFolder") = ""
Session("AddFolder") = Session.SessionID
Session("DeleteFolder") = Session("HsImagePath") & Session.SessionID


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/HSReportsJSFunc.js"></script>
    <script type="text/javascript" src="JsFiles/HSReportsAjaxFunc.js"></script>
    <script type="text/javascript" src="JsFiles/HSReportsDateTimePicker_css.js"></script>    
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:HSAddLogLoad();">
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
                <a href ="javascript:GoBackHSOption();" style="font-size:12px; color: <%=NewBlue%>;">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>
            </td>
             <td height="20px" width="34%" valign="bottom" align="center"> <!-- color #0069AA; Blue-->
                <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" />  
                <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">New&nbsp;<%=GroupTitleExtra%>&nbsp;Log</font>
            </td>
            <td valign="bottom" align="right" width="33%">
                <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;off</a>&nbsp;&nbsp;
            </td> 
        </tr>
        <tr>
            <td height="5px" colspan="3"></td>
        </tr>
        <tr>
            <td colspan="3" nowrap="nowrap" >
                <p id="toptext">
                    <font size="3" style="color: #0069AA">                        
                    &nbsp;&nbsp;Enter&nbsp;details&nbsp;in&nbsp;box&nbsp;below.&nbsp;Select&nbsp;reason&nbsp;&amp;&nbsp;severity.
                    Fields&nbsp;marked&nbsp;</font><font size="3" style="color: red">*</font>
                    <font size="3" style="color: #0069AA">are&nbsp;mandatory.<br />
                    &nbsp;&nbsp;If&nbsp;no&nbsp;date/time&nbsp;is&nbsp;selected&nbsp;the&nbsp;current&nbsp;date/time&nbsp;will&nbsp;be&nbsp;used.
                    You&nbsp;can't&nbsp;select&nbsp;a&nbsp;date/time&nbsp;in&nbsp;the&nbsp;future.
                    </font>
                </p>
            </td>
        </tr>                
    </table>
    
    <br />  
    <form method="post" action="HsUpdateLog.asp"  name="frmAddLog" id="frmAddlog" onsubmit="return HSValidateLog('Add');">
    
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;" >          
        <tr>
            <td align="left" width="60%"  >
                <label style=" color: <%=NewBlue%>">&nbsp;Details<font style="color: red">*</font></label><br />
                <textarea  rows="15" id="txtDetails" name="txtDetails" style="text-align: left; " cols="110" ></textarea>
            </td>            
                      
                     
        </tr>           
    </table>
    
    <br />
    <br />
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;">
 
    <tr>   
            <td valign="middle" width="100px" ><label style=" color: <%=NewBlue%>">&nbsp;Select&nbsp;Reason&nbsp;<font style="color: red">*</font></label></td>                      
            <td valign="middle"  width="200px" >
            <select id='cboReason' name='cboReason' onchange='javascript:HSAddReason("True");'>
                <option value="" >Select&nbsp;Reason</option>
                <%                                                                  
                While Not GroupRs.EOF                    
                    Response.Write ("<option value='" & GroupRs("Id") & "#" & GroupRs("ExcludedLocations") & "'>" & GroupRs("Description") & "</option>") & vbcrlf 
                    GroupRs.MoveNext                
                Wend
                
                GroupRs.Close
                Set GroupRs = Nothing
               
                %>                               
                </select>               
                       
            </td>             
            
            <td valign="middle" width="100px" ><label style=" color: <%=NewBlue%>">&nbsp;Select&nbsp;Severity&nbsp;<font style="color: red">*</font></label></td>         
            <td valign="middle" >
                
                <select id="cboSeverity" name="cboSeverity" onchange="javascript:HSSeverity();" disabled="disabled">
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
            <td valign="middle" ><label style=" color: <%=NewBlue%>">&nbsp;Date&nbsp;&amp;Time&nbsp;of&nbsp;Incident</label></td>
            <td align="left">
            <input name="txtDate" id="txtDate" type="text" value="" readonly="readonly" title="If not selected, current Date/Time will be used"
                    onclick="javascript:NewCssCal('txtDate','ddmmyyyy','arrow',true,'24','','past')" style="cursor:pointer"/> 
                
                <img id="DtPickerDate"  src="images/cal.gif" width="15" height="15" 
                alt="" onclick="javascript:NewCssCal('txtDate','ddmmyyyy','arrow',true,'24','','past')" style="cursor:pointer" />&nbsp;            
            </td>            
            <td align="left">
                <label style=" color: <%=NewBlue%>">&nbsp;Select&nbsp;Location&nbsp;<font style="color: red">*</font></label>    
            </td>            
            <td align="left">
                <select id="cboLocation" name="cboLocation" onchange="javascript:HSAddLocation();" disabled="disabled" style="width: 200px"> <!--disabled="disabled"-->
                    <option value="">Select&nbsp;Location</option>
                    <%
                    While Not LocationRs.EOF                    
                        Response.Write ("<option value='" & LocationRs("Id") & "'>" & LocationRs("LocationName") & "</option>") & vbcrlf 
                        LocationRs.MoveNext                
                    Wend
                    
                    LocationRs.Close
                    Set LocationRs = Nothing
                   
                    %>   
                </select>       
            </td>            
        </tr>
        
        <tr>
            <td height="5px" colspan="4"></td>
        </tr>
        
        <tr>
            <td valign="middle" ><label style=" color: <%=NewBlue%>">&nbsp;Issue&nbsp;Resolved&nbsp;</label></td>
            <td align="left" ><input type="checkbox" id="chkResolved" name="chkResolved" <%=ResolvedDisabled%> /></td>
            <td align="left" ><label style=" color: <%=NewBlue%>">&nbsp;Riddor&nbsp;</label></td>
            <td align="left">
                <select id="cboRiddor" name="cboRiddor" onchange="javascript:HSAddRiddor();" 
                <%If Session("ShowRiddor") = Cbool(False) Then Response.Write " disabled='disabled' " %> >
                    <option value="">Select&nbsp;Riddor</option>
                    <%
                    
                   
                    While Not RiddorRs.EOF                    
                        Response.Write ("<option value='" & RiddorRs("Id") & "#" & RiddorRs("Days") & "'>" & RiddorRs("ShortDescription") & "</option>") & vbcrlf 
                        RiddorRs.MoveNext                
                    Wend
                    
                    RiddorRs.Close
                    Set RiddorRs = Nothing                   
                    %>   
            
                </select>
                &nbsp;<label id="RiddorDays" style="text-align: left; font-size: medium; color: #009933;"></label>
                
            </td>
        </tr>
       
        <tr>
            <td height="30px" colspan="4"></td>
        </tr>
        
        <tr>
	        <td valign="top" height="30" colspan="4" >  <!--disabled="disabled"-->    
		        <input id="btnSubmit" name="btnSubmit" type="submit"  value="Update"  />&nbsp;&nbsp;
		        <input id="btnReset" name="btnReset" onclick="javascript:ResetPage();" type="button" value=" Reset " />&nbsp;&nbsp;
		        
		        <input id="btnAdd" name="btnAdd" onclick="javascript:AddImage('A');" type="button" value="Add Image" title="Any images added will be viewable after record is saved"/>&nbsp;&nbsp;
		        
		        <%
                If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                    If InputType = "text" Then
                        Response.Write ("Don't Save Data&nbsp;&nbsp;")
                        Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                    End If                    
                End If          
                %> 
                <br /><br />    
                <input type="<%=InputType%>" name="frmName" id="frmName" value="frmAddLog" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hSeverity"%>
                <input type="<%=InputType%>" name="hSeverity" id="hSeverity" value="" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hReasonId"%>
                <input type="<%=InputType%>" name="hReasonId" id="hReasonId" value="" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hGroupId"%>
                <input type="<%=InputType%>" name="hGroupId" id="hGroupId" value="<%=GroupType%>" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "Locked"%>
                <input type="<%=InputType %>" name="" id ="Locked" value="<%=Session("Locked")%>" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hDayNumber"%>
                <input type="<%=InputType %>" name="" id ="hDayNumber" value="<%=DayNumber%>" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hHourNumber"%>
                <input type="<%=InputType %>" name="" id ="hHourNumber" value="<%=HourNumber%>" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hMinuteNumber"%>
                <input type="<%=InputType %>" name="" id ="hMinuteNumber" value="<%=MinuteNumber%>" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hLocation"%>
                <input type="<%=InputType %>" name="hLocation" id ="hLocation" value="0" <%=InputSize%>/>&nbsp;
                
                <input type="<%=InputType%>" name="SMTP" id="SMTP" value="<%=Session("Smtp")%>" <%=InputSize %> />&nbsp;
                <input type="<%=InputType%>" name="PC-Name" id="PC-Name" value="<%=Session("PC-Name")%>" <%=InputSize %> />&nbsp;
                <input type="<%=InputType%>" name="SessionID" id="SessionID" value="<%=Session.SessionID%>" />&nbsp;
                
                
                <!--All below for Riddor-->
                <%If InputType = "text" Then  Response.Write "hRiddor"%>
                <!-- This is the level, update sets Riddor to True or False based in level number-->
                <input type="<%=InputType %>" name="hRiddor" id="hRiddor" value="0" <%=InputSize%>/>&nbsp;
                              
                <%If InputType = "text" Then  Response.Write "hRiddorDays"%>
                <input type="<%=InputType %>" name="hRiddorDays" id="hRiddorDays" value="" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hRiddorDateSerial"%>
                <input type="<%=InputType %>" name="hRiddorDateSerial" id="hRiddorDateSerial" value="" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hRiddorSubmitted "%>
                <input type="<%=InputType %>" name="hRiddorSubmitted" id="hRiddorSubmitted" value="" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hTodayDateSerial "%>
                <input type="<%=InputType %>" name="hTodayDateSerial" id="hTodayDateSerial" value="<%=TodayDateSerial%>" <%=InputSize%>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hLogCreatedSerial "%>
                <input type="<%=InputType %>" name="hLogCreatedSerial" id="hLogCreatedSerial" value="<%=TodayDateSerial%>" <%=InputSize%>/>
                                
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