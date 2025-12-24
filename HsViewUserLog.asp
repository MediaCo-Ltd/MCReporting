
<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="CommonFunctions.asp" -->

<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect "SystemLocked.asp"

If Session("UserName") = "" Then Response.Redirect ("Login.asp")
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

Dim LogToGet
LogToGet = Clng(Request.QueryString("id"))

Dim LogDataSql
Dim LogDataRs

LogDataSql = "SELECT  Logs.Id, Logs.UserId, Logs.GroupId, Logs.GroupSelection, Logs.Notes, Logs.CreatedDate, Logs.CreatedDateSerial,"
LogDataSql = LogDataSql & " Logs.Severity, Logs.SelectedDate, Logs.SelectedDateSerial, Logs.Resolved, Logs.LocationId,"
LogDataSql = LogDataSql & " Logs.ResponseNotes, Logs.LastModifiedDate, Logs.LastModifiedDateSerial, Logs.LastModifiedBy,"
LogDataSql = LogDataSql & " Logs.ResolvedDate, Logs.ResolvedDateSerial, Logs.ResolvedBy, Reasons.Description, Reasons.ExcludedLocations,"
LogDataSql = LogDataSql & " Logs.CreatedByName, Logs.Dormant"
LogDataSql = LogDataSql & " FROM Logs INNER JOIN"
LogDataSql = LogDataSql & " Reasons ON Logs.GroupSelection = Reasons.Id" 
LogDataSql = LogDataSql & " Where (Logs.Id = " & LogToGet & ") AND (UserId = " & Session("UserId") & ")"

Set LogDataRs = Server.CreateObject("ADODB.Recordset")
LogDataRs.ActiveConnection = Session("ConnHSReports")
LogDataRs.Source = LogDataSql
LogDataRs.CursorType = Application("adOpenForwardOnly")
LogDataRs.CursorLocation = Application("adUseClient") 
LogDataRs.LockType = Application("adLockReadOnly")
LogDataRs.Open
Set LogDataRs.ActiveConnection = Nothing

If LogDataRs.BOF = True Or LogDataRs.EOF = True Then
    LogDataRs.Close
    Set LogDataRs = Nothing
    Session("OrderOk") = "Record " & LogToGet & " does not exist!"        
    Response.Redirect "SelectOption.asp"
End If

Dim SeverityTxt
Select Case LogDataRs("Severity")
    Case 1 
        SeverityTxt = "Minor"
    Case 2
        SeverityTxt = "Medium"    
    Case Else
        SeverityTxt = "Critical" 
End Select 

Dim LocalTypeId
Dim LocalTypeName 
LocalTypeId = LogDataRs("GroupId")
If LocalTypeId = 1 then
    LocalTypeName = "Accident"
ElseIf LocalTypeId = 2 then
    LocalTypeName = "Incident"
Else
    LocalTypeName = "Unsafe"
End If  

Dim ReturnPage
Dim ReturnHtml
ReturnPage = "HsDisplayUserLogs.asp"
ReturnHtml = "&nbsp;&nbsp;"
ReturnHtml = ReturnHtml & "<a href ='javascript:window.location.replace(""" & ReturnPage & """);' style='color:" & NewBlue & "; font-size:12px;'>"
ReturnHtml = ReturnHtml & "Return&nbsp;to&nbsp;display&nbsp;page</a>"

Dim LocalReasonId
LocalReasonId = LogDataRs("GroupSelection")

Dim LocalResolved
LocalResolved = CBool(LogDataRs("Resolved"))

Dim InputType
Dim InputSize
If Session("PC-Name") = "Home" Then
    InputType = "text"
    InputSize = "size = '4'"
Else
    InputType = "hidden"
    InputSize = ""
End If

Dim LocalReasonNotes
Dim LocalResponseNotes

LocalReasonNotes = Ltrim(LogDataRs("Notes"))

If LogDataRs("ResponseNotes") <> "" Then LocalResponseNotes = Ltrim(LogDataRs("ResponseNotes"))

If LocalReasonNotes <> "" Then
    LocalReasonNotes = Replace(LocalReasonNotes,"<br />",VbCrLf,1,-1,1)
End If

If LocalResponseNotes <> "" Then
    LocalResponseNotes = Replace(LocalResponseNotes,"<br />",VbCrLf,1,-1,1)
Else
    LocalResponseNotes = ""
End If


Dim LoggedByName
LoggedByName = Trim(LogDataRs("CreatedByName"))

Dim ResolvedByName
ResolvedByName = GetMcUser (LogDataRs("ResolvedBy"))

Dim LastModifiedByName
LastModifiedByName = GetMcUser (LogDataRs("LastModifiedBy"))

Dim ResponseRows
Dim ResponseMsg
Dim ResponseHeaderExtra
Dim LockResponseText
Dim LocalHasNotes

If LogDataRs("LastModifiedDateSerial") > 0 Then
    ResponseMsg = "Last update by " & LastModifiedByName & "&nbsp;" & Cdate(LogDataRs("LastModifiedDate"))
    If LocalResponseNotes = "" Then
        ResponseMsg = ""
        LockResponseText = ""
        ResponseHeaderExtra = "*"
        LocalHasNotes = False
    Else
        LocalHasNotes = True
        LockResponseText = " readonly='readonly' "
        ResponseHeaderExtra = ""
    End If
Else
    ResponseMsg = ""
    LockResponseText = ""
    ResponseHeaderExtra = "*"
    LocalHasNotes = False
End If


Dim LocalLocationId
Dim LocationLocked
Dim LocatLocationName

LocalLocationId = LogDataRs("LocationId")
LocationLocked = " disabled='disabled' "


Dim LocationSql
Dim LocationRs
Dim WhereClause
Dim ExcludedList

Dim DataLocked

DataLocked = " disabled='disabled' "
LocationLocked = " disabled='disabled' "
    
LocationSql = "SELECT  Id, LocationName, LocationType, LocationTypeId, LocationGroup"
LocationSql = LocationSql & " FROM Location"
LocationSql = LocationSql & " WHERE (Id = " & LocalLocationId & ")"    

Set LocationRs = Server.CreateObject("ADODB.Recordset")
LocationRs.ActiveConnection = Session("ConnHSReports") 
LocationRs.Source = LocationSql
LocationRs.CursorType = Application("adOpenForwardOnly")
LocationRs.CursorLocation = Application("adUseClient") 
LocationRs.LockType = Application("adLockReadOnly")
LocationRs.Open
Set LocationRs.ActiveConnection = Nothing

LocatLocationName = LocationRs("LocationName")

LocationRs.Close
set LocationRs = Nothing

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
    <script type="text/javascript" src="JsFiles/HSReportsDateTimePicker_css.js"></script>   
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:HSEditLogLoad();">

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
                <%=ReturnHtml%>            
            </td>
             <td height="20px" width="34%" valign="bottom" align="center"> <!-- color #0069AA; Blue-->
                <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" />  
                <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">View&nbsp;<%=LocalTypeName%>&nbsp;Log&nbsp;#<%=LogToGet%></font>
            </td>
            <td valign="bottom" align="right" width="33%">
                <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;off</a>&nbsp;&nbsp;
            </td> 
        </tr>
        <tr>
        <td colspan="3" nowrap="nowrap">
                
                <p id="toptext">&nbsp;</p>
            </td>
        </tr>              
    </table>
    <br />
    <br />  
        
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;" >          
        <tr>
            <td align="left" width="50%">
                <label style=" color: <%=NewBlue%>; font-size: medium;">Created by <%=LoggedByName%>&nbsp;<%=LogDataRs("SelectedDate") %></label>
                <br /> <br />    
            </td>
            <td align="left" width="50%">
                <label style=" color: <%=NewBlue%>; font-size: medium;"><%=ResponseMsg%></label> 
                <br /> <br />    
            </td>
        </tr>
        
        <tr>
            <td align="left" width="50%" >
                <label style=" color: <%=NewBlue%>">Details</label><br />
                <textarea  rows="20" id="txtError" name="txtError" style="text-align: left; " 
                    cols="70" readonly="readonly"><%=Trim(LocalReasonNotes)%></textarea>
            </td>            
            <td align="left" width="50%" <%= DataLocked %>>
                <label style=" color: <%=NewBlue%>">Comments<font style="color: red"><%=ResponseHeaderExtra%></font></label><br />
                <textarea  onkeyup="javascript:EnableSubmit()" rows="20" id="txtResponse" <%=LockResponseText%>
                    name="txtResponse" style="text-align: left; " cols="70" ><%=Trim(LocalResponseNotes)%></textarea>
                
            </td>             
        </tr>           
    </table>
    
    <br />
    <br />
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;"> 
        <tr>   
            <td valign="middle" align="left">
                <label style=" color: <%=NewBlue%>">Reason&nbsp;</label>
                <input type="text" id="txtMachine" value="<%=LogDataRs("Description")%>" readonly="readonly" />
                <label style=" color: <%=NewBlue%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Severity&nbsp;</label>
                <input type="text" id="Severity" value="<%=SeverityTxt%>" readonly="readonly" />
                <label style=" color: <%=NewBlue%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Location&nbsp;</label>
                <input type="text" id="Location" value="<%=LocatLocationName%>" readonly="readonly" />
                <label style=" color: <%=NewBlue%>">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Resolved</label>
                <input type="checkbox" id="chkResolved" name="chkResolved" disabled="disabled"
                <%
                If LogDataRs("Resolved") = CBool(True) Then
                    Response.Write " checked='checked' " 
                    Response.Write " title='Marked resolved on " & Cdate(LogDataRs("ResolvedDateSerial")) & " by " & ResolvedByName & "' "
                 End If 
                %>                 
                 />
            </td>                     
        </tr>      
    </table>
    
    <br />
    <br />
           
    <table style="width: 100%; padding-right: 10px; padding-left: 10px;" >
        <tr>
            <td align="left" style="height: 45px" valign="bottom">&nbsp;&nbsp;
                <a href ="javascript:GoBackHSOption();" style="font-size:12px; color: <%=NewBlue%>;">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>
                <%=ReturnHtml%>            
            </td>
        </tr>
       
        <tr>
	        <td valign="top"  > 	        
		                       
                <br /><br />
                <%If InputType = "text" Then Response.Write "frmName"%>     
                <input type="<%=InputType%>" name="frmName" id="frmName" value="frmEditLog"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hType"%>
                <input type="<%=InputType%>" name="hType" id="hType" value="<%=LocalTypeId%>" <%=InputSize %> />&nbsp;
                <%If InputType = "text" Then Response.Write "hLogId"%>
                <input type="<%=InputType%>" name="hLogId" id="hLogId" value="<%=LogDataRs("Id")%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hSeverity"%> 
                <input type="<%=InputType%>" name="hSeverity" id="hSeverity" value="<%=LogDataRs("Severity")%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hResolved"%>
                <input type="<%=InputType%>" name="hResolved" id="hResolved" value="<%=LogDataRs("Resolved")%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hReasonId"%>
                <input type="<%=InputType%>" name="hReasonId" id="hReasonId" value="<%=LocalReasonId%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hHasNotes"%>
                <input type="<%=InputType%>" name="hHasNotes" id="hHasNotes" value="<%=LocalHasNotes%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hLocation"%>
                <input type="<%=InputType %>" name="hLocation" id ="hLocation" value="<%=LocalLocationId%>" <%=InputSize%>/>
                       
                
                <%
                LogDataRs.Close
                Set LogDataRs = Nothing
                %>
            </td>
        </tr>
        
    </table>  
   
</body>  
</html>

