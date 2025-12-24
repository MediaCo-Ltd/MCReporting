
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

Dim FolderName

If LogToGet < 10 Then
    FolderName = "0" & Cstr(LogToGet)
Else
    FolderName = Cstr(LogToGet)
End If

Session("EditFolder") = FolderName

Dim SendPage
Dim ReturnPage
Dim ReturnHtml

SendPage = Cint(Request.QueryString("sp"))

If SendPage = 1 Then
    ReturnPage = "HsDisplay.asp?st=1" 
ElseIf SendPage = 2 Then 
    ReturnPage = "HsDisplay.asp?st=2" 
ElseIf SendPage = 3 Then
    ReturnPage = "HsDisplay.asp?st=3" 
'ElseIf SendPage = 5 Then
'    ReturnPage = "HsDisplayRiddor.asp" 
Else
    ReturnPage = ""
End If
                

Dim LogDataSql
Dim LogDataRs

LogDataSql = "SELECT  Logs.Id, Logs.UserId, Logs.GroupId, Logs.GroupSelection, Logs.Notes, Logs.CreatedDate, Logs.CreatedDateSerial,"
LogDataSql = LogDataSql & " Logs.Severity, Logs.SelectedDate, Logs.SelectedDateSerial, Logs.Resolved, Logs.LocationId,"
LogDataSql = LogDataSql & " Logs.ResponseNotes, Logs.LastModifiedDate, Logs.LastModifiedDateSerial, Logs.LastModifiedBy,"
LogDataSql = LogDataSql & " Logs.ResolvedDate, Logs.ResolvedDateSerial, Logs.ResolvedBy, Reasons.Description, Reasons.ExcludedLocations,"
LogDataSql = LogDataSql & " Logs.CreatedByName, Logs.Dormant, Logs.Resolved, Logs.RiddorDays, Logs.RiddorDateSerial, Logs.RiddorSubmitted,"
LogDataSql = LogDataSql & " Logs.RiddorLevel, Logs.RiddorSubmitedDate, Logs.Riddor, Logs.HasImage"

LogDataSql = LogDataSql & " FROM Logs INNER JOIN"
LogDataSql = LogDataSql & " Reasons ON Logs.GroupSelection = Reasons.Id" 
LogDataSql = LogDataSql & " Where (Logs.Id = " & LogToGet & ")"


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


Dim LocalTypeId
Dim LocalTypeName 
LocalTypeId = LogDataRs("GroupId")
If LocalTypeId = 1 then
    LocalTypeName = "Accident"
ElseIf LocalTypeId = 2 then
    LocalTypeName = "Near Miss"
Else
    LocalTypeName = "Unsafe Condition or Damage"
End If    

Dim LocalReasonId
LocalReasonId = LogDataRs("GroupSelection")

Dim LocalResolved
LocalResolved = CBool(LogDataRs("Resolved"))

Dim HasImages
HasImages = LogDataRs("HasImage")


Dim InputType
Dim InputSize
If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
    If Session("UserId") = 4 Then
        InputType = "text"
        InputSize = "size = '4'"
    Else
        InputType = "hidden"
        InputSize = ""
    End If
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


'## Check If Locked
Dim LockStatus
LockStatus = ChkRecordLock


Dim LoggedByName
LoggedByName = Trim(LogDataRs("CreatedByName"))

Dim ResolvedByName
ResolvedByName = GetMcUser (LogDataRs("ResolvedBy"))

Dim LastModifiedByName
LastModifiedByName = GetMcUser (LogDataRs("LastModifiedBy"))

Dim DormantMsg
If LogDataRs("Dormant") = Cbool(True) Then
    DormantMsg = "Record is currently marked as dormant. You can still edit it, but it won't show in any reports"
Else
    DormantMsg = ""
End If


Dim ResponseRows
Dim ResponseMsg
Dim ResponseHeaderExtra
Dim LockResponseText
Dim NewResponseHtml
Dim LocalHasNotes

Dim DaysLeft
Dim LocalRiddor
Dim LocalRiddorDays
Dim LocalRiddorDateSerial
Dim LocalRiddorSubmitted
Dim TodayDateSerial
Dim LogCreatedSerial

LogCreatedSerial = LogDataRs("CreatedDateSerial")
TodayDateSerial = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))

LocalRiddor = LogDataRs("RiddorLevel")
LocalRiddorDays = LogDataRs("RiddorDays")
LocalRiddorDateSerial = LogDataRs("RiddorDateSerial")
LocalRiddorSubmitted = LogDataRs("RiddorSubmitted")

Dim RiddorDisabled
Dim RiddorDisabledText

If TodayDateSerial - LogCreatedSerial > 15 Then
    RiddorDisabled = Cbool(True)
    RiddorDisabledText = "Riddor submission date has been exceeded"
ElseIf Session("ShowRiddor") = Cbool(False) Then
    RiddorDisabled = Cbool(True)
    RiddorDisabledText = "Riddor submission date has been exceeded" 
Else
    RiddorDisabled = Cbool(False)
    RiddorDisabledText = ""
End If

DaysLeft = TodayDateSerial - LogCreatedSerial

If LocalRiddor > 0 And LocalRiddorSubmitted = False Then    
    DaysLeft = LocalRiddorDays - DaysLeft
    If DaysLeft > 0 Then RiddorDisabledText = " You have " & DaysLeft & " Days to submit"
End If

If LocalRiddorSubmitted = Cbool(True) Then
    RiddorDisabled = Cbool(True)
    RiddorDisabledText = "Riddor has been submitted"    
End If    

If LogDataRs("LastModifiedDateSerial") > 0 Then
    ResponseMsg = "Last update by " & LastModifiedByName & "&nbsp;" & Cdate(LogDataRs("LastModifiedDate"))
    If LocalResponseNotes = "" Then
        ResponseMsg = ""
        LockResponseText = ""
        NewResponseHtml = ""
        ResponseRows = 15
        ResponseHeaderExtra = "*"
        LocalHasNotes = False
    Else
        LocalHasNotes = True
        LockResponseText = " readonly='readonly' "
        ResponseRows = 8
        ResponseHeaderExtra = ""
        NewResponseHtml = "<label style=' color: " & NewBlue & ";'>New&nbsp;Comments</label><br />"
        NewResponseHtml = NewResponseHtml & " <textarea  onkeyup='javascript:EnableSubmit()' rows='8' id='txtResponseNew' "
        NewResponseHtml = NewResponseHtml & " name='txtResponseNew' style='text-align: left; ' cols='70' ></textarea>"
    End If
Else
    ResponseMsg = ""
    LockResponseText = ""
    NewResponseHtml = ""
    ResponseRows = 15
    ResponseHeaderExtra = "*"
    LocalHasNotes = False
End If


Dim LocalLocationId
Dim LocationLocked

LocalLocationId = LogDataRs("LocationId")

If LogDataRs("LocationId") = 0 Then 
    LocationLocked = ""
Else
    LocationLocked = " disabled='disabled' "
End If

Dim LocationSql
Dim LocationRs
Dim WhereClause
Dim ExcludedList

'LocationSql = "SELECT Id, LocationTypeId, LocationName" 
'LocationSql = LocationSql & " FROM Area Where (Active = 1) Order By LocationGroup, LocationTypeId"

ExcludedList = LogDataRs("ExcludedLocations")

If ExcludedList = "" Then
    '## Show all
    WhereClause = " WHERE (Active = 1)"
Else
    '## Only show allowwd locations
    WhereClause = " WHERE (Active = 1) AND (NOT (Id IN (" & ExcludedList & ")))"
End If
    
LocationSql = "SELECT  Id, LocationName, LocationType, LocationTypeId, LocationGroup"
LocationSql = LocationSql & " FROM Location"
LocationSql = LocationSql & WhereClause    
LocationSql = LocationSql & " ORDER BY LocationGroup, LocationTypeId"

Set LocationRs = Server.CreateObject("ADODB.Recordset")
LocationRs.ActiveConnection = Session("ConnHSReports") 
LocationRs.Source = LocationSql
LocationRs.CursorType = Application("adOpenForwardOnly")
LocationRs.CursorLocation = Application("adUseClient") 
LocationRs.LockType = Application("adLockReadOnly")
LocationRs.Open
Set LocationRs.ActiveConnection = Nothing


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


Dim ReturnExtra
If LocalResolved = Cbool(True) Then 
    ReturnPage = ReturnPage & "&rf=1"
Else
    ReturnPage = ReturnPage & "&rf=0"
End If

If SendPage = 5 Then ReturnPage = "HsDisplayRiddor.asp"

If ReturnPage <> "" Then
    ReturnHtml = "&nbsp;&nbsp;"
    ReturnHtml = ReturnHtml & "<a href ='javascript:window.location.replace(""" & ReturnPage & """);' style='color:" & NewBlue & "; font-size:12px;'>"
    ReturnHtml = ReturnHtml & "Return&nbsp;to&nbsp;display&nbsp;page</a>"
Else
    ReturnHtml = ""
End If

If SendPage = 0 Then ReturnHtml = ""

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
                <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">Edit&nbsp;<%=LocalTypeName%>&nbsp;Log&nbsp;#<%=LogToGet%></font>
            </td>
            <td valign="bottom" align="right" width="33%">
                <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;off</a>&nbsp;&nbsp;
            </td> 
        </tr>
        <tr>
        <td colspan="3" nowrap="nowrap">
                <br /><br />
                <p id="toptext">
                    <font size="3" style="color: #0069AA">                        
                    &nbsp;&nbsp;Details&nbsp;&amp;&nbsp;Reason&nbsp;are&nbsp;not&nbsp;editable.
                    Fields&nbsp;marked&nbsp;</font><font size="3" style="color: red">*</font>
                    <font size="3" style="color: #0069AA">are&nbsp;mandatory.</font><br />
                    <font size="3" style="color: #0069AA">&nbsp;&nbsp;If&nbsp;record&nbsp;is&nbsp;ticked&nbsp;as&nbsp;resolved&nbsp;you&nbsp;can't&nbsp;untick,&nbsp;but&nbsp;you&nbsp;can&nbsp;add&nbsp;more&nbsp;notes</font>
                </p>
            </td>
        </tr>              
    </table>
    <br />
    <br />  
    <form method="post" action="HsUpdateLog.asp"  name="frmEditLog" id="frmEditLog" onsubmit="return HSValidateLog('Edit');"><!--onsubmit="return ValidateLog('Edit');"-->
    
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
                <textarea  rows="15" id="txtError" name="txtError" style="text-align: left; " 
                    cols="70" readonly="readonly"><%=Trim(LocalReasonNotes)%></textarea>
            </td>            
            <td align="left" width="50%" >
                <label style=" color: <%=NewBlue%>">Comments<font style="color: red"><%=ResponseHeaderExtra%></font></label><br />
                <textarea  onkeyup="javascript:EnableSubmit()" rows="<%=ResponseRows%>" id="txtResponse" <%=LockResponseText%>
                    name="txtResponse" style="text-align: left; " cols="70" ><%=Trim(LocalResponseNotes)%></textarea>
                <br />
                <%=NewResponseHtml%>
            </td>             
        </tr>           
    </table>
    
    <br />
    <br />
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;"> 
        <tr>   
            <td valign="middle" width="10%" align="right"><label style=" color: <%=NewBlue%>">&nbsp;Reason</label></td>                      
            <td valign="middle" width="15%" align="left">
                <input type="text" id="txtMachine" value="<%=LogDataRs("Description")%>" readonly="readonly" />
            </td> 
            <td valign="middle" width="10%" align="right"><label style=" color: <%=NewBlue%>">&nbsp;Severity&nbsp;<font style="color: red">*</font></label></td>         
            <td valign="middle" width="15%" align="left">
                
                <select id="cboSeverity" name="cboSeverity" onchange="javascript:HSSeverity();">
                    <option value="">Select&nbsp;Severity</option>
                    <%
                    Dim Count
                    Dim LocalSeverity(3)
                    LocalSeverity(1) = "Minor"
                    LocalSeverity(2) = "Medium"
                    LocalSeverity(3) = "Critical"
                    
                    For Count = 1 To 3
                        If LogDataRs("Severity") = Count Then
                            Response.Write ("<option selected='selected' value='" & Count & "'>" & LocalSeverity(Count) & "</option>") & vbcrlf 
                        Else
                            Response.Write ("<option value='" & Count & "'>" & LocalSeverity(Count) & "</option>") & vbcrlf 
                        End If                    
                    Next
                    Erase LocalSeverity
                    %>      
                </select>                             
                </td>             
        
            <td valign="middle" width="10%" align="right"><label style=" color: <%=NewBlue%>">&nbsp;Location&nbsp;</label></td>
            

            <td valign="middle" width="15%" align="left" > <!--onchange="javascript:AddLocation();"  disabled="disabled"-->
            <select id="cboLocation" name="cboLocation"  onchange="javascript:HSAddLocation();"  style="width: 200px" <%=LocationLocked%>>
                    <option value="">Select&nbsp;Location</option>
                    <%
                    While Not LocationRs.EOF 
                        If LocationRs("Id") = LocalLocationId Then
                            Response.Write ("<option selected='selected' value='" & LocationRs("Id") & "'>" & LocationRs("LocationName") & "</option>") & vbcrlf
                        Else                  
                            Response.Write ("<option value='" & LocationRs("Id") & "'>" & LocationRs("LocationName") & "</option>") & vbcrlf
                        End If 
                        LocationRs.MoveNext                
                    Wend
                    
                    LocationRs.Close
                    Set LocationRs = Nothing
                   
                    %>   
                </select>
            
            </td>
        
            <td valign="middle" width="10%" align="right" ><label style=" color: <%=NewBlue%>">&nbsp;Resolved</label></td>
            <td valign="middle" width="15%" align="left" >
                <input type="checkbox" id="chkResolved" name="chkResolved" onclick="javascript:EnableSubmit()"
                <%If LogDataRs("Resolved") = CBool(True) Then Response.Write " checked='checked' "%> 
                <%If LogDataRs("Resolved") = CBool(True) Then Response.Write " disabled='disabled' "%>
                <%
                If LogDataRs("Resolved") = CBool(True) Then 
                    Response.Write " title='Marked resolved on " & Cdate(LogDataRs("ResolvedDateSerial")) & " by " & ResolvedByName & "' "
                End If
                %> 
                 />
            </td>            

        </tr>
        
        <tr>
            <td  colspan="8" >&nbsp;</td>
        </tr>
        
        <tr>
        <td valign="middle" width="10%" align="right" ><label style=" color: <%=NewBlue%>">&nbsp;Riddor</label></td>
            <td valign="middle"  align="left"  colspan="7">
                <select id="cboRiddor" name="cboRiddor" onchange="javascript:HSAddRiddor();" 
                <%If RiddorDisabled = Cbool(True) Then Response.Write " disabled='disabled' " %> >
                    <option value="0">Select&nbsp;Riddor</option>
                    <%                    
                    While Not RiddorRs.EOF 
                        If RiddorRs("id") = LocalRiddor Then
                            Response.Write ("<option selected='selected' value='" & RiddorRs("Id") & "#" & RiddorRs("Days") & "'>" & RiddorRs("ShortDescription") & "</option>") & vbcrlf 
                        Else                  
                            Response.Write ("<option value='" & RiddorRs("Id")  & "#" & RiddorRs("Days") & "'>" & RiddorRs("ShortDescription") & "</option>") & vbcrlf 
                        End If
                        
                        RiddorRs.MoveNext                
                     Wend
                    
                    RiddorRs.Close
                    Set RiddorRs = Nothing                   
                    %>             
                </select>
                &nbsp;<label id="RiddorDays" style="text-align: left; font-size: medium; color: #009933;"><%=RiddorDisabledText%></label>                 
            </td>        
        </tr>
       
        <tr>
            <td height="45px" colspan="8" style="vertical-align: middle; color: #FF0000; font-size: 16px;"><%=DormantMsg%></td>
        </tr>
        
       
        <tr>
	        <td valign="top" colspan="8" >    
		        <input id="btnSubmit" name="btnSubmit" type="submit"  value="Update" disabled="disabled" />&nbsp;&nbsp;
		        <input id="btnReset" name="btnReset" onclick="javascript:ResetPage();" type="button" value=" Reset " />&nbsp;&nbsp;
		        <input id="btnAdd" name="btnAdd" onclick="javascript:AddImage('E');" type="button" value="Add Image" />&nbsp;&nbsp;
		       
		        <%
		            If HasImages = True Then
		                Response.Write "<input id='btnImages' name='btnImages' type='button' value=' Images ' onclick='javascript:ShowHsImages();' />&nbsp;&nbsp;"
		            Else
	                    If ImgChk(FolderName) = True Then
	                        Response.Write "<input id='btnImages' name='btnImages' type='button' value=' Images ' onclick='javascript:ShowHsImages();' />&nbsp;&nbsp;"
	                        '## Update Image status
	                        UpdateImageStatus 1, LogToGet	                        
	                    Else
	                        Response.Write "&nbsp;"
	                        UpdateImageStatus 0, LogToGet
	                    End If
	                End If 
		        %>
		       
		       
		       
		        <%
                If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                    Response.Write ("Don't Save Data&nbsp;&nbsp;")
                    Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                End If          
                %>
                
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
                <%If InputType = "text" Then Response.Write "hLockStatus"%>
                <input type="<%=InputType%>" name="hLockStatus" id="hLockStatus" value= "<%=LockStatus%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "LockedByName"%> 
                <input type="<%=InputType%>" name="LockedByName" id="LockedByName" value= "<%=Session("LockedByName")%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "Locked"%> 
                <input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hReasonId"%>
                <input type="<%=InputType%>" name="hReasonId" id="hReasonId" value="<%=LocalReasonId%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hHasNotes"%>
                <input type="<%=InputType%>" name="hHasNotes" id="hHasNotes" value="<%=LocalHasNotes%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hLocation"%>
                <input type="<%=InputType %>" name="hLocation" id ="hLocation" value="<%=LocalLocationId%>" <%=InputSize%>/>&nbsp;
                
                <!--All below for Riddor-->
                <%If InputType = "text" Then  Response.Write "hRiddor"%>
                <!-- This is the level, update sets Riddor to True or False based in level number-->
                
                <input type="<%=InputType %>" name="hRiddor" id="hRiddor" value="<%=LocalRiddor%>" <%=InputSize%>/>&nbsp;                
                <%If InputType = "text" Then  Response.Write "hRiddorDays"%>
                <input type="<%=InputType %>" name="hRiddorDays" id="hRiddorDays" value="<%=LocalRiddorDays%>" <%=InputSize%>/>&nbsp;                
                <%If InputType = "text" Then  Response.Write "hRiddorDateSerial"%>
                <input type="<%=InputType %>" name="hRiddorDateSerial" id="hRiddorDateSerial" value="<%=LocalRiddorDateSerial%>" <%=InputSize%>/>&nbsp;                
                <%If InputType = "text" Then  Response.Write "hRiddorSubmitted"%>
                <input type="<%=InputType %>" name="hRiddorSubmitted" id="hRiddorSubmitted" value="<%=LocalRiddorSubmitted%>" <%=InputSize%>/>&nbsp;                                
                <%If InputType = "text" Then  Response.Write "hTodayDateSerial"%>
                <input type="<%=InputType %>" name="hTodayDateSerial" id="hTodayDateSerial" value="<%=TodayDateSerial%>" <%=InputSize%>/>&nbsp;                
                <%If InputType = "text" Then  Response.Write "hLogCreatedSerial"%>
                <input type="<%=InputType %>" name="hLogCreatedSerial" id="hLogCreatedSerial" value="<%=LogCreatedSerial%>" <%=InputSize%>/>&nbsp;                
                <%If InputType = "text" Then  Response.Write "hDaysLeft"%>
                <input type="<%=InputType %>" name="hDaysLeft" id="hDaysLeft" value="<%=DaysLeft%>" <%=InputSize%>/>
                  
                
                <%
                LogDataRs.Close
                Set LogDataRs = Nothing
                %>
            </td>
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

Private Function ChkRecordLock()

    '## Set locked status, so that other users can't open 

    Dim RecordLockSql
    Dim RecordLockRs
    Dim LocalRecordLock
    LocalRecordLock = False
    
    Set RecordLockRs = Server.CreateObject("ADODB.Recordset")
    RecordLockRs.ActiveConnection = Session("ConnMcLogon")
    RecordLockRs.Source = "Select * From RecordLocks"  
    RecordLockRs.Source = RecordLockRs.Source & " Where (RecordId = " & LogToGet & ") AND (SystemName = 'HS')"
    RecordLockRs.CursorType = Application("adOpenStatic")
    RecordLockRs.CursorLocation = Application("adUseClient")
    RecordLockRs.LockType = Application("adLockOptimistic")
    RecordLockRs.Open
    
    If RecordLockRs.BOF = True Or RecordLockRs.EOF = True Then
        '## No match add new
        RecordLockRs.AddNew
        RecordLockRs("RecordId") = LogToGet
        RecordLockRs("LockedbyId") = Session("UserId")
        RecordLockRs("LockedByName") = Session("UserName")
        RecordLockRs("SystemName") = "HS"
        Session("LockedByName") = Session("UserName")
        Session("LockedId") = LogToGet
        
        RecordLockRs.Update
        RecordLockRs.MoveLast        
    Else
        '## Already locked, do nothing
        If RecordLockRs("LockedById") = Session("UserId")  Then  'And RecordLockRs("SystemName") = "HS"
            '## Match found but was just a reset of page by same user
            LocalRecordLock = False
        Else
            '## Match found so status is true
            LocalRecordLock = True
            Session("LockedByName") = RecordLockRs("LockedByName")  
        End If
    End If
    
    Set RecordLockRs.ActiveConnection = Nothing
    RecordLockRs.Close
    Set RecordLockRs = Nothing
    
    ChkRecordLock = LocalRecordLock

End Function


'#######################################################

Function ImgChk(Id)
    
    '## Check if Record has images
    Dim Path
    Path = "C:\Web Sites\MC Reporting\HsImages\" & Cstr(Id)
    Dim ReturnValue
    ReturnValue = False    

    Dim objFSO 
    'Dim objFile
    'Dim Folder
    Dim objFileItem
    Dim objFolder
    Dim objFolderContents
    Dim TotalPics
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If objFSO.FolderExists(Path) = False Then
        ReturnValue = False
        Set objFSO = Nothing
    Else
        Set objFolder = objFSO.GetFolder(Path)
        Set objFolderContents = objFolder.Files      
        
        TotalPics = 0
        For Each objFileItem in objFolderContents
            If Ucase(Right(objFileItem.Name,4))=".GIF" OR Ucase(Right(objFileItem.Name,4))=".JPG" OR Ucase(Right(objFileItem.Name,4))=".PNG" Then
                TotalPics = TotalPics + 1
            End if
        Next
        
        If TotalPics > 0 Then            
            ReturnValue = True
        Else
            ReturnValue = False
        End If
        
        Set objFSO = Nothing
        Set objFolder = Nothing
        Set objFolderContents = Nothing    
    End If

    ImgChk = ReturnValue

End Function

'####################################################################

Sub UpdateImageStatus(Status,RecId)

    '## Updates log on new records

    Dim UpdateImageRs
       
    Set UpdateImageRs = Server.CreateObject("ADODB.Recordset")
    UpdateImageRs.ActiveConnection = Session("ConnHSReports") 
    UpdateImageRs.Source = "Select Id, HasImage From Logs Where (Id = " & RecId & ")"
    
    UpdateImageRs.CursorType = Application("adOpenStatic")
    UpdateImageRs.CursorLocation = Application("adUseClient")
    UpdateImageRs.LockType = Application("adLockOptimistic")
    UpdateImageRs.Open

    UpdateImageRs("HasImage") = Status
    
    UpdateImageRs.Update
    
    UpdateImageRs.Close
    Set UpdateImageRs = Nothing

End Sub
%>