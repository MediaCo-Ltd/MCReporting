<%@Language="VBScript" Codepage="1252" EnableSessionState=True%>
<%Option Explicit%>
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="CommonFunctions.asp" -->
<%

'on error resume next

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "cache-control","private"
Response.AddHeader "pragma","no-cache"
Response.CacheControl = "no-cache, no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

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

Dim LogDataSql
Dim LogDataRs

LogDataSql = "SELECT Logs.Id As RecordRef, Logs.UserId, Logs.MachineId, Logs.Status, Logs.ErrorDescription, Logs.FaultRepaired,"
LogDataSql = LogDataSql & " Logs.ErrorNotes, Logs.RepairNotes, Logs.RepairDate, Logs.Cost, Logs.RecurringFault, Logs.Dormant,"
LogDataSql = LogDataSql & " Logs.MachineTypeId, Logs.LogDate, Logs.LogDateSerial, Logs.RepairDateSerial, Logs.RepairUserId,"
LogDataSql = LogDataSql & " Logs.FaultGroups As SelectedGroup, Machine.FaultGroups, Machine.MachineName, Logs.CreatedByName, Logs.HasImage"
LogDataSql = LogDataSql & " FROM Logs INNER JOIN"
LogDataSql = LogDataSql & " Machine ON Logs.MachineId = Machine.Id"
LogDataSql = LogDataSql & " Where (Logs.Id = " & LogToGet & ")"

Set LogDataRs = Server.CreateObject("ADODB.Recordset")
LogDataRs.ActiveConnection = Session("ConnMachinefaults") 
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
LocalTypeId = LogDataRs("MachineTypeId")

Dim LocalMachineId
LocalMachineId = LogDataRs("MachineId")

Dim LocalFaultRepaired
LocalFaultRepaired = CBool(LogDataRs("FaultRepaired"))

Dim LoggedByName
LoggedByName = Trim(LogDataRs("CreatedByName"))

Dim DormantMsg
If LogDataRs("Dormant") = Cbool(True) Then
    DormantMsg = "Record is currently marked as dormant. You can still edit it, but it won't show in any reports"
Else
    DormantMsg = ""
End If

Dim HasImages
HasImages = LogDataRs("HasImage")

Dim LocalCost
LocalCost = 0
LocalCost = LogDataRs("Cost")

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

Dim LocalErrorNotes
Dim LocalRepairNotes

LocalErrorNotes = Ltrim(LogDataRs("ErrorNotes"))
LocalRepairNotes = Ltrim(LogDataRs("RepairNotes"))
If LocalErrorNotes <> "" Then
    LocalErrorNotes = Replace(LocalErrorNotes,"<br />",VbCrLf,1,-1,1)
End If

If LocalRepairNotes <> "" Then
    LocalRepairNotes = Replace(LocalRepairNotes,"<br />",VbCrLf,1,-1,1)
End If

'## Check If Locked
Dim LockStatus
LockStatus = ChkRecordLock


Dim ExistingFaultsRS
Dim ExistingFaultsSql
Dim ExistingFaultsList

ExistingFaultsSql = "Select Id, Description From FaultGroups Where (Id In(" & LogDataRs("SelectedGroup") & "))"

Set ExistingFaultsRS = Server.CreateObject("ADODB.Recordset")
ExistingFaultsRS.ActiveConnection = Session("ConnMachinefaults") 
ExistingFaultsRS.Source = ExistingFaultsSql
ExistingFaultsRS.CursorType = Application("adOpenForwardOnly")
ExistingFaultsRS.CursorLocation = Application("adUseClient") 
ExistingFaultsRS.LockType = Application("adLockReadOnly")
ExistingFaultsRS.Open
Set ExistingFaultsRS.ActiveConnection = Nothing

If ExistingFaultsRs.RecordCount = 1 Then
    ExistingFaultsList = ExistingFaultsRS("Description")
Else
    While Not ExistingFaultsRs.EOF
        If ExistingFaultsList = "" Then
            ExistingFaultsList = ExistingFaultsRS("Description")
        Else
            ExistingFaultsList = ExistingFaultsList & ", " & ExistingFaultsRS("Description")
        End If
        ExistingFaultsRS.MoveNext
    Wend
End If

ExistingFaultsRS.Close
Set ExistingFaultsRS = Nothing

Dim ResponseMsg
Dim RepairedByName


If LogDataRs("RepairDateSerial") > 0 Then
    RepairedByName = GetMcUser (LogDataRs("RepairUserId"))
    ResponseMsg = "Last update by " & RepairedByName & "&nbsp;" & Cdate(LogDataRs("RepairDate"))
Else
    ResponseMsg = ""
End If


Dim SendPage
Dim ReturnPage
Dim ReturnHtml

SendPage = Cint(Request.QueryString("sp"))

If SendPage = 0 Then
    ReturnPage = "MfDisplay.asp?st=0"
ElseIf SendPage = 1 Then 
    ReturnPage = "MfDisplay.asp?st=1"
ElseIf SendPage = 2 Then
    ReturnPage = "MfDisplay.asp?st=2"
ElseIf SendPage = 3 Then
    ReturnPage = "MfDisplayUserLogs.asp"    
Else
    ReturnPage = "MfDisplayMachine.asp?id=" & LocalMachineId
End If

If ReturnPage <> "" Then
    ReturnHtml = "&nbsp;&nbsp;"
    ReturnHtml = ReturnHtml & "<a href ='javascript:window.location.replace(""" & ReturnPage & """);' style='color:" & NewBlue & "; font-size:12px;'>"
    ReturnHtml = ReturnHtml & "Return&nbsp;to&nbsp;display&nbsp;page</a>"
Else
    ReturnHtml = ""
End If




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

    <script language="javascript" type="text/javascript">
// <!CDATA[

        function cboFaultGroup_onclick() 
        {
            
        }

// ]]>
    </script>
</head>

<body style="padding: 0px; margin: 0px" onload="javascript:EditLogLoad();">
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
                <%=ReturnHtml%>
            </td>
             <td height="20px" width="34%" valign="bottom" align="center"> <!-- color #0069AA; Blue-->
                <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" /> 
                <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">Edit&nbsp;Log&nbsp;<%=LogToGet%></font>
            </td>
            <td valign="bottom" align="right" width="33%">
                <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;off</a>&nbsp;&nbsp;
            </td> 
        </tr>
        <tr>
            <td colspan="3" nowrap="nowrap">
                <p id="toptext">
                    <font size="3" style="color: #0069AA">                        
                    <!--&nbsp;&nbsp;Subject,&nbsp;-->
                    &nbsp;&nbsp;Fault&nbsp;Details&nbsp;&amp;&nbsp;Machine&nbsp;are&nbsp;not&nbsp;editable.
                    Fields&nbsp;marked&nbsp;</font><font size="3" style="color: red">*</font>
                    <font size="3" style="color: #0069AA">are&nbsp;mandatory.</font>
                </p>
            </td>
        </tr>              
    </table>
    <br />
    <br />  
    <form method="post" action="MfUpdateLog.asp"  name="frmEditLog" id="frmEditLog" onsubmit="return MFValidateLog('Edit');">
    
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;" >          
        <tr>
            <td align="left" width="50%">
                <label style=" color: <%=NewBlue%>; font-size: medium;">Created by <%=LoggedByName%>&nbsp;<%=LogDataRs("LogDate") %></label>
                <br /> <br />    
            </td>
            <td align="left" width="50%">
                <label style=" color: <%=NewBlue%>; font-size: medium;"><%=ResponseMsg%></label>
                <br /> <br />    
            </td>
        </tr>
        
        <tr>
            <td align="left" width="50%" >
                <label style=" color: <%=NewBlue%>">Fault&nbsp;Details</label><br />
                <textarea  rows="15" id="txtError" name="txtError" style="text-align: left; " cols="55" readonly="readonly"><%=Trim(LocalErrorNotes)%></textarea>
            </td>            
            <td align="left" width="50%" >
                <label style=" color: <%=NewBlue%>">Repair&nbsp;Comments&nbsp;<font style="color: red">*</font></label><br />
                <textarea  onkeyup="javascript:EnableSubmit()" rows="15" id="txtRepair" name="txtRepair" style="text-align: left; " cols="55" ><%=Trim(LocalRepairNotes)%></textarea>
            </td>             
        </tr>           
    </table>
    
    <br />
    <br />
    
    <table style="width: 100%; padding-right: 20px; padding-left: 40px;"> 
        <tr>   
            <td valign="middle" width="100px"><label style=" color: <%=NewBlue%>">&nbsp;Machine</label></td>                      
            <td valign="middle"  colspan="1" width="200px">
                <input type="text" id="txtMachine" value="<%=LogDataRs("MachineName")%>" readonly="readonly" />
            </td> 
            <td valign="middle" width="100px"><label style=" color: <%=NewBlue%>">&nbsp;Severity&nbsp;<font style="color: red">*</font></label></td>         
            <td valign="middle" >
                
                <select id="cboSeverity" name="cboSeverity" onchange="javascript:Severity();">
                    <option value="">Select&nbsp;Severity</option>
                    <%
                    Dim Count
                    Dim LocalSeverity(3)
                    LocalSeverity(1) = "Minor"
                    LocalSeverity(2) = "Medium"
                    LocalSeverity(3) = "Critical"
                    
                    For Count = 1 To 3
                        If LogDataRs("Status") = Count Then
                            Response.Write ("<option selected='selected' value='" & Count & "'>" & LocalSeverity(Count) & "</option>") & vbcrlf 
                        Else
                            Response.Write ("<option value='" & Count & "'>" & LocalSeverity(Count) & "</option>") & vbcrlf 
                        End If                    
                    Next
                    Erase LocalSeverity
                    %>      
                </select>                
            </td>             
        </tr>
        
        <tr>
            <td height="5px" colspan="4"></td>
        </tr>
        
        
        <tr>
            <td valign="middle"  ><label style=" color: <%=NewBlue%>">&nbsp;Recurring&nbsp;Fault</label>              
                
            </td>
            
            <td colspan="1" align="left">
                <input type="checkbox" id="chkRecurring" name="chkRecurring" onclick="javascript:EnableSubmit()"
                <%If LogDataRs("RecurringFault") = CBool(True) Then Response.Write " checked='checked' "%> /></td>
            
            
            <td colspan="1" align="left"><label id="GroupLabel" style="color: <%=NewBlue%>">&nbsp;Select&nbsp;Fault&nbsp;Group&nbsp;<font style="color: red">*</font></label></td>
            <td colspan="1" align="left">
                <select id="cboFaultGroup" name="cboFaultGroup" onchange="javascript:FaultSelectEdit();" onclick="return cboFaultGroup_onclick()">
                    <option value="">Select&nbsp;Fault&nbsp;Group</option>             
                </select>
            
            </td>
        </tr>
       
       <tr>
            <td height="5px" colspan="4"></td>
        </tr>
        
        
        <tr>
            <td valign="middle"  ><label style=" color: <%=NewBlue%>">&nbsp;Fault&nbsp;Fixed</label></td>
            <td align="left" colspan="1">
                <input type="checkbox" id="chkFixed" name="chkFixed" onclick="javascript:EnableSubmit()"
                <%If LogDataRs("FaultRepaired") = CBool(True) Then Response.Write " checked='checked' "%> />
            </td>
            <td align="left"><label style=" color: <%=NewBlue%>">&nbsp;Exitsing&nbsp;Selection&nbsp;</label></td>
            <td align="left">
                <label id="OldGrouplist" style="text-align: left; font-size: medium; color: #009933;"><%=ExistingFaultsList%></label>
            </td>        
        </tr>
       
        <tr>
            <td height="5px" colspan="4"></td>
        </tr>
        
        
        <tr>
            <td valign="middle"  ><label style=" color: <%=NewBlue%>">&nbsp;Cost</label></td>
            
            <td align="left">
                <input id="txtCost" name="txtCost" value="<%=LocalCost%>" />
            </td>
            
            <!--CCur(FormatCurrency(LocalCost))-->
            <!--CCur(LocalCost)-->
            
            <!--<td height="25px" align="left"><label style=" color: <%=NewBlue%>">&nbsp;New&nbsp;Selection&nbsp;lspan="2"> <label id="Grouplist" style="text-align: left; font-size: medium; color: #FF0000;"></label></td>-->
            <td height="25px" align="left"><label style=" color: <%=NewBlue%>">&nbsp;New&nbsp;Selection&nbsp;</label></td>
            <td colspan="2"><label id="Grouplist" style="text-align: left; font-size: medium; color: #FF0000;"></label></td>
        </tr>
        
         <tr>
            <td height="30px" colspan="4" style="vertical-align: middle; color: #FF0000; font-size: 16px;"><%=DormantMsg%></td>
        </tr>
        
        <tr>
	        <td valign="top" colspan="4" >    
		        <input id="btnSubmit" name="btnSubmit" type="submit"  value="Update" disabled="disabled" />&nbsp;&nbsp;
		        <input id="btnReset" name="btnReset" onclick="javascript:ResetPage();" type="button" value=" Reset " />&nbsp;&nbsp;
		        <input id="btnAdd" name="btnAdd" onclick="javascript:AddImage('E');" type="button" value="Add Image" />&nbsp;&nbsp;
		        
		        <%
		            If HasImages = True Then
		                Response.Write "<input id='btnImages' name='btnImages' type='button' value=' Images ' onclick='javascript:ShowImages();' />&nbsp;&nbsp;"
		            Else
	                    If ImgChk(FolderName) = True Then
	                        Response.Write "<input id='btnImages' name='btnImages' type='button' value=' Images ' onclick='javascript:ShowImages();' />&nbsp;&nbsp;"
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
                
                <%If InputType = "text" Then Response.Write "hType"%>
                <input type="<%=InputType%>" name="hType" id="hType" value="<%=LocalTypeId%>" <%=InputSize %> />&nbsp;
                <%If InputType = "text" Then Response.Write "hLogId"%>
                <input type="<%=InputType%>" name="hLogId" id="hLogId" value="<%=LogDataRs("RecordRef")%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hSeverity"%> 
                <input type="<%=InputType%>" name="hSeverity" id="hSeverity" value="<%=LogDataRs("Status")%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hRepaired"%>
                <input type="<%=InputType%>" name="hRepaired" id="hRepaired" value="<%=LocalFaultRepaired%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hLockStatus"%>
                <input type="<%=InputType%>" name="hLockStatus" id="hLockStatus" value= "<%=LockStatus%>"  <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "LockedByName"%> 
                <input type="<%=InputType%>" name="LockedByName" id="LockedByName" value= "<%=Session("LockedByName")%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "Locked"%> 
                <input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then  Response.Write "hGroups"%>
                <input type="<%=InputType%>" name="hGroups" id="hGroups" value="<%=LogDataRs("FaultGroups")%>" <%=InputSize %>/>&nbsp;
                <%If InputType = "text" Then Response.Write "hGroupId"%>
                <input type="<%=InputType%>" name="hGroupId" id="hGroupId" value="" <%=InputSize %>/>
                <%If InputType = "text" Then Response.Write "hSelectId"%>
                <input type="<%=InputType%>" name="hSelectId" id="hSelectId" value="" <%=InputSize %>/>
                
                <%If InputType = "text" Then Response.Write "frmName"%>     
                <input type="<%=InputType%>" name="frmName" id="frmName" value="frmEditLog"  />
                       
                
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

'##################################################################

Private Function ChkRecordLock()

    '## Set locked status, so that other users can't open 

    Dim RecordLockSql
    Dim RecordLockRs
    Dim LocalRecordLock
    LocalRecordLock = False
    
    Set RecordLockRs = Server.CreateObject("ADODB.Recordset")
    RecordLockRs.ActiveConnection = Session("ConnMcLogon")
    RecordLockRs.Source = "Select * From RecordLocks"  
    RecordLockRs.Source = RecordLockRs.Source & " Where (RecordId = " & LogToGet & ") AND (SystemName = 'MF')"
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
        RecordLockRs("SystemName") = "MF"
        Session("LockedByName") = Session("UserName")
        Session("LockedId") = LogToGet
        'Session("LockedSid") = Session.SessionID
        'Session("LockedbyId") = Session("UserId")
        RecordLockRs.Update
        RecordLockRs.MoveLast 
        
    Else
        '## Already locked, do nothing
        If RecordLockRs("LockedById") = Session("UserId") Then
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
    Path = "C:\Web Sites\MC Reporting\FaultImages\" & Cstr(Id)
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
    UpdateImageRs.ActiveConnection = Session("ConnMachinefaults") 
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