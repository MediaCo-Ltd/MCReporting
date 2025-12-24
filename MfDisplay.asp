<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="MfGetData.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<% 
'## On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim RepairStatus

RepairStatus = Cint(Request.QueryString("st"))

'## If Cookie already set don't bother with check

Dim OkStatus 
Dim ReturnAddress
Dim Digital
Dim NotesTitleText
OkStatus = 0
Dim EditId
EditId = ""
Dim NoDisplay
Dim FlushCount
Dim Msg
Dim StatusBg
Dim StatusText
Dim HeaderText
Dim LoggedByName

Session("AddFolder") = ""

If RepairStatus = 0 Then
    HeaderText = "All Pending Records"
ElseIf RepairStatus = 1 Then
    HeaderText = "All Complete Records"
Else
    HeaderText = "All Records"
End If

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
    <title>Machine Faults</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/W-S48.png" type="image/x-icon" />
    <link href="CSS/MachineFaultCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/MachineFaultExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/MachineFaultJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" id="Display" onload="javascript:ViewRecordsLoad();">

<table style="width: 100%;padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100" colspan="3" >
            <img align="left" alt="mediaco logo" src="Images/mediaco_logo.jpg" width="160" />
        </td>
    </tr>
    <tr>
	    <td height="8" valign="top" colspan="3" >
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	    </td>
    </tr>    
	
    <tr>
        <td align="left" style="height: 20px" valign="bottom" width="33%" >
            &nbsp;&nbsp;<a href ="javascript:GoBackOption();" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>            
        </td>
        <td align="center" valign="bottom" width="34%">
            <img align="top" alt="mediaco logo" src="Images/W-S48.png" style="width: 20px; height: 20px;" />
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >&nbsp;<%=HeaderText%></font>                     
        </td>
        <td align="right" width="33%">
            <a href ="javascript:LogOff()" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;Out</a>&nbsp;&nbsp;    
        </td>
    </tr>
   
    <tr>     
        <td align="left" valign="middle" colspan="3">        
            <%
            If Session("AdminUser") = Cbool(True) Or Session("EditMF") = Cbool(True) Then
                Response.Write "<br />&nbsp;To&nbsp;edit&nbsp;a&nbsp;record&nbsp;click&nbsp;anywhere&nbsp;in&nbsp;the&nbsp;Record&nbsp;Id&nbsp;area" 
            Else 
                Response.Write "&nbsp;"                            
            End If
            %>   
        </td>
    </tr>
</table>


<%

Call OpenMachine(RepairStatus,"0")    

While Not MachineRs.Eof

    Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;" & MachineRs("MachineName") & " </h4>"    
    Call OpenLogs (MachineRs("Id"),RepairStatus,"0")    
    '## If no records don't draw tables
    If LogRs.Bof = False or LogRs.EOF = False Then
        Response.Write (Msg)
    %>    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">    
	   <tr >
	        <th class="styleTHleft" width="5%" >Record Id</th>
	        <th class="styleTHstd" width="5%">Log Date</th>	        
	        <th class="styleTHstd" width="5%">Severity</th>
	        <th class="styleTHstd" width="8%">Logged By</th>
	        <th class="styleTHstd" width="4%">Images</th>
	        <th class="styleTHstd" width="20%">Reasons</th>
	        <th class="styleTHstd" width="5%">Repair Status</th>
	        <th class="styleTHstd" width="5%">Repair Date</th>
	        <th class="styleTHstd" width="20%">Repair Notes</th>
	        <th class="styleTHstd" width="5%">Cost</th>
	        <th class="styleTHstd" width="5%">Recurring Fault</th>
	        
        </tr>    
                
        <%
        While Not LogRs.EOF
        
        If Session("AdminUser") = Cbool(True) Or Session("EditMF") = Cbool(True) Then
            EditId = LogRs("Id")
        Else
            EditId = ""
        End If 
        %>        
        <tr>
            <td title="" id="<%=EditId%>" style="color: #0069AA" class="styleTDleft"
                <%If Session("AdminUser") = Cbool(True) Or Session("EditMF") = Cbool(True) Then Response.Write " onclick='javascript:MFLoadEditOrder(" & EditId & ");'" %> ><%=LogRs("Id")%></td>
	        <td  class="styleTDstd"><%=CDate(LogRs("LogDateSerial"))%></td>	    
	        <%
	            If LogRs("Status") = 1 Then
	                StatusBg = "#00BB00"
	                StatusText = "Minor Fault"
	            ElseIf LogRs("Status") = 2 Then
	                StatusBg = "#FFA500"  'lighter "#FFBC55" 
	                StatusText = "Medium Fault"
	            Else
	                StatusBg = "#FF0000"
	                StatusText = "Critical Fault"
	            End If
	            
	            LoggedByName = Trim(LogRs("CreatedByName"))	            
	            NotesTitleText = Replace(LogRs("ErrorNotes"),"<br />",VbCrLf,1,-1,1)
	        %>
	        
	        <td title="<%=StatusText%>" class="styleTDstd" style="background-color:<%=StatusBg%>;"></td>
	        <td  class="styleTDstd"><%=LoggedByName%></td>
	        
	        <td  class="styleTDstd">       
	        <%
	        If LogRs("HasImage") = Cbool(True) Then
	        	Response.Write "<img align='middle' alt='' title='Click to view images' src='Images/checkmark.ico' width='15' height='auto' onclick='javascript:LoadImage(" & EditId & ");' />"
   	        	'Response.Write "<img align='middle' alt='' title='Click to view images' src='Images/camera.ico' width='20' height='auto' onclick='javascript:LoadImage(" & EditId & ");' />"
	        Else
	            Response.Write "&nbsp;"
	        End If	        
	        %>	        
	        </td>
	        
	        <!--Response.Write FaultGroupDescription (LogRs("FaultGroups"))-->
	        <td  class="styleTDstd" style="text-align: left; padding-left: 5px;" title="<%=NotesTitleText%>">
	            <%=FaultGroupDescription (LogRs("FaultGroups"))%></td>
	        <td  class="styleTDstd">
	        <%
	        If LogRs("FaultRepaired") = Cbool(True) Then
	            Response.Write "<img align='middle' alt='' title='Repaired' src='Images/checkmark.ico' width='15' height='15'  />"
	        Else
	            Response.Write "<img align='middle' alt='' title='Pending' src='Images/W-S48.png' width='15' height='15'  />"
	        End If	        
	        %>	        
	        </td>
	        <td  class="styleTDstd">
	            <%
	            If LogRs("RepairDateSerial") > 0 Then
	                Response.Write CDate(LogRs("RepairDateSerial"))
	            Else
    	            Response.Write ""
	            End If
	        %>
	        </td>
	        <td  class="styleTDstd"><%=LogRs("RepairNotes")%></td>
	        
	       	<td  class="styleTDstd"><%
	       	    If LogRs("Cost") > 0 Then
	       	        Response.Write LogRs("Cost")
	       	    Else
	       	        Response.Write "&nbsp;"    
	       	    End If
	       	%></td>       	
	       	
	       	
	       	<td  class="styleTDstd">
	       	<%
	       	    If LogRs("RecurringFault") = Cbool(True) Then
	       	        Response.Write "<img align='middle' alt='' title='Recurring Fault' src='Images/X2.ico' width='15' height='15'  />"
	       	    Else
	       	        Response.Write "&nbsp;"
	       	    End If	       	
	       	%>
	       	</td>
	       	
	       	
	       	
	      </tr>
        <%
        
        OkStatus = 0
        StatusBg = ""
        StatusText = ""
        LoggedByName = ""
        EditId = ""
        NotesTitleText = ""
        LogRs.MoveNext 
        
        FlushCount = FlushCount + 1
        If FlushCount >= 50 Then
            Response.Flush
            FlushCount = 0
        End If
                                             
        Wend
        %>
        
        </table>
        <br />
        <hr style="border-style: none; width: 98%; height: 1px; background-color: <%=NewCyan%>; display: block;" />         
        
        <%
        LogRs.Close
        Set LogRs = Nothing
                   
      End If

    MachineRs.MoveNext 
    Wend

    MachineRs.Close  
    Set MachineRs = Nothing    
        
%>   
<br /><br /><br />
<table>
<tr>
<td>
<%If InputType = "text" Then Response.Write "Locked"%> 
<input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" />&nbsp;
<%If InputType = "text" Then Response.Write "Sendpage"%> 
<input type="<%=InputType %>" name="Sendpage" id ="Sendpage" value="<%=RepairStatus%>" />&nbsp;
</td>
</tr>
</table>
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