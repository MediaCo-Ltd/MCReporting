<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="HsGetData.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<% 
'##On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect "SystemLocked.asp"

Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim LogTypeId
Dim ReasonId
Dim UserId
Dim ShiftId
Dim StartSerial
Dim EndSerial
Dim DateSourceId
Dim LocationId

'## Below will allways be fixed only UserId changes 
LogTypeId = Cint(4)
ReasonId = Cint(0)
UserId = Cint(Session("UserId"))
ShiftId = Cint(0)
StartSerial = Clng(0)
EndSerial = Clng(0)
DateSourceId = Cint(0)
LocationId = Cint(0)

Dim EditId
Dim OkStatus 
Dim ReturnAddress
Dim Titletext
Dim NoDisplay
Dim FlushCount
Dim Msg
Dim StatusBg
Dim StatusText
Dim HeaderText
Dim OtherColour
Dim NotesTxt

NotesTxt = ""
 
HeaderText = "All Records created by " & Session("UserName")


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
    <title>HS Reporting</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link href="CSS/HSReportsCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/HSReportsExtraCss.css" rel="stylesheet" type="text/css" />
    <link rel="shortcut icon" href="Images/Plus-icon.png" type="image/x-icon" />
    <script type="text/javascript" src="JsFiles/HSReportsJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" id="Display" >  <!--onload="javascript:ViewRecordsLoad();"-->

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
            &nbsp;&nbsp;<a href ="javascript:GoBackHSOption();" style="font-size:12px; color:<%=NewBlue%>">Return&nbsp;to&nbsp;select&nbsp;option&nbsp;page</a>            
        </td>
        <td align="center" valign="bottom" width="34%">
            <img align="top" alt="mediaco logo" src="Images/Plus-icon.png" style="width: 20px; height: 20px;" /> 
            <font style="font-weight: bold; color:<%=NewBlue%>; font-size: 16px;" >&nbsp;<%=HeaderText%></font>                     
        </td>
        <td align="right" width="33%">
            <a href ="javascript:LogOff()" style="font-size:12px; color:<%=NewBlue%>">Log&nbsp;Out</a>&nbsp;&nbsp;    
        </td>
    </tr>
   
    <tr>     
        <td align="left" valign="middle" colspan="3">
            <br />&nbsp;To&nbsp;view&nbsp;a&nbsp;record&nbsp;click&nbsp;anywhere&nbsp;in&nbsp;the&nbsp;Record&nbsp;Id&nbsp;area 
        </td>
    </tr>
</table>


<%

Call OpenCustomRs(LogTypeId,ReasonId,UserId,ShiftId,LocationId,Clng(StartSerial),Clng(EndSerial),DateSourceId,"Display")

Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;&nbsp;&nbsp;" & HeaderText

If CustomRs.Bof = False or CustomRs.EOF = False Then

       
    %>    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">    
	   <tr >
	        <th class="styleTHleft" width="5%" >Record Id</th>
	        <th class="styleTHstd" width="10%">Created</th>	        
	        <th class="styleTHstd" width="5%">Severity</th>
	        <th class="styleTHstd" width="10%">Log Type</th>
	        <th class="styleTHstd" width="15%">Reason</th>
	        <th class="styleTHstd" width="5%">Resolved</th>
	        <th class="styleTHstd" width="10%">Date/Time of Incident</th>
	        <th class="styleTHstd" width="10%">Last Updated</th>
        </tr>    
                
        <%
   
   Msg = Msg & "</h4>"
   Response.Write (Msg)         
            
   While Not CustomRs.EOF
        
                
        NotesTxt = ""
        EditId = CustomRs("Id") 
        %>        
        <tr>
            <td title="" style="color: #0069AA" class="styleTDleft"
            onclick="javascript:HSLoadViewOrderUser(<%=EditId%>);"><%=CustomRs("Id")%>
            </td>
	        <td class="styleTDstd"><%=Cdate(CustomRs("CreatedDate"))%></td>	    
	        <%
	            If CustomRs("Severity") = 1 Then
	                StatusBg = "#00BB00"
	                StatusText = "Minor"
	            ElseIf CustomRs("Severity") = 2 Then
	                StatusBg = "#FFA500"  'lighter "#FFBC55" 
	                StatusText = "Medium"
	            Else
	                StatusBg = "#FF0000"
	                StatusText = "Critical"
	            End If
	            
	            If CustomRs("GroupSelection") = 15 Then
	                OtherColour = " style='color: #FF0000' "
	            Else
	                OtherColour = " style='color: #000000' "
	            End If
	        %>
	        
	        <td title="<%=StatusText%>" class="styleTDstd" style="background-color:<%=StatusBg%>;"></td>
	        <td class="styleTDstd">
	            <%
	                If CustomRs("GroupId") = 1 Then
	                    Response.Write ("Accident") 
	                ElseIf CustomRs("GroupId") = 2 Then
	                    Response.Write ("Near Miss")
	                Else
	                    Response.Write ("Unsafe")
	                End If
	                
	                NotesTxt = Replace(CustomRs("Notes"),"<br />",VbCrLf,1,-1,1)
	            %>
	        </td>
	        <td class="styleTDstd" <%=OtherColour%> title="<%=NotesTxt%>">
	            <%
	            If CustomRs("GroupSelection") = 15 Then
	                If CustomRs("GroupId") = 1 Then
	                    Response.Write ("Accident - Other") 
	                ElseIf CustomRs("GroupId") = 2 Then
	                    Response.Write ("Near Miss - Other")
	                Else
	                    Response.Write ("Unsafe - Other")
	                End If
	            Else
	                Response.Write CustomRs("Description")
	            End If
	            %>
	        </td>
	        <td class="styleTDstd">
	        <%
	        If CustomRs("Resolved") = Cbool(True) Then
	            Response.Write "<img align='middle' alt='' title='Resolved' src='Images/checkmark.ico' width='15' height='15'  />"
	        Else
	            Response.Write "<img align='middle' alt='' title='Pending' src='Images/X2.ico' width='15' height='15'  />"
	        End If	        
	        %>	        
	        </td>
	        <td class="styleTDstd"><%=CustomRs("SelectedDate")%></td>
	        <td class="styleTDstd">
	            <%If CustomRs("LastModifiedDateSerial") > 0 Then Response.Write Cdate(CustomRs("LastModifiedDateSerial"))%>
	        </td>
	       	
	       
	      </tr>
        <%
        
        OkStatus = 0
        StatusBg = ""
        StatusText = ""
       
        EditId = ""
        NotesTxt = ""
        
        FlushCount = FlushCount + 1
        If FlushCount >= 50 Then
            Response.Flush
            FlushCount = 0
        End If

        CustomRs.MoveNext 
    Wend
    
    %>
        </table>
        <!--<br />
        <hr style="border-style: none; width: 98%; height: 1px; background-color: <%'=NewCyan%>; display: block;" />-->         
        
        <%
End If

CustomRs.Close  
Set CustomRs = Nothing   
        
%>   
<br /><br /><br />
<table>
<tr>
<td>
<%If InputType = "text" Then Response.Write "Locked"%> 
<input type="<%=InputType %>" name="Locked" id ="Locked" value="<%=Session("Locked")%>" />&nbsp;

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