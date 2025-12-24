<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="IhrGetQcData.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<% 
'On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim DateRange
DateRange = Request.QueryString("sd")

Dim TheDateSerial
TheDateSerial = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))

Dim SwapDateStart
Dim SwapDateEnd

Dim TheDay
TheDay = WeekdayName(Weekday(Date,1) ,False,1)

Dim DisplayRange
Dim ReturnAddress

Dim StartDate
Dim EndDate
Dim MyEndDate

If DateRange = 1 Then   '## Last Night

    ReturnAddress = "1"

    If TheDay = "Monday" Then   '## Show weekend as well
        SwapDateStart = Cdate(TheDateSerial-3)
        SwapDateEnd = Cdate(TheDateSerial)
        DisplayRange = "&nbsp;&nbsp;Created&nbsp;between&nbsp;Friday&nbsp;19:00&nbsp;&amp;&nbsp;08:00&nbsp;this&nbsp;morning"
    Else
        SwapDateStart = Cdate(TheDateSerial-1)
        SwapDateEnd = Cdate(TheDateSerial)
        DisplayRange = "&nbsp;&nbsp;Created&nbsp;between&nbsp;19:00&nbsp;last&nbsp;night&nbsp;&amp;&nbsp;08:00&nbsp;this&nbsp;morning"        
    End If
    
    SwapDateStart = Mid(SwapDateStart,4,3) & Left(SwapDateStart,3) & Mid(SwapDateStart,7,4)
    SwapDateEnd = Mid(SwapDateEnd,4,3) & Left(SwapDateEnd,3) & Mid(SwapDateEnd,7,4)
    DateRange = " Created Between #" & SwapDateStart & " 18:59:59# AND #" & SwapDateEnd & " 07:59:59#)"
    
ElseIf DateRange = 2 Then '## This week
    
    ReturnAddress = "2"
    
    If TheDay = "Monday" Then
        SwapDateStart = TheDateSerial
        DisplayRange = "Created&nbsp;This&nbsp;Week"   
    ElseIf TheDay = "Tuesday" Then  
        SwapDateStart = TheDateSerial-1
        DisplayRange = "Created&nbsp;This&nbsp;Week"
    ElseIf TheDay = "Wednesday" Then 
        SwapDateStart = TheDateSerial-2
        DisplayRange = "Created&nbsp;This&nbsp;Week"
    ElseIf TheDay = "Thursday" Then  
        SwapDateStart = TheDateSerial-3
        DisplayRange = "Created&nbsp;This&nbsp;Week"
    ElseIf TheDay = "Friday" Then 
        SwapDateStart = TheDateSerial-4
        DisplayRange = "Created&nbsp;This&nbsp;Week"
    ElseIf TheDay = "Saturday" Then  
        SwapDateStart = TheDateSerial-5
        DisplayRange = "Created&nbsp;This&nbsp;Week"
    ElseIf TheDay = "Sunday" Then  
        SwapDateStart = TheDateSerial-6
        DisplayRange = "Created&nbsp;This&nbsp;Week"
    End If
    
    SwapDateEnd = TheDateSerial
    DateRange = " CreatedSerial Between " & SwapDateStart & " AND " & SwapDateEnd & ")"

ElseIf DateRange = 3 Then '## This month
    
    ReturnAddress = "3"
    DisplayRange = "Created&nbsp;This&nbsp;Month"    
    
    StartDate = CDbl(DateSerial(Year(Date),Month(Date),"01"))    
    DateRange = " CreatedSerial Between " & StartDate & " AND " & TheDateSerial & ")"
    
ElseIf DateRange = 6 Then '## Today
    
    ReturnAddress = "6"
    DisplayRange = "Created&nbsp;Today"    
    
    StartDate = TheDateSerial    
    DateRange = " CreatedSerial Between " & StartDate & " AND " & TheDateSerial & ")"
    
Else    '## This year
    
    ReturnAddress = "4"    
    DisplayRange = "Created&nbsp;This&nbsp;Year" 
         
    StartDate = CDbl(DateSerial(Year(Date),"01","01"))    
    DateRange = " CreatedSerial Between " & StartDate & " AND " & TheDateSerial & ")"
        
End If

Dim OkStatus 
Dim Digital
Dim Titletext
OkStatus = 0
Dim EditId
EditId = ""
Dim NoDisplay
Dim FlushCount
Dim Msg
Dim StatusBg
Dim StatusText
Dim HeaderText
Dim DisplayRs
Dim GetDatesRS

Dim DisplayDept
Dim DisplayReason

Dim InputType
InputType = Session("ShowHidden")

Dim MatPos
Dim MatTitle
Dim MatDisplay

Dim RevisionTxt
RevisionTxt = ""
Dim RevisionTitle
RevisionTitle = ""

Dim NotActiveColour
Dim NotActiveTitle

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en" >
<head>
    <title>In House Redo</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/X1.ico" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/IhRedoJSFunc.js"></script>
</head>

<body style="padding: 0px; margin: 0px" id="Display" >  <!--onload="javascript:ViewRecordsLoad();"-->

<table style="width: 100%;padding-right: 10px; padding-left: 10px;" >
    <tr>
	    <td align="left" valign="bottom" height="100" colspan="3" >
            <img align="left" alt="mediaco logo" src="<%=CompanyLogo%>" width="160" />
        </td>
    </tr>
    <tr>
	    <td height="8" valign="top" colspan="3" >
	        <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
	    </td>
    </tr>    
	
    <tr>
        <td align="left" style="height: 20px" valign="bottom" width="33%" >
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('IhrJobNo.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return&nbsp;to&nbsp;main&nbsp;page</a>            
        	&nbsp;&nbsp;<a href ="javascript:window.location.replace('IhrSelectDate.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to select date</a>
        </td>
        <td align="center" valign="bottom" width="34%">
            <img align="top" alt="mediaco logo" src="Images/X1.ico" style="width: 20px; height: 20px;" />&nbsp;
            <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">&nbsp;Redo&nbsp;Records&nbsp;<%=DisplayRange%></font>                     
        </td>
        <td align="right" width="33%">
        <%If Session("ShowHidden") = "text" Then Response.Write "ReturnAddress" %>
        &nbsp;<input id="ReturnAddress" value="<%=ReturnAddress%>" type="<%=Session("ShowHidden")%>" />
        <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;     
        </td>
    </tr>
   
    <tr>     
        <td align="left" valign="middle" colspan="3" style="font-size:12px; color: <%=NewCyan%>;">        
            <%
            Response.Write "<br />&nbsp;To&nbsp;view&nbsp;a&nbsp;record&nbsp;click&nbsp;anywhere&nbsp;in&nbsp;the&nbsp;Quote&nbsp;Ref&nbsp;area"            
            %>   
        </td>
    </tr>
</table>


<%

'Call OpenDisplayLastNightRs(DateRange)


'While Not GetDatesRS.Eof 
    Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;" & DisplayRange & "</h4>"   
      
    Call OpenDisplayByDateRs (DateRange)
    '## If no records don't draw tables
        
    If DisplayRs.Bof = False or DisplayRs.EOF = False Then
        Response.Write (Msg)
    %>    
    <table  align="center" cellpadding="0" cellspacing="0"  width="95%">    
	   <tr >
	        <th class="styleTHleft" width="5%">Job</th>
	        <th class="styleTHstd" width="5%">Redo&nbsp;Date</th>	        
	        <th class="styleTHstd" width="8%">Client</th>
	        <th class="styleTHstd" width="12%">Description</th>
	        <th class="styleTHstd" width="10%">Substrate</th>
	        <th class="styleTHstd" width="5%">Redo Sqm</th>
	        <th class="styleTHstd" width="10%">Departments</th>	        
	        <th class="styleTHstd" width="10%">Reason</th>
	        <th class="styleTHstd" width="5%">Created By</th>
	        <%
	            If Session("ShowRedoCost") = CBool(True) Then Response.Write "<th class='styleTHstd' width='5%'>Total Cost</th>"
	        %>
	        
        </tr>    
                
        <%
        While Not DisplayRs.Eof
            '## Record Id
            EditId = DisplayRs("Id") 
            '## QuoteId
            '## EditId = DisplayRs("QuoteId")
            
            Call GetDepartments (DisplayRs("DeptId"))
            Call GetReasons (DisplayRs("ReasonCodes"))
            
            If Instr(1,DisplayRs("Substrates"),"#",1) > 0 Then
                MatPos = Instr(1,DisplayRs("Substrates"),"#",1)
                MatDisplay = Left(DisplayRs("Substrates"),MatPos-1) & " ..."
                MatTitle = Replace(DisplayRs("Substrates"),"#",VbCrLf,1,-1,1)
            Else
                MatTitle = ""
                MatDisplay = DisplayRs("Substrates")
            End If
            
            If Cint(DisplayRs("Revision")) = 1 Then
                RevisionTxt = ""
                RevisionTitle = ""
            Else
                RevisionTxt = " (" & DisplayRs("Revision") & ")"
                RevisionTitle = " title ='Redo No "  & DisplayRs("Revision") & " for this job' "
            End If
            
            If DisplayRs("Active") = Cbool(False) Then
                NotActiveTitle = " title='Record is dormant and is not listed in emailed reports' "
                NotActiveColour = " style='background-color: #FFFF99' "            
            Else
                NotActiveTitle = ""
                NotActiveColour = " style='background-color: #FFFFFF' "
            End If
       
        %>        
        <tr <%=NotActiveColour%> <%=NotActiveTitle%> >
            <td id="<%=EditId%>" style="color: #0069AA; text-align: left; text-indent: 5px;" class="styleTDleft" <%=RevisionTitle%>
                <%Response.Write " onclick='javascript:LoadRedo(" & EditId & ");'" %> ><%="REF" & DisplayRs("QuoteRef") & RevisionTxt%>
            </td>	        
	        <td  class="styleTDstd"><%=DisplayRs("Created")%></td>
	        <td class="styleTDstd"><%=DisplayRs("Client")%></td>                    <!--style=" text-align: left;  text-indent: 5px;"-->
	        <td  class="styleTDstd"><%=DisplayRs("Description")%></td>              <!--style=" text-align: left;  text-indent: 5px;"-->
	        <td  class="styleTDstd" title="<%=MatTitle%>"><%=MatDisplay%></td>	        
	        <td  class="styleTDstd"><%=DisplayRs("RedoSqm")%></td>
	        <td  class="styleTDstd"><%=DisplayDept%></td>	        
	        <td  class="styleTDstd"><%=DisplayReason%></td>
	        <td  class="styleTDstd"><%=DisplayRs("CreatedBy")%></td>
	        <%
	            If Session("ShowRedoCost") = CBool(True) Then Response.Write "<td class='styleTDstd' >" & DisplayRs("RedoCost") & "</td>"
	        %>
	        
	      </tr>
        <%
        
        OkStatus = 0
        StatusBg = ""
        StatusText = ""
        DisplayDept = ""
        DisplayReason = ""
        EditId = ""
        
        MatDisplay = ""
        MatTitle = ""
        MatPos = 0
        RevisionTxt = ""
        RevisionTitle = ""
        
        DisplayRs.MoveNext        
        
        
        'FlushCount = FlushCount + 1
        'If FlushCount >= 50 Then
        '    Response.Flush
        '    FlushCount = 0
        'End If
                                             
        Wend
        %>
        
        </table>
        <br />
        <hr style="border-style: none; width: 98%; height: 1px; background-color: <%=NewCyan%>; display: block;" />         
        
        <%
        DisplayRs.Close
        Set DisplayRs = Nothing
                   
      End If

      'GetDatesRS.MoveNext 
'Wend

'GetDatesRS.Close  
'Set GetDatesRS = Nothing   
        
%>   
<br /><br /><br />

</body>
</html>
