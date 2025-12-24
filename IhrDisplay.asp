<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="IhrGetQcData.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->


<% 
'## On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim OkStatus 
Dim ReturnAddress
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

HeaderText = "All Redo Records"

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
            <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">&nbsp;<%=HeaderText%></font>                             
        </td>
        <td align="right" width="33%">
            <%If Session("ShowHidden") = "text" Then Response.Write "ReturnAddress" %>
            &nbsp;<input id="ReturnAddress" value="0" type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%>/>
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

Call OpenGetDatesRS()


While Not GetDatesRS.Eof 
    Msg = "<h4 style = 'color: Black' >&nbsp;&nbsp;Created&nbsp;on&nbsp;" & Cdate(GetDatesRS("CreatedSerial")) & " </h4>"   
      
    Call OpenDisplayRs (GetDatesRS("CreatedSerial"))
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
	        <td class="styleTDstd" ><%=DisplayRs("Client")%></td>                   <!--style=" text-align: left;  text-indent: 5px;"-->
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

      GetDatesRS.MoveNext 
Wend

GetDatesRS.Close  
Set GetDatesRS = Nothing   
        
%>   
<br /><br /><br />

</body>
</html>
