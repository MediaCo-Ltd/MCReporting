<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\Declarations.asp" -->
<!--#include file="..\##GlobalFiles\PkId.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\connIpAddressDB.asp" -->

<% 
'## On Error Resume Next 

Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"

If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim DetailRS
Dim strDetailSql
Dim DetailJob
Dim DetailItems
Dim DetailItemsParts
Dim DetailItemsMatData
Dim ItemMatData

Dim DetailItemsDept
Dim DetailItemsDeptID
Dim DetailItemsQcDeptID
Dim DetailItemsWcID
Dim ItemAlpha
Dim ItemDept
Dim ItemDeptId
Dim ItemQcDeptId
Dim ItemWcID

Dim DetailPartUnitCost
Dim DetailPartSqm
Dim DetailSubPartSqm
Dim DetailOperationStatus
Dim DetailPartDescription
Dim DetailPartSaveDescription

Dim LocalSubstrate
Dim Fabrication(3,1)
Dim FabDummy
Dim ProdNotes
Dim DisplayStatus
Dim DisplaySub

Dim OpNotesExist
Dim OperationNotes(6)

Dim PLCWPrivateNotes
Dim DetailPartCode
Dim DetailPartQty
Dim DetailSubPartQty
Dim DetailSubPartItemCount
Dim DetailPrinterName

Dim Client
Dim DeliveryAddress

Dim SubCount

Dim Combined
Dim JobIsOutdoor

Dim AlphaItemsParts
Dim AlphaItems
Dim TotalRows

Dim ReqDelHeader
Dim ItemTitle
Dim DetailMaterialCount

Dim MgsColSpan

Dim ReasonSize

If Session("ShowHiddenSize") = "" Then
    ReasonSize = " size='40' "
Else
    ReasonSize = ""
End If    

DetailJob = Session("JobNo")

'Not Distinct  Can't use if showing notes 
strDetailSql = "SELECT Job.Reference AS JobRef, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId, Quotation.user11 As ArtworkStatus,"
strDetailSql = strDetailSql & " Quotation.Description AS JobDesc, Quotation.Notes, Job.Id as JobId, Quotation.User6,"
strDetailSql = strDetailSql & " Case When Quotation.User15 = NULL Then '' Else Quotation.User15 END AS ArtworkMatch,"
strDetailSql = strDetailSql & " Quotation.status AS QuoteStatus, Quotation.JobTypeId, Quotation.completedDate AS DelDate, Contacts.Forename + ' ' + Contacts.Surname AS BookedBy"
strDetailSql = strDetailSql & " FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId INNER JOIN Contacts ON Job.CreatedBy = Contacts.Id"
strDetailSql = strDetailSql & " Where(Quotation.id = '" & DetailJob & "')"
 
Set DetailRS = Server.CreateObject("ADODB.Recordset")

DetailRS.ActiveConnection = Session("ClarityConn")   'strConnClarity
DetailRS.Source = strDetailSql
DetailRS.CursorType = Application("adOpenForwardOnly")
DetailRS.CursorLocation = Application("adUseClient")
DetailRS.LockType = Application("adLockReadOnly")
DetailRS.Open
Set DetailRS.ActiveConnection = Nothing

If DetailRS.BOF = True Or DetailRS.EOF = True Then
    DetailRS.Close
    Set DetailRS = Nothing
    Err.Clear
    Response.Redirect "Error.asp"
End If


If DetailRS("JobTypeId") =	OutJobTypeId Then
    JobIsOutdoor = True
    MgsColSpan = 6
Else
    JobIsOutdoor = False
    MgsColSpan = 7
End If

Select Case DetailRS("QuoteStatus")
    Case 3
        DisplayStatus = "Confirmed"
        ReqDelHeader = "Required"
    Case 5
        DisplayStatus = "Delivered" 
        ReqDelHeader = "Delivered On"
    Case 6
        DisplayStatus = "Invoiced"
        ReqDelHeader = "Delivered On"
    Case 13
        DisplayStatus = "Part Delivered"
        ReqDelHeader = "Required"
End Select

Dim VisibilityTxt
VisibilityTxt = "style='visibility: collapse;'"

'## Get Address details and Company Name
DelAddress DetailRS("QuoteId")

Dim hQuoteRef
Dim hSqm
Dim hDesc

hSqm = 0
hQuoteRef = Mid(DetailRS("JobRef"),4)
hDesc = DetailRS("JobDesc")

Dim RedoQty
RedoQty = 0

Dim BookedInBy
BookedInBy = DetailRS("BookedBy")

Dim SpItemsTrue
SpItemsTrue = False


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>

    <title>In House Redo</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/X1.ico" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/IhRedoJSFunc.js"></script>
</head>


<body style="padding: 0px; margin: 0px" onload="javascript:DataPageLoad();">

<table style="padding-right: 10px; padding-left: 10px; width: 100%;" >
	<tr>
		<td align="left" valign="bottom" height="100">
            <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
        </td> 
    </tr>	
	<tr>
		<td height="8" valign="top" align="left" width="100%" colspan="4">
            <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
        </td>
	</tr>
</table>

<table style="padding-right: 10px; padding-left: 10px; width: 100%;">
    <tr>
        <td align="left" style="height: 20px" width="33%">
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('IhrJobNo.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to main page</a><br /><br />
        </td>
        <td align="center" width="34%" >
            <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">New In House Redo Record</font>                  
        </td>
        <td align="right" valign="bottom" width="33%">
            <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;
        </td> 
    </tr>
</table>

<br />
<form action="IhrUpdate.asp" method="post" onsubmit="return ValidateData();">
<table style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">
<tr >
    
    <th width="50" class="styleTHleft">Job</th>
    <th width="50" class="styleTHstd"><%=ReqDelHeader%></th>
    <%
        If JobIsOutdoor = False Then
            Response.Write "<th width='50' class='styleTHstd'>Created By</th>"
        End If
    %>
    <th width="130" class="styleTHstd">Company</th>        
    <th width="130" class="styleTHstd">Description</th>
    <th width="50" class="styleTHstd" >Status</th>
    <th width="50" class="styleTHstd">Delivery Method</th>
    <th width="50" class="styleTHstd">Colour Match</th>
</tr>
<tr >
    
    <td width="50" class="styleTDleft" style="color: #000000"><%=DetailRS("JobRef") %></td>
    <td width="50" class="styleTDstd">
    <%
        If ReqDelHeader = "Delivered On" Then
            Response.Write Trim(Left(DetailRS("DelDate"),10))    
        Else
            Response.Write Trim(Left(DetailRS("ReqDate"),10))
        End If    
    %>
    </td> 
    
    <%
        If JobIsOutdoor = False Then
            Response.Write "<td width='50' class='styleTDstd'>" & DetailRS("BookedBy") & "</td>"
        End If
    %>
    <td width="130" class="styleTDstd"><%=Client%></td>       
    <td width="130" class="styleTDstd"><%=DetailRS("JobDesc") %></td>
    <td width="50" class="styleTDstd" ><%=DisplayStatus%></td>
    <td width="50" class="styleTDstd"><%=DetailRS("User6") %></td>
    <td width="50" class="styleTDstd"><%=DetailRS("ArtworkMatch") %></td>

</tr>
</table>

<br />

<table style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">
<tr>
<th align="left" style="font-size: 12px; color: #0069AA;" nowrap="nowrap">
<label id="SelectTxt">Select&nbsp;which&nbsp;Item(s)&nbsp;you&nbsp;need&nbsp;to&nbsp;create&nbsp;a&nbsp;record&nbsp;for.</label>
<label style="color: #FF0000">Welding,&nbsp;Sewing,&nbsp;&amp;&nbsp;Eyeletting&nbsp;are&nbsp;all&nbsp;grouped&nbsp;under&nbsp;Finishing.</label>
<br /><br />
If&nbsp;Item&nbsp;redo&nbsp;qty&nbsp;is&nbsp;different&nbsp;to&nbsp;Item&nbsp;original&nbsp;qty.&nbsp;<label style="color: #FF0000">Change&nbsp;value&nbsp;in&nbsp;box.</label>&nbsp;
If&nbsp;needed&nbsp;add&nbsp;a&nbsp;brief&nbsp;description&nbsp;in&nbsp;the&nbsp;reason&nbsp;box.
<br />
Depending&nbsp;on&nbsp;department&nbsp;&amp;&nbsp;Item&nbsp;selected,&nbsp;the&nbsp;reason&nbsp;will&nbsp;be&nbsp;automaticaly&nbsp;selected.
But&nbsp;you&nbsp;can&nbsp;still&nbsp;select&nbsp;others&nbsp;in&nbsp;the&nbsp;list.
<br /><!--<br />-->
Only&nbsp;one&nbsp;Department/Reason&nbsp;can&nbsp;be&nbsp;selected&nbsp;per&nbsp;item,&nbsp;so&nbsp;select&nbsp;the&nbsp;most&nbsp;relevant.&nbsp;
If&nbsp;you&nbsp;need&nbsp;more&nbsp;create&nbsp;a&nbsp;new&nbsp;record.&nbsp;There&nbsp;is&nbsp;no&nbsp;limit&nbsp;to&nbsp;No&nbsp;of&nbsp;records&nbsp;for&nbsp;the&nbsp;same&nbsp;job.
<br /><br />
<!--<label id="NoHold" style="color: #FF0000;" >
Your&nbsp;Operation&nbsp;must&nbsp;be&nbsp;put&nbsp;on&nbsp;hold&nbsp;in&nbsp;Clarity&nbsp;before&nbsp;creating&nbsp;the&nbsp;record.
</label>-->

</th>
</tr>

<tr>
<td>
<%
    '## Create tick box for each Item that has been created by a Calc Wizard
    '## If only 1 Item no boxes show & Item is visible
    GetItems DetailRS("QuoteId")
    
    If TotalRows = 0 Then          
        Session("JobNoError") = "Job No. " & DetailRS("JobRef") & " has no valid data"
        DetailRS.Close
        Set DetailRS = Nothing
        Response.Redirect("IhrJobNo.asp")
    End If
    
    DetailItemsParts = Split(DetailItems,";",-1,1)
    DetailItems = ""   
    
    AlphaItemsParts = Split(AlphaItems,";",-1,1)    
    AlphaItems = ""
    
       
    
    If TotalRows = 1 Then
        If GetOpStatus (DetailItemsParts(0)) = True Then 
            VisibilityTxt = ""
        Else
            VisibilityTxt = "style='visibility: collapse;'"            
        End If
    Else
       
        For Count = 0 to Ubound(DetailItemsParts)
            If DetailItemsParts(Count) <> "" Then    
                '## Create the Item Letter for normal quote lines
                If AlphaItemsParts(Count) <= 26 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)+64))
                ElseIf AlphaItemsParts(Count) >= 27 And AlphaItemsParts(Count) <= 52 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)+38))
                    ItemAlpha = "A" & ItemAlpha
                ElseIf AlphaItemsParts(Count) >= 53 And AlphaItemsParts(Count) <= 78 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)+12))
                    ItemAlpha = "B" & ItemAlpha 
                ElseIf AlphaItemsParts(Count) >= 79 And AlphaItemsParts(Count) <= 104 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)-14))
                    ItemAlpha = "C" & ItemAlpha 
                ElseIf AlphaItemsParts(Count) >= 105 And AlphaItemsParts(Count) <= 130 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)-14))
                    ItemAlpha = "D" & ItemAlpha
                ElseIf AlphaItemsParts(Count) >= 131 Then
                    ItemAlpha = "?"  
                End If         
            
            If GetOpStatus (DetailItemsParts(Count)) = True Then
                Response.Write "<input type= 'checkbox' value='0' onclick='javascript:ShowRow(" & Cstr(Count) & ");' id='chk" & Count & "' name='chk" & Count & "' />" & ItemAlpha & VbCrLf 
            Else             
                Response.Write "<input title='To enable selection your operation for item " & ItemAlpha & "." & VbCrLf & "Must be put on hold in Clarity before creating the record." & VbCrLf & "Once done reset page and continue.' type= 'checkbox' value='0' id='chk" & Count & "' name='chk" & Count & "' disabled='disabled'/>" & ItemAlpha & VbCrLf 
            End If
                    
            If Count = 40 Then Response.Write "<br />"        
            If Count = 80 Then Response.Write "<br />"        
            If Count = 120 Then Response.Write "<br />"
            If Count = 160 Then Response.Write "<br />"
            
            End If
        Next
    End If
    
    ItemAlpha = ""
 

%>
</td>
</tr>
</table>

<br />

<table  style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">	
    <tr >		 
	    <th width="10%" class="styleTHleft" >Item</th>
	    <th width="10%" class="styleTHstd" title="Add spreadsheet items" >SP Items</th> 
	    <th width="30%" class="styleTHstd" >Description</th>
	    <th width="20%" class="styleTHstd" >Department</th>
	    <!--<th width="10%" class="styleTHstd" >(Original)&nbsp;&nbsp;</th>-->
	    
	    <th width="10%" class="styleTHstd" >(Original)&nbsp;&nbsp;<label style="color: #FF0000">Redo&nbsp;Qty</label></th>
	    
        <th width="15%" class="styleTHstd" nowrap="nowrap" >       
        Select&nbsp;Reason&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
        </th>	
        <th width="15%" class="styleTHstd" nowrap="nowrap" >
        &nbsp;Add&nbsp;Additional&nbsp;Info&nbsp;if&nbsp;needed&nbsp;&nbsp;
        </th>	    		    
    </tr>

    <!-- do loop -->     
    <%   
    
    For Count = 0 to Ubound(DetailItemsParts)   
    
        If DetailItemsParts(Count) <> "" Then      
            
     
            DetailPartSqm = 0
            Combined = False
            DetailPartUnitCost = 0
            RedoQty = 0
            
            '## Get Sqm, Substrate, Pass QuotationItems.Id 
            CalculateDetailSqm DetailItemsParts(Count)
            
            If DetailPartQty > 1 Then
                RedoQty = 0
            Else
                RedoQty = DetailPartQty
            End If
            
            'RedoQty = DetailPartQty
            
            If Instr(1,DetailPartDescription,"Spreadsheet",1) > 0 Then SpItemsTrue = True       
                        
            '## Title ignores anything after a Quotation mark, may do same with Apostrophe
            ItemTitle = Cstr(DetailPartDescription)          
            ItemTitle = Replace(ItemTitle,Chr(34),"&#34;",1,-1,1) '## Quotation mark
            ItemTitle = Replace(ItemTitle,Chr(39),"&#39;",1,-1,1) '## Apostrophe
            DetailPartSaveDescription = Cstr(DetailPartDescription)
            
            DetailPartSaveDescription = Replace(DetailPartSaveDescription,Chr(13),"<br />",1,-1,1) '## Enter
            
            
            If Len(DetailPartDescription) > 70 Then 
                DetailPartDescription = Left(DetailPartDescription,60) & " ..."
            Else
                ItemTitle = ""
            End If
            
            DetailPrinterName = ""
            
            DetailPrinterName = GetPrinterName (DetailItemsParts(Count))        
            If DetailPrinterName <> "" Then 
                If DetailPrinterName = "CJV 300" Then    
                    DetailPrinterName = "CJV-JV 300"
                End If
            End If   
            
            DisplaySub = ""
            
            DisplaySub = LocalSubstrate
            If Right(DisplaySub,1) = "#" Then DisplaySub = Left(DisplaySub,Len(DisplaySub)-1)
            
       
            
    %>   
     
    <tr id="Row<%=Count%>"  <%=VisibilityTxt%> >                     
        <td width="25" class="styleTDleft" >
            <%
            '## ItemAlpha for all normal quote lines & PWlist items
            '## Is got in Get Items & split into a array
            '## If Combined it is got in OpStatusCombined
            If Combined = True Then
                Response.Write "<font style='color: #FF0000'>" & ItemAlpha & "</font>"
            Else
                '## Create the Item Letter for normal quote lines
                If AlphaItemsParts(Count) <= 26 Then
                     ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)+64))   
                ElseIf AlphaItemsParts(Count) >= 27 And AlphaItemsParts(Count) <= 52 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)+38))
                    ItemAlpha = "A" & ItemAlpha
                ElseIf AlphaItemsParts(Count) >= 53 And AlphaItemsParts(Count) <= 78 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)+12))
                    ItemAlpha = "B" & ItemAlpha 
                ElseIf AlphaItemsParts(Count) >= 79 And AlphaItemsParts(Count) <= 104 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)-14))
                    ItemAlpha = "C" & ItemAlpha 
                ElseIf AlphaItemsParts(Count) >= 105 And AlphaItemsParts(Count) <= 130 Then
                    ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)-14))
                    ItemAlpha = "D" & ItemAlpha
                ElseIf AlphaItemsParts(Count) >= 131 Then
                    ItemAlpha = "?"  
                End If                               
                    Response.Write "<font style='color: #000000'>" & ItemAlpha & "</font>"
                
            End If
            
              
            %>
            &nbsp;
            <input type="<%=Session("ShowHidden")%>"  value="<%=DetailItemsParts(Count)%>" id="ItemIdRow<%=Count%>" 
            <%=Session("ShowHiddenSize")%>  name="ItemIdRow<%=Count%>" />
            &nbsp;
            <input type="<%=Session("ShowHidden")%>"  value="<%=ItemAlpha%>" id="ItemIdAlpha<%=Count%>" 
            <%=Session("ShowHiddenSize")%>  name="ItemIdAlpha<%=Count%>" />
            
            <% 
            DetailItemsMatData = GetMaterialData(DetailItemsParts(Count))
            ItemMatData = Split(DetailItemsMatData,"#",-1,1)
            DetailItemsMatData = ""
            %>
               
            <input type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%> id='MatCost<%=Count%>'
                name="MatCost<%=Count%>" value="<%=ItemMatData(0)%>" />
                
            <input type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%> id='MatGroup<%=Count%>'
                name="MatGroup<%=Count%>" value="<%=ItemMatData(1)%>" /> 
                
            <input type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%> id='MatCategory<%=Count%>'
                name="MatCategory<%=Count%>" value="<%=ItemMatData(2)%>" />
             
        </td>
        
        <td class="styleTDstd" align="left" >
            &nbsp;<input type="text" id="SpId<%=Count%>" name="SpId<%=Count%>" onkeypress='return DisableEnterKey(event);'/>
            <input type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%> id='SpTrue<%=Count%>'
                name="SpTrue<%=Count%>" value="<%=SpItemsTrue%>" />
        </td>
              
        <td class="styleTDstd" title="<%=ItemTitle%>" > <%=DetailPartDescription%> 
            <input type="<%=Session("ShowHidden")%>" id="ItemDesc<%=Count%>" <%=Session("ShowHiddenSize")%>  
                name="ItemDesc<%=Count%>" value="<%=DetailPartSaveDescription%>"/> 
            
            <%If Session("ShowHidden") = "text" Then Response.Write "&nbsp;DisplaySub"%>
            <input type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%> id="ItemSubtrate<%=Count%>"
                name="ItemSubtrate<%=Count%>" value="<%=DisplaySub%>" />
                
               
            <%If Session("ShowHidden") = "text" Then Response.Write "&nbsp;ItemCost"%>  
            <input type="<%=Session("ShowHidden")%>" <%=Session("ShowHiddenSize")%> id='ItemCost<%=Count%>'
                name="ItemCost<%=Count%>" value="<%=DetailPartUnitCost%>" />
                
              
                                       
        </td> 
        <td class="styleTDstd">
            <select id="CboDept<%=Count%>" onclick="javascript:AddWc('<%=Count%>')" style="width: 180px" >
            
            <%
            Erase ItemMatData
            
            GetDepartments DetailItemsParts(Count)
            DetailItemsDept = Split(ItemDept,";",-1,1)
            DetailItemsDeptID = Split(ItemDeptID,";",-1,1)
            DetailItemsQcDeptID = Split(ItemQcDeptId,";",-1,1)
            DetailItemsWcID = Split(ItemWcID,";",-1,1)  
                        
            If ItemDept = "" Then
                Response.Write "<option value='' disabled='disabled'>Select Department</option>" & VbCrLf
                ItemDept = ""
            Else
                Response.Write "<option value=''>Select Department</option>" & VbCrLf
            End If
            ItemDeptID = ""  
            
            '## Add Booking in
            Response.Write "<option value='0#0#0'>Account Management</option>" & VbCrLf          
                       
            For SubCount = 0 To Ubound(DetailItemsDept) -1              
                Response.Write "<option value=" & DetailItemsDeptID(SubCount) & "#" & DetailItemsQcDeptID(SubCount) &  "#" & DetailItemsWcID(SubCount) & ">" & DetailItemsDept(SubCount) & "</option>" & VbCrLf
            Next
            
            
            '## Add stock option
            Response.Write "<option value='0#7#0'>Stock</option>" & VbCrLf            
            
            %>
            </select>
            &nbsp;
            <input type="<%=Session("ShowHidden")%>" id="WCID<%=Count%>" value="" <%=Session("ShowHiddenSize")%> name="WCID<%=Count%>" />  
            <input type="<%=Session("ShowHidden")%>" id="OrigQty<%=Count%>" value="<%=DetailPartQty%>" <%=Session("ShowHiddenSize")%>  name="OrigQty<%=Count%>" />     
            <input type="<%=Session("ShowHidden")%>" id="Grp<%=Count%>" value="" <%=Session("ShowHiddenSize")%>  name="Grp<%=Count%>" />
        </td>
        
        <td class="styleTDstd" >&nbsp;&nbsp;&nbsp;&nbsp;<label style="margin-left: 10px;"><%="(" & DetailPartQty & ")"%></label><!--style="margin-left: 40px;"-->
        
        
            <input type="text" id="Qty<%=Count%>" size= "5" value="<%=RedoQty%>" 
            style="margin-top: 2px; margin-bottom: 2px; color: #FF0000; font-weight: bold; margin-left: 30px;" 
            onkeypress='return DisableEnterKey(event);' onfocus='this.select();' onmouseup='return false;' name="Qty<%=Count%>" />
        </td>
        
        
        
        <td class="styleTDstd" align="left" >
            &nbsp;
            <select id="CboReason<%=Count%>" style=" visibility: hidden; width: 180px;" onclick="javascript:AddCodes('<%=Count%>');" > <!--visibility: hidden;-->
                <option value='' >Select Reason</option>            
            </select>
            
            <%If Session("ShowHidden") = "text" Then Response.Write "<br />"%>
            
            <input type="<%=Session("ShowHidden")%>" id="Code<%=Count%>"  name="Code<%=Count%>" <%=Session("ShowHiddenSize")%> />
            <%If Session("ShowHidden") = "text" Then Response.Write "Code&nbsp;"%>
           
            <input type="<%=Session("ShowHidden")%>" id="Dept<%=Count%>"  name ="Dept<%=Count%>" <%=Session("ShowHiddenSize")%> />
            <%If Session("ShowHidden") = "text" Then Response.Write "Dept&nbsp;"%> 
            
            <input type="<%=Session("ShowHidden")%>" id="Prt<%=Count%>"  name ="Prt<%=Count%>" <%=Session("ShowHiddenSize")%>
            value = "<%=DetailPrinterName%>" />
            <%If Session("ShowHidden") = "text" Then Response.Write "Printer&nbsp;"%>           
           
       </td>
           
       <td class="styleTDstd" align="left" >  
       
            <textarea  rows="1" id="Reason<%=Count%>" name="Reason<%=Count%>" style="text-align: left; height: 13px; text-indent: 1px; width: 75%;"
                cols="22" ></textarea>
            
            <%If Session("ShowHidden") = "text" Then Response.Write "<br />"%>
            
            <input type="<%=Session("ShowHidden")%>" id="ReasonTxt<%=Count%>"  name ="ReasonTxt<%=Count%>" <%=Session("ShowHiddenSize")%> />
            <%If Session("ShowHidden") = "text" Then Response.Write "ReasonTxt&nbsp;"%>
            
            <input type="<%=Session("ShowHidden")%>" id="Sqm<%=Count%>"  name ="Sqm<%=Count%>" <%=Session("ShowHiddenSize")%>
            value="<%If DetailPartSqm > 0 Then
                         Response.Write Round(DetailPartSqm,2)
                         hSqm = hSqm + Round(DetailPartSqm,2)
                     Else
                         Response.Write "0"
                     End If 
                   %>"/>  <%If Session("ShowHidden") = "text" Then Response.Write "Sqm&nbsp;"%>
        </td>
        
    </tr>
        <%
        
        End If
        
        Combined = False
        
        DisplaySub = ""
        DetailPartQty = 0
        DetailSubPartQty = 0
        DetailSubPartSqm = 0
        DetailPartSqm = 0
        DetailSubPartItemCount = 0
        DetailPartUnitCost = 0
        RedoQty = 0
       
        ItemDept = ""
        ItemAlpha = ""
        ItemTitle = ""
        DetailPartCode = ""
        DetailPrinterName = ""
        DetailPartDescription = ""
        DetailPartSaveDescription = "" 
        SpItemsTrue = False        
          
    Next 
  
Erase DetailItemsParts
Erase AlphaItemsParts
Erase DetailItemsDept
Erase DetailItemsDeptID 
Erase DetailItemsQcDeptID    

DetailRS.Close
Set DetailRS = Nothing

%>
</table>
<br />

<table align="center" style="padding-right: 10px; padding-left: 10px;" cellpadding="0" cellspacing="0"  width="95%">
    <tr>
        <td align="left">
            &nbsp;&nbsp;<input type="submit" value="update" disabled="disabled" id="btnSubmit" />
            &nbsp;&nbsp;<input type="button" value="Reset" id="reset"  onclick="javascript:window.location.replace('IhrAddData.asp');"/>
            <%
            If TotalRows > 1 Then
            %>
                &nbsp;&nbsp;<label title="Selects all items, But you still have to enter data for each row" id="TickAll"  onmouseover="this.style.color='<%=MplMagenta%>'" onmouseout="this.style.color='<%=DarkGrey%>'" style="font-size: 12px; color: <%=DarkGrey%>; font-weight: bold;" onclick="javascript:ShowAllRows();" >Select All Items</label>
            <%
            End If      
            %>
              
        </td>
    </tr>
    
    <tr>
        <td>
            <br />
            <%If Session("ShowHidden") = "text" Then Response.Write "TotalRows" %>&nbsp;<input id="TotalRows" name="TotalRows" value="<%=TotalRows%>" type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "ActualRows" %>&nbsp;<input id="ActualRows" name="ActualRows" value="0" type="<%=Session("ShowHidden")%>"  />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "Quote Ref" %>&nbsp;<input id="QuoteRef" name="QuoteRef" value="<%=hQuoteRef%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "QuoteID" %>&nbsp;<input id="QuoteID" name="QuoteID" value="<%=DetailJob%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "Client" %>&nbsp;<input id="Client" name="Client" value="<%=Client%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "Desc" %>&nbsp;<input id="Desc" name="Desc" value="<%=hDesc%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "OrigSqm" %>&nbsp;<input id="OrigSqm" name="OrigSqm" value="<%=hSqm%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "DepAjax" %>&nbsp;<input id="ReasonTest" name="ReasonTest" value=""  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "BookedInBy" %>&nbsp;<input id='BookedInBy' name="BookedInBy" value="<%=BookedInBy%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "CreatedById" %>&nbsp;<input id="CreatedById" name="CreatedById" value="<%=Session("UserId")%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
        </td>
    </tr>      
</table>

</form>

</body>

</html>

<%

Sub GetItems (QuotationId) '## Passed Quote.ID to get QuotationItems.ID for each quote item

    Dim GetItemsSql
    Dim GetItemsRs
    
    DetailItems = ""
        
    GetItemsSql = "SELECT QuotationId, Id AS QuoteItemId, ItemId, RecState, PartType" 
    GetItemsSql = GetItemsSql & " FROM QuotationItems"
    GetItemsSql = GetItemsSql & " WHERE (QuotationItems.QuotationId = '" & Cstr(QuotationId) & "')"
    GetItemsSql = GetItemsSql & " AND (QuotationItems.RecState = 0) AND (PartType = 9)"
    GetItemsSql = GetItemsSql & " ORDER BY QuotationItems.ItemId"
  

    Set GetItemsRs = Server.CreateObject("ADODB.Recordset")
    GetItemsRs.ActiveConnection = Session("ClarityConn")  'strConnClarity
	GetItemsRs.Source = GetItemsSql 	
	GetItemsRs.CursorType = Application("adOpenForwardOnly")
	GetItemsRs.CursorLocation = Application("adUseClient")
	GetItemsRs.LockType = Application("adLockReadOnly")
	GetItemsRs.Open
	Set GetItemsRs.ActiveConnection = Nothing
	
	If GetItemsRs.BOF = True Or GetItemsRs.EOF = True Then
	    GetItemsRs.Close
        Set GetItemsRs = Nothing
        TotalRows = 0
	    Exit Sub
	End If
	
	
	TotalRows = GetItemsRs.RecordCount

    While not GetItemsRs.EOF
        If GetItemsRs("QuoteItemId") > 0 Then
            If DetailItems = "" Then
                DetailItems = GetItemsRs("QuoteItemId") &  ";"
                AlphaItems = GetItemsRs("ItemId") &  ";"
            Else
                DetailItems = DetailItems & GetItemsRs("QuoteItemId") & ";"
                AlphaItems = AlphaItems & GetItemsRs("ItemId") &  ";"
            End If
        End If
        GetItemsRs.MoveNext
    Wend
    
    GetItemsRs.Close
    Set GetItemsRs = Nothing

End Sub

'#############################################################################################

Sub GetDepartments (QuoteItemId) '## Passed QuotationItems.ID for each quote item

    Dim GetDepartmentsSql
    Dim GetDepartmentsRs
    
    ItemDept = ""
    
                                            
    GetDepartmentsSql = "SELECT QuotationItems.QuotationId, QuotationItems.Id AS QuoteItemId, QuotationItems.ItemId, ProdOperations.WorkCentreId,"
    GetDepartmentsSql = GetDepartmentsSql & " ProdOperations.StatusNumber, QuotationItems.PartType, QuotationItems.RecState," 
    GetDepartmentsSql = GetDepartmentsSql & " ProdJobCards.Description, ProdWorkCentres.Description AS WorkCentre, ProdOperations.GroupId," 
    GetDepartmentsSql = GetDepartmentsSql & " ProdWorkCentreGroups.Description AS GroupName, ProdOperations.OPNumber"
    GetDepartmentsSql = GetDepartmentsSql & " FROM ProdJobCards INNER JOIN"
    GetDepartmentsSql = GetDepartmentsSql & " QuotationItems ON ProdJobCards.QuoteItemId = QuotationItems.Id INNER JOIN"
    GetDepartmentsSql = GetDepartmentsSql & " ProdOperations ON ProdJobCards.Id = ProdOperations.ParentId INNER JOIN"
    GetDepartmentsSql = GetDepartmentsSql & " ProdWorkCentres ON ProdOperations.WorkCentreId = ProdWorkCentres.Id INNER JOIN"
    GetDepartmentsSql = GetDepartmentsSql & " ProdWorkCentreGroups ON ProdOperations.GroupId = ProdWorkCentreGroups.Id"
    GetDepartmentsSql = GetDepartmentsSql & " WHERE (QuotationItems.Id = " & QuoteItemId & ")" 
    GetDepartmentsSql = GetDepartmentsSql & " AND (ProdOperations.StatusNumber IN (0, 1, 2, 3)) AND (QuotationItems.RecState = 0)"
    GetDepartmentsSql = GetDepartmentsSql & " ORDER BY ProdOperations.OPNumber"

    Set GetDepartmentsRs = Server.CreateObject("ADODB.Recordset")
    GetDepartmentsRs.ActiveConnection = Session("ClarityConn")  'strConnClarity
	GetDepartmentsRs.Source = GetDepartmentsSql 	
	GetDepartmentsRs.CursorType = Application("adOpenForwardOnly")
	GetDepartmentsRs.CursorLocation = Application("adUseClient")
	GetDepartmentsRs.LockType = Application("adLockReadOnly")
	GetDepartmentsRs.Open
	Set GetDepartmentsRs.ActiveConnection = Nothing
	

    While Not GetDepartmentsRs.EOF
        If ItemDept = "" Then           
            If GetDepartmentsRs("GroupId") = 5 Then
                ItemDept = GetDepartmentsRs("GroupName") & " - "  &  GetDepartmentsRs("WorkCentre") & ";"
            ElseIf GetDepartmentsRs("GroupId") = 12 Then
                    ItemDept = "Laminate"
            Else
                    ItemDept = GetDepartmentsRs("GroupName") & ";"
            End If
            ItemDeptID = GetDepartmentsRs("GroupId") &  ";"
            ItemQcDeptId = GetQcGroup(GetDepartmentsRs("GroupId")) &  ";"
            ItemWcID = GetDepartmentsRs("WorkCentreId") &  ";"
         Else
            
            '## Check for duplicate may need for dulicate welding
            '## Sewing in twice on 244221 once for pockets & once for edges
            If Instr(1,ItemDept,"Sewing",1) > 0 And GetDepartmentsRs("WorkCentre") = "Sewing" Then
                '## Already in so don't duplicate
            Else
            
                If GetDepartmentsRs("GroupId") = 5 Then
                    ItemDept = ItemDept & GetDepartmentsRs("GroupName") & " - "  &  GetDepartmentsRs("WorkCentre") & ";"
                ElseIf GetDepartmentsRs("GroupId") = 12 Then
                    ItemDept = ItemDept &  "Laminate" & ";"              
                Else
                    ItemDept = ItemDept & GetDepartmentsRs("GroupName") & ";"
                End If
                
                ItemDeptID = ItemDeptID & GetDepartmentsRs("GroupId") &  ";"
                ItemQcDeptId = ItemQcDeptId & GetQcGroup(GetDepartmentsRs("GroupId")) &  ";"
                ItemWcID = ItemWcID & GetDepartmentsRs("WorkCentreId") &  ";"  
            End If               
        End If
                    
                    
                    
        GetDepartmentsRs.MoveNext
    Wend
    
    GetDepartmentsRs.Close
    Set GetDepartmentsRs = Nothing

End Sub

'#############################################################################################

Function GetQcGroup(ClarityId)

    Dim LocalId

    Select Case ClarityId        
        
        Case 3
            LocalId = 4 '## All Cutting is under Finishing        
        Case 4
            LocalId = 1 '## Prepress
        Case 7
            LocalId = 2 '## Heat Press, comes under Printing
        Case 10
            LocalId = 6 '## Installation
        Case 11
            LocalId = 5 '## Build 
        Case Else
            LocalId = ClarityId   
    
    End Select

    GetQcGroup = LocalId

End Function

'#############################################################################################

Sub CalculateDetailSqm (QuoteItemID) '## Passed QuotationItems.Id 
 
    Dim strDetailSqmSql
    Dim DetailSqmRs
    Dim SqmStartPos
    Dim SqmTempNotes
    
    SqmStartPos = 0
	SqmTempNotes = ""
	DetailMaterialCount = 0
	DetailPartCode = ""
	DetailPartDescription = ""
	DetailPartUnitCost = 0
    
	strDetailSqmSql = "Select QuotationItems.Quantity, QuotationItems.SizeX, QuotationItems.SizeY, QuotationItems.ItemId,"
    strDetailSqmSql = strDetailSqmSql & " QuotationItems.Id, QuotationItems.Description, QuotationItems.Partcode,"
    strDetailSqmSql = strDetailSqmSql & " QuotationItems.UnitCost, QuotationItems.Notes, QuotationItems.RecState From QuotationItems"
    strDetailSqmSql = strDetailSqmSql & " Where(QuotationItems.Id = '" & Cstr(QuoteItemID) & "')"
    strDetailSqmSql = strDetailSqmSql & "  AND (QuotationItems.RecState = 0)" 
 
    Set DetailSqmRs = Server.CreateObject("ADODB.Recordset")

    DetailSqmRs.ActiveConnection = Session("ClarityConn")  'strConnClarity
	DetailSqmRs.Source = strDetailSqmSql 	
	DetailSqmRs.CursorType = Application("adOpenForwardOnly")
	DetailSqmRs.CursorLocation = Application("adUseClient")
	DetailSqmRs.LockType = Application("adLockReadOnly")
	DetailSqmRs.Open
	Set DetailSqmRs.ActiveConnection = Nothing
	
	DetailPartSqm = 0
	DetailPartQty = 0
	
	If DetailSqmRs.EOF or DetailSqmRs.BOF Then
	    DetailSqmRs.Close
	    Set DetailSqmRs = Nothing
	    Exit Sub	    
	End If
	    
    If DetailSqmRs("SizeX") > 0 then	        
        Material DetailSqmRs("Id") '## populate substrate
        DetailPartSqm = DetailSqmRs("Quantity") * ((DetailSqmRs("SizeX") * DetailSqmRs("SizeY")) / 1000000) 
        DetailPartSqm = Round(DetailPartSqm, 2)
        DetailPartQty = DetailSqmRs("Quantity")
                
        If InStr(1, DetailSqmRs("Notes"), "Double Sided: ", 1) > 0 Then
            SqmStartPos = InStr(1, DetailSqmRs("Notes"), "Double Sided: ", 1)                
            SqmTempNotes = Mid(DetailSqmRs("Notes"), SqmStartPos, 17)
            
            If Right(SqmTempNotes, 3) = "Yes" Then
                DetailPartSqm = DetailPartSqm * 2
            End If
        End If	
    End If	    
    
    If DetailPartSqm = 0 Then
        '## Check DetailItems Detail to see if it's a price wizard made of calc wizards
        '## Pass QuotationItems.Id & Qty
        CalculateDetailSqmSubItems DetailSqmRs("Id"), DetailSqmRs("Quantity")
        If DetailPartSqm > 0 Then DetailPartQty = DetailSqmRs("Quantity")                
    End If	    
    
    DetailPartCode = DetailSqmRs("PartCode")
    
    DetailPartDescription = Trim(DetailSqmRs("Description"))
    
    DetailPartUnitCost = DetailSqmRs("UnitCost") 
    
   
    
    LocalSubstrate = MaterialList
    MaterialList = ""
	    
	'## Strip out ; from list
'	If MaterialList <> "" Then
'	    LocalSubstrate = Split(MaterialList,"#",-1,1)
'	    MaterialList = ""
'	Else
'	    Redim LocalSubstrate(0)
'	    LocalSubstrate(0) = "N/A"
	
'	End If

	DetailSqmRs.close
	Set DetailSqmRs = Nothing
	
	Err.Clear
 
 End Sub
 
'#############################################################################################
 
 Sub CalculateDetailSqmSubItems (DetailItemsId, ItemQty)  '## Passed QuotationItems.Id & Qty
 
    Dim DetailSubSqmSql 
    Dim DetailSubSqmSqlRS
        
    DetailSubPartQty = 0
    DetailSubPartSqm = 0
    
    DetailSubSqmSql = "SELECT  QuotationItemDetail.QuotationItemId, QuotationItemDetail.SizeX,QuotationItemDetail.SizeY, "
    DetailSubSqmSql = DetailSubSqmSql & " QuotationItemDetail.PartCode, QuotationItemDetail.Quantity, QuotationItemDetail.Id"
    DetailSubSqmSql = DetailSubSqmSql & " FROM QuotationItemDetail"
    DetailSubSqmSql = DetailSubSqmSql & " WHERE (QuotationItemDetail.QuotationItemId = " & DetailItemsId & ")"
    
    DetailSubSqmSql = DetailSubSqmSql & " AND (ScalingType = 1)  AND (PartType IN(8,9))"
 
    Set DetailSubSqmSqlRS = Server.CreateObject("ADODB.Recordset")

    DetailSubSqmSqlRS.ActiveConnection = Session("ClarityConn")  'strConnClarity
	DetailSubSqmSqlRS.Source = DetailSubSqmSql 	
	DetailSubSqmSqlRS.CursorType = Application("adOpenForwardOnly")
	DetailSubSqmSqlRS.CursorLocation = Application("adUseClient")
	DetailSubSqmSqlRS.LockType = Application("adLockReadOnly")
	DetailSubSqmSqlRS.Open
	Set DetailSubSqmSqlRS.ActiveConnection = Nothing
		
	If DetailSubSqmSqlRS.EOF Or DetailSubSqmSqlRS.BOF Then
	    DetailSubPartQty = 0 
	    DetailSubSqmSqlRS.Close
	    Set DetailSubSqmSqlRS = Nothing
	    Exit Sub
	End If
       
    If DetailSubSqmSqlRS("Quantity") <> 0 Then        
        While Not DetailSubSqmSqlRS.EOF
            If DetailSubSqmSqlRS("SizeX") <> 0 Then
                Material DetailSubSqmSqlRS("QuotationItemId") '## populate substrate             
                DetailSubPartSqm = DetailSubPartSqm + DetailSubSqmSqlRS("Quantity") * ((DetailSubSqmSqlRS("SizeX") * DetailSubSqmSqlRS("SizeY")) / 1000000)
                DetailSubPartQty = DetailSubPartQty + DetailSubSqmSqlRS("Quantity")
                DetailSubPartItemCount = DetailSubPartItemCount + 1
                DetailSubPartSqm = Round(DetailSubPartSqm,2)
            End If
            DetailSubSqmSqlRS.MoveNext                
        Wend        
    Else
        DetailSubPartSqm = 0
    End If    
    
    DetailPartSqm = DetailSubPartSqm * ItemQty
   
    DetailSubSqmSqlRS.Close
    Set DetailSubSqmSqlRS = Nothing
    
    Err.Clear
 
End Sub

'#############################################################################################
 
Private Sub Material (MatItemID)    '## Passed QuotationItems.Id 

'## On Error Resume Next


    '## Using this in main report
    
    '## This works ok but can't get Stock Requisitioned
   ' MaterialSql = "SELECT Part.Description, Part.SizeY, PartCategory.Name"
   ' MaterialSql = MaterialSql & " FROM  QuotationItemDetail INNER JOIN"
  '  MaterialSql = MaterialSql & " Part ON QuotationItemDetail.TSLId = Part.TSLId INNER JOIN"
 '   MaterialSql = MaterialSql & " PartCategory ON Part.PartCategoryId = PartCategory.Id"
        
    '## If add this can still have GlobalPartList.StockRequisitioned
    '## INNER JOIN GlobalPartList ON QuotationItemDetail.TSLId = GlobalPartList.TSLId
    '## But slows it all down 
    
 '   MaterialSql = MaterialSql & " WHERE (QuotationItemDetail.QuotationItemId = " & MatItemID & ")"
'    MaterialSql = MaterialSql & " AND (QuotationItemDetail.PartType IN (1, 5))" 
 '   MaterialSql = MaterialSql & " AND (QuotationItemDetail.PriceListId = 4)"    '## 11 = Redundant
 '   MaterialSql = MaterialSql & " AND (Part.SizeY > 150)" 
  '  MaterialSql = MaterialSql & " AND (NOT (PartCategory.Name IN (N'Consumables', N'Finishing', N'Laminate')))"
 
    Dim MaterialSql
    Dim MaterialRS
     
    MaterialSql = "SELECT QuotationItemDetail.QuotationItemId, QuotationItemDetail.PartCode, GlobalPartList.Description,"
    MaterialSql = MaterialSql & " QuotationItemDetail.PartType, QuotationItemDetail.SizeY,GlobalPartList.Level3"
    MaterialSql = MaterialSql & " FROM QuotationItemDetail INNER JOIN"
    MaterialSql = MaterialSql & " GlobalPartList ON QuotationItemDetail.PartCode = GlobalPartList.Partcode"
    MaterialSql = MaterialSql & " WHERE (QuotationItemDetail.QuotationItemId = " & MatItemID & ")"
   
    MaterialSql = MaterialSql & " AND (GlobalPartList.Level3 IN ('Fabric', 'PVC', 'Rigid', 'SAV', 'Film', 'Pull-up', 'Paper'))"
    MaterialSql = MaterialSql & " AND (QuotationItemDetail.PriceListId = 4)" 
    
    'MaterialSql = MaterialSql & " AND (GlobalPartList.Level3 NOT IN ('Consumables'))"
                                    '## was PartType = 1) material, now material or non stock = 5
    MaterialSql = MaterialSql & " AND (QuotationItemDetail.PartType IN (1,5))"
    MaterialSql = MaterialSql & "  AND (QuotationItemDetail.PartCode <> '<Unknown>')"
    
    '##AND (QuotationItemDetail.SizeY > 150)
    
    Set MaterialRS = Server.CreateObject("ADODB.Recordset")

    MaterialRS.ActiveConnection = Session("ClarityConn")  'strConnClarity
	MaterialRS.Source = MaterialSql	
	MaterialRS.CursorType = Application("adOpenForwardOnly")
	MaterialRS.CursorLocation = Application("adUseClient")
	MaterialRS.LockType = Application("adLockReadOnly")
	MaterialRS.Open
	Set MaterialRS.ActiveConnection = Nothing
	
	If MaterialRS.EOF Or MaterialRS.BOF Then
	    MaterialRS.Close
        Set MaterialRS = Nothing
        Exit Sub
    End If
        
    If MaterialRS.RecordCount > 1 then    
        While Not MaterialRS.EOF        
            If MaterialList = "" Then
                MaterialList = MaterialRS("Description") 
            Else
                MaterialList = MaterialList & "#" & MaterialRS("Description") 
            End If        
            MaterialRS.MoveNext        
        Wend
    Else 
        If MaterialList = "" Then
            MaterialList = MaterialRS("Description") 
        Else
            MaterialList = MaterialList & "#" &  MaterialRS("Description") 
        End If 
    End If     
    
    MaterialRS.Close
    Set MaterialRS = Nothing
    
    Err.Clear      
 
End Sub

'#############################################################################################
 
 Private Function GetPrinterName (QItemId) 
 
    Dim PrinterSql
    Dim PrinterRs
    Dim LocalPrinter
    
    PrinterSql = "SELECT ProdOperations.WorkCentreId, ProdWorkCentres.Description, ProdJobCards.QuoteItemId, ProdWorkCentres.WorkCentreGroupId"
    PrinterSql = PrinterSql & " FROM ProdOperations INNER JOIN"
    PrinterSql = PrinterSql & " ProdJobCards ON ProdOperations.ParentId = ProdJobCards.Id INNER JOIN"
    PrinterSql = PrinterSql & " ProdWorkCentres ON ProdOperations.WorkCentreId = ProdWorkCentres.Id"
    PrinterSql = PrinterSql & " WHERE (ProdJobCards.QuoteItemId =  " & QItemId & ")"
    PrinterSql = PrinterSql & " AND (ProdWorkCentres.WorkCentreGroupId = " & PrintGroupId & ")"
    
    Set PrinterRs = Server.CreateObject("ADODB.Recordset")

    PrinterRs.ActiveConnection = Session("ClarityConn")  'strConnClarity
	PrinterRs.Source = PrinterSql	
	PrinterRs.CursorType = Application("adOpenForwardOnly")
	PrinterRs.CursorLocation = Application("adUseClient")
	PrinterRs.LockType = Application("adLockReadOnly")
	PrinterRs.Open
	Set PrinterRs.ActiveConnection = Nothing	
	
	If PrinterRs.BOF = True Or PrinterRs.EOF = True then
	    LocalPrinter = ""	    
	Else
	    LocalPrinter = PrinterRs("Description")
	End If
	
	'Set PrinterRs.ActiveConnection = Nothing
	PrinterRs.Close
	Set PrinterRs = Nothing 
	
	GetPrinterName = LocalPrinter   
 
End Function

'#############################################################################################

Private Function GetMaterialData (QItemId)

    Dim MaterialDataSql
    Dim MaterialDataRs
    Dim LocalResult
    
    Dim MaterialDataCost
    Dim LocalGroup
    Dim LocalLevel2
    
    MaterialDataCost = CCur(0)
    LocalGroup = ""
    LocalLevel2 = ""
    
    MaterialDataSql = "SELECT QuotationItemDetail.QuotationItemId, QuotationItemDetail.Description, QuotationItemDetail.TSLId,"
    MaterialDataSql = MaterialDataSql & " Round(QuotationItemDetail.UnitCost,2) AS UnitCost,QuotationItemDetail.PartCode,"
    MaterialDataSql = MaterialDataSql & " CASE WHEN vwFlatUser.User11 IS NULL THEN 'No Group' ELSE vwFlatUser.User11 END AS GroupName,"
    MaterialDataSql = MaterialDataSql & " Round(QuotationItemDetail.TotalCost,2) AS TotalCost, GlobalPartList.Level2"

    MaterialDataSql = MaterialDataSql & " FROM QuotationItemDetail INNER JOIN"
    MaterialDataSql = MaterialDataSql & " GlobalPartList ON QuotationItemDetail.TSLId = GlobalPartList.TSLId INNER JOIN"
    MaterialDataSql = MaterialDataSql & " vwPartAttributesAsUserFields AS vwFlatUser ON vwFlatUser.PartId = GlobalPartList.Id"
    MaterialDataSql = MaterialDataSql & " WHERE (QuotationItemDetail.QuotationItemId = " & QItemId & ")"
    MaterialDataSql = MaterialDataSql & " AND (GlobalPartList.Level1 = 'Supplier Price List')"
    MaterialDataSql = MaterialDataSql & " AND (GlobalPartList.Level3 IN ('Fabric', 'PVC', 'Rigid', 'SAV', 'Film', 'Pull-up', 'Paper'))"


    Set MaterialDataRs = Server.CreateObject("ADODB.Recordset")

    MaterialDataRs.ActiveConnection = Session("ClarityConn")  'strConnClarity
	MaterialDataRs.Source = MaterialDataSql	
	MaterialDataRs.CursorType = Application("adOpenForwardOnly")
	MaterialDataRs.CursorLocation = Application("adUseClient")
	MaterialDataRs.LockType = Application("adLockReadOnly")
	MaterialDataRs.Open
	Set MaterialDataRs.ActiveConnection = Nothing	
	
	If MaterialDataRs.BOF = True Or MaterialDataRs.EOF = True then
	    LocalResult = "0#No Group# "	    
	Else
	    LocalGroup = MaterialDataRs("GroupName")
	    LocalLevel2 = MaterialDataRs("Level2")
	    
	    While Not MaterialDataRs.EOF
	        MaterialDataCost =  CCur(MaterialDataCost) + CCur(MaterialDataRs("TotalCost"))
	        
	        '## If do this adds twice
	        'LocalGroup = LocalGroup & MaterialDataRs("GroupName")
	        'LocalLevel2 = LocalLevel2 & MaterialDataRs("Level2")
	        
	        
	    'LocalResult = Cstr(Round(MaterialDataRs("TotalCost"),2)) & "#" & MaterialDataRs("GroupName") & "#" & MaterialDataRs("Level2")    
	        MaterialDataRs.MoveNext
	    Wend
	    LocalResult = Cstr(MaterialDataCost) & "#" & LocalGroup & "#" & LocalLevel2
	   ' LocalResult = Cstr(Round(MaterialDataRs("TotalCost"),2)) & "#" & MaterialDataRs("GroupName") & "#" & MaterialDataRs("Level2")
	End If	
	
	MaterialDataRs.Close
	Set MaterialDataRs = Nothing 	
	
	'LocalResult = Cstr(Round(LocalCost,2)) & "#" & MaterialDataRs("GroupName") & "#" & MaterialDataRs("Level2")
	
    GetMaterialData = LocalResult

End Function

'#############################################################################################

Sub DelAddress (QuoteId)   '## Passed Quote Id also gets Client Name
 
    Dim strAddressSql
    Dim AddressRs
    Dim LocalCompanyAddress
    Dim LocalDeliveryAddress 
    
    DeliveryAddress = ""  
 
    strAddressSql = "SELECT  Quotation.DeliveryContact, Quotation.DeliveryCompany, Quotation.DeliveryAddress, Quotation.DeliveryAddress2,"
    strAddressSql = strAddressSql & " Quotation.DeliveryAddress3, Quotation.DeliveryCity, Quotation.DeliveryCounty, "
    strAddressSql = strAddressSql & " Quotation.DeliveryPostcode, Quotation.contactId, Contacts.CompanyId, Company.Id,"
    strAddressSql = strAddressSql & " Company.Name, Quotation.id, Contacts.Address, Contacts.Address2, Contacts.Address3,"
    strAddressSql = strAddressSql & " Contacts.City, Contacts.County, Contacts.Postcode"
    strAddressSql = strAddressSql & " FROM     Contacts INNER JOIN"
    strAddressSql = strAddressSql & " Quotation ON Contacts.Id = Quotation.contactId INNER JOIN"
    strAddressSql = strAddressSql & " Company ON Contacts.CompanyId = Company.Id"
    strAddressSql = strAddressSql & " Where (Quotation.id = '" & QuoteId & "')" 
    
    Set AddressRs = Server.CreateObject("ADODB.Recordset")

    AddressRs.ActiveConnection = Session("ClarityConn")  'strConnClarity
	AddressRs.Source = strAddressSql	
	AddressRs.CursorType = Application("adOpenForwardOnly")
	AddressRs.CursorLocation = Application("adUseClient")
	AddressRs.LockType = Application("adLockReadOnly")
	AddressRs.Open
	Set AddressRs.ActiveConnection = Nothing
	
	If AddressRs.EOF = True Or AddressRs.BOF = True Then
	    AddressRs.Close
	    Set AddressRs = Nothing
	    Exit Sub
	End if
	
	Client = AddressRs("Name")
	
    
    AddressRs.Close
	Set AddressRs = Nothing 
	
 
End Sub

'#############################################################################################
 
Private Function GetOpStatus (OpItemId) 
 
    Dim OpStatusRs
    Dim OpStatusSql 
    Dim LocalResult
    LocalResult = False
    
    OpStatusSql = "SELECT ProdJobCards.QuoteItemId, ProdOperations.ParentId, ProdOperations.StatusNumber AS ProdOpStatus,"
    OpStatusSql = OpStatusSql & " ProdOperations.WorkCentreId, ProdOperations.OPNumber"
    OpStatusSql = OpStatusSql & " FROM ProdOperations INNER JOIN"
    OpStatusSql = OpStatusSql & " ProdJobCards ON ProdOperations.ParentId = ProdJobCards.Id"
    OpStatusSql = OpStatusSql & " WHERE (ProdJobCards.QuoteItemId = " & OpItemId & ")"
    OpStatusSql = OpStatusSql & " AND (ProdOperations.StatusNumber = 2)"
    OpStatusSql = OpStatusSql & " ORDER BY ProdOperations.OPNumber"
    
    Set OpStatusRs = Server.CreateObject("ADODB.Recordset")

    OpStatusRs.ActiveConnection = Session("ClarityConn")  
	OpStatusRs.Source = OpStatusSql	
	OpStatusRs.CursorType = Application("adOpenForwardOnly")
	OpStatusRs.CursorLocation = Application("adUseClient")
	OpStatusRs.LockType = Application("adLockReadOnly")
	OpStatusRs.Open
	Set OpStatusRs.ActiveConnection = Nothing
	
	If OpStatusRs.BOF =  True  Or OpStatusRs.EOF = True Then
	    LocalResult = False
    Else
        LocalResult = True
    End If
    
    OpStatusRs.Close
    Set OpStatusRs = Nothing
    
    '## overide for me
    If Session("UserId") = 4 Then LocalResult = True
    
    GetOpStatus = LocalResult

End function
%>