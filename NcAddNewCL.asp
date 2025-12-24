<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>


<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\Declarations.asp" -->
<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\PkId.asp" -->


<% 
 On Error Resume Next

'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")


Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.asp"

Dim DetailRS
Dim strDetailSql
Dim DetailJob
Dim DetailItems
Dim DetailItemsParts
Dim DetailPartSqm
Dim DetailSubPartSqm
Dim DetailOperationStatus
Dim DetailPartDescription
Dim LocalSubstrate
Dim Fabrication(3,1)
Dim FabDummy
Dim ProdNotes
Dim DisplayStatus

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
Dim ItemAlpha
Dim Combined
Dim JobIsOutdoor

Dim AlphaItemsParts
Dim AlphaItems

Dim ReqDelHeader
Dim ItemDescription
Dim DetailMaterialCount

Dim MgsColSpan
Dim hQuoteId
Dim hJobId
Dim hJobRef

DetailJob = Session("JobNo")

'Not Distinct  Can't use if showing notes 
strDetailSql = "SELECT Job.Reference AS JobRef, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId, Quotation.user11 As ArtworkStatus,"
strDetailSql = strDetailSql & " Quotation.Description AS JobDesc, Quotation.Notes, Job.Id as JobId, Quotation.User6,"
strDetailSql = strDetailSql & " Quotation.status AS QuoteStatus, Quotation.JobTypeId, Quotation.completedDate AS DelDate, Contacts.Forename + ' ' + Contacts.Surname AS BookedBy"
strDetailSql = strDetailSql & " FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId INNER JOIN Contacts ON Job.CreatedBy = Contacts.Id"
strDetailSql = strDetailSql & " Where(Quotation.id = '" & DetailJob & "')"
 
Set DetailRS = Server.CreateObject("ADODB.Recordset")

DetailRS.ActiveConnection = strConnClarity
DetailRS.Source = strDetailSql
DetailRS.CursorType = 0
DetailRS.CursorLocation = 3
DetailRS.LockType = 1
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


hQuoteId = DetailRS("QuoteId")
hJobId = DetailRS("JobId")
hJobRef = Trim(DetailRS("JobRef"))

'## Get Address details and Company Name
DelAddress DetailRS("QuoteId")

If DetailRS("Notes") <> "" Then
    ProdNotes = Trim(Cstr(DetailRS("Notes")))
    If Left(ProdNotes,2) = VbCrLf Then ProdNotes = Mid(ProdNotes,2)
    ProdNotes = Replace(ProdNotes, VbCrLf, "<br />",1,-1,1)
    ProdNotes = Replace(ProdNotes, "<br /><br />", "<br />",1,-1,1)  
Else    
    ProdNotes = ""   
End If
If Left(ProdNotes,6) = "<br />" Then ProdNotes = Mid(ProdNotes,7)

'## Remove any extra <br />
If Right(ProdNotes,6) = "<br />" Then
    Do Until Right(ProdNotes,6) <> "<br />"
        If Right(ProdNotes,6) = "<br />" Then
            ProdNotes = Left(ProdNotes,Len(ProdNotes)-6)
        Else
            Exit Do
        End If
    Loop
End If

Dim DeptSql
Dim DeptRs

DeptSql = "SELECT DISTINCT vwProdGantt.GroupName, vwProdGantt.WorkCentre, vwProdGantt.GroupId, vwProdGantt.WorkCentreId,"
DeptSql = DeptSql & " ProdWorkCentreGroups.[Order] AS BolOrder FROM vwProdGantt INNER JOIN"
DeptSql = DeptSql & " ProdWorkCentreGroups ON vwProdGantt.GroupId = ProdWorkCentreGroups.Id"
DeptSql = DeptSql & " WHERE (JobId = " & hJobId  & ") ORDER BY BolOrder, vwProdGantt.WorkCentreId"

Set DeptRs = Server.CreateObject("ADODB.Recordset")

DeptRs.ActiveConnection = strConnClarity
DeptRs.Source = DeptSql
DeptRs.CursorType = 0
DeptRs.CursorLocation = 3
DeptRs.LockType = 1
DeptRs.Open

If DeptRs.BOF = True Or DeptRs.EOF = True Then
    DeptRs.Close
    Set DeptRs = Nothing
    Err.Clear
    Response.Redirect "Error.asp"
End If

Dim DeptArray()
ReDim DeptArray(DeptRs.RecordCount,1)

Dim DeptCount

While Not DeptRs.EOF
    
        If DeptRs("GroupId") < 10 Then
            DeptArray(DeptRs.AbsolutePosition,0) = "0" & DeptRs("GroupId") '## & "#" & DeptRs("WorkCentreId")  'DeptRs("GroupId") & "#" &
            If DeptRs("WorkCentreId") < 10 Then
                DeptArray(DeptRs.AbsolutePosition,0) = DeptArray(DeptRs.AbsolutePosition,0) & "#" & "0" & DeptRs("WorkCentreId")
            Else
                DeptArray(DeptRs.AbsolutePosition,0) = DeptArray(DeptRs.AbsolutePosition,0) & "#" & DeptRs("WorkCentreId")
            End If
        Else           
            DeptArray(DeptRs.AbsolutePosition,0) = DeptRs("GroupId")'## & "#" & DeptRs("WorkCentreId")  'DeptRs("GroupId") & "#" & 
            If DeptRs("WorkCentreId") < 10 Then
                DeptArray(DeptRs.AbsolutePosition,0) = DeptArray(DeptRs.AbsolutePosition,0) & "#" & "0" & DeptRs("WorkCentreId")
            Else
                DeptArray(DeptRs.AbsolutePosition,0) = DeptArray(DeptRs.AbsolutePosition,0) & "#" & DeptRs("WorkCentreId")
            End If
        
        End If
        DeptArray(DeptRs.AbsolutePosition,1) = DeptRs("WorkCentre")     'DeptRs("GroupName")  & " - " &    

    DeptRs.MoveNext
Wend

DeptRs.Close
Set DeptRs = Nothing

Dim TextType
Dim InputSize
If Session("ShowText") = Cbool(True) Then
    TextType = "text"
    InputSize = "size = '4'"
Else
    TextType = "hidden"
    InputSize = ""
End If

%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="en" lang="en">
<head>

    <title>NC Reporting job <%=DetailRS("JobRef")%></title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <link rel="shortcut icon" href="Images/warning.png" type="image/x-icon" />
    <link href="CSS/ClarityCss.css" rel="stylesheet" type="text/css" />
    <link href="CSS/ClarityExtraCss.css" rel="stylesheet" type="text/css" />
    <script type="text/javascript" src="JsFiles/NCReportsJSFunc.js"></script>
    <script type="text/javascript" src="JsFiles/NCReportsAjaxFunc.js"></script>
</head>


<body style="padding: 0px; margin: 0px">

<table style="padding-right: 10px; padding-left: 10px; width: 100%;" >
	<tr>
	    
		<td align="left" valign="bottom" height="100" colspan="3">
            <img align="left" alt="mediaco logo" src='<%=CompanyLogo%>' width="160" />
        </td> 
	    
    </tr>	
	<tr>
		<td height="8" valign="top" align="left" width="100%" colspan="3">
            <hr style="border-style: none;  height: 4px; background-color: <%=NewCyan%>; display: block;" />
        </td>
	</tr>
<!--</table>

<table style="padding-right: 10px; padding-left: 10px; width: 100%;">-->
<tr>
        <td align="left" width="33%" valign="top">
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('NcJobNo.asp');" style="font-size:12px; color: <%=NewBlue%>;">Return to Job Search Page</a><br /><br />
        </td> 
        <td align="center" width="34%">
            <img align="top" alt="mediaco logo" src="Images/warning.png" style="width: 20px; height: 20px;" />
            <font style="color: #0069AA; font-weight: bold; font-size: 16px;">New&nbsp;Log</font>
        </td>          
        <td align="right" width="33%" valign="top"><a id="logoffR" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>; ">Log&nbsp;Out</a>&nbsp;&nbsp;</td>        
    </tr>
</table>

<br />

<form method="post" action="NcUpdateLog.asp"  name="frmAddLog" id="frmAddLog" onsubmit="return NCValidateLogCL('Add');"> <!---->

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
    <th width="80" class="styleTHstd" >Delivery Address</th>
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
                                                    
    <td class="styleTDstd" width="80" style="text-align: left; padding-left: 3px; padding-top: 3px; padding-bottom: 3px;" rowspan="3">
        <%=DeliveryAddress %><!-- If no delivery address shows Client Address -->
    </td>
</tr>
<tr>
    <th style='border-style: none solid solid solid; border-width: 1px; border-color: #000000; color:#0069AA;' colspan='<%=MgsColSpan%>' >Production Notes</th>    
</tr>
<tr>
    <td class="styleTDleft" id='ProdNotes' colspan='<%=MgsColSpan%>' 
    style='text-align: left; padding-left: 3px; padding-top: 3px; height: 50px; color: #FF0000;'><%=ProdNotes%></td>    
</tr>


</table>

<br />

<table  style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">	
    <tr>
        <td colspan="3" nowrap="nowrap">
            <label style="font-weight: bold; color: <%=NewBlue%>;">Making a selection in a dropdown, enables 
            the next one. The final one enables notes, you can multi select in all</label>
        </td>
    </tr>
    
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
	    <th align="left" width="20%">Item</th>
	    <th align="left" width="40%">Selected</th>
	    <th align="left" width="40%" >&nbsp;</th> 
    </tr>
    
    
    

    <!-- do loop -->     
    <%
       
    GetItems DetailRS("QuoteId")
    DetailItemsParts = Split(DetailItems,";",-1,1)
    DetailItems = ""
    
    AlphaItemsParts = Split(AlphaItems,";",-1,1)    
    AlphaItems = ""
           
    %>   
     <tr >                                               
        <td  align="left">        
            <select id="cboItem" onchange="javascript:AddItem();" >
                <option value="">Select Item</option>
            <%
            
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
                        ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)-40)) 
                        ItemAlpha = "D" & ItemAlpha
                    ElseIf AlphaItemsParts(Count) >= 131 And AlphaItemsParts(Count) <= 156 Then
                        ItemAlpha = Cstr(Chr(AlphaItemsParts(Count)-66)) 
                        ItemAlpha = "D" & ItemAlpha
                    Else
                        ItemAlpha = "?"  
                    End If                
                    
                    If AlphaItemsParts(Count) < 10 Then AlphaItemsParts(Count) = 0 & AlphaItemsParts(Count)                  
                    Response.Write "<option value='" & AlphaItemsParts(Count) & "'>" & ItemAlpha & "</option>" & VbCrLf
                
                End If
            Next 
            %>

            </select>             
        </td>
        
        <td align="left" colspan="2">
            <input type="text" id="SelectedItem" value = "" 
                style="border-style: none; display: inline; width: 98%; font-size: 14px; color: #009933"/>
        </td>     
    </tr>
        
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" width="20%">Work centre</th>
        <th align="left" width="40%">Selected</th>
        <th align="left" width="40%" >&nbsp;</th>	        
    </tr>
    
    <tr>
        <td align="left">
            <select id="cboDeptSelect" onchange="javascript:AddDept();"  disabled="disabled">
                 <option value="">Select Dept</option>                 
                <%
                    If JobIsOutdoor = False Then
                        Response.Write "<option value='100#100'>Sales - CS</option>" & VbCrLf
                    End If
                
                    For DeptCount = 1 to Ubound(DeptArray)                    
                        Response.Write "<option value='" & DeptArray(DeptCount,0) & "'>" & DeptArray(DeptCount,1) & "</option>" & VbCrLf
                    Next
                    
                    Response.Write "<option value='101#101'>Stock</option>" & VbCrLf
                    Response.Write "<option value='0#0'>Other</option>"
                    
                    Erase DeptArray    
                %>
            </select>
        
        </td>
        <td align="left" colspan="2">        
            <input type="text" id="SelectedDept" value = ""
                style="border-style: none; width: 98%; display: inline; font-size: 14px; color: #009933"/>
        </td>        
    </tr> 
    
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>
    
    <tr >		 
        <th align="left" width="20%">Issue/Problem</th>
        <th align="left" width="40%">Selected</th>
        <th align="left" width="40%" >&nbsp;</th>	        
    </tr>
    
    <tr>
        <td align="left">
            <select id="cboReason" onchange="javascript:AddReason();" disabled="disabled">
                 <option value="">Select Issue/Problem</option>
                 <!--<option value="1">something</option> 
                 <option value="2">something else</option> -->          
            </select>        
        </td>
        <td align="left" colspan="2">
            <input type="text" id="SelectedReason" value = ""
                style="border-style: none; width: 98%; display: inline; font-size: 14px; color: #009933"/>        
        </td>        
    </tr>        
    
    <tr>        
        <td colspan="3" align="left">&nbsp;</td>
    </tr>    
    
    
    <tr>       
    <th  colspan="3" align="left">Notes</th>
    </tr>        
    
    <tr>
        <td colspan="3">
            <textarea  rows="15" id="txtDetails" name="txtDetails" style="text-align: left; " cols="110"  onkeyup="javascript:EnableSubmit();" disabled="disabled"></textarea>
        </td>
    </tr>
    
        
        
        <%

  
Erase DetailItemsParts
Erase AlphaItemsParts     

DetailRS.Close
Set DetailRS = Nothing

%>
        
        
</table>

<br />



<table style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">
    <tr>
        
        <td>
            <input type="submit" value="Update" disabled="disabled" id="btnSubmit"/>        
            &nbsp;&nbsp;&nbsp;
            <input name="btnReset" id="btnReset" type="reset" value="Reset" onclick="javascript:ResetPage();"/>
            
            <%
            If Session("PC-Name") = "Home" or Session("PC-Name") = "Work" Then
                If TextType = "text" Then
                    Response.Write ("&nbsp;Don't Save Data&nbsp;&nbsp;")
                    Response.Write ("<input type='checkbox' id='chkUpdate' checked='checked' name='chkUpdate'/>")
                End If                    
            End If          
            %>   
            
            <br /><br />
            <%If TextType = "text" Then Response.Write "hQuoteId"%>
            <input type="<%=TextType%>" id="hQuoteId" name="hQuoteId" value="<%=hQuoteId%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hJobId"%>           
            <input type="<%=TextType%>" id="hJobId" name="hJobId" value="<%=hJobId%>" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hItemId"%> 
            <input type="<%=TextType%>" id="hItemId" name="hItemId" value="" <%=InputSize %> />            
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hItemsAlpha"%> 
            <input type="<%=TextType%>" id="hItemsAlpha" name="hItemsAlpha" value="" <%=InputSize %> />                        
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hGroup"%> 
            <input type="<%=TextType%>" id="hGroup" name="hGroup" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hGroupSelected"%> 
            <input type="<%=TextType%>" id="hGroupSelected" name="hGroupSelected" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hDept"%> 
            <input type="<%=TextType%>" id="hDept" name="hDept" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hDeptSelected"%> 
            <input type="<%=TextType%>" id="hDeptSelected" name="hDeptSelected" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hReason"%> 
            <input type="<%=TextType%>" id="hReason" name="hReason" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hReasonSelected"%> 
            <input type="<%=TextType%>" id="hReasonSelected" name="hReasonSelected" value="" <%=InputSize %> />
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hUserId"%>            
            <input type="<%=TextType%>" id="hUserId" name="hUserId" value="<%=Session("UserId")%>" <%=InputSize %> />   
            &nbsp;&nbsp;<%If TextType = "text" Then Response.Write "hJobType"%> 
            <input type="<%=TextType%>" id="hJobType" name="hJobType" value="<%=Session("JobTypeNo")%>" <%=InputSize %> />
            <br /><br />
            <input type="<%=TextType%>" id="hJobRef" name="hJobRef" value="<%=hJobRef%>"  />
            &nbsp;&nbsp;<input type="<%=TextType%>" id="frmName" name="frmName" value="frmNcAddNewCL" />
           
        </td>
    </tr>
    <tr>
        <td id="PrivateNotes" style="visibility: visible"><%=Trim(PLCWPrivateNotes)%></td>
    </tr>    
</table>

</form>

</body>

<!-- PriceList CalcWizard Notes are not displayed. -->
<!-- This script adds it to any production notes and overwrites the Production Notes at the Top -->
<script type="text/javascript" >
    var PdNotes
    var PlCwNotes
    var CombinedNotes
    PdNotes = document.getElementById('ProdNotes').innerHTML
    PlCwNotes = document.getElementById('PrivateNotes').innerHTML
    if (PlCwNotes == "")
    { CombinedNotes = PdNotes }
    else
    { CombinedNotes = PdNotes + '<br />' + PlCwNotes }
    //CombinedNotes = PdNotes + '<br />' + PlCwNotes 
    document.getElementById('ProdNotes').innerHTML = CombinedNotes;   
</script>
</html>

<%

Sub GetItems (QuotationId) '## Passed Quote.ID to get QuotationItems.ID for each quote item

    Dim GetItemsSql
    Dim GetItemsRs
    
    DetailItems = ""

    GetItemsSql = "Select QuotationItems.QuotationId, QuotationItems.ItemId, QuotationItems.Id, QuotationItems.RecState"
    GetItemsSql = GetItemsSql & " From QuotationItems Where(QuotationItems.QuotationId = '" & Cstr(QuotationId) & "')"
    GetItemsSql = GetItemsSql & " AND (QuotationItems.RecState = 0)"
    GetItemsSql = GetItemsSql & " ORDER BY QuotationItems.ItemId"

    Set GetItemsRs = Server.CreateObject("ADODB.Recordset")
    GetItemsRs.ActiveConnection = strConnClarity
	GetItemsRs.Source = GetItemsSql 	
	GetItemsRs.CursorType = 0
	GetItemsRs.CursorLocation = 3
	GetItemsRs.LockType = 1
	GetItemsRs.Open
	Set GetItemsRs.ActiveConnection = Nothing

    While not GetItemsRs.EOF
        If GetItemsRs("Id") > 0 Then
            If DetailItems = "" Then
                DetailItems = GetItemsRs("Id") &  ";"
                AlphaItems = GetItemsRs("ItemId") &  ";"
            Else
                DetailItems = DetailItems & GetItemsRs("Id") & ";"
                AlphaItems = AlphaItems & GetItemsRs("ItemId") &  ";"
            End If
        End If
        GetItemsRs.MoveNext
    Wend
    
    GetItemsRs.Close
    Set GetItemsRs = Nothing

End Sub

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

    AddressRs.ActiveConnection = strConnClarity
	AddressRs.Source = strAddressSql	
	AddressRs.CursorType = 0
	AddressRs.CursorLocation = 3
	AddressRs.LockType = 1
	AddressRs.Open
	Set AddressRs.ActiveConnection = Nothing
	
	If AddressRs.EOF = True Or AddressRs.BOF = True Then
	    AddressRs.Close
	    Set AddressRs = Nothing
	    Exit Sub
	End if
	
	Client = AddressRs("Name")
	
	If AddressRs("Name") <> "" Then LocalCompanyAddress = AddressRs("Name")
	If AddressRs("Address") <> "" Then LocalCompanyAddress = LocalCompanyAddress & "<br />" & AddressRs("Address")
	If AddressRs("Address2") <> "" Then LocalCompanyAddress = LocalCompanyAddress & "<br />" & AddressRs("Address2")
	If AddressRs("Address3") <> "" Then LocalCompanyAddress = LocalCompanyAddress & "<br />" & AddressRs("Address3")
	If AddressRs("City") <> "" Then LocalCompanyAddress = LocalCompanyAddress & "<br />" & AddressRs("City")
	If AddressRs("County") <> "" Then LocalCompanyAddress = LocalCompanyAddress & "<br />" & AddressRs("County")
	If AddressRs("Postcode") <> "" Then LocalCompanyAddress = LocalCompanyAddress & "<br />" & AddressRs("Postcode")
	
	If JobIsOutdoor = False Then
	   If AddressRs("DeliveryContact") <> "" Then LocalDeliveryAddress = AddressRs("DeliveryContact")
	   If AddressRs("DeliveryCompany") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryCompany")
	Else
	    If AddressRs("DeliveryCompany") <> "" Then LocalDeliveryAddress = AddressRs("DeliveryCompany")
	End If
	
	If AddressRs("DeliveryAddress") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryAddress")
	If AddressRs("DeliveryAddress2") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryAddress2")
	If AddressRs("DeliveryAddress3") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryAddress3")
	If AddressRs("DeliveryCity") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryCity")
	If AddressRs("DeliveryCounty") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryCounty")
	If AddressRs("DeliveryPostcode") <> "" Then LocalDeliveryAddress = LocalDeliveryAddress & "<br />" & AddressRs("DeliveryPostcode")
	
    If LocalDeliveryAddress = "" Then
        If AddressRs("DeliveryContact") <> "" Then
            DeliveryAddress = AddressRs("DeliveryContact")
        Else
            DeliveryAddress = LocalCompanyAddress
        End If
    Else
        DeliveryAddress = LocalDeliveryAddress
    End If  
    
    AddressRs.Close
	Set AddressRs = Nothing 
	
	If Left(DeliveryAddress,6) = "<br />" Then DeliveryAddress = Mid(DeliveryAddress,7)
 
 End Sub

'#############################################################################################
 
 Private Sub Material (MatItemID)    '## Passed QuotationItems.Id 

'## On Error Resume Next
 
    Dim MaterialSql
    Dim MaterialRS
     
    MaterialSql = "SELECT QuotationItemDetail.QuotationItemId, QuotationItemDetail.PartCode, GlobalPartList.Description,"
    MaterialSql = MaterialSql & " QuotationItemDetail.PartType, QuotationItemDetail.SizeY,GlobalPartList.Level3"
    MaterialSql = MaterialSql & " FROM QuotationItemDetail INNER JOIN"
    MaterialSql = MaterialSql & " GlobalPartList ON QuotationItemDetail.PartCode = GlobalPartList.Partcode"
    MaterialSql = MaterialSql & " WHERE (QuotationItemDetail.QuotationItemId = " & MatItemID & ")"
    MaterialSql = MaterialSql & " AND (GlobalPartList.Level3 NOT IN ('Consumables'))"
                                    '## was PartType = 1) material, now material or non stock = 5
    MaterialSql = MaterialSql & " AND (QuotationItemDetail.PartType IN (1,5))"
    MaterialSql = MaterialSql & " AND (QuotationItemDetail.SizeY > 150) AND (QuotationItemDetail.PartCode <> '<Unknown>')"
    
    Set MaterialRS = Server.CreateObject("ADODB.Recordset")

    MaterialRS.ActiveConnection = strConnClarity
	MaterialRS.Source = MaterialSql	
	MaterialRS.CursorType = adOpenForwardOnly
	MaterialRS.CursorLocation = adUseClient
	MaterialRS.LockType = adLockReadOnly
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
                MaterialList = MaterialRS("Description") & ";"
            Else
                MaterialList = MaterialList &  MaterialRS("Description") & ";"
            End If        
            MaterialRS.MoveNext        
        Wend
    Else 
        If MaterialList = "" Then
            MaterialList = MaterialRS("Description") & ";"
        Else
            MaterialList = MaterialList &  MaterialRS("Description") & ";"
        End If 
    End If     
    
    MaterialRS.Close
    Set MaterialRS = Nothing
    
    Err.Clear      
 
End Sub


%>