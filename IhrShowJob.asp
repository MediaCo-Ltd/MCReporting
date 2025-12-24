<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\Declarations.asp" -->
<!--#include file="..\##GlobalFiles\PkId.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="IhrGetQcData.asp" -->

<% 
'## On Error Resume Next

Response.Expires = -1
Response.AddHeader "Pragma", "no-cache"
Response.AddHeader "cache-control", "no-store"
Response.AddHeader "Refresh",CStr(CInt(Session.Timeout + 1) * 60)& "; URL=SessionExpired.html"
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim DetailRS
Dim strDetailSql
Dim RedoItemDataRs
Dim DetailJob
Dim DisplayStatus
Dim Client
Dim DeliveryAddress
Dim DisplaySub
DisplaySub = ""

Dim JobIsOutdoor

Dim ReqDelHeader
Dim ItemTitle
Dim DetailDescription

Dim MgsColSpan
Dim ReasonSize
Dim ReasonTxtColour

If Session("ShowHiddenSize") = "" Then
    ReasonSize = " size='40' "
Else
    ReasonSize = ""
End If  

Dim RedoId
RedoId = Request.QueryString("Rid")
DetailJob = GetQuoteIdRs (RedoId)

Dim ReturnPage
'ReturnPage = Request.QueryString("rp")

'If ReturnPage = "0" Then
    ReturnPage = "IhrSelectdate.asp"
'Else
'    ReturnPage = "IhrDisplayByDate.asp?sd=" & Request.QueryString("rp")
'End If


'Not Distinct  Can't use if showing notes 
strDetailSql = "SELECT Job.Reference AS JobRef, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId, Quotation.user11 As ArtworkStatus,"
strDetailSql = strDetailSql & " Quotation.Description AS JobDesc, Quotation.Notes, Job.Id as JobId, Quotation.User6,"
strDetailSql = strDetailSql & " Case When Quotation.User15 = NULL Then '' Else Quotation.User15 END AS ArtworkMatch,"
strDetailSql = strDetailSql & " Quotation.status AS QuoteStatus, Quotation.JobTypeId, Quotation.completedDate AS DelDate, Contacts.Forename + ' ' + Contacts.Surname AS BookedBy"
strDetailSql = strDetailSql & " FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId INNER JOIN Contacts ON Job.CreatedBy = Contacts.Id"
strDetailSql = strDetailSql & " Where(Quotation.id = '" & DetailJob & "')"
 
Set DetailRS = Server.CreateObject("ADODB.Recordset")

DetailRS.ActiveConnection = Session("ClarityConn") 'strConnClarity
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
    Response.Redirect "IhrError.asp"
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
VisibilityTxt = ""

'## Get Address details and Company Name
DelAddress DetailRS("QuoteId")

Dim hQuoteRef
Dim hDesc
Dim DispalyRedoDate

Dim NonActiveTxt
NonActiveTxt = ""

If GetRedoStatus = Cbool(False) Then
    NonActiveTxt = "&nbsp;&nbsp;&nbsp;&nbsp;Record is dormant."
Else
    NonActiveTxt = ""
End If

DispalyRedoDate = "Redo&nbsp;Created&nbsp;" & GetRedoDate 

hQuoteRef = Mid(DetailRS("JobRef"),4)
hDesc = DetailRS("JobDesc")

Dim RevisionTxt
RevisionTxt = ""
Dim RevisionTitle
RevisionTitle = ""
Dim RevisionNo
RevisionNo = GetRevision

If Cint(RevisionNo) = 1 Then
    RevisionTxt = ""
    RevisionTitle = ""
Else
    RevisionTxt = " (" & RevisionNo & ")"
    RevisionTitle = " title ='Redo No "  & RevisionNo & " for this job' "
End If

Dim DetailColSpan
If Session("ShowRedoCost") = CBool(True) Then
    DetailColSpan = 10
Else
    DetailColSpan = 9
End If

Dim SpTitle 
SpTitle = ""

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


<body style="padding: 0px; margin: 0px" >

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
        <td align="left" width="20%" style="height: 20px">
            &nbsp;&nbsp;<a href ="javascript:window.location.replace('<%=ReturnPage%>');" style="font-size:12px; color: <%=NewBlue%>;">Return to select date Page</a><br /><br />
        </td>
        <td align="center">
            <img align="top" alt="mediaco logo" src="Images/X1.ico" style="width: 20px; height: 20px;" />&nbsp;
            <font style="font-weight: bold; color:<%=NewBlue%>;" size="3">Redo record (<%=DetailRS("JobRef") %>)&nbsp;Created&nbsp;By&nbsp;<%=GetRedoCreator%></font>
        </td>
        <td align="right" width="20%">
            <a id="logoff" href ="javascript:LogOff();" style="font-size:12px; color: <%=NewBlue%>;">Log&nbsp;Out</a>&nbsp;&nbsp;
        </td> 
    </tr>
</table>
<br />

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
    
    <td width="50" class="styleTDleft" style="color: #000000" <%=RevisionTitle%> ><%=DetailRS("JobRef") & RevisionTxt %></td>
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
<br />

<table  style="padding-right: 10px; padding-left: 10px;" align="center" cellpadding="0" cellspacing="0"  width="95%">	
    <tr>
        <th align="left" colspan="<%=DetailColSpan%>" style="font-size: 14px; padding-left: 5px; padding-bottom: 5px"><%=DispalyRedoDate%>
        <%
        If NonActiveTxt <> "" Then Response.Write "<label style='color: #FF0000;'>" & NonActiveTxt & "</label>"
        %>
        </th>    
    </tr>
    <tr >		 
	    <th width="5%" class="styleTHleft" >Item</th> 
	    <th width="7%" class="styleTHstd" >SP&nbsp;Items</th>
	    <th width="25%" class="styleTHstd" >Description</th>
	    <th width="20%" class="styleTHstd" >Substrate</th>	    
	    <th width="6%" class="styleTHstd" >Redo&nbsp;Qty</th>
	    
	    <th width="6%" class="styleTHstd" >Redo&nbsp;Sqm</th>
	    <th width="10%" class="styleTHstd" >Department</th>
	    <th width="10%" class="styleTHstd" >Printer</th>
        <th width="10%" class="styleTHstd" >Reason</th>
        <%
            If Session("ShowRedoCost") = CBool(True) Then Response.Write "<th class='styleTHstd' width='10%'>Unit Cost</th>"        
        %>
           		    
    </tr>
         
    <%
    '## Data Loop
    DetailRS.Close
    Set DetailRS = Nothing 
    
    
    Dim ReasonDescription
    Dim GroupDescription
    Dim ReasonDescriptionTitle
    ReasonDescriptionTitle = ""
    
    '## Get Item Qc Details
    RedoItemData
    
    If RedoItemDataRs.BOF = True Or RedoItemDataRs.EOF = True Then
        RedoItemDataRs.Close
        Set RedoItemDataRs = Nothing        
        Err.Clear
        Response.Redirect "Error.asp"
    End If
    
    While Not RedoItemDataRs.EOF   
    
        GroupDescription = GetGroupDesc (RedoItemDataRs("GroupId"))
        ReasonDescription = GetReasonDesc (RedoItemDataRs("ReasonCode"))
        
        If GroupDescription = "Printing" And RedoItemDataRs("WorkCentreName") = "Heat Press" Then
            GroupDescription = "Heat Press"
        End If
        
        ReasonTxtColour = "#000000"
        
        If Instr(1,ReasonDescription,"Other",1) > 0 Then 
            If RedoItemDataRs("Reason") <> "" Then
                ReasonDescription = GroupDescription & "&nbsp;(Other)<br />" & RedoItemDataRs("Reason")        
            Else
                ReasonTxtColour = "#FF0000"
                ReasonDescription = GroupDescription & "&nbsp;(Other)&nbsp;Selected<br />But&nbsp;no&nbsp;reason&nbsp;was&nbsp;entered."
            End If
        Else
            If RedoItemDataRs("Reason") <> "" Then
                'ReasonDescription = RedoItemDataRs("Reason")
                ReasonDescription = ReasonDescription & ".<br />" & RedoItemDataRs("Reason")        
            End If
        
        End If        
        
        ItemTitle = RedoItemDataRs("ItemDescription")
        ItemTitle = Replace(ItemTitle,"<br />",Chr(13),1,-1,1)
        DetailDescription = RedoItemDataRs("ItemDescription")
        DetailDescription = Replace(DetailDescription,"<br />",Chr(13),1,-1,1)
        If Len(DetailDescription) > 70 Then DetailDescription = Left(DetailDescription,60) & " ..."
        
        DisplaySub = Replace(RedoItemDataRs("Substrates"),"#","<br />",1,-1,1)
        
        If Len(ReasonDescription) > 70 Then
            ReasonDescriptionTitle = Replace(ReasonDescription,"<br />",Chr(13),1,-1,1)
            ReasonDescription = Left(ReasonDescription,60) & " ..."
        End If
       
        
    %>   
     
    <tr   >                     
        <td width="25" class="styleTDleft" ><%=RedoItemDataRs("ItemAlpha")%></td>
        <td class="styleTDstd"><%=RedoItemDataRs("SpId")%></td>              
        <td class="styleTDstd" title="<%=ItemTitle%>" ><%=DetailDescription%></td>
        <td class="styleTDstd"><%=DisplaySub%></td>                       
        <td class="styleTDstd"><%=RedoItemDataRs("NewQty")%></td>
        
        <td class="styleTDstd" align="left" ><%=RedoItemDataRs("ItemNewSqm")%></td>        
        <td class="styleTDstd"><%=GroupDescription%></td>
        <td class="styleTDstd"><%=RedoItemDataRs("Printer")%></td>
        <td class="styleTDstd" title="<%=ReasonDescriptionTitle%>" align="left" style="color: <%=ReasonTxtColour%>"><%=ReasonDescription%></td>
        
        <%
            If Session("ShowRedoCost") = CBool(True) Then Response.Write "<td class='styleTDstd' >" & RedoItemDataRs("UnitCost") & "</td>"                    
        %>
        
        
        
    </tr>
        <%

        ItemTitle = ""
        DetailDescription = ""
        ReasonDescription = ""
        GroupDescription = ""
        DisplaySub = ""
        ReasonDescriptionTitle = ""
        
        RedoItemDataRs.MoveNext          
          
    Wend 
    
RedoItemDataRs.Close
Set RedoItemDataRs = Nothing 

%>
</table>
<br />

<table align="center" style="padding-right: 10px; padding-left: 10px;" cellpadding="0" cellspacing="0"  width="95%">
    <tr>
        <td>&nbsp;&nbsp;<br />
            <%If Session("ShowHidden") = "text" Then Response.Write "Quote Ref" %>&nbsp;<input id="QuoteRef" name="QuoteRef" value="<%=hQuoteRef%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "QuoteID" %>&nbsp;<input id="QuoteID" name="QuoteID" value="<%=DetailJob%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "Client" %>&nbsp;<input id="Client" name="Client" value="<%=Client%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "Desc" %>&nbsp;<input id="Desc" name="Desc" value="<%=hDesc%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
            <%If Session("ShowHidden") = "text" Then Response.Write "RedoId" %>&nbsp;<input id="RedoId" name="RedoId" value="<%=RedoId%>"  type="<%=Session("ShowHidden")%>" />&nbsp;
        </td>
    </tr>      
</table>


</body>

</html>

<%

Sub RedoItemData()

    Dim RedoItemDataSql
    RedoItemDataSql = "Select * From OHQuoteItemData Where(OHQuoteDataId = " & RedoId & ") Order By ItemAlpha"
    
	Set RedoItemDataRs = Server.CreateObject("ADODB.Recordset")
	RedoItemDataRs.ActiveConnection = Session("ConnQC") 
	RedoItemDataRs.Source = RedoItemDataSql
	RedoItemDataRs.CursorType = Application("adOpenForwardOnly")
	RedoItemDataRs.CursorLocation = Application("adUseClient")
	RedoItemDataRs.LockType = Application("adLockReadOnly")
	RedoItemDataRs.Open

End Sub

'################################################################

Function GetReasonDesc(RCode)

    Dim GetCodesRs
    Dim GetCodesSql
    Dim LocalDesc

    GetCodesSql = "Select ReasonDescription From ReasonCodes Where (ReasonCode = " & RCode & ")"
    Set GetCodesRs = Server.CreateObject("ADODB.Recordset")
    GetCodesRs.ActiveConnection = Session("ConnQC")
    GetCodesRs.Source = GetCodesSql

    GetCodesRs.CursorType = Application("adOpenForwardOnly")
    GetCodesRs.CursorLocation = Application("adUseClient")
    GetCodesRs.LockType = Application("adLockReadOnly")
    GetCodesRs.Open
    
    LocalDesc = GetCodesRs("ReasonDescription")
    GetCodesRs.Close
    Set GetCodesRs = Nothing
    
    GetReasonDesc = LocalDesc

End Function

'################################################################

Function GetGroupDesc(GrpId)

    Dim GetGroupRs
    Dim GetGroupSql
    Dim LocalDesc

    GetGroupSql = "Select GroupName From Department Where (GroupId = " & GrpId & ")"
    Set GetGroupRs = Server.CreateObject("ADODB.Recordset")
    GetGroupRs.ActiveConnection = Session("ConnQC")
    GetGroupRS.Source = GetGroupSql

    GetGroupRs.CursorType = Application("adOpenForwardOnly")
    GetGroupRs.CursorLocation = Application("adUseClient")
    GetGroupRs.LockType = Application("adLockReadOnly")
    GetGroupRs.Open
    
    LocalDesc = GetGroupRs("GroupName")
    GetGroupRs.Close
    Set GetGroupRs = Nothing
    
    GetGroupDesc = LocalDesc

End Function

'################################################################

Function GetRedoStatus()

    Dim GetRedoStatusSql
    Dim GetRedoStatusRs
    Dim LocalStatus

    GetRedoStatusSql = "Select ID, Active From OhQuoteData Where (ID = " & RedoId & ")"
    Set GetRedoStatusRs = Server.CreateObject("ADODB.Recordset")
    GetRedoStatusRs.ActiveConnection = Session("ConnQC")
    GetRedoStatusRS.Source = GetRedoStatusSql

    GetRedoStatusRs.CursorType = Application("adOpenForwardOnly")
    GetRedoStatusRs.CursorLocation = Application("adUseClient")
    GetRedoStatusRs.LockType = Application("adLockReadOnly")
    GetRedoStatusRs.Open
    
    LocalStatus = GetRedoStatusRs("Active")
        
    GetRedoStatusRs.Close
    Set GetRedoStatusRs = Nothing
    
    GetRedoStatus = LocalStatus

End Function

'################################################################

Function GetRedoDate()

    Dim GetRedoDateSql
    Dim GetRedoDateRs
    Dim LocalDate

    GetRedoDateSql = "Select ID, Created From OhQuoteData Where (ID = " & RedoId & ")"
    Set GetRedoDateRs = Server.CreateObject("ADODB.Recordset")
    GetRedoDateRs.ActiveConnection = Session("ConnQC")
    GetRedoDateRS.Source = GetRedoDateSql

    GetRedoDateRs.CursorType = Application("adOpenForwardOnly")
    GetRedoDateRs.CursorLocation = Application("adUseClient")
    GetRedoDateRs.LockType = Application("adLockReadOnly")
    GetRedoDateRs.Open
    
    LocalDate = GetRedoDateRs("Created")
    
    GetRedoDateRs.Close
    Set GetRedoDateRs = Nothing
    
    GetRedoDate = LocalDate

End Function

'################################################################

Function GetRedoCreator()

    Dim GetRedoCreatorSql
    Dim GetRedoCreatorRs
    Dim LocalName

    GetRedoCreatorSql = "Select ID, CreatedBy From OhQuoteData Where (ID = " & RedoId & ")"
    Set GetRedoCreatorRs = Server.CreateObject("ADODB.Recordset")
    GetRedoCreatorRs.ActiveConnection = Session("ConnQC")
    GetRedoCreatorRS.Source = GetRedoCreatorSql

    GetRedoCreatorRs.CursorType = Application("adOpenForwardOnly")
    GetRedoCreatorRs.CursorLocation = Application("adUseClient")
    GetRedoCreatorRs.LockType = Application("adLockReadOnly")
    GetRedoCreatorRs.Open
    
    LocalName = GetRedoCreatorRs("CreatedBy")
    GetRedoCreatorRs.Close
    Set GetRedoCreatorRs = Nothing
    
    GetRedoCreator = LocalName

End Function

'################################################################

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

    AddressRs.ActiveConnection = Session("ClarityConn") 'strConnClarity
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
 
 '################################################################

Function GetRevision()

    Dim RevisionSql
    Dim RevisionRs
    Dim LocalRevision
    LocalRevision = 1

    RevisionSql = "Select ID, Revision From OhQuoteData Where (ID = " & RedoId & ")"
    Set RevisionRs = Server.CreateObject("ADODB.Recordset")
    RevisionRs.ActiveConnection = Session("ConnQC")
    RevisionRS.Source = RevisionSql

    RevisionRs.CursorType = Application("adOpenForwardOnly")
    RevisionRs.CursorLocation = Application("adUseClient")
    RevisionRs.LockType = Application("adLockReadOnly")
    RevisionRs.Open
    
    If RevisionRs.BOF = True Or RevisionRs.EOF = True Then
        LocalRevision = 1    
    Else
        LocalRevision = RevisionRs("Revision")
    End If
    
    
    RevisionRs.Close
    Set RevisionRs = Nothing
    
    GetRevision = LocalRevision

End Function


%>