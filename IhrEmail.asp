<%@Language="VBScript" Codepage="1252" %>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\Declarations.asp" -->
<!--#include file="..\##GlobalFiles\PkId.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="IhrGetQcData.asp" -->

<% 
'## On Error Resume Next

Dim DetailRS
Dim strDetailSql
Dim RedoItemDataRs
Dim DetailJob
Dim DisplayStatus
Dim Client
Dim DeliveryAddress
Dim DisplaySub
DisplaySub = ""
Dim MsgHtml

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

Dim SendToAw 
SendToAw  = Request.QueryString("AW")

Dim SendToBookingIn
SendToBookingIn = Request.QueryString("AM")

'Not Distinct  Can't use if showing notes 
strDetailSql = "SELECT Job.Reference AS JobRef, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId, Quotation.user11 As ArtworkStatus,"
strDetailSql = strDetailSql & " Quotation.Description AS JobDesc, Quotation.Notes, Job.Id as JobId, Quotation.User6,"
strDetailSql = strDetailSql & " Case When Quotation.User15 = NULL Then '' Else Quotation.User15 END AS ArtworkMatch,"
strDetailSql = strDetailSql & " Quotation.status AS QuoteStatus, Quotation.JobTypeId, Quotation.completedDate AS DelDate, Contacts.Forename + ' ' + Contacts.Surname AS BookedBy"
strDetailSql = strDetailSql & " FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId INNER JOIN Contacts ON Job.CreatedBy = Contacts.Id"
strDetailSql = strDetailSql & " Where(Quotation.id = '" & DetailJob & "')"
 
Set DetailRS = Server.CreateObject("ADODB.Recordset")

DetailRS.ActiveConnection = Session("ClarityConn")   
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

DispalyRedoDate = "Redo&nbsp;Created&nbsp;" & GetRedoDate 

hQuoteRef = Mid(DetailRS("JobRef"),4)
hDesc = DetailRS("JobDesc")

Dim RevisionTxt
RevisionTxt = ""
Dim RevisionNo
RevisionNo = GetRevision

If Cint(RevisionNo) = 1 Then
    RevisionTxt = ""
Else
    RevisionTxt = " (" & RevisionNo & ")"
End If

'############################################ Set up email

Dim WhoTo
Dim WhoToName
Dim WhoToCC
Dim WhoToNameCC
Dim Emsg
Dim ContentId


Set Emsg = Server.CreateOBject( "JMail.Message" )

Emsg.Logging = true
Emsg.Silent = true
Emsg.From = "no-reply@mediaco.co.uk"
Emsg.FromName = "Redo Reporting"
Emsg.Subject = "New Redo Record - " & hQuoteRef & " " & Client '## & " SendToAw = " & SendToAw

'ContentId = Emsg.AddAttachment(Session("RootPath") & "\Images\X1.png",true,"image/png")

Emsg.HTMLBody = CreateMsg

Session("Smtp") = "mx496502.smtp-engine.com" 
Emsg.MailServerUserName = "warren.morris@mediaco.co.uk"
Emsg.MailServerPassWord = "Q6)9l7Sit6dAV*" 

        
If Session("PC-Name") = "Home" Then
    Emsg.AddRecipient "alan.holgate@yahoo.co.uk", "Alan Holgate"
Else        
    Emsg.AddRecipient "RedoProduction@mediaco.co.uk", "Redo Production"    
    If SendToAw = "True" Then Emsg.AddRecipient "artwork@mediaco.co.uk", "artwork"
    If SendToBookingIn = "True" Then Emsg.AddRecipient "vicky.joines@mediaco.co.uk", "Vicky Joines"
End If

'## For Testing
'## Emsg.ClearRecipients
'## Emsg.AddRecipient "alan.holgate@mediaco.co.uk", "Alan Holgate"    

If Not Emsg.Send(Session("Smtp")) Then
    Session("JobNoError") = "Failed to send Email"
Else
    Session("JobNoError") = "Update ok"
End If

Emsg.Close
Set Emsg = nothing

Response.Redirect "IhrJobNo.asp"

'################################################################

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

    AddressRs.ActiveConnection = Session("ClarityConn")   'strConnClarity
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

'#####################################

Function CreateMsg()

'**************************
'   Create Html Document 
'**************************

MsgHtml = ""

MsgHtml = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"  & VbCrLf 

MsgHtml = MsgHtml & "<html xmlns='http://www.w3.org/1999/xhtml' xml:lang='en' lang='en'>" & VbCrLf 
MsgHtml = MsgHtml & "<head>" & VbCrLf 

MsgHtml = MsgHtml & "<title>In House Redo</title>" & VbCrLf 
MsgHtml = MsgHtml & "<meta http-equiv='Content-Type' content='text/html; charset=iso-8859-1' />" & VbCrLf 

MsgHtml = MsgHtml & "<style type='text/css'>" & VbCrLf
MsgHtml = MsgHtml & "body" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "font-family:  'Verdana', 'Arial', 'Geneva',  'Lucida', 'San-Serif', 'Lucida Sans', 'Lucida Console', 'MS Sans Serif';" & VbCrLf
MsgHtml = MsgHtml & "font-size: 69%;" & VbCrLf
MsgHtml = MsgHtml & "font-style: normal;" & VbCrLf
MsgHtml = MsgHtml & "font-weight: normal;" & VbCrLf
MsgHtml = MsgHtml & "font-variant:  normal;" & VbCrLf
MsgHtml = MsgHtml & "text-transform:  none;" & VbCrLf
MsgHtml = MsgHtml & "text-decoration:  none;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".styleTHstd" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "height: 30px;" & VbCrLf
MsgHtml = MsgHtml & "color: #0069AA;" & VbCrLf
MsgHtml = MsgHtml & "text-align: center;" & VbCrLf
MsgHtml = MsgHtml & "border-style: solid solid solid none;" & VbCrLf 
MsgHtml = MsgHtml & "border-width: 1px;" & VbCrLf
MsgHtml = MsgHtml & "border-color: Black;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".styleTHleft" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "height: 30px;" & VbCrLf
MsgHtml = MsgHtml & "color: #0069AA;" & VbCrLf
MsgHtml = MsgHtml & "text-align: center;" & VbCrLf
MsgHtml = MsgHtml & "border-style: solid solid solid solid;" & VbCrLf 
MsgHtml = MsgHtml & "border-width: 1px;" & VbCrLf
MsgHtml = MsgHtml & "border-color: Black;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".styleTDstd" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "color: Black;" & VbCrLf
MsgHtml = MsgHtml & "height: 20px;" & VbCrLf 
MsgHtml = MsgHtml & "text-align: center;" & VbCrLf
MsgHtml = MsgHtml & "border-style: none solid solid none;" & VbCrLf 
MsgHtml = MsgHtml & "border-width: 1px;" & VbCrLf
MsgHtml = MsgHtml & "border-color: Black;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & ".styleTDleft" & VbCrLf
MsgHtml = MsgHtml & "{" & VbCrLf
MsgHtml = MsgHtml & "height: 20px;" & VbCrLf
MsgHtml = MsgHtml & "color: #0069AA;" & VbCrLf
MsgHtml = MsgHtml & "text-align: center;" & VbCrLf
MsgHtml = MsgHtml & "border-style: none solid solid solid;" & VbCrLf
MsgHtml = MsgHtml & "border-width: 1px;" & VbCrLf
MsgHtml = MsgHtml & "border-color: Black;" & VbCrLf
MsgHtml = MsgHtml & "}" & VbCrLf

MsgHtml = MsgHtml & "</style>" & VbCrLf

MsgHtml = MsgHtml & "</head>" & VbCrLf 
MsgHtml = MsgHtml & "<body style='padding: 0px; margin: 0px' >" & VbCrLf 

'## With logo can't get it correct
'MsgHtml = MsgHtml & "<table width='100%' >" & VbCrLf
'MsgHtml = MsgHtml & "<tr>" & VbCrLf
'MsgHtml = MsgHtml & "<td align='left' valign='bottom' width='25%' >" & VbCrLf
'MsgHtml = MsgHtml & "<img align='left' alt='Redo logo' width='30px' src=""cid:" & ContentId & """ /> " & VbCrLf  ' width ='90px'
'MsgHtml = MsgHtml & "</td>" & VbCrLf
'MsgHtml = MsgHtml & "<font style='font-weight: bold; color:" & NewBlue & ";' size='3'>Redo record (" & DetailRS("JobRef") & ")&nbsp;Created&nbsp;By&nbsp;" & GetRedoCreator & "</font>" & VbCrLf 
'MsgHtml = MsgHtml & "<td align='right' valign='bottom' width='25%' >&nbsp;</td>" & VbCrLf
'MsgHtml = MsgHtml & "</tr>" & VbCrLf
'MsgHtml = MsgHtml & "</table>" & VbCrLf

'## Without Logo
MsgHtml = MsgHtml & "<table style='padding-right: 10px; padding-left: 10px; width: 100%;'>" & VbCrLf 
MsgHtml = MsgHtml & "<tr>" & VbCrLf 
MsgHtml = MsgHtml & "<td align='left' width='20%' style='height: 20px'>" & VbCrLf 
MsgHtml = MsgHtml & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "<td align='center'>" & VbCrLf 
MsgHtml = MsgHtml & "<font style='font-weight: bold; color:" & NewBlue & ";' size='3'>Redo record (" & DetailRS("JobRef") & ")&nbsp;Created&nbsp;By&nbsp;" & GetRedoCreator & "</font>" & VbCrLf 
MsgHtml = MsgHtml & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "<td align='right' width='20%'>" & VbCrLf 
MsgHtml = MsgHtml & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "</tr>" & VbCrLf

'## Spacing between tables
MsgHtml = MsgHtml & "<tr>"
MsgHtml = MsgHtml & "<td style='height: 10px' colspan='3'>&nbsp;</td>"
MsgHtml = MsgHtml & "</tr>"
MsgHtml = MsgHtml & "</table>" & VbCrLf

MsgHtml = MsgHtml & "<table style='padding-right: 10px; padding-left: 10px;' align='center' cellpadding='0' cellspacing='0'  width='95%'>" & VbCrLf 
MsgHtml = MsgHtml & "<tr >" & VbCrLf 
    
MsgHtml = MsgHtml & "<th width='50' class='styleTHleft'>Job</th>" & VbCrLf 
MsgHtml = MsgHtml & "<th width='50' class='styleTHstd'>" & ReqDelHeader & "</th>" & VbCrLf 
MsgHtml = MsgHtml & "<th width='50' class='styleTHstd'>Created By</th>" & VbCrLf 

    
MsgHtml = MsgHtml & "<th width='130' class='styleTHstd'>Company</th>" & VbCrLf         
MsgHtml = MsgHtml & "<th width='130' class='styleTHstd'>Description</th>" & VbCrLf 
MsgHtml = MsgHtml & "<th width='50' class='styleTHstd' >Status</th>" & VbCrLf 
MsgHtml = MsgHtml & "<th width='50' class='styleTHstd'>Delivery Method</th>" & VbCrLf 
MsgHtml = MsgHtml & "<th width='50' class='styleTHstd'>Colour Match</th>" & VbCrLf 
MsgHtml = MsgHtml & "</tr>" & VbCrLf 

MsgHtml = MsgHtml & "<tr >" & VbCrLf     
MsgHtml = MsgHtml & "<td width='50' class='styleTDleft' style='color: #000000'>" & DetailRS("JobRef") & RevisionTxt & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "<td width='50' class='styleTDstd'>" 
 
If ReqDelHeader = "Delivered On" Then
    MsgHtml = MsgHtml &  Trim(Left(DetailRS("DelDate"),10))    
Else
    MsgHtml = MsgHtml &  Trim(Left(DetailRS("ReqDate"),10))
End If    

MsgHtml = MsgHtml & "</td>"  & VbCrLf  

MsgHtml = MsgHtml & "<td width='50' class='styleTDstd'>" & DetailRS("BookedBy") & "</td>" & VbCrLf
MsgHtml = MsgHtml & "<td width='130' class='styleTDstd'>" & Client & "</td>" & VbCrLf      
MsgHtml = MsgHtml & "<td width='130' class='styleTDstd'>" & DetailRS("JobDesc") & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "<td width='50' class='styleTDstd' >" & DisplayStatus & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "<td width='50' class='styleTDstd'>" & DetailRS("User6") & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "<td width='50' class='styleTDstd'>" & DetailRS("ArtworkMatch")  & "</td>" & VbCrLf 
MsgHtml = MsgHtml & "</tr>" & VbCrLf

'## Spacing between tables
MsgHtml = MsgHtml & "<tr>"
MsgHtml = MsgHtml & "<td style='height: 15px' colspan='8'>&nbsp;</td>"
MsgHtml = MsgHtml & "</tr>"
MsgHtml = MsgHtml & "</table>" & VbCrLf 

MsgHtml = MsgHtml & "<table  style='padding-right: 10px; padding-left: 10px;' align='center' cellpadding='0' cellspacing='0'  width='95%'>" & VbCrLf 	
MsgHtml = MsgHtml & "<tr>" & VbCrLf
MsgHtml = MsgHtml & "<th align='left' colspan='9' style='font-size: 14px; padding-left: 5px; padding-bottom: 5px'>" & DispalyRedoDate & "</th>" & VbCrLf    
MsgHtml = MsgHtml & "</tr>" & VbCrLf
MsgHtml = MsgHtml & "<tr >" & VbCrLf		 
MsgHtml = MsgHtml & "<th width='5%' class='styleTHleft' >Item</th>" & VbCrLf 
MsgHtml = MsgHtml & "<th width='7%' class='styleTHstd' >SP&nbsp;Items</th>" & VbCrLf
MsgHtml = MsgHtml & "<th width='25%' class='styleTHstd' >Description</th>" & VbCrLf
MsgHtml = MsgHtml & "<th width='20%' class='styleTHstd' >Substrate</th>" & VbCrLf	    
MsgHtml = MsgHtml & "<th width='6%' class='styleTHstd' >Redo&nbsp;Qty</th>" & VbCrLf

MsgHtml = MsgHtml & "<th width='6%' class='styleTHstd' >Redo&nbsp;Sqm</th>" & VbCrLf
MsgHtml = MsgHtml & "<th width='10%' class='styleTHstd' >Department</th>" & VbCrLf
MsgHtml = MsgHtml & "<th width='10%' class='styleTHstd' >Printer</th>" & VbCrLf
MsgHtml = MsgHtml & "<th width='10%' class='styleTHstd' >Reason</th>" & VbCrLf
MsgHtml = MsgHtml & "</tr>" & VbCrLf


DetailRS.Close
Set DetailRS = Nothing 

Dim ReasonDescription
Dim GroupDescription
'## Get Item Qc Details
RedoItemData

If RedoItemDataRs.BOF = True Or RedoItemDataRs.EOF = True Then
    RedoItemDataRs.Close
    Set RedoItemDataRs = Nothing        
    Err.Clear
    Response.Redirect "Error.asp"
End If

'## Data Loop
While Not RedoItemDataRs.EOF   

    GroupDescription = GetGroupDesc (RedoItemDataRs("GroupId"))
    ReasonDescription = GetReasonDesc (RedoItemDataRs("ReasonCode"))
    
    If GroupDescription = "Printing" And RedoItemDataRs("WorkCentreName") = "Heat Press" Then
        GroupDescription = "Heat Press"
    End If
    
    ReasonTxtColour = "#000000"
    
    If Instr(1,ReasonDescription,"Other",1) > 0 Then 
        If RedoItemDataRs("Reason") <> "" Then
            ReasonDescription = GroupDescription & "&nbsp;(Other)<br/>" & RedoItemDataRs("Reason")        
        Else
            ReasonTxtColour = "#FF0000"
            ReasonDescription = GroupDescription & "&nbsp;(Other)&nbsp;Selected<br />But&nbsp;no&nbsp;reason&nbsp;was&nbsp;entered."
        End If
    Else
        If RedoItemDataRs("Reason") <> "" Then
            ReasonDescription = ReasonDescription & ".<br/>" & RedoItemDataRs("Reason")       
        End If    
    End If  
    
    If Len(ReasonDescription) > 60 Then
        ReasonDescription = Left(ReasonDescription,50) & " ..."
    End If      
       
    DetailDescription = RedoItemDataRs("ItemDescription")
    DetailDescription = Replace(DetailDescription,"<br />","&nbsp;",1,-1,1)
    If Len(DetailDescription) > 70 Then DetailDescription = Left(DetailDescription,60) & " ..."
    
    DisplaySub = Replace(RedoItemDataRs("Substrates"),"#","<br />",1,-1,1)
     
    MsgHtml = MsgHtml & "<tr>" & VbCrLf                     
    MsgHtml = MsgHtml & "<td width='25' class='styleTDleft' >" & RedoItemDataRs("ItemAlpha") & "</td>" & VbCrLf
    MsgHtml = MsgHtml & "<td class='styleTDstd'>" & RedoItemDataRs("SpId") & "</td>" & VbCrLf             
    MsgHtml = MsgHtml & "<td class='styleTDstd' align='left' >" & Trim(DetailDescription) & "</td>" & VbCrLf
    MsgHtml = MsgHtml & "<td class='styleTDstd'>" & DisplaySub & "</td>" & VbCrLf                       
    MsgHtml = MsgHtml & "<td class='styleTDstd'>" & RedoItemDataRs("NewQty") & "</td>" & VbCrLf
    
    MsgHtml = MsgHtml & "<td class='styleTDstd' align='left' >" & RedoItemDataRs("ItemNewSqm") & "</td>" & VbCrLf       
    MsgHtml = MsgHtml & "<td class='styleTDstd'>" & GroupDescription & "</td>" & VbCrLf
    MsgHtml = MsgHtml & "<td class='styleTDstd'>" & RedoItemDataRs("Printer") & "</td>" & VbCrLf
    MsgHtml = MsgHtml & "<td class='styleTDstd' align='left' style='color: " & ReasonTxtColour & "'>" & ReasonDescription & "</td>" & VbCrLf
        
    MsgHtml = MsgHtml & "</tr>" & VbCrLf

    ItemTitle = ""
    DetailDescription = ""
    ReasonDescription = ""
    GroupDescription = ""
    DisplaySub = ""

    RedoItemDataRs.MoveNext          
          
Wend 
    
RedoItemDataRs.Close
Set RedoItemDataRs = Nothing

'## Spacing at end so final row has bottom line, may not be needed
MsgHtml = MsgHtml & "<tr>"
MsgHtml = MsgHtml & "<td style='height: 15px' colspan='9'>&nbsp;</td>"
MsgHtml = MsgHtml & "</tr>"

MsgHtml = MsgHtml & "</table>" & VbCrLf
MsgHtml = MsgHtml & "</body>" & VbCrLf
MsgHtml = MsgHtml & "</html>" & VbCrLf

CreateMsg = MsgHtml

End Function

%>