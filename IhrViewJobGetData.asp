<% 

Dim errMsg
Dim DueDateRS
Dim ClarityRS
Dim OperationRs
Dim ClaritySqmRs
Dim PrinterDateRS
Dim PrinterDataRS
Dim ClientJobsRS
Dim strClientJobsSql
Dim strClientSql
Dim strDateSql
Dim strClaritySql
Dim strPrinterSql
Dim strOperationSql
Dim strClaritySqmSql
Dim strPrinterDateSql
Dim strClarityDispWipSql
Dim strPrinterQuoteItemsSql

Dim MyStartDate
Dim MyEndDate

errMsg = ""

'##############################################

' Set for Mgs (Quotation.JobTypeId = 1)

'## Join in Main & Single line printer Was
'## FROM Job INNER JOIN Quotation ON Job.Reference = Quotation.reference
'## Changed to
'## FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId

'Session("JobType") = "MGS"
'Session("JobTypeNo") = 1

'################################## Date Sql = Distinct Req' Dates with time removed

strDateSql = "SELECT DISTINCT LEFT(RequiredDate, 11) AS DateRequired, DAY(RequiredDate) AS DayRequired,"
strDateSql = strDateSql & " MONTH(RequiredDate) AS MonthRequired, YEAR(RequiredDate) AS YearRequired,"
strDateSql = strDateSql & " DATENAME(weekday, Quotation.requiredDate) AS DayName"
strDateSql = strDateSql & " FROM Quotation"
strDateSql = strDateSql & " WHERE (status IN (3, 13)) AND (NOT (LEFT(requiredDate, 11) IS NULL))"  
strDateSql = strDateSql & " AND (Quotation.JobTypeId = " & Session("JobTypeNo") & ")"
strDateSql = strDateSql & " ORDER BY YearRequired, MonthRequired, DayRequired"

'############################################## Main Clarity Sql

strClaritySql = "SELECT Distinct Job.Reference AS JobRef, Job.Id AS JobId, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId, Quotation.Description AS JobDesc, Quotation.createdBy, Contacts.Forename AS ConfirmedName,"
strClaritySql = strClaritySql & " vwContacts.CompanyName, Job.RouteStatus AS JobRoute, Job.Status AS JobStatus, Quotation.status AS QuoteStatus, Quotation.user3 as Fabrication, Quotation.DeliveryContact As DPCode,"

strClaritySql = strClaritySql & " Contacts.Forename + ' ' + Contacts.Surname AS ConfirmedByFull, Quotation.user10 As Undeliverable, Quotation.user6 As DispMethod,"

'## Due to way had to get outer date loop need this
strClaritySql = strClaritySql & " LTRIM(LEFT(Quotation.requiredDate, 11)) AS DateExpr"

strClaritySql = strClaritySql & " FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId"
strClaritySql = strClaritySql & " INNER JOIN vwContacts ON Job.ContactId = vwContacts.ContactId"
strClaritySql = strClaritySql & " INNER JOIN Contacts ON Quotation.createdBy = Contacts.Id"

'### Only get jobs that have a job card, Scheduled, unscheduled, W.I.P. or part delivered
strClaritySql = strClaritySql & " WHERE (Quotation.status In (3,13)) AND (NOT (Job.RouteStatus IS NULL))"
strClaritySql = strClaritySql & " AND (Quotation.JobTypeId = " & Session("JobTypeNo") & ")"  
'strClaritySql = strClaritySql & " AND (Quotation.User10 <> 'True')" 

'############################################## New Customer By Id

strClientSql = "SELECT DISTINCT Job.Reference AS JobRef, Job.Id AS JobId, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId," 
strClientSql = strClientSql & " Quotation.description AS JobDesc, Quotation.createdBy, Contacts.Forename AS ConfirmedName,"
strClientSql = strClientSql & " Job.RouteStatus AS JobRoute, Job.Status AS JobStatus, Quotation.status AS QuoteStatus,"
strClientSql = strClientSql & " Quotation.User3 AS Fabrication, Quotation.DeliveryContact AS DPCode, Quotation.user6 As DispMethod,"
strClientSql = strClientSql & " Contacts.Forename + ' ' + Contacts.Surname AS ConfirmedByFull, Quotation.User10 AS Undeliverable,"
strClientSql = strClientSql & " LTRIM(LEFT(Quotation.requiredDate, 11)) AS DateExpr, Company.Name As CompanyName, Company.Id"


strClientSql = strClientSql & " FROM Job INNER JOIN"
strClientSql = strClientSql & " Quotation ON Job.Id = Quotation.JobId INNER JOIN"
strClientSql = strClientSql & " Contacts ON Quotation.createdBy = Contacts.Id INNER JOIN"
strClientSql = strClientSql & " Contacts AS Contacts_1 ON Quotation.contactId = Contacts_1.Id INNER JOIN"
strClientSql = strClientSql & " Company ON Contacts_1.CompanyId = Company.Id"

'### Only get jobs that have a job card, Scheduled, unscheduled, W.I.P. or part delivered
strClientSql = strClientSql & " WHERE (Quotation.status In (3,13)) AND (NOT (Job.RouteStatus IS NULL))"
strClientSql = strClientSql & " AND (Quotation.JobTypeId = " & Session("JobTypeNo") & ")"  	

'############################################## Get jobs that are deliverd but not finished

strClarityDispWipSql = "SELECT Distinct Job.Reference AS JobRef, Job.Id AS JobId, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId, Quotation.Description AS JobDesc, Quotation.createdBy, Contacts.Forename AS ConfirmedName,"
strClarityDispWipSql = strClarityDispWipSql & " vwContacts.CompanyName, Job.RouteStatus AS JobRoute, Job.Status AS JobStatus, Quotation.status AS QuoteStatus, Quotation.user3 as Fabrication,"
'## Due to way had to get outer date loop need this
strClarityDispWipSql = strClarityDispWipSql & " LTRIM(LEFT(Quotation.requiredDate, 11)) AS DateExpr"

strClarityDispWipSql = strClarityDispWipSql & " FROM Job INNER JOIN Quotation ON Job.Id = Quotation.JobId"
strClarityDispWipSql = strClarityDispWipSql & " INNER JOIN vwContacts ON Job.ContactId = vwContacts.ContactId"
strClarityDispWipSql = strClarityDispWipSql & " INNER JOIN Contacts ON Quotation.createdBy = Contacts.Id"

'### Only get jobs that have a job card, Scheduled, unscheduled, W.I.P. or part delivered
strClarityDispWipSql = strClarityDispWipSql & " WHERE (Quotation.status In (5,6)) AND (Job.RouteStatus <> 3)"
strClarityDispWipSql = strClarityDispWipSql & " AND (Quotation.JobTypeId = " & Session("JobTypeNo") & ") ORDER BY JobRef"  	


'############################################## Get single line Printer Sql

strPrinterDateSql = "SELECT DISTINCT "                                             
strPrinterDateSql = strPrinterDateSql & " vwProdGantt.JobId, vwProdGantt.WorkCentre, vwProdGantt.JobRef, vwProdGantt.CurrentVersion,"
strPrinterDateSql = strPrinterDateSql & " vwProdGantt.JobStatus, Quotation.status, Quotation.id AS QuoteID, DATENAME(weekday, Quotation.requiredDate) AS DayName," 
strPrinterDateSql = strPrinterDateSql & " Quotation.requiredDate, LTRIM(LEFT(Quotation.requiredDate, 11)) AS DateExpr,"
strPrinterDateSql = strPrinterDateSql & " Quotation.user3 as Fabrication,Quotation.Description AS JobDesc"

strPrinterDateSql = strPrinterDateSql & " FROM vwProdGantt INNER JOIN Quotation ON vwProdGantt.JobId = Quotation.JobId"

strPrinterDateSql = strPrinterDateSql & " WHERE (vwProdGantt.CurrentVersion = 1) AND (vwProdGantt.JobStatus <= 3)"
strPrinterDateSql = strPrinterDateSql & " AND (Quotation.JobTypeId = " & Session("JobTypeNo") & ") AND (Quotation.status IN (3, 13))"
'strPrinterDateSql = strPrinterDateSql & " AND (Quotation.User10 <> 'True')"


'############################################## Get QuoteItems For Printer

strPrinterQuoteItemsSql = "SELECT  ProdJobCards.QuoteItemId, ProdJobCards.JobId, ProdJobCards.Id AS RCid,ProdJobCards.Description AS ItemName,"
strPrinterQuoteItemsSql = strPrinterQuoteItemsSql  & " ProdOperations.WorkCentreId, ProdWorkCentres.Description AS WorkCentre"
strPrinterQuoteItemsSql = strPrinterQuoteItemsSql  & " FROM ProdJobCards "
strPrinterQuoteItemsSql = strPrinterQuoteItemsSql  & " INNER JOIN ProdOperations ON ProdJobCards.Id = ProdOperations.ParentId"
strPrinterQuoteItemsSql = strPrinterQuoteItemsSql  & " INNER JOIN ProdWorkCentres ON ProdOperations.WorkCentreId = ProdWorkCentres.Id"

'############################# Get All Customer Jobs Invoced Or delivered, also used by DPC

strClientJobsSql = "SELECT DISTINCT Job.Reference AS JobRef, Job.Id AS JobId, Quotation.requiredDate AS ReqDate, Quotation.id AS QuoteId," 
strClientJobsSql = strClientJobsSql & " Quotation.description AS JobDesc, Quotation.createdBy, Contacts.Forename AS ConfirmedName,"
strClientJobsSql = strClientJobsSql & " Job.RouteStatus AS JobRoute, Job.Status AS JobStatus, Quotation.status AS QuoteStatus,"
strClientJobsSql = strClientJobsSql & " Quotation.User3 AS Fabrication, Quotation.DeliveryContact AS DPCode,LEFT(confirmedDate, 11) AS ConfDate,"
strClientJobsSql = strClientJobsSql & " Contacts.Forename + ' ' + Contacts.Surname AS ConfirmedByFull, Quotation.User10 AS Undeliverable,"
strClientJobsSql = strClientJobsSql & " LTRIM(LEFT(Quotation.requiredDate, 11)) AS DateExpr, Company.Name As CompanyName, Company.Id,"
strClientJobsSql = strClientJobsSql & " LEFT(Quotation.completedDate, 11) AS DelDate, Quotation.OrderNo, Quotation.User4 As Redo,"
strClientJobsSql = strClientJobsSql & " Quotation.user12 As ClientRedo,"
 
'strClientJobsSql = strClientJobsSql & " CASE WHEN LEFT(Quotation.completedDate, 11) > LEFT(Quotation.requiredDate, 11)" 
'strClientJobsSql = strClientJobsSql & " THEN CAST(1 AS bit) ELSE CAST(0 AS bit) END AS OverDue,"

strClientJobsSql = strClientJobsSql & " CASE WHEN datediff(d, Quotation.requiredDate, Quotation.completedDate) > 0 THEN CAST(1 AS int)" 
strClientJobsSql = strClientJobsSql & " ELSE CASE WHEN datediff(d, Quotation.requiredDate, Quotation.completedDate) = 0 THEN CAST(0 AS int)" 
strClientJobsSql = strClientJobsSql & " ELSE CAST(- 1 AS int) END END AS DelStatus, DATEDIFF(d, requiredDate, completedDate) AS DelDif,"

strClientJobsSql = strClientJobsSql & " MONTH(confirmedDate) AS MonthRequired, YEAR(confirmedDate) AS YearRequired, DAY(confirmedDate) AS DayRequired"

strClientJobsSql = strClientJobsSql & " FROM Job INNER JOIN"
strClientJobsSql = strClientJobsSql & " Quotation ON Job.Id = Quotation.JobId INNER JOIN"
strClientJobsSql = strClientJobsSql & " Contacts ON Quotation.createdBy = Contacts.Id INNER JOIN"
strClientJobsSql = strClientJobsSql & " Contacts AS Contacts_1 ON Quotation.contactId = Contacts_1.Id INNER JOIN"
strClientJobsSql = strClientJobsSql & " Company ON Contacts_1.CompanyId = Company.Id"

'### Only get jobs that have a job card, Scheduled, unscheduled, W.I.P. or part delivered
strClientJobsSql = strClientJobsSql & " WHERE (Quotation.status In (5,6)) AND (NOT (Job.RouteStatus IS NULL))"
strClientJobsSql = strClientJobsSql & " AND (Quotation.JobTypeId = " & Session("JobTypeNo") & ")" 

'############################################################################################################## 

Sub OpenDueDateRS 
    
    Set DueDateRS = Server.CreateObject("ADODB.Recordset")

	DueDateRS.ActiveConnection = Session("ClarityConn") 'strConnClarity 
	DueDateRS.Source = strDateSql  	
	DueDateRS.CursorType = Application("adOpenForwardOnly")
	DueDateRS.CursorLocation = Application("adUseClient")
	DueDateRS.LockType = Application("adLockReadOnly")
	DueDateRS.Open
	Set DueDateRS.ActiveConnection = Nothing
	
	'Response.Write "In open Due Date"	

End Sub

Sub OpenClarityRS(RsFilter)

    Set ClarityRS = Server.CreateObject("ADODB.Recordset")

    ClarityRS.ActiveConnection = Session("ClarityConn") 'strConnClarity
       
    ClarityRS.Source = strClaritySql  & _
    " AND (LEFT(Quotation.requiredDate, 11) like '" & RsFilter & "') ORDER BY JobRef"    
 	
    ClarityRS.CursorType = Application("adOpenForwardOnly")
    ClarityRS.CursorLocation = Application("adUseClient")
    ClarityRS.LockType = Application("adLockReadOnly")
    ClarityRS.Open
    Set ClarityRS.ActiveConnection = Nothing	
		        
End Sub

Sub OpenCreatedByRS(RsFilter,AccountFilter)

    Set ClarityRS = Server.CreateObject("ADODB.Recordset")

    ClarityRS.ActiveConnection = Session("ClarityConn") 'strConnClarity
    
    ClarityRS.Source = strClaritySql  & _
    " AND (LEFT(Quotation.requiredDate, 11) like '" & RsFilter & "')" & _ 
    " AND (Quotation.createdBy like " & AccountFilter & ") ORDER BY JobRef"
    
    '## 16788325 Me 
    '## 704647167 Simon Gabriel 
    '## 33554433 Karen Sewell 
    '## 33649307 Simon O'Donnell
    '## 50340658 Craig Howard
    '## 201344568 Stephen Arthur
    '## 704663124 Hena Sadler
     	
    ClarityRS.CursorType = Application("adOpenForwardOnly")
    ClarityRS.CursorLocation = Application("adUseClient")
    ClarityRS.LockType = Application("adLockReadOnly")
    ClarityRS.Open
    Set ClarityRS.ActiveConnection = Nothing	
		        
End Sub

Sub OpenClarityDispatchWipRS()

    'http://127.0.0.1/wip/ClarityReportMgsDispatchedWip.asp

    Set ClarityRS = Server.CreateObject("ADODB.Recordset")

    ClarityRS.ActiveConnection = Session("ClarityConn") 'strConnClarity
       
    ClarityRS.Source = strClarityDispWipSql  
 	
    ClarityRS.CursorType = Application("adOpenForwardOnly")
    ClarityRS.CursorLocation = Application("adUseClient")
    ClarityRS.LockType = Application("adLockReadOnly")
    ClarityRS.Open
    Set ClarityRS.ActiveConnection = Nothing	
		        
End Sub

Sub OpenCustomerRS(RsFilter,CustFilter)

    Set ClarityRS = Server.CreateObject("ADODB.Recordset")

    ClarityRS.ActiveConnection = Session("ClarityConn") 'strConnClarity
    
    ClarityRS.Source = strClientSql  & _
    " AND (LEFT(Quotation.requiredDate, 11) like '" & RsFilter & "')" & _ 
    " AND (Company.Id = " & CustFilter & ") ORDER BY JobRef"    
    
 	
    ClarityRS.CursorType = Application("adOpenForwardOnly")
    ClarityRS.CursorLocation = Application("adUseClient")
    ClarityRS.LockType = Application("adLockReadOnly")
    ClarityRS.Open
    Set ClarityRS.ActiveConnection = Nothing	
		        
End Sub

'## this needs to be changed to allow by Company.Id confirmed by ??????????? deliverd date ???
Sub OpenCustomerJobsRS()

    MyStartDate = Session("StartDate")
    MyStartDate = Mid(MyStartDate, 4, 3) & Left(MyStartDate, 3) & Right(MyStartDate, 4) & " 00:00:00"
    MyEndDate = Session("EndDate")
    MyEndDate = Mid(MyEndDate, 4, 3) & Left(MyEndDate, 3) & Right(MyEndDate, 4) & " 23:59:59"
    
    strClientJobsSql = strClientJobsSql & " AND (Company.Id = '" & Session("Customer") & "')"
    strClientJobsSql = strClientJobsSql & " AND (confirmedDate BETWEEN '" & MyStartDate & "' AND ' " & MyEndDate & "')"
    strClientJobsSql = strClientJobsSql & " ORDER BY YearRequired, MonthRequired, DayRequired"

    Set ClarityRS = Server.CreateObject("ADODB.Recordset")
    ClarityRS.ActiveConnection = Session("ClarityConn") 'strConnClarity  
 	ClarityRS.Source = strClientJobsSql 	
    ClarityRS.CursorType = Application("adOpenForwardOnly")
    ClarityRS.CursorLocation = Application("adUseClient")
    ClarityRS.LockType = Application("adLockReadOnly")
    ClarityRS.Open
    Set ClarityRS.ActiveConnection = Nothing
		        
End Sub

'## work on this by id ????
Sub OpenDpcRS()

    MyStartDate = Session("StartDate")
    MyStartDate = Mid(MyStartDate, 4, 3) & Left(MyStartDate, 3) & Right(MyStartDate, 4) & " 00:00:00"
    MyEndDate = Session("EndDate")
    MyEndDate = Mid(MyEndDate, 4, 3) & Left(MyEndDate, 3) & Right(MyEndDate, 4) & " 23:59:59"
    
    strClientJobsSql = strClientJobsSql & " AND (DeliveryContact Like '" & Session("Customer") & "')"
    strClientJobsSql = strClientJobsSql & " AND (confirmedDate BETWEEN '" & MyStartDate & "' AND ' " & MyEndDate & "')"
    strClientJobsSql = strClientJobsSql & " ORDER BY YearRequired, MonthRequired, DayRequired"

    Set ClarityRS = Server.CreateObject("ADODB.Recordset")
    ClarityRS.ActiveConnection = Session("ClarityConn") 'strConnClarity  
 	ClarityRS.Source = strClientJobsSql 	
    ClarityRS.CursorType = Application("adOpenForwardOnly")
    ClarityRS.CursorLocation = Application("adUseClient")
    ClarityRS.LockType = Application("adLockReadOnly")
    ClarityRS.Open
    Set ClarityRS.ActiveConnection = Nothing
		        
End Sub

Sub OpenPrinterDateRS(DateFilter,PrinterFilter)
    
    Set PrinterDateRS = Server.CreateObject("ADODB.Recordset")

    PrinterDateRS.ActiveConnection = Session("ClarityConn") 'strConnClarity    
    PrinterDateRS.Source = strPrinterDateSql & _
    " AND (LEFT(Quotation.requiredDate, 11) like '" & DateFilter & "')" & _
    " AND (vwProdGantt.WorkCentre Like '" & PrinterFilter & "%') ORDER BY JobRef" 
 	
    PrinterDateRS.CursorType = Application("adOpenForwardOnly")
    PrinterDateRS.CursorLocation = Application("adUseClient")
    PrinterDateRS.LockType = Application("adLockReadOnly")
    PrinterDateRS.Open
    Set PrinterDateRS.ActiveConnection = Nothing	
		        
End Sub

Sub OpenPrinterDataRS(JobId,PrinterFilter)  

    Set PrinterDataRS = Server.CreateObject("ADODB.Recordset")

    PrinterDataRS.ActiveConnection = Session("ClarityConn") 'strConnClarity
    PrinterDataRS.Source = strPrinterQuoteItemsSql &_	
 	" Where (ProdJobCards.JobId ='" & JobId & "') And (ProdWorkCentres.Description Like '" &_
 	PrinterFilter & "%') Order By ProdJobCards.Description"
 	'## Seems to be working ok 
 	
    PrinterDataRS.CursorType = Application("adOpenForwardOnly")
    PrinterDataRS.CursorLocation = Application("adUseClient")
    PrinterDataRS.LockType = Application("adLockReadOnly")
    PrinterDataRS.Open
    Set PrinterDataRS.ActiveConnection = Nothing	

    '## Has to be done as filter  ?????????????????????
   ' PrinterDataRS.Filter = "JobId =" & JobId & " And WorkCentre Like '" & PrinterFilter & "*'" 
		        
End Sub

Function GetPrice(Qid)

    Dim PriceRs
    Dim PriceSql
    Dim LocalPrice
    
    LocalPrice = 0
    PriceSql = "SELECT TotalPrice, QuotationId FROM vwQuotationTotals"
    PriceSql = PriceSql & " WHERE (QuotationId = '" & Qid & "')"

    Set PriceRs = Server.CreateObject("ADODB.Recordset")
    PriceRs.ActiveConnection = Session("ClarityConn") 'strConnClarity
    PriceRs.Source = PriceSql
    
    PriceRs.CursorType = Application("adOpenForwardOnly")
    PriceRs.CursorLocation = Application("adUseClient")
    PriceRs.LockType = Application("adLockReadOnly")
    PriceRs.Open
    Set PriceRs.ActiveConnection = Nothing
    
    If PriceRs.BOF = True Or PriceRs.EOF Then
        LocalPrice = 0
    Else
        LocalPrice = PriceRs("TotalPrice")
    End If   
    
    PriceRs.Close
    Set PriceRs = Nothing
    
    GetPrice = LocalPrice    

End Function		
%>