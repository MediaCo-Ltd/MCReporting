
<%

Dim errMsg
Dim DateSql
Dim DisplaySql
Dim DisplayMultiSql
Dim DeptSql
Dim ReasonCodeSql
Dim DeptRs
Dim ReasonCodeRs
Dim QuoteIdRs
Dim DateRangeSql


errMsg = ""

'On Error Resume Next

DateSql = "SELECT Distinct CreatedSerial FROM OhQuoteData Where (CreatedSerial > 1) Order by CreatedSerial DESC" 
DisplaySql = "SELECT * FROM OhQuoteData Where ("
DeptSql = "SELECT Distinct GroupName, GroupId FROM Department Where ("
ReasonCodeSql = "SELECT Distinct ReasonCode, ReasonDescription FROM ReasonCodes Where ("
DisplayMultiSql = "SELECT * FROM OhQuoteData Where ("


Sub OpenGetDatesRS ()
		
	Set GetDatesRS = Server.CreateObject("ADODB.Recordset")
	GetDatesRS.ActiveConnection = Session("ConnQC") 
	GetDatesRS.Source = DateSql
	GetDatesRS.CursorType = Application("adOpenForwardOnly")
	GetDatesRS.CursorLocation = Application("adUseClient")
	GetDatesRS.LockType = Application("adLockReadOnly")
	GetDatesRS.Open	
              
End Sub

'#####################################

Sub OpenMyLogsRs (RsFilterDateSerial)
	
	Set DisplayRs = Server.CreateObject("ADODB.Recordset")
	DisplayRs.ActiveConnection = Session("ConnQC") 
	DisplayRs.Source = DisplaySql & "CreatedSerial = " & RsFilterDateSerial & ") AND (CreatedById = " & Session("UserId") & ") Order By QuoteRef, Revision"      
	DisplayRs.CursorType = Application("adOpenForwardOnly")
	DisplayRs.CursorLocation = Application("adUseClient")
	DisplayRs.LockType = Application("adLockReadOnly")

	DisplayRs.Open	

End Sub

'#####################################

Sub OpenDisplayRs (RsFilterDateSerial)
	
	Set DisplayRs = Server.CreateObject("ADODB.Recordset")
	DisplayRs.ActiveConnection = Session("ConnQC") 
	DisplayRs.Source = DisplaySql & "CreatedSerial = " & RsFilterDateSerial & ") Order By QuoteRef, Revision"      
	DisplayRs.CursorType = Application("adOpenForwardOnly")
	DisplayRs.CursorLocation = Application("adUseClient")
	DisplayRs.LockType = Application("adLockReadOnly")

	DisplayRs.Open	

End Sub

'#####################################

Sub OpenDisplayByDateRs (RsFilterDateRange)
	
	Set DisplayRs = Server.CreateObject("ADODB.Recordset")
	DisplayRs.ActiveConnection = Session("ConnQC") 
	DisplayRs.Source = DisplaySql &  RsFilterDateRange & " Order By QuoteRef, Revision"     
	DisplayRs.CursorType = Application("adOpenForwardOnly")
	DisplayRs.CursorLocation = Application("adUseClient")
	DisplayRs.LockType = Application("adLockReadOnly")

	DisplayRs.Open	

End Sub

'#####################################

Sub OpenDisplayMultiRs (RsFilterMulti)
	
	Set DisplayMultiRs = Server.CreateObject("ADODB.Recordset")
	DisplayMultiRs.ActiveConnection = Session("ConnQC") 
	DisplayMultiRs.Source = DisplayMultiSql & "QuoteId = " & RsFilterMulti & ") Order By Revision"      
	DisplayMultiRs.CursorType = Application("adOpenForwardOnly")
	DisplayMultiRs.CursorLocation = Application("adUseClient")
	DisplayMultiRs.LockType = Application("adLockReadOnly")

	DisplayMultiRs.Open	

End Sub

'#####################################

Sub GetDepartments (RsFilterGroupId)

    DisplayDept = ""
    
    Set DeptRs = Server.CreateObject("ADODB.Recordset")
	DeptRs.ActiveConnection = Session("ConnQC") 
	DeptRs.Source = DeptSql & " GroupId In (" & RsFilterGroupId & ")) Order By GroupId"  ' order by GroupName DESC"  ' Order by JobNo DESC"     
	DeptRs.CursorType = Application("adOpenForwardOnly")
	DeptRs.CursorLocation = Application("adUseClient")
	DeptRs.LockType = Application("adLockReadOnly")

	DeptRs.Open
	
	While Not DeptRs.EOF
	    If DisplayDept = "" Then
	        DisplayDept = DeptRs("GroupName")
	    Else
	        DisplayDept = DisplayDept & ",&nbsp;" & DeptRs("GroupName")
	    End If	    
	    DeptRs.MoveNext
	Wend
    
    DeptRs.Close
    Set DeptRs = Nothing

End Sub

'#####################################

Sub GetReasons (RsFilterReasonId)

    DisplayReason = ""
    
    Set ReasonCodeRs = Server.CreateObject("ADODB.Recordset")
	ReasonCodeRs.ActiveConnection = Session("ConnQC") 
	ReasonCodeRs.Source = ReasonCodeSql & " ReasonCode IN( " & RsFilterReasonId & ")) Order By ReasonCode"  ' Order by JobNo DESC"     
	ReasonCodeRs.CursorType = Application("adOpenForwardOnly")
	ReasonCodeRs.CursorLocation = Application("adUseClient")
	ReasonCodeRs.LockType = Application("adLockReadOnly")

	ReasonCodeRs.Open
	
	While Not ReasonCodeRs.EOF
	    If DisplayReason = "" Then
	        If ReasonCodeRs("ReasonCode") = 205 Then 
	            DisplayReason = "Other (Studio)"
	        ElseIf ReasonCodeRs("ReasonCode") = 311 Then
	            DisplayReason = "Other (Printing)"
	        ElseIf ReasonCodeRs("ReasonCode") = 409 Then
	            DisplayReason = "Other (Finishing)"
	        ElseIf ReasonCodeRs("ReasonCode") = 505 Then
	            DisplayReason = "Other (Dispatch - Installs)"
	        ElseIf ReasonCodeRs("ReasonCode") = 602 Then
	            DisplayReason = "Other (Stock)"
	        Else
	            DisplayReason = ReasonCodeRs("ReasonDescription")
	        End If
	    Else
	        If ReasonCodeRs("ReasonCode") = 205 Then 
	            DisplayReason = DisplayReason & ",&nbsp;" & "Other (Studio)"
	        ElseIf ReasonCodeRs("ReasonCode") = 311 Then
	            DisplayReason = DisplayReason & ",&nbsp;" & "Other (Printing)"
	        ElseIf ReasonCodeRs("ReasonCode") = 409 Then
	            DisplayReason = DisplayReason & ",&nbsp;" & "Other (Finishing)"
	        ElseIf ReasonCodeRs("ReasonCode") = 505 Then
	            DisplayReason = DisplayReason & ",&nbsp;" & "Other (Dispatch - Installs)"
	        ElseIf ReasonCodeRs("ReasonCode") = 602 Then
	            DisplayReason = DisplayReason & ",&nbsp;" & "Other (Stock)"
	        Else
	            DisplayReason = DisplayReason & ",&nbsp;" & ReasonCodeRs("ReasonDescription")
	        End If
	    End If	    
	    ReasonCodeRs.MoveNext
	Wend
    
    ReasonCodeRs.Close
    Set ReasonCodeRs = Nothing


End Sub

'#####################################

Function GetQuoteIdRs (RsFilterQid)
		
	Dim LocalResult
	Set QuoteIdRs = Server.CreateObject("ADODB.Recordset")
	QuoteIdRs.ActiveConnection = Session("ConnQC") 
	QuoteIdRs.Source = DisplaySql & " Id = " & RsFilterQid & ")"
	QuoteIdRs.CursorType = Application("adOpenForwardOnly")
	QuoteIdRs.CursorLocation = Application("adUseClient")
	QuoteIdRs.LockType = Application("adLockReadOnly")
	QuoteIdRs.Open	
	
	LocalResult = QuoteIdRs("QuoteId")
	QuoteIdRs.Close
	Set QuoteIdRs = Nothing
	
	GetQuoteIdRs = LocalResult
           
End Function
		
%>