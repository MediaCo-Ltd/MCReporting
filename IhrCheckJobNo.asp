<%@language="vbscript" codepage="1252" EnableSessionState="True"%>
<%Option Explicit%>

<!--#include file="..\##GlobalFiles\PkId.asp" -->
<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->

<%
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp")
If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim ExtraMsg
Dim JobNo
Dim CheckSql
Dim Check_Data
Dim ProdJobId

JobNo = Trim(Request.Form("txtJobNo"))

If JobNo = "" Then Response.Redirect("IhrJobNo.asp")

If Ucase(Left(JobNo,3)) = "REF" THEN
    JobNo = Mid(JobNo,4)
End If

If IsNumeric(JobNo) Then

    If Session("Location") = "WorkTest" Or Session("Location") = "Home" Then
        CheckSql = "SELECT reference, status, Quotation.id, JobTypeId, JobId FROM Quotation where (Reference = 'REF" & JobNo & "') And (status In (3,13,5,6))"
    Else
        CheckSql = "SELECT reference, status, Quotation.id, JobTypeId, JobId FROM Quotation where (Reference = 'REF" & JobNo & "') And (status In (3,13))"
    End If
    
    Set Check_Data = Server.CreateObject("ADODB.Recordset")

    Check_Data.ActiveConnection = Session("ClarityConn")
    Check_Data.Source = CheckSql
    Check_Data.CursorType = Application("adOpenForwardOnly")
    Check_Data.CursorLocation = Application("adUseClient")
    Check_Data.LockType = Application("adLockReadOnly")
    Check_Data.Open
    Set Check_Data.ActiveConnection = Nothing

    If Check_Data.Bof Or Check_Data.Eof Then    
        Session("JobNoError") = "Job No. " & Trim(Request.Form("txtJobNo")) & " does not exist or is not a live job" 
        Check_Data.Close
        Set Check_Data = Nothing
	    Response.Redirect("IhrJobNo.asp")		
    End If
    
    If Check_Data("JobTypeId") <> 1 Then
        Session("JobNoError") = "Job No. " & Trim(Request.Form("txtJobNo")) & " is an Outdoor Job"
        Check_Data.Close
        Set Check_Data = Nothing
        Response.Redirect("IhrJobNo.asp")
    End If
    
    ProdJobId = Check_Data("JobId")

    Session("JobNoError") = ""
    Session("JobNo") = Check_Data("id")
        
    Check_Data.Close
    Set Check_Data = Nothing
    
    If ChkonHold(ProdJobId) = False Then
        Session("JobNoError") = "Job No. " & Request.Form("txtJobNo") & " has no operations on hold" & VbCrLf & "Put your operation on on hold in Clarity first"
	    Response.Redirect("IhrJobNo.asp")
	Else
        Response.Redirect("IhrAddData.asp")
    End If
Else
    '## Non numeric data entered
    Session("JobNoError") = "Job No. " & Request.Form("txtJobNo") & " is not a valid job number !"
	Response.Redirect("IhrJobNo.asp")
	
End If

'#################################

Function ChkOnHold (Qjid)

    Dim ChkOnHoldRs
    Dim ChkOnHoldSql
    Dim ChkRresult
    ChkRresult = False

    'ChkOnHoldSql = "SELECT vwProdOperations.JobRef, vwProdOperations.OPStatus, vwProdOperations.CurrentVersion,"
    'ChkOnHoldSql = ChkOnHoldSql & " vwProdOperations.JobId, Quotation.id"
    'ChkOnHoldSql = ChkOnHoldSql & " FROM vwProdOperations INNER JOIN"
    'ChkOnHoldSql = ChkOnHoldSql & " Quotation ON vwProdOperations.JobId = Quotation.JobId"
    'ChkOnHoldSql = ChkOnHoldSql & " WHERE (vwProdOperations.CurrentVersion = 1) AND (vwProdOperations.OPStatus = 2)" 
    'ChkOnHoldSql = ChkOnHoldSql & " AND (Quotation.id = " & Qid & ")"

    ChkOnHoldSql = "SELECT JobRef, OPStatus, CurrentVersion, JobId"
    ChkOnHoldSql = ChkOnHoldSql & " FROM vwProdOperations"
    ChkOnHoldSql = ChkOnHoldSql & " WHERE (CurrentVersion = 1) AND (OPStatus = 2) AND (NOT(WorkCentreId = 14)) AND (JobId = " & Qjid & ")"

    Set ChkOnHoldRs = Server.CreateObject("ADODB.Recordset")

    ChkOnHoldRs.ActiveConnection = Session("ClarityConn")
    ChkOnHoldRs.Source = ChkOnHoldSql
    ChkOnHoldRs.CursorType = Application("adOpenForwardOnly")
    ChkOnHoldRs.CursorLocation = Application("adUseClient")
    ChkOnHoldRs.LockType = Application("adLockReadOnly")
    ChkOnHoldRs.Open
    Set ChkOnHoldRs.ActiveConnection = Nothing
    
    If ChkOnHoldRs.BOF = True Or ChkOnHoldRs.EOF = True Then
        ChkRresult = False
    Else
        ChkRresult = True
    End If
    
    ChkOnHoldRs.Close
    Set ChkOnHoldRs = Nothing
    
    '## Overide for me
    If Session("UserId") = 4 Then ChkRresult = True

    ChkOnHold = ChkRresult

End Function

%>