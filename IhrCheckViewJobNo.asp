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
Dim JobId
Dim CheckSql
Dim Check_Data
Dim MultiQid


JobNo = Trim(Request.Form("txtJobNo"))

If JobNo = "" Then Response.Redirect("IhrSelectDate.asp")

If Ucase(Left(JobNo,3)) = "REF" THEN
    JobNo = Mid(JobNo,4)
End If

If IsNumeric(JobNo) Then
    
    CheckSql = "SELECT QuoteRef, QuoteId, ID FROM OhQuoteData where (QuoteRef = '" & JobNo & "')"    
    
    Set Check_Data = Server.CreateObject("ADODB.Recordset")

    Check_Data.ActiveConnection = Session("ConnQC")
    Check_Data.Source = CheckSql
    Check_Data.CursorType = Application("adOpenForwardOnly")
    Check_Data.CursorLocation = Application("adUseClient")
    Check_Data.LockType = Application("adLockReadOnly")
    Check_Data.Open
    Set Check_Data.ActiveConnection = Nothing

    If Check_Data.Bof Or Check_Data.Eof Then    
        Session("JobNoError") = "Job No. " & Trim(Request.Form("txtJobNo")) & " has no redo records" 
        Check_Data.Close
        Set Check_Data = Nothing
	    Response.Redirect("IhrSelectDate.asp")		
    End If
    
    If Check_Data.RecordCount > 1 Then
        MultiQid = Check_Data("QuoteId")
        'Session("JobNoError") = "Job No. " & Trim(Request.Form("txtJobNo")) & " has multiple redo records" & VbCrLf & "Select a date range to view all" 
        Check_Data.Close
        Set Check_Data = Nothing
	    Response.Redirect("IhrDisplayMulti.asp?Qid=" & MultiQid & "&qr=" & JobNo )    
    End If
    
    JobId = Check_Data("ID")

    Check_Data.Close
    Set Check_Data = Nothing
    
    
    
    

    Response.Redirect("IhrShowJob.asp?Rid=" & JobId)
'Else
    '## Non numeric data entered
'    Session("JobNoError") = "' " & Request.Form("txtJobNo") & " ' is not a valid job number !"
'	Response.Redirect("IhrJobNo.asp")
	
End If

%>