<%@language="vbscript" codepage="1252"%>
<%Option Explicit%>


<!--#include file="..\##GlobalFiles\DefaultColours.asp" -->
<!--#include file="..\##GlobalFiles\connClarityDB.asp" -->
<!--#include file="..\##GlobalFiles\PkId.asp" -->


<%
'## Check if System is locked
Server.Execute "SystemLockedCheck.asp"
If Session("Locked") = True Then Response.Redirect ("SystemLocked.asp") 

Dim ExtraMsg
Dim JobNo
Dim CheckSql
Dim Check_Flag
Dim Check_Data


JobNo = Request.Form("txtJobNo")

If Ucase(Left(JobNo,3)) = "REF" THEN
    JobNo = Mid(JobNo,4)
End If

If IsNumeric(JobNo) Then

    CheckSql = "SELECT reference, status, Quotation.id, JobTypeId FROM Quotation where (Reference = 'REF" & JobNo & "' And (Quotation.status In (3,5,6,13)))"
    'CheckSql = CheckSql & "And (JobTypeId = " & Session("JobTypeNo") & ")"


    Check_flag="ADODB.Recordset"
    Set Check_Data = Server.CreateObject(Check_flag)

    Check_Data.ActiveConnection = strConnClarity 
    Check_Data.Source = CheckSql
    Check_Data.CursorType = 0
    Check_Data.CursorLocation = 3
    Check_Data.LockType = 1
    Check_Data.Open
    Set Check_Data.ActiveConnection = Nothing

    If Check_Data.Bof Or Check_Data.Eof Then    
        Session("JobNoError") = "Job No. " & Request.Form("txtJobNo") & " does not exist !" '& vbcrlf & ExtraMsg
	    Check_Data.Close
        Set Check_Data = Nothing
	    Response.Redirect("NcJobNo.asp")		
    End If

    If Check_Data("JobTypeId") = 1 Then
        Session("JobType") = "MGS"
        Session("JobTypeNo") = MgsJobTypeId
    ElseIf Check_Data("JobTypeId") = 2 Then
        Session("JobType") = "OUT"
        Session("JobTypeNo") = OutJobTypeId
    Else
        Session("JobNoError") = "Job No. " & Request.Form("txtJobNo") & " Is a Outdoor Paper Job" '& vbcrlf & ExtraMsg
	    Check_Data.Close
        Set Check_Data = Nothing
        Session("JobType") = ""
        Session("JobTypeNo") = ""
	    Response.Redirect("NcJobNo.asp")
    End If
    
    Session("JobNoError") = ""
    Session("JobNo") = Check_Data("id")

    Check_Data.Close
    Set Check_Data = Nothing

    Response.Redirect("NcAddNewCL.asp")
Else
    '## Non numeric data entered
    If Ucase(Left(Request.Form("txtJobNo"),3)) = "REF" Then
        Session("JobNoError") = "Only enter the number !"
    Else
        Session("JobNoError") = "' " & Request.Form("txtJobNo") & " ' is not a valid job number !"
	End If
	Response.Redirect("NcJobNo.asp")
	
End If

%>