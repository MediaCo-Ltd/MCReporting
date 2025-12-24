<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%> 

 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1

Dim Count
Dim Shifts()
Dim ShiftRs
Dim ShiftSql
Dim Nodata
Dim ReturnData
Dim SentType
Dim SentReason
Dim SentUser

Nodata = False
SentType = Cint(Request.Form("xType"))
SentReason = Cint(Request.Form("xReason"))
SentUser = Cint(Request.Form("xUser"))

'## Common to all
ShiftSql = "SELECT DISTINCT SelectedTimeSlot As TimeId," 
ShiftSql = ShiftSql & " CASE WHEN SelectedTimeSlot = 1 THEN 'Earlies' ELSE CASE WHEN SelectedTimeSlot = 2 THEN 'Lates' ELSE 'Nights' END END AS Shift"
ShiftSql = ShiftSql & " FROM Logs Where (Id >= 1 )"

If SentType = 4 Then
    '## Do Nothing
Else
    ShiftSql = ShiftSql & " AND (GroupId = " & SentType & ")"
End If

If SentReason > 0 Then
    ShiftSql = ShiftSql & " AND (GroupSelection = " & SentReason & ")"        
End If

If SentUser > 0 Then
    ShiftSql = ShiftSql & " AND (UserId = " & SentUser & ")"        
End If

ShiftSql = ShiftSql & " ORDER BY TimeId" 

Set ShiftRs = Server.CreateObject("ADODB.Recordset")
ShiftRs.ActiveConnection = Session("ConnHSReports")
ShiftRs.Source = ShiftSql
ShiftRs.CursorType = Application("adOpenForwardOnly")
ShiftRs.CursorLocation = Application("adUseClient")
ShiftRs.LockType = Application("adLockReadOnly")
ShiftRs.Open
Set ShiftRs.ActiveConnection = Nothing


If ShiftRs.BOF = True Or ShiftRs.EOF = True Then
    Nodata = True
Else
    Redim Shifts(ShiftRs.RecordCount,1)
    While Not ShiftRs.EOF
        Shifts(ShiftRs.AbsolutePosition,0) = ShiftRs("TimeId")
        Shifts(ShiftRs.AbsolutePosition,1) = ShiftRs("Shift")
        ShiftRs.MoveNext
    Wend        
End If

ShiftRs.Close
Set ShiftRs = Nothing

If Nodata = False Then

    ReturnData = "<option value='' >Select&nbsp;Shift</option>"
    For Count = 1 To Ubound(Shifts)
        ReturnData = ReturnData & ("<option value='" & Cstr(Shifts(Count,0)) & "'>" & Trim(Shifts(Count,1)) & "</option>)") & VbCrLf
    Next
    Erase Shifts

End If

Response.Write ReturnData

%>
