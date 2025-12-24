<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%>

 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1


Dim Count
Dim RecordId
Dim Status

Dim NoConnection
Dim ReturnData

If Session("ConnMcLogon") = "" Then
    NoConnection = True
    ReturnData = ""
Else
    NoConnection = False
End If

If NoConnection = False Then
    RecordId = Clng(ExtractOnlyNumbers(Request.Form("xRecordId")))
    Status = Cint(Request.Form("xStatus"))

    If IsNumeric(RecordId) Then

        Dim LogDataSql
        Dim LogDataRs

        LogDataSql = "SELECT  Id, Dormant From Logs Where (Id = " & RecordId  & ")"

        Set LogDataRs = Server.CreateObject("ADODB.Recordset")
        LogDataRs.ActiveConnection = Session("ConnMachinefaults")
        LogDataRs.Source = LogDataSql
        LogDataRs.CursorType = Application("adOpenStatic")
        LogDataRs.CursorLocation = Application("adUseClient") 
        LogDataRs.LockType = Application("adLockOptimistic")
        LogDataRs.Open
        
        If LogDataRs.BOF = True Or LogDataRs.EOF = True Then
            ReturnData = "Record " & RecordId & " doesn't exist"
        Else
            ReturnData = ""        
            If Status = 1 Then
                LogDataRs("Dormant") = 1
                LogDataRs.Update
                ReturnData = "Record " & RecordId & " set to dormant"        
            Else
                LogDataRs("Dormant") = 0
                LogDataRs.Update
                ReturnData = "Record " & RecordId & " is visible again" 
            End If    
        End If
        
        LogDataRs.Close
        Set LogDataRs = Nothing    
    Else
        ReturnData = "Enter a valid number"
    End If
End If

Response.Write ReturnData

'################################################################

Private Function ExtractOnlyNumbers(StrToCheck)

   ' Dim i 
    'Dim J 
    Dim PosToCheck 
    Dim ReturnValue 
    
    For Count = 1 To Len(StrToCheck)
        PosToCheck = Mid(StrToCheck, Count, 1)
        If IsNumeric(PosToCheck) Then
            ReturnValue = ReturnValue & PosToCheck
        End If
    Next
    
    ExtractOnlyNumbers = ReturnValue

End Function



%>
