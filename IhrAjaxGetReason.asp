<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%> 
 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1

On Error Resume Next

Dim Count
Dim Nodata
Dim ReturnData
Dim GroupToGet
Dim WcCheck
Dim Selected
Dim ReturnRow
Dim AutoReturnCode
Dim Reasons()

GroupToGet = Cint(Request.Form("xGrpId"))
ReturnRow  = Cint(Request.Form("xRowId"))
WcCheck = Cstr(Request.Form("xWcId"))

ReturnData = ""
Selected = ""
AutoReturnCode = ""
Dim AutoGroupCode
AutoGroupCode = ""

Dim GetCodesRs
Dim GetCodesSql

GetCodesSql = "Select Id, QcGroupId, ReasonCode, ReasonDescription, WorkCentreCode From ReasonCodes Where (QcGroupId = " & GroupToGet & ")"

'GetCodesSql = "SelectReasonCodes.ID, ReasonCodes.QcGroupId, ReasonCodes.ReasonCode, ReasonCodes.ReasonDescription,"
'GetCodesSql = GetCodesSql & " ReasonCodes.WorkCentreCode, Department.GroupName"
'GetCodesSql = GetCodesSql & " FROM (ReasonCodes INNER JOIN"
'GetCodesSql = GetCodesSql & " Department ON ReasonCodes.QcGroupId = Department.GroupId)"
'GetCodesSql = GetCodesSql & " Where (QcGroupId = " & GroupToGet & ")"



If Instr(1,Session("ConnQC"),"MSDASQL",1) > 0 Then
    GetCodesSql = GetCodesSql & " AND (Active = 1 )"
Else
    GetCodesSql = GetCodesSql & " AND (Active = True )"
End If  
    

GetCodesSql = GetCodesSql & " ORDER BY ReasonCode"

Set GetCodesRs = Server.CreateObject("ADODB.Recordset")
GetCodesRs.ActiveConnection = Session("ConnQC")
GetCodesRs.Source = GetCodesSql

GetCodesRs.CursorType = Application("adOpenForwardOnly")
GetCodesRs.CursorLocation = Application("adUseClient")
GetCodesRs.LockType = Application("adLockReadOnly")
GetCodesRs.Open
Set GetCodesRs.ActiveConnection = Nothing


If GetCodesRs.BOF = True Or GetCodesRs.EOF = True Then
    Nodata = True
    ReturnData = "BOF"
End If

If Nodata = False Then

    Redim Reasons(GetCodesRs.RecordCount,2)    
    
    While Not GetCodesRs.EOF
        '## 102 Wrong info from client is disabled in Program so skip this
        If GetCodesRs("ReasonCode") = 102 Then 
            GetCodesRs.MoveNext
        Else
            Reasons(GetCodesRs.AbsolutePosition,0) = GetCodesRs("ReasonCode")
            Reasons(GetCodesRs.AbsolutePosition,1) = GetCodesRs("ReasonDescription")
            Reasons(GetCodesRs.AbsolutePosition,2) = GetCodesRs("WorkCentreCode")
            'Reasons(GetCodesRs.AbsolutePosition,3) = GetCodesRs("GroupName")
            GetCodesRs.MoveNext
        End If
    Wend
    
    GetCodesRs.Close
    Set GetCodesRs = Nothing
    
    ReturnData = ReturnRow & "#"
    
    ReturnData = ReturnData & "<option value='' >Select&nbsp;Reason</option>"
    
    If WcCheck = "26" Or WcCheck = "37" Or WcCheck = "44"Then 
            ReturnData = ReturnData & "<option value='" & Cstr(302) & "'" & ">" & Trim("Creases") & "</option>" & VbCrLf
            ReturnData = ReturnData & "<option value='" & Cstr(305) & "'" & ">" & Trim("Gassing") & "</option>" & VbCrLf
            ReturnData = ReturnData & "<option value='" & Cstr(311) & "'" & ">" & Trim("Other (Please Specify)") & "</option>" & VbCrLf
    Else
        For Count = 1 To Ubound(Reasons) 
    
            If Reasons(Count,2) = "0" Then
                ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & "'" & ">" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
            Else
                '## 34 = Zund
                If Reasons(Count,2) = "34" And WcCheck = "34" Then 
                    ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & " 'selected='selected' >" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
                    AutoReturnCode = Cstr(Reasons(Count,0))
                '## 10= Bullmer
                ElseIf Reasons(Count,2) = "10" And WcCheck = "10" Then
                    ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & " 'selected='selected' >" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
                    AutoReturnCode = Cstr(Reasons(Count,0)) 
                '## Laminating    
                ElseIf Reasons(Count,2) = "12" And WcCheck = "17" Then
                    ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & " 'selected='selected' >" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
                    AutoReturnCode = Cstr(Reasons(Count,0))               
                '## Welding               
                ElseIf Reasons(Count,2) = "2024" And WcCheck = "20" Or Reasons(Count,2) = "2024" And WcCheck = "24" Then
                    ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & " 'selected='selected' >" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
                    AutoReturnCode = Cstr(Reasons(Count,0))
                '## Sewing Meevo               
                ElseIf Reasons(Count,2) = "1647" And WcCheck = "16" Or Reasons(Count,2) = "1647" And WcCheck = "47" Then
                    ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & " 'selected='selected' >" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
                    AutoReturnCode = Cstr(Reasons(Count,0))
                ElseIf Reasons(Count,2) =  WcCheck Then
                    ReturnData = ReturnData & "<option value='" & Cstr(Reasons(Count,0)) & " 'selected='selected' >" & Trim(Reasons(Count,1)) & "</option>" & VbCrLf
                    AutoReturnCode = Cstr(Reasons(Count,0))
                End If
            End If
            
            'AutoGroupCode = Cstr(Reasons(Count,3))
        
        Next     
    End If
    
    Erase Reasons
    
    ReturnData = ReturnData & "#" & AutoReturnCode '& "~" & AutoGroupCode

End If


Response.Write ReturnData


%>
