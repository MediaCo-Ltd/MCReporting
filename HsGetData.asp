<% 
Dim errMsg
Dim LogSql
Dim LogRs
Dim LogExtra
Dim HsUserRs

Dim CustomRs
Dim SingleLogRs
Dim ReasonSql
Dim ReasonRs
Dim ReasonsExtra

Dim MyStartDate
Dim MyEndDate

Dim SourceExtra

errMsg = ""
LogExtra = ""
ReasonsExtra = "" 

'## Reasons Sql
ReasonSql = " Select Id, GroupId, Usage, Description From Reasons"

'## Logs Sql
LogSql = "SELECT  Logs.Id, Logs.UserId, Logs.GroupId, Logs.GroupSelection, Logs.Notes, Logs.CreatedDate, Logs.CreatedDateSerial,"
LogSql = LogSql & " Logs.Severity, Logs.SelectedDate, Logs.SelectedDateSerial, Logs.Resolved, Logs.LastModifiedDate,"
LogSql = LogSql & " Logs.LastModifiedDateSerial, Logs.LastModifiedBy, Logs.ResolvedDate, Logs.ResolvedDateSerial, Logs.HasImage,"
LogSql = LogSql & " Logs.ResolvedBy, Reasons.Description, Reasons.Usage, Logs.LocationId, Logs.CreatedByName, Logs.Dormant,"    

LogSql = LogSql & " Logs.Riddor, Logs.RiddorLevel, Logs.RiddorDays, Logs.RiddorDateSerial, Logs.RiddorSubmitted, Logs.RiddorSubmitedDate"

LogSql = LogSql & " From Logs INNER JOIN Reasons ON Logs.GroupSelection = Reasons.Id"

'###############################

Sub OpenReasons(GroupFilter)

    If GroupFilter = 4 Or GroupFilter = 5 Then
        '## Get all     
    Else
        ReasonsExtra = " Where (GroupId In(0, " & GroupFilter & "))"
    End If

    Set ReasonRs = Server.CreateObject("ADODB.Recordset")
	ReasonRs.ActiveConnection = Session("ConnHSReports")
	ReasonRs.Source = ReasonSql & ReasonsExtra 
	ReasonRs.CursorType = Application("adOpenForwardOnly")
	ReasonRs.CursorLocation = Application("adUseClient") 
	ReasonRs.LockType = Application("adLockReadOnly")
	ReasonRs.Open
	Set ReasonRs.ActiveConnection = Nothing

End Sub

Sub OpenLogs(GroupFilter, ReasonFilter, ResolvedFilter)   'and (Resolved = 0) 
    
    Dim LocalResolved
    
    If ResolvedFilter = 0 Then
        LocalResolved = " And (Logs.Resolved = 0) "
    Else
        LocalResolved = ""
    End If
    
    If GroupFilter = 4 Then
        '## Get all
        LogExtra = " Where(Logs.GroupSelection = " & ReasonFilter & ") AND (Dormant = 0) Order By Logs.GroupId, Logs.CreatedDateSerial" 
    ElseIf GroupFilter = 5 Then 
        LogExtra = " Where(Logs.Riddor = 1) And (Logs.GroupSelection = " & ReasonFilter & ") AND (Dormant = 0) Order By Logs.GroupId, Logs.CreatedDateSerial"  
    Else
        LogExtra = " Where(Logs.GroupId = " & GroupFilter & ") AND (Logs.GroupSelection = " & ReasonFilter & " )" & LocalResolved & " AND (Logs.Dormant = 0) Order By Logs.CreatedDateSerial"    
    End If
    
    
        
    Set LogRs = Server.CreateObject("ADODB.Recordset")
	LogRs.ActiveConnection = Session("ConnHSReports")
	LogRs.Source = LogSql & LogExtra 
	LogRs.CursorType = Application("adOpenForwardOnly")
	LogRs.CursorLocation = Application("adUseClient") 
	LogRs.LockType = Application("adLockReadOnly")
	LogRs.Open
	Set LogRs.ActiveConnection = Nothing

End Sub

'##############################################  Built Up Reports For Custom Export 


Sub OpenCustomRs(TypeFilter, ReasonFilter, UserFilter, ShiftFilter, LocationFilter, StartFilter, EndFilter, SourceFilter, ReturnPageType)

    SourceExtra = ""
    LogExtra = ""
    
    Dim OrderFilter
    
    If SourceFilter = 0 Then
        SourceExtra = "Logs.CreatedDateSerial"
    ElseIf SourceFilter = 1 Then
        SourceExtra = "Logs.LastModifiedDateSerial"
    Else
        SourceExtra = "Logs.CreatedDateSerial"   
    End If
    
    OrderFilter = " ORDER BY Logs.CreatedDateSerial"     
   
    If TypeFilter = "1" Then
        LogExtra = " Where (Logs.GroupId = 1) AND (Logs.Dormant = 0)"        
    ElseIf TypeFilter = "2" Then
        LogExtra = " Where (Logs.GroupId = 2) AND (Logs.Dormant = 0)" 
    ElseIf TypeFilter = "3" Then
        LogExtra = " Where (Logs.GroupId = 3) AND (Logs.Dormant = 0)" 
    ElseIf TypeFilter = "4" Then
        LogExtra = " Where (Logs.GroupId In(1,2,3)) AND (Logs.Dormant = 0)"
        OrderFilter = " ORDER BY Logs.GroupId, Logs.GroupSelection"
    End If
         
    If ReasonFilter = "0" Then
        '## No extra
    Else
        If LogExtra = "" Then
            LogExtra = " AND (Logs.GroupSelection = " & ReasonFilter & ")"
        Else
            LogExtra = LogExtra & " AND (Logs.GroupSelection = " & ReasonFilter & ")"
        End If
    End If
    
    If UserFilter = "0" Then
        '## No extra
    Else
        If LogExtra = "" Then
            LogExtra = " AND (Logs.UserId = " & UserFilter & ")"
        Else
            LogExtra = LogExtra & " AND (Logs.UserId = " & UserFilter & ")"
        End If
    End If
    
    If ShiftFilter = "0" Then
        '## No extra
    Else
        If LogExtra = "" Then
            LogExtra = " AND (Logs.SelectedTimeSlot = " & ShiftFilter & ")"
        Else
            LogExtra = LogExtra & " AND (Logs.SelectedTimeSlot = " & ShiftFilter & ")"
        End If
    End If
    
    If LocationFilter = "0" Then
        '## No extra
    Else
        If LogExtra = "" Then
            LogExtra = " AND (Logs.LocationId = " & ShiftFilter & ")"
        Else
            LogExtra = LogExtra & " AND (Logs.LocationId = " & LocationFilter & ")"
        End If
    End If
    
      
    If StartFilter = "0" Then
            '## No extra
    Else
        If StartFilter > 0 And EndFilter > 0 Then   '##WHERE        (Id > 11) AND (InChargeSerial BETWEEN 42537 AND 42604)
            If LogExtra = "" Then
                If StartFilter = EndFilter Then
                    LogExtra = " And (" & SourceExtra & " = " & StartFilter & ")"
                Else
                    LogExtra = " And (" & SourceExtra & " Between " & StartSerial & " And " & EndSerial & ")"
                End If
            Else
                If StartFilter = EndFilter Then 'LogExtra &
                    LogExtra = LogExtra & " And (" & SourceExtra & " = " & StartFilter & ")"
                Else
                    LogExtra = LogExtra & " And (" & SourceExtra & " Between " & StartSerial & " And " & EndSerial & ")"
                End If
            End If
        End If 
    End If


    Set CustomRs = Server.CreateObject("ADODB.Recordset")

    CustomRs.ActiveConnection = Session("ConnHSReports")      '## Now uses date serial
    
   ' If ReturnPageType = "Export" Then
   '     LogRs.Source = strMplExportSql
   ' Else
        CustomRs.Source = LogSql
   ' End If
    
    CustomRs.Source = CustomRs.Source & LogExtra & OrderFilter
     	
    CustomRs.CursorType = Application("adOpenForwardOnly")
    CustomRs.CursorLocation = Application("adUseClient")
    CustomRs.LockType = Application("adLockReadOnly")
    CustomRs.Open
    Set CustomRs.ActiveConnection = Nothing	
		        
End Sub

%>