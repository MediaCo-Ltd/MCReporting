<% 
Dim errMsg
Dim LogSql
Dim LogRs
Dim LogExtra
Dim NcUserRs

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

'## Logs Sql
LogSql = "SELECT * FROM Logs" 



'## Reasons Sql
ReasonSql = " Select Id, GroupId, Usage, Description From Reasons"



'############################## User Name

Function GetNcUser(UserId)

  '  Dim LocalName
  '  LocalName = ""

  '  Set NcUserRs = Server.CreateObject("ADODB.Recordset")
'	NcUserRs.ActiveConnection = Session("ConnMcLogon")
	'NcUserRs.Source = "Select UserName From Users Where (Id = " & UserId & ")" 
	'NcUserRs.CursorType = Application("adOpenForwardOnly")
'	NcUserRs.CursorLocation = Application("adUseClient") 
'	NcUserRs.LockType = Application("adLockReadOnly")
'	NcUserRs.Open
'	Set NcUserRs.ActiveConnection = Nothing

  '  If NcUserRs.BOF = True Or NcUserRs.EOF = True Then
  '      LocalName = ""
  '  Else
 '       LocalName = NcUserRs("UserName")
  '  End If
    
  '  NcUserRs.Close
  '  Set NcUserRs = Nothing

 '   GetHsUser = LocalName

End Function

'###############################

Sub OpenReasons(GroupFilter)

  '  If GroupFilter = 4 Then
        '## Get all     
  '  Else
  '      ReasonsExtra = " Where (Id In(" & GroupFilter & "))"
  '  End If

  '  Set ReasonRs = Server.CreateObject("ADODB.Recordset")
'	ReasonRs.ActiveConnection = Session("ConnNCReports")
	'ReasonRs.Source = ReasonSql & ReasonsExtra 
'	ReasonRs.CursorType = Application("adOpenForwardOnly")
	'ReasonRs.CursorLocation = Application("adUseClient") 
	'ReasonRs.LockType = Application("adLockReadOnly")
'	ReasonRs.Open
'	Set ReasonRs.ActiveConnection = Nothing

End Sub



Sub OpenLogs(Filter)      
    
    If Filter = "" Then
        '## Get all
        LogExtra = ""    
    Else
        LogExtra = " Where(Logs.Id = " & Filter & ")"    
    End If
        
    Set LogRS = Server.CreateObject("ADODB.Recordset")
    LogRS.ActiveConnection = Session("ConnNCReports")
    LogRS.Source = LogSql & LogExtra
    LogRS.CursorType = Application("adOpenForwardOnly")
    LogRS.CursorLocation = Application("adUseClient")
    LogRS.LockType = Application("adLockReadOnly")
    LogRS.Open

End Sub


Sub OpenLogByUser()      
        
    Set LogRS = Server.CreateObject("ADODB.Recordset")
    LogRS.ActiveConnection = Session("ConnNCReports")
    LogRS.Source = LogSql & " Where (UserId = " & Clng(Session("UserId")) & ")"
    LogRS.CursorType = Application("adOpenForwardOnly")
    LogRS.CursorLocation = Application("adUseClient")
    LogRS.LockType = Application("adLockReadOnly")
    LogRS.Open

End Sub

'##############################################  Built Up Reports For Custom Export 


Sub OpenCustomRs(TypeFilter, ReasonFilter, UserFilter, ShiftFilter, LocationFilter, StartFilter, EndFilter, SourceFilter, ReturnPageType)

    SourceExtra = ""
    LogExtra = ""
    
   ' Dim OrderFilter
    
   ' If SourceFilter = 0 Then
   '     SourceExtra = "Logs.CreatedDateSerial"
   ' ElseIf SourceFilter = 1 Then
   '     SourceExtra = "Logs.LastModifiedDateSerial"
   ' Else
   '     SourceExtra = "Logs.CreatedDateSerial"   
   ' End If
    
   ' OrderFilter = " ORDER BY Logs.CreatedDateSerial"     
   
   ' If TypeFilter = "1" Then
   '     LogExtra = " Where (Logs.GroupId = 1)"        
  '  ElseIf TypeFilter = "2" Then
   '     LogExtra = " Where (Logs.GroupId = 2)" 
   ' ElseIf TypeFilter = "3" Then
   '     LogExtra = " Where (Logs.GroupId = 3)" 
   ' ElseIf TypeFilter = "4" Then
   '     LogExtra = " Where (Logs.GroupId In(1,2,3))"
   '     OrderFilter = " ORDER BY Logs.GroupId, Logs.GroupSelection"
   ' End If
         
   ' If ReasonFilter = "0" Then
        '## No extra
  '  Else
  '      If LogExtra = "" Then
  '          LogExtra = " AND (Logs.GroupSelection = " & ReasonFilter & ")"
  '      Else
  '          LogExtra = LogExtra & " AND (Logs.GroupSelection = " & ReasonFilter & ")"
  '      End If
  '  End If
    
 '   If UserFilter = "0" Then
        '## No extra
 '   Else
 '       If LogExtra = "" Then
 '           LogExtra = " AND (Logs.UserId = " & UserFilter & ")"
 '       Else
 '           LogExtra = LogExtra & " AND (Logs.UserId = " & UserFilter & ")"
 '       End If
 '   End If
    
 '   If ShiftFilter = "0" Then
        '## No extra
 '   Else
 '       If LogExtra = "" Then
 '           LogExtra = " AND (Logs.SelectedTimeSlot = " & ShiftFilter & ")"
 '       Else
 '           LogExtra = LogExtra & " AND (Logs.SelectedTimeSlot = " & ShiftFilter & ")"
 '       End If
 '   End If
    
 '   If LocationFilter = "0" Then
        '## No extra
 '   Else
 '       If LogExtra = "" Then
 '           LogExtra = " AND (Logs.LocationId = " & ShiftFilter & ")"
 '       Else
 '           LogExtra = LogExtra & " AND (Logs.LocationId = " & LocationFilter & ")"
 '       End If
 '   End If
    
      
  '  If StartFilter = "0" Then
            '## No extra
 '   Else
 '       If StartFilter > 0 And EndFilter > 0 Then   '##WHERE        (Id > 11) AND (InChargeSerial BETWEEN 42537 AND 42604)
 '           If LogExtra = "" Then
 '               If StartFilter = EndFilter Then
 '                   LogExtra = " And (" & SourceExtra & " = " & StartFilter & ")"
  '              Else
  '                  LogExtra = " And (" & SourceExtra & " Between " & StartSerial & " And " & EndSerial & ")"
 '               End If
 '           Else
 '               If StartFilter = EndFilter Then 'LogExtra &
 '                   LogExtra = LogExtra & " And (" & SourceExtra & " = " & StartFilter & ")"
 '               Else
 '                   LogExtra = LogExtra & " And (" & SourceExtra & " Between " & StartSerial & " And " & EndSerial & ")"
 '               End If
 '           End If
 '       End If 
 '   End If


  '  Set CustomRs = Server.CreateObject("ADODB.Recordset")

 '   CustomRs.ActiveConnection = Session("ConnNCReports")      '## Now uses date serial
    
   
'        CustomRs.Source = LogSql
 
    
 '   CustomRs.Source = CustomRs.Source & LogExtra & OrderFilter
     	
 '   CustomRs.CursorType = Application("adOpenForwardOnly")
 '   CustomRs.CursorLocation = Application("adUseClient")
 '   CustomRs.LockType = Application("adLockReadOnly")
 '   CustomRs.Open
 '   Set CustomRs.ActiveConnection = Nothing	
		        
End Sub

Private Function DeptDescription (DeptId) 

'## Work Centres from Clarity

    Dim DeptSql
    Dim DeptRs
    Dim DeptDesc
    
    
    DeptSql = "SELECT Id, Description From ProdWorkCentres Where (Id IN(" & DeptId & ")) Order By WorkCentreGroupId"

    Set DeptRs = Server.CreateObject("ADODB.Recordset")
    DeptRs.ActiveConnection = strConnClarity
    DeptRs.Source = DeptSql
    DeptRs.CursorType = Application("adOpenForwardOnly")
    DeptRs.CursorLocation = Application("adUseClient")
    DeptRs.LockType = Application("adLockReadOnly")
    DeptRs.Open

    If DeptRs.BOF= True Or DeptRs.EOF = True Then
        DeptDesc = ""
    Else
        While Not DeptRs.EOF    
            If DeptDesc = "" Then
                DeptDesc = DeptRs("Description")
            Else
                DeptDesc = DeptDesc & ", " & DeptRs("Description")
            End If
            DeptRs.MoveNext
        Wend
    End If
    
    'If DeptId = "0" Then DeptDesc = "Other"
    
    DeptRs.Close
    Set DeptRs = Nothing

    DeptDescription = DeptDesc

End Function

Private Function GroupDescription (GrpId)

    Dim GrpSql
    Dim GrpRs
    Dim GrpDesc
    GrpSql = "SELECT GroupId, GroupName, SortOrder FROM GroupData Where (GroupId IN(" & GrpId & ")) AND (Active = 1) Order By SortOrder"

    Set GrpRs = Server.CreateObject("ADODB.Recordset")
    GrpRs.ActiveConnection = Session("ConnNCReports")
    GrpRs.Source = GrpSql
    GrpRs.CursorType = Application("adOpenForwardOnly")
    GrpRs.CursorLocation = Application("adUseClient")
    GrpRs.LockType = Application("adLockReadOnly")
    GrpRs.Open

    If GrpRs.BOF= True Or GrpRs.EOF = True Then
        GrpDesc = ""
    Else
        While Not GrpRs.EOF    
            If GrpDesc = "" Then
                GrpDesc = GrpRs("GroupName")
            Else
                GrpDesc = GrpDesc & ", " & GrpRs("GroupName")
            End If
            GrpRs.MoveNext
        Wend
    End If

    GrpRs.Close
    Set GrpRs = Nothing

    GroupDescription = GrpDesc

End Function

Private Function ReasonDescription (RsId)

    Dim RsdSql
    Dim RsdRs
    Dim RsdDesc
    
    RsdSql = "SELECT Description, SortOrder FROM Reasons Where (Id IN(" & RsId & ")) AND (Active = 1) Order By SortOrder"

    Set RsdRs = Server.CreateObject("ADODB.Recordset")
    RsdRs.ActiveConnection = Session("ConnNCReports")
    RsdRs.Source = RsdSql
    RsdRs.CursorType = Application("adOpenForwardOnly")
    RsdRs.CursorLocation = Application("adUseClient")
    RsdRs.LockType = Application("adLockReadOnly")
    RsdRs.Open

    If RsdRs.BOF= True Or RsdRs.EOF = True Then
        RsdDesc = ""
    Else
        While Not RsdRs.EOF    
            If RsdDesc = "" Then
                RsdDesc = RsdRs("Description")
            Else
                RsdDesc = RsdDesc & ", " & RsdRs("Description")
            End If
            RsdRs.MoveNext
        Wend
    End If

    RsdRs.Close
    Set RsdRs = Nothing

    ReasonDescription = RsdDesc

End Function


%>