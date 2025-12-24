<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp")

Dim Error
Dim UserId
Dim UpdateOk
Dim Count
Dim Notes
Dim NewNotes
Dim SaveNotes

Dim AddRs
Dim AddSql
Dim UpdateRs
Dim UpdateSql
Dim RedirectUrl
Dim CreatedDateSerial
Dim EditRecordId
Dim NewRecordId
Dim SortArray
Dim SortString
Dim SortCount


RedirectUrl = "NcSelectOption.asp"

If Request.Form("chkUpdate") = "on" Then
    Session("OrderOk") = "chkUpdate Trap"
    '## Testing don't save data, will only be on if at home or on my pc at work
Else
   
'############################## Add New Log ##############################
    If Left(Request.Form("frmName"),8) = "frmNcAdd" Then  
    
    '## frmNcAdd

    '############################## Create log 

        AddSql = "SELECT * FROM Logs" 

        Set AddRs = Server.CreateObject("ADODB.Recordset")
        AddRs.ActiveConnection = Session("ConnNCReports")
        AddRs.Source = AddSql
        AddRs.CursorType = Application("adOpenStatic")
        AddRs.CursorLocation = Application("adUseClient")
        AddRs.LockType = Application("adLockOptimistic")
        AddRs.Open

        AddRs.AddNew
        
        '## Linked to Job
        If Request.Form("frmName") = "frmNcAddNewCL" Then
            AddRs("JobId") = Clng(Request.Form("hJobId")) 
            AddRs("JobTypeId") = Cint(Request.Form("hJobType")) 
            AddRs("QuoteId") = Clng(Request.Form("hQuoteId"))        
            AddRs("QuoteRef") = Trim(Request.Form("hJobRef"))
            
            '## All Ids  to be sorted, then rebuild string to save
            '## Item Alpha to be saved from sorted Item Id so in correct order
            SortArray = Split(Request.Form("hItemId"),",",-1,1)    
            SortIds SortArray
    
            For SortCount = 0 To Ubound(SortArray)
                SortString = SortString & SortArray(SortCount) & ","    
            Next
    
            If Right(SortString,1) = "," Then SortString = Left(SortString,Len(SortString)-1)
            
            AddRs("QuoteItemID") = SortString
            AddRs("QuoteItemsAlpha") = RebuildAlpha (SortArray)
            Erase SortArray
            SortString = ""        
            
            '## Sort Department
            SortArray = Split(Request.Form("hDeptSelected"),",",-1,1)    
            SortIds SortArray
    
            For SortCount = 0 To Ubound(SortArray)
                SortString = SortString & SortArray(SortCount) & ","    
            Next
    
            If Right(SortString,1) = "," Then SortString = Left(SortString,Len(SortString)-1)
            AddRs("DeptSelected") = SortString            
            SortString = ""
            
            '## Save Dept Id from Selected & strip out leading zero's
            For SortCount = 0 To Ubound(SortArray)
                SortString = SortString & Cstr(Cint(SortArray(SortCount))) & ","    
            Next
            If Right(SortString,1) = "," Then SortString = Left(SortString,Len(SortString)-1)
            
            AddRs("DeptIds") = SortString
            Erase SortArray
            SortString = ""          
                 
        Else
            AddRs("JobId") = 0 
            AddRs("JobTypeId") = 0 
            AddRs("QuoteId") = 0        
            AddRs("QuoteRef") = ""
            AddRs("QuoteItemID") = ""
            AddRs("QuoteItemsAlpha") = "" 
            AddRs("DeptIds") = ""
            AddRs("DeptSelected") = ""        
        End If       

        '## Common
        
        '## Sort Groups
        SortArray = Split(Request.Form("hGroupSelected"),",",-1,1)    
        SortIds SortArray

        For SortCount = 0 To Ubound(SortArray)
            SortString = SortString & SortArray(SortCount) & ","    
        Next

        If Right(SortString,1) = "," Then SortString = Left(SortString,Len(SortString)-1)
        AddRs("GroupSelection") = SortString            
        SortString = ""
        
        '## Save Group Id from Selected & strip out leading zero's
        For SortCount = 0 To Ubound(SortArray)
            SortString = SortString & Cstr(Cint(SortArray(SortCount))) & ","    
        Next
        If Right(SortString,1) = "," Then SortString = Left(SortString,Len(SortString)-1)
        
        AddRs("GroupIds") = SortString
        
        Erase SortArray
        SortString = "" 
        
        AddRs("ReasonIds") = Trim(Request.Form("hReason"))
        AddRs("ReasonSelection") = Trim(Request.Form("hReasonSelected"))    
        
       
        If Request.Form("txtDetails") <> "" Then        
            '## need to do formatting so vbcrlf = <br />            
            Notes = Trim(Cstr(Request.Form("txtDetails")))
            Notes = Replace(Notes,VbCrLf,"<br />",1,-1,1)
            '## If triple line feeds cut one out
            '## They seem to like having lots of gaps, making the notes to long
            '## Triple = 1 for return of 1st line then 2 gaps before next line
            Notes = Replace(Notes,"<br /><br /><br />","<br /><br />",1,-1,1)
            Notes = Replace(Notes,"&","&amp;",1,-1,1)
            AddRs("Notes") = Trim(Notes)
        End If
        
        AddRs("CreatedDate") = Date
        AddRs("CreatedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        
        
        If Request.Form("txtDate") <> "" Then
          '  If Request.Form("txtDate") > Now Then
          '      LocalTimeStr = Time()
          '      AddRs("SelectedDate") = Now
          '      AddRs("SelectedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
          '  Else 
                'LocalTimeStr = Right(Request.Form("txtDate"),5) & ":00"
                AddRs("SelectedDate") = Request.Form("txtDate")
                AddRs("SelectedDateSerial") = CDbl(DateSerial(Year(Request.Form("txtDate")),Month(Request.Form("txtDate")),Day(Request.Form("txtDate"))))
          '  End If
        Else    
            'LocalTimeStr = Time()
            AddRs("SelectedDate") = Now
            AddRs("SelectedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        End If
        
               
        AddRs("UserId") = Cint(Session("UserId"))
        AddRs("CreatedByName") = Trim(Session("UserName"))
        
        If Request.Form("chkRecurring") = "on" Then  
            'AddRs("RecurringFault") = 1
        Else
            'AddRs("RecurringFault") = 0
        End If
        
        If Request.Form("chkFixed") = "on" Then  
            'AddRs("FaultRepaired") = 1
        Else
            'AddRs("FaultRepaired") = 0
        End If
        
        AddRs.Update
        AddRs.MoveLast

        NewRecordId = AddRs("Id")
        
        '### Update reasons has logs ?
        
        Set AddRs.ActiveConnection = Nothing
        AddRs.Close
        Set AddRs = Nothing
               
        If Err = 0 Then
            UpdateOk = True
            Session("OrderOk") = "New record " & NewRecordId & " added"
           ' If Cint(Request.Form("hSeverity")) = 3 Then  '??????????????
            '    RedirectUrl = "NcEmail.asp"
           '     Session("EmailId") = NewRecordId                
           ' End If
        Else
            Session("SystemError") = "Add NC Log page"
            Response.Redirect ("SystemError.asp")       
        End If
    
'############################## Edit Log Update ############################## 
    
    ElseIf Left(Request.Form("frmName"),12) = "frmNcEditLog" Then
    
            
        EditRecordId = Request.Form("hLogId")
        UpdateSql = "SELECT * FROM Logs Where (Id = " & EditRecordId & ")" 

        Set UpdateRs = Server.CreateObject("ADODB.Recordset")
        UpdateRs.ActiveConnection = Session("ConnNCReports")
        UpdateRs.Source = UpdateSql
        UpdateRs.CursorType = Application("adOpenStatic")
        UpdateRs.CursorLocation = Application("adUseClient")
        UpdateRs.LockType = Application("adLockOptimistic")
        UpdateRs.Open

        UpdateRs("LastModifiedBy") = Cint(Session("UserId"))       
        UpdateRs("LastModifiedDate") = Date
        UpdateRs("LastModifiedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))
        
        If Cbool(Request.Form("hHasNotes")) = False Then        
            If Request.Form("txtResponse") <> "" Then        
                '## need to do formatting so vbcrlf = <br />            
                Notes = Trim(Cstr(Request.Form("txtResponse")))
                Notes = Replace(Notes,VbCrLf,"<br />",1,-1,1)
                '## If triple line feeds cut one out
                '## They seem to like having lots of gaps, making the notes to long
                '## Triple = 1 for return of 1st line then 2 gaps before next line
                Notes = Replace(Notes,"<br /><br /><br />","<br /><br />",1,-1,1)
                Notes = Replace(Notes,"&","&amp;",1,-1,1)
                Notes = Notes & "<br />" & Session("UserName") & " " & Date
                SaveNotes = Trim(Notes) & "<br />"
            End If
        Else
            If Request.Form("txtResponseNew") <> "" Then        
                '## need to do formatting so vbcrlf = <br />            
                NewNotes = Trim(Cstr(Request.Form("txtResponseNew")))
                NewNotes = Replace(NewNotes,VbCrLf,"<br />",1,-1,1)
                '## If triple line feeds cut one out
                '## They seem to like having lots of gaps, making the notes to long
                '## Triple = 1 for return of 1st line then 2 gaps before next line
                NewNotes = Replace(NewNotes,"<br /><br /><br />","<br /><br />",1,-1,1)
                NewNotes = Replace(NewNotes,"&","&amp;",1,-1,1)
                NewNotes = NewNotes & "<br />" & Session("UserName") & " " & Date 
                SaveNotes = "<br />" & Trim(NewNotes) & "<br />"          
            End If
        End If
        
        If Cbool(Request.Form("hHasNotes")) = False Then 
            If SaveNotes <> "" Then        
                UpdateRs("ResponseNotes") = SaveNotes
            End If
        Else
            If SaveNotes <> "" Then 
                UpdateRs("ResponseNotes") = UpdateRs("ResponseNotes") & SaveNotes
            End If
        End If
        
        If UpdateRs("ResolvedDateSerial") = 0 And Request.Form("chkResolved") = "on" Then
            UpdateRs("ResolvedBy") = Cint(Session("UserId"))
            UpdateRs("ResolvedDate") = Date
            UpdateRs("ResolvedDateSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))            
        End If
                                       
        If Request.Form("chkResolved") = "on" Then  
            UpdateRs("Resolved") = 1
        Else
            UpdateRs("Resolved") = 0
        End If     
        
        UpdateRs.Update
        UpdateRs.MoveLast 
        
        Set UpdateRs.ActiveConnection = Nothing
        UpdateRs.Close
        Set UpdateRs = Nothing   
        
        If Err = 0 Then
            UpdateOk = True
        Else
            Session("SystemError") = "Edit NC Edit Log page"
            Response.Redirect ("SystemError.asp")       
        End If
        
         '################################### Clear record lock         
         '## Done when return to option Page Or Login
        
    Else
        Response.Redirect (RedirectUrl)
    End If
End If

Response.Redirect (RedirectUrl)

'#########################################################################

Sub UpdateMachine ()

    'Dim UpdateMachineRs
       
    'Set UpdateMachineRs = Server.CreateObject("ADODB.Recordset")
    'UpdateMachineRs.ActiveConnection = Session("ConnNCReports") 
    'UpdateMachineRs.Source = "Select Id, HasLogs From Machine Where (Id = " & UpdateMachineId & ")"
    
    'UpdateMachineRs.CursorType = Application("adOpenStatic")
    'UpdateMachineRs.CursorLocation = Application("adUseClient")
    'UpdateMachineRs.LockType = Application("adLockOptimistic")
    'UpdateMachineRs.Open

    'UpdateMachineRs("HasLogs") = 1
    'UpdateMachineRs.Update
    
    'UpdateMachineRs.Close
    'Set UpdateMachineRs = Nothing    

End Sub

'################################################## Sort Out Prices, Bubble Sort

Sub SortIds( byRef arrArray )

    Dim row, j
    Dim StartingKeyValue, NewKeyValue, swap_pos

    For row = 0 To UBound( arrArray ) - 1
    'Take a snapshot of the first element
    'in the array because if there is a 
    'smaller value elsewhere in the array 
    'we'll need to do a swap.
        StartingKeyValue = arrArray ( row )
        NewKeyValue = arrArray ( row )
        swap_pos = row
	    	
        For j = row + 1 to UBound( arrArray )
        'Start inner loop.
            If arrArray ( j ) < NewKeyValue Then
            'This is now the lowest number - 
            'remember it's position.
                swap_pos = j
                NewKeyValue = arrArray ( j )
            End If
        Next
	    
        If swap_pos <> row Then
        'If we get here then we are about to do a swap
        'within the array.		
            arrArray ( swap_pos ) = StartingKeyValue
            arrArray ( row ) = NewKeyValue
        End If	
    Next
    
End Sub

'################################### Rebuild Alpha fom sorted Array 

Function RebuildAlpha (SentArray)

Dim Count
Dim ItemAlpha
Dim Localstring
For Count = 0 to Ubound(SentArray)   
    
    If SentArray(Count) <> "" Then
                
        '## Create the Item Letter for normal quote lines
        If SentArray(Count) <= 26 Then
             ItemAlpha = Cstr(Chr(SentArray(Count)+64))   
        ElseIf SentArray(Count) >= 27 And SentArray(Count) <= 52 Then
            ItemAlpha = Cstr(Chr(SentArray(Count)+38))
            ItemAlpha = "A" & ItemAlpha
        ElseIf SentArray(Count) >= 53 And SentArray(Count) <= 78 Then
            ItemAlpha = Cstr(Chr(SentArray(Count)+12))
            ItemAlpha = "B" & ItemAlpha 
        ElseIf SentArray(Count) >= 79 And SentArray(Count) <= 104 Then
            ItemAlpha = Cstr(Chr(SentArray(Count)-14))
            ItemAlpha = "C" & ItemAlpha 
        ElseIf SentArray(Count) >= 105 And SentArray(Count) <= 130 Then
            ItemAlpha = Cstr(Chr(SentArray(Count)-40)) 
            ItemAlpha = "D" & ItemAlpha
        ElseIf SentArray(Count) >= 131 And SentArray(Count) <= 156 Then
            ItemAlpha = Cstr(Chr(SentArray(Count)-66)) 
            ItemAlpha = "D" & ItemAlpha
        Else
            ItemAlpha = "?"  
        End If                         
        
        Localstring = Localstring & ItemAlpha & ","
    
     End If
     
Next 
 
If Right(Localstring,1) = "," Then Localstring = Left(Localstring,Len(Localstring)-1)   

RebuildAlpha  = Localstring

End Function
%>
