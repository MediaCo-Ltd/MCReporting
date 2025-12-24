<%@Language="VBScript" Codepage="1252" EnableSessionState="True"%>
<%Option Explicit%> 

<%

If Session("UserName") = "" Then Response.Redirect ("SessionExpired.asp?uf=" & Trim(Request.Form("QuoteRef")))

'## Rem below on live
'## Session("JobNoError") = "Update ok" 
'## Response.Redirect "IhrJobNo.asp"

Dim Error
Dim NewId
Dim UpdateOk
Dim Count
Dim RowCount
Dim AddRs
Dim AddSql
Dim EditRs
Dim EditSql
Dim RedirectUrl
Dim DeptId
Dim CodeId

Dim OkCount
OkCount = 0

Dim RevisionCount

RowCount = Cint(Request.Form("TotalRows"))

'Session("JobNoError") = "RowCount = " & RowCount

'## Do quick Check of Rows

For Count = 0 To RowCount
    If Request.Form("Code" & Cstr(Count)) <> "" Then OkCount = OkCount +1
Next    
    
If OkCount = 0 Then
    Session("JobNoError") = "No Departments/Reasons selected. Nothing has been saved" 
    Response.Redirect ("IhrJobNo.asp")
End If  

Dim QuoteId
QuoteId = Trim(Request.Form("QuoteID"))  

'############################## Add New Record ##############################

AddSql = "SELECT * FROM OhQuoteData" 

Set AddRs = Server.CreateObject("ADODB.Recordset")
AddRs.ActiveConnection = Session("ConnQC")
AddRs.Source = AddSql
AddRs.CursorType = Application("adOpenStatic")
AddRs.CursorLocation = Application("adUseClient")
AddRs.LockType = Application("adLockOptimistic")
AddRs.Open

AddRs.AddNew

AddRs("Active") = 1

AddRs("QuoteRef") = Trim(Request.Form("QuoteRef"))
AddRs("QuoteId") = Trim(Request.Form("QuoteID"))
AddRs("Client") = Trim(Request.Form("Client"))
AddRs("Description") = Trim(Request.Form("Desc"))
AddRs("Sqm") = Trim(Request.Form("OrigSqm"))

AddRs("BookedInBy") = Trim(Request.Form("BookedInBy"))
AddRs("CreatedBy") = Session("UserName")
AddRs("CreatedById") = Session("UserId")

AddRs("Created") = Now()
AddRs("CreatedSerial") = CDbl(DateSerial(Year(Date),Month(Date),Day(Date)))

AddRs("Revision") = GetRevision

AddRs.Update
AddRs.MoveLast

NewId = AddRs("Id")

Set AddRs.ActiveConnection = Nothing
AddRs.Close
Set AddRs = Nothing


'############################## Add items data

AddSql = "SELECT * FROM OHQuoteItemData" 

Set AddRs = Server.CreateObject("ADODB.Recordset")
AddRs.ActiveConnection = Session("ConnQC")
AddRs.Source = AddSql
AddRs.CursorType = Application("adOpenStatic")
AddRs.CursorLocation = Application("adUseClient")
AddRs.LockType = Application("adLockOptimistic")
AddRs.Open

Dim NewSqm
Dim TotalSqm
Dim NewCost
Dim TotalCost
Dim NewMatCost
Dim TotalMatCost
Dim LocalItemDesc
NewCost = 0
NewSqm = 0
TotalSqm = 0
NewMatCost = 0
TotalMatCost = 0

Dim SendToArtwork
SendToArtwork = False

Dim SendToBookingIn
SendToBookingIn = False

Dim ItemIdList
ItemIdList = ""

Dim ProductionList
ProductionList = ""

For Count = 0 To RowCount
    If Request.Form("Code" & Cstr(Count)) <> "" Then
    
        '## Upadte Production, puts previous OP's on hold
        '## If source op (who discovered it ) isn't on hold all ops go on hold
        '## Prepress only put on hold if dept code = 1
        UpdateProduction Trim(Request.Form("ItemIdRow" & Cstr(Count))), Trim(Request.Form("Dept" & Cstr(Count))), GetOnHoldOp(Trim(Request.Form("ItemIdRow" & Cstr(Count))))
        
        AddRs.AddNew
        AddRs("OHQuoteDataId") = NewId
        AddRs("QuoteId") = Trim(Request.Form("QuoteID"))
        
        AddRs("ItemId") = Trim(Request.Form("ItemIdRow" & Cstr(Count)))
        AddRs("ItemAlpha") = Trim(Request.Form("ItemIdAlpha" & Cstr(Count)))
        AddRs("ReasonCode") = Trim(Request.Form("Code" & Cstr(Count)))
        
        AddRs("OriginalQty") = Trim(Request.Form("OrigQty" & Cstr(Count)))
        AddRs("NewQty") = Trim(Request.Form("Qty" & Cstr(Count)))
        
        '## Item Unit Cost
        NewCost = Request.Form("Qty" & Cstr(Count)) * Request.Form("ItemCost" & Cstr(Count))        
        AddRs("UnitCost") = CCur(NewCost)
        TotalCost = TotalCost + NewCost
        NewCost = 0
        
        AddRs("ItemOriginalSqm") = Trim(Request.Form("Sqm" & Cstr(Count)))
        
        '## Additional reason
        If Request.Form("Reason" & Cstr(Count)) <> "" Then 
            AddRs("Reason") = Trim(Request.Form("Reason" & Cstr(Count)))
        Else
            AddRs("Reason") = ""
        End If
               
        If Request.Form("ItemDesc"& Cstr(Count)) <> "" Then
            LocalItemDesc = Trim(Request.Form("ItemDesc" & Cstr(Count)))            
            If Right(LocalItemDesc,6) = "<br />" Then
                LocalItemDesc = Left(LocalItemDesc,Len(LocalItemDesc)-6)
            End If
            
            AddRs("ItemDescription") = LocalItemDesc
            LocalItemDesc = ""            
        Else
            AddRs("ItemDescription") = ""
        End If
        
        
        If Request.Form("OrigQty" & Cstr(Count)) = Request.Form("Qty" & Cstr(Count)) Then
            AddRs("ItemNewSqm") = Trim(Request.Form("Sqm" & Cstr(Count))) 
            TotalSqm = TotalSqm + Round(Request.Form("Sqm" & Cstr(Count)),2)      
        Else
            NewSqm = Request.Form("Sqm" & Cstr(Count)) / Request.Form("OrigQty" & Cstr(Count))
            NewSqm = NewSqm * Request.Form("Qty" & Cstr(Count))
            AddRs("ItemNewSqm") = Round(NewSqm,2)
            TotalSqm = TotalSqm + Round(NewSqm,2)
        End If
        
        '## Spreadsheet Items
        If Request.Form("SpId" & Cstr(Count)) <> "" Then
            AddRs("SpId") = Trim(Request.Form("SpId" & Cstr(Count)))
        Else
            AddRs("SpId") = ""
        End if
        
        AddRs("GroupId") = Trim(Request.Form("Dept" & Cstr(Count)))
        
        '## Sends email to artwork as well as normal group
        '## Sends if Booking in or Prepress
        If Trim(Request.Form("Dept" & Cstr(Count))) = "0" Then SendToArtwork = True
        If Trim(Request.Form("Dept" & Cstr(Count))) = "1" Then SendToArtwork = True
        If Trim(Request.Form("Dept" & Cstr(Count))) = "0" Then SendToBookingIn = True        
        
        If DeptId = "" Then
            DeptId = Trim(Request.Form("Dept" & Cstr(Count)))    
        Else
            DeptId = DeptId & "," & Trim(Request.Form("Dept" & Cstr(Count)))
        End If
        
        AddRs("ReasonCode") = Trim(Request.Form("Code" & Cstr(Count)))
        
        If CodeId = "" Then
            CodeId = Trim(Request.Form("Code" & Cstr(Count)))
        Else
            CodeId = CodeId & "," & Trim(Request.Form("Code" & Cstr(Count)))
        End If
        
        AddRs("Substrates") = Trim(Request.Form("ItemSubtrate" & Cstr(Count)))
        
        If ItemIdList = "" Then
            ItemIdList = Trim(Request.Form("ItemIdRow" & Cstr(Count)))
        Else
            ItemIdList = ItemIdList & "," & Trim(Request.Form("ItemIdRow" & Cstr(Count)))
        End If                
               
        If Request.Form("Prt"& Cstr(Count)) <> "" Then
            AddRs("Printer") = Trim(Request.Form("Prt" & Cstr(Count)))
        Else
            AddRs("Printer") = "No Printing"
        End If
        
        AddRs("StockGroup") = Request.Form("MatGroup" & Cstr(Count))
        
        If Request.Form("OrigQty" & Cstr(Count)) = Request.Form("Qty" & Cstr(Count)) Then
            AddRs("MaterialCost") = CCur(Request.Form("MatCost" & Cstr(Count)))
            TotalMatCost = TotalMatCost + CCur(Request.Form("MatCost" & Cstr(Count)))
        Else
            NewMatCost = CCur(Request.Form("MatCost" & Cstr(Count)))
            NewMatCost = NewMatCost / Request.Form("OrigQty" & Cstr(Count))
            NewMatCost = NewMatCost * Request.Form("Qty" & Cstr(Count))
            TotalMatCost = TotalMatCost + NewMatCost
            AddRs("MaterialCost") = NewMatCost
            NewMatCost = 0
        End If
        
        If Request.Form("MatCategory" & Cstr(Count)) = "Stock Items" Then
            AddRs("StockCategory") = "Stock"
        Else
            AddRs("StockCategory") = Request.Form("MatCategory" & Cstr(Count))
        End If
        
        If Request.Form("WCID" & Cstr(Count)) = "7#2#26" Or Request.Form("WCID" & Cstr(Count)) = "7#2#37" Or Request.Form("WCID" & Cstr(Count)) = "7#2#44" Then
            AddRs("WorkCentreName") = "Heat Press"
            'AddRs("ReasonText") = "Heat Press"
        Else
            AddRs("WorkCentreName") = ""
        End If
       
        AddRs("GroupName") = Trim(Request.Form("Grp" & Cstr(Count)))
        
        Select Case Trim(Request.Form("Code" & Cstr(Count)))
            Case 103
                AddRs("ReasonText") = "Other (Account Management)"
            Case 205
                AddRs("ReasonText") = "Other (Prepress)"
            Case 311
                AddRs("ReasonText") = "Other (Printing)"
            Case 409
                AddRs("ReasonText") = "Other (Finishing)"
            Case 505
                AddRs("ReasonText") = "Other (Pack/Disp)"
            Case 602
                AddRs("ReasonText") = "Other (Stock)"
            Case 401
                AddRs("ReasonText") = "Bullmer Cutting"
            Case 402
                AddRs("ReasonText") = "Zund Cutting"
            Case 403
                AddRs("ReasonText") = "Sewing"
            Case 404
                AddRs("ReasonText") = "Welding"
            Case 405
                AddRs("ReasonText") = "Eyeleting" 
            Case 406
                AddRs("ReasonText") = "Laminating"
            Case Else
                AddRs("ReasonText") = Trim(Request.Form("ReasonTxt" & Cstr(Count)))               
        End Select   
        
        
        AddRs.Update
        AddRs.MoveLast            
    End If       
Next

Set AddRs.ActiveConnection = Nothing
AddRs.Close
Set AddRs = Nothing

'###################################### Update Total Sqm & other data


EditSql = "SELECT ID, RedoSqm, DeptId, ReasonCodes, Substrates, RedoCost, RedoMatCost, Reasons, ActionText, ActionId, CboActionTxt"
EditSql = EditSql & " FROM OhQuoteData Where (ID = " & NewId & ")"

Set EditRs = Server.CreateObject("ADODB.Recordset")
EditRs.ActiveConnection = Session("ConnQC")
EditRs.Source = EditSql
EditRs.CursorType = Application("adOpenStatic")
EditRs.CursorLocation = Application("adUseClient")
EditRs.LockType = Application("adLockOptimistic")
EditRs.Open

EditRs("RedoSqm") = TotalSqm
EditRs("DeptId") = DeptId
EditRs("ReasonCodes") = CodeId
EditRs("Substrates") = GetTotalMaterial(ItemIdList) 
EditRs("RedoCost") = TotalCost
EditRs("RedoMatCost") = TotalMatCost
EditRs("Reasons") = GetTotalReasons(NewId)
EditRs("ActionText") = ""
EditRs("ActionId") = 0
EditRs("CboActionTxt") = ""

EditRs.Update

Set EditRs.ActiveConnection = Nothing
EditRs.Close
Set EditRs = Nothing

'## Rem below on live
'## Session("JobNoError") = "Update ok" 
'## Response.Redirect "IhrJobNo.asp"

'## No Email on Test or Home, only update Update Clarity on live
If Session("PC-Name") = "Home" Then
    Session("JobNoError") = "Update ok"
    Response.Redirect "IhrJobNo.asp"
Else
    If Session("Location") = "WorkTest" Then        
        Session("JobNoError") = "Update ok"
        Response.Redirect "IhrJobNo.asp"
    Else
        UpdateClarity
        Response.Redirect "IhrEmail.asp?Rid=" & NewId & "&AW=" & SendToArtwork & "&AM=" & SendToBookingIn
    End If
End If

'#################################

Sub UpdateClarity

    Dim UpdateClarityRs
    Dim UpdateClaritySql

    UpdateClaritySql = "UPDATE Quotation SET User10 = 'True', User16 = 1 WHERE (id = " & QuoteId & ")"

    Set UpdateClarityRs = Server.CreateObject("ADODB.Recordset")
    UpdateClarityRs.ActiveConnection = Session("ClarityConn")  
    UpdateClarityRs.Source = UpdateClaritySql
    UpdateClarityRs.CursorType = Application("adOpenStatic")
    UpdateClarityRs.CursorLocation = Application("adUseClient")
    UpdateClarityRs.LockType = Application("adLockOptimistic")

    UpdateClarityRs.Open

    Set UpdateClarityRs = Nothing

End Sub

'#################################

Function GetTotalReasons (RedoId)

    Dim TotalReasonsRs
    Dim TotalReasonsSql
    Dim LocalReason
    LocalReason = ""
    
    
    TotalReasonsSql = "SELECT DISTINCT OHQuoteDataId, ReasonCode, Reason, ReasonText"
    TotalReasonsSql = TotalReasonsSql & " FROM OHQuoteItemData"
    TotalReasonsSql = TotalReasonsSql & " WHERE (OHQuoteDataId = " & RedoId & ") ORDER BY ReasonCode"
    
    Set TotalReasonsRs = Server.CreateObject("ADODB.Recordset")
    TotalReasonsRs.ActiveConnection = Session("ConnQC")
    TotalReasonsRs.Source = TotalReasonsSql
    TotalReasonsRs.CursorType =  Application("adOpenForwardOnly")
    TotalReasonsRs.CursorLocation = Application("adUseClient")
    TotalReasonsRs.LockType = Application("adLockReadOnly")
    TotalReasonsRs.Open
    
    If TotalReasonsRs.BOF = True Or TotalReasonsRs.EOF = True Then
        LocalReason = "EOF or BOF"
    Else
        While Not TotalReasonsRs.EOF            
            If LocalReason = "" Then
                If Left(TotalReasonsRs("ReasonText"),5) = "Other" Then
                    LocalReason = TotalReasonsRs("ReasonText") & " " & TotalReasonsRs("Reason")
                Else
                    LocalReason = TotalReasonsRs("ReasonText")
                End If
            Else
                If Left(TotalReasonsRs("ReasonText"),5) = "Other" Then
                    LocalReason = LocalReason & ", " & TotalReasonsRs("ReasonText") & " " & TotalReasonsRs("Reason")
                Else
                    LocalReason = LocalReason & ", " & TotalReasonsRs("ReasonText")
                End If
            End If    
            TotalReasonsRs.MoveNext
        Wend
    End If

    TotalReasonsRs.Close
    Set TotalReasonsRs = Nothing
    
    GetTotalReasons = LocalReason

End Function

'#################################

Function GetTotalMaterial (IdList)

    Dim TotalMaterialRs
    Dim TotalMaterialSql  
    Dim LocalTotal  

    TotalMaterialSql = "SELECT DISTINCT GlobalPartList.Description, GlobalPartList.Level1"
    TotalMaterialSql = TotalMaterialSql & " FROM QuotationItemDetail INNER JOIN"
    TotalMaterialSql = TotalMaterialSql & " GlobalPartList ON QuotationItemDetail.TSLId = GlobalPartList.TSLId"
    TotalMaterialSql = TotalMaterialSql & " WHERE (QuotationItemDetail.QuotationItemId IN (" & IdList & "))"
    TotalMaterialSql = TotalMaterialSql & " AND (GlobalPartList.Level1 = N'Supplier Price List')"    
    TotalMaterialSql = TotalMaterialSql & " AND (GlobalPartList.Level3 IN ('Fabric', 'PVC', 'Rigid', 'SAV', 'Film', 'Pull-up', 'Paper'))"
    
    
    Set TotalMaterialRs = Server.CreateObject("ADODB.Recordset")
    TotalMaterialRs.ActiveConnection = Session("ClarityConn")
    TotalMaterialRs.Source = TotalMaterialSql
    TotalMaterialRs.CursorType =  Application("adOpenForwardOnly")
    TotalMaterialRs.CursorLocation = Application("adUseClient")
    TotalMaterialRs.LockType = Application("adLockReadOnly")
    TotalMaterialRs.Open
    
    
    If TotalMaterialRs.BOF = True Or TotalMaterialRs.EOF = True Then
        LocalTotal = "No Data"    
    Else
        While Not TotalMaterialRs.EOF
            If LocalTotal = "" Then
                LocalTotal = TotalMaterialRs("Description")
            Else
                LocalTotal = LocalTotal & "#" & TotalMaterialRs("Description")
            End If        
            TotalMaterialRs.MoveNext
        Wend
    End If
    
    TotalMaterialRs.Close
    Set TotalMaterialRs = Nothing 
    
    GetTotalMaterial = LocalTotal       

End Function

'#################################

Function GetRevision

    Dim RevisionSql
    Dim RevisionRS
    Dim LocalRevision
    LocalRevision = 1

    RevisionSql= "SELECT QuoteRef FROM OhQuoteData WHERE (QuoteRef = '" & Trim(Request.Form("QuoteRef")) & "')"

    Set RevisionRS = Server.CreateObject("ADODB.Recordset")
    RevisionRS.ActiveConnection = Session("ConnQC")
    RevisionRS.Source = RevisionSql
    RevisionRS.CursorType =  Application("adOpenForwardOnly")
    RevisionRS.CursorLocation = Application("adUseClient")
    RevisionRS.LockType = Application("adLockReadOnly")
    RevisionRS.Open
    
    
    If RevisionRS.BOF = True Or RevisionRS.EOF = True Then
        LocalRevision = 1
    Else
        LocalRevision = RevisionRS.RecordCount +1
    End If

    RevisionRS.Close
    Set RevisionRS = Nothing
    
    GetRevision = LocalRevision

End Function

'#################################

Sub UpdateProduction (QitemId, PP, OPNo)

    On Error Resume Next
    
    Dim UpdateProductionRs
    Dim UpdateProductionSql   

    UpdateProductionSql = "UPDATE ProdOperations SET StatusNumber = 2, ModifiedDate = GETDATE()"
    UpdateProductionSql = UpdateProductionSql & " FROM ProdOperations INNER JOIN"
    UpdateProductionSql = UpdateProductionSql & " ProdJobCards ON ProdOperations.ParentId = ProdJobCards.Id"
    UpdateProductionSql = UpdateProductionSql & " WHERE (ProdJobCards.QuoteItemId = " & QitemId & ")"
    
    If Cint(PP) = 0 Or Cint(PP) = 1 Then
        '## Update Prepess if booking in or prepress
    Else
        UpdateProductionSql = UpdateProductionSql & " AND (NOT (ProdOperations.WorkCentreId = 14))"
    End If
    
    If Cint(OPNo) <> 0 Then 
        UpdateProductionSql = UpdateProductionSql & " AND (ProdOperations.OPNumber < " & Cint(OPNo) & ")"
    End If
    
    Set UpdateProductionRs = Server.CreateObject("ADODB.Recordset")
    UpdateProductionRs.ActiveConnection = Session("ClarityConn")
    UpdateProductionRs.Source = UpdateProductionSql
    UpdateProductionRs.CursorType = Application("adOpenStatic")
    UpdateProductionRs.CursorLocation = Application("adUseClient")
    UpdateProductionRs.LockType = Application("adLockOptimistic")
    UpdateProductionRs.Open

    Set UpdateProductionRs = Nothing
    
    Err.Clear


End Sub

'#################################

Function GetOnHoldOp (QitemId)

    Dim OnHoldOpRs
    Dim OnHoldOpSql
    Dim LocalOpResult
    LocalOpResult = 0


    OnHoldOpSql = "SELECT ProdJobCards.QuoteItemId, ProdOperations.ParentId, ProdOperations.StatusNumber AS ProdOpStatus,"
    OnHoldOpSql = OnHoldOpSql & " ProdOperations.WorkCentreId, ProdOperations.OPNumber"
    OnHoldOpSql = OnHoldOpSql & " FROM ProdOperations INNER JOIN"
    OnHoldOpSql = OnHoldOpSql & " ProdJobCards ON ProdOperations.ParentId = ProdJobCards.Id"
    OnHoldOpSql = OnHoldOpSql & " WHERE (ProdJobCards.QuoteItemId = " & QitemId & ") AND (ProdOperations.StatusNumber = 2)"
    OnHoldOpSql = OnHoldOpSql & " ORDER BY ProdOperations.OPNumber DESC"

    Set OnHoldOpRs = Server.CreateObject("ADODB.Recordset")
    OnHoldOpRs.ActiveConnection = Session("ClarityConn")
    OnHoldOpRs.Source = OnHoldOpSql
    OnHoldOpRs.CursorType = Application("adOpenForwardOnly")
    OnHoldOpRs.CursorLocation = Application("adUseClient")
    OnHoldOpRs.LockType = Application("adLockReadOnly")
    OnHoldOpRs.Open
    
    If OnHoldOpRs.BOF = True Or OnHoldOpRs.EOF = True Then
        LocalOpResult = 0
    Else
        LocalOpResult = OnHoldOpRs("OPNumber")
    End If
    
    OnHoldOpRs.Close
    Set OnHoldOpRs = Nothing

    GetOnHoldOp = LocalOpResult

End Function

%>