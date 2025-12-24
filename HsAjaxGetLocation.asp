<%@Language="VBScript"  EnableSessionState="True"%>
<%Option Explicit%> 

 
<%

Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1

Dim Count
Dim Location()
Dim LocationRs
Dim GetLocationSql
Dim Nodata
Dim ReturnData
Dim ExcludedList
Dim WhereClause
Dim SentType

WhereClause = ""

Nodata = False

SentType = Cstr(Request.Form("xType"))
ExcludedList = Cstr(Request.Form("xExcluded"))


'## Get Locations based on Excluded list, used by Add/Edit Log
If SentType = "Log" Then
    If ExcludedList = "" Then
        '## Show all
        WhereClause = " WHERE (Active = 1)"
    Else
        '## Only show allowwd locations
        WhereClause = " WHERE (Active = 1) AND (NOT (Id IN (" & ExcludedList & ")))"
    End If
        
    GetLocationSql = "SELECT  Id, LocationName"
    GetLocationSql = GetLocationSql & " FROM Location"
    GetLocationSql = GetLocationSql & WhereClause    
    GetLocationSql = GetLocationSql & " ORDER BY LocationGroup, LocationTypeId"

    Set LocationRs = Server.CreateObject("ADODB.Recordset")
    LocationRs.ActiveConnection = Session("ConnHSReports")
    LocationRs.Source = GetLocationSql
    LocationRs.CursorType = Application("adOpenForwardOnly")
    LocationRs.CursorLocation = Application("adUseClient")
    LocationRs.LockType = Application("adLockReadOnly")
    LocationRs.Open
    Set LocationRs.ActiveConnection = Nothing

    If LocationRs.BOF = True Or LocationRs.EOF = True Then
        Nodata = True
    Else
        Redim Location(LocationRs.RecordCount,1)
        While Not LocationRs.EOF
            Location(LocationRs.AbsolutePosition,0) = LocationRs("Id")
            Location(LocationRs.AbsolutePosition,1) = LocationRs("LocationName")
            LocationRs.MoveNext
        Wend        
    End If

    LocationRs.Close
    Set LocationRs = Nothing

    If Nodata = False Then
        ReturnData = "<option value='' >Select&nbsp;Location</option>"
        For Count = 1 To Ubound(Location)
            ReturnData = ReturnData & ("<option value='" & Cstr(Location(Count,0)) & "'>" & Trim(Location(Count,1)) & "</option>)") & VbCrLf
        Next
        Erase Location
    End If
Else
    '## Get locations based on Logs for Custom select page
    GetLocationSql = "SELECT DISTINCT Location.LocationName, Logs.LocationId, Location.LocationGroup, Location.LocationTypeId"
    GetLocationSql = GetLocationSql & " FROM Logs INNER JOIN"
    GetLocationSql = GetLocationSql & " Location ON Logs.LocationId = Location.Id"
    If ExcludedList = "4" Then
        GetLocationSql = GetLocationSql & " WHERE (Location.Active = 1)"
    Else
        GetLocationSql = GetLocationSql & " WHERE (Location.Active = 1) AND (Logs.GroupId = " & ExcludedList & ")"
    End If
    GetLocationSql = GetLocationSql & " ORDER BY LocationGroup, LocationTypeId"
                         
    Set LocationRs = Server.CreateObject("ADODB.Recordset")
    LocationRs.ActiveConnection = Session("ConnHSReports")
    LocationRs.Source = GetLocationSql
    LocationRs.CursorType = Application("adOpenForwardOnly")
    LocationRs.CursorLocation = Application("adUseClient")
    LocationRs.LockType = Application("adLockReadOnly")
    LocationRs.Open
    Set LocationRs.ActiveConnection = Nothing
    
    If LocationRs.BOF = True Or LocationRs.EOF = True Then
        Nodata = True
    Else
        Redim Location(LocationRs.RecordCount,1)
        While Not LocationRs.EOF
            Location(LocationRs.AbsolutePosition,0) = LocationRs("LocationId")
            Location(LocationRs.AbsolutePosition,1) = LocationRs("LocationName")
            LocationRs.MoveNext
        Wend        
    End If

    LocationRs.Close
    Set LocationRs = Nothing

    If Nodata = False Then
        ReturnData = "<option value='' >Select&nbsp;Location</option>"
        For Count = 1 To Ubound(Location)
            ReturnData = ReturnData & ("<option value='" & Cstr(Location(Count,0)) & "'>" & Trim(Location(Count,1)) & "</option>)") & VbCrLf
        Next
        Erase Location
    End If
    
    
End If

Response.Write ReturnData

%>
