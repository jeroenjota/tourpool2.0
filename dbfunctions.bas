Attribute VB_Name = "dbfunctions"
Option Explicit

Function getOrganisation(cn As ADODB.Connection, Optional field As String) As String
'get the name for the organisation of this pool / or just the content of field
Dim adoCmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sqlstr As String
Dim result As String
    sqlstr = "Select * from tblOrganisation"
    Set adoCmd = New ADODB.Command
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        Set rs = .Execute
    End With

    If Not rs.EOF Then
        If field = "" Then
            result = Trim(rs!firstname)
            If rs!middlename > "" Then
                result = result & " " & Trim(rs!middlename)
            End If
            If rs!lastname > "" Then
                result = result & " " & Trim(rs!lastname)
            End If
'            result = result & vbNewLine & Trim(rs!address) & vbNewLine & Trim(rs!postalcode) & " " & Trim(rs!city)
        Else
            result = rs(field)
        End If
    End If
    getOrganisation = result
    rs.Close
    Set rs = Nothing
End Function

Function getPoolInfo(fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tblPools
Dim adoCmd As ADODB.Command
Dim rs As ADODB.Recordset
Dim sqlstr As String

    Set adoCmd = New ADODB.Command
    sqlstr = "Select " & fldName & " from tblPools where poolid = ?"
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").Value = thisPool
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        getPoolInfo = rs(fldName)
    Else
        getPoolInfo = Null
    End If
    
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing

End Function

Function getTourInfo(fldName As String, cn As ADODB.Connection)
'return the value of fieldnmame in tbltours
    Dim adoCmd As ADODB.Command
    Set adoCmd = New ADODB.Command
    Dim sqlstr As String
    Dim result As Variant
    Dim rs As ADODB.Recordset
    
    sqlstr = "Select * from tblTours Where tourID = ? "
    With adoCmd
        .ActiveConnection = cn
        .CommandType = adCmdText
        .CommandText = sqlstr
        .Prepared = True
        .Parameters.Append .CreateParameter("id", adInteger, adParamInput)
        .Parameters("id").Value = thisTour
        Set rs = .Execute
    End With
    If Not rs.EOF Then
        ' add description as extra - Access doesn't understand concat
        If fldName = "description" Then
            result = rs!tourName & " - " & rs!tourYear
        Else
            If rs(fldName).Type = adBoolean Then
                result = CBool(rs(fldName)) * 1
            Else
                result = rs(fldName)
            End If
        End If
    Else
        result = Null
    End If
    getTourInfo = result
    rs.Close
    Set rs = Nothing
    Set adoCmd = Nothing
End Function

Function chkPoolHasCompetitors(pool As Long, cn As ADODB.Connection)
'are there competitors for this pool
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
        
        sqlstr = "Select  poolID from tblCompetitors Where poolid = " & pool
        rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        chkPoolHasCompetitors = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function chkTourHasPools(tour As Long, cn As ADODB.Connection)
'are there pools for this tour?
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
        sqlstr = "Select tourID from tblPools Where tourid = " & tour
        rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
        chkTourHasPools = Not rs.EOF
    rs.Close
    Set rs = Nothing
End Function

Function getThisPooltourId(cn As ADODB.Connection) As Long
'return the tour for the current pool
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    getThisPooltourId = 0
    Dim sqlstr As String
    sqlstr = "Select tourID from tblPools Where poolid = " & thisPool
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getThisPooltourId = rs!tourId
    End If
    rs.Close
    Set rs = Nothing
End Function

Function chkTourStarted(cn As ADODB.Connection)
'check to see if tour already started

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sqlstr As String
    chkTourStarted = False
    sqlstr = "Select * from tblTours Where tourid = " & getThisPooltourId(cn)
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        chkTourStarted = CDbl(rs!TourStartDate) < CDate(Now())
    End If
    rs.Close
    Set rs = Nothing
End Function

Function supportsTransactions(cn As ADODB.Connection) As Boolean
'check if connection supports transactions
    On Error GoTo err_supportsTransactions:
        Dim lValue As Long
        lValue = cn.Properties("Transaction DDL").Value
        supportsTransactions = True
    Exit Function
err_supportsTransactions:
    Select Case Err.number
    Case adErrItemNotFound:
        supportsTransactions = False
    Case Else
        MsgBox Err.description
    End Select
End Function

Function getDeelnemerInfo(id As Long, fld As String, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblDeelnemers where DeelnemId = " & id
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getDeelnemerInfo = rs(fld)
    Else
        getDeelnemerInfo = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function getTeamId(tourTeamCode As Long, cn As ADODB.Connection)
'get the basic id  of a tour teamcode
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    sqlstr = "Select * from tbltourTeamCodes where teamCodeId = " & tourTeamCode
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    If Not rs.EOF Then
        getTeamId = rs(rs!teamId)
    Else
        getTeamId = Null
    End If
    rs.Close
    Set rs = Nothing
End Function

Function rennerInTourTeam(playerId As Long, teamId As Long, cn As ADODB.Connection)
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblTeamRenners where teamId = " & teamId
    sqlstr = sqlstr & " AND RennerId = " & rennerID
    sqlstr = sqlstr & " AND tourId = " & thisTour
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    rennerInTourTeam = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function RennerExists(fName As String, mName As String, lName As String, NickName As String, cn As ADODB.Connection)
    'check double entries
    Dim sqlstr As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    sqlstr = "Select * from tblPeople where (firstname = '" & fName
    sqlstr = sqlstr & "' AND middleName = '" & mName
    sqlstr = sqlstr & "' AND lastName = '" & lName
    sqlstr = sqlstr & "') OR nickName = '" & NickName & "'"
    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
    
    RennerExists = Not rs.EOF
    
    rs.Close
    Set rs = Nothing
End Function

Function getStageCount(cn As ADODB.Connection, Optional tourId As Long)
  'return number of Stages for current tour or given tourID
  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  If Not tourId Then tourId = thisTour
  Set rs = New ADODB.Recordset
  sqlstr = "Select COUNT(*) as recAant from tbltourSchedule "
  sqlstr = sqlstr & "WHERE tourID = " & tourId
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getStageCount = rs!recAant
  Else
    getStageCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getCount(strSQL As String, cn As ADODB.Connection)
  'return number of records in fromTbl
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  rs.Open strSQL, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    rs.MoveLast
    getCount = rs.RecordCount
  Else
    getCount = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getPoolPoints(description As String, cn As ADODB.Connection) As Integer()
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim ret(1 To 2)  As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select pointPointsAward, pointPointsMargin  from tblPoolpoints "
  sqlstr = sqlstr & " WHERE poolid = " & thisPool
  sqlstr = sqlstr & " AND pointTypeID IN ("
  sqlstr = sqlstr & "Select pointTypeID from tblPointtypes WHERE "
  sqlstr = sqlstr & "pointTypeDescription = '" & description & "')"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    ret(1) = rs!pointpointsAward
    ret(2) = nz(rs!pointpointsAward, 0)
    getPoolPoints = ret()
  Else
    getPoolPoints = ret()
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getLastDeelnemerID(thisAddress As Long, cn As ADODB.Connection)
'return the last deelnemerID for thisAddress
Dim sqlstr As String
Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select max(deelnemID) as ID from tblDeelnemers "
  sqlstr = sqlstr & " WHERE poolID = " & thisPool
  'sqlstr = sqlstr & " AND adresID = " & thisAddress
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getLastDeelnemerID = rs!id
  Else
    getLastDeelnemerID = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Function getLastPoolID(cn As ADODB.Connection)
'get the ID of the last pool that was added
Dim sqlstr As String
Dim rs As ADODB.Recordset
  sqlstr = "Select MAX(poolID) from tblPools"
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    getLastPoolID = rs!poolid
  Else
    getLastPoolID = 0
  End If
  rs.Close
  Set rs = Nothing
End Function

Sub renumRennerPos(deelnemID As Long, cn As ADODB.Connection)
 'renumber the renners in the tblDeelnemerPloegen for renners that are still with us
 Dim sqlstr As String
 Dim rs As ADODB.Recordset
 Dim i As Integer
 Set rs = New ADODB.Recordset
 sqlstr = "Select * from tblDeelnemerPloegen"
 sqlstr = sqlstr & " WHERE deelnemID = " & deelnemID
 sqlstr = sqlstr & " ORDER BY rennerpos"
 rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
 For i = 1 To rs.RecordCount
  rs!rennerpos = i
  rs.Update
  rs.MoveNext
 Next
 rs.Close
 Set rs = Nothing
End Sub


Function FindInTable(cn As ADODB.Connection, tbl As String, fieldName As String, lookUpText As Variant, Optional skipID As Long, Optional idField As String)
  Dim sqlstr As String
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from " & tbl
  sqlstr = sqlstr & " WHERE " & fieldName
  If IsNumeric(lookUpText) Then
    sqlstr = sqlstr & " = " & lookUpText
  Else
    sqlstr = sqlstr & " = '" & lookUpText & "'"
  End If
  If skipID Then
    sqlstr = sqlstr & " AND " & idField & " <> " & skipID
  End If
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  FindInTable = Not rs.EOF
End Function


Sub swapRenners(cn As ADODB.Connection, deelnemerID As Long, posA As Integer, posB As Integer)
'swap the position of the renners in a deelnemer team
  Dim savPos As Integer
  Dim rs As ADODB.Recordset
  Set rs = New ADODB.Recordset
  Dim sqlstr As String
  sqlstr = "Select * from tblDeelnemerPloegen Where deelnemId = " & deelnemerID
  sqlstr = sqlstr & " ORDER BY rennerPos"
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  If posA < posB Then
    rs.Find "rennerPos = " & posA
    If Not rs.EOF Then
      savPos = rs!rennerpos
      rs!rennerpos = posB
  Else
    findPos = posB
    rs!rennerpos = posB
    
    rs.MoveFirst
    rs.Find "rennerpos = " & posB
    If Not rs.EOF Then
      savPosB = rs!rennerpos
      rs!rennerpos = savPosA
      
  rs.Close
  Set rs = Nothing
End Sub
