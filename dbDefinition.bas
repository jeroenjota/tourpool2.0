Attribute VB_Name = "dbDefinition"
''''''''''
'Routines and functions for the definition and copying the database

Option Explicit

Public Const dbName = "tourpool2"


Function lclConn()
Dim fullPath As String
    fullPath = App.Path & "\" & dbName & ".mdb"
    lclConn = "PROVIDER='Microsoft.Jet.OLEDB.4.0';Data Source=" & fullPath & ";"
End Function

Function mySqlConn()
    
    Dim server As String
    Dim driver As String
    Dim cnStr As String
    Dim passwd As String
    passwd = "!xjer56!"
    'server = "192.168.178.14"
    server = "jotaservices.duckdns.org"
    driver = "{MariaDB ODBC 3.1 Driver}"
    mySqlConn = "DRIVER=" & driver & ";TCPIP=1;SERVER=" & server & ";DATABASE=" & dbName & ";UID=jeroen;PWD=" & passwd & ";port=3306;"

End Function

Function tableExists(srcTable As String, cn As ADODB.Connection)
'check if table exists in local database
Dim rs As ADODB.Recordset
    Set rs = cn.OpenSchema(adSchemaColumns, Array(Empty, Empty, srcTable, Empty))
    tableExists = Not (rs.BOF And rs.EOF)
    rs.Close
    Set rs = Nothing
End Function

Function recordsExist(tblName As String, cn As ADODB.Connection)
    Dim rs As ADODB.Recordset
    If tableExists(tblName, cn) Then
        Set rs = New ADODB.Recordset
        rs.Open "Select * from " & tblName, cn, adOpenKeyset, adLockReadOnly
        recordsExist = Not rs.EOF
    Else
        recordsExist = False
    End If
    rs.Close
    Set rs = Nothing
End Function

Function cFieldType(fldType As String) As Integer
'convert mySQL fldType to ADODB type
    Dim returnType As Integer
    If Left(fldType, 7) = "varchar" Then
        returnType = adVarWChar  'default type
    Else
        Select Case LCase(fldType)
        Case "Date", "time", "datetime", "timestamp"
            returnType = adDate
        Case "int(11)"
            returnType = adInteger
        Case "double"
            returnType = adDouble
        Case "decimal(19,4)"
            returnType = adCurrency
        Case "tinyint(3)", "tinyint(3) unsigned"
            returnType = adUnsignedTinyInt
        Case "tinyint(1)"
            returnType = adBoolean
        Case Else
            Stop
        End Select
    End If
    cFieldType = returnType
End Function


'create the database
Sub createDb()
    Dim setupDb As String
    Dim newDb As String
    Dim msg As String
    Dim fileName As String
    
    ' MDB to be created. In app.path
    newDb = App.Path & "\" & dbName & ".mdb"
    ' Drop the existing database, if any.
    If Dir(newDb) > "" Then
        msg = "Er is al een database " & newDb & vbNewLine
        msg = msg & "Wil je een kopie hiervan bewaren?" & vbNewLine
        If MsgBox(msg, vbYesNo, "Nieuwe database aanmaken") = vbYes Then
            FileCopy newDb, newDb & ".bak"
        End If
        Kill newDb 'remove the old db
    End If
    setupDb = App.Path & "\tourpoolSetup.mdb"
    FileCopy setupDb, newDb
End Sub

Sub fillDefaultValues()
'fill some tables with default values
'    Dim cn As ADODB.Connection
'    Dim rs As ADODB.Recordset
'    Dim sqlstr As String
'    Dim cmd As ADODB.Command
'    Dim tour As String
'    Dim orgID As Long
'    Set cn = New ADODB.Connection
'    With cn
'        .ConnectionString = lclConn
'        .Open
'    End With
'    If thisTour = 0 Then
'        MsgBox "Geen tour geselcteerd ??, Heel vreemd!", vbOKOnly + vbCritical, "Database Fout in fillDefaultValues"
'        Exit Sub
'    End If
'    tour = getTourInfo("description", cn)
'    'fill the first pool record
'    'get the OrganisationID - should be only one organisation
'    orgID = 1 'no longer need the organisation field, but just in case...
'    'create a first record in tblPools
'    sqlstr = "INSERT INTO tblPools (tourId, OrganisationId, poolName, poolFormsFrom, poolFormsTill, "
'    sqlstr = sqlstr & "poolCost, prizeHighDayScore, prizeHighDayPosition, prizeLowDayposition, "
'    sqlstr = sqlstr & "prizePercentageFirst, prizePercentageSecond, prizePercentageThird, prizePercentageFourth, "
'    sqlstr = sqlstr & "prizeLowFinalPosition ) VALUES ("
'    sqlstr = sqlstr & thisTour & ", " & orgID & ", '" & tour & " pool" & "', " & CDbl(Date) & ", " & CDbl(getTourInfo("tourStartDate", cn) - 7) & ", "
'    sqlstr = sqlstr & "10, 2.5, 1, 0.1, 50, 30, 20, 0, 10)"
'    cn.Execute sqlstr
'
'    'set the thisPool global variable
'    sqlstr = "Select * from tblPools order by poolID"
'    Set rs = New ADODB.Recordset
'    rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
'    rs.MoveLast
'    thisPool = rs!poolid
'
'    'default data for points
'    sqlstr = "INSERT into tblPoolPoints ( poolID, pointTypeId, pointPointsAward, pointPointsMargin )"
'    sqlstr = sqlstr & " Select " & thisPool & ", pointtypeId , pointDefaultPoints, pointDefaultMargin from tblpointTypes"
'    If Not getTourInfo("tourThirdPlace", cn) Then 'do not import 3rd and 4th place categories
'        sqlstr = sqlstr & " WHERE pointTypeCategory <> 7"
'        If getTourInfo("tourTeamCount", cn) <= 16 Then 'do not import 8th final categories
'            sqlstr = sqlstr & " AND pointTypeCategory <> 2"
'        End If
'    End If
'    cn.Execute sqlstr
'    If Not rs Is Nothing Then
'        If (rs.State And adStateOpen) = adStateOpen Then
'            rs.Close
'        End If
'        Set rs = Nothing
'    End If
'
'    If Not cn Is Nothing Then
'        If (cn.State And adStateOpen) = adStateOpen Then
'            cn.Close
'        End If
'        Set cn = Nothing
'    End If
End Sub

Sub makeTables(cn As ADODB.Connection)
    Dim sqlstr As String
    'OBSOLETE, just leave it here
    'new local tables just in case
    
    'address table
    sqlstr = "CREATE TABLE tblAddressen ( "
    sqlstr = sqlstr & "addressID INTEGER NOT NULL, "
    sqlstr = sqlstr & "firstname VARCHAR(50), "
    sqlstr = sqlstr & "middlename VARCHAR(32), "
    sqlstr = sqlstr & "lastname VARCHAR(50), "
    sqlstr = sqlstr & "shortname VARCHAR(24), "
    sqlstr = sqlstr & "address VARCHAR(50), "
    sqlstr = sqlstr & "postalcode VARCHAR(10), "
    sqlstr = sqlstr & "city VARCHAR(50), "
    sqlstr = sqlstr & "telephone VARCHAR(20), "
    sqlstr = sqlstr & "email VarChar(255) "
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblAddress (addressID) WITH PRIMARY"
    cn.Execute sqlstr
    
    'competitorpoints
    sqlstr = "CREATE TABLE tblCompetitorPoints ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL,"
    sqlstr = sqlstr & "stageNumber INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointsGeneral INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsPunten INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsYoung INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsBerg INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsEtappeUitslag INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsDayTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointsGrandTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "poisitionDay INTEGER,"
    sqlstr = sqlstr & "positionTotal INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "moneyDay DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayPosition DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayLast DECIMAL(19,4),"
    sqlstr = sqlstr & "moneyTotal DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & "moneyDayTotal DECIMAL(19,4) DEFAULT 0,"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    
    'deelnemers
    sqlstr = "CREATE TABLE tblCompetitors ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL, "
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "addressID INTEGER NOT NULL,"
    sqlstr = sqlstr & "nickName VARCHAR(50) NOT NULL,"
    sqlstr = sqlstr & "payed YESNO DEFAULT 0,"
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblCompetitors (competitorID) WITH PRIMARY"
    cn.Execute sqlstr
    
    'deelnemerPloegen
    sqlstr = "CREATE TABLE tblCompetitorRenners ("
    sqlstr = sqlstr & "competitorID INTEGER NOT NULL, "
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "rennerID INTEGER NOT NULL,"
    sqlstr = sqlstr & ") "
    cn.Execute sqlstr
    
    
    'pool points
    sqlstr = "CREATE TABLE tblPoolPoints ("
    sqlstr = sqlstr & "poolid INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointTypeID INTEGER NOT NULL,"
    sqlstr = sqlstr & "pointPointsAward INTEGER DEFAULT 0,"
    sqlstr = sqlstr & "pointPointsMargin byte DEFAULT 0 )"
    cn.Execute sqlstr

    'pools
    sqlstr = "CREATE TABLE tblPools ("
    sqlstr = sqlstr & "poolID INTEGER NOT NULL DEFAULT 0,"
    sqlstr = sqlstr & "tourID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "organisationID INTEGER DEFAULT NULL,"
    sqlstr = sqlstr & "poolName varchar(50) DEFAULT NULL,"
    sqlstr = sqlstr & "poolStartAcceptForms datetime DEFAULT NULL,"
    sqlstr = sqlstr & "poolEndAcceptForms datetime DEFAULT NULL,"
    sqlstr = sqlstr & "poolCost decimal(19,4) DEFAULT 10.0000,"
    sqlstr = sqlstr & "prizeHighDayScore decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizeHighDayPosition decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizeLowDayPosition decimal(19,4) DEFAULT 0.0000,"
    sqlstr = sqlstr & "prizePercentageFirst double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentageSecond double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentageThird double DEFAULT 0,"
    sqlstr = sqlstr & "prizePercentageFourth double DEFAULT 0,"
    sqlstr = sqlstr & "prizeLowFinalPosition decimal(19,4) DEFAULT 0.0000"
    sqlstr = sqlstr & ")"
    cn.Execute sqlstr
    sqlstr = "CREATE INDEX PrimaryKey on tblPools (poolID) WITH PRIMARY"
    cn.Execute sqlstr
    
End Sub

Sub copyDefaultPoints()
'copy the default points table to current pool
Dim cnFrom As ADODB.Connection
Dim cnTo As ADODB.Connection
Dim cnStr As String
Dim fileName As String
Dim sqlstr As String
Dim rs As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim msg As String
Dim fld As field
  
  Set cnTo = New ADODB.Connection
  Set cnFrom = New ADODB.Connection
  Set rs = New ADODB.Recordset
  Set rs2 = New ADODB.Recordset
  
  fileName = App.Path & "\vbpSetup.mdb"
  If Dir(fileName) = "" Then
    msg = "Bestand 'vbpSetup.mdb' niet gevonden in installatie map"
    msg = msg & vbNewLine & "Kan standaard puntentabel niet kopieëren"
    MsgBox msg, vbOKOnly + vbCritical, "Puntentabel"
    Exit Sub
  End If
  
  cnStr = "PROVIDER='Microsoft.Jet.OLEDB.4.0';Data Source=" & fileName & ";"
  With cnFrom
    .ConnectionString = cnStr
    .Open
  End With
  
  With cnTo
    .ConnectionString = lclConn
    .Open
    .CursorLocation = adUseClient
  End With
  'selete current records
  sqlstr = "Delete from tblPoolPoints WHERE poolid = " & thisPool
  cnTo.Execute sqlstr
  
  sqlstr = "Select * from tblPoolPoints"
  rs2.Open sqlstr, cnTo, adOpenKeyset, adLockOptimistic
  
  sqlstr = "Select * from tblPoolPoints where poolID = 0"
  rs.Open sqlstr, cnFrom, adOpenKeyset, adLockReadOnly
  'copy the records
  Do While Not rs.EOF
    rs2.AddNew
      For Each fld In rs.Fields
        rs2(fld.Name) = rs(fld.Name)
        If fld.Name = "poolid" Then rs2(fld.Name) = thisPool
      Next
    rs2.Update
    rs.MoveNext
  Loop
  'tidy up
  If (rs.State And adStateOpen) = adStateOpen Then rs.Close
  Set rs = Nothing
  If (rs2.State And adStateOpen) = adStateOpen Then rs2.Close
  Set rs2 = Nothing
  If (cnTo.State And adStateOpen) = adStateOpen Then cnTo.Close
  Set cnTo = Nothing
  If (cnFrom.State And adStateOpen) = adStateOpen Then cnFrom.Close
  Set cnFrom = Nothing
End Sub
