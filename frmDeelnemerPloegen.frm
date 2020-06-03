VERSION 5.00
Object = "{0DF5D14C-08DD-4806-8BE2-B59CB924CFC9}#1.7#0"; "VBCCR16.OCX"
Begin VB.Form frmDeelnemerPloegen 
   Caption         =   "Deelnemer teams"
   ClientHeight    =   9765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VBCCR16.CommandButtonW btnSaveAdres 
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Opslaan"
   End
   Begin VBCCR16.CommandButtonW btnNewPool 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   1085
      Enabled         =   0   'False
      Caption         =   "Voeg pool toe voor dit adres"
   End
   Begin VBCCR16.CommandButtonW btnNewAddress 
      Height          =   375
      Left            =   3240
      TabIndex        =   16
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      Caption         =   "Nieuw"
   End
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   2760
      TabIndex        =   15
      Top             =   2160
      Width           =   3615
      Begin VBCCR16.ListView lstDeelnemRenners 
         Height          =   5295
         Left            =   120
         TabIndex        =   27
         Top             =   1440
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   9340
         View            =   3
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         ClickableColumnHeaders=   0   'False
         AutoSelectFirstItem=   0   'False
      End
      Begin VBCCR16.CommandButtonW btnMoveUp 
         Height          =   615
         Left            =   3120
         TabIndex        =   23
         Top             =   1920
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Caption         =   "^"
         Picture         =   "frmDeelnemerPloegen.frx":0000
      End
      Begin VBCCR16.ComboBoxW cmbRenners 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
      End
      Begin VBCCR16.TextBoxW txtPoolName 
         Height          =   405
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   714
      End
      Begin VBCCR16.CommandButtonW btnDeleteRenner 
         Height          =   615
         Left            =   3120
         TabIndex        =   24
         Top             =   2700
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Caption         =   "X"
         Picture         =   "frmDeelnemerPloegen.frx":0452
      End
      Begin VBCCR16.CommandButtonW btnMoveDn 
         Height          =   615
         Left            =   3120
         TabIndex        =   25
         Top             =   3360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         Caption         =   "v"
         Picture         =   "frmDeelnemerPloegen.frx":08A4
      End
      Begin VBCCR16.LabelW lblTotals 
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   6240
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Selecteer renner"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Poolnaam"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
   End
   Begin VBCCR16.CommandButtonW btnSave 
      Height          =   390
      Left            =   2880
      TabIndex        =   13
      Top             =   9120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
      Caption         =   "Opslaan"
   End
   Begin VBCCR16.ComboBoxW cmbAddresses 
      Height          =   315
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   556
   End
   Begin VBCCR16.CommandButtonW btnClose 
      Height          =   390
      Left            =   4560
      TabIndex        =   14
      Top             =   9120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
      Caption         =   "Sluiten"
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   0
      Left            =   840
      TabIndex        =   2
      Top             =   1140
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   1140
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   2
      Left            =   3720
      TabIndex        =   4
      Top             =   1140
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   3
      Left            =   840
      TabIndex        =   5
      Top             =   1560
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   4
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   714
   End
   Begin VBCCR16.ListBoxW lstCompetitorTeams 
      Height          =   6300
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   11113
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin VBCCR16.LabelW LabelW4 
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Deelnemer pools"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Naam"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefoon"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   1605
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   255
      Index           =   4
      Left            =   3000
      TabIndex        =   17
      Top             =   1620
      Width           =   2655
   End
   Begin VBCCR16.LabelW LabelW1 
      Height          =   270
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   476
      BackStyle       =   0
      Caption         =   "Zoek"
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Deelnemers && Ploegen"
      Height          =   375
      Left            =   -360
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "frmDeelnemerPloegen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim thisDeelnemer As Long
Dim thisAddress As Long
Dim teamSize As Integer
Dim reserveSize As Integer
Dim teamBackupSize As Integer

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnDeleteRenner_Click()
  Dim renID As Long
  Dim sqlstr As String
  Dim i As Integer
  For i = 1 To Me.lstDeelnemRenners.ListItems.Count
    If Me.lstDeelnemRenners.ListItems(i).Selected Then
      renID = Me.lstDeelnemRenners.ListItems(i).Text
      Exit For
    End If
  Next
  If renID <> 0 Then
    sqlstr = "Delete from tblDeelnemerPloegen"
    sqlstr = sqlstr & " WHERE deelnemID = " & thisDeelnemer
    sqlstr = sqlstr & " AND rennerID = " & renID
    cn.Execute sqlstr
    'renumber
    renumRennerPos thisDeelnemer, cn
    
    'update lists
    updateDeelnemerRenners
    fillRennerCombo
  End If

End Sub

Private Sub btnMoveDn_Click()
  Dim posA As Integer, posB As Integer
  Dim savPos As Integer
  Dim sqlstr As String
  Dim i As Integer
  With Me.lstDeelnemRenners
    For i = 1 To .ListItems.Count - 1
      If .ListItems(i).Selected Then
        savPos = .ListItems(i).SubItems(1)
        .ListItems(i).SubItems(1) = .ListItems(i + 1).SubItems(1)
        .ListItems(i + 1).SubItems(1) = savPos
        Exit For
      End If
    Next
  End With
  'update lists
  saveDeelnemerRenners
  updateDeelnemerRenners
  With Me.lstDeelnemRenners
    If i < .ListItems.Count Then
      Set .SelectedItem = .ListItems(i + 1)
      '.ListItems(i + 1).Selected = True
      .HideSelection = False
      '.ListItems(i + 1).EnsureVisible
    End If
  End With
End Sub

Private Sub btnMoveUp_Click()
  Dim posA As Integer, posB As Integer
  Dim savPos As Integer
  Dim sqlstr As String
  Dim i As Integer
  With Me.lstDeelnemRenners
    For i = 2 To .ListItems.Count
      If .ListItems(i).Selected Then
        savPos = .ListItems(i).SubItems(1)
        .ListItems(i).SubItems(1) = .ListItems(i - 1).SubItems(1)
        .ListItems(i - 1).SubItems(1) = savPos
        Exit For
      End If
    Next
  End With
  'update lists
  saveDeelnemerRenners
  updateDeelnemerRenners
  With Me.lstDeelnemRenners
    If i < .ListItems.Count Then
      Set .SelectedItem = .ListItems(i - 1)
      .HideSelection = False
    End If
  End With
End Sub

Sub saveDeelnemerRenners()
  
  Dim sqlstr As String
  Dim i As Integer
  sqlstr = "Delete from tblDeelnemerPloegen WHERE deelnemID = " & thisDeelnemer
  cn.Execute sqlstr
  With Me.lstDeelnemRenners
    For i = 1 To .ListItems.Count
      saveDeelnemerRenner .ListItems(i), .ListItems(i).SubItems(1)
    Next
  End With
End Sub

Sub saveDeelnemerRenner(rennerID As Long, pos As Integer)
  Dim sqlstr As String
  Dim adoCmd As ADODB.Command
  Set adoCmd = New ADODB.Command
  
  sqlstr = "INSERT INTO tblDeelnemerPloegen (deelnemID, rennerID, rennerPos) VALUES (?, ?, ?)"
  With adoCmd
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = sqlstr
    .Prepared = True
    .Parameters.Append .CreateParameter("DeelnemID", adInteger, adParamInput, , thisDeelnemer)
    .Parameters.Append .CreateParameter("rennerID", adInteger, adParamInput, , rennerID)
    .Parameters.Append .CreateParameter("posities", adInteger, adParamInput, , pos)
    .Execute
  End With
  Set adoCmd = Nothing

End Sub

Private Sub btnNewPool_Click()
  Dim sqlstr As String
  Dim adoCmd As ADODB.Command
  Set adoCmd = New ADODB.Command
  sqlstr = "INSERT INTO tblDeelnemers (poolID, adresID) VALUES (?, ?)"
  With adoCmd
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = sqlstr
    .Prepared = True
    .Parameters.Append .CreateParameter("poolID", adInteger, adParamInput, , thisPool)
    .Parameters.Append .CreateParameter("adresID", adInteger, adParamInput, , thisAddress)
    .Execute
  End With
  thisDeelnemer = getLastDeelnemerID(thisAddress, cn)
  Set adoCmd = Nothing
  Me.lstDeelnemRenners.ListItems.Clear
  Me.txtPoolName.Text = ""
  Me.txtPoolName.SetFocus
End Sub

Private Sub cmbAddresses_Click()
  'get the addresss data
  Dim i As Integer
  Dim sqlstr As String
  Set rs = New ADODB.Recordset
  sqlstr = "Select * from tblCompetitors where competitorID = " & Me.cmbAddresses.ItemData(Me.cmbAddresses.ListIndex)
  rs.Open sqlstr, cn, adOpenKeyset, adLockReadOnly
  If Not rs.EOF Then
    For i = 1 To 5
      Me.txtField(i - 1) = nz(rs.Fields(i), "")
    Next
    Me.btnNewPool.Enabled = True
  End If
  thisAddress = Me.cmbAddresses.ItemData(Me.cmbAddresses.ListIndex)
  Me.btnNewPool.Caption = "Nieuwe pool voor " & vbNewLine & Me.cmbAddresses.Text
End Sub

Sub initForm()
  Dim sqlstr As String
  'adressen combo
  sqlstr = "Select competitorId as id, trim(lastname & ', ' & firstname & ' ' & middlename) as Name "
  sqlstr = sqlstr & " from tblCompetitors order by lastName"
  FillCombo Me.cmbAddresses, sqlstr, cn, "name", "id"
  'list with pools
  fillPoolList
  'combobox with renners
  fillRennerCombo

End Sub

Sub fillPoolList()
Dim sqlstr As String
  sqlstr = "Select deelnemID as id, roepnaam "
  sqlstr = sqlstr & " from tblDeelnemers WHERE poolID = " & thisPool
  sqlstr = sqlstr & " ORDER BY roepnaam"
  fillList Me.lstCompetitorTeams, sqlstr, cn, "roepnaam", "id"

End Sub

Sub fillRennerCombo()
Dim sqlstr As String
  sqlstr = "Select id, trim(aNaam & ', ' & vNaam & ' ' & tNaam) as Renner"
  sqlstr = sqlstr & " from tblRenners a INNER JOIN tblRondeRenners b on a.id = b.rennerID"
  sqlstr = sqlstr & " WHERE b.tourID = " & thisTour
  sqlstr = sqlstr & " AND id NOT in ( SELECT rennerID from tblDeelnemerPloegen "
  sqlstr = sqlstr & " WHERE deelnemID = " & thisDeelnemer & ")"
  sqlstr = sqlstr & " ORDER BY aNaam "
  FillCombo Me.cmbRenners, sqlstr, cn, "Renner", "id"

End Sub

Private Sub cmbRenners_Click()
  If Me.lstDeelnemRenners.ListItems.Count >= teamSize + reserveSize + teamBackupSize Then Exit Sub
  saveDeelnemerRenner Me.cmbRenners.ItemData(Me.cmbRenners.ListIndex), Me.lstDeelnemRenners.ListItems.Count + 1
  updateDeelnemerRenners
  fillRennerCombo
End Sub

Private Sub Form_Load()
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .Open
  End With
  
  teamSize = getPoolInfo("deelNemPloegAant", cn)
  reserveSize = getPoolInfo("deelNemResAant", cn)
  teamBackupSize = getPoolInfo("deelNemResExtra", cn)
  
  Me.btnMoveUp.Caption = StringFromCodepoint(&H2B9D)
  Me.btnMoveDn.Caption = StringFromCodepoint(&H2B9F)
  Me.btnDeleteRenner.Caption = StringFromCodepoint(&H26D2)
  initForm
  
  UnifyForm Me
  centerForm Me
  
'  If Me.lstCompetitorTeams.ListCount > 0 Then
'    Me.lstCompetitorTeams.ListIndex = 0
'  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    If Not rs Is Nothing Then
      If (rs.State And adStateOpen) = adStateOpen Then
        rs.Close
      End If
      Set rs = Nothing
    End If
End Sub

Private Sub lstCompetitorTeams_Click()
  thisDeelnemer = Me.lstCompetitorTeams.ItemData(Me.lstCompetitorTeams.ListIndex)
  getDeelnemAdres
  Me.txtPoolName = Me.lstCompetitorTeams
  updateDeelnemerRenners
  fillRennerCombo
End Sub

Sub getDeelnemAdres()
  
  'get the address data for this deelnemer
  Dim i As Integer
  thisAddress = nz(getDeelnemerInfo(thisDeelnemer, "adresID", cn), 0)
  If thisAddress Then
    For i = 0 To Me.cmbAddresses.ListCount - 1
      If Me.cmbAddresses.ItemData(i) = thisAddress Then
        Me.cmbAddresses.ListIndex = i
        Exit For
      End If
    Next
  End If
  
End Sub

Sub updateDeelnemerRenners()
  Dim sqlstr As String
  Dim i As Integer, r As Integer
  Set rs = New ADODB.Recordset
  sqlstr = "Select id, rennerpos as nr, trim(aNaam & ', ' & vNaam & ' ' & tNaam) as Renner"
  sqlstr = sqlstr & " from tblRenners a INNER JOIN tblDeelnemerPloegen b on a.id = b.rennerID"
  sqlstr = sqlstr & " WHERE deelnemID = " & thisDeelnemer
  sqlstr = sqlstr & " ORDER BY rennerPos "
  rs.Open sqlstr, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
  With Me.lstDeelnemRenners
    .ColumnHeaders.Clear
    .ListItems.Clear
    .View = LvwViewReport
    For i = 0 To rs.Fields.Count - 1
      .ColumnHeaders.Add , , rs.Fields(i).Name
    Next
    Do While Not rs.EOF      '(without ItemData)
      r = r + 1
      With .ListItems.Add(, , rs.Fields(0))
        .Bold = r <= teamSize
        For i = 1 To rs.Fields.Count - 1
          .SubItems(i) = rs.Fields(i)
          .ListSubItems(i).Bold = r <= teamSize
          If r > teamSize Then .ListSubItems(i).ForeColor = vbRed
          If r > teamSize + reserveSize Then .ListSubItems(i).ForeColor = vbBlue
         ' .SubItems(i).Text.ForeColor = vbBlue
        Next
      End With
      rs.MoveNext
    Loop
    .ColumnHeaders(1).Width = 0
    .ColumnHeaders(2).AutoSize (LvwColumnHeaderAutoSizeToItems)
    .ColumnHeaders(3).Width = .Width - .ColumnHeaders(2).Width
  End With
  
End Sub

Private Sub lstReserves_Click()
  'Me.lstTeamRenners.ListIndex = -1
End Sub

Private Sub lstTeamRenners_Click()
  'Me.lstReserves.ListIndex = -1
End Sub

Public Sub FillDeelnemerRenners(objLV As ListView, _
                     strSQL As String, _
                     cn As ADODB.Connection, _
                     strFieldToShow As String, _
                     Optional hideID As Boolean)

'Fills a combobox with values from a database

'adapted code from VBforums
  Dim i As Integer
  Dim oRS As ADODB.Recordset  'Load the data
  Set oRS = New ADODB.Recordset
  oRS.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly, adCmdText
  If oRS.EOF Then
      MsgBox "Geen records in recordset", vbCritical + vbOKOnly, "FillCombo"
      Exit Sub
  End If
  With objLV          'Fill the combo box
  End With
  
  oRS.Close                 'Tidy up
  Set oRS = Nothing

End Sub

Private Sub lstDeelnemRenners_ItemSelect(ByVal Item As VBCCR16.LvwListItem, ByVal Selected As Boolean)
  Me.btnDeleteRenner.Enabled = Item.Selected
  Me.btnMoveUp.Enabled = Item.Index > 1
  Me.btnMoveDn.Enabled = Item.Index < Me.lstDeelnemRenners.ListItems.Count

End Sub

Private Sub txtPoolName_LostFocus()
  'save this pool name
  Dim sqlstr As String
  Dim msg As String
  If FindInTable(cn, "tblDeelnemers", "roepnaam", Me.txtPoolName, thisDeelnemer, "deelnemID") Then
    msg = "De naam '" & Me.txtPoolName & "' bestaat al, voer een andere in"
    MsgBox msg, vbOKOnly + vbCritical, "Poolnaam"
    Me.txtPoolName.SetFocus
    Exit Sub
  End If
  sqlstr = "UPDATE tblDeelnemers SET roepnaam= '" & Me.txtPoolName & "'"
  sqlstr = sqlstr & " WHERE deelnemID = " & thisDeelnemer
  cn.Execute sqlstr
  fillPoolList
End Sub
