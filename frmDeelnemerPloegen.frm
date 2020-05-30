VERSION 5.00
Object = "{0DF5D14C-08DD-4806-8BE2-B59CB924CFC9}#1.7#0"; "VBCCR16.OCX"
Begin VB.Form frmDeelnemerPloegen 
   Caption         =   "Deelnemer teams"
   ClientHeight    =   7740
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7920
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
   ScaleHeight     =   7740
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VBCCR16.CommandButtonW btnSaveAdres 
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4440
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      Caption         =   "Adresgegevens opslaan"
   End
   Begin VBCCR16.CommandButtonW btnNewPool 
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   5040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   1508
      Caption         =   "Voeg pool toe voor dit adres"
   End
   Begin VBCCR16.CommandButtonW btnNewAddress 
      Height          =   375
      Left            =   1200
      TabIndex        =   19
      Top             =   480
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      Caption         =   "Nieuw adres"
   End
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   5040
      TabIndex        =   17
      Top             =   480
      Width           =   2775
      Begin VBCCR16.CommandButtonW btnMoveUp 
         Height          =   615
         Left            =   2280
         TabIndex        =   28
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
         Caption         =   "^"
         Picture         =   "frmDeelnemerPloegen.frx":0000
      End
      Begin VBCCR16.ListBoxW lstTeamRenners 
         Height          =   2595
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   4577
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
      End
      Begin VBCCR16.ComboBoxW cmbRenners 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   556
      End
      Begin VBCCR16.ListBoxW listReserves 
         Height          =   1815
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   3201
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
      End
      Begin VBCCR16.TextBoxW txtPoolName 
         Height          =   405
         Left            =   120
         TabIndex        =   10
         Top             =   420
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   714
      End
      Begin VBCCR16.CommandButtonW btnDeleteRenner 
         Height          =   615
         Left            =   2280
         TabIndex        =   29
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
         Caption         =   "X"
         Picture         =   "frmDeelnemerPloegen.frx":0452
      End
      Begin VBCCR16.CommandButtonW btnMoveDn 
         Height          =   615
         Left            =   2280
         TabIndex        =   30
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
         Caption         =   "v"
         Picture         =   "frmDeelnemerPloegen.frx":08A4
      End
      Begin VBCCR16.LabelW lblTotals 
         Height          =   255
         Left            =   120
         TabIndex        =   31
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
         TabIndex        =   27
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Poolnaam"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   1695
      End
      Begin VBCCR16.LabelW LabelW3 
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         Caption         =   "Reserve renners"
      End
   End
   Begin VBCCR16.CommandButtonW btnSave 
      Height          =   390
      Left            =   4440
      TabIndex        =   15
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
      Caption         =   "Opslaan"
   End
   Begin VBCCR16.ComboBoxW cmbAddresses 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
   End
   Begin VBCCR16.CommandButtonW btnClose 
      Height          =   390
      Left            =   6120
      TabIndex        =   16
      Top             =   7200
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   688
      Caption         =   "Sluiten"
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1740
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   1740
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   2475
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   3225
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   3960
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   714
   End
   Begin VBCCR16.ListBoxW lstCompetitorTeams 
      Height          =   5910
      Left            =   2880
      TabIndex        =   9
      Top             =   960
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   10425
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin VBCCR16.LabelW LabelW4 
      Height          =   255
      Left            =   3000
      TabIndex        =   26
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      BackStyle       =   0
      Caption         =   "Deelnemer pools"
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Voornaam"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "tussenvg"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   23
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Achternaam"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   22
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Telefoon"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   2925
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   3660
      Width           =   2655
   End
   Begin VBCCR16.LabelW LabelW1 
      Height          =   270
      Left            =   240
      TabIndex        =   14
      Top             =   600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   476
      BackStyle       =   0
      Caption         =   "Adressenlijst"
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DeelnemerTeams"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Tag             =   "kop"
      Top             =   120
      Width           =   8055
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
Dim teamSize As Integer
Dim reserveSize As Integer

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnNewPool_Click()
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
  End If
  Me.btnNewPool.Caption = "Nieuwe pool voor " & vbNewLine & Me.cmbAddresses.Text
End Sub

Sub initForm()
  Dim sqlstr As String
  'adressenlijst
  sqlstr = "Select competitorId as id, trim(lastname & ', ' & firstname & ' ' & middlename) as Name "
  sqlstr = sqlstr & " from tblCompetitors order by lastName"
  FillCombo Me.cmbAddresses, sqlstr, cn, "name", "id"
  
  sqlstr = "Select deelnemID as id, roepnaam "
  sqlstr = sqlstr & " from tblDeelnemers WHERE poolID = " & thisPool
  fillList Me.lstCompetitorTeams, sqlstr, cn, "roepnaam", "id"
  fillRennerCombo
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
  Dim sqlstr As String
  Dim adoCmd As ADODB.Command
  Set adoCmd = New ADODB.Command
  If Me.lstTeamRenners.ListCount + Me.listReserves.ListCount >= teamSize + reserveSize Then Exit Sub
  sqlstr = "INSERT INTO tblDeelnemerPloegen (deelnemID, rennerID, rennerPos) VALUES (?, ?, ?)"
  With adoCmd
    .ActiveConnection = cn
    .CommandType = adCmdText
    .CommandText = sqlstr
    .Prepared = True
    .Parameters.Append .CreateParameter("DeelnemID", adInteger, adParamInput, , thisDeelnemer)
    .Parameters.Append .CreateParameter("rennerID", adInteger, adParamInput, , Me.cmbRenners.ItemData(Me.cmbRenners.ListIndex))
    .Parameters.Append .CreateParameter("posities", adInteger, adParamInput, , Me.lstTeamRenners.ListCount + Me.listReserves.ListCount + 1)
    .Execute
  End With
  updatePoolRenners
  fillRennerCombo
End Sub

Private Sub Form_Load()
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .Open
  End With
  
  teamSize = getPoolInfo("deelNemPloegAant", cn)
  reserveSize = getPoolInfo("deelNemresAant", cn) + getPoolInfo("deelNemresExtra", cn)
  
  Me.btnMoveUp.Caption = StringFromCodepoint(&H2B9D)
  Me.btnMoveDn.Caption = StringFromCodepoint(&H2B9F)
  Me.btnDeleteRenner.Caption = StringFromCodepoint(&H26D2)
  initForm
  
  UnifyForm Me
  centerForm Me
  
  If Me.lstCompetitorTeams.ListCount > 0 Then
    Me.lstCompetitorTeams.ListIndex = 0
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    
End Sub

Private Sub lstCompetitorTeams_Click()
  thisDeelnemer = Me.lstCompetitorTeams.ItemData(Me.lstCompetitorTeams.ListIndex)
  
  updatePoolRenners
  fillRennerCombo
End Sub

Sub updatePoolRenners()
  Dim sqlstr As String
  
  sqlstr = "Select top " & teamSize & " id, trim(aNaam & ', ' & vNaam & ' ' & tNaam) as Renner"
  sqlstr = sqlstr & " from tblRenners a INNER JOIN tblDeelnemerPloegen b on a.id = b.rennerID"
  sqlstr = sqlstr & " WHERE deelnemID = " & thisDeelnemer
  sqlstr = sqlstr & " ORDER BY rennerPos "
  fillList Me.lstTeamRenners, sqlstr, cn, "Renner", "id"
  
  sqlstr = "Select id, trim(aNaam & ', ' & vNaam & ' ' & tNaam) as Renner"
  sqlstr = sqlstr & " from tblRenners a INNER JOIN tblDeelnemerPloegen b on a.id = b.rennerID"
  sqlstr = sqlstr & " WHERE deelnemID = " & thisDeelnemer
  sqlstr = sqlstr & " and rennerPos > " & teamSize
  sqlstr = sqlstr & " ORDER BY rennerPos "
  fillList Me.listReserves, sqlstr, cn, "Renner", "id"
  
  Me.lblTotals = Me.lstTeamRenners.ListCount & "/" & teamSize & " renners; " & Me.listReserves.ListCount & "/" & reserveSize & " reserves"
  Me.cmbRenners.Enabled = Me.lstTeamRenners.ListCount + Me.listReserves.ListCount < teamSize + reserveSize

End Sub
