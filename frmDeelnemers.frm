VERSION 5.00
Object = "{3E5D9624-07F7-4D22-90F8-1314327F7BAC}#1.0#0"; "VBFLXGRD14.OCX"
Object = "{0DF5D14C-08DD-4806-8BE2-B59CB924CFC9}#1.7#0"; "VBCCR16.OCX"
Begin VB.Form frmCompetitors 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Adressen"
   ClientHeight    =   7935
   ClientLeft      =   645
   ClientTop       =   1215
   ClientWidth     =   11775
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   Tag             =   "adressen"
   Begin VB.ComboBox cmbAdresList 
      Height          =   315
      Left            =   240
      TabIndex        =   17
      Top             =   7440
      Width           =   2655
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Wissen"
      Height          =   495
      Left            =   9360
      Picture         =   "frmDeelnemers.frx":0000
      TabIndex        =   11
      Top             =   7320
      Width           =   1095
   End
   Begin VBFLXGRD14.VBFlexGrid grdCompetitors 
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   9763
      SelectionMode   =   3
      Redraw          =   0   'False
      DirectionAfterReturn=   2
   End
   Begin VB.CommandButton btnNew 
      Caption         =   "Nieuw"
      Height          =   375
      Left            =   10560
      Picture         =   "frmDeelnemers.frx":0442
      TabIndex        =   9
      Top             =   1250
      Width           =   1095
   End
   Begin VB.CommandButton btnSave 
      Caption         =   "Opslaan"
      Height          =   375
      Left            =   9360
      Picture         =   "frmDeelnemers.frx":0884
      TabIndex        =   8
      Top             =   1250
      Width           =   1095
   End
   Begin VB.PictureBox picDummy 
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Sluiten"
      Height          =   495
      Left            =   10560
      TabIndex        =   0
      Top             =   7320
      Width           =   1095
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   1
      Left            =   1440
      TabIndex        =   13
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   2
      Left            =   2760
      TabIndex        =   14
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   3
      Left            =   4080
      TabIndex        =   15
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
   End
   Begin VBCCR16.TextBoxW txtField 
      Height          =   405
      Index           =   4
      Left            =   5400
      TabIndex        =   16
      Top             =   840
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   714
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Adreslijst"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Tag             =   "kop"
      Top             =   120
      Width           =   11535
   End
   Begin VB.Label Label1 
      Caption         =   "Email"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   6
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Telefoon"
      Height          =   255
      Index           =   3
      Left            =   4080
      TabIndex        =   5
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Achternaam"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "tussenvg"
      Height          =   255
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Voornaam"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   735
   End
End
Attribute VB_Name = "frmCompetitors"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim currentID As Long
Dim editMode As Boolean

Private Sub btnClose_Click()
  Unload Me
End Sub

Private Sub btnNew_Click()
  Dim i As Integer
  For i = 0 To Me.txtField.Count - 1
    Me.txtField(i) = ""
  Next
  currentID = 0
  editMode = True
  Me.btnDelete.Visible = True
End Sub

Private Sub btnSave_Click()
Dim i As Integer
  With rs
  
    If currentID = 0 Then
      .AddNew
      currentID = !id
    Else
      .Find "id = " & currentID
    End If
    If Not .EOF Then
      For i = 1 To rs.Fields.Count - 1
        .Fields(i) = Me.txtField(i - 1)
      Next
      .Update
    End If
  End With
  editMode = False
  initForm
End Sub

Private Sub cmbAdresList_Click()
  Dim i As Integer
  Dim id As Long
  With Me.grdCompetitors
    id = Me.cmbAdresList.ItemData(Me.cmbAdresList.ListIndex)
    For i = 0 To Me.grdCompetitors.Rows - 1
      If .TextMatrix(i + 1, 0) = id Then
        Exit For
      End If
    Next
    .Row = i + 1
    .TopRow = i + 1
  End With
End Sub

Private Sub Form_Load()
Dim sqlstr As String
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .Open
  End With
  sqlstr = "Select competitorID as id, firstName as Voornaam, middleName as Tussvg, "
  sqlstr = sqlstr & " lastName as Achternaam, telephone as telefoon, email as Email from tblCompetitors"
  sqlstr = sqlstr & " ORDeR BY firstName, lastName"
  Set rs = New ADODB.Recordset
  rs.Open sqlstr, cn, adOpenKeyset, adLockOptimistic
  initForm
  UnifyForm Me
  centerForm Me
End Sub

Sub initForm()
Dim i As Integer, j As Integer
Dim grdWidth As Integer
Dim savColWidth(6)
  Dim sqlstr As String
  sqlstr = "Select competitorID, trim(lastName & ', ' & firstName & ' ' & middleName) as name from tblCompetitors order by lastname"
  FillCombo Me.cmbAdresList, sqlstr, cn, "name", "competitorID"
  'fill the grid
  i = 0
'  Me.Visible = False
  
  With Me.grdCompetitors
  .Redraw = False
    .Clear
    .Cols = rs.Fields.Count
    For j = 0 To rs.Fields.Count - 1
      If Not IsNull(rs.Fields(j).Name) Then
        .TextMatrix(i, j) = rs.Fields(j).Name
        .ColAlignment(j) = flexAlignLeftCenter
        If j > 0 Then Me.Label1(j - 1).Caption = rs.Fields(j).Name
      End If
    Next
    If rs.RecordCount = 0 Then Exit Sub
    .Rows = rs.RecordCount + 1
    rs.MoveFirst
    Me.picDummy.Font.Name = Me.grdCompetitors.Font.Name
    Me.picDummy.Font.Size = Me.grdCompetitors.Font.Size + 1
    Me.picDummy.ScaleMode = vbTwips
    Do While Not rs.EOF
      i = i + 1
      For j = 0 To rs.Fields.Count - 1
        If Not IsNull(rs.Fields(j).Value) Then
          .TextMatrix(i, j) = rs.Fields(j).Value
        Else
          .TextMatrix(i, j) = ""
        End If
        If j = 0 Then
          .ColWidth(j) = 0
        Else
          If Me.picDummy.TextWidth(.TextMatrix(i, j) & "XXX") > savColWidth(j) Then
            savColWidth(j) = Me.picDummy.TextWidth(.TextMatrix(i, j) & "XXX")
          End If
        End If
      Next
      DoEvents
      rs.MoveNext
    Loop
    Me.Label1(0).Left = Me.grdCompetitors.Left
    Me.txtField(0).Left = Me.grdCompetitors.Left
    For j = 1 To rs.Fields.Count - 1
      .ColWidth(j) = savColWidth(j)
      Me.Label1(j - 1).Width = savColWidth(j)
      Me.txtField(j - 1).Width = savColWidth(j)
      If j > 1 Then
        Me.Label1(j - 1).Left = Me.Label1(j - 2).Left + Me.Label1(j - 2).Width
        Me.txtField(j - 1).Left = Me.Label1(j - 2).Left + Me.Label1(j - 2).Width
      End If
    Next
    Me.grdCompetitors.Width = Me.Label1(rs.Fields.Count - 2).Left + Me.Label1(rs.Fields.Count - 2).Width + 240
    Me.btnClose.Left = Me.grdCompetitors.Width - Me.btnClose.Width - 260
    Me.btnNew.Left = Me.grdCompetitors.Width - Me.btnNew.Width - 240
    Me.btnDelete.Left = Me.btnClose.Left - Me.btnDelete.Width - 20
    Me.btnSave.Left = Me.btnNew.Left - Me.btnSave.Width - 20
    Me.lblTitle.Width = Me.grdCompetitors.Width
    Me.Width = Me.grdCompetitors.Width + 240
    .Redraw = True
  End With
 ' Me.Visible = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'Clean-up procedure
    If Not cn Is Nothing Then
        'first, check if the state is open, if yes then close it
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        'set them to nothing
        Set cn = Nothing
    End If
    'same with rs
    If Not rs Is Nothing Then
        If (rs.State And adStateOpen) = adStateOpen Then
            rs.Close
        End If
        Set rs = Nothing
    End If

End Sub

Private Sub grdCompetitors_RowColChange()
  Dim i As Integer
  rs.MoveFirst
  currentID = val(Me.grdCompetitors.TextMatrix(Me.grdCompetitors.Row, 0))
  rs.Find "id = " & currentID
  With rs
    If Not .EOF Then
      For i = 1 To .Fields.Count - 1
        Me.txtField(i - 1) = nz(.Fields(i), "")
      Next
    End If
  End With
  editMode = True
End Sub
