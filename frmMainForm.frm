VERSION 5.00
Object = "{0DF5D14C-08DD-4806-8BE2-B59CB924CFC9}#1.7#0"; "VBCCR16.OCX"
Begin VB.Form frmMain 
   Caption         =   "Tour Pool 2.0"
   ClientHeight    =   6690
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12435
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VBCCR16.LabelW lblTitle 
      Height          =   2175
      Left            =   -120
      TabIndex        =   1
      Top             =   1560
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Cambria"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Alignment       =   2
      BackStyle       =   0
      Caption         =   "Tourpool"
   End
   Begin VBCCR16.LabelW lblCopyright 
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   6360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Garamond"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   1
      BackStyle       =   0
      Caption         =   "jota services"
   End
   Begin VB.Image bgImage 
      Height          =   6495
      Left            =   0
      Picture         =   "frmMainForm.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8175
   End
   Begin VB.Menu mnuBestand 
      Caption         =   "&Bestand"
      Begin VB.Menu mnuRondes 
         Caption         =   "&Rondes"
      End
      Begin VB.Menu mnuPools 
         Caption         =   "&Pools"
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Afdrukken"
      End
      Begin VB.Menu mnuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSluiten 
         Caption         =   "Af&sluiten"
      End
   End
   Begin VB.Menu muDeelnemers 
      Caption         =   "&Deelnemers"
      Begin VB.Menu mnuAdressen 
         Caption         =   "&Adreslijst"
      End
      Begin VB.Menu mnuDeelnemersInvoeren 
         Caption         =   "&Pool deelnemers"
      End
   End
   Begin VB.Menu mnuRonde 
      Caption         =   "&Ronde"
      Begin VB.Menu mnuEtappes 
         Caption         =   "&Etappes"
      End
      Begin VB.Menu mnuRitUitslagen 
         Caption         =   "&Rit uitslagen"
      End
   End
   Begin VB.Menu mnuOverdeApp 
      Caption         =   "&Over deze app"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cn As ADODB.Connection

Private Sub Form_Load()
  Set cn = New ADODB.Connection
  With cn
    .ConnectionString = lclConn
    .Open
  End With
  
  UnifyForm Me
  centerForm Me
  
End Sub

Sub initform()
  Dim picAspectRatio As Double
  LockWindowUpdate Me.hwnd
  With Me.bgImage
    .Visible = True
    .Picture = LoadPicture(App.Path & "\mainpic05.jpg")
    picAspectRatio = .Picture.Height / .Picture.Width
    If Me.Height < 9000 Then Me.Height = 9000
    Me.Width = Me.Height / picAspectRatio
    
    .Top = 0
    .Width = Me.ScaleWidth
    .Height = Me.ScaleHeight
  End With
  
  LockWindowUpdate 0
  Me.lblTitle.Top = (Me.ScaleHeight - Me.lblTitle.Height) / 5 * 2
  Me.lblTitle.Left = 0
  Me.lblTitle.Width = Me.ScaleWidth
  Me.lblTitle.Font.Size = 36
  Me.lblTitle.Caption = getOrganisation(cn) & vbNewLine & "Tourpool"
  Me.lblTitle.ZOrder
  Me.lblCopyright = "© 2004 - " & Year(Now) & " jota services"
  Me.lblCopyright.AutoSize = True
  Me.lblCopyright.Left = Me.ScaleWidth - Me.lblCopyright.Width - 150
  Me.lblCopyright.Top = Me.ScaleHeight - Me.lblCopyright.Height - 100
  Me.lblCopyright.ZOrder
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim objForm As Form
    
    If Not cn Is Nothing Then
        If (cn.State And adStateOpen) = adStateOpen Then
            cn.Close
        End If
        Set cn = Nothing
    End If
    
    For Each objForm In Forms
        If objForm.Name <> Me.Name Then
            Unload objForm
            Set objForm = Nothing
        End If
    Next
    write2Log "App ended", True
End Sub

Private Sub Form_Resize()
  initform
End Sub

Private Sub mnuAdressen_Click()
frmCompetitors.Show 1
End Sub

Private Sub mnuDeelnemersInvoeren_Click()
  frmDeelnemerPloegen.Show 1
End Sub
