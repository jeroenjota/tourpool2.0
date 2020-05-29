VERSION 5.00
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Tourpool 2.0"
   ClientHeight    =   9345
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10665
   LinkTopic       =   "frmMain"
   Picture         =   "frmMain.frx":0000
   ScrollBars      =   0   'False
   StartUpPosition =   3  'Windows Default
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

Private Sub MDIForm_Resize()
Dim client_rect As RECT
Dim client_hwnd As Long

    picStretched.Move 0, 0, _
        ScaleWidth, ScaleHeight

    ' Copy the original picture into picStretched.
    picStretched.PaintPicture _
        picOriginal.Picture, _
        0, 0, _
        picStretched.ScaleWidth, _
        picStretched.ScaleHeight, _
        0, 0, _
        picOriginal.ScaleWidth, _
        picOriginal.ScaleHeight
        
    ' Set the MDI form's picture.
    Picture = picStretched.Image

    ' Invalidate the picture.
    client_hwnd = FindWindowEx(Me.hWnd, 0, "MDIClient", _
        vbNullChar)
    GetClientRect client_hwnd, client_rect
    InvalidateRect client_hwnd, client_rect, 1
End Sub
