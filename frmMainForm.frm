VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Tour Pool 2.0"
   ClientHeight    =   6060
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6060
   ScaleWidth      =   11880
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

Private Sub mnuAdressen_Click()
frmCompetitors.Show 1
End Sub

