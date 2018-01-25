VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnEksport 
         Caption         =   "&Eksport"
      End
      Begin VB.Menu mnPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnInsert 
      Caption         =   "&Insert"
      Begin VB.Menu mnBacaDataNominatif 
         Caption         =   "&Baca Data Nominatif"
      End
   End
   Begin VB.Menu mnEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnDataFisik 
         Caption         =   "&Data Fisik"
      End
      Begin VB.Menu mnDataNonfisik 
         Caption         =   "&Data Nonfisik"
      End
   End
   Begin VB.Menu mnSetting 
      Caption         =   "&Setting"
   End
   Begin VB.Menu mnReport 
      Caption         =   "&Report"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnExit_Click()
End
End Sub

