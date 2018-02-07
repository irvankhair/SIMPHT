VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Transfer 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferring"
   ClientHeight    =   495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14160
   FillStyle       =   2  'Horizontal Line
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   495
   ScaleWidth      =   14160
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Align           =   1  'Align Top
      DragMode        =   1  'Automatic
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   873
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
End
Attribute VB_Name = "Transfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    ProgressBar1.Value = ProgressBar1.Min
End Sub
