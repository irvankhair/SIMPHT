VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   15930
   Begin VB.CommandButton Command2 
      Caption         =   "New Button"
      Height          =   735
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   8280
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6375
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   14415
      _ExtentX        =   25426
      _ExtentY        =   11245
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "teamInformatikaAmalia2018"
      Height          =   375
      Left            =   8760
      TabIndex        =   0
      Top             =   7080
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "okke"
End Sub

Private Sub Command2_Click()
Dim dbStruk As ADODB.Connection
Set dbStruk = New ADODB.Connection
Dim rsStrukUmum As ADODB.Recordset
dbStruk.CursorLocation = adUseClient
dbStruk.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\EDit sumber baku.xls;Extended Properties='Excel 8.0;HDR=Yes;IMEX=0';"
            

Set rsStrukUmum = New ADODB.Recordset
rsStrukUmum.Open "select * from [sheet1$]", dbStruk, adOpenDynamic, adLockOptimistic
MsgBox rsStrukUmum.RecordCount
MsgBox rsStrukUmum.Fields(4).Name

End Sub
