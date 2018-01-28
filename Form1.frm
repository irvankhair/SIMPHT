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
   Begin VB.CommandButton Transfer 
      Caption         =   "Transfer"
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   1320
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
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9128
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


Private Sub Transfer_Click()

    Dim excel As New ADODB.Recordset
    Dim access As New ADODB.Recordset
    
    Dim i, j As Integer
    Set excel = importExcel
    Set access = konekAccess
    
    Dim temp As String
   
    temp = excel.Fields("nis")
    For i = 1 To 602
        With access
            .AddNew
            For j = 1 To 21
                .Fields(j) = excel.Fields(j - 1)
            Next
        End With
        If IsNull(access.Fields("nis")) Then
            access.Fields("nis") = temp
        ElseIf (Not IsNull(access.Fields("nis")) And (access.Fields("nis") <> temp)) Then
            temp = access.Fields("nis")
        End If
        access.Update
        excel.MoveNext
        access.MoveNext
    Next

    MsgBox "Transfered"
    excel.Close
    access.Close
End Sub
