VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "SIMPHT"
   ClientHeight    =   3135
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim excelPath As String
Public Function importExcel() As ADODB.Recordset

    Dim dbStruk As ADODB.Connection
    Set dbStruk = New ADODB.Connection
    Dim rsStrukUmum As ADODB.Recordset
    
    dbStruk.CursorLocation = adUseClient
    dbStruk.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & excelPath & ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=0';"
    Set rsStrukUmum = New ADODB.Recordset
    rsStrukUmum.Open "select * from [sheet1$]", dbStruk, adOpenDynamic, adLockOptimistic
    Set importExcel = rsStrukUmum
    
End Function
Public Function konekAccess() As ADODB.Recordset

    Dim conConnection As New ADODB.Connection
    Dim cmdCommand As New ADODB.Command
    Dim rstRecordSet As New ADODB.Recordset
    
    With rstRecordSet
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
    End With
    
    conConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "gian.mdb;Mode=Read|Write"
    conConnection.Open
    rstRecordSet.Open "Select * from Table1", conConnection, adOpenStatic, adLockOptimistic
    Set konekAccess = rstRecordSet
    
End Function
Private Sub mnExit_Click()
End
End Sub

Private Sub mnOpen_Click()
CommonDialog1.Filter = "Excel 2003 (*.xls)|*.xls|Excel 2007 (*.xlsx)|*.xlsx"
CommonDialog1.ShowOpen
    excelPath = CommonDialog1.FileName
    'MsgBox CommonDialog1.FileName
    'If Right(CommonDialog1.FileName, 15) = "Data\Master.mdb" Then
    'Open App.Path & "\PATH" For Output As #1
    'Write #1, Left(CommonDialog1.FileName, Len(CommonDialog1.FileName) - 15)
    'Close #1
    'MsgBox "Silahkan jalankan ulang kembali!"
    'End
    'GoTo mulai
x = MsgBox("Apakah Anda akan mengambil data nominatif dari file " & exelpath, vbYesNo, "Konfirmasi Input Data Nominatif")
    If x = vbYes Then
        CommonDialog1.Filter = "Acces 2003 (*.mdb)|*.mdb"
        CommonDialog1.ShowSave
        FileCopy App.Path & "\master.mdb", CommonDialog1.FileName
    End If
        
End Sub

