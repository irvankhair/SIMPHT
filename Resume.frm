VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ResumeNilai 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Resume"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15930
   HasDC           =   0   'False
   LinkTopic       =   "DaftarNominatif"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   15930
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog saveToExcelDialog 
      Left            =   4440
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton TransferKEExcell 
      Caption         =   "Transfer Ke Excell"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   24
      Top             =   2040
      Width           =   2535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      ItemData        =   "Resume.frx":0000
      Left            =   6960
      List            =   "Resume.frx":000D
      TabIndex        =   22
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Atur Tampilan Tabel"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   21
      Top             =   2040
      Width           =   2535
   End
   Begin VB.PictureBox Tampilan 
      BackColor       =   &H00E0E0E0&
      Height          =   7155
      Left            =   8760
      ScaleHeight     =   7095
      ScaleWidth      =   7635
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   7695
      Begin VB.ListBox ListPeralihan 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4590
         Left            =   2280
         MultiSelect     =   1  'Simple
         TabIndex        =   18
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         Height          =   375
         Left            =   5640
         Picture         =   "Resume.frx":0025
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Height          =   375
         Left            =   5640
         Picture         =   "Resume.frx":0367
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Height          =   375
         Left            =   5640
         Picture         =   "Resume.frx":06A9
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Height          =   375
         Left            =   5640
         Picture         =   "Resume.frx":09EB
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Batal"
         Height          =   375
         Left            =   3390
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   3360
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2400
         Width           =   855
      End
      Begin VB.ListBox List5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6330
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   11
         Top             =   360
         Width           =   3135
      End
      Begin VB.CommandButton Command14 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   3600
         Picture         =   "Resume.frx":0D2D
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   4560
         Width           =   420
      End
      Begin VB.CommandButton Command13 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   3600
         Picture         =   "Resume.frx":142F
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3960
         Width           =   420
      End
      Begin VB.CommandButton Command12 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   3600
         Picture         =   "Resume.frx":1B31
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   420
      End
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Height          =   390
         Left            =   3600
         Picture         =   "Resume.frx":2233
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1200
         Width           =   420
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         DragIcon        =   "Resume.frx":2937
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6030
         ItemData        =   "Resume.frx":9189
         Left            =   4440
         List            =   "Resume.frx":918B
         MultiSelect     =   1  'Simple
         TabIndex        =   6
         ToolTipText     =   "Geser data dengan drag  dan dropp"
         Top             =   360
         Width           =   3135
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Daftar Kolom Tersedia"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   600
         TabIndex        =   20
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Urutan Kolom Yang Tampil"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   0
         Width           =   2655
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10560
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Simpan Ukuran Kolom"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11040
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Transfer 
      Caption         =   "Transfer Ke Excell"
      Height          =   495
      Left            =   11040
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Refresh 
      Caption         =   "Refress"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid grdNominatif 
      Height          =   5175
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9128
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   23
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
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
         MarqueeStyle    =   4
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "RESUME"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   6255
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      DrawMode        =   9  'Not Mask Pen
      X1              =   0
      X2              =   3240
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "teamInformatikaAmalia2018"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   6255
   End
End
Attribute VB_Name = "ResumeNilai"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function ShellEx Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As Any, ByVal lpDirectory As Any, ByVal nShowCmd As Long) As Long

Private Declare Function GetTempPath Lib "kernel32" _
                                     Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                                           ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName Lib "kernel32" _
                                         Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                                                   ByVal lpPrefixString As String, ByVal wUnique As Long, _
                                                                   ByVal lpTempFileName As String) As Long

'untuk mouse list
Private mintDragIndex As Integer

Dim posisi As String
Private Declare Function SendMessage Lib "user32" Alias _
                                     "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
                                                     ByVal wParam As Long, lParam As Long) As Long
Private Const LB_SETCURSEL = &H186
Private Const LB_GETCURSEL = &H188
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" _
                                        (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function LBItemFromPt Lib "COMCTL32.DLL" _
                                      (ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, _
                                       ByVal bAutoScroll As Long) As Long


Private Function RapihkanGrid()
'    MsgBox grdNominatif.Columns(1).Caption
    grdNominatif.Columns("id").Visible = False
    grdNominatif.Columns("urutid").Visible = False
    grdNominatif.Columns("idNIB").Visible = False

    Dim i As Integer
    Dim db As ADODB.Connection

    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
    Set rsKolom = New ADODB.Recordset

    rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Resume' ", db, adOpenDynamic, adLockOptimistic


    For i = 0 To grdNominatif.Columns.Count - 1
        rsKolom.MoveFirst
        rsKolom.Find "isi='" & grdNominatif.Columns(i).Caption & "'"
        If Not rsKolom.EOF Then
            grdNominatif.Columns(i).Width = rsKolom![lebar kolom]

        End If
    Next i
End Function

Private Sub Command12_Click()
    Dim i As Integer
    For i = 0 To List5.ListCount - 1
        List6.AddItem List5.List(i)
    Next i
    'bulanLap = Format(CDate(List6.List(0)), "mmmm yyyy")
    List5.Clear
End Sub

Private Sub Command13_Click()
    Dim i As Integer
    'If Not List6 = "Nama" Then
pertama:
    For i = 0 To List6.ListCount - 1
        If List6.Selected(i) Then
            List5.AddItem List6.List(i)
            List6.RemoveItem (i)
            GoTo pertama

        End If

    Next i
    'End If
End Sub

Private Sub Command14_Click()
    Dim i As Integer
    For i = 0 To List6.ListCount - 1
        List5.AddItem List6.List(i)
    Next i
    List6.Clear
    'List6.AddItem "Nama Obat"
End Sub

Private Sub Command18_Click()
    Dim i As Integer
    Dim db As ADODB.Connection
    Dim Sumber As String
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
    Set rsKolom = New ADODB.Recordset
mulai:
    rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Resume' order by [indeks kolom]", db, adOpenDynamic, adLockOptimistic
    If rsKolom.RecordCount < List5.ListCount + List6.ListCount Then
        For i = 1 To List5.ListCount + List6.ListCount - rsKolom.RecordCount
            rsKolom.AddNew
            rsKolom![nama tabel] = "Daftar Resume"
            rsKolom![indeks kolom] = rsKolom.RecordCount
            rsKolom![lebar kolom] = 1170
            'rskolom!isi=
            rsKolom.Update

        Next i
        rsKolom.Requery
        rsKolom.MoveFirst
    End If
    'List5.Clear
    'List6.Clear

    '
    For i = 0 To List6.ListCount - 1
        rsKolom!isi = List6.List(i)
        rsKolom!tampil = True
        rsKolom.MoveNext
    Next i
    If Not List5.ListCount = 0 Then
        For i = 0 To List5.ListCount - 1
            rsKolom!isi = List5.List(i)
            rsKolom!tampil = False
            rsKolom.MoveNext
        Next i
    End If
    rsKolom.MoveFirst
    Sumber = ""
    For i = 0 To List6.ListCount - 2
        Sumber = Sumber & "[" & List6.List(i) & "],"
    Next i
    Sumber = Sumber & "[" & List6.List(List6.ListCount - 1) & "]"
    rsKolom!Source = "select " & Sumber & " from [Daftar Nominatif] order by urutid"
    rsKolom.Update
    '    rsKolom.Update
    Tampilan.Visible = False
    Set RSDN = New ADODB.Recordset
    'MsgBox rsKolom!Source

    Sumber = rsKolom!Source
    RSDN.Open Sumber, db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"

    Set grdNominatif.DataSource = RSDN
    grdNominatif.ReBind

    RapihkanGrid
    'While Not rsKolom.EOF
    '    X = rsKolom![indeks kolom]
    '    If rsKolom!tampil = True Then
    '        List6.AddItem rsdn.Fields(X).Name
    '    Else
    '        List5.AddItem rsdn.Fields(X).Name
    '    End If
    'rsKolom.MoveNext
    'Wend
    'Tampilan.Visible = True
    'rsKolom.Requery
    'RapihkanGrid

    Exit Sub
Adaeror:
    InputBox "t", "t", Err.Number
    If Err.Number = 3021 Then
        rsKolom.AddNew
        rsKolom![nama tabel] = "Daftar Nominatif"
        rsKolom![indeks kolom] = rsKolom.RecordCount
        rsKolom![lebar kolom] = 1170
        'rskolom!isi=
        rsKolom.Update
        rsKolom.Close
        GoTo mulai
    End If

End Sub

Private Sub Command19_Click()
    Dim i As Integer
pertama:
    For i = 0 To List5.ListCount - 1
        If List5.Selected(i) Then
            List6.AddItem List5.List(i)
            List5.RemoveItem (i)
            GoTo pertama

        End If

    Next i
End Sub

Private Sub Command2_Click()
    Dim rsSemua As ADODB.Recordset
    Dim rsKolom As ADODB.Recordset
    Dim db As ADODB.Connection
    Set db = New ADODB.Connection
    db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database "    ';pwd=globalisasi"
    db.CursorLocation = adUseClient
    Set rsSemua = New ADODB.Recordset
    rsSemua.Open "select * from [Daftar Nominatif] ", db, adOpenDynamic, adLockOptimistic

    Set rsKolom = New ADODB.Recordset
    rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Nominatif'", db, adOpenDynamic, adLockOptimistic
    While Not rsKolom.EOF
        rsKolom.Delete
        rsKolom.MoveNext
    Wend
    For i = 0 To rsSemua.Fields.Count - 1
        rsKolom.AddNew
        rsKolom![nama tabel] = "Daftar Resume"
        rsKolom![indeks kolom] = i
        rsKolom!isi = rsSemua.Fields(i).Name
        rsKolom![lebar kolom] = 1000    'rsSemua.Fields(i).ActualSize
        rsKolom!tipe = rsSemua.Fields(i).Type
        rsKolom.Update
    Next i



End Sub

Private Sub Command3_Click()
    Dim i As Integer
    Dim db As ADODB.Connection
    'Tampilan.Top = txtNama.Top
    'Tampilan.Left = txtNama.Left
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
    Set rsKolom = New ADODB.Recordset
    rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Resume'", db, adOpenDynamic, adLockOptimistic
    List5.Clear
    List6.Clear
    While Not rsKolom.EOF
        'X = rsKolom![indeks kolom]
        If rsKolom!tampil = True Then
            List6.AddItem "" & rsKolom!isi
        Else
            List5.AddItem rsKolom!isi
        End If
        rsKolom.MoveNext
    Wend
    Tampilan.Visible = True
End Sub

Private Sub Command4_Click()
    Tampilan.Visible = False

End Sub

Private Sub EditIsi_Click()
    If Not RSDN.EOF Then
        grdNominatif.Splits(0).MarqueeStyle = dbgFloatingEditor
        grdNominatif.Splits(0).Locked = False
    End If
End Sub
Public Sub pindah(ByVal lbhwnd As Long, _
                  ByVal X As Single, ByVal Y As Single)
'untuk select otomatis saat meuse ke list

    Dim ItemIndex As Long
    Dim AtThisPoint As POINTAPI
    AtThisPoint.X = X \ Screen.TwipsPerPixelX
    AtThisPoint.Y = Y \ Screen.TwipsPerPixelY
    Call ClientToScreen(lbhwnd, AtThisPoint)
    ItemIndex = LBItemFromPt(lbhwnd, AtThisPoint.X, _
                             AtThisPoint.Y, False)
    If ItemIndex <> SendMessage(lbhwnd, LB_GETCURSEL, 0, 0) Then
        Call SendMessage(lbhwnd, LB_SETCURSEL, ItemIndex, 0)
    End If

End Sub

Private Sub Refresh_Click()
    RSDN.Requery
    RapihkanGrid
End Sub

Private Sub TransferKEExcell_Click()
    saveToExcelDialog.Filter = "Excel 2003 (*.xlsx)|*.xlsx"
    saveToExcelDialog.ShowSave
    FileCopy App.Path & "\Temp.xlsx", saveToExcelDialog.FileName
    exportExcel = saveToExcelDialog.FileName

    Dim excelApp As Excel.Application
    Dim excelWB As Excel.Workbook
    Dim excelWS As Excel.Worksheet
    Set excelApp = CreateObject("Excel.Application")
    Dim i As Integer
    Dim j As Integer

    Dim conConnection As New ADODB.Connection
    Dim cmdCommand As New ADODB.Command
    Dim rstRecordSet As New ADODB.Recordset

    With rstRecordSet
        .CursorType = adOpenStatic
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
    End With

    conConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";"
    conConnection.Open
    rstRecordSet.Open "Select * from [daftar nominatif]", conConnection, adOpenStatic, adLockOptimistic

    excelApp.Visible = False
    excelApp.Workbooks.Open exportExcel
    Set excelWB = excelApp.Workbooks(1)
    Set excelWS = excelWB.Worksheets(1)
    'MsgBox RSDN.Fields("nomor urut")
    excelWS.Activate
    'excelWS.Cells(16, 2) = 3
    excelWS.Cells(2, 2) = NamaProjek
    'MsgBox excelWS.Cells(16, 2)
    For i = 16 To rstRecordSet.RecordCount
        excelWS.Cells(i, 2) = rstRecordSet.Fields("nomor urut")
        excelWS.Cells(i, 3) = rstRecordSet.Fields("Index Identitas")
        excelWS.Cells(i, 4) = rstRecordSet.Fields("Identitas")
        excelWS.Cells(i, 7) = rstRecordSet.Fields("NIB")
        excelWS.Cells(i, 10) = rstRecordSet.Fields("Luas Hasil Ukur di Dalam Trase")
        'excelWS.Cells(i, 12) = rstRecordSet.Fields("Surat Tanda Bukti")
        'excelWS.Cells(i, 13) = rstRecordSet.Fields("Status Ruang Atas Bawah")
        'excelWS.Cells(i, 14) = rstRecordSet.Fields("Luas Atas Bawah")
        ' For j = 12 To 35
        '    excelWS.Cells(i, j + 3) = rstRecordSet.Fields(j)
        'Next
        excelWS.Cells(i, 15) = rstRecordSet.Fields("Indexs Jenis Bangunan")
        excelWS.Cells(i, 16) = rstRecordSet.Fields("jenis bangunan")
        excelWS.Cells(i, 17) = rstRecordSet.Fields("Jumlah Jenis Bangunan")
        excelWS.Cells(i, 18) = rstRecordSet.Fields("Luas bangunan")
        excelWS.Cells(i, 19) = rstRecordSet.Fields("index tanaman")
        excelWS.Cells(i, 20) = rstRecordSet.Fields("Jenis Musim Tanaman")
        excelWS.Cells(i, 21) = rstRecordSet.Fields("Jumlah Jenis Musim Tanaman")
        excelWS.Cells(i, 22) = rstRecordSet.Fields("nomor tanaman")
        excelWS.Cells(i, 23) = rstRecordSet.Fields("jenis tanaman")
        excelWS.Cells(i, 24) = rstRecordSet.Fields("Ukuran Jenis Tanaman")
        excelWS.Cells(i, 25) = rstRecordSet.Fields("Jumlah tanaman")
        excelWS.Cells(i, 26) = rstRecordSet.Fields("Index Benda Lain yang Berkaitan")
        excelWS.Cells(i, 27) = rstRecordSet.Fields("Jenis Benda Lain yang Berkaitan")
        excelWS.Cells(i, 28) = rstRecordSet.Fields("Jumlah Benda Lain yang Berkaitan")

        excelWS.Cells(i, 33) = rstRecordSet.Fields("Nilai Tanah per Meter Persegi")
        excelWS.Cells(i, 34) = rstRecordSet.Fields("Nilai Pasar Tanah")
        excelWS.Cells(i, 35) = rstRecordSet.Fields("Nilai Bangunan per Meter Persegi")
        excelWS.Cells(i, 36) = rstRecordSet.Fields("Jumlah Nilai Bangunan")
        excelWS.Cells(i, 37) = rstRecordSet.Fields("Nilai Tanaman per Meter Persegi")
        excelWS.Cells(i, 38) = rstRecordSet.Fields("Jumlah Nilai Tanaman")
        excelWS.Cells(i, 39) = rstRecordSet.Fields("Nilai Pasar Tanaman")
        excelWS.Cells(i, 40) = rstRecordSet.Fields("Total Nilai Fisik")
        excelWS.Cells(i, 41) = rstRecordSet.Fields("Kerugian Usaha")
        excelWS.Cells(i, 42) = rstRecordSet.Fields("Solatium")
        excelWS.Cells(i, 43) = rstRecordSet.Fields("Pindah")
        excelWS.Cells(i, 44) = rstRecordSet.Fields("Pajak")
        excelWS.Cells(i, 45) = rstRecordSet.Fields("Masa Tunggu")
        excelWS.Cells(i, 46) = rstRecordSet.Fields("Total Nilai Non Fisik")
        excelWS.Cells(i, 47) = rstRecordSet.Fields("Grand Total Penggantian Wajar")
        rstRecordSet.MoveNext
    Next
    excelWB.Save
    rstRecordSet.Close
    conConnection.Close
    excelWB.Close
    excelApp.Quit
    Set excelApp = Nothing
    Set excelWB = Nothing
    Set excelWS = Nothing
    MsgBox "Selsai"



End Sub

Private Function CreateTempFile(sPrefix As String) As String
    Dim sTmpPath As String * 512
    Dim sTmpName As String * 576
    Dim nRet As Long

    nRet = GetTempPath(512, sTmpPath)
    If (nRet > 0 And nRet < 512) Then
        nRet = GetTempFileName(sTmpPath, sPrefix, 0, sTmpName)
        If nRet <> 0 Then
            CreateTempFile = Left$(sTmpName, _
                                   InStr(sTmpName, vbNullChar) - 1)
        End If
    End If
End Function
Private Sub Form_Click()
'List1.Visible = False
End Sub

Private Sub Form_Load()

    Dim i As Integer
    Dim db As ADODB.Connection
    Dim rsKolom As ADODB.Recordset
    Dim Sumber As String
    Set db = New ADODB.Connection
    Set rsKolom = New ADODB.Recordset
    Set RSDN = New ADODB.Recordset

    If pROJECTPATH = "" Then
        MsgBox "Maaf belum ada file projeck yang dipilih......!", vbInformation
    Else
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database "    ';pwd=globalisasi"
        db.CursorLocation = adUseClient

        rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Resume' ", db, adOpenDynamic, adLockOptimistic
        Sumber = rsKolom!Source
        'RSDN.Open "SELECT * FROM [DAFTAR NOMINATIF] order by UrutId", db, adOpenDynamic, adLockOptimistic
        RSDN.Open Sumber, db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"
        'MsgBox Sumber
        Set grdNominatif.DataSource = RSDN
        grdNominatif.ReBind


        RapihkanGrid

        Set MyProperty = grdNominatif   'nama datagrid yang inigin di scroll dengan mouse
        WheelHook grdNominatif
    End If

End Sub

Private Sub Form_Resize()
    If Not Me.ScaleHeight = 0 Then
        Tampilan.Left = 0
        Tampilan.Top = Line1.Y1 + 30
        Label1.Width = Me.Width
        Line1.X2 = Me.ScaleWidth
        grdNominatif.Top = Line1.Y1 + 30
        grdNominatif.Width = Me.ScaleWidth
        grdNominatif.Height = Me.ScaleHeight - Line1.Y1 - 30
        grdNominatif.Left = 0
    End If
End Sub


Private Sub grdNominatif_Click()
'List1.Visible = False
End Sub

Private Sub grdNominatif_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    Dim db As ADODB.Connection
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False"
    db.CursorLocation = adUseClient
    Set rsKolom = New ADODB.Recordset
    rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Resume'", db, adOpenDynamic, adLockOptimistic
    rsKolom.Find "isi='" & grdNominatif.Columns(ColIndex).Caption & "'"
    If Not rsKolom.EOF Then
        rsKolom![lebar kolom] = grdNominatif.Columns(ColIndex).Width
        'rskolom!tipe = rsdn.Fields(0).Type
        rsKolom.Update
    End If
End Sub

Private Sub grdNominatif_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not RSDN.EOF And Not RSDN.BOF Then
            NIBTerpilih = "" & RSDN!nib
            'PopupMenu MainForm.mnEditBaris
            List1.Visible = True
            List1.Top = Y + grdNominatif.Top
            List1.Left = X

        End If
    End If
End Sub

Private Sub grdNominatif_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If Not LastRow = "" Then
        'MsgBox LastRow
        grdNominatif.Splits(0).Locked = True
        grdNominatif.Splits(0).MarqueeStyle = dbgHighlightRowRaiseCell

    End If
    List1.Visible = False

End Sub

Private Sub HapusBaris_Click()
    If Not RSDN.EOF And Not RSDN.BOF Then
        X = MsgBox("Yakin akan menghapus kompionen bidang pada NIB : " & NIBTerpilih, vbYesNo, "Konfirmasi Hapus Baris Pada Bidang")
        If X = vbYes Then
            RSDN.Delete
            RSDN.MoveNext
        End If
    End If

End Sub

Private Sub List1_Click()
    If List1.text = "Edit" Then
        EditIsi_Click
        List1.Visible = False
    ElseIf List1.text = "Sisip" Then
        Sisip_Click
        List1.Visible = False
    ElseIf List1.text = "Hapus" Then
        HapusBaris_Click
        List1.Visible = False
    End If

End Sub

Private Sub List1_LostFocus()
    List1.Visible = False
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pindah List1.hwnd, X, Y
End Sub
'untuk dragdropp list -----------------------------------------------------------------------------
Public Function ListRowCalc(pobjLB As ListBox, ByVal Y As Single) As Integer    '-----------------------------------------------------------------------------

    Const LB_GETITEMHEIGHT = &H1A1

    Dim intItemHeight As Integer
    Dim intRow As Integer

    intItemHeight = SendMessage(pobjLB.hwnd, LB_GETITEMHEIGHT, 0, 0)

    intRow = ((Y / Screen.TwipsPerPixelY) \ intItemHeight) + pobjLB.TopIndex

    If intRow < pobjLB.ListCount - 1 Then
        ListRowCalc = intRow
    Else
        ListRowCalc = pobjLB.ListCount - 1
    End If

End Function
'untuk dragdropp list-----------------------------------------------------------------------------
Public Sub ListRowMove(pobjLB As ListBox, _
                       ByVal pintOldRow As Integer, _
                       ByVal pintNewRow As Integer)
'-----------------------------------------------------------------------------

    Dim strSavedItem As String
    Dim intX As Integer

    If pintOldRow = pintNewRow Then Exit Sub

    strSavedItem = pobjLB.List(pintOldRow)

    If pintOldRow > pintNewRow Then
        For intX = pintOldRow To pintNewRow + 1 Step -1
            pobjLB.List(intX) = pobjLB.List(intX - 1)
        Next intX
    Else
        For intX = pintOldRow To pintNewRow - 1
            pobjLB.List(intX) = pobjLB.List(intX + 1)
        Next intX
    End If

    pobjLB.List(pintNewRow) = strSavedItem

End Sub
Private Sub List6_DragDrop(Source As Control, X As Single, Y As Single)
    ListRowMove Source, mintDragIndex, ListRowCalc(Source, Y)

End Sub

Private Sub List6_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    List6.ListIndex = ListRowCalc(List6, Y)

End Sub
Private Sub List6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mintDragIndex = ListRowCalc(List6, Y)
    List6.Drag
End Sub



Private Sub Sisip_Click()
    Dim urutSebelum As String
    Dim nAMApEMILIK As String
    If Not RSDN.EOF = True And Not RSDN.BOF = True Then
        posisi = RSDN.AbsolutePosition

        RSDN.MovePrevious
        If RSDN.BOF Then
            RSDN.Move 1
            urutSebelum = RSDN!urutid - 0.2
        Else
            urutSebelum = RSDN!urutid
        End If
        nAMApEMILIK = "" & RSDN!pemilik

        RSDN.AddNew
        RSDN!urutid = urutSebelum + 0.1
        RSDN!nib = NIBTerpilih
        RSDN!pemilik = nAMApEMILIK
        RSDN.Update
        RSDN.Requery
        RSDN.Move posisi - 1

        RapihkanGrid
    End If
End Sub
'=======
Private Sub Command1_Click()
'MsgBox "okke"
'    Dim excel As New ADODB.Recordset
'    Set excel = importExcel
'    MsgBox excel.RecordCount
'Command2_Click
    RSDN.Requery

End Sub

Public Function importExcel() As ADODB.Recordset

    Dim dbStruk As ADODB.Connection
    Set dbStruk = New ADODB.Connection
    Dim rsStrukUmum As ADODB.Recordset

    dbStruk.CursorLocation = adUseClient
    dbStruk.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "EDit sumber baku.xls;Extended Properties='Excel 8.0;HDR=Yes;IMEX=0';"
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

    conConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "Master.mdb;Mode=Read|Write"
    conConnection.Open
    rstRecordSet.Open "Select * from Daftar_Nominatif", conConnection, adOpenStatic, adLockOptimistic
    Set konekAccess = rstRecordSet

End Function

'>>>>>>> d3d07549d5308c63fd676dc3c8bf461ab7c97286

Private Sub Transfer_Click()


    Dim Excel As New ADODB.Recordset
    Dim access As New ADODB.Recordset

    Dim i, j As Integer
    Set Excel = importExcel
    Set access = konekAccess

    Dim temp As String

    temp = Excel.Fields("NIB")
    For i = 1 To Excel.RecordCount
        With access
            .AddNew
            For j = 2 To 19
                .Fields(j) = Excel.Fields(j - 2)

            Next

            If ((Excel.Fields(17) <> 0) And (Excel.Fields(18) <> 0) And (Excel.Fields(19) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = Excel.Fields(17)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If


                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = Excel.Fields(18)
                .Fields("jenis tanaman") = Excel.Fields(16)
                .MoveNext
                .AddNew
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = Excel.Fields(19)
                .Fields("jenis tanaman") = Excel.Fields(16)
            ElseIf ((Excel.Fields(17) <> 0) And (Excel.Fields(18) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = Excel.Fields(17)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")

                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = Excel.Fields(18)
                .Fields("jenis tanaman") = Excel.Fields(16)
            ElseIf ((Excel.Fields(18) <> 0) And (Excel.Fields(19) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = Excel.Fields(18)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = Excel.Fields(19)
                .Fields("jenis tanaman") = Excel.Fields(16)
            ElseIf ((Excel.Fields(17) <> 0) And (Excel.Fields(19) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = Excel.Fields(17)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = Excel.Fields(19)
                .Fields("jenis tanaman") = Excel.Fields(16)
            ElseIf (Excel.Fields(17) <> 0) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = Excel.Fields(17)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

            ElseIf (Excel.Fields(18) <> 0) Then
                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = Excel.Fields(18)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

            ElseIf (Excel.Fields(19) <> 0) Then
                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = Excel.Fields(19)
                .Fields("jenis tanaman") = Excel.Fields(16)
                If IsNull(access.Fields("idNIB")) Then
                    access.Fields("idNIB") = temp
                ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                    temp = access.Fields("idNIB")
                End If
                If (Not IsNull(access.Fields(2))) Then
                    access.Fields("NIB") = access.Fields("idNIB")
                End If

            End If
        End With
        If IsNull(access.Fields("idNIB")) Then
            access.Fields("idNIB") = temp
        ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
            temp = access.Fields("idNIB")

        End If

        'Call cekNIB(temp, access.Fields("idNIB"))
        If (Not IsNull(access.Fields(1))) Then
            access.Fields("NIB") = access.Fields("idNIB")
        End If

        access.Update

        Excel.MoveNext
        access.MoveNext
    Next
    access.MoveFirst

    For i = 1 To access.RecordCount

        access.Fields(1) = access.Fields(0)
        access.Update
        access.MoveNext
    Next
    MsgBox "Transfered"
    Excel.Close
    access.Close
End Sub
