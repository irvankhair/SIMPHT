VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Tanaman 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   14220
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Tampilan 
      BackColor       =   &H00E0E0E0&
      Height          =   7155
      Left            =   480
      ScaleHeight     =   7095
      ScaleWidth      =   5595
      TabIndex        =   12
      Top             =   2280
      Visible         =   0   'False
      Width           =   5655
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
         TabIndex        =   25
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Height          =   375
         Left            =   5640
         Picture         =   "Tanaman.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Height          =   375
         Left            =   5640
         Picture         =   "Tanaman.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Height          =   375
         Left            =   5640
         Picture         =   "Tanaman.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Height          =   375
         Left            =   5640
         Picture         =   "Tanaman.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Batal"
         Height          =   375
         Left            =   2430
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
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
         Sorted          =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command16 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2640
         Picture         =   "Tanaman.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   5040
         Width           =   420
      End
      Begin VB.CommandButton Command17 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2640
         Picture         =   "Tanaman.frx":140A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4440
         Width           =   420
      End
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2640
         Picture         =   "Tanaman.frx":1B0C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2280
         Width           =   420
      End
      Begin VB.CommandButton Command20 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Height          =   390
         Left            =   2640
         Picture         =   "Tanaman.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   420
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         DragIcon        =   "Tanaman.frx":2912
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
         ItemData        =   "Tanaman.frx":9164
         Left            =   3360
         List            =   "Tanaman.frx":9166
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   13
         ToolTipText     =   "Geser data dengan drag  dan dropp"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Daftar NIB Tersedia"
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
         Left            =   120
         TabIndex        =   27
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "NIB YANG TAMPIL"
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
         Left            =   3480
         TabIndex        =   26
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5055
      Left            =   120
      TabIndex        =   10
      Top             =   5160
      Width           =   17295
      Begin VB.CommandButton Command22 
         Caption         =   "&Terapkan Kalkulasi NPW Tanaman"
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
         Left            =   10080
         TabIndex        =   34
         Top             =   4440
         Width           =   4215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "&Edit"
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
         Left            =   1560
         TabIndex        =   33
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "&Bersihkan"
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
         TabIndex        =   32
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&Simpan"
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
         Left            =   3000
         TabIndex        =   31
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command8 
         Caption         =   "&Hapus"
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
         Left            =   5880
         TabIndex        =   30
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton Command9 
         Caption         =   "&Pilih NIB"
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
         TabIndex        =   29
         Top             =   4440
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid GRDTanaman 
         Height          =   3495
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   6165
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Penetapan Harga Tanaman"
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
         Height          =   615
         Left            =   3480
         TabIndex        =   11
         Top             =   120
         Width           =   6255
      End
   End
   Begin MSDataListLib.DataList List1 
      Height          =   3375
      Left            =   8760
      TabIndex        =   9
      Top             =   6120
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5953
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   11655
      Begin VB.CommandButton Command23 
         Caption         =   "&Tetapkan Harga"
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
         Left            =   9360
         TabIndex        =   36
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command21 
         Caption         =   "&Ambil Data"
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
         Left            =   240
         TabIndex        =   35
         Top             =   3600
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Edit"
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
         Left            =   2880
         TabIndex        =   5
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Bersihkan"
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
         Left            =   6120
         TabIndex        =   4
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "&Simpan"
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
         Left            =   4320
         TabIndex        =   3
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "&Hapus"
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
         Left            =   7560
         TabIndex        =   2
         Top             =   3600
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grdHarga 
         Height          =   2655
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4683
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   23
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Klasifikasi Jenis Tanaman"
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
         Left            =   3480
         TabIndex        =   7
         Top             =   120
         Width           =   6255
      End
   End
   Begin VB.CommandButton Command10 
      Caption         =   "&Input Harga Zona Tanah"
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
      Left            =   6720
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "PENILAIAN HARGA TANAMAN"
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
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "Tanaman"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHarga As ADODB.Recordset
Dim rstanaman As ADODB.Recordset
Private Declare Function SendMessage Lib "user32" Alias _
"SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
ByVal wParam As Long, lParam As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function ClientToScreen Lib "user32" _
(ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Private Declare Function LBItemFromPt Lib "COMCTL32.DLL" _
(ByVal hLB As Long, ByVal ptX As Long, ByVal ptY As Long, _
ByVal bAutoScroll As Long) As Long

Private Sub Command1_Click()
grdHarga.Splits(0).MarqueeStyle = dbgFloatingEditor
grdHarga.Splits(0).Locked = False
End Sub

Private Sub Command10_Click()
Frame1.Visible = True

End Sub

Private Sub Command15_Click()
Tampilan.Visible = False
End Sub

Private Sub Command16_Click()
Dim i As Integer
For i = 0 To List6.ListCount - 1
List5.AddItem List6.List(i)
Next i
List6.Clear
'List6.AddItem "Nama Obat"
End Sub

Private Sub Command17_Click()
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

Private Sub Command18_Click()
Dim rsDaftarNIB As ADODB.Recordset
 Dim i As Integer
 
    Dim db As ADODB.Connection
    'Tampilan.Top = txtNama.Top
    'Tampilan.Left = txtNama.Left
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False"
    db.CursorLocation = adUseClient
    'Set rstanaman = New ADODB.Recordset
    'rstanaman.Open "select * from [Zona Tanah] ", db, adOpenDynamic, adLockOptimistic
    Set rsDaftarNIB = New ADODB.Recordset
    rsDaftarNIB.Open "SELECT * from [daftar tanaman] order by urutid", db, adOpenDynamic, adLockOptimistic
        If Not rstanaman.RecordCount = 0 Then
            rstanaman.MoveFirst
        End If
        While Not rstanaman.EOF
            rstanaman.Delete
            rstanaman.MoveNext
        Wend
        
        For i = 0 To List6.ListCount - 1
            If Not rsDaftarNIB.RecordCount = 0 Then
                rsDaftarNIB.MoveFirst
            End If
'metode bangunan bukan find tapi filter karena bis banyak bangunan untuk satu nib
            rsDaftarNIB.Filter = "idNIB='" & List6.List(i) & "'"
            While Not rsDaftarNIB.EOF
                rstanaman.AddNew
                    rstanaman!idtanaman = rsDaftarNIB!ID
                    rstanaman!nomor = rsDaftarNIB![urutid]
                    rstanaman!nib = rsDaftarNIB!idnib
                    rstanaman!Pemilik = rsDaftarNIB!Pemilik
                    rstanaman![Jenis tanaman] = rsDaftarNIB![Jenis tanaman]
                    rstanaman![ukuran jenis tanaman] = rsDaftarNIB![ukuran jenis tanaman]
                    rstanaman![jumlah tanaman] = rsDaftarNIB![jumlah tanaman]
                    
                rstanaman.Update
                rsDaftarNIB.MoveNext
            Wend
        Next i
        Tampilan.Visible = False
        GRDTanaman.ReBind
End Sub

Private Sub Command19_Click()
Dim i As Integer
For i = 0 To List5.ListCount - 1
List6.AddItem List5.List(i)
Next i
'bulanLap = Format(CDate(List6.List(0)), "mmmm yyyy")
List5.Clear
End Sub

Private Sub Command20_Click()
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

Private Sub Command21_Click()
Dim rsDaftarTanaman As ADODB.Recordset
 Dim i As Integer
 
    Dim db As ADODB.Connection
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
    Set rsDaftarTanaman = New ADODB.Recordset
     rsDaftarTanaman.Open "SELECT DISTINCTROW [Daftar Nominatif].[jenis tanaman], [Daftar Nominatif].[Ukuran Jenis Tanaman], Sum([Daftar Nominatif].[Jumlah tanaman]) AS [Sum Of Jumlah tanaman], Avg([Daftar Nominatif].[Nilai Tanaman per Meter Persegi]) AS [Harga], Count(*) AS [Jumlah Tanaman] From [Daftar Nominatif] GROUP BY [Daftar Nominatif].[jenis tanaman], [Daftar Nominatif].[Ukuran Jenis Tanaman] HAVING ((([Daftar Nominatif].[jenis tanaman]) Is Not Null));", db, adOpenDynamic, adLockOptimistic
     
     While Not rsDaftarTanaman.EOF
        If Not rsHarga.RecordCount = 0 Then
        rsHarga.MoveFirst
        End If
        rsHarga.Filter = "[jenis tanaman]='" & rsDaftarTanaman![Jenis tanaman] & "' and [ukuran]='" & rsDaftarTanaman![ukuran jenis tanaman] & "'"
        'MsgBox "[jenis tanaman]='" & rsDaftarTanaman![jenis tanaman] & "' and ukuran = '" & rsDaftarTanaman![Ukuran Jenis Tanaman] & "'"
        If rsHarga.EOF Then
            rsHarga.AddNew
                rsHarga!harga = rsDaftarTanaman!harga
                rsHarga![Jenis tanaman] = rsDaftarTanaman![Jenis tanaman]
                rsHarga![ukuran] = rsDaftarTanaman![ukuran jenis tanaman]
            rsHarga.Update
        End If
        rsDaftarTanaman.MoveNext
    Wend
    rsHarga.Filter = ""
    rsHarga.Sort = "[jenis tanaman],ukuran"
    rsHarga.Requery
End Sub

Private Sub Command22_Click()
Dim JumlahItem As Single
Dim JumlahTanaman As Single
Dim NilaiSusut As Single
X = MsgBox("Apakah Anda yakin untuk melakukan proses kalkulasi NPW tanaman dengan harga pada klasifikasi yang telah ditetapkan?", vbYesNo, "Konfirmasi Kalkulasi NPW Tanah")
If X = vbYes Then
    rstanaman.Filter = "harga > 0"
    rstanaman.MoveFirst
    Dim db As ADODB.Connection
    Dim RSDN As ADODB.Recordset
    Dim rsKalkulasi As ADODB.Recordset
    Dim Sumber As String
    Set db = New ADODB.Connection
      Set RSDN = New ADODB.Recordset
      Set rsKalkulasi = New ADODB.Recordset
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database "    ';pwd=globalisasi"
        db.CursorLocation = adUseClient
   
        RSDN.Open "select id,idnib,[Jumlah Tanaman],[Nilai Tanaman per meter persegi],[jumlah nilai tanaman],[nilai pasar Tanaman] from [daftar nominatif] where [jenis tanaman] is not null ", db, adOpenDynamic, adLockOptimistic
        If Not RSDN.EOF Then
'kalkulasi tahap 1 jumlah tiap bangunan dengan penyusutan
            While Not rstanaman.EOF
                RSDN.MoveFirst
                RSDN.Find "id='" & rstanaman!idtanaman & "'"
                If Not RSDN.EOF Then
                    If IsNull(RSDN![jumlah tanaman]) Then
                        JumlahItem = 1
                    Else
                        JumlahItem = RSDN![jumlah tanaman]
                    End If
                    If IsNull(RSDN![jumlah tanaman]) Then
                        JumlahTanaman = 1
                    Else
                        luasBangunan = RSDN![jumlah tanaman]
                    End If
                    
                    'MsgBox rstanaman!harga & Chr(13) & JumlahItem & Chr(13) & luasbangunan & Chr(13) & NilaiSusut
                    RSDN![Nilai tanaman per meter persegi] = rstanaman!harga '* JumlahItem * luasbangunan * NilaiSusut / 100
                    RSDN![jumlah nilai tanaman] = rstanaman!harga * JumlahItem * JumlahTanaman
                    
                    'RSDN![Nilai Pasar Tanah] = rstanaman!harga * RSDN![Luas Hasil Ukur di Dalam Trase]
                    RSDN.Update
                    
                    rstanaman!keterangan = "NPW telah diupdate untuk tanaman pada tanggal " & Format(Date, "dd-mm-yyyy")
                    
                    
                Else
                    rstanaman!keterangan = "NIB tidak ditemukan pada database"
                End If
                rstanaman.MoveNext
            Wend
            
            
            
'kalkulasi tahap 2 nilai pasar bangunan
            rstanaman.MoveFirst
            While Not rstanaman.EOF
                RSDN.Close
                rsKalkulasi.Open "select sum([jumlah tanaman]) from [daftar nominatif] where idnib='" & rstanaman!nib & " ' and pemilik='" & rstanaman!Pemilik & "' group by idnib", db, adOpenDynamic, adLockOptimistic
                RSDN.Open "select id,urutid,[Nilai pasar tanaman] from [daftar nominatif] where idnib='" & rstanaman!nib & " ' and pemilik='" & rstanaman!Pemilik & "' order by urutid"
                If Not RSDN.EOF Then
                    RSDN![nilai pasar tanaman] = rsKalkulasi.Fields(0)
                    RSDN.Update
                'MsgBox rstanaman!identitas
                'MsgBox RSDN.Fields(0)
                'RSDN.Close
                End If
                rsKalkulasi.Close
                rstanaman.MoveNext
            Wend
        End If
        
        'Sumber = rsKolom!Source
        'Set RSDN = New ADODB.Recordset
        'RSDN.Open "SELECT * FROM [DAFTAR NOMINATIF] order by UrutId", db, adOpenDynamic, adLockOptimistic
        'RSDN.Open Sumber, db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"

        'Set DaftarNominatif.grdNominatif.DataSource = RSDN
        'DaftarNominatif.grdNominatif.ReBind

       ' RapihkanGrid

       
  rstanaman.Filter = ""
    rstanaman.Requery
End If
End Sub

Private Sub Command23_Click()

 
 Dim i As Integer
    If Not rstanaman.RecordCount = 0 Then
        rstanaman.MoveFirst
    End If
     While Not rstanaman.EOF
        If Not rsHarga.RecordCount = 0 Then
        rsHarga.MoveFirst
        End If
        rsHarga.Filter = "[jenis tanaman]='" & rstanaman![Jenis tanaman] & "' and [ukuran]='" & rstanaman![ukuran jenis tanaman] & "'"
        'MsgBox "[jenis tanaman]='" & rsDaftarTanaman![jenis tanaman] & "' and ukuran = '" & rsDaftarTanaman![Ukuran Jenis Tanaman] & "'"
        If Not rsHarga.EOF Then
            rstanaman!harga = rsHarga!harga
            rsHarga.Update
        End If
        rstanaman.MoveNext
    Wend
    rsHarga.Filter = ""
    rsHarga.Sort = "[jenis tanaman],ukuran"
    rsHarga.Requery
    
    
End Sub

Private Sub Command3_Click()
grdHarga.Splits(0).MarqueeStyle = dbgHighlightRow
grdHarga.Splits(0).Locked = True

End Sub

Private Sub Command9_Click()
 Dim rsDaftarNIB As ADODB.Recordset
 Dim i As Integer
 
    Dim db As ADODB.Connection
    'Tampilan.Top = txtNama.Top
    'Tampilan.Left = txtNama.Left
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
    'Set rstanaman = New ADODB.Recordset
    'rstanaman.Open "select * from [Zona Tanah] ", db, adOpenDynamic, adLockOptimistic
    Set rsDaftarNIB = New ADODB.Recordset
    'rsDaftarNIB.Open "SELECT [Daftar Bangunan].[id],[Daftar Bangunan].urutid, [Daftar Bangunan].idNIB, [Daftar Bangunan].Pemilik FROM [Daftar Bangunan] LEFT JOIN [Bangunan] ON [Daftar Bangunan].urutid = [Bangunan].[ID] WHERE ((([Bangunan].ID) Is Null));", db, adOpenDynamic, adLockOptimistic
     rsDaftarNIB.Open "SELECT [Daftar tanaman].idNIB FROM [Daftar tanaman] LEFT JOIN tanaman ON [Daftar tanaman].[idNIB] = tanaman.[NIB] GROUP BY [Daftar tanaman].idNIB, tanaman.NIB HAVING (((tanaman.NIB) Is Null));", db, adOpenDynamic, adLockOptimistic
     
  
    
    List5.Clear
    List6.Clear
    While Not rsDaftarNIB.EOF
        
        List5.AddItem "" & rsDaftarNIB!idnib
        rsDaftarNIB.MoveNext
    Wend
    rsDaftarNIB.Close
    rsDaftarNIB.Open "SELECT nib from tanaman where nib is not null group by nib", db, adOpenDynamic, adLockOptimistic
    While Not rsDaftarNIB.EOF
        
        List6.AddItem rsDaftarNIB!nib
        rsDaftarNIB.MoveNext
    Wend
        'X = rsKolom![indeks kolom]
        'If rsKolom!tampil = True Then
         '   List6.AddItem "" & rsKolom!isi
        'Else
        '    List5.AddItem rsKolom!isi
        'End If
        'rsKolom.MoveNext
    'Wend
    Tampilan.Visible = True
End Sub



Private Sub Form_Load()

 Set MyProperty = GRDTanaman   'nama datagrid yang inigin di scroll dengan mouse
        WheelHook GRDTanaman
'Set MyProperty = grdHarga   'nama datagrid yang inigin di scroll dengan mouse
'        WheelHook grdHarga
If pROJECTPATH = "" Then
MsgBox "Maaf belum ada file project yang dipilih", vbInformation
Me.Hide
Exit Sub
End If
Dim db As ADODB.Connection
'Dim rsKolom As ADODB.Recordset
'Dim Sumber As String
    Set db = New ADODB.Connection
     ' Set rsKolom = New ADODB.Recordset
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database "    ';pwd=globalisasi"
        db.CursorLocation = adUseClient
   
        'rsKolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Nominatif' ", db, adOpenDynamic, adLockOptimistic
        'Sumber = rsKolom!Source
        Set rsHarga = New ADODB.Recordset
            rsHarga.Open "select * from [Harga tanaman] order by [jenis tanaman]", db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"

        Set grdHarga.DataSource = rsHarga
        grdHarga.ReBind
        grdHarga.Columns("ID").Visible = False
        Set rstanaman = New ADODB.Recordset
        rstanaman.Open "SELECT *  from tanaman ORDER BY NOMOR", db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"
        
        Set GRDTanaman.DataSource = rstanaman
        GRDTanaman.ReBind
        GRDTanaman.Columns("idTanaman").Visible = False
        GRDTanaman.Columns("ID").Visible = False
        'grdTanaman.Columns("Zona Tanah").Button = True
        'MsgBox rsHarga.RecordCount
        'Set List1.DataSource = rsHarga
        'Set List1.RowSource = rsHarga 'rsHarga![zona tanah]
        ' List1.ListField = "Zona Tanah"
        'List1.ReFill
End Sub

Private Sub Form_Resize()
Frame1.Top = Label2.Top + Label2.Height + 100 ' Command10.Top + Command10.Height + 100
Frame1.Left = 0
'label3.Top=frame1.
End Sub

Private Sub grdTanaman_AfterColEdit(ByVal ColIndex As Integer)
If GRDTanaman.Col = GRDTanaman.Columns("Penyusutan").ColIndex Then
    'rsHarga.MoveFirst
    'rsHarga.Find "Nomor ='" & grdTanaman.Columns("Penyusutan").Value & "'"
    'If Not rsHarga.EOF Then
    '    grdTanaman.Columns("Harga").Value = rsHarga!harga
        'Sendkeys "{down}"
        rstanaman.MoveNext
        GRDTanaman.Col = GRDTanaman.Col - 1
    'Else
    'MsgBox "Maaf nomor zona tersebut belum diidentifikasi, silahkan tambahkan pada daftar harga zona!", vbCritical
    'grdTanaman.SetFocus
    'End If
    Exit Sub
End If
If GRDTanaman.Col = GRDTanaman.Columns("Klasifikasi").ColIndex Then
    rsHarga.MoveFirst
    rsHarga.Find "Nomor ='" & GRDTanaman.Columns("Klasifikasi").Value & "'"
    If Not rsHarga.EOF Then
        GRDTanaman.Columns("Harga").Value = rsHarga!harga
        GRDTanaman.Col = GRDTanaman.Col + 1
        'Sendkeys "{down}"
        'rstanaman.MoveNext
    Else
    MsgBox "Maaf nomor zona tersebut belum diidentifikasi, silahkan tambahkan pada daftar harga zona!", vbCritical
    GRDTanaman.SetFocus
    End If
End If


End Sub

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
Private Sub grdTanaman_ButtonClick(ByVal ColIndex As Integer)
'tampilList
Dim co
    Set co = GRDTanaman.Columns(GRDTanaman.Col)
    List1.Left = DataGrid1.Left + co.Left + co.Width
    List1.Top = DataGrid1.Top + GRDTanaman.RowTop(GRDTanaman.Row)
    List1.Visible = True
        
        
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
pindah List1.hwnd, X, Y

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

