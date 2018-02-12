VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form Usaha 
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
      Left            =   360
      ScaleHeight     =   7095
      ScaleWidth      =   5595
      TabIndex        =   13
      Top             =   1920
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
         TabIndex        =   26
         Top             =   6600
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         Height          =   375
         Left            =   5640
         Picture         =   "Usaha.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2040
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Height          =   375
         Left            =   5640
         Picture         =   "Usaha.frx":0342
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2520
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Height          =   375
         Left            =   5640
         Picture         =   "Usaha.frx":0684
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   3000
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Height          =   375
         Left            =   5640
         Picture         =   "Usaha.frx":09C6
         Style           =   1  'Graphical
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   3840
         Width           =   855
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Simpan"
         Height          =   375
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command16 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2640
         Picture         =   "Usaha.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5040
         Width           =   420
      End
      Begin VB.CommandButton Command17 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2640
         Picture         =   "Usaha.frx":140A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4440
         Width           =   420
      End
      Begin VB.CommandButton Command19 
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   2640
         Picture         =   "Usaha.frx":1B0C
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2280
         Width           =   420
      End
      Begin VB.CommandButton Command20 
         Appearance      =   0  'Flat
         BackColor       =   &H0000C0C0&
         Height          =   390
         Left            =   2640
         Picture         =   "Usaha.frx":220E
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   420
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         DragIcon        =   "Usaha.frx":2912
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
         ItemData        =   "Usaha.frx":9164
         Left            =   3360
         List            =   "Usaha.frx":9166
         MultiSelect     =   1  'Simple
         Sorted          =   -1  'True
         TabIndex        =   14
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
         TabIndex        =   28
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
         TabIndex        =   27
         Top             =   0
         Width           =   2655
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   10455
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   17295
      Begin VB.CommandButton Command22 
         Caption         =   "&Terapkan Kalkulasi NPW Kerugian Usaha"
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
         Left            =   9240
         TabIndex        =   35
         Top             =   8280
         Width           =   5055
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
         Left            =   1680
         TabIndex        =   34
         Top             =   8280
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
         Left            =   4560
         TabIndex        =   33
         Top             =   8280
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
         Left            =   3120
         TabIndex        =   32
         Top             =   8280
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
         Left            =   6000
         TabIndex        =   31
         Top             =   8280
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
         Left            =   240
         TabIndex        =   30
         Top             =   8280
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid GRDzona 
         Height          =   8055
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   16815
         _ExtentX        =   29660
         _ExtentY        =   14208
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
         Caption         =   "PERHITUNGAN KERUGIAN USAHA"
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
         TabIndex        =   12
         Top             =   120
         Width           =   6975
      End
   End
   Begin MSDataListLib.DataList List1 
      Height          =   3375
      Left            =   8760
      TabIndex        =   10
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
      Visible         =   0   'False
      Width           =   11655
      Begin VB.CommandButton Command21 
         Caption         =   "&Tutup"
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
         Left            =   9960
         TabIndex        =   9
         Top             =   240
         Width           =   1335
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
         Left            =   240
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
         Left            =   3480
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
         Left            =   1680
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
         Left            =   4920
         TabIndex        =   2
         Top             =   3600
         Width           =   1335
      End
      Begin MSDataGridLib.DataGrid grdHargaZona 
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
         Caption         =   "Harga Zona Tanah"
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
      Caption         =   "PENILAIAN KERUGIAN USAHA"
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
Attribute VB_Name = "Usaha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsHargaZona As ADODB.Recordset
Dim rsUsaha As ADODB.Recordset
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
grdHargaZona.Splits(0).MarqueeStyle = dbgFloatingEditor
grdHargaZona.Splits(0).Locked = False
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
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
    'Set rsUsaha = New ADODB.Recordset
    'rsUsaha.Open "select * from [Kerugian Usaha] ", db, adOpenDynamic, adLockOptimistic
    Set rsDaftarNIB = New ADODB.Recordset
    rsDaftarNIB.Open "SELECT * from [Daftar NIB Pemilik]", db, adOpenDynamic, adLockOptimistic
        If Not rsUsaha.RecordCount = 0 Then
            rsUsaha.MoveFirst
        End If
        While Not rsUsaha.EOF
            rsUsaha.Delete
            rsUsaha.MoveNext
        Wend
        
        For i = 0 To List6.ListCount - 1
            If Not rsDaftarNIB.RecordCount = 0 Then
                rsDaftarNIB.MoveFirst
            End If
            rsDaftarNIB.Find "idNIB='" & List6.List(i) & "'"
            If Not rsDaftarNIB.EOF Then
                rsUsaha.AddNew
                    rsUsaha!idusaha = rsDaftarNIB!ID
                    rsUsaha!nomor = rsDaftarNIB![nomor urut]
                    rsUsaha!nib = rsDaftarNIB!idnib
                    rsUsaha!Pemilik = rsDaftarNIB!Pemilik
                rsUsaha.Update
            End If
        Next i
        Tampilan.Visible = False
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
Frame1.Visible = False
End Sub

Private Sub Command22_Click()
X = MsgBox("Apakah Anda yakin untuk melakukan proses kalkulasi NPW tanah dengan harga pada zona yang telah ditetapkan?", vbYesNo, "Konfirmasi Kalkulasi NPW Tanah")
If X = vbYes Then
    rsUsaha.Filter = "harga > 0"
    rsUsaha.MoveFirst
    Dim db As ADODB.Connection
    Dim RSDN As ADODB.Recordset
    Dim Sumber As String
    Set db = New ADODB.Connection
      Set RSDN = New ADODB.Recordset
        db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database "    ';pwd=globalisasi"
        db.CursorLocation = adUseClient
   
        RSDN.Open "select id,nib,[Luas Hasil Ukur di Dalam Trase],[Nilai Tanah per Meter Persegi],[Nilai Pasar Tanah] from [daftar nominatif] where nib is not null ", db, adOpenDynamic, adLockOptimistic
        If Not RSDN.EOF Then
        
            While Not rsUsaha.EOF
                RSDN.MoveFirst
                RSDN.Find "nib='" & rsUsaha!nib & "'"
                If Not RSDN.EOF Then
                    RSDN![Nilai Tanah per Meter Persegi] = rsUsaha!harga
                    RSDN![Nilai Pasar Tanah] = rsUsaha!harga * RSDN![Luas Hasil Ukur di Dalam Trase]
                    RSDN.Update
                    
                    rsUsaha!keterangan = "NPW telah diupdate pada tanggal " & Format(Date, "dd-mm-yyyy")
                    
                Else
                    rsUsaha!keterangan = "NIB tidak ditemukan pada database"
                End If
                rsUsaha.MoveNext
            Wend
        End If
        
        'Sumber = rsKolom!Source
        'Set RSDN = New ADODB.Recordset
        'RSDN.Open "SELECT * FROM [DAFTAR NOMINATIF] order by UrutId", db, adOpenDynamic, adLockOptimistic
        'RSDN.Open Sumber, db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"

        'Set DaftarNominatif.grdNominatif.DataSource = RSDN
        'DaftarNominatif.grdNominatif.ReBind

       ' RapihkanGrid

       
  rsUsaha.Filter = "harga > 0"
    
End If

End Sub

Private Sub Command3_Click()
grdHargaZona.Splits(0).MarqueeStyle = dbgHighlightRow
grdHargaZona.Splits(0).Locked = True

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
    'Set rsUsaha = New ADODB.Recordset
    'rsUsaha.Open "select * from [Kerugian Usaha] ", db, adOpenDynamic, adLockOptimistic
    Set rsDaftarNIB = New ADODB.Recordset
    'rsDaftarNIB.Open "SELECT [Daftar NIB Pemilik].[id],[Daftar NIB Pemilik].[nomor urut], [Daftar NIB Pemilik].idNIB, [Daftar NIB Pemilik].Pemilik FROM [Daftar NIB Pemilik] LEFT JOIN [Kerugian Usaha] ON [Daftar NIB Pemilik].[nomor urut] = [Kerugian Usaha].[ID] WHERE ((([Kerugian Usaha].ID) Is Null));", db, adOpenDynamic, adLockOptimistic
    rsDaftarNIB.Open "SELECT [Daftar NIB Pemilik].[nomor urut], [Daftar NIB Pemilik].idNIB, [Daftar NIB Pemilik].Pemilik, [Daftar NIB Pemilik].id FROM [Daftar NIB Pemilik] LEFT JOIN [Kerugian Usaha] ON [Daftar NIB Pemilik].[idNIB] = [Kerugian Usaha].[NIB] WHERE ((([Kerugian Usaha].NIB) Is Null));", db, adOpenDynamic, adLockOptimistic
     
   
    
    List5.Clear
    List6.Clear
    While Not rsDaftarNIB.EOF
        
        List5.AddItem rsDaftarNIB!idnib
        rsDaftarNIB.MoveNext
    Wend
    rsDaftarNIB.Close
    rsDaftarNIB.Open "SELECT * from [Kerugian Usaha] where nib is not null", db, adOpenDynamic, adLockOptimistic
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

 Set MyProperty = GRDzona   'nama datagrid yang inigin di scroll dengan mouse
        WheelHook GRDzona

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
    '    Set rsHargaZona = New ADODB.Recordset
    '        rsHargaZona.Open "SELECT id,NOMOR,[Kerugian Usaha],HARGA,KETERANGAN from [harga zona tanah]", db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"

     '   Set grdHargaZona.DataSource = rsHargaZona
     '   grdHargaZona.ReBind

        Set rsUsaha = New ADODB.Recordset
        rsUsaha.Open "SELECT *  from [Kerugian Usaha] ORDER BY nib", db, adOpenDynamic, adLockOptimistic  '"SELECT * FROM [DAFTAR NOMINATIF] order by UrutId"
        'grdHarga.Columns("ID").Visible = False

        Set GRDzona.DataSource = rsUsaha
        GRDzona.ReBind
        GRDzona.Columns("idusaha").Visible = False
        GRDzona.Columns("ID").Visible = False
        'GRDzona.Columns("Zona Tanah").Button = True
        'MsgBox rsHargaZona.RecordCount
        'Set List1.DataSource = rsHargaZona
        'Set List1.RowSource = rsHargaZona 'rsHargaZona![Kerugian Usaha]
        ' List1.ListField = "Zona Tanah"
        'List1.ReFill
End Sub

Private Sub Form_Resize()
Frame1.Top = Label2.Top + Label2.Height + 100 ' Command10.Top + Command10.Height + 100
Frame1.Left = 0
'label3.Top=frame1.
End Sub

Private Sub GRDzona_AfterColEdit(ByVal ColIndex As Integer)
If GRDzona.Col = GRDzona.Columns("Kerugian Per Bulan").ColIndex Then
    'rsHargaZona.MoveFirst
    'rsHargaZona.Find "Nomor ='" & GRDzona.Columns("Zona Tanah").Value & "'"
    'If Not rsHargaZona.EOF Then
    '    GRDzona.Columns("Harga").Value = rsHargaZona!harga
          ' rsUsaha.MoveNext
    'Else
    'MsgBox "Maaf nomor zona tersebut belum diidentifikasi, silahkan tambahkan pada daftar harga zona!", vbCritical
    'MsgBox GRDzona.Columns("Jumlah Bulan")
    
    If GRDzona.Columns("Jumlah Bulan") = "" Then
       GRDzona.Columns("Total").Value = 0
    Else
        GRDzona.Columns("Total").Value = GRDzona.Columns("Kerugian Per Bulan").Value * GRDzona.Columns("Jumlah Bulan").Value
    End If
    GRDzona.Col = GRDzona.Col + 1
    GRDzona.SetFocus
    Exit Sub
End If
If GRDzona.Col = GRDzona.Columns("Jumlah Bulan").ColIndex Then
   GRDzona.Columns("Total").Value = GRDzona.Columns("Kerugian Per Bulan").Value * GRDzona.Columns("Jumlah Bulan").Value
    rsUsaha.MoveNext
    GRDzona.Col = GRDzona.Col - 1
End If


End Sub

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
Private Sub GRDzona_ButtonClick(ByVal ColIndex As Integer)
'tampilList
Dim co
    Set co = GRDzona.Columns(GRDzona.Col)
    List1.Left = DataGrid1.Left + co.Left + co.Width
    List1.Top = DataGrid1.Top + GRDzona.RowTop(GRDzona.Row)
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

