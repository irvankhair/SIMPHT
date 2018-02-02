VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DaftarNominatif 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Daftar Nominatif"
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
   Visible         =   0   'False
   WindowState     =   2  'Maximized
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
      Left            =   3840
      TabIndex        =   4
      Top             =   2040
      Width           =   2895
   End
   Begin VB.CommandButton Transfer 
      Caption         =   "Transfer Ke Excell"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6255
   End
   Begin VB.Menu EditBaris 
      Caption         =   "Edit Baris"
      Visible         =   0   'False
      Begin VB.Menu Sisip 
         Caption         =   "Sisipkan"
      End
      Begin VB.Menu HapusBaris 
         Caption         =   "Hapus"
      End
   End
End
Attribute VB_Name = "DaftarNominatif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<<<<<<< HEAD

Private Function RapihkanGrid()
grdNominatif.Columns("id").Visible = False
grdNominatif.Columns("urutid").Visible = False
grdNominatif.Columns("idnib").Visible = False


Dim i As Integer
Dim db As ADODB.Connection
    
    
    
    
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Jet OLEDB:Database Password=globalisasi;Persist Security Info=False"
    db.CursorLocation = adUseClient
      Set rskolom = New ADODB.Recordset
   
     rskolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Nominatif' ", db, adOpenDynamic, adLockOptimistic
      

For i = 0 To grdNominatif.Columns.Count - 1
    rskolom.MoveFirst
    rskolom.Find "isi='" & grdNominatif.Columns(i).Caption & "'"
    If Not rskolom.EOF Then
        grdNominatif.Columns(i).Width = rskolom![lebar kolom]
    
    End If
Next i
End Function

Private Sub Command2_Click()
Dim rskolom As ADODB.Recordset
Dim db As ADODB.Connection
 Set db = New ADODB.Connection
 db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database " ';pwd=globalisasi"
db.CursorLocation = adUseClient
    Set rskolom = New ADODB.Recordset
    rskolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Nominatif'", db, adOpenDynamic, adLockOptimistic
   While Not rskolom.EOF
   rskolom.Delete
   rskolom.MoveNext
   Wend
   For i = 0 To grdNominatif.Columns.Count - 1
        rskolom.AddNew
        rskolom![nama tabel] = "Daftar Nominatif"
        rskolom![indeks kolom] = grdNominatif.Columns(i).ColIndex
        rskolom!isi = grdNominatif.Columns(i).Caption
        rskolom![lebar kolom] = grdNominatif.Columns(i).Width
        rskolom!tipe = RSDN.Fields(i).Type
        rskolom.Update
    Next i
     
  

End Sub

Private Sub Form_Load()
Dim db As ADODB.Connection
 Set db = New ADODB.Connection
 db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False;Jet OLEDB:Database " ';pwd=globalisasi"
db.CursorLocation = adUseClient
Set RSDN = New ADODB.Recordset
RSDN.Open "SELECT * FROM [DAFTAR NOMINATIF] order by urutid", db, adOpenDynamic, adLockOptimistic

Set DaftarNominatif.grdNominatif.DataSource = RSDN
DaftarNominatif.grdNominatif.ReBind

RapihkanGrid

Set MyProperty = DaftarNominatif.grdNominatif   'nama datagrid yang inigin di scroll dengan mouse
WheelHook grdNominatif
End Sub

Private Sub Form_Resize()
If Not Me.ScaleHeight = 0 Then
    Label1.Width = Me.Width
    Line1.X2 = Me.ScaleWidth
    grdNominatif.Top = Line1.Y1 + 30
    grdNominatif.Width = Me.ScaleWidth
    grdNominatif.Height = Me.ScaleHeight - Line1.Y1 - 30
    grdNominatif.Left = 0
End If
End Sub


Private Sub grdNominatif_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
Dim db As ADODB.Connection
    Set db = New ADODB.Connection
    db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pROJECTPATH & ";Persist Security Info=False"
    db.CursorLocation = adUseClient
    Set rskolom = New ADODB.Recordset
    rskolom.Open "select * from [kostum tabel] where [nama tabel]='Daftar Nominatif'", db, adOpenDynamic, adLockOptimistic
   rskolom.Find "isi='" & grdNominatif.Columns(ColIndex).Caption & "'"
   If Not rskolom.EOF Then
   rskolom![lebar kolom] = grdNominatif.Columns(ColIndex).Width
   'rskolom!tipe = rs.Fields(0).Type
    rskolom.Update
    End If
End Sub

Private Sub grdNominatif_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
PopupMenu EditBaris
End If
End Sub

Private Sub Sisip_Click()
Dim urutSebelum As String
Dim posisi As String
If Not RSDN.EOF = True And Not RSDN.BOF = True Then
    posisi = RSDN.AbsolutePosition
    RSDN.MovePrevious
    urutSebelum = RSDN!urutid
    RSDN.AddNew
    RSDN!urutid = urutSebelum + 0.1
    RSDN.Update
    RSDN.Requery
    RSDN.Move posisi
    
    RapihkanGrid
End If
End Sub
'=======
Private Sub Command1_Click()
    'MsgBox "okke"
    Dim excel As New ADODB.Recordset
    Set excel = importExcel
    MsgBox excel.RecordCount
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

    conConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & "gian.mdb;Mode=Read|Write"
    conConnection.Open
    rstRecordSet.Open "Select * from Daftar_Nominatif", conConnection, adOpenStatic, adLockOptimistic
    Set konekAccess = rstRecordSet

End Function

Sub cekNIB(ByRef temp As String, ByRef nib As Variant)
    If IsNull(nib) Then
        nib = temp
    ElseIf (Not IsNull(nib) And (nib <> temp)) Then
        temp = nib
    End If
End Sub
>>>>>>> d3d07549d5308c63fd676dc3c8bf461ab7c97286

Private Sub Transfer_Click()


    Dim excel As New ADODB.Recordset
    Dim access As New ADODB.Recordset

    Dim i, j As Integer
    Set excel = importExcel
    Set access = konekAccess

    Dim temp As String

    temp = excel.Fields("NIB")
    For i = 1 To excel.RecordCount
        With access
            .AddNew
            For j = 1 To 18
                .Fields(j) = excel.Fields(j - 1)
            Next


            If ((excel.Fields(17) <> 0) And (excel.Fields(18) <> 0) And (excel.Fields(19) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = excel.Fields(17)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If


                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = excel.Fields(18)
                .Fields("jenis tanaman") = excel.Fields(16)
                .MoveNext
                .AddNew
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = excel.Fields(19)
                .Fields("jenis tanaman") = excel.Fields(16)
            ElseIf ((excel.Fields(17) <> 0) And (excel.Fields(18) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = excel.Fields(17)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")

                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = excel.Fields(18)
                .Fields("jenis tanaman") = excel.Fields(16)
            ElseIf ((excel.Fields(18) <> 0) And (excel.Fields(19) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = excel.Fields(18)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = excel.Fields(19)
                .Fields("jenis tanaman") = excel.Fields(16)
            ElseIf ((excel.Fields(17) <> 0) And (excel.Fields(19) <> 0)) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = excel.Fields(17)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .MoveNext
                .AddNew
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = excel.Fields(19)
                .Fields("jenis tanaman") = excel.Fields(16)
            ElseIf (excel.Fields(17) <> 0) Then
                .Fields("Ukuran Jenis Tanaman") = "Besar"
                .Fields("Jumlah tanaman") = excel.Fields(17)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

            ElseIf (excel.Fields(18) <> 0) Then
                .Fields("Ukuran Jenis Tanaman") = "Sedang"
                .Fields("Jumlah tanaman") = excel.Fields(18)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

            ElseIf (excel.Fields(19) <> 0) Then
                .Fields("Ukuran Jenis Tanaman") = "Kecil"
                .Fields("Jumlah tanaman") = excel.Fields(19)
                .Fields("jenis tanaman") = excel.Fields(16)
                If IsNull(access.Fields("NIB")) Then
                    access.Fields("NIB") = temp
                ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
                    temp = access.Fields("NIB")
                End If
                If (Not IsNull(access.Fields(1))) Then
                    access.Fields("NIB Palsu") = access.Fields("NIB")
                End If

            End If
        End With
        If IsNull(access.Fields("NIB")) Then
            access.Fields("NIB") = temp
        ElseIf (Not IsNull(access.Fields("NIB")) And (access.Fields("NIB") <> temp)) Then
            temp = access.Fields("NIB")

        End If

        'Call cekNIB(temp, access.Fields("NIB"))
        If (Not IsNull(access.Fields(1))) Then
            access.Fields("NIB Palsu") = access.Fields("NIB")
        End If

        access.Update

        excel.MoveNext
        access.MoveNext
    Next

    MsgBox "Transfered"
    excel.Close
    access.Close
End Sub
