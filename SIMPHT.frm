VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SIMPHT"
   ClientHeight    =   6690
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   10275
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   0
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":0000
            Key             =   "s_Key1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":049A
            Key             =   "s_Key2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":0934
            Key             =   "s_Key3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":0DCE
            Key             =   "s_Key4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":1268
            Key             =   "s_Key5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":1702
            Key             =   "s_Key6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":1B9C
            Key             =   "s_Key7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":2036
            Key             =   "s_Key8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":24D0
            Key             =   "s_Key9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":296A
            Key             =   "s_Key10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SIMPHT.frx":2E04
            Key             =   "s_Key11"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Tag             =   "&File|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=-1)"
      WindowList      =   -1  'True
      Begin VB.Menu mnOpen 
         Caption         =   "&Open"
         Tag             =   "&Open|#s_Key7|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnSave 
         Caption         =   "&Save"
         Tag             =   "&Save|#s_Key10|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnEksport 
         Caption         =   "&Eksport"
         Tag             =   "&Eksport|#s_Key4|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnPrint 
         Caption         =   "&Print"
         Tag             =   "&Print|#s_Key8|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnExit 
         Caption         =   "&Exit"
         Tag             =   "&Exit|#s_Key3|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
   Begin VB.Menu mnInsert 
      Caption         =   "&Insert"
      Tag             =   "&Insert|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnBacaDataNominatif 
         Caption         =   "&Baca Data Nominatif"
         Tag             =   "&Baca Data Nominatif|#s_Key2|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
   Begin VB.Menu mnEdit 
      Caption         =   "&Edit"
      Tag             =   "&Edit|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      Begin VB.Menu mnDaftarNominatif 
         Caption         =   "&Daftar Nominatif"
         Tag             =   "&Daftar Nominatif|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnResume 
         Caption         =   "&Resume"
         Tag             =   "&Resume|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu mnDataFisik 
         Caption         =   "Data &Fisik"
         Tag             =   "Data &Fisik|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mntanaH 
            Caption         =   "&Tanah"
            Tag             =   "&Tanah|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu mNbangunan 
            Caption         =   "&Bangunan"
            Tag             =   "&Bangunan|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         End
         Begin VB.Menu MnTanaman 
            Caption         =   "&Tanaman"
         End
      End
      Begin VB.Menu mnDataNonfisik 
         Caption         =   "Data &Nonfisik"
         Tag             =   "Data &Nonfisik|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
         Begin VB.Menu mnUsaha 
            Caption         =   "&Kerugian Usaha"
         End
      End
   End
   Begin VB.Menu mnSetting 
      Caption         =   "&Setting"
      Tag             =   "&Setting|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu mnReport 
      Caption         =   "&Report"
      Tag             =   "&Report|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
   End
   Begin VB.Menu mnEditBaris 
      Caption         =   "Edit Baris"
      Tag             =   "Edit Baris|(Checked=0)(Enabled=-1)(Visible=0)(WindowList=0)"
      Visible         =   0   'False
      Begin VB.Menu EditBaris 
         Caption         =   "Edit"
         Tag             =   "Edit|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu SisipBaris 
         Caption         =   "Sisipkan"
         Tag             =   "Sisipkan|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
      Begin VB.Menu Hapusbaris 
         Caption         =   "Hapus Baris"
         Tag             =   "Hapus Baris|(Checked=0)(Enabled=-1)(Visible=-1)(WindowList=0)"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents MenuEvents As CEvents
Attribute MenuEvents.VB_VarHelpID = -1

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

    conConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & namaMdb & ";Mode=Read|Write"
    conConnection.Open
    rstRecordSet.Open "Select * from [daftar nominatif]", conConnection, adOpenStatic, adLockOptimistic
    Set konekAccess = rstRecordSet

End Function

Private Sub mnBacaDataNominatif_click()
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
    X = MsgBox("Apakah Anda akan mengambil data nominatif dari file " & CommonDialog1.FileTitle, vbYesNo, "Konfirmasi Input Data Nominatif")
    If X = vbYes Then
        CommonDialog1.Filter = "Acces 2003 (*.mdb)|*.mdb"
        CommonDialog1.ShowSave
        FileCopy App.Path & "\master.mdb", CommonDialog1.FileName
        namaMdb = CommonDialog1.FileName
        Transfer.Show
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
                For j = 2 To 19
                    .Fields(j) = excel.Fields(j - 2)

                Next
                If (Not IsNull(access.Fields("nomor urut"))) Then
                    access.Fields("Pemilik") = access.Fields("Identitas")
                End If
                .Fields("Index Benda Lain yang Berkaitan") = excel.Fields("Index Benda Lain yang Berkaitan")
                .Fields("Jenis Benda Lain yang Berkaitan") = excel.Fields("Jenis Benda Lain yang Berkaitan")
                .Fields("Jumlah Benda Lain yang Berkaitan") = excel.Fields("Jumlah Benda Lain yang Berkaitan")
                If ((excel.Fields(18) <> 0) And (excel.Fields(19) <> 0) And (excel.Fields(20) <> 0)) Then
                    .Fields("Ukuran Jenis Tanaman") = "Besar"
                    .Fields("Jumlah tanaman") = excel.Fields(18)
                    .Fields("jenis tanaman") = excel.Fields(17)
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
                    .Fields("Jumlah tanaman") = excel.Fields(19)
                    .Fields("jenis tanaman") = excel.Fields(17)
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
                    .Fields("Jumlah tanaman") = excel.Fields(20)
                    .Fields("jenis tanaman") = excel.Fields(17)
                ElseIf ((excel.Fields(18) <> 0) And (excel.Fields(19) <> 0)) Then
                    .Fields("Ukuran Jenis Tanaman") = "Besar"
                    .Fields("Jumlah tanaman") = excel.Fields(18)
                    .Fields("jenis tanaman") = excel.Fields(17)
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
                    .Fields("Jumlah tanaman") = excel.Fields(19)
                    .Fields("jenis tanaman") = excel.Fields(17)
                ElseIf ((excel.Fields(19) <> 0) And (excel.Fields(20) <> 0)) Then
                    .Fields("Ukuran Jenis Tanaman") = "Sedang"
                    .Fields("Jumlah tanaman") = excel.Fields(19)
                    .Fields("jenis tanaman") = excel.Fields(17)
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
                    .Fields("Jumlah tanaman") = excel.Fields(20)
                    .Fields("jenis tanaman") = excel.Fields(17)
                ElseIf ((excel.Fields(18) <> 0) And (excel.Fields(20) <> 0)) Then
                    .Fields("Ukuran Jenis Tanaman") = "Besar"
                    .Fields("Jumlah tanaman") = excel.Fields(18)
                    .Fields("jenis tanaman") = excel.Fields(17)
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
                    .Fields("Jumlah tanaman") = excel.Fields(20)
                    .Fields("jenis tanaman") = excel.Fields(17)
                ElseIf (excel.Fields(18) <> 0) Then
                    .Fields("Ukuran Jenis Tanaman") = "Besar"
                    .Fields("Jumlah tanaman") = excel.Fields(18)
                    .Fields("jenis tanaman") = excel.Fields(17)
                    If IsNull(access.Fields("idNIB")) Then
                        access.Fields("idNIB") = temp
                    ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                        temp = access.Fields("idNIB")
                    End If
                    If (Not IsNull(access.Fields(2))) Then
                        access.Fields("NIB") = access.Fields("idNIB")
                    End If

                ElseIf (excel.Fields(19) <> 0) Then
                    .Fields("Ukuran Jenis Tanaman") = "Sedang"
                    .Fields("Jumlah tanaman") = excel.Fields(19)
                    .Fields("jenis tanaman") = excel.Fields(17)
                    If IsNull(access.Fields("idNIB")) Then
                        access.Fields("idNIB") = temp
                    ElseIf (Not IsNull(access.Fields("idNIB")) And (access.Fields("idNIB") <> temp)) Then
                        temp = access.Fields("idNIB")
                    End If
                    If (Not IsNull(access.Fields(2))) Then
                        access.Fields("NIB") = access.Fields("idNIB")
                    End If

                ElseIf (excel.Fields(20) <> 0) Then
                    .Fields("Ukuran Jenis Tanaman") = "Kecil"
                    .Fields("Jumlah tanaman") = excel.Fields(20)
                    .Fields("jenis tanaman") = excel.Fields(17)
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
            If (Not IsNull(access.Fields(2))) Then
                access.Fields("NIB") = access.Fields("idNIB")
            End If

            access.Update

            excel.MoveNext
            access.MoveNext
            If Transfer.ProgressBar1.Value < 100 Then
                Transfer.ProgressBar1.Value = ((i / excel.RecordCount) * 100)
            ElseIf (i >= excel.RecordCount) Then
                Unload Transfer
            End If

        Next
        access.MoveFirst
        temp = access.Fields("Identitas")
        For i = 1 To access.RecordCount
            If IsNull(access.Fields("nomor urut")) Then
                access.Fields("Pemilik") = temp
            ElseIf (Not IsNull(access.Fields("nomor urut")) And (access.Fields("nomor urut") <> temp)) Then
                temp = access.Fields("Identitas")

            End If
            access.Fields(1) = access.Fields(0)
            access.Update
            access.MoveNext
        Next
        Unload Transfer
        MsgBox "Transfered"
        excel.Close
        access.Close
    End If
    Set excel = Nothing
    Set access = Nothing
    tmp = Empty
    excelPath = Empty
End Sub


Private Sub mNbangunan_Click()
    Bangunan.Show
End Sub

Private Sub mnDaftarNominatif_Click()
    If (RSDN.State = 0) Then
        MsgBox "Anda Belum membuka file Project"
    Else
        DaftarNominatif.Show
    End If
End Sub

Private Sub mnExit_Click()
Unload Me
    End
End Sub

Private Sub mnOpen_Click()

    CommonDialog1.Filter = "Acces 2003 (*.mdb)|*.mdb"
    CommonDialog1.ShowOpen
    If CommonDialog1.FileTitle = "" Then
        Exit Sub
    End If
    pROJECTPATH = CommonDialog1.FileName
    NamaProjek = Left(CommonDialog1.FileTitle, Len(CommonDialog1.FileTitle) - 4)
    DaftarNominatif.Label1 = NamaProjek

    DaftarNominatif.Show
End Sub

Private Sub tanaH_Click()

End Sub

Private Sub mnResume_Click()
    ResumeNilai.Show
End Sub

Private Sub mntanaH_Click()
    Tanah.Show
End Sub
Private Sub MnTanaman_Click()
Tanaman.Show
End Sub

Private Sub MenuEvents_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    If gbSubClassMenu Then
        '/ this event show Decription menu-item to a StatusBar control
        '/ customize for your project, if you need this.
        '/ Note: MenuText contain the menu Caption.
        '/       MenuHelp contain the Description string.
        '/ example: sbStatusBar.Panels(1).text = MenuHelp
        '/          lblMenuHelp.Caption = MenuHelp
    End If
End Sub
Public Sub SubClassMenuXP()

'/ this code is made by MenuCreator add-in

'/ prepare the caption for subclassing. Warning! Don't remove this comment!!!
    mnFile.Caption = "&File"
    mnOpen.Caption = "&Open|#s_Key7"
    mnSave.Caption = "&Save|#s_Key10"
    mnEksport.Caption = "&Eksport|#s_Key4"
    mnPrint.Caption = "&Print|#s_Key8"
    mnExit.Caption = "&Exit|#s_Key3"
    mnInsert.Caption = "&Insert"
    mnBacaDataNominatif.Caption = "&Baca Data Nominatif|#s_Key2"
    mnEdit.Caption = "&Edit"
    mnDaftarNominatif.Caption = "&Daftar Nominatif"
    mnResume.Caption = "&Resume"
    mnDataFisik.Caption = "Data &Fisik"
    mntanaH.Caption = "&Tanah"
    mNbangunan.Caption = "&Bangunan"
    mnDataNonfisik.Caption = "Data &Nonfisik"
    mnSetting.Caption = "&Setting"
    mnReport.Caption = "&Report"
    mnEditBaris.Caption = "Edit Baris"
    EditBaris.Caption = "Edit"
    SisipBaris.Caption = "Sisipkan"
    Hapusbaris.Caption = "Hapus Baris"


    '/ Subclassing menu. Warning! Don't remove this comment!!!

    Set MenuEvents = New CEvents
    Set objMenuEx = New cMenuEx
    Call objMenuEx.Install(Me.hwnd, App.Path & "\" & Me.Name, ImageList1, 2, MenuEvents)

End Sub

Public Sub MenuDesigner()
'/ Open Menu Designer tool
    objMenuEx.MenuDesigner Me.hwnd
End Sub

Private Sub MDIForm_Load()
'/ This MDIForm_Load is made by MenuCreator

'/ If gbSubClassMenu is False, the menu is not subclassed
    gbSubClassMenu = True

    If gbSubClassMenu Then SubClassMenuXP

End Sub

Private Sub MDIForm_UnLoad(Cancel As Integer)
'/ This MDIForm_UnLoad is made by MenuCreator

    If gbSubClassMenu Then
        '/ prevent error if the menu is not subclassed
        On Error Resume Next
        '/ release object
        Call objMenuEx.Uninstall(Me.hwnd, ImageList1, MenuEvents)
        Set MenuEvents = Nothing
        Set objMenuEx = Nothing
    End If

End Sub

Private Sub mnUsaha_Click()
Usaha.Show
End Sub
