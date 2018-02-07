VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      WindowList      =   -1  'True
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
      Begin VB.Menu mnDaftarNominatif 
         Caption         =   "&Daftar Nominatif"
      End
      Begin VB.Menu mnDataFisik 
         Caption         =   "Data &Fisik"
         Begin VB.Menu mntanaH 
            Caption         =   "&Tanah"
         End
         Begin VB.Menu mNbangunan 
            Caption         =   "&Bangunan"
         End
      End
      Begin VB.Menu mnDataNonfisik 
         Caption         =   "Data &Nonfisik"
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
            If (Not IsNull(access.Fields(1))) Then
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

        For i = 1 To access.RecordCount

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
    DaftarNominatif.Show
End Sub

Private Sub mnExit_Click()
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

Private Sub mntanaH_Click()
    Tanah.Show
End Sub


