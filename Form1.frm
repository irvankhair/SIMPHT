VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3135
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1800
      Width           =   1575
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6135
      Left            =   3120
      TabIndex        =   1
      Top             =   3720
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10821
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
            LCID            =   18441
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
            LCID            =   18441
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
   Begin MSComDlg.CommonDialog saveToExcelDialog 
      Left            =   2400
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   855
      Left            =   720
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
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
    
    Dim NamaFileTemp As String
    Dim oExcel As Object
    Dim oBook As Object
    Dim oSheet As Object
    Dim sel As String



    NamaFileTemp = CreateTempFile("Pjn")
    Set oExcel = CreateObject("Excel.Application")
    Set oBook = oExcel.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    oSheet.Range("A1:I1").Select
    
    'coding untuk membuat garis pada cell
    'With oSheet.Selection
    '.HorizontalAlignment = xlCenter
    '.VerticalAlignment = xlBottom
    '.WrapText = False
    '.Orientation = 0
    '.AddIndent = False
    '.IndentLevel = 0
    '.ShrinkToFit = False
    '.ReadingOrder = xlContext
    '.MergeCells = False
    'End With
    'oSheet.range("A1:I1").Merge
    oSheet.Range("A1:I1").Select

    With oSheet.Range("A1:I1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    'oSheet.Range("g3:I3").MergeCells = True
    'oSheet.Range("g3:I3").WrapText = True

    oSheet.Range("a" & RSDN.RecordCount + 8).Value = "Pembuat Laporan"
    oSheet.Range("a" & RSDN.RecordCount + 12).Value = "(……………………………)"
    oSheet.Range("f" & RSDN.RecordCount + 8).Value = "Diperiksa Oleh"
    oSheet.Range("f" & RSDN.RecordCount + 12).Value = "(……………………………)"
    oSheet.Range("J" & RSDN.RecordCount + 8).Value = "Diterima Oleh"
    oSheet.Range("J" & RSDN.RecordCount + 12).Value = "(……………………………)"
    Range("a" & RSDN.RecordCount + 6 & ":K" & RSDN.RecordCount + 6).Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("a" & RSDN.RecordCount + 6 & ":K" & RSDN.RecordCount + 6).Borders(xlEdgeTop).Weight = xlThin



    oSheet.Range("a1").Value = "RESUME"
    oSheet.Range("a4").Value = "Hari/tgl : " & Date
    oSheet.Range("a3").Value = NamaProjek  '"Sumber : " & lblLokal.Caption
    'oSheet.Range("g4").Value = "Kode Akses : " & txtKodeAkses.text
    'oSheet.Range("g3").Value = "Tujuan : " & lblTujuan.Caption
    Range("a5:K5").Borders(xlEdgeLeft).LineStyle = xlContinuous
    Range("a5:K5").Borders(xlEdgeLeft).Weight = xlThin
    Range("a5:K5").Borders(xlEdgeLeft).ColorIndex = 0
    Range("a5:K5").Borders(xlEdgeLeft).TintAndShade = 0
    Range("a5:K5").Borders(xlEdgeTop).LineStyle = xlContinuous
    Range("a5:K5").Borders(xlEdgeTop).Weight = xlThin
    Range("a5:K5").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("a5:K5").Borders(xlEdgeBottom).Weight = xlThin
    Range("a5:K5").Borders(xlEdgeRight).LineStyle = xlContinuous
    Range("a5:K5").Borders(xlEdgeRight).Weight = xlThin
    'oSheet.Columns("c:c").ColumnWidth = 18
    'oSheet.Columns("d:d").ColumnWidth = 18

    'coding untuk mengatur kebar kolom
    oSheet.Columns("e:e").ColumnWidth = 33
    oSheet.Columns("f:f").ColumnWidth = 7.57
    oSheet.Columns("g:g").ColumnWidth = 8.43
    oSheet.Columns("i:i").ColumnWidth = 30
    oSheet.Columns("a:a").ColumnWidth = 4

    'pilih kolom mana yang mau di hide
    'oSheet.Columns("a:a").Hidden = True
    'oSheet.Columns("h:h").Hidden = True
    ''oSheet.Columns("b:b").Hidden = True
    'oSheet.Columns("c:c").Hidden = True
    'oSheet.Columns("d:d").Hidden = True
    'oSheet.Columns("J:J").Hidden = True
    'oSheet.Columns("I:I").Hidden = True

    'Selection.EntireColumn.Hidden = True
    oSheet.Rows("1:1").RowHeight = 30.75
    'oSheet.range("A1:I1").Select
    For i = 1 To RSDN.Fields.Count
        oSheet.Cells(5, i) = RSDN.Fields(i - 1).Name
    Next i
    'Transfer the data to Excel


    oSheet.Range("A6").CopyFromRecordset RSDN



    If oExcel.Version > "11.0" Then
        oBook.SaveAs NamaFileTemp & ".xlsx"
        NamaFileTemp = NamaFileTemp & ".xlsx"
    Else
        oBook.SaveAs NamaFileTemp & ".xls"
        NamaFileTemp = NamaFileTemp & ".xls"
    End If
    oExcel.Quit
    X = ShellEx(Me.hwnd, "open", NamaFileTemp, "", "", 10)
    '               tandaSelesai.Visible = True
    '              Shape1.Visible = True
    '             If txtKodeAkses = Operator & Format(Now, "ddMMyyyyhhmm") Then
    '            txtKodeAkses = Operator & Format(Now, "ddMMyyyyhhmm") + 1
    ''           Else
    '         txtKodeAkses = Operator & Format(Now, "ddMMyyyyhhmm")
    '        End If


End Sub

Private Sub Command2_Click()
saveToExcelDialog.Filter = "Excel 2003 (*.xlsx)|*.xlsx"
    saveToExcelDialog.ShowSave
    FileCopy App.Path & "\Temp.xlsx", saveToExcelDialog.FileName
    exportExcel = saveToExcelDialog.FileName

    Dim excelApp As Excel.Application
    Dim excelWB As Excel.Workbook
    Dim excelWS As Excel.Worksheet
    Set excelApp = CreateObject("Excel.Application")
     excelApp.Visible = False
    excelApp.Workbooks.Open exportExcel
    Set excelWB = excelApp.Workbooks(1)
    Set excelWS = excelWB.Worksheets(1)
    MsgBox excelWS.Cells(2, 2)
    excelWS.Activate
End Sub

