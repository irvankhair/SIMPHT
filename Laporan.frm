VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form Form9 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Pra Cetak"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Laporan.frx":0000
   LinkTopic       =   "Form9"
   MDIChild        =   -1  'True
   ScaleHeight     =   11535
   ScaleWidth      =   19080
   WindowState     =   2  'Maximized
   Begin CRVIEWERLibCtl.CRViewer CRViewer1 
      Height          =   5535
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   10095
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   -1  'True
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   -1  'True
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xApp As New CRAXDRT.Application
Dim xRpT As New CRAXDRT.Report
Dim xDbf As CRAXDRT.DatabaseTable
'dim xdbf as CRAXDRT.d





Private Sub Form_Load()

'pathdrive = "d:"
Select Case LaporanTerpilih

Case "Kartu Stok"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\Kartu Stok.rpt")
xRpT.ReportTitle = tanggalLaporan

Case "Rekapitulasi Fast Moving"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\REKAPITULASI FAST MOVING.rpt")
Case "Rekapitulasi Omzet Harian"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\REKAPITULASI OMZET HARIAN.rpt")

Case "Rekapitulasi Tiap Golongan"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\REKAPITULASI HASIL PENJUALAN TIAP GOLONGAN.rpt")
Case "Rekapitulasi Tiap Jenis"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\REKAPITULASI HASIL PENJUALAN TIAP JENIS.rpt")


Case "sembeli"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\daftar pembelian.rpt")

Case "Rekapitulasi Cara Bayar"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\Rekapitulasi Cara Bayar.rpt")
xRpT.ReportTitle = tanggalLaporan


Case "Laporan Kondisi Barang"

Set xRpT = xApp.OpenReport(PathDrive & "\laporan\LAPORAN STOK AKHIR BARANG.rpt")
xRpT.ReportTitle = tanggalLaporan

Case "Laporan Narkotika Psikotropika"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\LAPORAN NARKOTIKA PSIKOTROPIKA.rpt")
Case "Laporan Bulanan Penjualan"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\laporan bulanan penjualan.rpt")
Case "Laporan Pembelian"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\pembelian bulanan.rpt")
Case "Laporan Barang Masuk"
LaporanTerpilih = "Laporan Kondisi Barang"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\Laporan Barang Masuk.rpt")
Case "Laporan Laba Rugi"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\laporan laba rugi.rpt")
Case "Laporan Neraca Bulanan"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\laporan neraca bulanan.rpt")
Case "Laporan Harian Penjualan"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\laporan harian penjualan.rpt")
Case "Laporan Jasa Resep"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\laporan jasa resep.rpt")
Case "Daftar Expired"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\expired.rpt")
'xRpT.Database.SetDataSource
Case "Faktur Penjualan"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\faktur.rpt", 1)
Case "Defekta"
Set xRpT = xApp.OpenReport(PathDrive & "\laporan\defekta.rpt", 1)
'Set xRpT = xapp.OpenReport((pathdrive & "\laporan\faktur.rpt"),1
End Select


'YANG BARU DICOBA ternyata menjadikan ga perlu di refress
xRpT.DiscardSavedData

'xRpT.Database.Tab
'xRpT.Database.Parent.ReportAlerts.Parent.
'CRViewer1.ActivateView

If Not cobaLaporanGaError = True Then
If LaporanTerpilih = "Kartu Stok" Then
xRpT.Database.SetDataSource rsLaporan
xRpT.Database.Verify
'MsgBox rsLaporan.Fields(1)
xRpT.ReportTitle = "Per " & tanggalLaporan
End If

If LaporanTerpilih = "Rekapitulasi Omzet Harian" Then
'xRpT.Database.SetDataSource rsLaporan
'xRpT.Database.Verify
'xRpT.Database.LogOffServer
'For Each xDbf In xRpT.Database.Tables
'xDbf.Location = "D:\Sima 08 jl\Laporan\Februari 11"
'Next
'MsgBox rsLaporan.Fields(1)
Set xDbf = xRpT.Database.Tables(1)
xDbf.SetDataSource rsLaporan
xRpT.Database.Verify
End If
If LaporanTerpilih = "sembeli" Then
xRpT.Database.SetDataSource rsLaporan
xRpT.Database.Verify
End If
' Rekapitulasi Cara Bayar
If LaporanTerpilih = "Rekapitulasi Cara Bayar" Then
xRpT.Database.SetDataSource rsLaporan
rsLaporan.MoveFirst
xRpT.Database.Verify
'MsgBox rsLaporan.Fields(1)
xRpT.ReportTitle = "Rekapitulasi Hasil Penjualan " & tanggalLaporan
End If
If LaporanTerpilih = "Laporan Kondisi Barang" Then
xRpT.Database.SetDataSource rsLaporan
xRpT.Database.Verify
'MsgBox rsLaporan.Fields(1)
xRpT.ReportTitle = "Per " & tanggalLaporan
End If
End If
cobaLaporanGaError = False
'MsgBox "Laporan Telah Dibuat"
'MsgBox xRpT.Database.Tables
' CRViewer1.Refresh
 '   CRViewer1.ReportSource = CRXReport
  '  CRViewer1.ViewReport

'CRViewer1.r
'crviewer1.Refresh
'CRViewer1_RefreshButtonClicked
'CRViewer1.ReportSource


CRViewer1.ReportSource = xRpT
CRViewer1.ViewReport
'CRViewer1.Refresh

Exit Sub
Keluar:
'MsgBox "Ada Masalah Dengan Laporan !!", vblnformation,
End Sub
'Private Sub Form_ResizeO()
'CRViewerl.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
'End Sub

'End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
'CRViewer1.Refresh

End Sub

