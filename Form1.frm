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
