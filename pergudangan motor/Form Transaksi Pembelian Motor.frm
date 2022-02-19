VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   Caption         =   "Transaksi Pembelian Motor"
   ClientHeight    =   5745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   5745
   ScaleWidth      =   7860
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form Transaksi Pembelian Motor.frx":0000
      Height          =   1455
      Left            =   240
      TabIndex        =   15
      Top             =   4200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2566
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   17
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "IdSupplier"
         Caption         =   "IdSupplier"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NamaSupplier"
         Caption         =   "NamaSupplier"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "NoTelp"
         Caption         =   "NoTelp"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "TanggalBeli"
         Caption         =   "TanggalBeli"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "HargaBeli"
         Caption         =   "HargaBeli"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14345
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1395,213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1440
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcPembelian 
      Height          =   375
      Left            =   3960
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=koneksiodbc"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "koneksiodbc"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "TransaksiPembelian"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<- Kembali     "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   12
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdSimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   11
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton cmdTambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   975
   End
   Begin VB.TextBox txtHarga 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txtNoTelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txtNama 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1440
      Width           =   1935
   End
   Begin VB.TextBox txtID 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker dtpTanggal 
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   2400
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   142540801
      CurrentDate     =   44609
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   7920
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Label Label6 
      Caption         =   "Harga Beli"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Tanggal Beli"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "No Telp"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Nama Supplier"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "ID Supplier"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   8520
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "TRANSAKSI PEMBELIAN MOTOR"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()

End Sub

Private Sub cmdHapus_Click()
    AdodcPembelian.Recordset.Delete
    bersih
    AdodcPembelian.Recordset.Update
    MsgBox "Data sudah dihapus"
    AdodcPembelian.Refresh
End Sub

Sub bersih()
    txtID = ""
    txtNama = ""
    txtNoTelp = ""
    txtHarga = ""
    
    cmdTambah.Enabled = True
    cmdSimpan.Enabled = False
    cmdHapus.Enabled = False
End Sub

Private Sub cmdSimpan_Click()
    
    AdodcPembelian.Recordset!IdSupplier = txtID.Text
    AdodcPembelian.Recordset!NamaSupplier = txtNama.Text
    AdodcPembelian.Recordset!NoTelp = txtNoTelp.Text
    AdodcPembelian.Recordset!TanggalBeli = dtpTanggal
    AdodcPembelian.Recordset!HargaBeli = txtHarga.Text
    AdodcPembelian.Recordset.Update
    
    MsgBox "Data telah diupdate"
    AdodcPembelian.Refresh
    Call bersih
End Sub

Private Sub cmdTambah_Click()
    If txtID.Text = "" And txtNama.Text = "" And txtNoTelp.Text = "" And txtHarga.Text = "" Then
        MsgBox "Data tidak boleh kosong"
    Else
        AdodcPembelian.Recordset.AddNew
        AdodcPembelian.Recordset!IdSupplier = txtID.Text
        AdodcPembelian.Recordset!NamaSupplier = txtNama.Text
        AdodcPembelian.Recordset!NoTelp = txtNoTelp.Text
        AdodcPembelian.Recordset!TanggalBeli = dtpTanggal
        AdodcPembelian.Recordset!HargaBeli = txtHarga.Text
        AdodcPembelian.Recordset.Update
        
        MsgBox "Data Disimpan"
        AdodcPembelian.Refresh
        Call bersih
     End If
End Sub

Private Sub Command4_Click()
    Form4.Show
    Me.Hide
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub DataGrid1_Click()
    txtID.Text = AdodcPembelian.Recordset!IdSupplier
    txtNama.Text = AdodcPembelian.Recordset!NamaSupplier
    txtNoTelp.Text = AdodcPembelian.Recordset!NoTelp
    dtpTanggal = AdodcPembelian.Recordset!TanggalBeli
    txtHarga.Text = AdodcPembelian.Recordset!HargaBeli
    
    cmdTambah.Enabled = False
    cmdSimpan.Enabled = True
    cmdHapus.Enabled = True
End Sub

Private Sub Form_Load()
    cmdTambah.Enabled = True
    cmdSimpan.Enabled = False
    cmdHapus.Enabled = False
End Sub

Private Sub txtHarga_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
    Else
    KeyAscii = 0
    End If
End Sub

Private Sub txtNoTelp_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
    Else
    KeyAscii = 0
    End If
End Sub
