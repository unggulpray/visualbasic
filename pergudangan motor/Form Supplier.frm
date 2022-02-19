VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Supplier"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5820
   LinkTopic       =   "Form Supplier"
   ScaleHeight     =   4860
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   12
      Top             =   0
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Form Supplier.frx":0000
      Height          =   1215
      Left            =   240
      TabIndex        =   11
      Top             =   3120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2143
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "Id"
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
         DataField       =   "Nama"
         Caption         =   "Nama"
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
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739,906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc AdodcSupplier 
      Height          =   330
      Left            =   1080
      Top             =   3480
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      RecordSource    =   "Supplier"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Command8"
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   240
      Width           =   855
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
      Left            =   3600
      TabIndex        =   9
      Top             =   2400
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
      Left            =   2040
      TabIndex        =   8
      Top             =   2400
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
      Left            =   600
      TabIndex        =   7
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtNoTelp 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtNama 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtID 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   5760
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label3 
      Caption         =   "No Telepon"
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
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label4 
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
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5760
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "SUPPLIER"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "ID Suplier"
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
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bersih()
    txtID = ""
    txtNama = ""
    txtNoTelp = ""
    
    cmdTambah.Enabled = True
    cmdSimpan.Enabled = False
    cmdHapus.Enabled = False
End Sub

Private Sub cmdHapus_Click()
    AdodcSupplier.Recordset.Delete
    Call bersih
    AdodcSupplier.Recordset.Update
    MsgBox "Data Telah dihapus"
    AdodcSupplier.Refresh
End Sub

Private Sub cmdSimpan_Click()
    AdodcSupplier.Recordset!Id = txtID.Text
    AdodcSupplier.Recordset!Nama = txtNama.Text
    AdodcSupplier.Recordset!NoTelp = txtNoTelp.Text
    AdodcSupplier.Recordset.Update
    MsgBox "Data telah diupdate"
    AdodcSupplier.Refresh
End Sub

Private Sub cmdTambah_Click()
    If txtID.Text = "" Or txtNama.Text = "" Or txtNoTelp.Text = "" Then
        MsgBox "Data tidak boleh kosong"
    Else
        AdodcSupplier.Recordset.AddNew
        AdodcSupplier.Recordset!Id = txtID.Text
        AdodcSupplier.Recordset!Nama = txtNama.Text
        AdodcSupplier.Recordset!NoTelp = txtNoTelp.Text
        AdodcSupplier.Recordset.Update
        MsgBox "Data telah ditambahkan"
        AdodcSupplier.Refresh
    End If
    Call bersih
End Sub

Private Sub Command4_Click()
    Form4.Show
    Me.Hide
End Sub

Private Sub Command6_Click()
    Form4.Show
    Me.Hide
End Sub

Private Sub DataGrid1_Click()
    txtID = AdodcSupplier.Recordset!Id
    txtNama = AdodcSupplier.Recordset!Nama
    txtNoTelp = AdodcSupplier.Recordset!NoTelp

    cmdTambah.Enabled = False
    cmdSimpan.Enabled = True
    cmdHapus.Enabled = True
End Sub

Private Sub Form_Load()
    cmdTambah.Enabled = True
    cmdSimpan.Enabled = False
    cmdHapus.Enabled = False
End Sub

Private Sub txtNoTelp_KeyPress(KeyAscii As Integer)
    If (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = vbKeyBack Then
    Else
    KeyAscii = 0
    End If
End Sub

