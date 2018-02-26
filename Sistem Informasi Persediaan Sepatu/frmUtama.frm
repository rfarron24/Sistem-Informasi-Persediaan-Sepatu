VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.MDIForm frmUtama 
   BackColor       =   &H00004080&
   Caption         =   $"frmUtama.frx":0000
   ClientHeight    =   6345
   ClientLeft      =   165
   ClientTop       =   210
   ClientWidth     =   10050
   LinkTopic       =   "MDIForm1"
   MousePointer    =   1  'Arrow
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8640
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      Height          =   135
      Left            =   0
      ScaleHeight     =   75
      ScaleWidth      =   9990
      TabIndex        =   1
      Top             =   0
      Width           =   10050
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5970
      Width           =   10050
      _ExtentX        =   17727
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8017
            MinWidth        =   8017
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "12:08 PM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "4/11/07"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   45050
            MinWidth        =   45050
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9360
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   25
      ImageHeight     =   25
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":00A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0101
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":015F
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":01BD
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":021B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0279
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":02D7
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0335
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0393
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":03F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":044F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":04AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":050B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0569
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":05C7
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0625
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0683
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":06E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":073F
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":079D
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":07FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0859
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":08B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUtama.frx":0915
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Crystal.CrystalReport Crpt2 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin VB.Menu mnuinput 
      Caption         =   "&File"
      Begin VB.Menu mnu01 
         Caption         =   "Barang/ Sepatu "
      End
      Begin VB.Menu mnu02 
         Caption         =   "Jenis Sepatu"
      End
      Begin VB.Menu mnu03 
         Caption         =   "Produk"
      End
      Begin VB.Menu mnu04 
         Caption         =   "Barang/ Sepatu Masuk"
      End
      Begin VB.Menu mnu05 
         Caption         =   "Pengeluaran Sepatu"
      End
   End
   Begin VB.Menu mnuedit 
      Caption         =   "Edit"
      Begin VB.Menu mnu001 
         Caption         =   "Ubah"
         Begin VB.Menu mnu002 
            Caption         =   "Barang/ Sepatu "
         End
         Begin VB.Menu mnu003 
            Caption         =   "Jenis Sepatu"
         End
         Begin VB.Menu mnu004 
            Caption         =   "Produk"
         End
         Begin VB.Menu mnu005 
            Caption         =   "Barang/ Sepatu Masuk"
         End
         Begin VB.Menu mnu00523 
            Caption         =   "Pengeluaran Sepatu"
         End
      End
      Begin VB.Menu mnu007 
         Caption         =   "Hapus"
         Begin VB.Menu mnu0010 
            Caption         =   "Barang/ Sepatu "
         End
         Begin VB.Menu mnu0011 
            Caption         =   "Jenis Sepatu"
         End
         Begin VB.Menu mnu0012 
            Caption         =   "Produk"
         End
         Begin VB.Menu mnu0013 
            Caption         =   "Barang/ Sepatu Masuk"
         End
         Begin VB.Menu mnu0014 
            Caption         =   "Pengeluaran Sepatu"
         End
      End
   End
   Begin VB.Menu mnulap 
      Caption         =   "Laporan"
      Begin VB.Menu a1 
         Caption         =   "Laporan Barang (Sepatu)  Masuk/ Tanggal/ Jenis Sepatu"
      End
      Begin VB.Menu a2 
         Caption         =   "Laporan Barang (Sepatu)  Masuk/ Bulan/ Jenis Sepatu"
      End
      Begin VB.Menu sdbg 
         Caption         =   "Laporan Pengeluaran Barang (Sepatu)/Tanggal/ Jenis Sepatu"
      End
      Begin VB.Menu hbs 
         Caption         =   "Laporan Pengeluaran Barang (Sepatu)/ Bulan/ Jenis Sepatu"
      End
      Begin VB.Menu sdfg 
         Caption         =   "Laporan Persediaan Barang (Sepatu)/ Bulan/ Jenis Sepatu"
      End
   End
   Begin VB.Menu qsys 
      Caption         =   "Keluar"
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub a1_Click()
Lap_01.Show
End Sub

Private Sub a2_Click()
lap_02.Show
End Sub

Private Sub hbs_Click()
lap_04.Show
End Sub

Private Sub MDIForm_DblClick()
End
End Sub

Private Sub MDIForm_Load()
 
        
   Connect

StatusBar1.Panels(4).Text = " System is Ready "


End Sub

Private Sub mnu0010_Click()
frmHpsSepatu.Show
End Sub

Private Sub mnu0011_Click()
frmHpsJenisSepatu.Show
End Sub

Private Sub mnu0012_Click()
frmHpsProduk.Show
End Sub

Private Sub mnu0013_Click()
frmHpsBarangMasuk.Show
End Sub

Private Sub mnu0014_Click()
frmHpsPengeluaran.Show
End Sub

Private Sub mnu002_Click()
frmUbahSepatu.Show
End Sub

Private Sub mnu003_Click()
frmUbahJenisSepatu.Show
End Sub

Private Sub mnu004_Click()
frmUbahProduk.Show
End Sub

Private Sub mnu005_Click()
frmUbahBrgMasuk.Show
End Sub

Private Sub mnu00523_Click()
frmUbahBarangKeluar.Show
End Sub

Private Sub mnu01_Click()
frmSepatu.Show
End Sub

Private Sub mnu02_Click()
frmJenisSepatu.Show
End Sub

Private Sub mnu03_Click()
frmProduk.Show
End Sub

Private Sub mnu04_Click()
frmBarangMasuk.Show
End Sub

Private Sub mnu05_Click()
frmPengeluaranSepatu.Show
End Sub

Private Sub qsys_Click()
End
End Sub

Private Sub sdbg_Click()
Lap_03.Show
End Sub

Private Sub sdfg_Click()
lap_05.Show
End Sub

Private Sub Timer1_Timer()
Me.Caption = Right$(Me.Caption, Len(Me.Caption) - 1) + Left$(Me.Caption, 1)
StatusBar1.Panels(1) = Format(Date, "dd mmmm yyyy")
StatusBar1.Panels(2) = Format(Time, "hh:mm:ss")
End Sub
