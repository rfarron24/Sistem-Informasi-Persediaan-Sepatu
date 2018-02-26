VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tbh_penggajian 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5325
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10350
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "S&impan"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   4800
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Penggajian Pegawai"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Width           =   4365
      Begin VB.TextBox Gaber 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   6
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox Pot_simpanan 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   5
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox T_beras 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   4
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox T_anak 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   3
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox T_suami_istri 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   2
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox Nm_peg 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   23
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox T_jabat 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox Gapok 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1560
         Width           =   1935
      End
      Begin VB.ComboBox NIP 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         TabIndex        =   0
         Text            =   "NIP"
         Top             =   480
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker Tgl_gaji 
         Height          =   375
         Left            =   1800
         TabIndex        =   1
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   55115779
         CurrentDate     =   38037
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Gaji Bersih:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   4200
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Potongan Simpanan:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   3840
         Width           =   1485
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Beras:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   3480
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Anak:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   3120
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Suami/Istri:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   2760
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jabatan:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai:"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tunjangan:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Gajian:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   1320
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIP:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gaji Pokok:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1680
         Width           =   825
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   23
      ImageHeight     =   23
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_penggajian.frx":0E52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5055
      Left            =   4560
      TabIndex        =   13
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nama Pegawai"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tanggal Gajian"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Gaji Pokok"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tunjangan Jabatan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Tunjangan  Suami/Istri"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Tunjangan Anak"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Tunjangan Beras"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Potongan Simpnanan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Gaji Bersih"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   15
         Left            =   2160
         Top             =   720
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2760
         TabIndex        =   18
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label9 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   16
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "tbh_penggajian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

 
 
 
 
 
Private Sub cmdBatal_Click()
Call CleanControls
Me.NIP.SetFocus
 
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

 

Private Sub CmdSimpan_Click()
 
 
   If Me.NIP.Text <> "" And _
   Me.NIP.Text <> "" And _
      Me.NIP.Text <> "" Then
       
                Call Simpan
                Frame3.Visible = True
                Timer1.Enabled = True
                cmdBatal_Click
                Call LoadDataToListView("SELECT * FROM [QryGaji]", rsRS, Me.lv1, 9)
                MsgBox "Data sudah tersipan!", vbInformation
                Me.NIP.SetFocus
             
     Else
         PesanKosong tbh_penggajian
         Exit Sub
    End If
          
  
End Sub
 


 

 
  

 

 

Private Sub Form_Activate()
Me.NIP.SetFocus
 
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
   End If
End Sub

Private Sub Form_Load()
 Call LoadDataToListView("SELECT * FROM [QryGaji]", rsRS, lv1, 9)
 Call LoadnipToCombo("SELECT*FROM [pegawai]", rsRS, Me.NIP)

 
  
EditFlag = False
AddFlag = False

End Sub

Private Sub CleanControls()
For Each txt In Me.Controls
If TypeOf txt Is TextBox Then
txt.Text = ""
ElseIf TypeOf txt Is ComboBox Then
txt.ListIndex = -1
End If
Next
 

End Sub

 

 
 
 

 
  

Private Sub Simpan()
 
 SQlSimpan = "INSERT INTO Gaji VALUES('" & Me.NIP.Text & "'," & _
               "'" & Me.Tgl_gaji.Value & "'," & _
               "'" & Me.Gapok.Text & "'," & _
               "'" & Me.T_jabat.Text & "'," & _
               "'" & Me.T_suami_istri.Text & "'," & _
               "'" & Me.T_anak.Text & "'," & _
               "'" & Me.T_beras.Text & "'," & _
               "'" & Me.Pot_simpanan.Text & "'," & _
               "'" & Me.Gaber.Text & "');"



  Conn.Execute (SQlSimpan)
 
 
End Sub

 
 





Private Sub Pot_simpanan_Change()
Me.Gaber.Text = (Val(Me.T_anak.Text) + Val(Me.T_beras.Text) + Val(Me.T_jabat.Text) - Val(Me.Pot_simpanan.Text))
End Sub

Private Sub NIP_Click()
 Call OpenTable("SELECT * FROM [qrypegawai] WHERE nip='" & Me.NIP.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                             Me.Nm_peg.Text = .Fields(1)
                              Me.T_jabat.Text = .Fields("bsr_tunj")
                              Me.Gapok.Text = .Fields("Gapok")
                           
                          End If
                        End With
End Sub
