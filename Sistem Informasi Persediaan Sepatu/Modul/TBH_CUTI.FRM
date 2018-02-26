VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tbh_cuti 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3690
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "S&impan"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   3120
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Cuti Pegawai"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4845
      Begin VB.TextBox ket 
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
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox No_Reg 
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
         MaxLength       =   10
         TabIndex        =   0
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox Lm_cuti 
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
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1920
         Width           =   735
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
         TabIndex        =   2
         Text            =   "NIP"
         Top             =   840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker Tgl_Akhir_Cuti 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1560
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   70385667
         CurrentDate     =   38037
      End
      Begin MSComCtl2.DTPicker Tgl_Awal_Cuti 
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
         Format          =   70385667
         CurrentDate     =   38037
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Keterangan:"
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   2400
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Awal Cuti :"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   1320
         Width           =   1380
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "No. Registrasi Cuti:"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIP:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Akhir Cuti:"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Width           =   1350
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Lama Cuti:"
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Width           =   750
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
            Picture         =   "tbh_cuti.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "tbh_cuti.frx":0E52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3375
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   5953
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
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No. Registrasi Cuti"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NIP"
         Object.Width           =   4305
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nama Pegawai"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tanggal Awal Cuti"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tanggal Akhir Cuti"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Lama Cuti"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Keterangan"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   360
      TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   15
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label9 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "tbh_cuti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

 
 
 
 
 
Private Sub cmdBatal_Click()
Call CleanControls
Me.No_Reg.SetFocus
 
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

 

Private Sub CmdSimpan_Click()
Dim cNIP
'If AddFlag Then
 
   If Me.NIP.Text <> "" And _
   Me.NIP.Text <> "" And _
      Me.NIP.Text <> "" Then
      
            cNo_Reg = Len(Trim(Me.No_Reg.Text))
            If cNo_Reg <> 4 Then
               MsgBox "No_Reg Harus 4 Karakter!", vbExclamation, "Peringatan"
               Me.NIP.SetFocus
               Exit Sub
            Else
               Call OpenTable("SELECT * FROM [Cuti] WHERE No_Reg='" & Me.No_Reg.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                             PesanSudahAda tbh_cuti
                             Me.No_Reg.SetFocus
                             SendKeys "{home}+{End}"
                             Exit Sub
                          End If
                        End With
                Call Simpan
                Frame3.Visible = True
                Timer1.Enabled = True
                cmdBatal_Click
                Call LoadDataToListView("SELECT * FROM [QryCuti]", rsRS, Me.lv1, 6)
                MsgBox "Data sudah tersipan!", vbInformation
                Me.No_Reg.SetFocus
            End If
     Else
         PesanKosong tbh_cuti
         Exit Sub
    End If
          
  
End Sub
 


 

 
  

 

 

Private Sub Form_Activate()
No_Reg.SetFocus
 
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
   End If
End Sub

Private Sub Form_Load()
 Call LoadDataToListView("SELECT * FROM [QryCuti]", rsRS, lv1, 6)
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



 

 

 

 
 

 
 

 

 
Private Sub NIP_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cNIP

 If KeyCode = 13 Then
     cNIP = Len(Me.NIP.Text)
            If cNIP <> 9 Then
               MsgBox "NIP Harus 9 Karakter!", vbExclamation, "Peringatan"
               Me.NIP.SetFocus
               SendKeys "{Home}+{End}"
               Exit Sub
            Else
               Call OpenTable("SELECT * FROM [Pegawai] WHERE NIP='" & Me.NIP.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                             PesanSudahAda tbh_cuti
                             Me.NIP.SetFocus
                             SendKeys "{Home}+{End}"
                             Exit Sub
                          End If
                        End With
            End If
 End If
End Sub

 

 
  

Private Sub Simpan()
 
 SQlSimpan = "INSERT INTO Cuti VALUES('" & Me.No_Reg.Text & "'," & _
               "'" & Me.NIP.Text & "'," & _
               "'" & Me.Tgl_Awal_Cuti.Value & "'," & _
               "'" & Me.Tgl_Akhir_Cuti.Value & "'," & _
               "'" & Me.Lm_cuti.Text & "'," & _
                "'" & Me.ket.Text & "');"



  Conn.Execute (SQlSimpan)
 
 
End Sub

 
 




Private Sub Tgl_Akhir_Cuti_Click()
Me.Lm_cuti.Text = Me.Tgl_Akhir_Cuti.Value - Me.Tgl_Awal_Cuti.Value
End Sub
