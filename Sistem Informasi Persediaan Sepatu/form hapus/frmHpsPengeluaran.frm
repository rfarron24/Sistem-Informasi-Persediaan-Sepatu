VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHpsPengeluaran 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   12300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   5640
      TabIndex        =   16
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton cmdHapus 
      Caption         =   "&Hapus"
      Height          =   495
      Left            =   4440
      TabIndex        =   15
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   495
      Left            =   6840
      TabIndex        =   14
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "FORM HAPUS PENGELUARAN BARANG/SEPATU"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5805
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   6405
      Begin VB.ComboBox no_keluar 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker Tgl_Keluar 
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   19267587
         CurrentDate     =   38037
      End
      Begin MSComctlLib.ListView lv2 
         Height          =   2535
         Left            =   360
         TabIndex        =   7
         Top             =   3000
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4471
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Kode Produk"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Harga"
            Object.Width           =   4304
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Jumlah"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.TextBox kd_sepatu 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   2
         Top             =   2160
         Width           =   975
      End
      Begin VB.TextBox kd_jen 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox jen_sepatu 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox nm_sepatu 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Nama Sepatu:"
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Kode Sepatu:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Sepatu:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   1680
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode Jenis:"
         Height          =   195
         Left            =   360
         TabIndex        =   10
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal:"
         Height          =   195
         Left            =   3720
         TabIndex        =   9
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "No. Keluar:"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   795
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
            Picture         =   "frmHpsPengeluaran.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHpsPengeluaran.frx":0E52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   5775
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   10186
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No. Keluar"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tgl. Keluar"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Jenis Sepatu"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nama Sepatu"
         Object.Width           =   6068
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   360
      TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   22
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label9 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmHpsPengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


 
 
 
 
 

 

Private Sub cmdBatal_Click()

Call CleanControls

Call LoadDataToListView("SELECT * FROM [Qrypengeluaran_sepatu]", rsRS, lv1, 3)
Me.lv2.ListItems.Clear
Me.no_keluar.SetFocus

End Sub

 
Private Sub cmdHapus_Click()
If Me.no_keluar.ListIndex = -1 Then
Exit Sub

Else


 Call OpenTable("SELECT * FROM [pengeluaran_sepatu] WHERE No_Keluar='" & Me.no_keluar.Text & "'", rsRS)
        With rsRS
           If Not .EOF Then
              reply = MsgBox("Data Akan Dihapus?", vbQuestion + vbYesNo, "Konfirmasi")
                    If reply = vbYes Then
                       SQLHapus = "DELETE FROM [pengeluaran_sepatu] WHERE No_Keluar='" & Me.no_keluar.Text & "'"
                       Conn.Execute (SQLHapus)
                       
                        Call OpenTable("SELECT * FROM [detail_pengeluaran_sepatu] WHERE No_Keluar='" & Me.no_keluar.Text & "'", rsRS)
        With rsRS
           Do While Not .EOF
                SQLHapus = "DELETE FROM detail_pengeluaran_sepatu WHERE No_Keluar='" & Me.no_keluar.Text & "'"
                       Conn.Execute SQLHapus
            .MoveNext
           Loop
         End With
                       
                       
                       
                       cmdBatal_Click
                       MsgBox "Data Dihapus!", vbInformation, "Hapus Data"
                     End If
             End If
        End With










         
         
         
         
         
 
End If
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

 
 
 

 

 

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Me.no_keluar.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
   End If
End Sub

Private Sub Form_Load()
Call LoadDataToListView("SELECT * FROM [Qrypengeluaran_sepatu]", rsRS, lv1, 3)
  Call LoadNo_KeluarToCombo("SELECT*FROM pengeluaran_sepatu", rsRS, Me.no_keluar)
 Me.Tgl_Keluar.Value = Date
 Call SetFormCenter(Me)

 
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



 

 

 

 
  

 

 

 

 

 
 

Private Sub kd_jen_Change()
Call OpenTable("SELECT * FROM jenis_sepatu WHERE kd_jen='" & Me.kd_jen.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                      
                                Me.jen_sepatu.Text = .Fields(1)
                       
                          End If
                        End With
End Sub
 

 

Private Sub kd_sepatu_Change()
Call OpenTable("SELECT * FROM sepatu WHERE kd_sepatu='" & Me.kd_sepatu.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                      
                                Me.nm_sepatu.Text = .Fields(1)
                       
                          End If
                        End With
End Sub

Private Sub lv1_Click()
On Error GoTo e:

Call OpenTable("SELECT * FROM pengeluaran_sepatu WHERE No_Keluar='" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'", rsRS)
                                          
                        With rsRS
                         If Not .EOF Then
                         
                                Me.no_keluar.Text = .Fields("No_Keluar")
                                Me.Tgl_Keluar.Value = .Fields("Tgl_Keluar")
                                Me.kd_jen.Text = .Fields("kd_jen")
                                Me.kd_sepatu.Text = .Fields("kd_sepatu")
                         End If
                        End With




 Call OpenTable("SELECT * FROM Qrydetail_pengeluaran_sepatu WHERE No_Keluar='" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'", rsRS)
                         Me.lv2.ListItems.Clear
                        
                        With rsRS
                          Do While Not .EOF
                    
                    Set j = Me.lv2.ListItems.Add(, , .Fields(0))
                    j.SubItems(1) = .Fields(1)
                    j.SubItems(2) = .Fields(2)
                 
                       
                        .MoveNext
                        Loop
                        End With
e:
Exit Sub
End Sub

 

Private Sub No_Keluar_Click()
Call OpenTable("SELECT * FROM pengeluaran_sepatu WHERE No_Keluar='" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'", rsRS)
                                          
                        With rsRS
                         If Not .EOF Then
                         
                            
                                Me.Tgl_Keluar.Value = .Fields("Tgl_Keluar")
                                Me.kd_jen.Text = .Fields("kd_jen")
                                Me.kd_sepatu.Text = .Fields("kd_sepatu")
                         End If
                        End With




 Call OpenTable("SELECT * FROM Qrydetail_pengeluaran_sepatu WHERE No_Keluar='" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'", rsRS)
                         Me.lv2.ListItems.Clear
                        
                        With rsRS
                          Do While Not .EOF
                    
                    Set j = Me.lv2.ListItems.Add(, , .Fields(0))
                    j.SubItems(1) = .Fields(1)
                    j.SubItems(2) = .Fields(2)
                 
                       
                        .MoveNext
                        Loop
                        End With

End Sub

Private Sub Timer1_Timer()
Bar1.Value = Bar1.Value + 10
Me.Label9.Caption = Bar1.Value
If Bar1.Value = 100 Then
Timer1.Enabled = False

Frame3.Visible = False
Bar1.Value = 0
PesanSimpan frmUbahjenis_sepatu
End If
End Sub

 
 

Private Sub tampilkan()
With rsRS

  Me.kd_jen.Text = .Fields(0)
  Me.jen_sepatu.Text = .Fields(1)
   
   
  
End With
End Sub



 







