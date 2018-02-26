VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmJenisSepatu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4290
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11130
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   11130
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "FORM JENIS SEPATU"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3285
      Left            =   5760
      TabIndex        =   5
      Top             =   120
      Width           =   5205
      Begin VB.TextBox kd_jen 
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
         Left            =   1560
         MaxLength       =   2
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox jen_sepatu 
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   1
         Top             =   1200
         Width           =   2895
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Sepatu:"
         Height          =   195
         Left            =   360
         TabIndex        =   7
         Top             =   1320
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Kode Jenis:"
         Height          =   195
         Left            =   360
         TabIndex        =   6
         Top             =   840
         Width           =   825
      End
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   495
      Left            =   6120
      TabIndex        =   4
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "S&impan"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   3600
      Width           =   1095
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   3600
      Width           =   1095
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
            Picture         =   "frmJenisSepatu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmJenisSepatu.frx":0E52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   5741
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Kode Jenis Sepatu"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Jenis Sepatu"
         Object.Width           =   11360
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   360
      TabIndex        =   9
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
         TabIndex        =   10
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label9 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2760
         TabIndex        =   11
         Top             =   480
         Width           =   90
      End
   End
End
Attribute VB_Name = "frmJenisSepatu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

 

 
 
 
 
 
Private Sub cmdBatal_Click()
Call CleanControls
Call LoadDataToListView("SELECT * FROM [jenis_sepatu]", rsRS, lv1, 1)
Me.kd_jen.SetFocus
 
End Sub

Private Sub CmdKeluar_Click()
Unload Me
End Sub

 

Private Sub CmdSimpan_Click()
  
  
   If Me.kd_jen.Text <> "" And _
      Me.jen_sepatu.Text <> "" Then
      
            
             ckd_jen = Len(Me.kd_jen.Text)
            If ckd_jen <> 2 Then
               MsgBox "kd_jen Harus 2 Karakter!", vbExclamation, "Peringatan"
               Me.kd_jen.SetFocus
               SendKeys "{Home}+{End}"
               Exit Sub
            Else
            
              
            
                Call simpan
                Frame3.Visible = True
                Timer1.Enabled = True
                cmdBatal_Click
                MsgBox "Data sudah tersimpan!", vbExclamation, "Simpan Data"
              End If
     Else
         PesanKosong frmJenisSepatu
         Exit Sub
    End If
          
  
End Sub
 


 

 
  

 

 

Private Sub Form_Activate()
Me.kd_jen.SetFocus
 
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
   End If
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



  
 
Private Sub simpan()
 
 SQlSimpan = "INSERT INTO jenis_sepatu VALUES('" & Me.kd_jen.Text & "'," & _
               "'" & Me.jen_sepatu.Text & "');"
 
  Conn.Execute (SQlSimpan)
 
 
End Sub
 

Private Sub Form_Load()
Call LoadDataToListView("SELECT * FROM [jenis_sepatu]", rsRS, lv1, 1)
 Call SetFormCenter(Me)

End Sub


Private Sub kd_jen_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     ckd_jen = Len(Me.kd_jen.Text)
            If ckd_jen <> 2 Then
               MsgBox "kd_jen Harus 2 Karakter!", vbExclamation, "Peringatan"
               Me.kd_jen.SetFocus
               SendKeys "{Home}+{End}"
               Exit Sub
            Else
            
               Call OpenTable("SELECT * FROM [jenis_sepatu] WHERE kd_jen='" & Me.kd_jen.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                             PesanSudahAda frmJenisSepatu
                             Me.kd_jen.SetFocus
                             SendKeys "{Home}+{End}"
                             Exit Sub
                          End If
                        End With
            End If
 End If
End Sub


