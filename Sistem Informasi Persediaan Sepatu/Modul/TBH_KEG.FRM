VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tbh_keg 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4260
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10290
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4260
   ScaleWidth      =   10290
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "S&impan"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   3720
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3975
      Left            =   5160
      TabIndex        =   8
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7011
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
         Text            =   "Kode Kegiatan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tanggal Kegiatan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NIP"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nama Kegiatan"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Kegiatan Pegawai"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3525
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   4845
      Begin VB.TextBox Nm_keg 
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
         Top             =   1920
         Width           =   2895
      End
      Begin VB.TextBox Kd_keg 
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
         MaxLength       =   4
         TabIndex        =   0
         Top             =   840
         Width           =   735
      End
      Begin VB.ComboBox nip 
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
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1560
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker Tgl_keg 
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
         Format          =   70189059
         CurrentDate     =   38037
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode Kegiatan:"
         Height          =   195
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tanggal Kegiatan::"
         Height          =   195
         Left            =   240
         TabIndex        =   16
         Top             =   1200
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "NIP:"
         Height          =   195
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nama Kegiatan:"
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   9
      Top             =   0
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
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2760
         TabIndex        =   13
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label11 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label12 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   11
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "tbh_keg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 
 

 
 
 
 
 
Private Sub cmdBatal_Click()
Call CleanControls
Me.Kd_keg.SetFocus
 
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
      
            cKd_keg = Len(Trim(Me.Kd_keg.Text))
            If cKd_keg <> 4 Then
               MsgBox "Kd_keg Harus 4 Karakter!", vbExclamation, "Peringatan"
               Me.NIP.SetFocus
               Exit Sub
            Else
               Call OpenTable("SELECT * FROM [Kegiatan] WHERE Kd_keg='" & Me.Kd_keg.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                             PesanSudahAda tbh_keg
                             Me.Kd_keg.SetFocus
                             SendKeys "{home}+{End}"
                             Exit Sub
                          End If
                        End With
                Call Simpan
                Frame3.Visible = True
                Timer1.Enabled = True
                cmdBatal_Click
                Call LoadDataToListView("SELECT * FROM [QryKegiatan]", rsRS, Me.ListView1, 3)
                MsgBox "Data sudah tersipan!", vbInformation
                Me.Kd_keg.SetFocus
            End If
     Else
         PesanKosong tbh_keg
         Exit Sub
    End If
          
  
End Sub
 


 

 
  

 

 

Private Sub Form_Activate()
Kd_keg.SetFocus
 
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
   End If
End Sub

Private Sub Form_Load()
 Call LoadDataToListView("SELECT * FROM [QryKegiatan]", rsRS, ListView1, 3)
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
                             PesanSudahAda tbh_keg
                             Me.NIP.SetFocus
                             SendKeys "{Home}+{End}"
                             Exit Sub
                          End If
                        End With
            End If
 End If
End Sub

 

 
  

Private Sub Simpan()
 
 SQlSimpan = "INSERT INTO Kegiatan VALUES('" & Me.Kd_keg.Text & "'," & _
               "'" & Me.Tgl_keg.Value & "'," & _
               "'" & Me.NIP.Text & "'," & _
                "'" & Me.Nm_keg.Text & "');"



  Conn.Execute (SQlSimpan)
 
 
End Sub

 
 


