VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form tbh_pegawai 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6915
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   375
      Left            =   3840
      TabIndex        =   16
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "S&impan"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   6360
      Width           =   855
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   375
      Left            =   1200
      TabIndex        =   15
      Top             =   6360
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   4800
      TabIndex        =   24
      Top             =   240
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   10610
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
      NumItems        =   14
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
         Text            =   "Alamat"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Tempat Lahir"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Tgl. Lahir"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Jenis Kelamin"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Agama"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Jabatan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Gaji Pokok"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Jumlah Anak"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Pendidikan"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Tgl. Masuk"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Lama Kerja"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Biodata Pegawai"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6165
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4605
      Begin VB.TextBox Lm_kerja 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   13
         Top             =   5280
         Width           =   1335
      End
      Begin VB.ComboBox Pend 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   4560
         Width           =   1335
      End
      Begin VB.TextBox Jlh_anak 
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   10
         Top             =   4200
         Width           =   975
      End
      Begin VB.ComboBox Status 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ComboBox Kd_jab 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3120
         Width           =   1335
      End
      Begin VB.ComboBox Agama 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Almt 
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox T_lahir 
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1560
         Width           =   2415
      End
      Begin VB.ComboBox Jen_kel 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox Nm_peg 
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
         Left            =   1440
         MaxLength       =   30
         TabIndex        =   1
         Top             =   840
         Width           =   2415
      End
      Begin VB.ComboBox Kd_gol 
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
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox NIP 
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
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker Tgl_msk 
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   4920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   70189059
         CurrentDate     =   38037
      End
      Begin MSComCtl2.DTPicker Tgl_lahir 
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   1920
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   70189059
         CurrentDate     =   38037
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Lama Kerja:"
         Height          =   195
         Left            =   240
         TabIndex        =   37
         Top             =   5400
         Width           =   840
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Masuk:"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   5040
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Pendidikan:"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   4680
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Jumlah Anak:"
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   4320
         Width           =   960
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Status:"
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Kode Golongan:"
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   3480
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode Jabatan:"
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   3120
         Width           =   1035
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Agama:"
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   2760
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tgl. Lahir:"
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Kelamin:"
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2400
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Alamat:"
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tempat Lahir:"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Nama Pegawai:"
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "NIP:"
         Height          =   195
         Left            =   240
         TabIndex        =   18
         Top             =   600
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   0
      TabIndex        =   19
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
         TabIndex        =   20
         Top             =   240
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label12 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "0"
         Height          =   195
         Left            =   2760
         TabIndex        =   21
         Top             =   480
         Width           =   90
      End
   End
End
Attribute VB_Name = "tbh_pegawai"
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
Dim cNIP
'If AddFlag Then
 
   If Me.NIP.Text <> "" And _
   Me.NIP.Text <> "" And _
      Me.NIP.Text <> "" Then
      
            cNIP = Len(Trim(Me.NIP.Text))
            If cNIP <> 9 Then
               MsgBox "NIP Harus 9 Karakter!", vbExclamation, "Peringatan"
               Me.NIP.SetFocus
               Exit Sub
            Else
               Call OpenTable("SELECT * FROM [Pegawai] WHERE NIP='" & Me.NIP.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                             PesanSudahAda tbh_pegawai
                             Me.NIP.SetFocus
                             SendKeys "{home}+{End}"
                             Exit Sub
                          End If
                        End With
                Call Simpan
                Frame3.Visible = True
                Timer1.Enabled = True
                cmdBatal_Click
                Call LoadDataToListView("SELECT * FROM [QryPegawai]", rsRS, Me.ListView1, 13)
                MsgBox "Data sudah tersipan!", vbInformation
                Me.NIP.SetFocus
            End If
     Else
         PesanKosong tbh_pegawai
         Exit Sub
    End If
          
  
End Sub
 


 

 
  

 

 

Private Sub Form_Activate()
NIP.SetFocus
 
 End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
ElseIf KeyAscii = 13 Then
   SendKeys "{Tab}"
   End If
End Sub

Private Sub Form_Load()
 Call LoadDataToListView("SELECT * FROM [QryPegawai]", rsRS, ListView1, 13)
 Call LoadKd_jabToCombo("SELECT*FROM [jabatan]", rsRS, Me.Kd_jab)
 Call LoadKd_golToCombo("SELECT*FROM [Golongan]", rsRS, Me.Kd_gol)
 

With Me.Jen_kel
   .AddItem "Pria"
   .AddItem "Wanita"

End With

With Me.Pend
   .AddItem "SMA"
   .AddItem "D3"
   .AddItem "S1"
   .AddItem "S2"

End With



With Me.Status
   .AddItem "K"
   .AddItem "B"

End With


With Me.Agama
   .AddItem "Islam"
   .AddItem "Protestan"
   .AddItem "Katolik"
   .AddItem "Budha"
   .AddItem "Hindu"

End With


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
                             PesanSudahAda tbh_pegawai
                             Me.NIP.SetFocus
                             SendKeys "{Home}+{End}"
                             Exit Sub
                          End If
                        End With
            End If
 End If
End Sub

 

Private Sub timer1_pegawaier()
Bar1.Value = Bar1.Value + 10
Me.Label10.Caption = Bar1.Value
If Bar1.Value = 100 Then
Timer1.Enabled = False

Frame3.Visible = False
Bar1.Value = 0
PesanSimpan tbh_pegawai
End If
End Sub

  

Private Sub Simpan()
 
 SQlSimpan = "INSERT INTO Pegawai VALUES('" & Me.NIP.Text & "'," & _
               "'" & Me.Nm_peg.Text & "'," & _
               "'" & Me.Almt.Text & "'," & _
                "'" & Me.T_lahir.Text & "'," & _
               "'" & Me.Tgl_lahir.Value & "'," & _
               "'" & Me.Jen_kel.Text & "'," & _
               "'" & Me.Agama.Text & "'," & _
                "'" & Me.Kd_jab.Text & "'," & _
               "'" & Me.Kd_gol.Text & "'," & _
               "'" & Me.Status.Text & "'," & _
               "'" & Me.Jlh_anak.Text & "'," & _
               "'" & Me.Pend.Text & "'," & _
               "'" & Me.Tgl_msk.Value & "'," & _
               "'" & Me.Lm_kerja.Text & "');"



  Conn.Execute (SQlSimpan)
 
 
End Sub

 
 
