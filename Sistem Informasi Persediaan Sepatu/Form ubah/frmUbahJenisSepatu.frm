VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUbahJenisSepatu 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4410
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdKeluar 
      Caption         =   "Kel&uar"
      Height          =   495
      Left            =   6600
      TabIndex        =   5
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdSimpan 
      Caption         =   "Ubah"
      Height          =   495
      Left            =   4080
      TabIndex        =   4
      Top             =   3720
      Width           =   975
   End
   Begin VB.CommandButton CmdBatal 
      Caption         =   "&Batal"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   3720
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "FORM UBAH JENIS SEPATU"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   0
      Width           =   5085
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
         Left            =   1680
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1080
         Width           =   2895
      End
      Begin VB.ComboBox kd_jen 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Kode Jenis:"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Sepatu:"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   1200
         Width           =   960
      End
   End
   Begin MSComctlLib.ListView lv1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      _ExtentX        =   8705
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
            Picture         =   "frmUbahJenisSepatu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":005E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":00BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":011A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0178
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":01D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0234
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0292
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":02F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":034E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":03AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":040A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0468
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":04C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0582
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":05E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":063E
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":069C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":06FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0758
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":07B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0872
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":08D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":092E
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":098C
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":09EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0A48
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0B04
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0B62
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0BC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0CDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0D38
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0D96
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0DF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUbahJenisSepatu.frx":0E52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   360
      TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   480
         Width           =   90
      End
      Begin VB.Label Label9 
         Caption         =   "Saving New Data"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label8 
         Caption         =   "%"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmUbahJenisSepatu"
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
   
    
   If Me.jen_sepatu.Text <> "" Then
      
       Call Perbaiki
            cmdBatal_Click
   Call LoadDataToListView("SELECT * FROM [jenis_sepatu]", rsRS, lv1, 1)
        
            Frame3.Visible = True
            Timer1.Enabled = True
   Else
            PesanKosong frmUbahJenisSepatu
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

Private Sub Form_Load()
Call LoadDataToListView("SELECT * FROM [jenis_sepatu]", rsRS, lv1, 1)
 Call Loadkd_jenToCombo("SELECT*FROM jenis_sepatu", rsRS, Me.kd_jen)
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



 

 

 

 
  

 

 

 

 

 
 

Private Sub kd_jen_Click()
Call OpenTable("SELECT * FROM jenis_sepatu WHERE kd_jen='" & Me.kd_jen.Text & "'", rsRS)
                        With rsRS
                          If Not .EOF Then
                      
                                Me.jen_sepatu.Text = .Fields(1)
                       
                          End If
                        End With
End Sub

Private Sub lv1_Click()
On Error GoTo e:
Call OpenTable("SELECT * FROM [jenis_sepatu] WHERE kd_jen='" & lv1.ListItems.Item(lv1.SelectedItem.Index).Text & "'", rsRS)
         
        With rsRS
           If Not .EOF Then
              Call tampilkan
              
            End If
         End With
e:
Exit Sub
End Sub

Private Sub Timer1_Timer()
Bar1.Value = Bar1.Value + 10
Me.Label9.Caption = Bar1.Value
If Bar1.Value = 100 Then
Timer1.Enabled = False

Frame3.Visible = False
Bar1.Value = 0
PesanSimpan frmUbahJenisSepatu
End If
End Sub

 
 


 Sub Perbaiki()
      SQLPerbaiki = "UPDATE [jenis_sepatu] SET jen_sepatu ='" & Me.jen_sepatu.Text & "' WHERE [kd_jen]='" & Me.kd_jen.Text & "'"
                     
      Conn.Execute (SQLPerbaiki)
End Sub

 


Private Sub tampilkan()
With rsRS

  Me.kd_jen.Text = .Fields(0)
  Me.jen_sepatu.Text = .Fields(1)
   
   
  
End With
End Sub



 




