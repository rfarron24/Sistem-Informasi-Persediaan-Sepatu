VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form lap_04 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5400
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "Laporan Barang Keluar"
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox kd_jen 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1095
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Keluar"
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   2520
         Width           =   1695
      End
      Begin VB.CommandButton cmdPrinter 
         BackColor       =   &H00008000&
         Caption         =   "Cetak Laporan"
         Height          =   375
         Left            =   960
         TabIndex        =   1
         Top             =   2520
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker tgl 
         Height          =   375
         Left            =   1440
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MM/yyyy"
         Format          =   19267587
         CurrentDate     =   38037
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Jenis Sepatu:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Kode Jenis:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Bulan:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   450
      End
   End
   Begin Crystal.CrystalReport Crpt1 
      Left            =   2640
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
   Begin Crystal.CrystalReport Crpt2 
      Left            =   1800
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
   End
End
Attribute VB_Name = "lap_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

 
                

 

Private Sub cmdClose_Click()
Unload Me

End Sub

 

Private Sub cmdPrinter_Click()
  
                       With Crpt2
                        
                        .DataFiles(0) = App.Path & "\Database\db.Mdb"
                        .ReportFileName = "D:\program_pipit\Laporan\lap_04.rpt"
                        .SelectionFormula = "month({pengeluaran_sepatu.Tgl_Keluar})= " & Month(Me.tgl.Value) & " and year({pengeluaran_sepatu.Tgl_Keluar})= " & Year(Me.tgl.Value) & " AND {jenis_sepatu.kd_jen}='" & Me.kd_jen.Text & "'"
                        .RetrieveDataFiles
                        .WindowState = crptMaximized

                        .Action = 2
                        
                        End With
                            
                  
 
End Sub

  



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
   Unload Me
   End If
End Sub

Private Sub Form_Load()
Call SetFormCenter(Me)
  Call Loadkd_jenToCombo("SELECT*FROM jenis_sepatu", rsRS, Me.kd_jen)


  End Sub

 

Private Sub clrTxt()
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


