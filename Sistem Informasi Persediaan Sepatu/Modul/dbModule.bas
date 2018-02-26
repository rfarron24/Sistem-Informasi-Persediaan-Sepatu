Attribute VB_Name = "dbModule"
 
Public Conn As New ADODB.Connection
Public rsRS As New ADODB.Recordset

 
 
Public AddFlag As Boolean
Public EditFlag As Boolean
Public Isitext As String
Public List As ListItem
Public I As Integer
Public CariItem
Public txt As Control
Public reply As String
Public StrSql As String
Public SQlSimpan As String
Public SQLHapus As String
Public SQLPerbaiki As String

Public Sub Connect()
    Set Conn = New ADODB.Connection
    Conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                          "Data Source=" & App.Path & "\Database\db.mdb"
    Conn.Open
End Sub
 

Public Sub LoadDataToListView(StrSql As String, rs As ADODB.Recordset, Grid As ListView, CountFields As Integer)
On Error Resume Next

Call OpenTable(StrSql, rs)
Grid.ListItems.Clear
Do While Not rs.EOF
   Set List = Grid.ListItems.Add(, , rs.Fields(0))
   For I = 1 To CountFields
      List.SubItems(I) = rs.Fields(I)
   Next I
   rs.MoveNext
Loop
End Sub

 
 Public Sub LoadDataToListViewxx(StrSql As String, rs As ADODB.Recordset, Grid As ListView, CountFields As Integer)
On Error Resume Next

Call OpenTable(StrSql, rs)
Grid.ListItems.Clear
Do While Not rs.EOF
   Set List = Grid.ListItems.Add(, , rs.Fields(0))
   For I = 1 To CountFields
      List.SubItems(I) = rs.Fields(I)
   Next I
   rs.MoveNext
Loop
End Sub

Public Sub SetFormCenter(frm As Form)
frm.Move (frmUtama.ScaleWidth \ 2) - (frm.Width \ 2), (frmUtama.ScaleHeight / 2) - (frm.Height / 2)
End Sub

Public Sub Loadkd_petugasToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub
 
 

 Public Sub Loadno_masukToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub


Public Sub Loadkd_sepatuToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub LoadNo_KeluarToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub Loadkd_jenToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

Public Sub Loadkd_pengirimanToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

 

Public Sub OpenTable(StrSql As String, rs As ADODB.Recordset)
    Set rs = New ADODB.Recordset
        If rs.State = adStateOpen Then Set rs = Nothing
        rs.Open StrSql, Conn, adOpenDynamic, adLockOptimistic
        
    
End Sub

Public Sub Loadkd_produkToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub

 
 

Public Sub Loadno_pesertaToCombo(StrSql As String, rs As ADODB.Recordset, Combo As ComboBox)
Call OpenTable(StrSql, rs)
Combo.Clear
Do While Not rs.EOF
   Combo.AddItem rs.Fields(0)
   rs.MoveNext
Loop
End Sub


 Public Sub PesanSudahAda(frm As Form)
 MsgBox "Data sudah ada!", vbCritical, "Data Suda Ada"
 End Sub
 Public Sub PesanKosong(frm As Form)
 MsgBox "Data tidak boleh kosong!", vbCritical, "Data Kosong"
  
 End Sub
 

 
 Public Sub PesanSimpan(frm As Form)
 MsgBox "Data sudah disimpan!", vbInformation, "Simpan Data"
 End Sub
  Public Sub PesanUpdate(frm As Form)
 MsgBox "Data sudah di-update!", vbInformation, "Update Data"
 End Sub
 
 Public Sub PesanHapus(frm As Form)
 MsgBox "Data sudah terhapus!", vbInformation, "Hapus Data"
 End Sub
 
 Public Sub IsiDataText1()
     Isitext = "AaBbCcDdEeFfGgHhIiJjKkLlMmNnOoPpQqRrSsTtUuVvWwXxYyZz.,"
 End Sub
Public Sub IsiDataText2()
     Isitext = "0123456789"
End Sub
Public Sub IsiDataText3()
     Isitext = "()-0123456789"
End Sub
 


