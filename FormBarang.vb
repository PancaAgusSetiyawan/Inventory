Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class FormBarang
    Private strSQL As String
    Private objDataTable As DataTable
    Private objReader As OleDbDataReader
    Private objAdapter As OleDbDataAdapter
    Private objDataset As DataSet
    Private myCon As OleDbConnection
    Private objCommand As OleDbCommand

    Private Sub LockText()
        txtHargaSatuan.ReadOnly = True
        txtKodeBarang.ReadOnly = True
        txtNamaBarang.ReadOnly = True
        txtSatuanBarang.ReadOnly = True
        txtSpecBarang.ReadOnly = True
    End Sub

    Private Sub UnLockText()
        txtHargaSatuan.ReadOnly = False
        txtKodeBarang.ReadOnly = False
        txtNamaBarang.ReadOnly = False
        txtSatuanBarang.ReadOnly = False
        txtSpecBarang.ReadOnly = False
    End Sub

    Private Sub TextKosong()
        txtHargaSatuan.Text = ""
        txtKodeBarang.Text = ""
        txtNamaBarang.Text = ""
        txtSatuanBarang.Text = ""
        txtSpecBarang.Text = ""
    End Sub

    Private Sub FormBarang_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TextKosong()
        Call LockText()
        Call ListGrid()
        ButtonSave.Enabled = False
        ButtonCancel.Enabled = False
    End Sub

    Private Sub ButtonAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAdd.Click
        Call UnLockText()
        Call TextKosong()
        txtKodeBarang.Focus()
        ButtonAdd.Enabled = False
        ButtonEdit.Enabled = False
        ButtonDelete.Enabled = False
        ButtonExit.Enabled = False
        ButtonSave.Enabled = True
        ButtonCancel.Enabled = True
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        If txtKodeBarang.Text = "" Then
            MsgBox("Kode barang tidak boleh kosong")
            txtKodeBarang.Focus()
        ElseIf txtNamaBarang.Text = "" Then
            MsgBox("Nama barang tidak boleh kosong")
            txtNamaBarang.Focus()
        ElseIf txtSpecBarang.Text = "" Then
            MsgBox("Spec barang tidak boleh kosong")
            txtSpecBarang.Focus()
        ElseIf txtSatuanBarang.Text = "" Then
            MsgBox("Satuan barang tidak boleh kosong")
            txtSatuanBarang.Focus()
        ElseIf txtHargaSatuan.Text = "" Then
            MsgBox("Harga satuan tidak boleh kosong")
            txtHargaSatuan.Focus()
        Else
            myCon = New OleDbConnection(strCon)
            Try
                myCon.Open()
                strSQL = "SELECT * FROM TBL_BARANG WHERE KD_BRG = '" & Trim(txtKodeBarang.Text) & "'"
                objCommand = New OleDbCommand(strSQL, myCon)
                objReader = objCommand.ExecuteReader(CommandBehavior.Default)
                If objReader.HasRows Then
                    MsgBox("Duplicate Data")
                Else
                    objCommand.Dispose()
                    strSQL = "INSERT INTO TBL_BARANG (KD_BRG,NM_BRG,SAT_BRG,SPEC_BRG,HRG_SAT) VALUES('" & txtKodeBarang.Text & "','" & txtNamaBarang.Text & "','" & txtSatuanBarang.Text & "','" & txtSpecBarang.Text & "','" & CDbl(txtHargaSatuan.Text) & "')"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    If objCommand.ExecuteNonQuery Then
                        MsgBox("Data telah di simpan")
                    Else
                        MsgBox("Data error di simpan")
                    End If
                    objReader.Close()
                End If
            Catch ex As Exception
                MsgBox("ERROR")
            Finally
                myCon.Close()
                objCommand = Nothing
                objReader = Nothing
                myCon = Nothing
                Call ListGrid()
                Call LockText()
                Call TextKosong()
                ButtonAdd.Enabled = True
                ButtonEdit.Enabled = True
                ButtonDelete.Enabled = True
                ButtonExit.Enabled = True
                ButtonSave.Enabled = False
                ButtonCancel.Enabled = False
            End Try
        End If
    End Sub

    Private Sub ListGrid()
        Try
            myCon = New OleDbConnection(strCon)
            objDataTable = New DataTable
            myCon.Open()
            strSQL = "SELECT * FROM TBL_BARANG"
            objCommand = New OleDbCommand(strSQL, myCon)
            objReader = objCommand.ExecuteReader(CommandBehavior.Default)
            objDataTable.Load(objReader)
            grdBarang.DataSource = objDataTable
            objCommand.Dispose()
            objReader.Close()
            objCommand = Nothing
            objReader = Nothing
            myCon = Nothing
        Catch ex As Exception
            grdBarang = Nothing
        End Try
    End Sub

    Private Sub txtHargaSatuan_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHargaSatuan.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) And e.KeyChar <> Chr(Asc(".")) Then
            e.Handled = True
        End If
    End Sub

    Private Sub grdBarang_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdBarang.Click
        Try
            txtKodeBarang.Text = Trim(grdBarang.SelectedCells(0).Value)
            txtNamaBarang.Text = grdBarang.SelectedCells(1).Value
            txtSatuanBarang.Text = grdBarang.SelectedCells(2).Value
            txtSpecBarang.Text = grdBarang.SelectedCells(3).Value
            txtHargaSatuan.Text = Format(grdBarang.SelectedCells(4).Value, "#,##0.00")
            Call GetStock(txtKodeBarang.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub ButtonCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonCancel.Click
        Call TextKosong()
        Call LockText()
        ButtonAdd.Enabled = True
        ButtonEdit.Enabled = True
        ButtonDelete.Enabled = True
        ButtonExit.Enabled = True
        ButtonSave.Enabled = False
        ButtonCancel.Enabled = False
        ButtonEdit.Text = "Edit"
    End Sub

    Private Sub ButtonDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDelete.Click
        If txtKodeBarang.Text = "" Or txtNamaBarang.Text = "" Or txtHargaSatuan.Text = "" Or txtSpecBarang.Text = "" Or txtSatuanBarang.Text = "" Then
            MsgBox("Data kosong")
        Else
            myCon = New OleDbConnection(strCon)
            Try
                myCon.Open()
                strSQL = "DELETE FROM TBL_BARANG WHERE KD_BRG = '" & Trim(txtKodeBarang.Text) & "'"
                objCommand = New OleDbCommand(strSQL, myCon)
                If objCommand.ExecuteNonQuery Then
                    MsgBox("Data telah dihapus")
                Else
                    MsgBox("Hapus gagal")
                End If
            Catch ex As Exception
                MsgBox("ERROR")
            Finally
                myCon.Close()
                objCommand = Nothing
                objReader = Nothing
                myCon = Nothing
                Call ListGrid()
                Call TextKosong()
                Call LockText()
            End Try
        End If
    End Sub

    Private Sub ButtonEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEdit.Click
        If txtKodeBarang.Text = "" Or txtNamaBarang.Text = "" Or txtHargaSatuan.Text = "" Or txtSpecBarang.Text = "" Or txtSatuanBarang.Text = "" Then
            MsgBox("Data kosong")
        Else
            If ButtonEdit.Text = "Edit" Then
                ButtonEdit.Text = "Update"
                ButtonAdd.Enabled = False
                ButtonEdit.Enabled = True
                ButtonSave.Enabled = False
                ButtonDelete.Enabled = False
                ButtonCancel.Enabled = True
                ButtonExit.Enabled = False
                Call UnLockText()
                txtKodeBarang.ReadOnly = True
                txtNamaBarang.Focus()
            Else
                myCon = New OleDbConnection(strCon)
                Try
                    myCon.Open()
                    strSQL = "UPDATE TBL_BARANG SET NM_BRG = '" & txtNamaBarang.Text & "', SAT_BRG = '" & txtSatuanBarang.Text & "', SPEC_BRG = '" & txtSpecBarang.Text & "', HRG_SAT = '" & CDbl(txtHargaSatuan.Text) & "' WHERE KD_BRG = '" & txtKodeBarang.Text & "'"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    If objCommand.ExecuteNonQuery Then
                        MsgBox("Data telah diupdate")
                    Else
                        MsgBox("Update gagal")
                    End If
                Catch ex As Exception
                    MsgBox("ERROR")
                Finally
                    myCon.Close()
                    objCommand = Nothing
                    objReader = Nothing
                    myCon = Nothing
                    Call ListGrid()
                    Call TextKosong()
                    Call LockText()
                End Try
                ButtonEdit.Text = "Edit"
                ButtonAdd.Enabled = True
                ButtonEdit.Enabled = True
                ButtonDelete.Enabled = True
                ButtonExit.Enabled = True
                ButtonSave.Enabled = False
                ButtonCancel.Enabled = False
            End If
        End If
    End Sub

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        Me.Close()
    End Sub

    Private Sub GetStock(ByVal strKode As String)
        Try
            myCon = New OleDbConnection(strCon)
            myCon.Open()
            strSQL = "SELECT QTY_IN,QTY_OUT,QTY_AKHIR FROM TBL_STOCK WHERE KD_BRG = '" & Trim(strKode) & "'"
            objCommand = New OleDbCommand(strSQL, myCon)
            objReader = objCommand.ExecuteReader(CommandBehavior.Default)
            If objReader.HasRows Then
                objReader.Read()
                txtIn.Text = Format(objReader(0), "#,##0.00")
                txtOut.Text = Format(objReader(1), "#,##0.00")
                txtAkhir.Text = Format(objReader(2), "#,##0.00")
            Else
                txtIn.Text = 0
                txtOut.Text = 0
                txtAkhir.Text = 0
            End If
            objReader.Close()
            myCon.Close()
        Catch ex As Exception

        End Try
    End Sub
End Class