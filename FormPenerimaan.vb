Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class FormPenerimaan
    Private strSQL As String
    Private objDataTable As DataTable
    Private objReader As OleDbDataReader
    Private objAdapter As OleDbDataAdapter
    Private objDataset As DataSet
    Private myCon As OleDbConnection
    Private objCommand As OleDbCommand

    Private blnStock As Boolean
    Private dblNewIn As Double
    Private dblNewAkhir As Double
    Private tmpNewIn As Double
    Private tmpNewAkhir As Double
    Private tmpQty As Double

    Private Sub LockText()
        txtNomor.ReadOnly = True
        cboBarang.Enabled = False
        txtQty.ReadOnly = True
    End Sub

    Private Sub UnLockText()
        txtNomor.ReadOnly = False
        txtQty.ReadOnly = False
        cboBarang.Enabled = True
    End Sub

    Private Sub TextKosong()
        txtHargaSatuan.Text = ""
        cboBarang.Text = ""
        txtNamaBarang.Text = ""
        txtSatuanBarang.Text = ""
        txtSpecBarang.Text = ""
        txtQty.Text = ""
        txtTgl.Text = ""
        txtNomor.Text = ""
        txtIn.Text = ""
        txtOut.Text = ""
        txtAkhir.Text = ""
    End Sub

    Private Sub PopulateBarang()
        Try
            cboBarang.Items.Clear()
            myCon = New OleDbConnection(strCon)
            objDataTable = New DataTable
            myCon.Open()
            strSQL = "SELECT KD_BRG FROM TBL_BARANG"
            objCommand = New OleDbCommand(strSQL, myCon)
            objReader = objCommand.ExecuteReader(CommandBehavior.Default)
            If objReader.HasRows Then
                While objReader.Read
                    cboBarang.Items.Add(objReader(0))
                End While
            End If
            objCommand.Dispose()
            objReader.Close()
            myCon.Close()
            objCommand = Nothing
            objReader = Nothing
            myCon = Nothing
        Catch ex As Exception
            grdBarang = Nothing
        End Try
    End Sub

    Private Sub ListGrid()
        Try
            myCon = New OleDbConnection(strCon)
            objDataTable = New DataTable
            myCon.Open()
            strSQL = "SELECT * FROM TBL_PENERIMAAN"
            objCommand = New OleDbCommand(strSQL, myCon)
            objReader = objCommand.ExecuteReader(CommandBehavior.Default)
            objDataTable.Load(objReader)
            grdBarang.DataSource = objDataTable
            objCommand.Dispose()
            objReader.Close()
            myCon.Close()
            objCommand = Nothing
            objReader = Nothing
            myCon = Nothing
        Catch ex As Exception
            grdBarang = Nothing
        End Try
    End Sub

    Private Sub FormPenerimaan_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call TextKosong()
        Call LockText()
        Call ListGrid()
        Call PopulateBarang()
        ButtonSave.Enabled = False
        ButtonCancel.Enabled = False
    End Sub

    Private Sub ButtonAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonAdd.Click
        Call UnLockText()
        Call TextKosong()
        txtTgl.Text = Now.Date
        txtNomor.Focus()
        ButtonAdd.Enabled = False
        ButtonEdit.Enabled = False
        ButtonDelete.Enabled = False
        ButtonExit.Enabled = False
        ButtonSave.Enabled = True
        ButtonCancel.Enabled = True
    End Sub

    Private Sub ButtonSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSave.Click
        If txtNomor.Text = "" Then
            MsgBox("Nomor penerimaan tidak boleh kosong")
            txtNomor.Focus()
        ElseIf cboBarang.Text = "" Then
            MsgBox("Kode barang tidak boleh kosong")
            cboBarang.Focus()
        ElseIf txtQty.Text = "" Then
            MsgBox("Quantity tidak boleh kosong")
            txtQty.Focus()
        Else
            myCon = New OleDbConnection(strCon)
            Try
                myCon.Open()
                strSQL = "SELECT * FROM TBL_PENERIMAAN WHERE NO_PENERIMAAN = '" & Trim(txtNomor.Text) & "'"
                objCommand = New OleDbCommand(strSQL, myCon)
                objReader = objCommand.ExecuteReader(CommandBehavior.Default)
                If objReader.HasRows Then
                    MsgBox("Duplicate Data")
                Else
                    objCommand.Dispose()
                    strSQL = "INSERT INTO TBL_PENERIMAAN (NO_PENERIMAAN,TGL_TERIMA,KD_BRG,QTY) VALUES('" & txtNomor.Text & "','" & txtTgl.Text & "','" & cboBarang.Text & "','" & CDbl(txtQty.Text) & "')"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    If objCommand.ExecuteNonQuery Then
                        MsgBox("Data telah di simpan")
                    Else
                        MsgBox("Data error di simpan")
                    End If
                End If
                objReader.Close()
                dblNewAkhir = CDbl(txtAkhir.Text) + CDbl(txtQty.Text) ' pendklarasian
                dblNewIn = CDbl(txtIn.Text) + CDbl(txtQty.Text) ' pendklarasian
                If blnStock = True Then
                    strSQL = "UPDATE TBL_STOCK SET QTY_IN = " & dblNewIn & ",QTY_AKHIR = " & dblNewAkhir & " WHERE KD_BRG ='" & cboBarang.Text & "'"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    objCommand.ExecuteNonQuery()
                Else
                    strSQL = "INSERT INTO TBL_STOCK (KD_BRG,QTY_IN,QTY_OUT,QTY_AKHIR) VALUES('" & cboBarang.Text & "'," & CDbl(txtQty.Text) & ",0,'" & CDbl(txtQty.Text) & "')"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    objCommand.ExecuteNonQuery()
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

    Private Sub ButtonEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEdit.Click
        If txtNomor.Text = "" Or txtQty.Text = "" Or txtTgl.Text = "" Or cboBarang.Text = "" Then
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
                tmpQty = CDbl(txtQty.Text)
                Call UnLockText()
                txtNomor.ReadOnly = True
                txtQty.Focus()
            Else
                myCon = New OleDbConnection(strCon)
                Try
                    myCon.Open()
                    strSQL = "UPDATE TBL_PENERIMAAN SET KD_BRG = '" & cboBarang.Text & "', QTY = '" & CDbl(txtQty.Text) & "' WHERE NO_PENERIMAAN = '" & txtNomor.Text & "'"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    If objCommand.ExecuteNonQuery Then
                        MsgBox("Data telah diupdate")
                    Else
                        MsgBox("Update gagal")
                    End If
                    dblNewAkhir = (CDbl(txtAkhir.Text) + CDbl(txtQty.Text)) - tmpQty
                    dblNewIn = (CDbl(txtIn.Text) + CDbl(txtQty.Text)) - tmpQty
                    strSQL = "UPDATE TBL_STOCK SET QTY_IN = " & dblNewIn & ",QTY_AKHIR = " & dblNewAkhir & " WHERE KD_BRG ='" & cboBarang.Text & "'"
                    objCommand = New OleDbCommand(strSQL, myCon)
                    objCommand.ExecuteNonQuery()
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

    Private Sub ButtonDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonDelete.Click
        If txtNomor.Text = "" Or txtQty.Text = "" Or txtTgl.Text = "" Or cboBarang.Text = "" Then
            MsgBox("Data kosong")
        Else
            myCon = New OleDbConnection(strCon)
            Try
                myCon.Open()
                strSQL = "DELETE FROM TBL_PENERIMAAN WHERE NO_PENERIMAAN = '" & Trim(txtNomor.Text) & "'"
                objCommand = New OleDbCommand(strSQL, myCon)
                If objCommand.ExecuteNonQuery Then
                    MsgBox("Data telah dihapus")
                Else
                    MsgBox("Hapus gagal")
                End If
                dblNewAkhir = CDbl(txtAkhir.Text) - CDbl(txtQty.Text)
                dblNewIn = CDbl(txtIn.Text) - CDbl(txtQty.Text)
                strSQL = "UPDATE TBL_STOCK SET QTY_IN = " & dblNewIn & ",QTY_AKHIR = " & dblNewAkhir & " WHERE KD_BRG ='" & cboBarang.Text & "'"
                objCommand = New OleDbCommand(strSQL, myCon)
                objCommand.ExecuteNonQuery()
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

    Private Sub ButtonExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonExit.Click
        Me.Close()
    End Sub

    Private Sub grdBarang_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdBarang.Click
        Try
            txtNomor.Text = grdBarang.SelectedCells(0).Value
            txtTgl.Text = grdBarang.SelectedCells(1).Value
            cboBarang.Text = grdBarang.SelectedCells(2).Value
            txtQty.Text = Format(grdBarang.SelectedCells(3).Value, "#,##0.00")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub txtQty_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtQty.KeyPress
        If (e.KeyChar < Chr(48) Or e.KeyChar > Chr(57)) And e.KeyChar <> Chr(8) And e.KeyChar <> Chr(Asc(".")) Then
            e.Handled = True
        End If
    End Sub ' KEYPRES HANYA ANGKA

    Private Sub cboBarang_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboBarang.TextChanged
        Try
            myCon = New OleDbConnection(strCon)
            objDataTable = New DataTable
            myCon.Open()
            strSQL = "SELECT * FROM TBL_BARANG WHERE KD_BRG = '" & cboBarang.Text & "'"
            objCommand = New OleDbCommand(strSQL, myCon)
            objReader = objCommand.ExecuteReader(CommandBehavior.Default)
            If objReader.HasRows Then
                objReader.Read()
                txtNamaBarang.Text = objReader("NM_BRG")
                txtHargaSatuan.Text = Format(objReader("HRG_SAT"), "#,##0.00")
                txtSatuanBarang.Text = objReader("SAT_BRG")
                txtSpecBarang.Text = objReader("SPEC_BRG")
            Else
                objReader.Read()
                txtNamaBarang.Text = ""
                txtHargaSatuan.Text = ""
                txtSatuanBarang.Text = ""
                txtSpecBarang.Text = ""
            End If
            objCommand.Dispose()
            objReader.Close()
            myCon.Close()
            objCommand = Nothing
            objReader = Nothing
            myCon = Nothing
            Call GetStock(cboBarang.Text)
        Catch ex As Exception
            grdBarang = Nothing
        End Try
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
                blnStock = True
            Else
                txtIn.Text = 0
                txtOut.Text = 0
                txtAkhir.Text = 0
                blnStock = False
            End If
            objReader.Close()
            myCon.Close()
        Catch ex As Exception

        End Try        
    End Sub

    Private Sub grdBarang_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles grdBarang.CellContentClick

    End Sub
End Class