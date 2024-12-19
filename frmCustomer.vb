Imports System.Data.OleDb
Public Class frmCustomer
    Dim adCus As New OleDbDataAdapter
    Dim ds As New DataSet
    Dim n As Integer
    Dim chrDBCommand As Char
    Private Sub frmCustomer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        con.Open()
        Dim cmCus As New OleDbCommand
        cmCus.Connection = con
        cmCus.CommandText = "SELECT * FROM пользователь"
        adCus.SelectCommand = cmCus
        adCus.Fill(ds, "пользователь")
        n = ds.Tables("пользователь").Rows.Count - 1
        con.Close()
        showRecords()

    End Sub
    Sub showRecords()
        Dim drCus As DataRow
        If n >= 0 Then
            drCus = ds.Tables("пользователь").Rows(n)
            With drCus
                txtCNum.Text = .Item("Код")
                c_Name.Text = .Item("имя")
                Surname.Text = .Item("фамилия")
                If Not IsDBNull(.Item("отчество")) Then
                    Patronymic.Text = .Item("отчество")
                Else
                    Patronymic.Text = ""
                End If
                If Not IsDBNull(.Item("телефон")) Then
                    Phone.Text = .Item("телефон")
                Else
                    Phone.Text = ""
                End If
                If Not IsDBNull(.Item("почта")) Then
                    Email.Text = .Item("почта")
                Else
                    Email.Text = ""
                End If
                If Not IsDBNull(.Item("счет")) Then
                    Bank.Text = .Item("счет")
                Else
                    Bank.Text = ""
                End If

            End With
        End If
    End Sub

    Private Sub btnFirst_Click(sender As Object, e As EventArgs) Handles btnFirst.Click
        n = 0
        showRecords()
    End Sub

    Private Sub btnPrevious_Click(sender As Object, e As EventArgs) Handles btnPrevious.Click
        If n > 0 Then
            n = n - 1
            showRecords()
        End If
        txtMsg.Text = ""
    End Sub

    Private Sub btnNext_Click(sender As Object, e As EventArgs) Handles btnNext.Click
        If n < ds.Tables("пользователь").Rows.Count - 1 Then
            n = n + 1
            showRecords()
        End If
        txtMsg.Text = ""
    End Sub

    Private Sub btnLast_Click(sender As Object, e As EventArgs) Handles btnLast.Click
        n = ds.Tables("пользователь").Rows.Count - 1
        showRecords()
        txtMsg.Text = ""
    End Sub

    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        chrDBCommand = "A"
        clearControls()
        txtMsg.Text = ""
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click
        If n >= 0 Then
            Dim cmBuilder As New OleDbCommandBuilder
            cmBuilder.DataAdapter = adCus
            Dim tbCus As DataTable
            Dim dcPrimaryKey(0) As DataColumn
            tbCus = ds.Tables("пользователь")
            dcPrimaryKey(0) = tbCus.Columns("Код")
            tbCus.PrimaryKey = dcPrimaryKey
            Dim drCus As DataRow = tbCus.Rows.Find(txtCNum.Text)
            With drCus
                .Item("Код") = txtCNum.Text
                .Item("имя") = c_Name.Text
                .Item("фамилия") = Surname.Text
                .Item("отчество") = Patronymic.Text
                .Item("телефон") = Phone.Text
                .Item("почта") = Email.Text
                .Item("счет") = Bank.Text
            End With
            adCus.UpdateCommand = cmBuilder.GetUpdateCommand
            txtMsg.Text = "Данные обновлены"
            txtMsg.ForeColor = Color.Orange
            Try
                adCus.Update(ds, "пользователь")
                txtMsg.Text = "Данные обновлены"
                txtMsg.ForeColor = Color.Orange
            Catch ex As Exception
                txtMsg.Text = "Ошибка: " & ex.Message
                txtMsg.ForeColor = Color.Red
            End Try
        End If
    End Sub

    Private Sub btn_search_Click(sender As Object, e As EventArgs) Handles btn_search.Click
        txtMsg.Text = c_Search.Text
        con.Open()
        Dim cmCus As New OleDbCommand
        cmCus.Connection = con
        cmCus.CommandText = "SELECT * FROM пользователь WHERE пользователь.фамилия LIKE " & Chr(34) & c_Search.Text & Chr(34)
        adCus.SelectCommand = cmCus
        adCus.Fill(ds, "пользователь")
        n = ds.Tables("пользователь").Rows.Count - 1
        con.Close()
        showRecords()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        If n >= 0 Then
            Dim cmBuilder As New OleDbCommandBuilder
            cmBuilder.DataAdapter = adCus
            ds.Tables("пользователь").Rows(n).Delete()
            adCus.DeleteCommand = cmBuilder.GetDeleteCommand
            n = n - 1
            txtMsg.Text = "Запись удалена"
            txtMsg.ForeColor = Color.Red

            con.Open()
            Try
                adCus.Update(ds, "пользователь")
                clearControls()
                con.Close()
                showRecords()
            Catch ex As Exception
                con.Close()
                MessageBox.Show("Некорректные данные")
            End Try
        End If
        con.Open()
        Dim cmCus As New OleDbCommand
        cmCus.Connection = con
        cmCus.CommandText = "SELECT * FROM пользователь"
        adCus.SelectCommand = cmCus
        adCus.Fill(ds, "пользователь")
        n = ds.Tables("пользователь").Rows.Count - 1
        con.Close()
        showRecords()
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim cmBuilder As New OleDbCommandBuilder
        cmBuilder.DataAdapter = adCus
        If chrDBCommand = "A" Then
            txtMsg.Clear()
            If c_Name.Text = "" Or Surname.Text = "" Then
                MessageBox.Show("Введите все данные")
            Else
                Dim drCus As DataRow
                drCus = ds.Tables("пользователь").NewRow
                With drCus
                    .Item("имя") = c_Name.Text
                    .Item("фамилия") = Surname.Text
                    .Item("отчество") = Patronymic.Text
                    .Item("телефон") = Phone.Text
                    .Item("почта") = Email.Text
                    .Item("счет") = Bank.Text
                End With
                ds.Tables("пользователь").Rows.Add(drCus)
                adCus.InsertCommand = cmBuilder.GetInsertCommand
                n = n + 1
                txtMsg.Text = "Сохранено"
                txtMsg.ForeColor = Color.Green
            End If
        End If
        con.Open()
        Try
            adCus.Update(ds, "пользователь")
            clearControls()
            'showRecords()
        Catch ex As Exception
            MessageBox.Show("Некорректные данные")
        End Try
        con.Close()
        con.Open()
        Dim cmCus As New OleDbCommand
        cmCus.Connection = con
        cmCus.CommandText = "SELECT * FROM пользователь"
        adCus.SelectCommand = cmCus
        adCus.Fill(ds, "пользователь")
        n = ds.Tables("пользователь").Rows.Count - 1
        con.Close()
        showRecords()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        showRecords()
    End Sub

    Sub clearControls()
        txtCNum.Clear()
        c_Name.Clear()
        Surname.Clear()
        Patronymic.Clear()
        Phone.Clear()
        Bank.Clear()
        Email.Clear()

    End Sub

    Private Sub btn_reset_search_Click(sender As Object, e As EventArgs) Handles btn_reset_search.Click
        txtMsg.Text = ""
        con.Open()
        Dim cmCus As New OleDbCommand
        cmCus.Connection = con
        cmCus.CommandText = "SELECT * FROM пользователь"
        adCus.SelectCommand = cmCus
        adCus.Fill(ds, "пользователь")
        n = ds.Tables("пользователь").Rows.Count - 1
        con.Close()
        showRecords()
    End Sub

End Class
