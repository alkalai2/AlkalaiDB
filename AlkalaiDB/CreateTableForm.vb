Public Class CreateTableForm
    Public entry As Entry
    Public range As Excel.Range

    Private Sub CreateTableForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        range = xlApp.Range(xlApp.ActiveCell, xlApp.ActiveCell.Cells(xlApp.Selection.rows.count, xlApp.Selection.columns.count))
        entry = New Entry(range)

        txt_createFormTableName.Text = entry.tname
        With combo_createFormDataTypes.Items
            .Add("real")
            .Add("integer")
            .Add("character(30)")
            .Add("date")
            .Add("serial")
        End With
        For Each a In entry.attr
            list_createFormAttributes.Items.Add(a)
        Next

        'populate localhost config
        txt_localServer.Text = "localhost"
        txt_localPort.Text = "5432"
        txt_localDB.Text = "MyDB"
        txt_localUser.Text = "postgres"
        txt_localPassword.Text = "*********"

    End Sub

    Private Sub btn_createFormCreate_Click(sender As Object, e As EventArgs) Handles btn_createFormCreate.Click

        ' check pk
        Dim pks As Integer = 0
        For i = 0 To entry.constr.Count - 1
            If (InStr(entry.constr(i), "PRIMARY KEY")) Then
                pks = pks + 1
            End If
        Next
        If (pks > 1) Then
            MsgBox("Please selct one Primary Key")
            Return
        End If


        Dim result As Integer = 0
        result = entry.createTable()
        If (result = 1) Then
            'populate excel with DB data
            entry.allowEventChanges = False
            entry.populateTableValues()
            ' change event for tables

            entry.onChangeEvent()
        End If

        

        Me.Close()
    End Sub
    Private Sub btn_createFormCancel_Click(sender As Object, e As EventArgs) Handles btn_createFormCancel.Click
        Me.Close()
    End Sub

    Private Sub list_createFormAttributes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles list_createFormAttributes.SelectedIndexChanged
        group_createFormEdit.Enabled = True
        Dim index As Integer = list_createFormAttributes.SelectedIndex
        For i As Integer = 0 To combo_createFormDataTypes.Items.Count - 1
            If (InStr(entry.types(index), combo_createFormDataTypes.Items(i))) Then
                combo_createFormDataTypes.SelectedIndex = i
            End If
        Next
        check_createFormPK.Checked = InStr(entry.constr(index), "PRIMARY KEY") > 0
        check_createFormNotNull.Checked = InStr(entry.constr(index), "NOT NULL") > 0

    End Sub

    Private Sub combo_createFormDataTypes_SelectedIndexChanged(sender As Object, e As EventArgs) Handles combo_createFormDataTypes.SelectedIndexChanged
        Dim index As Integer = list_createFormAttributes.SelectedIndex
        entry.types(index) = combo_createFormDataTypes.SelectedItem

    End Sub

    Private Sub check_createFormPK_CheckedChanged(sender As Object, e As EventArgs) Handles check_createFormPK.CheckedChanged

        Dim index As Integer = list_createFormAttributes.SelectedIndex
        If (check_createFormPK.Checked = True) Then
            entry.constr(index) = entry.constr(index) + "  PRIMARY KEY "

        Else
            entry.constr(index) = Replace(entry.constr(index), "PRIMARY KEY", "")

        End If

    End Sub

    Private Sub check_createFormNotNull_CheckedChanged(sender As Object, e As EventArgs) Handles check_createFormNotNull.CheckedChanged
        Dim index As Integer = list_createFormAttributes.SelectedIndex
        If (check_createFormNotNull.Checked = True) Then
            entry.constr(index) = entry.constr(index) + " NOT NULL "
        Else
            entry.constr(index) = Replace(entry.constr(index), "NOT NULL", "")
        End If
    End Sub

    Private Sub txt_createFormTableName_TextChanged(sender As Object, e As EventArgs) Handles txt_createFormTableName.TextChanged
        entry.tname = txt_createFormTableName.Text
    End Sub

    Private Sub check_createFormRemote_CheckedChanged(sender As Object, e As EventArgs) Handles check_createFormRemote.CheckedChanged
        If (check_createFormRemote.Checked = True) Then
            group_createFormLocal.Enabled = False
        Else
            group_createFormLocal.Enabled = True
        End If
    End Sub
End Class