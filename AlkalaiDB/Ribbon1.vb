﻿Imports Microsoft.Office.Tools.Ribbon

Public Class Ribbon1

    Private Sub Ribbon1_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub
   
    

    Private Sub btn_createTable_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_createTable.Click
        Dim createTableForm As CreateTableForm
        createTableForm = New CreateTableForm
        createTableForm.Show()

        
    End Sub

    Private Sub btn_insertRow_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_insertRow.Click
        Dim found As Entry = find_entry(xlApp.Selection, True)
        If (IsNothing(found)) Then
            Return
        End If
        found.insertRow(xlApp.ActiveCell, xlApp.Selection.columns.count)
        'update range/size of entry
        'found.range = xlApp.Range(found.loc, found.loc.Cells(found.rows + 1, found.cols))
        found.populateTableValues()
        'MsgBox(found.range.Address)
        Dim valueRange As Excel.Range = xlApp.Range(found.loc.Cells(3, 1), found.loc.Cells(found.rows, found.cols))
        'MsgBox(valueRange.Address)

    End Sub

    Private Sub btn_deleteRow_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_deleteRow.Click
        Dim found As Entry = find_entry(xlApp.Selection, False)
        If (IsNothing(found)) Then
            Return
        End If
        found.deleteRow(xlApp.ActiveCell, xlApp.Selection.columns.count)
        'update range/size of entry
        'found.range = xlApp.Range(found.loc, found.loc.Cells(found.rows - 1, found.cols))
        found.populateTableValues()
        'MsgBox(found.range.Address)
        Dim valueRange As Excel.Range = xlApp.Range(found.loc.Cells(3, 1), found.loc.Cells(found.rows, found.cols))
        'MsgBox(valueRange.Address)
    End Sub


    Private Sub btn_deleteTable_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_deleteTable.Click

        Dim found As Entry = remove_entry(xlApp.ActiveCell)
        If IsNothing(found) = False Then
            found.deleteTable()
            found = Nothing
        End If
    End Sub

    
    Private Sub btn_customQuery_Click(sender As Object, e As RibbonControlEventArgs) Handles btn_customQuery.Click
        Dim customQueryForm As CustomQueryForm
        customQueryForm = New CustomQueryForm
        customQueryForm.Show()
    End Sub
End Class
