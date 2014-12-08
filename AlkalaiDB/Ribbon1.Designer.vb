﻿Partial Class Ribbon1
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btn_createTable = Me.Factory.CreateRibbonButton
        Me.btn_deleteTable = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.btn_insertRow = Me.Factory.CreateRibbonButton
        Me.btn_deleteRow = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Label = "TabAddIns"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btn_createTable)
        Me.Group1.Items.Add(Me.btn_deleteTable)
        Me.Group1.Name = "Group1"
        '
        'btn_createTable
        '
        Me.btn_createTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_createTable.Image = Global.AlkalaiDB.My.Resources.Resources.table_icon
        Me.btn_createTable.Label = "Create Table"
        Me.btn_createTable.Name = "btn_createTable"
        Me.btn_createTable.ShowImage = True
        '
        'btn_deleteTable
        '
        Me.btn_deleteTable.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.btn_deleteTable.Image = Global.AlkalaiDB.My.Resources.Resources.DeleteRed
        Me.btn_deleteTable.Label = "Delete Table"
        Me.btn_deleteTable.Name = "btn_deleteTable"
        Me.btn_deleteTable.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.btn_insertRow)
        Me.Group2.Items.Add(Me.btn_deleteRow)
        Me.Group2.Label = "Edit Table"
        Me.Group2.Name = "Group2"
        '
        'btn_insertRow
        '
        Me.btn_insertRow.Label = "Insert Row"
        Me.btn_insertRow.Name = "btn_insertRow"
        '
        'btn_deleteRow
        '
        Me.btn_deleteRow.Label = "Delete Row"
        Me.btn_deleteRow.Name = "btn_deleteRow"
        '
        'Ribbon1
        '
        Me.Name = "Ribbon1"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_createTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Protected WithEvents btn_deleteTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btn_insertRow As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btn_deleteRow As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As Ribbon1
        Get
            Return Me.GetRibbon(Of Ribbon1)()
        End Get
    End Property
End Class
