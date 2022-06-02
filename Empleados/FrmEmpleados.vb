Public Class FrmEmpleados

    Dim empleados As New LMBADataSetTableAdapters.EmpleadosDgvTableAdapter
    Dim emp As New LMBADataSetTableAdapters.EmpleadoTableAdapter
    Dim idEmp As New Int32

    Private Sub FrmEmpleados_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenarGrid()
        BtnMod.Enabled = False
        BtnEliminar.Enabled = False
    End Sub

    Sub llenarGrid()
        DgvEmpleados.DataSource = empleados.GetData()
        DgvEmpleados.Columns("idHorario").Visible = False
        DgvEmpleados.Columns(0).Visible = False
        DgvEmpleados.Refresh()
    End Sub

    Private Sub DgvEmpleados_DoubleClick(sender As Object, e As EventArgs) Handles DgvEmpleados.DoubleClick
        Dim row = DgvEmpleados.CurrentRow.Index
        Dim emp = DgvEmpleados.Rows(row)
        Dim id = emp.Cells(0).Value
        idEmp = id
        BtnMod.Enabled = True
        BtnEliminar.Enabled = True
    End Sub

    Private Sub BtnMod_Click(sender As Object, e As EventArgs) Handles BtnMod.Click
        FrmModEmp.idEmp = idEmp
        FrmModEmp.Show()
    End Sub

    Private Sub BtnAgregar_Click(sender As Object, e As EventArgs) Handles BtnAgregar.Click
        FrmAddEmp.Show()
    End Sub

    Private Sub BtnBuscar_Click(sender As Object, e As EventArgs) Handles BtnBuscar.Click
        Try
            Dim texto As String = TxtBuscar.Text & "%"
            DgvEmpleados.DataSource = empleados.Search(texto)
            DgvEmpleados.Refresh()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub BtnEliminar_Click(sender As Object, e As EventArgs) Handles BtnEliminar.Click

        Dim row = DgvEmpleados.CurrentRow.Index
        Dim empleado = DgvEmpleados.Rows(row)
        Dim id = empleado.Cells(0).Value
        Dim empStr = id & " - " & empleado.Cells(2).Value & " " & empleado.Cells(3).Value

        Dim delete = MsgBox("¿Seguro que desea eliminar a " & empStr & "?", MsgBoxStyle.YesNo, "Eliminar")

        If (delete = delete.Yes) Then
            emp.DeleteEmpleado(idEmp)
            MsgBox(empStr & " eliminado con éxito", MsgBoxStyle.Information, "Eliminar")
            Me.llenarGrid()
            BtnEliminar.Enabled = False
            BtnMod.Enabled = False
        End If
    End Sub
End Class