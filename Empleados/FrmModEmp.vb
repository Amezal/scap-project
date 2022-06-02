Imports System.Text.RegularExpressions

Public Class FrmModEmp

    Public idEmp As New Int32
    Dim empleado As New LMBADataSetTableAdapters.EmpleadoTableAdapter
    Dim cargos As New LMBADataSetTableAdapters.CargoTableAdapter
    Dim dptos As New LMBADataSetTableAdapters.DepartamentoTableAdapter
    Dim horarios As New LMBADataSetTableAdapters.HorarioTableAdapter
    Dim tblEmp As New LMBADataSet.EmpleadoDataTable


    Private Sub FrmModEmp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TxtCedula.Select()
        tblEmp = empleado.GetEmpleado(idEmp)
        fillCampos()
    End Sub

    Private Sub fillCampos()
        Dim emp = tblEmp.Item(0)
        txtID.Text = emp.idEmpleado
        TxtCedula.Text = emp.numCedula
        TxtPrimerNombre.Text = emp.primerNombre
        TxtSegundoNombre.Text = If(emp.IsNull("segundoNombre"), "", emp.segundoNombre)
        TxtPrimerApellido.Text = emp.primerApellido
        TxtSegundoApellido.Text = If(emp.IsNull("segundoApellido"), "", emp.segundoApellido)

        'cargo y dpto
        CbxCargo.DataSource = cargos.GetData()
        CbxCargo.DisplayMember = "nombreCargo"
        CbxCargo.ValueMember = "idCargo"
        CbxCargo.SelectedValue = emp.idCargo

        CbxDpto.DataSource = dptos.GetData()
        CbxDpto.DisplayMember = "nombreDepartamento"
        CbxDpto.ValueMember = "idDepartamento"
        CbxDpto.SelectedValue = cargos.GetData().Select("idCargo='" & emp.idCargo & "'")(0)("idDepartamento")
        '

        TxtEmailCorp.Text = emp.emailCorporativo
        TxtEmail.Text = If(emp.IsNull("emailPersonal"), "", emp.emailPersonal)
        TxtTelefono.Text = emp.telefono

        'sexo y horario
        CbxSexo.Items.Add("Hombre")
        CbxSexo.Items.Add("Mujer")
        CbxSexo.SelectedText = If(emp.sexo, "Hombre", "Mujer")

        CbxHorario.DataSource = horarios.GetData()
        CbxHorario.DisplayMember = "nombre"
        CbxHorario.ValueMember = "idHorario"
        CbxHorario.SelectedValue = emp.idHorario
        '

        ChkEstado.Checked = emp.estado
        TxtPIN.Text = emp.PIN
        TxtDireccion.Text = emp.direccion
        TxtObservacion.Text = If(emp.IsNull("observacion"), "", emp.observacion)
        'fechas
        DtpIngreso.Value = If(emp.IsNull("fechaIngreso"), Date.Today, emp.fechaIngreso)
        DtpNac.Value = emp.fechaNacimiento
    End Sub

    Private Function Validar() As Boolean
        If (String.IsNullOrEmpty(TxtCedula.Text)) Then
            MsgBox("No puede quedar vacío la cédula", MsgBoxStyle.Critical, "ERROR")
            TxtCedula.Focus()
            Return False
        End If
        Dim cedula As Regex = New Regex("\d{3}-\d{6}-\d{4}[A-Z]")
        If (Not cedula.IsMatch(TxtCedula.Text)) Then
            MsgBox("La cédula debe tener el formato 000-000000-0000A", MsgBoxStyle.Critical, "ERROR")
            TxtCedula.Focus()
            Return False
        End If
        If (empleado.GetData.Select($"numCedula='{TxtCedula.Text}' AND idEmpleado <> '{txtID.Text}'").Count > 0) Then
            MsgBox("La cédula debe ser única", MsgBoxStyle.Critical, "ERROR")
            TxtCedula.Focus()
            Return False
        End If
        If (String.IsNullOrEmpty(TxtPrimerNombre.Text)) Then
            MsgBox("No puede quedar vacío el nombre", MsgBoxStyle.Critical, "ERROR")
            TxtPrimerNombre.Focus()
            Return False
        End If
        If (String.IsNullOrEmpty(TxtPrimerApellido.Text)) Then
            MsgBox("No puede quedar vacío el apellido", MsgBoxStyle.Critical, "ERROR")
            TxtSegundoNombre.Focus()
            Return False
        End If
        If (String.IsNullOrEmpty(TxtPIN.Text)) Then
            MsgBox("Debe añadir un PIN", MsgBoxStyle.Critical, "ERROR")
            TxtPIN.Focus()
            Return False
        End If
        If (TxtPIN2.TextLength = 0) Then
            MsgBox("Debe confirmar el PIN", MsgBoxStyle.Critical, "ERROR")
            TxtPIN2.Focus()
            Return False
        End If
        If (Not TxtPIN.Text.Equals(TxtPIN2.Text)) Then
            MsgBox("Los PINs deben coincidir", MsgBoxStyle.Critical, "ERROR")
            TxtPIN2.Focus()
            Return False
        End If
        If (String.IsNullOrEmpty(TxtDireccion.Text)) Then
            MsgBox("La dirección no puede quedar vacía", MsgBoxStyle.Critical, "ERROR")
            TxtPIN2.Focus()
            Return False
        End If
        Return True
    End Function

    Private Sub BtnPIN_MouseDown(sender As Object, e As MouseEventArgs) Handles BtnPIN.MouseDown
        TxtPIN.PasswordChar = ""
    End Sub

    Private Sub BtnPIN_MouseUp(sender As Object, e As MouseEventArgs) Handles BtnPIN.MouseUp
        TxtPIN.PasswordChar = "*"
    End Sub

    Private Sub BtnPIN2_MouseDown(sender As Object, e As MouseEventArgs) Handles BtnPIN2.MouseDown
        TxtPIN2.PasswordChar = ""
    End Sub

    Private Sub BtnPIN2_MouseUp(sender As Object, e As MouseEventArgs) Handles BtnPIN2.MouseUp
        TxtPIN2.PasswordChar = "*"
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles BtnGuardar.Click
        If (Not Validar()) Then
            Exit Sub
        End If

        empleado.UpdateQuery(
            TxtCedula.Text,
            If(ChkEstado.Checked, 2, 3),
            TxtPIN.Text,
            TxtPrimerNombre.Text.Trim,
            TxtSegundoNombre.Text.Trim,
            TxtPrimerApellido.Text.Trim,
            TxtSegundoApellido.Text.Trim,
            DtpNac.Value.Date.ToString(),
            If(CbxSexo.SelectedIndex.Equals(0), True, False),
            DtpIngreso.Value.Date.ToString(),
            TxtDireccion.Text.Trim,
            TxtObservacion.Text.Trim,
            TxtTelefono.Text.Trim,
            TxtEmail.Text.Trim,
            TxtEmailCorp.Text.Trim,
            CbxCargo.SelectedValue,
            CbxHorario.SelectedValue,
            idEmp
            )

        MsgBox("Empleado modificado", MsgBoxStyle.Information, "Éxito")
        Me.Close()
        FrmEmpleados.llenarGrid()
    End Sub

    Private Sub BtnCancelar_Click(sender As Object, e As EventArgs) Handles BtnCancelar.Click
        Me.Close()
    End Sub
End Class