Imports System.Text.RegularExpressions

Public Class FrmAddEmp

    Dim cargos As New LMBADataSetTableAdapters.CargoTableAdapter
    Dim dptos As New LMBADataSetTableAdapters.DepartamentoTableAdapter
    Dim horarios As New LMBADataSetTableAdapters.HorarioTableAdapter
    Dim empleados As New LMBADataSetTableAdapters.EmpleadoTableAdapter

    Public caller

    Private Sub FrmAddEmp_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        llenarCampos()
    End Sub

    Private Sub llenarCampos()
        CbxCargo.DataSource = cargos.GetData()
        CbxCargo.DisplayMember = "nombreCargo"
        CbxCargo.ValueMember = "idCargo"

        CbxDpto.DataSource = dptos.GetData()
        CbxDpto.DisplayMember = "nombreDepartamento"
        CbxDpto.ValueMember = "idDepartamento"

        CbxHorario.DataSource = horarios.GetData()
        CbxHorario.DisplayMember = "nombre"
        CbxHorario.ValueMember = "idHorario"

        CbxSexo.Items.Add("Hombre")
        CbxSexo.Items.Add("Mujer")
        CbxSexo.SelectedIndex = 0

        ChkEstado.Checked = True
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
        If (empleados.GetData.Select($"numCedula='{TxtCedula.Text}'").Count > 0) Then
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

        empleados.InsertQuery(
            TxtCedula.Text,
            If(ChkEstado.Checked, 1, 3),
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
            CbxHorario.SelectedValue
            )

        MsgBox("Empleado añadido", MsgBoxStyle.Information, "Éxito")
        Me.Close()
        FrmEmpleados.llenarGrid()
    End Sub

    Private Sub CbxCargo_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles CbxCargo.SelectionChangeCommitted
        Dim idDpto = cargos.GetData.Select("idCargo='" & CbxCargo.SelectedValue & "'")(0)("idDepartamento")
        CbxDpto.SelectedValue = idDpto
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Me.Close()
    End Sub
End Class