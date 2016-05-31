Imports System.Data.SqlClient
Public Class Form1
    Dim dep As String() = {5, 10, 20, 25, 30, 35, 40, 50, 55, 60, 65, 70, 75, 80, 85, 91, 93}
    Dim m5 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60}
    Dim m10 As String() = {5, 10, 12, 15, 20, 25, 30, 35}
    Dim m20 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45}
    Dim m25 As String() = {5, 10, 15, 20, 25, 30}
    Dim m30 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65}
    Dim m35 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45, 50}
    Dim m40 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45, 50, 55, 60, 65}
    Dim m50 As String() = {5, 10, 15, 20, 25, 30}
    Dim m55 As String() = {5, 10, 15, 20, 22, 25, 30, 32, 35}
    Dim m60 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45}
    Dim m65 As String() = {5, 7, 10, 15, 20, 25, 30, 35, 40, 45}
    Dim m70 As String() = {5, 10, 15, 20}
    Dim m75 As String() = {5, 10, 15, 20, 25, 30, 35, 40}
    Dim m80 As String() = {5, 10, 15, 20, 25, 30, 35, 40, 45, 50}
    Dim m85 As String() = {5, 10, 15, 20, 25, 30}
    Dim m91 As String() = {5, 10, 15, 20, 25, 27, 30, 35}
    Dim m93 As String() = {5, 10, 12, 15, 16, 20, 23, 25, 30, 35, 40, 45}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn()
        For Each item As String In dep
            tbDepartamentoProductor.Items.Add(item)
            tbDepartamentoExplotacion.Items.Add(item)
        Next

        For index = 1 To 37
            tb0241.Items.Add(index)
            tb0242.Items.Add(index)
            tb0243.Items.Add(index)
            tb0244.Items.Add(index)
            tb0245.Items.Add(index)
            'Expectativas
            tb0151.Items.Add(index)
            tb0152.Items.Add(index)
            tb0153.Items.Add(index)
            tb0154.Items.Add(index)
            tb0155.Items.Add(index)
            tb0181.Items.Add(index)
            tb0182.Items.Add(index)
            tb0183.Items.Add(index)
            tb0184.Items.Add(index)
            tb0185.Items.Add(index)
            tb0211.Items.Add(index)
            tb0212.Items.Add(index)
            tb0213.Items.Add(index)
            tb0214.Items.Add(index)
            tb0215.Items.Add(index)
            'sembrado
            tb0021.Items.Add(index)
            tb0026.Items.Add(index)
            tb0031.Items.Add(index)
            tb0036.Items.Add(index)
            tb0041.Items.Add(index)
            tb0046.Items.Add(index)
            tb0051.Items.Add(index)
            tb0056.Items.Add(index)
        Next
        For index = 1 To 8
            tb5000.Items.Add(index)
        Next
        For index = 1 To 19
            tb0166.Items.Add(index)
            tb0167.Items.Add(index)
            tb0168.Items.Add(index)
            tb0169.Items.Add(index)
            tb0170.Items.Add(index)
            tb0176.Items.Add(index)
            tb0177.Items.Add(index)
            tb0178.Items.Add(index)
            tb0179.Items.Add(index)
            tb0180.Items.Add(index)
            tb0196.Items.Add(index)
            tb0197.Items.Add(index)
            tb0198.Items.Add(index)
            tb0199.Items.Add(index)
            tb0200.Items.Add(index)
            tb0206.Items.Add(index)
            tb0207.Items.Add(index)
            tb0208.Items.Add(index)
            tb0209.Items.Add(index)
            tb0210.Items.Add(index)
            tb0226.Items.Add(index)
            tb0227.Items.Add(index)
            tb0228.Items.Add(index)
            tb0229.Items.Add(index)
            tb0230.Items.Add(index)
            tb0236.Items.Add(index)
            tb0237.Items.Add(index)
            tb0238.Items.Add(index)
            tb0239.Items.Add(index)
            tb0240.Items.Add(index)
        Next

    End Sub

    Private Sub bGuardar_Click(sender As Object, e As EventArgs) Handles bGuardar.Click
        Dim val = validar()

        If val = "True" Then
            cmd = New SqlCommand
            cmd.Connection = con
            query = createQuery()
            cmd.CommandText = query
            cmd.ExecuteNonQuery()
            MsgBox("Registro Agregado")
            clean()
            tc1.SelectedTab = TabPage1
            tbIdCuestionario.Focus()
        Else
            MsgBox(val)
        End If
    End Sub
    Private Sub bNuevo_Click(sender As Object, e As EventArgs) Handles bNuevo.Click
        'Console.Write(createQuery())
        clean()
        tc1.SelectedTab = TabPage1
        tbIdCuestionario.Focus()
    End Sub

    Private Function createQuery() As String
        Return "INSERT INTO Encuesta VALUES ('" &
            tbIdCuestionario.Text & "', '" &
            tbCopia.Text & "', '" &
            tbSegmento.Text & "', '" &
            tbNombreProductor.Text & "', '" &
            tbDepartamentoProductor.Text & "', '" &
            tbMunicipioProductor.Text & "', '" &
            tbTelefonoProductor.Text & "', '" &
            tbCiudadProductor.Text & "', '" &
            tbReferenciaProductor.Text & "', '" &
            tbDireccionProductor.Text & "', '" &
            tbNombreExplotacion.Text & "', '" &
            tbDepartamentoExplotacion.Text & "', '" &
            tbMunicipioExplotacion.Text & "', '" &
            tbCiudadExplotacion.Text & "', '" &
            tbReferenciaExplotacion.Text & "', '" &
            tbDireccionExplotacion.Text & "', '" &
            tb0020.Text & "', '" &
            tb0407.Text & "', '" &
            tb0021.Text & "', '" &
            tb0022.Text & "', '" &
            tb0023.Text & "', '" &
            tb0024.Text & "', '" &
            tb0025.Text & "', '" &
            tb0026.Text & "', '" &
            tb0027.Text & "', '" &
            tb0028.Text & "', '" &
            tb0029.Text & "', '" &
            tb0030.Text & "', '" &
            tb0031.Text & "', '" &
            tb0032.Text & "', '" &
            tb0033.Text & "', '" &
            tb0034.Text & "', '" &
            tb0035.Text & "', '" &
            tb0036.Text & "', '" &
            tb0037.Text & "', '" &
            tb0038.Text & "', '" &
            tb0039.Text & "', '" &
            tb0040.Text & "', '" &
            tb0041.Text & "', '" &
            tb0042.Text & "', '" &
            tb0043.Text & "', '" &
            tb0044.Text & "', '" &
            tb0045.Text & "', '" &
            tb0046.Text & "', '" &
            tb0047.Text & "', '" &
            tb0048.Text & "', '" &
            tb0049.Text & "', '" &
            tb0050.Text & "', '" &
            tb0051.Text & "', '" &
            tb0052.Text & "', '" &
            tb0053.Text & "', '" &
            tb0054.Text & "', '" &
            tb0055.Text & "', '" &
            tb0056.Text & "', '" &
            tb0057.Text & "', '" &
            tb0058.Text & "', '" &
            tb0059.Text & "', '" &
            tb0060.Text & "', '" &
            tb0151.Text & "', '" &
            tb0156.Text & "', '" &
            tb0161.Text & "', '" &
            tb0166.Text & "', '" &
            tb0171.Text & "', '" &
            tb0176.Text & "', '" &
            tb0152.Text & "', '" &
            tb0157.Text & "', '" &
            tb0162.Text & "', '" &
            tb0167.Text & "', '" &
            tb0172.Text & "', '" &
            tb0177.Text & "', '" &
            tb0153.Text & "', '" &
            tb0158.Text & "', '" &
            tb0163.Text & "', '" &
            tb0168.Text & "', '" &
            tb0173.Text & "', '" &
            tb0178.Text & "', '" &
            tb0154.Text & "', '" &
            tb0159.Text & "', '" &
            tb0164.Text & "', '" &
            tb0169.Text & "', '" &
            tb0174.Text & "', '" &'80
            tb0179.Text & "', '" &
            tb0155.Text & "', '" &
            tb0160.Text & "', '" &
            tb0165.Text & "', '" &
            tb0170.Text & "', '" &
            tb0175.Text & "', '" &
            tb0180.Text & "', '" &
            tb0181.Text & "', '" &
            tb0186.Text & "', '" &
            tb0191.Text & "', '" &
            tb0196.Text & "', '" &
            tb0201.Text & "', '" &
            tb0206.Text & "', '" &
            tb0182.Text & "', '" &
            tb0187.Text & "', '" &
            tb0192.Text & "', '" &
            tb0197.Text & "', '" &
            tb0202.Text & "', '" &
            tb0207.Text & "', '" &
            tb0183.Text & "', '" &
            tb0188.Text & "', '" &
            tb0193.Text & "', '" &
            tb0198.Text & "', '" &
            tb0203.Text & "', '" &
            tb0208.Text & "', '" &
            tb0184.Text & "', '" &
            tb0189.Text & "', '" &
            tb0194.Text & "', '" &
            tb0199.Text & "', '" &
            tb0204.Text & "', '" &
            tb0209.Text & "', '" &
            tb0185.Text & "', '" &
            tb0190.Text & "', '" &
            tb0195.Text & "', '" &
            tb0200.Text & "', '" &
            tb0205.Text & "', '" &
            tb0210.Text & "', '" &
            tb0211.Text & "', '" &
            tb0216.Text & "', '" &
            tb0221.Text & "', '" &
            tb0226.Text & "', '" &
            tb0231.Text & "', '" &
            tb0236.Text & "', '" &
            tb0212.Text & "', '" &
            tb0217.Text & "', '" &
            tb0222.Text & "', '" &
            tb0227.Text & "', '" &
            tb0232.Text & "', '" &
            tb0237.Text & "', '" &
            tb0213.Text & "', '" &'130
            tb0218.Text & "', '" &
            tb0223.Text & "', '" &
            tb0228.Text & "', '" &
            tb0233.Text & "', '" &
            tb0238.Text & "', '" &
            tb0214.Text & "', '" &
            tb0219.Text & "', '" &
            tb0224.Text & "', '" &
            tb0229.Text & "', '" &
            tb0234.Text & "', '" &
            tb0239.Text & "', '" &
            tb0215.Text & "', '" &
            tb0220.Text & "', '" &
            tb0225.Text & "', '" &
            tb0230.Text & "', '" &
            tb0235.Text & "', '" &
            tb0240.Text & "', '" &
            tb0241.Text & "', '" &
            tb0241MM.Text & "', '" &
            tb0241MP.Text & "', '" &
            tb0241JM.Text & "', '" &
            tb0241JP.Text & "', '" &
            tb0241JLM.Text & "', '" &
            tb0241JLP.Text & "', '" &
            tb0242.Text & "', '" &
            tb0242MM.Text & "', '" &
            tb0242MP.Text & "', '" &
            tb0242JM.Text & "', '" &
            tb0242JP.Text & "', '" &
            tb0242JLM.Text & "', '" &
            tb0242JLP.Text & "', '" &
            tb0243.Text & "', '" &
            tb0243MM.Text & "', '" &
            tb0243MP.Text & "', '" &
            tb0243JM.Text & "', '" &
            tb0243JP.Text & "', '" &
            tb0243JLM.Text & "', '" &
            tb0243JLP.Text & "', '" &
            tb0244.Text & "', '" &
            tb0244MM.Text & "', '" &'170
            tb0244MP.Text & "', '" &
            tb0244JM.Text & "', '" &
            tb0244JP.Text & "', '" &
            tb0244JLM.Text & "', '" &
            tb0244JLP.Text & "', '" &
            tb0245.Text & "', '" &
            tb0245MM.Text & "', '" &
            tb0245MP.Text & "', '" &
            tb0245JM.Text & "', '" &
            tb0245JP.Text & "', '" &
            tb0245JLM.Text & "', '" &
            tb0245JLP.Text & "', '" &
            tb5000.Text & "', '" &
            tbEnumerador.Text & "', '" &
            tbSupervisor.Text & "', '" &
            dtpFecha.Value.Year & "-" & dtpFecha.Value.Month & "-" & dtpFecha.Value.Day & "', '" &
            tbObservaciones.Text & "')"

    End Function

    Private Function validar()
        If tbIdCuestionario.Text = "" Then
            Return "Falta ID cuestionario!"
        End If
        If tbNombreProductor.Text = "" Then
            Return "Falta nombre del productor!"
        End If
        If tbTelefonoProductor.Text = "" Then
            Return "Falta telefono del productor!"
        End If
        If tbDepartamentoProductor.Text = "" Then
            Return "Falta departamento del productor!"
        End If
        If tbMunicipioProductor.Text = "" Then
            Return "Falta municipio del productor!"
        End If
        If tbCiudadProductor.Text = "" Then
            Return "Falta ciudad del productor!"
        End If
        If tbReferenciaProductor.Text = "" Then
            Return "Falta referencia del productor!"
        End If
        If tbDireccionProductor.Text = "" Then
            Return "Falta direccion del productor!"
        End If
        If tb0020.Text = "" Then
            Return "Falta codigo 0020"
        End If
        If tb0407.Text = "" Then
            Return "Falta codigo 0407"
        End If
        If tb5000.Text = "" Then
            Return "Falta codigo 5000"
        End If
        If tbEnumerador.Text = "" Then
            Return "Falta nombre del enumerador!"
        End If
        If tbSupervisor.Text = "" Then
            Return "Falta nombre del supervisor!"
        End If
        If tbCopia.Text = "" Then
            If tbNombreExplotacion.Text = "" Then
                Return "Falta nombre de la explotacion!"
            End If
            If tbDepartamentoExplotacion.Text = "" Then
                Return "Falta departamento del productor!"
            End If
            If tbMunicipioExplotacion.Text = "" Then
                Return "Falta municipio del productor!"
            End If
            If tbCiudadExplotacion.Text = "" Then
                Return "Falta ciudad del productor!"
            End If
            If tbReferenciaExplotacion.Text = "" Then
                Return "Falta referencia del productor!"
            End If
            If tbDireccionExplotacion.Text = "" Then
                Return "Falta direccion del productor!"
            End If
        End If
        Return "True"
    End Function

    Private Sub clean()
        tbIdCuestionario.Text = ""
        tbCopia.Text = ""
        tbSegmento.Text = ""
        tbNombreProductor.Text = ""
        tbDepartamentoProductor.Text = ""
        tbMunicipioProductor.Text = ""
        tbTelefonoProductor.Text = ""
        tbCiudadProductor.Text = ""
        tbReferenciaProductor.Text = ""
        tbDireccionProductor.Text = ""
        tbNombreExplotacion.Text = ""
        tbDepartamentoExplotacion.Text = ""
        tbMunicipioExplotacion.Text = ""
        tbCiudadExplotacion.Text = ""
        tbReferenciaExplotacion.Text = ""
        tbDireccionExplotacion.Text = ""
        tb0020.Text = ""
        tb0407.Text = ""
        tb0021.Text = ""
        tb0022.Text = ""
        tb0023.Text = ""
        tb0024.Text = ""
        tb0025.Text = ""
        tb0026.Text = ""
        tb0027.Text = ""
        tb0028.Text = ""
        tb0029.Text = ""
        tb0030.Text = ""
        tb0031.Text = ""
        tb0032.Text = ""
        tb0033.Text = ""
        tb0034.Text = ""
        tb0035.Text = ""
        tb0036.Text = ""
        tb0037.Text = ""
        tb0038.Text = ""
        tb0039.Text = ""
        tb0040.Text = ""
        tb0041.Text = ""
        tb0042.Text = ""
        tb0043.Text = ""
        tb0044.Text = ""
        tb0045.Text = ""
        tb0046.Text = ""
        tb0047.Text = ""
        tb0048.Text = ""
        tb0049.Text = ""
        tb0050.Text = ""
        tb0051.Text = ""
        tb0052.Text = ""
        tb0053.Text = ""
        tb0054.Text = ""
        tb0055.Text = ""
        tb0056.Text = ""
        tb0057.Text = ""
        tb0058.Text = ""
        tb0059.Text = ""
        tb0060.Text = ""
        tb0151.Text = ""
        tb0156.Text = ""
        tb0161.Text = ""
        tb0166.Text = ""
        tb0171.Text = ""
        tb0176.Text = ""
        tb0152.Text = ""
        tb0157.Text = ""
        tb0162.Text = ""
        tb0167.Text = ""
        tb0172.Text = ""
        tb0177.Text = ""
        tb0153.Text = ""
        tb0158.Text = ""
        tb0163.Text = ""
        tb0168.Text = ""
        tb0173.Text = ""
        tb0178.Text = ""
        tb0154.Text = ""
        tb0159.Text = ""
        tb0164.Text = ""
        tb0169.Text = ""
        tb0174.Text = ""
        tb0179.Text = ""
        tb0155.Text = ""
        tb0160.Text = ""
        tb0165.Text = ""
        tb0170.Text = ""
        tb0175.Text = ""
        tb0180.Text = ""
        tb0181.Text = ""
        tb0186.Text = ""
        tb0191.Text = ""
        tb0196.Text = ""
        tb0201.Text = ""
        tb0206.Text = ""
        tb0182.Text = ""
        tb0187.Text = ""
        tb0192.Text = ""
        tb0197.Text = ""
        tb0202.Text = ""
        tb0207.Text = ""
        tb0183.Text = ""
        tb0188.Text = ""
        tb0193.Text = ""
        tb0198.Text = ""
        tb0203.Text = ""
        tb0208.Text = ""
        tb0184.Text = ""
        tb0189.Text = ""
        tb0194.Text = ""
        tb0199.Text = ""
        tb0204.Text = ""
        tb0209.Text = ""
        tb0185.Text = ""
        tb0190.Text = ""
        tb0195.Text = ""
        tb0200.Text = ""
        tb0205.Text = ""
        tb0210.Text = ""
        tb0211.Text = ""
        tb0216.Text = ""
        tb0221.Text = ""
        tb0226.Text = ""
        tb0231.Text = ""
        tb0236.Text = ""
        tb0212.Text = ""
        tb0217.Text = ""
        tb0222.Text = ""
        tb0227.Text = ""
        tb0232.Text = ""
        tb0237.Text = ""
        tb0213.Text = ""
        tb0218.Text = ""
        tb0223.Text = ""
        tb0228.Text = ""
        tb0233.Text = ""
        tb0238.Text = ""
        tb0214.Text = ""
        tb0219.Text = ""
        tb0224.Text = ""
        tb0229.Text = ""
        tb0234.Text = ""
        tb0239.Text = ""
        tb0215.Text = ""
        tb0220.Text = ""
        tb0225.Text = ""
        tb0230.Text = ""
        tb0235.Text = ""
        tb0240.Text = ""
        tb0241.Text = ""
        tb0241MM.Text = ""
        tb0241MP.Text = ""
        tb0241JM.Text = ""
        tb0241JP.Text = ""
        tb0241JLM.Text = ""
        tb0241JLP.Text = ""
        tb0242.Text = ""
        tb0242MM.Text = ""
        tb0242MP.Text = ""
        tb0242JM.Text = ""
        tb0242JP.Text = ""
        tb0242JLM.Text = ""
        tb0242JLP.Text = ""
        tb0243.Text = ""
        tb0243MM.Text = ""
        tb0243MP.Text = ""
        tb0243JM.Text = ""
        tb0243JP.Text = ""
        tb0243JLM.Text = ""
        tb0243JLP.Text = ""
        tb0244.Text = ""
        tb0244MM.Text = ""
        tb0244MP.Text = ""
        tb0244JM.Text = ""
        tb0244JP.Text = ""
        tb0244JLM.Text = ""
        tb0244JLP.Text = ""
        tb0245.Text = ""
        tb0245MM.Text = ""
        tb0245MP.Text = ""
        tb0245JM.Text = ""
        tb0245JP.Text = ""
        tb0245JLM.Text = ""
        tb0245JLP.Text = ""
        tb5000.Text = ""
        tbEnumerador.Text = ""
        tbSupervisor.Text = ""
        dtpFecha.Value = Today
        tbObservaciones.Text = ""
    End Sub

    Private Sub tbIdCuestionario_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbIdCuestionario.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbCopia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbCopia.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbSegmento_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbSegmento.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbDepartamentoProductor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDepartamentoProductor.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbMunicipioProductor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbMunicipioProductor.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbTelefonoProductor_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbTelefonoProductor.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbDepartamentoExplotacion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbDepartamentoExplotacion.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tbMunicipioExplotacion_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tbMunicipioExplotacion.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0407_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0407.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            If (e.KeyChar = "1" Or e.KeyChar = "2") Then
                e.Handled = False
            Else
                e.Handled = True
            End If

        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0407_TextChanged(sender As Object, e As EventArgs) Handles tb0407.TextChanged
        If (tb0407.Text <> "") Then
            If (tb0407.Text = 1) Then
                tb0021.Enabled = True
                tb0026.Enabled = True
                tb0031.Enabled = True
                tb0036.Enabled = True
                tb0041.Enabled = True
                tb0046.Enabled = True
                tb0051.Enabled = True
                tb0056.Enabled = True
            ElseIf (tb0407.Text = 2 Or tb0407.Text = "") Then
                tb0021.Enabled = False
                tb0026.Enabled = False
                tb0031.Enabled = False
                tb0036.Enabled = False
                tb0041.Enabled = False
                tb0046.Enabled = False
                tb0051.Enabled = False
                tb0056.Enabled = False
            End If
        End If
    End Sub

    Private Sub tb5000_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb5000.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            If (e.KeyChar = "1" Or e.KeyChar = "2" Or e.KeyChar = "3" Or
            e.KeyChar = "4" Or e.KeyChar = "5" Or e.KeyChar = "6" Or
            e.KeyChar = "7" Or e.KeyChar = "8") Then
                e.Handled = False
            Else
                e.Handled = True
            End If

        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0020_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0020.KeyPress
        If (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".") Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0020_TextChanged(sender As Object, e As EventArgs) Handles tb0020.TextChanged
        If (tb0020.Text <> "") Then
            tb0407.Enabled = True
        Else
            tb0407.Enabled = False
        End If
    End Sub

    Private Sub bClose_Click(sender As Object, e As EventArgs) Handles bClose.Click
        Dispose()
    End Sub

    Private Sub tbCopia_Leave(sender As Object, e As EventArgs) Handles tbCopia.Leave
        If (tbCopia.Text <> "") Then
            cmd = New SqlCommand
            cmd.Connection = con
            cmd.CommandText = "SELECT Encuesta.id_cuestionario FROM Encuesta WHERE id_cuestionario = '" & tbCopia.Text & "'"
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            If Not reader.HasRows Then
                MsgBox("El N° ID del cuestionario copia escrito no existe!")
                tbCopia.Text = ""
            End If
            reader.Close()
            'Do While reader.HasRows
            'Do While reader.Read()
            'Console.WriteLine(reader.GetString(0))
            'Loop
            'reader.NextResult()
            'Loop
        End If
    End Sub

    Private Sub numeric(e As KeyPressEventArgs)
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Private Sub numericPoint(e As KeyPressEventArgs)
        If (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".") Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241_KeyPress_1(sender As Object, e As KeyPressEventArgs) Handles tb0241.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0241_TextChanged_1(sender As Object, e As EventArgs) Handles tb0241.TextChanged
        If (tb0241.Text <> "") Then
            tb0241JLM.Enabled = True
            tb0241JLP.Enabled = True
            tb0241JM.Enabled = True
            tb0241JP.Enabled = True
            tb0241MM.Enabled = True
            tb0241MP.Enabled = True
        Else
            tb0241JLM.Text = ""
            tb0241JLP.Text = ""
            tb0241JM.Text = ""
            tb0241JP.Text = ""
            tb0241MM.Text = ""
            tb0241MP.Text = ""
            'desactivar
            tb0241JLM.Enabled = False
            tb0241JLP.Enabled = False
            tb0241JM.Enabled = False
            tb0241JP.Enabled = False
            tb0241MM.Enabled = False
            tb0241MP.Enabled = False
        End If
    End Sub

    Private Sub tb0242_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0242.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0243_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0243.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0244_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0244.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0245_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0245.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0242_TextChanged(sender As Object, e As EventArgs) Handles tb0242.TextChanged
        If (tb0242.Text <> "") Then
            tb0242JLM.Enabled = True
            tb0242JLP.Enabled = True
            tb0242JM.Enabled = True
            tb0242JP.Enabled = True
            tb0242MM.Enabled = True
            tb0242MP.Enabled = True
        Else
            tb0242JLM.Text = ""
            tb0242JLP.Text = ""
            tb0242JM.Text = ""
            tb0242JP.Text = ""
            tb0242MM.Text = ""
            tb0242MP.Text = ""
            'desactivar
            tb0242JLM.Enabled = False
            tb0242JLP.Enabled = False
            tb0242JM.Enabled = False
            tb0242JP.Enabled = False
            tb0242MM.Enabled = False
            tb0242MP.Enabled = False
        End If
    End Sub
    Private Sub tb0243_TextChanged(sender As Object, e As EventArgs) Handles tb0243.TextChanged
        If (tb0243.Text <> "") Then
            tb0243JLM.Enabled = True
            tb0243JLP.Enabled = True
            tb0243JM.Enabled = True
            tb0243JP.Enabled = True
            tb0243MM.Enabled = True
            tb0243MP.Enabled = True
        Else
            tb0243JLM.Text = ""
            tb0243JLP.Text = ""
            tb0243JM.Text = ""
            tb0243JP.Text = ""
            tb0243MM.Text = ""
            tb0243MP.Text = ""
            'desactivar
            tb0243JLM.Enabled = False
            tb0243JLP.Enabled = False
            tb0243JM.Enabled = False
            tb0243JP.Enabled = False
            tb0243MM.Enabled = False
            tb0243MP.Enabled = False
        End If
    End Sub
    Private Sub tb0244_TextChanged(sender As Object, e As EventArgs) Handles tb0244.TextChanged
        If (tb0244.Text <> "") Then
            tb0244JLM.Enabled = True
            tb0244JLP.Enabled = True
            tb0244JM.Enabled = True
            tb0244JP.Enabled = True
            tb0244MM.Enabled = True
            tb0244MP.Enabled = True
        Else
            tb0244JLM.Text = ""
            tb0244JLP.Text = ""
            tb0244JM.Text = ""
            tb0244JP.Text = ""
            tb0244MM.Text = ""
            tb0244MP.Text = ""
            'desactivar
            tb0244JLM.Enabled = False
            tb0244JLP.Enabled = False
            tb0244JM.Enabled = False
            tb0244JP.Enabled = False
            tb0244MM.Enabled = False
            tb0244MP.Enabled = False
        End If
    End Sub
    Private Sub tb0245_TextChanged(sender As Object, e As EventArgs) Handles tb0245.TextChanged
        If (tb0245.Text <> "") Then
            tb0245JLM.Enabled = True
            tb0245JLP.Enabled = True
            tb0245JM.Enabled = True
            tb0245JP.Enabled = True
            tb0245MM.Enabled = True
            tb0245MP.Enabled = True
        Else
            tb0245JLM.Text = ""
            tb0245JLP.Text = ""
            tb0245JM.Text = ""
            tb0245JP.Text = ""
            tb0245MM.Text = ""
            tb0245MP.Text = ""
            'desactivar
            tb0245JLM.Enabled = False
            tb0245JLP.Enabled = False
            tb0245JM.Enabled = False
            tb0245JP.Enabled = False
            tb0245MM.Enabled = False
            tb0245MP.Enabled = False
        End If
    End Sub

    Private Sub tb0241MM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241MM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0242MM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0242MM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0243MM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0243MM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0244MM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0244MM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0245MM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0245MM.KeyPress
        numericPoint(e)
    End Sub

    Private Sub tb0241MP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241MP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0242MP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0242MP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0243MP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0243MP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0244MP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0244MP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0245MP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0245MP.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0241JM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0242JM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0242JM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0243JM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0243JM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0244JM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0244JM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0245JM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0245JM.KeyPress
        numericPoint(e)
    End Sub

    Private Sub tb0241JP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0242JP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0242JP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0243JP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0243JP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0244JP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0244JP.KeyPress
        numeric(e)
    End Sub
    Private Sub tb0245JP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0245JP.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0241JLM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JLM.KeyPress, tb0245JLM.KeyPress, tb0244JLM.KeyPress, tb0243JLM.KeyPress, tb0242JLM.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0241JLP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JLP.KeyPress, tb0245JLP.KeyPress, tb0244JLP.KeyPress, tb0243JLP.KeyPress, tb0242JLP.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0156_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0156.KeyPress, tb0235.KeyPress, tb0234.KeyPress, tb0233.KeyPress, tb0232.KeyPress, tb0231.KeyPress, tb0225.KeyPress, tb0224.KeyPress, tb0223.KeyPress, tb0222.KeyPress, tb0221.KeyPress, tb0220.KeyPress, tb0219.KeyPress, tb0218.KeyPress, tb0217.KeyPress, tb0216.KeyPress, tb0205.KeyPress, tb0204.KeyPress, tb0203.KeyPress, tb0202.KeyPress, tb0201.KeyPress, tb0195.KeyPress, tb0194.KeyPress, tb0193.KeyPress, tb0192.KeyPress, tb0191.KeyPress, tb0190.KeyPress, tb0189.KeyPress, tb0188.KeyPress, tb0187.KeyPress, tb0186.KeyPress, tb0175.KeyPress, tb0174.KeyPress, tb0173.KeyPress, tb0172.KeyPress, tb0171.KeyPress, tb0165.KeyPress, tb0164.KeyPress, tb0163.KeyPress, tb0162.KeyPress, tb0161.KeyPress, tb0160.KeyPress, tb0159.KeyPress, tb0158.KeyPress, tb0157.KeyPress
        numericPoint(e)
    End Sub
    Private Sub tb0151_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0151.KeyPress, tb0240.KeyPress, tb0239.KeyPress, tb0238.KeyPress, tb0237.KeyPress, tb0236.KeyPress, tb0230.KeyPress, tb0229.KeyPress, tb0228.KeyPress, tb0227.KeyPress, tb0226.KeyPress, tb0215.KeyPress, tb0214.KeyPress, tb0213.KeyPress, tb0212.KeyPress, tb0211.KeyPress, tb0210.KeyPress, tb0209.KeyPress, tb0208.KeyPress, tb0207.KeyPress, tb0206.KeyPress, tb0200.KeyPress, tb0199.KeyPress, tb0198.KeyPress, tb0197.KeyPress, tb0196.KeyPress, tb0185.KeyPress, tb0184.KeyPress, tb0183.KeyPress, tb0182.KeyPress, tb0181.KeyPress, tb0180.KeyPress, tb0179.KeyPress, tb0178.KeyPress, tb0177.KeyPress, tb0176.KeyPress, tb0170.KeyPress, tb0169.KeyPress, tb0168.KeyPress, tb0167.KeyPress, tb0166.KeyPress, tb0155.KeyPress, tb0154.KeyPress, tb0153.KeyPress, tb0152.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0151_TextChanged(sender As Object, e As EventArgs) Handles tb0151.TextChanged
        If (tb0151.Text <> "") Then
            tb0156.Enabled = True
            tb0161.Enabled = True
            tb0166.Enabled = True
            tb0171.Enabled = True
            tb0176.Enabled = True
        Else
            tb0156.Text = ""
            tb0161.Text = ""
            tb0166.Text = ""
            tb0171.Text = ""
            tb0176.Text = ""
            'desactivar
            tb0156.Enabled = False
            tb0161.Enabled = False
            tb0166.Enabled = False
            tb0171.Enabled = False
            tb0176.Enabled = False
        End If
    End Sub
    Private Sub tb0152_TextChanged(sender As Object, e As EventArgs) Handles tb0152.TextChanged
        If (tb0152.Text <> "") Then
            tb0157.Enabled = True
            tb0162.Enabled = True
            tb0167.Enabled = True
            tb0172.Enabled = True
            tb0177.Enabled = True
        Else
            tb0157.Text = ""
            tb0162.Text = ""
            tb0167.Text = ""
            tb0172.Text = ""
            tb0177.Text = ""
            'desactivar
            tb0157.Enabled = False
            tb0162.Enabled = False
            tb0167.Enabled = False
            tb0172.Enabled = False
            tb0177.Enabled = False
        End If
    End Sub
    Private Sub tb0153_TextChanged(sender As Object, e As EventArgs) Handles tb0153.TextChanged
        If (tb0153.Text <> "") Then
            tb0158.Enabled = True
            tb0163.Enabled = True
            tb0168.Enabled = True
            tb0173.Enabled = True
            tb0178.Enabled = True
        Else
            tb0158.Text = ""
            tb0163.Text = ""
            tb0168.Text = ""
            tb0173.Text = ""
            tb0178.Text = ""
            'desactivar
            tb0158.Enabled = False
            tb0163.Enabled = False
            tb0168.Enabled = False
            tb0173.Enabled = False
            tb0178.Enabled = False
        End If
    End Sub
    Private Sub tb0154_TextChanged(sender As Object, e As EventArgs) Handles tb0154.TextChanged
        If (tb0154.Text <> "") Then
            tb0159.Enabled = True
            tb0164.Enabled = True
            tb0169.Enabled = True
            tb0174.Enabled = True
            tb0179.Enabled = True
        Else
            tb0159.Text = ""
            tb0164.Text = ""
            tb0169.Text = ""
            tb0174.Text = ""
            tb0179.Text = ""
            'desactivar
            tb0159.Enabled = False
            tb0164.Enabled = False
            tb0169.Enabled = False
            tb0174.Enabled = False
            tb0179.Enabled = False
        End If
    End Sub
    Private Sub tb0155_TextChanged(sender As Object, e As EventArgs) Handles tb0155.TextChanged
        If (tb0155.Text <> "") Then
            tb0160.Enabled = True
            tb0165.Enabled = True
            tb0170.Enabled = True
            tb0175.Enabled = True
            tb0180.Enabled = True
        Else
            tb0160.Text = ""
            tb0165.Text = ""
            tb0170.Text = ""
            tb0175.Text = ""
            tb0180.Text = ""
            'desactivar
            tb0160.Enabled = False
            tb0165.Enabled = False
            tb0170.Enabled = False
            tb0175.Enabled = False
            tb0180.Enabled = False
        End If
    End Sub

    Private Sub tb0181_TextChanged(sender As Object, e As EventArgs) Handles tb0181.TextChanged
        If (tb0181.Text <> "") Then
            tb0186.Enabled = True
            tb0191.Enabled = True
            tb0196.Enabled = True
            tb0201.Enabled = True
            tb0206.Enabled = True
        Else
            tb0186.Text = ""
            tb0191.Text = ""
            tb0196.Text = ""
            tb0201.Text = ""
            tb0206.Text = ""
            'desactivar
            tb0186.Enabled = False
            tb0191.Enabled = False
            tb0196.Enabled = False
            tb0201.Enabled = False
            tb0206.Enabled = False
        End If
    End Sub
    Private Sub tb0182_TextChanged(sender As Object, e As EventArgs) Handles tb0182.TextChanged
        If (tb0182.Text <> "") Then
            tb0187.Enabled = True
            tb0192.Enabled = True
            tb0197.Enabled = True
            tb0202.Enabled = True
            tb0207.Enabled = True
        Else
            tb0187.Text = ""
            tb0192.Text = ""
            tb0197.Text = ""
            tb0202.Text = ""
            tb0207.Text = ""
            'desactivar
            tb0187.Enabled = False
            tb0192.Enabled = False
            tb0197.Enabled = False
            tb0202.Enabled = False
            tb0207.Enabled = False
        End If
    End Sub
    Private Sub tb0183_TextChanged(sender As Object, e As EventArgs) Handles tb0183.TextChanged
        If (tb0183.Text <> "") Then
            tb0188.Enabled = True
            tb0193.Enabled = True
            tb0198.Enabled = True
            tb0203.Enabled = True
            tb0208.Enabled = True
        Else
            tb0188.Text = ""
            tb0193.Text = ""
            tb0198.Text = ""
            tb0203.Text = ""
            tb0208.Text = ""
            'desactivar
            tb0188.Enabled = False
            tb0193.Enabled = False
            tb0198.Enabled = False
            tb0203.Enabled = False
            tb0208.Enabled = False
        End If
    End Sub
    Private Sub tb0184_TextChanged(sender As Object, e As EventArgs) Handles tb0184.TextChanged
        If (tb0184.Text <> "") Then
            tb0189.Enabled = True
            tb0194.Enabled = True
            tb0199.Enabled = True
            tb0204.Enabled = True
            tb0209.Enabled = True
        Else
            tb0189.Text = ""
            tb0194.Text = ""
            tb0199.Text = ""
            tb0204.Text = ""
            tb0209.Text = ""
            'desactivar
            tb0189.Enabled = False
            tb0194.Enabled = False
            tb0199.Enabled = False
            tb0204.Enabled = False
            tb0209.Enabled = False
        End If
    End Sub
    Private Sub tb0185_TextChanged(sender As Object, e As EventArgs) Handles tb0185.TextChanged
        If (tb0185.Text <> "") Then
            tb0190.Enabled = True
            tb0195.Enabled = True
            tb0200.Enabled = True
            tb0205.Enabled = True
            tb0210.Enabled = True
        Else
            tb0190.Text = ""
            tb0195.Text = ""
            tb0200.Text = ""
            tb0205.Text = ""
            tb0210.Text = ""
            'desactivar
            tb0190.Enabled = False
            tb0195.Enabled = False
            tb0200.Enabled = False
            tb0205.Enabled = False
            tb0210.Enabled = False
        End If
    End Sub

    Private Sub tb0211_TextChanged(sender As Object, e As EventArgs) Handles tb0211.TextChanged
        If (tb0211.Text <> "") Then
            tb0216.Enabled = True
            tb0221.Enabled = True
            tb0226.Enabled = True
            tb0231.Enabled = True
            tb0236.Enabled = True
        Else
            tb0216.Text = ""
            tb0221.Text = ""
            tb0226.Text = ""
            tb0231.Text = ""
            tb0236.Text = ""
            'desactivar
            tb0216.Enabled = False
            tb0221.Enabled = False
            tb0226.Enabled = False
            tb0231.Enabled = False
            tb0236.Enabled = False
        End If
    End Sub
    Private Sub tb0212_TextChanged(sender As Object, e As EventArgs) Handles tb0212.TextChanged
        If (tb0212.Text <> "") Then
            tb0217.Enabled = True
            tb0222.Enabled = True
            tb0227.Enabled = True
            tb0232.Enabled = True
            tb0237.Enabled = True
        Else
            tb0217.Text = ""
            tb0222.Text = ""
            tb0227.Text = ""
            tb0232.Text = ""
            tb0237.Text = ""
            'desactivar
            tb0217.Enabled = False
            tb0222.Enabled = False
            tb0227.Enabled = False
            tb0232.Enabled = False
            tb0237.Enabled = False
        End If
    End Sub
    Private Sub tb0213_TextChanged(sender As Object, e As EventArgs) Handles tb0213.TextChanged
        If (tb0213.Text <> "") Then
            tb0218.Enabled = True
            tb0223.Enabled = True
            tb0228.Enabled = True
            tb0233.Enabled = True
            tb0238.Enabled = True
        Else
            tb0218.Text = ""
            tb0223.Text = ""
            tb0228.Text = ""
            tb0233.Text = ""
            tb0238.Text = ""
            'desactivar
            tb0218.Enabled = False
            tb0223.Enabled = False
            tb0228.Enabled = False
            tb0233.Enabled = False
            tb0238.Enabled = False
        End If
    End Sub
    Private Sub tb0214_TextChanged(sender As Object, e As EventArgs) Handles tb0214.TextChanged
        If (tb0214.Text <> "") Then
            tb0219.Enabled = True
            tb0224.Enabled = True
            tb0229.Enabled = True
            tb0234.Enabled = True
            tb0239.Enabled = True
        Else
            tb0219.Text = ""
            tb0224.Text = ""
            tb0229.Text = ""
            tb0234.Text = ""
            tb0239.Text = ""
            'desactivar
            tb0219.Enabled = False
            tb0224.Enabled = False
            tb0229.Enabled = False
            tb0234.Enabled = False
            tb0239.Enabled = False
        End If
    End Sub
    Private Sub tb0215_TextChanged(sender As Object, e As EventArgs) Handles tb0215.TextChanged
        If (tb0215.Text <> "") Then
            tb0220.Enabled = True
            tb0225.Enabled = True
            tb0230.Enabled = True
            tb0235.Enabled = True
            tb0240.Enabled = True
        Else
            tb0220.Text = ""
            tb0225.Text = ""
            tb0230.Text = ""
            tb0235.Text = ""
            tb0240.Text = ""
            'desactivar
            tb0220.Enabled = False
            tb0225.Enabled = False
            tb0230.Enabled = False
            tb0235.Enabled = False
            tb0240.Enabled = False
        End If
    End Sub

    Private Sub tb0156_Leave(sender As Object, e As EventArgs) Handles tb0156.Leave, tb0175.Leave, tb0174.Leave, tb0173.Leave, tb0172.Leave, tb0171.Leave, tb0165.Leave, tb0164.Leave, tb0163.Leave, tb0162.Leave, tb0161.Leave, tb0160.Leave, tb0159.Leave, tb0158.Leave, tb0157.Leave
        If tb0156.Text <> "" And Not IsNumeric(tb0156.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0156.Text = ""
            tb0156.Focus()
        End If
        If tb0157.Text <> "" And Not IsNumeric(tb0157.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0157.Text = ""
            tb0157.Focus()
        End If
        If tb0158.Text <> "" And Not IsNumeric(tb0158.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0158.Text = ""
            tb0158.Focus()
        End If
        If tb0159.Text <> "" And Not IsNumeric(tb0159.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0159.Text = ""
            tb0159.Focus()
        End If
        If tb0160.Text <> "" And Not IsNumeric(tb0160.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0160.Text = ""
            tb0160.Focus()
        End If
        '----------
        If tb0161.Text <> "" And Not IsNumeric(tb0161.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0161.Text = ""
            tb0161.Focus()
        End If
        If tb0162.Text <> "" And Not IsNumeric(tb0162.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0162.Text = ""
            tb0162.Focus()
        End If
        If tb0163.Text <> "" And Not IsNumeric(tb0163.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0163.Text = ""
            tb0163.Focus()
        End If
        If tb0164.Text <> "" And Not IsNumeric(tb0164.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0164.Text = ""
            tb0164.Focus()
        End If
        If tb0165.Text <> "" And Not IsNumeric(tb0165.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0165.Text = ""
            tb0165.Focus()
        End If
        '------------
        If tb0171.Text <> "" And Not IsNumeric(tb0171.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0171.Text = ""
            tb0171.Focus()
        End If
        If tb0172.Text <> "" And Not IsNumeric(tb0172.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0172.Text = ""
            tb0172.Focus()
        End If
        If tb0173.Text <> "" And Not IsNumeric(tb0173.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0173.Text = ""
            tb0173.Focus()
        End If
        If tb0174.Text <> "" And Not IsNumeric(tb0174.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0174.Text = ""
            tb0174.Focus()
        End If
        If tb0175.Text <> "" And Not IsNumeric(tb0175.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0175.Text = ""
            tb0175.Focus()
        End If
    End Sub
    Private Sub tb0186_Leave(sender As Object, e As EventArgs) Handles tb0205.Leave, tb0204.Leave, tb0203.Leave, tb0202.Leave, tb0201.Leave, tb0195.Leave, tb0194.Leave, tb0193.Leave, tb0192.Leave, tb0191.Leave, tb0190.Leave, tb0189.Leave, tb0188.Leave, tb0187.Leave, tb0186.Leave
        If tb0186.Text <> "" And Not IsNumeric(tb0186.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0186.Text = ""
            tb0186.Focus()
        End If
        If tb0187.Text <> "" And Not IsNumeric(tb0187.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0187.Text = ""
            tb0187.Focus()
        End If
        If tb0188.Text <> "" And Not IsNumeric(tb0188.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0188.Text = ""
            tb0188.Focus()
        End If
        If tb0189.Text <> "" And Not IsNumeric(tb0189.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0189.Text = ""
            tb0189.Focus()
        End If
        If tb0190.Text <> "" And Not IsNumeric(tb0190.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0190.Text = ""
            tb0190.Focus()
        End If
        '------------
        If tb0191.Text <> "" And Not IsNumeric(tb0191.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0191.Text = ""
            tb0191.Focus()
        End If
        If tb0192.Text <> "" And Not IsNumeric(tb0192.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0192.Text = ""
            tb0192.Focus()
        End If
        If tb0193.Text <> "" And Not IsNumeric(tb0193.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0193.Text = ""
            tb0193.Focus()
        End If
        If tb0194.Text <> "" And Not IsNumeric(tb0194.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0194.Text = ""
            tb0194.Focus()
        End If
        If tb0195.Text <> "" And Not IsNumeric(tb0195.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0195.Text = ""
            tb0195.Focus()
        End If
        '----------------
        If tb0201.Text <> "" And Not IsNumeric(tb0201.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0201.Text = ""
            tb0201.Focus()
        End If
        If tb0202.Text <> "" And Not IsNumeric(tb0202.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0202.Text = ""
            tb0202.Focus()
        End If
        If tb0203.Text <> "" And Not IsNumeric(tb0203.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0203.Text = ""
            tb0203.Focus()
        End If
        If tb0204.Text <> "" And Not IsNumeric(tb0204.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0204.Text = ""
            tb0204.Focus()
        End If
        If tb0205.Text <> "" And Not IsNumeric(tb0205.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0205.Text = ""
            tb0205.Focus()
        End If
    End Sub
    Private Sub tb0216_Leave(sender As Object, e As EventArgs) Handles tb0235.Leave, tb0234.Leave, tb0233.Leave, tb0232.Leave, tb0231.Leave, tb0225.Leave, tb0224.Leave, tb0223.Leave, tb0222.Leave, tb0221.Leave, tb0220.Leave, tb0219.Leave, tb0218.Leave, tb0217.Leave, tb0216.Leave
        If tb0216.Text <> "" And Not IsNumeric(tb0216.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0216.Text = ""
            tb0216.Focus()
        End If
        If tb0217.Text <> "" And Not IsNumeric(tb0217.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0217.Text = ""
            tb0217.Focus()
        End If
        If tb0218.Text <> "" And Not IsNumeric(tb0218.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0218.Text = ""
            tb0218.Focus()
        End If
        If tb0219.Text <> "" And Not IsNumeric(tb0219.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0219.Text = ""
            tb0219.Focus()
        End If
        If tb0220.Text <> "" And Not IsNumeric(tb0220.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0220.Text = ""
            tb0220.Focus()
        End If
        '----------------
        If tb0221.Text <> "" And Not IsNumeric(tb0221.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0221.Text = ""
            tb0221.Focus()
        End If
        If tb0222.Text <> "" And Not IsNumeric(tb0222.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0222.Text = ""
            tb0222.Focus()
        End If
        If tb0223.Text <> "" And Not IsNumeric(tb0223.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0223.Text = ""
            tb0223.Focus()
        End If
        If tb0224.Text <> "" And Not IsNumeric(tb0224.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0224.Text = ""
            tb0224.Focus()
        End If
        If tb0225.Text <> "" And Not IsNumeric(tb0225.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0225.Text = ""
            tb0225.Focus()
        End If
        '------------
        If tb0231.Text <> "" And Not IsNumeric(tb0231.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0231.Text = ""
            tb0231.Focus()
        End If
        If tb0232.Text <> "" And Not IsNumeric(tb0232.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0232.Text = ""
            tb0232.Focus()
        End If
        If tb0233.Text <> "" And Not IsNumeric(tb0233.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0233.Text = ""
            tb0233.Focus()
        End If
        If tb0234.Text <> "" And Not IsNumeric(tb0234.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0234.Text = ""
            tb0234.Focus()
        End If
        If tb0235.Text <> "" And Not IsNumeric(tb0235.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0235.Text = ""
            tb0235.Focus()
        End If
    End Sub

    Private Sub tb0020_Leave(sender As Object, e As EventArgs) Handles tb0020.Leave
        If tb0020.Text <> "" And Not IsNumeric(tb0020.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0020.Text = ""
            tb0020.Focus()
        End If
    End Sub

    Private Sub tb0056_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0056.KeyPress, tb0051.KeyPress, tb0046.KeyPress, tb0041.KeyPress, tb0036.KeyPress, tb0031.KeyPress, tb0026.KeyPress, tb0021.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0057_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0057.KeyPress, tb0052.KeyPress, tb0047.KeyPress, tb0042.KeyPress, tb0037.KeyPress, tb0032.KeyPress, tb0027.KeyPress, tb0022.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            If (e.KeyChar = "1" Or e.KeyChar = "2") Then
                e.Handled = False
            Else
                e.Handled = True
            End If

        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0058_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0058.KeyPress, tb0053.KeyPress, tb0048.KeyPress, tb0043.KeyPress, tb0038.KeyPress, tb0033.KeyPress, tb0028.KeyPress, tb0023.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            If (e.KeyChar = "1" Or e.KeyChar = "2" Or e.KeyChar = "3" Or
                e.KeyChar = "4" Or e.KeyChar = "5") Then
                e.Handled = False
            Else
                e.Handled = True
            End If

        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0059_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0059.KeyPress, tb0054.KeyPress, tb0049.KeyPress, tb0044.KeyPress, tb0039.KeyPress, tb0034.KeyPress, tb0029.KeyPress, tb0024.KeyPress
        numericPoint(e)
    End Sub

    Private Sub tb0060_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0060.KeyPress, tb0055.KeyPress, tb0050.KeyPress, tb0045.KeyPress, tb0040.KeyPress, tb0035.KeyPress, tb0030.KeyPress, tb0025.KeyPress
        numeric(e)
    End Sub

    Private Sub tb0059_Leave(sender As Object, e As EventArgs) Handles tb0059.Leave, tb0054.Leave, tb0049.Leave, tb0044.Leave, tb0039.Leave, tb0034.Leave, tb0029.Leave, tb0024.Leave
        If tb0024.Text <> "" And Not IsNumeric(tb0024.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0024.Text = ""
            tb0024.Focus()
        End If
        If tb0029.Text <> "" And Not IsNumeric(tb0029.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0029.Text = ""
            tb0029.Focus()
        End If
        If tb0034.Text <> "" And Not IsNumeric(tb0034.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0034.Text = ""
            tb0034.Focus()
        End If
        If tb0039.Text <> "" And Not IsNumeric(tb0039.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0039.Text = ""
            tb0039.Focus()
        End If
        If tb0044.Text <> "" And Not IsNumeric(tb0044.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0044.Text = ""
            tb0044.Focus()
        End If
        If tb0049.Text <> "" And Not IsNumeric(tb0049.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0049.Text = ""
            tb0049.Focus()
        End If
        If tb0054.Text <> "" And Not IsNumeric(tb0054.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0054.Text = ""
            tb0054.Focus()
        End If
        If tb0059.Text <> "" And Not IsNumeric(tb0059.Text) Then
            MsgBox("Numero escrito no válido!")
            tb0059.Text = ""
            tb0059.Focus()
        End If
    End Sub

    Private Sub tb0021_TextChanged(sender As Object, e As EventArgs) Handles tb0021.TextChanged
        If (tb0021.Text <> "") Then
            tb0022.Enabled = True
            tb0023.Enabled = True
            tb0024.Enabled = True
            tb0025.Enabled = True
        Else
            tb0022.Text = ""
            tb0023.Text = ""
            tb0024.Text = ""
            tb0025.Text = ""
            'desactivar
            tb0022.Enabled = False
            tb0023.Enabled = False
            tb0024.Enabled = False
            tb0025.Enabled = False
        End If
    End Sub
    Private Sub tb0026_TextChanged(sender As Object, e As EventArgs) Handles tb0026.TextChanged
        If (tb0026.Text <> "") Then
            tb0027.Enabled = True
            tb0028.Enabled = True
            tb0029.Enabled = True
            tb0030.Enabled = True
        Else
            tb0027.Text = ""
            tb0028.Text = ""
            tb0029.Text = ""
            tb0030.Text = ""
            'desactivar
            tb0027.Enabled = False
            tb0028.Enabled = False
            tb0029.Enabled = False
            tb0030.Enabled = False
        End If
    End Sub
    Private Sub tb0031_TextChanged(sender As Object, e As EventArgs) Handles tb0031.TextChanged
        If (tb0031.Text <> "") Then
            tb0032.Enabled = True
            tb0033.Enabled = True
            tb0034.Enabled = True
            tb0035.Enabled = True
        Else
            tb0032.Text = ""
            tb0033.Text = ""
            tb0034.Text = ""
            tb0035.Text = ""
            'desactivar
            tb0032.Enabled = False
            tb0033.Enabled = False
            tb0034.Enabled = False
            tb0035.Enabled = False
        End If
    End Sub
    Private Sub tb0036_TextChanged(sender As Object, e As EventArgs) Handles tb0036.TextChanged
        If (tb0036.Text <> "") Then
            tb0037.Enabled = True
            tb0038.Enabled = True
            tb0039.Enabled = True
            tb0040.Enabled = True
        Else
            tb0037.Text = ""
            tb0038.Text = ""
            tb0039.Text = ""
            tb0040.Text = ""
            'desactivar
            tb0037.Enabled = False
            tb0038.Enabled = False
            tb0039.Enabled = False
            tb0040.Enabled = False
        End If
    End Sub
    Private Sub tb0041_TextChanged(sender As Object, e As EventArgs) Handles tb0041.TextChanged
        If (tb0041.Text <> "") Then
            tb0042.Enabled = True
            tb0043.Enabled = True
            tb0044.Enabled = True
            tb0045.Enabled = True
        Else
            tb0042.Text = ""
            tb0043.Text = ""
            tb0044.Text = ""
            tb0045.Text = ""
            'desactivar
            tb0042.Enabled = False
            tb0043.Enabled = False
            tb0044.Enabled = False
            tb0045.Enabled = False
        End If
    End Sub
    Private Sub tb0046_TextChanged(sender As Object, e As EventArgs) Handles tb0046.TextChanged
        If (tb0046.Text <> "") Then
            tb0047.Enabled = True
            tb0048.Enabled = True
            tb0049.Enabled = True
            tb0050.Enabled = True
        Else
            tb0047.Text = ""
            tb0048.Text = ""
            tb0049.Text = ""
            tb0050.Text = ""
            'desactivar
            tb0047.Enabled = False
            tb0048.Enabled = False
            tb0049.Enabled = False
            tb0050.Enabled = False
        End If
    End Sub
    Private Sub tb0051_TextChanged(sender As Object, e As EventArgs) Handles tb0051.TextChanged
        If (tb0051.Text <> "") Then
            tb0052.Enabled = True
            tb0053.Enabled = True
            tb0054.Enabled = True
            tb0055.Enabled = True
        Else
            tb0052.Text = ""
            tb0053.Text = ""
            tb0054.Text = ""
            tb0055.Text = ""
            'desactivar
            tb0052.Enabled = False
            tb0053.Enabled = False
            tb0054.Enabled = False
            tb0055.Enabled = False
        End If
    End Sub
    Private Sub tb0056_TextChanged(sender As Object, e As EventArgs) Handles tb0056.TextChanged
        If (tb0056.Text <> "") Then
            tb0057.Enabled = True
            tb0058.Enabled = True
            tb0059.Enabled = True
            tb0060.Enabled = True
        Else
            tb0057.Text = ""
            tb0058.Text = ""
            tb0059.Text = ""
            tb0060.Text = ""
            'desactivar
            tb0057.Enabled = False
            tb0058.Enabled = False
            tb0059.Enabled = False
            tb0060.Enabled = False
        End If
    End Sub

    Private Sub tbDepartamentoProductor_Leave(sender As Object, e As EventArgs) Handles tbDepartamentoProductor.Leave, tbDepartamentoProductor.TextChanged
        If tbDepartamentoProductor.Text <> "" Then
            If tbDepartamentoProductor.Text = "5" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m5
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "10" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m10
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "20" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m20
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "25" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m25
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "30" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m30
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "35" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m35
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "40" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m40
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "50" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m50
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "55" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m55
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "60" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m60
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "65" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m65
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "70" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m70
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "75" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m75
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "80" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m80
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "85" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m85
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "91" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m91
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            ElseIf tbDepartamentoProductor.Text = "93" Then
                tbMunicipioProductor.Items.Clear()
                For Each item As String In m93
                    tbMunicipioProductor.Items.Add(item)
                Next
                tbMunicipioProductor.Enabled = True

            Else
                tbMunicipioProductor.Enabled = False
                tbMunicipioProductor.Text = ""
            End If
        Else
            tbMunicipioProductor.Enabled = False
            tbMunicipioProductor.Text = ""
        End If
    End Sub

    Private Sub tbDepartamentoExplotacion_Leave(sender As Object, e As EventArgs) Handles tbDepartamentoExplotacion.Leave, tbDepartamentoExplotacion.TextChanged
        If tbDepartamentoExplotacion.Text <> "" Then
            If tbDepartamentoExplotacion.Text = "5" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m5
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "10" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m10
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "20" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m20
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "25" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m25
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "30" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m30
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "35" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m35
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "40" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m40
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "50" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m50
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "55" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m55
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "60" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m60
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "65" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m65
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "70" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m70
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "75" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m75
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "80" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m80
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "85" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m85
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "91" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m91
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            ElseIf tbDepartamentoExplotacion.Text = "93" Then
                tbMunicipioExplotacion.Items.Clear()
                For Each item As String In m93
                    tbMunicipioExplotacion.Items.Add(item)
                Next
                tbMunicipioExplotacion.Enabled = True

            Else
                tbMunicipioExplotacion.Enabled = False
                tbMunicipioExplotacion.Text = ""
            End If
        Else
            tbMunicipioExplotacion.Enabled = False
            tbMunicipioExplotacion.Text = ""
        End If
    End Sub

    Private Sub tbTelefonoProductor_Leave(sender As Object, e As EventArgs) Handles tbTelefonoProductor.Leave
        If tbTelefonoProductor.Text <> "" And tbTelefonoProductor.TextLength <> 8 Then
            MsgBox("formato de telefono no valido!")
            tbTelefonoProductor.Text = ""
            tbTelefonoProductor.Focus()
        End If
    End Sub

    Private Sub tbIdCuestionario_Leave(sender As Object, e As EventArgs) Handles tbIdCuestionario.Leave
        If (tbIdCuestionario.Text <> "") Then
            cmd = New SqlCommand
            cmd.Connection = con
            cmd.CommandText = "SELECT Encuesta.id_cuestionario FROM Encuesta WHERE id_cuestionario = '" & tbIdCuestionario.Text & "'"
            Dim reader As SqlDataReader = cmd.ExecuteReader()
            If reader.HasRows Then
                MsgBox("El N° ID del cuestionario ya existe!")
                tbIdCuestionario.Text = ""
            End If
            reader.Close()
        End If
    End Sub
End Class
