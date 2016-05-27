Imports System.Data.SqlClient
Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn()
    End Sub

    Private Sub bGuardar_Click(sender As Object, e As EventArgs) Handles bGuardar.Click
        cmd = New SqlCommand
        cmd.Connection = con
        'Console.WriteLine(createQuery())
        query = createQuery()
        cmd.CommandText = query
        cmd.ExecuteNonQuery()
        MsgBox("Registro Agregado")
        'Funcion de limpieza
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
                panelCultivos.Enabled = True
            ElseIf (tb0407.Text = 2 Or tb0407.Text = "") Then
                panelCultivos.Enabled = False
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

    Private Sub tb0241_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241_TextChanged(sender As Object, e As EventArgs) Handles tb0241.TextChanged
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

    Private Sub tb0241MM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241MM.KeyPress
        If (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".") Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241JM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JM.KeyPress
        If (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".") Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241JLM_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JLM.KeyPress
        If (Char.IsDigit(e.KeyChar) Or e.KeyChar = ".") Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241MP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241MP.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241JP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JP.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub

    Private Sub tb0241JLP_KeyPress(sender As Object, e As KeyPressEventArgs) Handles tb0241JLP.KeyPress
        If (Char.IsDigit(e.KeyChar)) Then
            e.Handled = False
        ElseIf (Char.IsControl(e.KeyChar))
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
End Class
