Imports System.IO
Imports System.Data.SqlClient


Public Class AMB_Empresa
    Dim ar As String
    Dim con As New SqlConnection("data source= " & CStr(leerArchivo(ar)) & "; initial catalog=Northwind; integrated security=true")


    Function leerArchivo(ByVal archivo As String) As String
        If File.Exists("c:\ABM\ip.txt") = True Then
            Dim SR As StreamReader = File.OpenText("c:\ABM\ip.txt")
            Dim Line As String = SR.ReadLine()
            SR.Close()
            Return Line
        Else
            MsgBox("Verifique si falta el archivo de configuración en el Disco C: ")
            Me.Close()

        End If
    End Function


    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
    End Sub

    Sub buscar(ByVal condicion As String)
        Dim da As New SqlDataAdapter("SELECT TOP (100) PERCENT ID, NomEmp from buscar_Empresa where" & condicion & "order by NomEmp", con)
        Dim ds As New DataSet
        da.Fill(ds, "Customers")
        If ds.Tables("Customers").Rows.Count = 0 Then 'si esto es 0 todas las grillas no se muestran, pq significa q no tiene nada

            DataGridView1.Visible = False
            pBotones.Visible = False
            pCampos.Visible = False
            lLegajo.Visible = False

        Else
            DataGridView1.DataSource = ds.Tables("Customers")
            DataGridView1.Refresh()
            DataGridView1.Visible = True
            lLegajo.Visible = True

        End If

    End Sub

    Private Sub AMB_Empresa_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        buscar(" NomEmp like '" & tBuscador.Text & "%' ")
    End Sub

    Sub FilaClick(ByVal e As Object)
        Dim fila As Integer = e.RowIndex
        Dim tfila As String

        If IsNothing(DataGridView1.Rows(fila).Cells(0).Value) Then
            lLegajo.Text = "0"
            pBotones.Visible = False
            pCampos.Visible = False
            Exit Sub
        Else
            tfila = DataGridView1.Rows(fila).Cells(0).Value
            lLegajo.Text = tfila.ToString()
            CargarCamposEmpresas()

        End If
    End Sub

    Sub CargarCamposEmpresas()
        If Val(lLegajo.Text) = 0 Then
            pBotones.Visible = False
            pCampos.Visible = False
            Exit Sub

        Else
            pBotones.Visible = True
            pCampos.Visible = True

            Dim da As New SqlDataAdapter("SELECT upper (ltrim(rtrim(isnull(CompanyName, '****'))))  as empresa,  ltrim(rtrim(isnull(ContactName,''))) as contacto, 
                ltrim(rtrim(isnull(ContactTitle,''))) as cargo, ltrim(rtrim(isnull(Address,''))) as direccion, 
              ltrim(rtrim(isnull(City,''))) as ciudad, ltrim(rtrim(isnull(Region,''))) as localidad, 
                ltrim(rtrim(isnull(PostalCode,''))) as CP, ltrim(rtrim(isnull(Country,''))) as pais,
               ltrim(rtrim(isnull(Phone,''))) as telefono, ltrim(rtrim(isnull(Fax,''))) as fax from Customers where ID=" & Val(lLegajo.Text), con)
            Dim ds As New DataSet
            da.Fill(ds, "Customers")

            TextBox1.Text = ds.Tables("Customers").Rows(0)("empresa")
            TextBox2.Text = ds.Tables("Customers").Rows(0)("contacto")
            TextBox3.Text = ds.Tables("Customers").Rows(0)("pais")
            TextBox4.Text = ds.Tables("Customers").Rows(0)("direccion")
            TextBox5.Text = ds.Tables("Customers").Rows(0)("cargo")
            TextBox6.Text = ds.Tables("Customers").Rows(0)("ciudad")
            TextBox7.Text = ds.Tables("Customers").Rows(0)("localidad")
            TextBox8.Text = ds.Tables("Customers").Rows(0)("cp")
            TextBox9.Text = ds.Tables("Customers").Rows(0)("telefono")
            TextBox10.Text = ds.Tables("Customers").Rows(0)("fax")

        End If
    End Sub

    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick
        FilaClick(e)
    End Sub

    Private Sub DataGridView1_RowEnter(ByVal sende As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.RowEnter
        FilaClick(e)
    End Sub

    Private Sub bBuscar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bBuscar.Click
        buscar(" NomEmp like '" & tBuscador.Text & "%' ")
    End Sub
    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        tBuscador.Text = ""
        buscar(" NomEmp like '" & tBuscador.Text & "%' ")
    End Sub

    Private Sub bEliminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bEliminar.Click
        If MessageBox.Show("Está por eliminar DEFINITIVAMENTE a la empresa: " & TextBox1.Text.Trim.ToUpper & ". Es algo EXTREMO, se eliminarán todas las referencias asociadas a esta. ¿Desea eliminarla?", "Eliminar empresa", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2) = Windows.Forms.DialogResult.No Then Exit Sub
        If SQL_Accion("delete from Northwind.dbo.Customers  where  ID=" & Val(lLegajo.Text)) = False Then
            MsgBox("Hubo un error al intentar borrar esta empresa. Por favor, reintente, y si el error persiste, anote todos los datos que quizo ingresar y comuníquese con la programadora.")
        Else
            buscar(" ID=" & Val(lLegajo.Text))
            MsgBox("La empresa fue eliminada de la base de datos.")
        End If
    End Sub

    Private Sub bGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bGuardar.Click

        Dim errores As String = "", en As String = vbCrLf
        If TextBox1.Text.Trim.Length < 3 Then
            errores &= "Debe completar el nombre de la empresa." & en
        End If
        If TextBox2.Text.Trim.Length < 3 Then
            errores &= "Debe completar el campo de Contactos correctamente. " & en
        End If

        If TextBox9.Text.Trim.Length < 3 Then
            errores &= "Verifique por favor el número de teléfono. No es obligatorio, pero si lo escribe debe ser correcto." & en
        End If

        If TextBox10.Text.Trim.Length < 3 Then
            errores &= "Verifique por favor el Fax. No es obligatorio, pero si lo escribe debe ser correcto." & en
        End If

        If errores.Length > 0 Then
            MsgBox("Hubo errores, por favor verifique y corrija antes de intentar de nuevo." & en & en & errores)
            Exit Sub
        End If


        If SQL_Accion("update Northwind.dbo.Customers set CompanyName='" & TextBox1.Text.Trim.ToUpper.Replace("'", "´") & "', ContactName='" & TextBox2.Text.Trim.ToUpper.Replace("'", "´") & "', ContactTitle='" & TextBox5.Text.Trim.ToUpper.Replace("'", "´") & "', Address='" & TextBox4.Text.Trim.Replace("'", "´") & "', City='" & TextBox6.Text.Trim.ToUpper.Replace("'", "´") & "', Region='" & TextBox7.Text.Trim.ToUpper.Replace("'", "´") & "', PostalCode='" & TextBox8.Text.Trim.ToUpper.Replace("'", "´") & "', Country='" & TextBox3.Text.Trim.ToUpper.Replace("'", "´") & "', Phone='" & TextBox9.Text.Trim.ToUpper.Replace("'", "´") & "', Fax=" & TextBox10.Text.Trim.ToUpper.Replace("'", "´") & "' where ID=" & VNum(lLegajo.Text)) = True Then

            MsgBox("Cambios realizados correctamente.")
            buscar(" ID=" & VNum(lLegajo.Text))
        Else
            MsgBox("Se produjo un error al querer guardar los datos de la empresa. Reintente, y si el error persiste, anote todos los datos que quizo ingresar y comuníquese con la programadora.")
        End If

    End Sub

    Private Sub bNuevoCliente_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bNuevoCliente.Click

        If SQL_Accion("insert into Northwind.dbo.Customers (CompanyName, ContactName, ContactTitle, Address, City, Region, PostalCode, Country, Phone, Fax) values ('*****',           '',           '',           '',           '',           '',           '',           '',             '',          '')  ") Then
            buscar(" NomEmp like '****%' ")
            MsgBox("Se ha creado un nuevo registro para la empresa que desea ingresar. Seleccione la línea nueva, cargue los datos y luego confirme con el botón 'Guardar cambios'.")
        End If
    End Sub

End Class
