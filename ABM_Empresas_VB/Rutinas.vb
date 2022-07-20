Imports System.IO
Imports System.Data.SqlClient
Module Rutinas

    Function leerArchivo(ByVal archivo As String) As String
        If File.Exists("c:\ABM\ip.txt") = True Then
            Dim SR As StreamReader = File.OpenText("c:\ABM\ip.txt")
            Dim Line As String = SR.ReadLine()
            SR.Close()
            Return Line
        End If
    End Function

    Function YaExisteSQL(ByVal sql As String) As Boolean
        'sele pasa un select de sql que en teoria busca algo, por ejemplo un numero de cheque, si hay registros la fun devuelve 
        'true, estilo "el cheque ya existe", si no esta, devuelve false
        Dim ar As String
        Dim con As New SqlConnection("data source=" & CStr(leerArchivo(ar)) & "; initial catalog=Northwind; integrated security=true")
        Dim da1 As New SqlDataAdapter(sql, con)
        Dim ds1 As New DataSet
        da1.Fill(ds1, "afidesc")
        If ds1.Tables("afidesc").Rows.Count < 1 Then
            Return False
        Else
            Return True
        End If
    End Function

    Function SQL_Accion(ByVal Sql_de_accion As String) As Boolean
        'Ejecuta la consulta de accion 'sql_de_accion' abriendo la conexion automaticamente
        'se da cuenta si es de insert, update o delete, porque mira dentro de la sentencia que se le pasa
        'devuelve true si se pudo hacer, y false si hubo algún error
        Dim ar As String
        ' Dim conn As New SqlConnection(IIf(leerarchivo(ar).indexof("ULTRA") >= 0, "data source=ULTRABOOK\SQL2000;user id=sa;password=gato;initial catalog=icab", "data source=server01; initial catalog=icab; integrated security=true"))
        Dim conn As New SqlConnection("data source=" & CStr(leerarchivo(ar)) & "; initial catalog=Northwind; integrated security=true")

        Dim adapter As New SqlDataAdapter, salida As Boolean = True

        Try
            conn.Open()
            If Sql_de_accion.ToUpper.IndexOf("INSERT") Then
                'MsgBox(Sql_de_accion)
                adapter.InsertCommand = New SqlCommand(Sql_de_accion, conn)
                adapter.InsertCommand.ExecuteNonQuery()
            Else
                If Sql_de_accion.ToUpper.IndexOf("UPDATE") Then
                    adapter.UpdateCommand = New SqlCommand(Sql_de_accion, conn)
                    adapter.UpdateCommand.ExecuteNonQuery()
                Else
                    If Sql_de_accion.ToUpper.IndexOf("DELETE") Then
                        adapter.DeleteCommand = New SqlCommand(Sql_de_accion, conn)
                        adapter.DeleteCommand.ExecuteNonQuery()
                    Else
                        'esta mal la sintaxis porque no hay ni insert, ni delete ni update
                        salida = False
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox(ex.Message)
            salida = False
        End Try
        conn.Close()
        Return salida
    End Function

    Function VNum(ByVal NTexto As String) As Decimal
        'transforma un texto con algo, que puede ser un numero con coma o punto o perro, y devuelve un valor decimal siempre

        Return CDec(Val(NTexto.Trim.Replace(",", ".")))

    End Function

    Function NumSQL(ByVal numero As String) As String
        'Recibe un número desde un textbox por ejemplo, lo verifica como número válido, 
        'y luego le cambia la coma por punto para que sea válido en una sentencia de sql,
        'luego lo devuelve
        Return CStr(VNum(numero)).Trim.Replace(",", ".")
    End Function

    Function RellenaNum(ByVal numero As Integer, ByVal cantidad As Integer) As String
        'Rellena con 0s adelante el numero. Ideal para dias y meses:
        'RellenaNum(3,2)---> "03" RellenaNum(3,4)--->"0003"
        Dim snum As String = CStr(numero).Trim
        Return "00000000000000000000".Substring(0, cantidad - snum.Length) & snum
    End Function
    Function FechaSQL(ByVal fecha As Date) As String

        'Devuelve fecha en el formato 'AAAAMMDD'
        Return "'" & RellenaNum(Year(fecha), 4) & RellenaNum(Month(fecha), 2) & RellenaNum(fecha.Day, 2) & "'"

    End Function
















End Module
