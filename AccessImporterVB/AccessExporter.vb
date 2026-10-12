Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.OleDb

Namespace AccessImporterVB

    Public Class Cliente
        Public Property Codigo As String
        Public Property Nombre As String
        Public Property VatNumber As String
        Public Property Email As String
        Public Property EmpresaId As Long
    End Class

    Public Class Dossier
        Public Property NumeroExpediente As String
        Public Property CodigoCliente As String
        Public Property AnalyticAccount As Integer
        Public Property FechaEmbarque As Date?
        Public Property FechaLlegada As Date?
        Public Property BL As String
        Public Property CIF As Double
        Public Property Contenido As String
        Public Property Bultos As Long
        Public Property Buque As String
        Public Property Peso As Double
        Public Property Estatus As String
        Public Property EmpresaId As Long
    End Class

    Public Class Proveedor
        Public Property CodigoProveedor As String
        Public Property Nombre As String
        Public Property Email As String
        Public Property EmpresaId As Long
    End Class

    Public Class LineItem
        Public Property Producto As String
        Public Property Cantidad As Double
        Public Property CostoUnitario As Double
        Public Property CodigoArancelario As String
        Public Property Peso As Double
        Public Property FacturaComercial As String
    End Class

    Public Class Shipment
        Public Property EmpresaId As Long
        Public Property ShipmentId As String
        Public Property VendorId As String
        Public Property Proveedor As Proveedor
        Public Property Lineas As New List(Of LineItem)
    End Class

    Public Module AccessExporter

        Public Sub Saga2Adm(ByVal strNumExpediente As String, ByVal idEmpresa As Integer, ByVal strConnAccess As String, ByVal strConnSQLServer As String)
            If String.IsNullOrWhiteSpace(strNumExpediente) Then
                Throw New ArgumentException("Debe especificar un número de expediente válido.", NameOf(strNumExpediente))
            End If

            ExportarDesdeAccessASQLServer(strConnAccess, strConnSQLServer, strNumExpediente.Trim(), CLng(idEmpresa))
        End Sub

        Public Sub ExportarDesdeAccessASQLServer(ByVal strConnAccess As String,
                                                 ByVal strConnSQLServer As String,
                                                 ByVal strNumeroExpedienteBuscar As String,
                                                 ByVal id_empresa As Long)

            If String.IsNullOrWhiteSpace(strConnAccess) Then
                Throw New ArgumentException("Cadena de conexión Access vacía.", NameOf(strConnAccess))
            End If

            Using cnAccess As New OleDbConnection(strConnAccess)
                cnAccess.Open()

                Dim sqlDossier As String = "SELECT '' as x_analytic_account, expediente as x_name, codcli as x_Customer_ID, " &
                    "fechnoti as x_Boarding_date, feclleg as x_arrival_date, documento as x_bl, " &
                    "basimp as x_cif, ref as x_content, canbul as x_packages, nomveh as x_vessel, " &
                    "totkg as x_weight, facturado as x_status " &
                    "FROM davmae WHERE expediente = ?"

                Using cmdDossier As New OleDbCommand(sqlDossier, cnAccess)
                    cmdDossier.Parameters.AddWithValue("?", strNumeroExpedienteBuscar)

                    Using reader = cmdDossier.ExecuteReader()
                        If Not reader.Read() Then
                            Throw New InvalidOperationException($"No se encontró el expediente '{strNumeroExpedienteBuscar}'.")
                        End If

                        Dim dossier As New Dossier With {
                            .NumeroExpediente = SafeGetString(reader, "x_name"),
                            .CodigoCliente = SafeGetString(reader, "x_Customer_ID"),
                            .AnalyticAccount = SafeGetIntOrDefault(reader, "x_analytic_account", 0),
                            .FechaEmbarque = SafeGetDateNullable(reader, "x_Boarding_date"),
                            .FechaLlegada = SafeGetDateNullable(reader, "x_arrival_date"),
                            .BL = SafeGetString(reader, "x_bl"),
                            .CIF = SafeGetDoubleOrDefault(reader, "x_cif", 0),
                            .Contenido = SafeGetString(reader, "x_content"),
                            .Bultos = SafeGetLongOrDefault(reader, "x_packages", 0),
                            .Buque = SafeGetString(reader, "x_vessel"),
                            .Peso = SafeGetDoubleOrDefault(reader, "x_weight", 0),
                            .EmpresaId = id_empresa
                        }

                        Dim statusBool = SafeGetBoolOrDefault(reader, "x_status", False)
                        dossier.Estatus = If(statusBool, "Abierto", "Cerrado")

                        Dim cliente As Cliente = Nothing
                        Dim sqlCliente As String = "SELECT codcli as x_Customer_ID, nombre, rif as vat_number, email FROM Clientes WHERE codcli = ?"
                        Using cmdCliente As New OleDbCommand(sqlCliente, cnAccess)
                            cmdCliente.Parameters.AddWithValue("?", dossier.CodigoCliente)
                            Using rCli = cmdCliente.ExecuteReader()
                                If rCli.Read() Then
                                    cliente = New Cliente With {
                                        .Codigo = SafeGetString(rCli, "x_Customer_ID"),
                                        .Nombre = SafeGetString(rCli, "nombre"),
                                        .VatNumber = SafeGetString(rCli, "vat_number"),
                                        .Email = SafeGetString(rCli, "email"),
                                        .EmpresaId = id_empresa
                                    }
                                End If
                            End Using
                        End Using

                        Dim embarques As New List(Of Shipment)()
                        Dim sqlShipments As String = "SELECT id, expediente as x_shipment_id, codpro as x_vendor_id FROM davdet WHERE expediente = ?"
                        Using cmdShip As New OleDbCommand(sqlShipments, cnAccess)
                            cmdShip.Parameters.AddWithValue("?", strNumeroExpedienteBuscar)
                            Using rShip = cmdShip.ExecuteReader()
                                While rShip.Read()
                                    Dim shipment As New Shipment With {
                                        .EmpresaId = id_empresa,
                                        .ShipmentId = SafeGetString(rShip, "x_shipment_id"),
                                        .VendorId = SafeGetString(rShip, "x_vendor_id")
                                    }

                                    Dim prov As Proveedor = Nothing
                                    Dim vendorId = If(String.IsNullOrWhiteSpace(shipment.VendorId), String.Empty, shipment.VendorId.Trim())

                                    If Not String.IsNullOrEmpty(vendorId) Then
                                        Dim sqlProv = "SELECT TOP 1 codpro, nompro, email FROM tabpro WHERE codpro = ?"
                                        Using cmdProv As New OleDbCommand(sqlProv, cnAccess)
                                            cmdProv.Parameters.AddWithValue("?", vendorId)
                                            Using rProv = cmdProv.ExecuteReader()
                                                If rProv.Read() Then
                                                    prov = New Proveedor With {
                                                        .CodigoProveedor = SafeGetString(rProv, "codpro"),
                                                        .Nombre = SafeGetString(rProv, "nompro"),
                                                        .Email = SafeGetString(rProv, "email"),
                                                        .EmpresaId = id_empresa
                                                    }
                                                Else
                                                    prov = New Proveedor With {
                                                        .Nombre = "PROVEEDOR DESCONOCIDO",
                                                        .CodigoProveedor = vendorId,
                                                        .Email = String.Empty,
                                                        .EmpresaId = id_empresa
                                                    }
                                                End If
                                            End Using
                                        End Using
                                    Else
                                        prov = New Proveedor With {
                                            .Nombre = "SIN PROVEEDOR",
                                            .CodigoProveedor = "N/A",
                                            .Email = String.Empty,
                                            .EmpresaId = id_empresa
                                        }
                                    End If

                                    shipment.Proveedor = prov

                                    Dim sqlLines As String = "SELECT expediente as x_shipment_id, codara as x_tariff_code, nomcom as x_product, " &
                                        "codpro as x_vendor_id, cant as x_quantity, totaluni as x_unit_cost, 0 as x_weight, '' as x_commercial_invoice FROM davmerc " &
                                        "WHERE expediente = ? AND codpro = ?"
                                    Using cmdLines As New OleDbCommand(sqlLines, cnAccess)
                                        cmdLines.Parameters.AddWithValue("?", strNumeroExpedienteBuscar)
                                        cmdLines.Parameters.AddWithValue("?", vendorId)
                                        Using rLines = cmdLines.ExecuteReader()
                                            While rLines.Read()
                                                Dim line As New LineItem With {
                                                    .Producto = SafeGetString(rLines, "x_product"),
                                                    .Cantidad = SafeGetDoubleOrDefault(rLines, "x_quantity", 0),
                                                    .CostoUnitario = SafeGetDoubleOrDefault(rLines, "x_unit_cost", 0),
                                                    .CodigoArancelario = SafeGetString(rLines, "x_tariff_code"),
                                                    .Peso = SafeGetDoubleOrDefault(rLines, "x_weight", 0),
                                                    .FacturaComercial = SafeGetString(rLines, "x_commercial_invoice")
                                                }
                                                shipment.Lineas.Add(line)
                                            End While
                                        End Using
                                    End Using

                                    embarques.Add(shipment)
                                End While
                            End Using
                        End Using

                        If String.Equals(dossier.Estatus, "Cerrado", StringComparison.OrdinalIgnoreCase) Then
                            ' Aquí puedes invocar la lógica de SQL Server o dejarla para un siguiente paso.
                            ' ProcessFullImportFlow(strConnSQLServer, cliente, dossier, embarques)
                        Else
                            Throw New InvalidOperationException($"El expediente '{strNumeroExpedienteBuscar}' no está 'Cerrado' (Estatus: {dossier.Estatus}).")
                        End If
                    End Using
                End Using
            End Using
        End Sub

        Private Function SafeGetString(ByVal r As IDataRecord, ByVal name As String) As String
            Dim i = GetOrdinalSafe(r, name)
            If i = -1 OrElse r.IsDBNull(i) Then Return String.Empty
            Return Convert.ToString(r.GetValue(i))
        End Function

        Private Function SafeGetDoubleOrDefault(ByVal r As IDataRecord, ByVal name As String, ByVal defValue As Double) As Double
            Dim i = GetOrdinalSafe(r, name)
            If i = -1 OrElse r.IsDBNull(i) Then Return defValue
            Return Convert.ToDouble(r.GetValue(i))
        End Function

        Private Function SafeGetLongOrDefault(ByVal r As IDataRecord, ByVal name As String, ByVal defValue As Long) As Long
            Dim i = GetOrdinalSafe(r, name)
            If i = -1 OrElse r.IsDBNull(i) Then Return defValue
            Return Convert.ToInt64(r.GetValue(i))
        End Function

        Private Function SafeGetIntOrDefault(ByVal r As IDataRecord, ByVal name As String, ByVal defValue As Integer) As Integer
            Dim i = GetOrdinalSafe(r, name)
            If i = -1 OrElse r.IsDBNull(i) Then Return defValue
            Return Convert.ToInt32(r.GetValue(i))
        End Function

        Private Function SafeGetBoolOrDefault(ByVal r As IDataRecord, ByVal name As String, ByVal defValue As Boolean) As Boolean
            Dim i = GetOrdinalSafe(r, name)
            If i = -1 OrElse r.IsDBNull(i) Then Return defValue
            Dim v = r.GetValue(i)
            If TypeOf v Is Boolean Then Return DirectCast(v, Boolean)
            Try
                Return Convert.ToBoolean(v)
            Catch
                Return defValue
            End Try
        End Function

        Private Function SafeGetDateNullable(ByVal r As IDataRecord, ByVal name As String) As Date?
            Dim i = GetOrdinalSafe(r, name)
            If i = -1 OrElse r.IsDBNull(i) Then Return Nothing
            Return Convert.ToDateTime(r.GetValue(i))
        End Function

        Private Function GetOrdinalSafe(ByVal r As IDataRecord, ByVal name As String) As Integer
            Try
                Return r.GetOrdinal(name)
            Catch
                Return -1
            End Try
        End Function

    End Module
End Namespace
