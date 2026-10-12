Imports System
Imports System.Drawing
Imports System.Windows.Forms

Namespace AccessImporterVB
    Public Class FormPrincipal
        Inherits Form

        Private ReadOnly connAcc As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ruta\data\mibase.mdb;Jet OLEDB:Database Password=miclave;"
        Private ReadOnly connSQL As String = "Server=DESKTOP-S2GD18Tx7;Database=ExportDb;Integrated Security=true;"

        Private WithEvents btnProcesar As Button
        Private txtExpediente As TextBox
        Private txtResultado As TextBox

        Public Sub New()
            Me.Text = "AccessImporterVB"
            Me.Size = New Size(640, 520)
            Me.StartPosition = FormStartPosition.CenterScreen
            Me.InitializeComponent()
        End Sub

        Private Sub InitializeComponent()
            Me.txtExpediente = New TextBox()
            Me.btnProcesar = New Button()
            Me.txtResultado = New TextBox()

            Me.txtExpediente.Name = "txtExpediente"
            Me.txtExpediente.Text = "EXP-2026-001"
            Me.txtExpediente.Location = New Point(130, 30)
            Me.txtExpediente.Width = 220

            Me.btnProcesar.Name = "btnProcesar"
            Me.btnProcesar.Text = "Procesar"
            Me.btnProcesar.Location = New Point(380, 24)
            Me.btnProcesar.Width = 180
            Me.btnProcesar.Height = 36
            AddHandler Me.btnProcesar.Click, AddressOf Me.BtnProcesar_Click

            Me.txtResultado.Multiline = True
            Me.txtResultado.ScrollBars = ScrollBars.Vertical
            Me.txtResultado.ReadOnly = True
            Me.txtResultado.Location = New Point(15, 90)
            Me.txtResultado.Size = New Size(590, 360)

            Dim lblExpediente As New Label()
            lblExpediente.Text = "Expediente:"
            lblExpediente.Location = New Point(20, 35)
            lblExpediente.AutoSize = True

            Me.Controls.Add(lblExpediente)
            Me.Controls.Add(Me.txtExpediente)
            Me.Controls.Add(Me.btnProcesar)
            Me.Controls.Add(Me.txtResultado)
        End Sub

        Private Sub BtnProcesar_Click(sender As Object, e As EventArgs)
            Dim expediente = Me.txtExpediente.Text.Trim()
            If String.IsNullOrWhiteSpace(expediente) Then
                MessageBox.Show("Debe introducir un número de expediente.", "Validación", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Try
                Me.Log("Iniciando proceso...")
                AccessExporter.Saga2Adm(expediente, 1, Me.connAcc, Me.connSQL)
                Me.Log("Proceso finalizado correctamente.")
                MessageBox.Show("Proceso finalizado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                Me.Log("ERROR: " & ex.Message)
                MessageBox.Show("Error al procesar: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub Log(ByVal mensaje As String)
            Me.txtResultado.AppendText("[" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "] " & mensaje & Environment.NewLine)
        End Sub
    End Class
End Namespace
