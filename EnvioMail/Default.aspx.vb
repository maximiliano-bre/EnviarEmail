
Partial Class _Default
    Inherits System.Web.UI.Page


    Dim correo As New GSMail
    Dim ToList As New System.Net.Mail.MailAddressCollection

    Dim ElementoGN_Origen As String = String.Empty

    Dim EmailOrigen As String = String.Empty

    Dim ToListAttachment As New Generic.Dictionary(Of String, String)

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim _FOLDERIMAGE As String = "C:\Inetpub\wwwroot\enviomail\"
        Dim FOLDERIMAGE As String = "C:\Inetpub\wwwroot\enviomail\"
        Dim strfile As String = "prueba.txt"
        Dim strExtension As String = ".txt"
        ToListAttachment.Add(_FOLDERIMAGE + strfile, FOLDERIMAGE + "prueba" + strfile)

        Try
            Dim archivo As String

            EmailOrigen = TextBox1.Text
            ToList.Add(New Net.Mail.MailAddress(TextBox2.Text)) 

            Dim host As String = "10.100.0.180"
            
            Dim hostport As String = "25"
           
            Dim SSL As Integer 
            Dim SSLuser As String = ""
            Dim SSLpass As String = ""

            'If SSL = 1 Then
            '    SSLuser = usuario
            '    SSLpass = pass
			'End If
            correo.Enviar(ToList, EmailOrigen, "Prueba", " Mensaje N° ", _
                "Body" + TextBox3.Text, host, True, hostport, SSL, SSLuser, SSLpass, ToListAttachment)
            Response.Write("Se envió el Correo correctamente")
        Catch ex As Exception

            LabelMensaje.Visible = True
            LabelMensaje.Text = ex.Message
        End Try

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LabelMensaje.Visible = False

    End Sub
End Class

