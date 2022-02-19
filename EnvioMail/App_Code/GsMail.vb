Imports System.Net.Mail
Imports System.Net
Imports System.Net.Security
Imports System.Security.Cryptography.X509Certificates
Imports Ionic.Utils

Public Class GSMail
    Public Enum eTipoAttachment
        DeDisco = 1
        DeBase = 2
    End Enum
    Public Sub AddDestinatario(ByVal Direccion As String)
        m_Destinatarios.Add(Direccion)
    End Sub

    Private m_Destinatarios As Collections.Generic.List(Of String)

    Public Property Destinatarios() As System.Collections.Generic.List(Of String)
        Get
            Return m_Destinatarios
        End Get
        Set(ByVal value As System.Collections.Generic.List(Of String))
            m_Destinatarios = value
        End Set
    End Property

    Private m_TipoAttachment As eTipoAttachment
    Public Property TipoAttachment() As eTipoAttachment
        Get
            Return m_TipoAttachment
        End Get
        Set(ByVal value As eTipoAttachment)
            m_TipoAttachment = value
        End Set
    End Property

    Private m_ComprimirImagen As Boolean
    Public Property ComprimirImagen() As Boolean
        Get
            Return m_ComprimirImagen
        End Get
        Set(ByVal value As Boolean)
            m_ComprimirImagen = value
        End Set
    End Property

    Private m_Body As String = ""

    Public Property Body() As String
        Get
            Return m_Body
        End Get
        Set(ByVal value As String)
            m_Body = value
        End Set
    End Property

    ''' <summary>
    ''' Este método envía e-mails.
    ''' </summary>
    ''' <param name="MailDestinatarios">MailAddressCollection conteniendo objetos MailAddress para cada destinatario.</param>
    ''' <param name="MailDireccionOrigen">Dirección de origen.</param>
    ''' <param name="MailNombreOrigen">Nombre de origen.</param>
    ''' <param name="MailSubject">Título del email.</param>
    ''' <param name="MailBody">Cuerpo del email.</param>
    ''' <param name="MailHost">Dirección o Nro. IP del Host.</param>
    ''' <param name="MailIsBodyHtml">Flag para contenido HTML (default: true)</param>
    ''' <param name="MailPort">Puerto de salida (default: 25)</param>
    ''' ''' <param name="MailListAttachment">Coleccion de Rutas donde estan guardas las Imagenes del Procedimiento(Opcional)</param>
    ''' <remarks>Alejandro Hernández [09-2006 creado]/ FS 30-60-2008</remarks>
    Sub Enviar(ByVal MailDestinatarios As System.Net.Mail.MailAddressCollection, _
            ByVal MailDireccionOrigen As String, _
            ByVal MailNombreOrigen As String, _
            ByVal MailSubject As String, _
            ByVal MailBody As String, _
            ByVal MailHost As String, _
            Optional ByVal MailIsBodyHtml As Boolean = True, _
            Optional ByVal MailPort As Integer = 25, _
            Optional ByVal EnableSSL As Integer = 0, _
            Optional ByVal SSLuser As String = "", _
            Optional ByVal SSLpass As String = "", _
            Optional ByVal MailListAttachment As Generic.Dictionary(Of String, String) = Nothing)

        Dim vLstComprimidosABorrar As New List(Of String)
        'Se crea el mensaje
        Dim oMensaje As New System.Net.Mail.MailMessage

        'Se crea el cliente SMTP
        Dim oSMTP As New System.Net.Mail.SmtpClient()

        'Se agregan las direcciones
        oMensaje.From = New System.Net.Mail.MailAddress(MailDireccionOrigen, MailNombreOrigen)
        For Each addr As MailAddress In MailDestinatarios
            oMensaje.To.Add(addr)
        Next

        'Agrego los Attachment por si tiene imagenes del procedimiento o de la persona
        'FS - 27/6/2008
        If MailListAttachment IsNot Nothing Then
            If m_TipoAttachment = eTipoAttachment.DeDisco Then
                Dim aPair As KeyValuePair(Of String, String)
                For Each aPair In MailListAttachment
                    Dim valComprimir As String = aPair.Value 'Ruta del Archivo original con el nombre convertido para enviarlo (Apellido_1.jpg)
                    Dim strValOriginal As String = aPair.Key 'Ruta del Archivo original con el nombre convertido (26.06.2008_16054.jpg)
                    Dim valArchivoFinal As String = String.Empty
                    If m_ComprimirImagen = True Then
                        'Comprimo los Attachment
                        Dim zip As New Zip.ZipFile(valComprimir)
                        zip.AddFile(strValOriginal, "Imagen") '(Nombre del archivo a comprimir, nombre de la carpeta que contendra el Zip)
                        zip.Save()
                        valArchivoFinal = valComprimir
                    Else
                        'No comprimo nada
                        valArchivoFinal = strValOriginal
                    End If
                    Dim attach As New Attachment(valArchivoFinal)
                    oMensaje.Attachments.Add(attach)
                    attach = Nothing
                    vLstComprimidosABorrar.Add(valArchivoFinal) 'Agrego el achivo comprimido en un Array para borralo al final
                Next
            End If
        End If

        oMensaje.Subject = MailSubject
        oMensaje.Body = MailBody
        oMensaje.IsBodyHtml = MailIsBodyHtml
        'Seteo que el server notifique solamente en el error de entrega
        oMensaje.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure
        'smtp.UseDefaultCredentials = False
        'smtp.Credentials = CredentialCache.DefaultNetworkCredentials
        'servidor smtp (dirección IP ó nombre de host)
        oSMTP.Host = MailHost
        'Puerto a utilizar (25 por defecto)
        oSMTP.Port = MailPort

        'Si se solicitó SSL, lo activo
        If EnableSSL = 1 Then
            oSMTP.EnableSsl = True
            'Manganeta para validar el certificado aunque no esté registrado su ente emiso como autoridad certificante
            ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf ValidarCertificado)
        End If

        'Cargo las credenciales si hacen falta
        If SSLuser.Length <> 0 Then
            Dim credenciales As New System.Net.NetworkCredential(SSLuser, SSLpass)
            oSMTP.Credentials = credenciales
        End If

        Try
            oSMTP.Send(oMensaje)


        Catch ex As SmtpException
            Throw
        Catch ex As Exception
            Throw
        Finally
            oMensaje.Dispose()
            oMensaje = Nothing
            oSMTP = Nothing
            'Elimino los Archivos Zipeados en el Servidor
            For Each aPar As String In vLstComprimidosABorrar
                If ArchivoLockeado(aPar) Then
                    Dim fa As IO.FileAttributes = IO.File.GetAttributes(aPar)
                    If IO.FileAttributes.ReadOnly = IO.FileAttributes.ReadOnly Then IO.File.SetAttributes(aPar, IO.FileAttributes.Normal) 'Le seteo atributos normales al archivo para que me deje borrarlo
                    GC.Collect()
                End If
                System.IO.File.Delete(aPar) 'Borro el archivo Comprimido
            Next
        End Try
    End Sub
    Private Function ValidarCertificado(ByVal sender As Object, ByVal certificate As X509Certificate, ByVal chain As X509Chain, ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        'Manganeta para validar el certificado aunque no sea pulenta-pulenta, siempre retorna verdadero
        Return True
    End Function

    Private Function ArchivoLockeado(ByVal strFullFileName As String) As Boolean
        Dim blnReturn As Boolean = False
        Dim fs As System.IO.FileStream
        Try
            fs = System.IO.File.Open(strFullFileName, IO.FileMode.OpenOrCreate, _
                            IO.FileAccess.Read, IO.FileShare.None)
            fs.Close()
        Catch ex As System.IO.IOException
            blnReturn = True
        Finally
            fs = Nothing
        End Try
        Return blnReturn
    End Function

    Public Sub New()
        m_TipoAttachment = eTipoAttachment.DeDisco 'Inicio el tipo de Attachment por defecto en Disco (las imagenes estan almacenadas en el Disco Rigido)
        m_ComprimirImagen = True 'Por default comprimo las imagenes
    End Sub
End Class