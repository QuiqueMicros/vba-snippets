Attribute VB_Name = "modOutlook"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'                           >>> binarytopic.com <<<                            '
'                            coded by Diego F.C.                               '
'                                                                              '
'            http://binarytopic.com/enviar-correo-desde-excel-vba/             '
'                                                                              '
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

Function EnviarCorreo_Outlook( _
                                Para As String, _
                                Asunto As String, _
                                Cuerpo As String, _
                                Optional CC As String = vbNullString, _
                                Optional BCC As String = vbNullString, _
                                Optional Adjuntos As Collection, _
                                Optional Mostrar As Boolean = False _
                             ) As Boolean
'Envía correo a través de la aplicación Microsoft Outlook. Es necesario tener
'configurado Outlook en el equipo para poder usar esta función.
'
' ARGS:
'   Para: Destinatario directo en formato string. Puede ser una o varias
'       direcciones email separadas por ";".
'   Asunto: Mensaje de correo. String.
'   Cuerpo: Cuerpo del mensaje completo. String. El mail se enviará en texto
'       plano, es importante definir los saltos, tabulaciones...
'   CC: Destinatarios en copia. Tipo String. Opcional.
'   BCC: Destinatarios en copia oculta. Tipo String. Opcional.
'   Adjuntos: Colección con las rutas completas de los archivos a adjuntar al
'       correo. Tipo Colección. Opcional.
'   Mostrar: Parámetro que indica si el mensaje se envia automáticamente o si se
'       muestra en pantalla para realizar antes una revisión. Por defecto se envía
'       sin mostrar en pantalla. Boolean. Opcional.

    On Error GoTo ErrorEnvio
    
    'Definición de variables a usar
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim Adjunto As Variant

    'Inicialización de objetos
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)

    'Preparado de email
    With OutlookMail
        .To = Para
        If CC <> vbNullString Then .CC = CC
        If BCC <> vbNullString Then .BCC = BCC
        .Subject = Asunto
        .Body = Cuerpo
        If Not (Adjuntos Is Nothing) Then
            For Each Adjunto In Adjuntos
                .Attachments.Add (Adjunto)
            Next Adjunto
        End If
        
        If Mostrar Then
            .Display
        Else
            .Send
        End If
    End With
    
    'Limpieza
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    EnviarCorreo_Outlook = True
    Exit Function
    
ErrorEnvio:
    MsgBox "Error en el envío de correo:" & vbNewLine _
         & vbTab & Err.Number & ": " & Err.Description, vbCritical, "Error"
    On Error GoTo 0
    EnviarCorreo_Outlook = False
End Function
