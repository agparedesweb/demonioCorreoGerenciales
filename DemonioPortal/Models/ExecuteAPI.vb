Imports System.Net
Imports System.Text
Imports System.Text.RegularExpressions
Imports RestSharp
Imports Sistema.Comunes.Registros.EscribanoRegistros
Imports Sistema.Comunes.Registros.FabricaRegistros

Public Class ExecuteAPI

    'Configuracion de Headers y Retorno de Llamado de la API
    Public Function MGet(url As String)
        Try
            Dim statusOK = 200 'Estatus HTTP 200
            
            ServicePointManager.Expect100Continue = True
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12

            Dim client = New RestClient()
            client.BaseUrl = New Uri(url)

            Dim request = New RestRequest()
            request.Method = Method.GET

            ' Configura las credenciales de autenticación Basic Authorization
            'se obtienen valores de tabla CONFIG_CONEXIONES        
            Dim user As String = fgObtenerConexionBD("header_userAPI")
            Dim password As String = fgObtenerConexionBD("header_passAPI")

            Dim base64Credentials As String = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{user}:{password}"))
            request.AddHeader("Authorization", "Basic " & base64Credentials)

            ' Timeout para Respuesta API
            client.Timeout = 30000

            Dim response = client.Execute(request)
            Console.WriteLine(vbCrLf & "### Se ejecuta Petición a la API ###" & vbCrLf & "HTTP StatusCode: " & response.StatusCode & vbCrLf)

            ' Verificar si la respuesta tiene éxito OK Status 200
            If response.StatusCode = statusOK Then
                Dim responseApi = client.Execute(request).Content.ToString()

                ' Validar si la respuesta de la API es un arreglo JSON vacío ([])
                If Regex.IsMatch(responseApi, "^\[\s*\]$") Then
                    Console.WriteLine(vbCrLf & "### ERROR ###")
                    Console.WriteLine("### La respuesta de la API está vacía. No Hay información con ese Rango de fecha ###" & vbCrLf)
                End If

                'Retornamos la respuesta de la API'
                Console.WriteLine("### Finaliza Ejecución de API ###" & vbCrLf)
                Return responseApi
            Else
                ' La solicitud a la API no tuvo éxito
                Console.WriteLine(vbCrLf & "### La solicitud a la API no tuvo éxito. HTTP StatusCode: " & response.StatusCode & " ###" & vbCrLf)
                Return "Error: " & response.StatusDescription

            End If
        Catch ex As Exception
            ' Ocurrió un error en la solicitud a la API
            Console.WriteLine(vbCrLf & "### Ocurrió un error en la solicitud a la API. ###" & vbCrLf)
            Console.WriteLine("Error: " & ex.Message)
            Return "Error: " & ex.Message
        End Try
    End Function




End Class
