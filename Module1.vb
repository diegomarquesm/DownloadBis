Imports System.Net
Imports System.IO
Imports System.Configuration
Module Module1
    Private Function LerValorArquivoConfig(ByVal sValor As String) As String
        Dim sValorLido As String = ""
        sValorLido = configurationAppSettings.GetValue(sValor, GetType(System.String))
        Return sValorLido
    End Function
    Private Property http As String
    Private configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader()
    Private sDir As String = LerValorArquivoConfig("Diretorio do Arquivo a ser baixado")
    Private sNomeArquivo As String = LerValorArquivoConfig("Nome do Arquivo a ser baixado")
    Private sDirCompleto As String = sDir & "\" & sNomeArquivo
    Private sSiteBaixar As String = "http://200.98.141.29/executaveis/?action=download&file=L0JJUy5leGU="
    Private sSiteTesteConex As String = "https://www.google.com.br"
    Private myProxy As New WebProxy()
    Private Function ExisteConexaoInternet(ByVal sUrl As String) As Boolean
        Dim vRetorno As Boolean = False
        Dim vInURL As New System.Uri(sUrl)
        Dim vChamadaWeb As System.Net.WebRequest
        Dim vRespostaWeb As System.Net.WebResponse

        vChamadaWeb = System.Net.WebRequest.Create(sUrl)
        vChamadaWeb.Proxy = myProxy

        'vChamadaWeb.Credentials = New NetworkCredential("eas\diegoma", "p0w3r07")
        'myProxy.Credentials = New NetworkCredential(username, password)
        'myWebRequest.Proxy = myProxy

        Try
            vRespostaWeb = vChamadaWeb.GetResponse()
            vRetorno = True
        Catch ex As Exception
            vRetorno = False
        End Try

        Return vRetorno

    End Function
    Private Sub MoverArquivo()
        'Directory.CreateDirectory
        If File.Exists(sDirCompleto) Then
            If Not (Directory.Exists(sDir & "\BIS_bkp")) Then _
            Directory.CreateDirectory(sDir & "\BIS_bkp")
            File.Move(sDirCompleto, sDir & "\BIS_bkp\" & sNomeArquivo)
            File.Move(sDir & "\Temp\" & sNomeArquivo, sDir & "\" & sNomeArquivo)
            Directory.Delete(sDir & "\Temp")
            'File.Move(sDirCompleto & "\" & sNomeArquivo, sDir & "\BIS_bkp")
        End If
    End Sub
    Sub Main()
        Dim sDataModify As String
        sDataModify = FileDateTime(sDirCompleto)
        sDataModify = CDate(sDataModify).ToString("dd/MM/yyyy")

        If DateTime.Compare(sDataModify, DateTime.Today) < 0 Then
            'GoTo ir
            If Not (ExisteConexaoInternet(sSiteBaixar)) Then
                If (ExisteConexaoInternet(sSiteTesteConex)) Then
                    MsgBox("Nao existe atualizacao so sistema Bis")
                Else
                    MsgBox("Sem conex䯠com a internet, verifique as configura趥s")
                End If
            Else
                'ir:
                If Not (Directory.Exists(sDir & "\Temp")) Then Directory.CreateDirectory(sDir & "\Temp")
                'GoTo ir2

                Dim arquivo As String = "BIS.exe"
                Dim webClient As New WebClient()
                'webClient = System.Net.WebRequest.Create(sUrl)
                webClient.Proxy = myProxy
                'Console.WriteLine("Baixando arquivo ""{0}"" de ""{1}"" ......." _
                Dim stopwatch As Stopwatch = stopwatch.StartNew
                Console.WriteLine("Download iniciado...... ")
                Console.WriteLine(DateTime.Now.ToLongTimeString)
                Console.WriteLine("Baixando arquivo ""{0}"" do site da BIS ......." _
                 + ControlChars.Cr + ControlChars.Cr, arquivo, sSiteBaixar)
                Console.WriteLine("Aguarde um momento ..........")
                'webClient.DownloadFile(site & arquivo, "C:\" & arquivo)
                webClient.DownloadFile(sSiteBaixar, sDir & "\Temp\" & arquivo)
                Console.WriteLine("Arquivo foi baixado com sucesso.")
                Console.WriteLine(stopwatch.ElapsedMilliseconds)
                Console.WriteLine(DateTime.Now.ToLongTimeString)
                System.Threading.Thread.Sleep(5000)
                'ir2:            MoverArquivo()
                'Console.ReadKey()
            End If
            MoverArquivo()
        End If

    End Sub
End Module
