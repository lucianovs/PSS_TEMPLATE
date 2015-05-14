Public Class frmPrincipal

    Private Sub frmPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Argumentos() As String = Environment.GetCommandLineArgs
        Dim sNomeUsuario As String

        'Ativar os Parâmetros iniciais de Segurança
        'Resgatar as Informações da Chamada
        If Environment.GetCommandLineArgs.Length > 1 Then
            g_Login = Environment.CommandLine(1)
        Else
            'MsgBox("Este programa não tem permissão para ser executado. Contactar o administrador da rede!!", MsgBoxStyle.Critical)
            'Application.Exit()

            'Parâmetros Padrão 9utilizar somente quando estiver em desenvolvimento
            g_Login = ClassCrypt.Encrypt("Admin")
        End If
        'Conection String
        g_ConnectString = (LerDadosINI(nomeArquivoINI(), "CONEXAO", "ConnectString", _
            ClassCrypt.Encrypt("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=SSVP.accdb;Persist Security Info=False;")))

        'Conectar com o Banco de dados
        If Not ConectarBanco() Then
            Application.Exit()
        End If

        'Ler o Usuário e Validar o Acesso
        sNomeUsuario = LerUsuario(ClassCrypt.Decrypt(g_Login))

        If sNomeUsuario <> "" Then
            Me.Text = "Modulo Tamplate - Usuário: " & sNomeUsuario
            GravarDadosIni("CONEXAO", "ConnectString", g_ConnectString)
        Else
            Application.Exit()
        End If


    End Sub

    Private Sub SairToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SairToolStripMenuItem.Click
        Application.Exit()
    End Sub

    Private Sub ConfiguraçõesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConfiguraçõesToolStripMenuItem.Click

    End Sub
End Class
