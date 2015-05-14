Imports System.Data.OleDb

Public Class Parametros
    Dim bAlterar As Boolean = False
    Dim Nivel_Opcao As Integer

    Private Sub Parametros_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        TratarObjetos()

    End Sub

    Private Sub TratarObjetos()

        tssContReg.Text = "Registro 1/1"

        '*** Botoes da Barra de comandos ***
        'Tag=1 -> Somente leitura
        'tag=2 -> Leitura + Alteração
        'Tag=3 -> Leitura + Alteração + Inclusão + Exclusão dos regs inserido pelo mesmo
        'Tag=4 -> Acesso Total
        '************************************
        btnIncluir.Enabled = False 'Not bAlterar And Me.Tag = 4
        btnAlterar.Enabled = Not bAlterar And Me.Tag = 4
        btnExcluir.Enabled = False 'Not bAlterar And Me.Tag = 4
        btnGravar.Enabled = bAlterar
        btnCancelar.Enabled = bAlterar
        btnAnterior.Enabled = False 'Not bAlterar
        btnProximo.Enabled = False 'Not bAlterar
        btnLocalizar.Enabled = False 'Not bAlterar
        btnImprimir.Enabled = Not bAlterar

        'Campos
        lblStrConexao.Enabled = bAlterar
        txtConnectString.Enabled = bAlterar
        lblPathReport.Enabled = bAlterar
        txtPathReport.Enabled = bAlterar
        lblBlocoReg.Enabled = bAlterar
        txtBlocoReg.Enabled = bAlterar

        'Preencher Campos
        txtConnectString.Text = ClassCrypt.Decrypt(g_ConnectString)
        txtPathReport.Text = LerDadosINI(nomeArquivoINI(), "PATH", "Reports", "C:\Fontes\SSVP_Projeto\Report\")
        txtBlocoReg.Text = LerDadosINI(nomeArquivoINI(), "BROWSE", "Bloco_Browse", 50)

    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click

        bAlterar = False
        Call TratarObjetos()

    End Sub

    Private Sub btnGravar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGravar.Click
        Dim cMensagem As String = ""

        If Validardados(cMensagem) Then
            g_ConnectString = ClassCrypt.Encrypt(txtConnectString.Text)
            GravarDadosIni("CONEXAO", "ConnectString", g_ConnectString)

            GravarDadosIni("PATH", "Reports", txtPathReport.Text)

            GravarDadosIni("BROWSE", "Bloco_Browse", txtBlocoReg.Text)

            bAlterar = False
            Call TratarObjetos()
        Else
            MsgBox(cMensagem)
        End If

    End Sub

    Private Function Validardados(ByRef cMensagem As String) As Boolean
        Dim cnConnectBanco = New OleDbConnection()
        Dim bValidOK As Boolean

        'Verificar se a conexão é válida 
        Try
            cnConnectBanco.ConnectionString = txtConnectString.Text
            cnConnectBanco.Open()
        Catch ex As Exception
            cMensagem = ex.Message & Chr(13) & "Contactar o administrador da rede."
            bValidOK = False
        Finally
        End Try
        If Trim(txtBlocoReg.Text) = "" Or Not IsNumeric(txtBlocoReg.Text) Then
            txtBlocoReg.Text = "50"
        End If

        If cMensagem = "" Then
            cnConnectBanco.Close()
            bValidOK = True
        End If
        Return bValidOK

    End Function

    Private Sub btnAlterar_Click(sender As Object, e As EventArgs) Handles btnAlterar.Click
        bAlterar = True
        Call TratarObjetos()
    End Sub

End Class