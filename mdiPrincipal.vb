Imports System.Windows.Forms
Imports System.Data.OleDb
Imports System.IO

Public Class mdiPrincipal

    Private Sub mdiPrincipal_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Argumentos() As String = Environment.GetCommandLineArgs
        Dim sNomeUsuario As String

        g_Modulo = "Template"

        'Ativar os Parâmetros iniciais de Segurança
        'Resgatar as Informações da Chamada
        If Environment.GetCommandLineArgs.Length > 1 Then
            g_Login = Environment.CommandLine(1)
        Else
            'Ativar estas Linhas quando for colocar em produção
            'MsgBox("Este programa não tem permissão para ser executado. Contactar o administrador da rede!!", MsgBoxStyle.Critical)
            'Application.Exit()

            'Parâmetros Padrão - utilizar somente quando estiver em desenvolvimento
            'g_Login = (ClassCrypt.Encrypt("ssvp$00"))
            'MsgBox(g_Login)
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
        sNomeUsuario = LerUsuario(ClassCrypt.Decrypt(g_Login), Nothing)

        If sNomeUsuario <> "" Then
            Me.Text = "Modulo " & g_Modulo & " - Usuário: " & sNomeUsuario
        Else
            Application.Exit()
        End If

        'Verificar o acesso às opções do sistema
        Dim cModulo As Integer = getCodModulo(g_Modulo) 'Pegar o código do Módulo
        Dim nCodUsuario As Integer = getCodUsuario(ClassCrypt.Decrypt(g_Login)) 'pegar o código do usuario

        For Each _control As Object In Me.Controls
            If TypeOf (_control) Is MenuStrip Then
                For Each itm As ToolStripMenuItem In _control.items
                    If itm.Text <> "&Sair" And itm.Name.ToString.StartsWith("menu") Then
                        itm.Tag = NivelAcesso(nCodUsuario, cModulo, itm.Name, "")
                        itm.Enabled = itm.Tag > 0
                        'Função para Verificar os SubItens do menu
                        If itm.DropDownItems.Count > 0 Then LoopMenuItems(itm, nCodUsuario, cModulo, itm.Name)
                    End If
                Next
            End If
        Next


    End Sub

    Private Function LoopMenuItems(ByVal parent As ToolStripMenuItem, nCodUsuario As Integer, cModulo As Integer, fPrincOpcao As String) As Object
        Dim retval As Object = Nothing

        For Each child As Object In parent.DropDownItems

            'MessageBox.Show("Child : " & child.name)

            If TypeOf (child) Is ToolStripMenuItem Then
                If child.Text <> "Sair" And child.Name.ToString.StartsWith("menu") Then
                    child.Tag = NivelAcesso(nCodUsuario, cModulo, child.Name, fPrincOpcao)
                    child.Enabled = child.Tag > 0
                    If child.DropDownItems.Count > 0 Then
                        retval = LoopMenuItems(child, nCodUsuario, cModulo, child.name)
                        If Not retval Is Nothing Then Exit For
                    End If
                End If
            End If
        Next

        Return retval
    End Function

    'Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
    ' ' Create a new instance of the child form.
    ' Dim ChildForm As New System.Windows.Forms.Form
    ' ' Make it a child of this MDI form before showing it.
    '     ChildForm.MdiParent = Me
    '
    '        m_ChildFormNumber += 1
    '        ChildForm.Text = "Window " & m_ChildFormNumber
    '
    '        ChildForm.Show()
    '    End Sub

    Private Sub menuConfiguracoes_Click(sender As Object, e As EventArgs) Handles menuSisConfiguracoes.Click
        Dim ChildForm As New Parametros

        ChildForm.MdiParent = Me
        ChildForm.Tag = menuSisConfiguracoes.Tag 'é gravado no tag do menu o nível de acesso
        ChildForm.Show()

    End Sub

    Private Sub menuUsuarios_Click(sender As Object, e As EventArgs) Handles menuSisUsuarios.Click
        '?? Alterar os parâmetros para passar ao Browse (Entudade e Form. do Cadastro) ??
        Dim frmBrowse_Usuario As frmBrowse = New frmBrowse("ESI000", "frmUsuario")

        frmBrowse_Usuario.MdiParent = Me
        frmBrowse_Usuario.Tag = menuSisUsuarios.Tag 'é gravado no tag do menu o nível de acesso
        frmBrowse_Usuario.Text = menuSisUsuarios.Text
        frmBrowse_Usuario.Show()

    End Sub

    Private Sub menuRelatorios_Click(sender As Object, e As EventArgs) Handles menuRelatorios.Click

    End Sub

    Private Sub menuImpUsuario_Click(sender As Object, e As EventArgs) Handles menuRelUsuario.Click

    End Sub

    Private Sub CarregarUnidadesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CarregarUnidadesToolStripMenuItem.Click
        Dim arqCSV As New StreamReader("C:\Fontes\SSVP_Projeto\Documentos\DadosLegado\Unidade.dat")
        Dim LinhaArquivo As String
        Dim ArrayCampos() As String
        Dim ArrayDados() As String
        'Dim dt As DataTable = New DataTable("EUN000")
        Dim cmd As OleDbCommand
        Dim cSql As String = ""
        Dim cString As String
        Dim cValue As String = ""

        LinhaArquivo = arqCSV.ReadLine
        ArrayCampos = Split(LinhaArquivo, "|")
        While Not arqCSV.EndOfStream
            LinhaArquivo = arqCSV.ReadLine
            ArrayDados = Split(LinhaArquivo, "|")

            cSql = "INSERT INTO EUN000 ("
            cValue = " Values ("

            cSql += "UN000_CODRED,"
            cValue += ArrayDados(0).ToString & ","

            cSql += "UN000_NUMREG,"
            cValue += "0,"

            cSql += "UN000_CLAUNI,"
            cValue += "'" & ArrayDados(5).ToString & "',"

            cSql += "UN000_NOMUNI,"
            cString = Replace(ArrayDados(4).ToString, "'", "`")
            cValue += "'" & cString & "',"

            cSql += "UN000_DATFUN,"
            cValue += IIf(IsDate(ArrayDados(16).ToString), "'" & ArrayDados(16).ToString & "',", "null,")

            cSql += "UN000_CNPUNI,"
            cValue += "'',"

            cSql += "UN000_ENDUNI,"
            cString = Replace(ArrayDados(6).ToString, "'", "`")
            cValue += IIf(cString = "null", "null,", "'" & cString & "',")

            cSql += "UN000_BAIUNI,"
            cString = Replace(ArrayDados(8).ToString, "'", "`")
            cValue += IIf(cString = "null", "null,", "'" & cString & "',")

            cSql += "UN000_CEPUNI,"
            cString = Replace(ArrayDados(11).ToString, "'", "`")
            cValue += IIf(cString = "null", "null,", "'" & cString & "',")

            cSql += "UN000_CIDUNI,"
            cString = Replace(ArrayDados(9).ToString, "'", "`")
            cValue += IIf(cString = "null", "null,", "'" & cString & "',")

            cSql += "UN000_ESTUNI,"
            cString = Replace(ArrayDados(10).ToString, "'", "`")
            cValue += IIf(cString = "null", "null,", "'" & cString & "',")

            cSql += "UN000_DIOUNI,"
            cValue += "'',"

            cSql += "UN000_BCOUNI,"
            cValue += "'',"

            cSql += "UN000_AGEUNI,"
            cValue += "'',"

            cSql += "UN000_CCOUNI,"
            cValue += "'',"

            cSql += "UN000_TITUNI,"
            cValue += "'',"

            cSql += "UN000_OBSCCO,"
            cValue += "'',"

            cSql += "UN000_FREREU,"
            cValue += "'',"

            cSql += "UN000_APROCP,"
            cValue += "null,"

            cSql += "UN000_APROCC,"
            cValue += "null,"

            cSql += "UN000_APROCM,"
            cValue += "null,"

            cSql += "UN000_APROCN,"
            cValue += "null,"

            cSql += "UN000_APROCG,"
            cValue += "null,"

            cSql += "UN000_NIVUNI,"
            cValue += ArrayDados(3).ToString & ","

            cSql += "UN000_DATINS) "
            cValue += IIf(IsDate(ArrayDados(18).ToString), "'" & ArrayDados(18).ToString & "')", "null)")

            cSql += cValue 

            cmd = New OleDbCommand(cSql, g_ConnectBanco)

            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally
            End Try

        End While

        MsgBox("Importação de Unidades. Processo conluído !!")

    End Sub


End Class
