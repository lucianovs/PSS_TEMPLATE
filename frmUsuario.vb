Imports System.Data.OleDb
Imports System.Drawing.Printing

Public Class frmUsuario
    '?? Alterar para a Entidade Principal ??
    Dim dt As DataTable = New DataTable("ESI000")
    Dim dtGrupo As DataTable = New DataTable("ESI001")

    Dim i As Integer
    Dim bAlterar As Boolean = False
    Dim bIncluir As Boolean = False
    Dim cQuery As String

    Private Sub frmUsuario_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        g_Param(1) = getCodUsuario(txtLogin.Text) 'Voltar com a Chave do registro do formulário
    End Sub

    'Public Sub New(ByVal ValorForm As String)
    '    Call InitializeComponent()
    '    MsgBox(ValorForm)
    'End Sub

    Private Sub frmUsuario_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim cQuery_Grupo As String
        Dim i_point As Integer

        ToolStrip1.ShowItemToolTips = True

        'Criar um adaptador que vai fazer o download de dados da base de dados
        '?? Alterar o Código para a Entidade Principal ??
        If Me.Tag = 4 Or g_Param(1) = "INSERT" Then
            cQuery = "SELECT * FROM ESI000"
        Else
            cQuery = "SELECT * FROM ESI000 where SI000_CODUSU = " & g_Param(1)
        End If

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dt)
        End Using

        If g_Param(1) <> "INSERT" Then
            'Posicionar no registro selecionado
            '?? Alterar para localizar a chave da tabela ??
            For i_point = 0 To dt.Rows.Count() - 1
                If dt.Rows(i_point).Item("SI000_CODUSU").ToString = g_Param(1) Then
                    Exit For
                End If
            Next
            i = i_point

            'Iniciar com o comando passado
            If g_Comando = "incluir" Then
                bIncluir = True
                bAlterar = True
            ElseIf g_Comando = "alterar" Then
                bIncluir = False
                bAlterar = True
            Else
                bIncluir = False
                bAlterar = False
            End If
        Else
            bIncluir = True
            bAlterar = True
        End If

        'Carregar o Combo de Grupos
        cQuery_Grupo = "SELECT * FROM ESI001"
        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery_Grupo, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtGrupo)
            If dtGrupo.Rows.Count() > 0 Then
                For x = 0 To dtGrupo.Rows.Count() - 1
                    cbGrupoPrincipal.Items.Add(dtGrupo.Rows(x).Item("SI001_DESGRU"))
                Next
            Else
                cbGrupoPrincipal.Items.Clear()
            End If
        End Using
        dtGrupo.Clear()
        '**************

        TratarObjetos()

    End Sub

    Private Sub CarregarListBox_Grupos()
        Dim cquery_grupo As String

        lstGrupoAssoc.Items.Clear()
        lstGrupoAssoc.Items.Add("GRUPOS ASSOCIADOS")

        If cbGrupoPrincipal.Text <> "" Then
            lstGrupoAssoc.Items.Add(cbGrupoPrincipal.Text)
        End If

        If Not bIncluir Then
            'Carregar o Combo de Grupos Associados ao usuário
            cquery_grupo = "SELECT ESI001.SI001_DESGRU, ESI006.SI006_CODGRU FROM (ESI001 " & _
                    "INNER JOIN ESI006 ON ESI006.SI006_CODGRU=ESI001.SI001_CODGRU) where ESI006.SI006_CODUSU= " & _
                    getCodUsuario(dt.Rows(i).Item("SI000_LGIUSU"))
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cquery_grupo, g_ConnectBanco)

                ' Preencher o DataTable 
                da.Fill(dtGrupo)
                If dtGrupo.Rows.Count() > 0 Then
                    For x = 0 To dtGrupo.Rows.Count() - 1
                        If cbGrupoPrincipal.Text <> dtGrupo.Rows(x).Item("SI001_DESGRU") Then
                            lstGrupoAssoc.Items.Add(dtGrupo.Rows(x).Item("SI001_DESGRU"))
                        End If
                    Next
                End If
            End Using
            dtGrupo.Clear()
        End If

        lstGrupoAssoc.Items.Add("")
        lstGrupoAssoc.Items.Add("GRUPOS A ASSOCIAR")

        'Carregar o Combo de Grupos Não Associados 
        If bIncluir Then
            cquery_grupo = "SELECT ESI001.SI001_DESGRU, ESI001.SI001_CODGRU FROM ESI001 " 
        Else
            cquery_grupo = "SELECT ESI001.SI001_DESGRU FROM ESI001 " & _
                    "where not exists(select * from ESI006 where ESI006.SI006_CODGRU=ESI001.SI001_CODGRU " & _
                        " AND ESI006.SI006_CODUSU = " & getCodUsuario(dt.Rows(i).Item("SI000_LGIUSU")) & ")"
            '"where ESI006.SI006_CODGRU is null AND ESI006.SI006_CODUSU = " & getCodUsuario(dt.Rows(i).Item("SI000_LGIUSU"))
        End If

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cquery_grupo, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtGrupo)
            If dtGrupo.Rows.Count() > 0 Then
                For x = 0 To dtGrupo.Rows.Count() - 1
                    If cbGrupoPrincipal.Text <> dtGrupo.Rows(x).Item("SI001_DESGRU") Then
                        lstGrupoAssoc.Items.Add(dtGrupo.Rows(x).Item("SI001_DESGRU"))
                    End If
                Next
            End If
        End Using
        dtGrupo.Clear()

    End Sub

    Private Sub TratarObjetos()

        tssContReg.Text = "Registro " & (i + 1).ToString & "/" & dt.Rows.Count().ToString

        'Botoes da Barra de comandos
        btnIncluir.Enabled = Not bAlterar And Me.Tag = 4 'And Me.Tag > 1
        btnAlterar.Enabled = Not bAlterar 'And Me.Tag > 1
        btnExcluir.Enabled = Not bAlterar And Me.Tag = 4
        btnGravar.Enabled = bAlterar
        btnCancelar.Enabled = bAlterar
        btnAnterior.Enabled = Not bAlterar
        btnProximo.Enabled = Not bAlterar
        btnLocalizar.Enabled = Not bAlterar
        btnImprimir.Enabled = Not bAlterar

        'Campos
        '?? Alterar para os seus objetos da Tela ??
        lblLogin.Enabled = bIncluir
        txtLogin.Enabled = bIncluir
        lblNmUsuario.Enabled = bAlterar And Me.Tag = 4 'And Me.Tag > 1
        txtNmUsuario.Enabled = bAlterar And Me.Tag = 4 'And Me.Tag > 1
        lblSenha.Enabled = bAlterar
        lblSenha2.Enabled = bAlterar
        txtSenha.Enabled = bAlterar
        txtSenha2.Enabled = bAlterar
        lblGrupo.Enabled = bAlterar
        btnLocalizar.Enabled = bAlterar
        txtColaborador.Enabled = bAlterar
        lblColaborador.Enabled = bAlterar

        'Outros Controles
        cbGrupoPrincipal.Enabled = bAlterar And Me.Tag = 4 'And Me.Tag > 1
        lstGrupoAssoc.Enabled = bAlterar And Me.Tag = 4
        lblGrupoAssoc.Enabled = bAlterar And Me.Tag = 4
        dtpExpira.Enabled = bAlterar And Me.Tag = 4
        lblDtpExpira.Enabled = bAlterar And Me.Tag = 4
        chkAlterarSenha.Enabled = bAlterar And Me.Tag = 4
        chkGerAgr.Enabled = bAlterar And Me.Tag = 4
        chkValidade.Enabled = bAlterar And Me.Tag = 4
        '*****************

        'Preencher Campos
        If i > -1 And Not bIncluir Then
            txtLogin.Text = dt.Rows(i).Item("SI000_LGIUSU")
            txtNmUsuario.Text = dt.Rows(i).Item("SI000_NOMUSU")
            txtSenha.Text = "**********" 'ClassCrypt.Decrypt(IIf(IsDBNull(dt.Rows(i).Item("SI000_PASLGI")), "", dt.Rows(i).Item("SI000_PASLGI")))
            txtSenha2.Text = ""
            dtpExpira.Value = CDate(dt.Rows(i).Item("SI000_DATEXP"))
            chkValidade.Checked = Format(dtpExpira.Value, "dd/MM/yyyy") = "01/01/1900"
            'If chkValidade.Checked Then dtpExpira.Text = ""
            dtpExpira.Enabled = Not chkValidade.Checked And bAlterar
            lblDtpExpira.Enabled = Not chkValidade.Checked And bAlterar

            chkAlterarSenha.Checked = dt.Rows(i).Item("SI000_ALTPAS") = 1
            chkGerAgr.Checked = IIf(IsDBNull(dt.Rows(i).Item("SI000_GERAGR")), False, dt.Rows(i).Item("SI000_GERAGR") = 1)

            If Not IsDBNull(dt.Rows(i).Item("SI000_CODASS")) Then
                txtColaborador.Text = Format(dt.Rows(i).Item("SI000_CODASS"), "000000") & " - " & LerNome_Colaborador(dt.Rows(i).Item("SI000_CODASS"))
            Else
                txtColaborador.Text = ""
            End If

            'Outros Controles
            cbGrupoPrincipal.Text = getDescrGrupo(dt.Rows(i).Item("SI000_CODGRU"))
        End If

        'Outras Chamadas
        CarregarListBox_Grupos()



    End Sub

    Private Sub btnProximo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnProximo.Click

        i += 1
        If Not i = dt.Rows.Count() Then
            Call TratarObjetos()
        Else
            i = dt.Rows.Count() - 1
        End If

    End Sub

    Private Sub btnAnterior_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAnterior.Click

        i -= 1
        If Not i < 0 Then
            Call TratarObjetos()
        Else
            i = 0
        End If

    End Sub

    Private Sub btnAlterar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAlterar.Click

        bAlterar = True
        Call TratarObjetos()

    End Sub

    Private Sub btnCancelar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelar.Click

        If g_Comando = "incluir" Or g_Comando = "alterar" Then
            dt.Clear()
            Me.Close()
        Else
            bAlterar = False
            bIncluir = False
            TratarObjetos()
        End If
    End Sub

    Private Sub btnIncluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIncluir.Click
        bAlterar = True
        bIncluir = True

        'Inicializar os seus Componentes de Entrada de Dados
        txtLogin.Text = ""
        txtNmUsuario.Text = ""
        txtSenha.Text = ""
        txtSenha2.Text = ""

        Call TratarObjetos()

    End Sub

    Private Sub btnGravar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGravar.Click
        Dim cSql As String = ""
        Dim cMensagem As String = ""
        Dim cmd As OleDbCommand
        Dim nProxCod_Usuario As Integer

        If ConectarBanco() Then
            '?? Colocar o Comando SQL para Gravar (Update e Insert)
            If Validardados(cMensagem) Then
                If bIncluir Then
                    nProxCod_Usuario = ProxCodChave("ESI000", "SI000_CODUSU")
                    cSql = "INSERT INTO ESI000(SI000_CODUSU, SI000_LGIUSU, SI000_PASLGI, " & _
                        "SI000_NOMUSU, SI000_CODGRU, SI000_DATEXP, SI000_ALTPAS, SI000_GERAGR, SI000_CODASS)"
                    cSql += " values (" & nProxCod_Usuario.ToString & ",'" & txtLogin.Text & "', '" & ClassCrypt.Encrypt(txtSenha.Text) & "', '" & _
                        txtNmUsuario.Text & "', " & getCodGrupo(cbGrupoPrincipal.Text) & ",'" & FormatarData(dtpExpira.Value) & _
                        "'," & IIf(chkAlterarSenha.Checked, "1", "0") & _
                        "," & IIf(chkGerAgr.Checked, "1", "0") & "," & IIf(txtColaborador.Text = "", "0", Microsoft.VisualBasic.Left(txtColaborador.Text, 6)) & ")"
                ElseIf bAlterar Then
                    cSql = "UPDATE ESI000 set SI000_NOMUSU='" & Trim(txtNmUsuario.Text) & _
                            "', SI000_CODGRU=" & getCodGrupo(cbGrupoPrincipal.Text) & _
                            ", SI000_DATEXP='" & FormatarData(dtpExpira.Value) & _
                            "', SI000_ALTPAS=" & IIf(chkAlterarSenha.Checked, "1", "0") & _
                            ", SI000_GERAGR=" & IIf(chkGerAgr.Checked, "1", "0") & _
                            IIf(txtSenha.Text = txtSenha2.Text, ",SI000_PASLGI='" & _
                                    ClassCrypt.Encrypt(Trim(txtSenha.Text)) & "'", "") & _
                            ", SI000_CODASS=" & IIf(txtColaborador.Text = "", "0", Microsoft.VisualBasic.Left(txtColaborador.Text, 6)) & _
                            " where SI000_LGIUSU = '" & Trim(txtLogin.Text) & "'"
                    'acessoWEB=" & If(chkSIM.Checked = 0, False, True)
                End If
                cmd = New OleDbCommand(cSql, g_ConnectBanco)
                'MsgBox(cSql)
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.ToString())
                Finally
                    bIncluir = False
                    bAlterar = False

                    If g_Param(1) = "INSERT" Then
                        dt.Clear()
                        'fechar o form de cadastro
                        Me.Close()
                    Else
                        dt.Reset()
                        Using da As New OleDbDataAdapter()
                            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

                            ' Preencher o DataTable 
                            da.Fill(dt)
                        End Using

                        'Verificar se o comando veio do browse
                        If g_Comando = "incluir" Or g_Comando = "alterar" Then
                            dt.Clear()
                            Me.Close()
                        Else
                            TratarObjetos()
                        End If

                    End If
                End Try
            Else
                MsgBox(cMensagem)
            End If
        End If

    End Sub

    Private Function Validardados(ByRef cMensagem As String) As Boolean
        Dim bRetorno As Boolean = True

        '?? Acrescentar as validações da Tela ??
        If Trim(txtLogin.Text) = "" Then
            cMensagem += " <O Login não pode estar vazio> "
            bRetorno = False
        End If

        If Trim(txtNmUsuario.Text) = "" Then
            cMensagem += " <O Nome do usuário não pode estar vazio> "
            bRetorno = False
        End If

        If Not txtSenha.Text = txtSenha2.Text And Trim(txtSenha2.Text) <> "" Then
            cMensagem += " <As Senhas não conferem> "
            bRetorno = False
        End If

        Return (bRetorno)

    End Function

    Private Sub btnExcluir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExcluir.Click

        Call Excluir_Registro()

    End Sub

    Private Sub Excluir_Registro()
        Dim cSql As String
        Dim cMensagem As String = ""
        Dim cmd As OleDbCommand

        Application.DoEvents()

        If MsgBox("Deseja excluir este registro?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "cadastro de Usuarios") = MsgBoxResult.Yes Then
            '?? Alterar para a Tabela a ser Excluída ??
            cSql = "DELETE FROM ESI000 where SI000_LGIUSU = '" & txtLogin.Text & "'"
            cmd = New OleDbCommand(cSql, g_ConnectBanco)

            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            Finally

                cSql = "DELETE FROM ESI006 where SI006_CODUSU = " & getCodUsuario(txtLogin.Text)
                cmd = New OleDbCommand(cSql, g_ConnectBanco)
                cmd.ExecuteNonQuery()

                dt.Reset()
                Using da As New OleDbDataAdapter()
                    da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

                    'Preencher o DataTable 
                    da.Fill(dt)
                End Using

                If i > dt.Rows.Count() - 1 Then
                    i = dt.Rows.Count() - 1
                End If

                'Verificar se o comando veio do browse
                If g_Comando = "excluir" Then
                    dt.Clear() 'Limpar o DataTable
                    Me.Close()
                Else
                    TratarObjetos()
                End If

            End Try

        Else
            If Not Trim(cMensagem) = "" Then MsgBox(cMensagem)
        End If
    End Sub

    Private Sub lstGrupoAssoc_DoubleClick(sender As Object, e As EventArgs) Handles lstGrupoAssoc.DoubleClick
        Dim cQuery_grupo As String
        Dim bDelete As Boolean
        Dim cmd As OleDbCommand

        'Gravar / Excluir o Grupo

        'Verificar se o Grupo Já está associado ao usuário
        'Carregar o Combo de Grupos
        cQuery_grupo = "SELECT * FROM ESI006 " & _
        "WHERE ESI006.SI006_CODGRU=" & getCodGrupo(lstGrupoAssoc.SelectedItem) & _
        " AND ESI006.SI006_CODUSU= " & getCodUsuario(dt.Rows(i).Item("SI000_LGIUSU"))

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery_grupo, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtGrupo)
            If dtGrupo.Rows.Count() > 0 Then
                bDelete = True
            Else
                bDelete = False
            End If
        End Using
        dtGrupo.Clear()

        If bDelete Then
            cQuery_grupo = "DELETE FROM ESI006 WHERE ESI006.SI006_CODGRU=" & getCodGrupo(lstGrupoAssoc.SelectedItem) & _
                " AND ESI006.SI006_CODUSU= " & getCodUsuario(dt.Rows(i).Item("SI000_LGIUSU"))
        Else
            cQuery_grupo = "INSERT INTO ESI006 (SI006_CODUSU, SI006_CODGRU)" & _
                " values (" & getCodUsuario(dt.Rows(i).Item("SI000_LGIUSU")) & ", " & getCodGrupo(lstGrupoAssoc.SelectedItem) & ")"
        End If
        cmd = New OleDbCommand(cQuery_grupo, g_ConnectBanco)

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString())
        Finally
            CarregarListBox_Grupos()
        End Try

    End Sub

    Private Sub btnImprimir_Click(sender As Object, e As EventArgs) Handles btnImprimir.Click

        'cria um novo documento para impressão
        Dim pd As PrintDocument = New PrintDocument()
        Dim PrinterSetup As PrintDialog = New PrintDialog

        'Mostrar o PrinterSetup
        PrinterSetup.Document = pd
        PrinterSetup.ShowDialog()
        '**

        'relaciona o objeto pd ao procedimento rptProdutos
        AddHandler pd.PrintPage, AddressOf Me.rptFormulario

        'cria uma nova instância do objeto PrintPreviewDialog()
        Dim objPrintPreview = New PrintPreviewDialog()

        'define algumas propriedades do obejto
        With objPrintPreview
            'indica qual o documento vai ser visualizado
            .Document = pd
            .WindowState = FormWindowState.Maximized
            .PrintPreviewControl.Zoom = 1   'maxima a visualização
            .Text = Me.Text
            'exibe a janela de visualização para o usuário
            .ShowDialog()
        End With

    End Sub

    Private Sub rptFormulario(ByVal sender As Object, ByVal Relatorio As System.Drawing.Printing.PrintPageEventArgs)
        Dim FormControl As Control
        Dim FormListBox As ListBox
        Dim pLinhaList As String

        Dim LinhasPorPagina As Integer
        Dim LinhaAdic As Integer
        Dim posicaoDaLinha As Integer
        Dim y As Integer

        Dim margemEsq As Single = Relatorio.MarginBounds.Left
        Dim margemSup As Single = Relatorio.MarginBounds.Top
        Dim margemDir As Single = Relatorio.MarginBounds.Right
        Dim margemInf As Single = Relatorio.MarginBounds.Bottom

        Dim fonteTitulo As Font
        Dim fonteRodape As Font
        Dim fonteNormal As Font

        fonteTitulo = New Font("Verdana", 15, FontStyle.Bold)
        fonteRodape = New Font("Verdana", 8)
        fonteNormal = New Font("Verdana", 10)

        LinhasPorPagina = Relatorio.MarginBounds.Height / fonteNormal.GetHeight(Relatorio.Graphics) - 10

        margemEsq = 10
        'Imprimir o Cabeçalho
        Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, 40, margemDir, 40)
        Relatorio.Graphics.DrawImage(Image.FromFile("logo.png"), 10, 48)
        Relatorio.Graphics.DrawString(Me.Text, fonteTitulo, Brushes.Blue, margemEsq + 275, 80, New StringFormat())
        Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, 145, margemDir, 145)
        LinhaAdic = 2

        For Each FormControl In Me.Controls
            If (TypeOf FormControl Is Label) Then
                posicaoDaLinha = margemSup + (((FormControl.TabIndex * 2) + LinhaAdic) * fonteNormal.GetHeight(Relatorio.Graphics))
                Relatorio.Graphics.DrawString(FormControl.Text, fonteNormal, Brushes.Black, margemEsq, posicaoDaLinha, New StringFormat())
            ElseIf (TypeOf FormControl Is TextBox) Or (TypeOf FormControl Is ComboBox) Then
                posicaoDaLinha = margemSup + (((FormControl.TabIndex * 2) + LinhaAdic) * fonteNormal.GetHeight(Relatorio.Graphics))
                Relatorio.Graphics.DrawString(FormControl.Text, fonteNormal, Brushes.Black, margemEsq + 150, posicaoDaLinha, New StringFormat())
            ElseIf (TypeOf FormControl Is DateTimePicker) Then

                posicaoDaLinha = margemSup + (((FormControl.TabIndex * 2) + LinhaAdic) * fonteNormal.GetHeight(Relatorio.Graphics))
                Relatorio.Graphics.DrawString(FormControl.Text, fonteNormal, Brushes.Black, margemEsq + 150, posicaoDaLinha, New StringFormat())
            ElseIf (TypeOf FormControl Is ListBox) Then
                FormListBox = FormControl
                posicaoDaLinha = margemSup + (((FormListBox.TabIndex * 2) + LinhaAdic) * fonteNormal.GetHeight(Relatorio.Graphics))
                pLinhaList = ""
                For y = 0 To FormListBox.Items.Count - 1
                    If Trim(FormListBox.Items(y).ToString) = "" Then Exit For
                    pLinhaList += "(" & FormListBox.Items(y).ToString & ") "
                Next
                Relatorio.Graphics.DrawString(pLinhaList, fonteNormal, Brushes.Black, margemEsq + 150, posicaoDaLinha, New StringFormat())
            End If
        Next

        'imprime o rodape no relatorio
        Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, margemInf, margemDir, margemInf)
        Relatorio.Graphics.DrawString(System.DateTime.Now, fonteRodape, Brushes.Black, margemEsq, margemInf, New StringFormat())
        Relatorio.Graphics.DrawString("Formulário", fonteRodape, Brushes.Black, margemDir - 50, margemInf, New StringFormat())

        Relatorio.HasMorePages = False

    End Sub

    Private Sub btnLocalizar_Click(sender As Object, e As EventArgs) Handles btnLocalizar.Click
        dt.Clear() 'Limpar o DataTable

        Me.Close()
    End Sub

    Private Sub chkValidade_CheckedChanged(sender As Object, e As EventArgs) Handles chkValidade.CheckedChanged

        If chkValidade.Checked Then
            dtpExpira.Value = CDate("01/01/1900")
        End If
        dtpExpira.Enabled = Not chkValidade.Checked And bAlterar
        lblDtpExpira.Enabled = Not chkValidade.Checked And bAlterar

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick

        'Usado o Timer para poder carregar o formulário antes de excluir
        'Verificar se é para excluir o registro comandado pelo browse
        Timer1.Enabled = False
        If g_Comando = "excluir" Then
            Call Excluir_Registro()
        End If

    End Sub

    Private Sub cbGrupoPrincipal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbGrupoPrincipal.SelectedIndexChanged

    End Sub

    Private Sub cbGrupoPrincipal_SelectedValueChanged(sender As Object, e As EventArgs) Handles cbGrupoPrincipal.SelectedValueChanged
        Call CarregarListBox_Grupos()
    End Sub

    Private Sub btnLocColaborador_Click(sender As Object, e As EventArgs) Handles btnLocColaborador.Click
        'txtColaborador.Text = dlgColaborador.ShowDialog()
        Dim options = New dlgColaborador

        ' Did the user click Save?
        If options.ShowDialog() = Windows.Forms.DialogResult.OK Then
            ' Yes, so grab the values you want from the dialog here
            'Dim textBoxValue As String = options.txtPesquisa.Text
            txtColaborador.Text = Microsoft.VisualBasic.Left(options.txtPesquisa.Text, 6) & " - " & LerNome_Colaborador(Microsoft.VisualBasic.Left(options.txtPesquisa.Text, 6))
        End If
    End Sub
End Class