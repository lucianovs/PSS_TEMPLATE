Imports System.Data.OleDb
Imports System.Drawing.Printing

Public Class frmBrowse
    Dim Formulario As String
    Dim classEntidade As String
    Dim sFiltroInicial As String
    Dim sModulo As String
    Dim sWhere As String
    Dim sJoin As String
    Dim cSql As String = ""
    Dim nCod_Login As Integer
    Private paginaAtual As Integer = 1
    Private nBlocoReg As Double = 0

    Public Sub New(ByVal fTabela As String, Cadastro_Form As String, Optional fJoin As String = "", Optional fWhere As String = "")
        InitializeComponent()

        Formulario = Cadastro_Form
        classEntidade = fTabela
        sJoin = fJoin
        sWhere = fWhere
        sModulo = g_Modulo

        'sFiltroInicial = classLoadListView.loadlistview(fTabela, ListView_Browse)
        'A variavel cEntidade vai retornar o filtro do Banco, se houver
        'sFiltroInicial = fTabela

    End Sub

    Private Sub frmBrowse_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim i As Integer
        Dim nCont1 As Integer
        Dim nCont2 As Integer
        Dim sTextCB As String
        Dim sTempCampo As String
        Dim sTempCondicao As String
        Dim sTempValorCondicao As String
        Dim cmdInsert As OleDbCommand

        Dim dtSI902 As DataTable = New DataTable("ESI902")
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        nCod_Login = getCodUsuario(ClassCrypt.Decrypt(g_Login)) 'CARREGAR O USUÁRIO LOGADO
        If nCod_Login = 1 Then
            sJoin = ""
            sWhere = ""
        End If
        'MsgBox(sJoin & " - " & sWhere)
        sFiltroInicial = classLoadListView.loadlistview(classEntidade, ListView_Browse, sJoin, sWhere, nBlocoReg)
        Application.DoEvents()
        PosicionarListView(ListView_Browse)
        Call TratarPaginacao()

        If Not ConectarBanco() Then MsgBox("erro ao conectar com o bd")

        'Carregar o Combo de Campos
        cbCampo.Items.Clear()
        For i = 0 To Me.ListView_Browse.Columns.Count - 1
            If Me.ListView_Browse.Columns(i).Text.StartsWith("{") Then
                sTextCB = Replace(Me.ListView_Browse.Columns(i).Text, "{", "")
                sTextCB = Replace(sTextCB, "}", "")
            Else
                sTextCB = Me.ListView_Browse.Columns(i).Text
            End If
            cbCampo.Items.Add(sTextCB)
        Next
        cbCampo.Items.Add("")

        'Montar o filtro gravado na tabela
        If Trim(sFiltroInicial) <> "" Then
            'Separar as informações do filtro
            nCont1 = InStr(8, sFiltroInicial, " ", vbTextCompare)
            nCont2 = InStr(nCont1 + 1, sFiltroInicial, " ", vbTextCompare)
            sTempCampo = Trim(Mid(sFiltroInicial, 8, nCont1 - 7))
            sTempCondicao = Trim(Mid(sFiltroInicial, nCont1 + 1, nCont2 - nCont1 - 1))
            sTempValorCondicao = TRATACONDICAO(sFiltroInicial)
            'sTempValorCondicao = Trim(Mid(sFiltroInicial, nCont2 + 1, Len(sFiltroInicial) - nCont2))

            'Buscar a descrição do Campo para o Combo
            cSql = "SELECT * FROM ESI902 WHERE SI902_NOMCPO='" & sTempCampo & _
                            "' AND SI902_NOMENT = '" & classEntidade & "'"
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cSql, g_ConnectBanco)

                ' Preencher o DataTable 
                da.Fill(dtSI902)
            End Using
            If dtSI902.Rows.Count() > 0 Then
                cbCampo.Text = dtSI902.Rows(0).Item("SI902_DESCPO").ToString
            Else
                cbCampo.Text = ""
            End If
            txtValorCondicao.Text = Replace(sTempValorCondicao, "%", "")
            txtValorCondicao.Text = Replace(txtValorCondicao.Text, "*", "")
            txtValorCondicao.Text = Replace(txtValorCondicao.Text, "'", "")

            'Montar os conectores
            If sTempCondicao = "=" Then
                cbCondicao.Text = "igual a"
            ElseIf sTempCondicao = ">" Then
                cbCondicao.Text = "Maior que"
            ElseIf sTempCondicao = "<" Then
                cbCondicao.Text = "Menor que"
            ElseIf sTempCondicao = "like" Then
                cbCondicao.Text = "Contenha"
            ElseIf sTempCondicao = "<>" Then
                cbCondicao.Text = "Diferente de"
            Else
                cbCondicao.Text = ""
            End If
        End If

        'Carregar o tamanho do form
        Dim dtSI903 As DataTable = New DataTable("ESI903")
        'Carregar as informações do tamanho do form na tabela 
        cSql = "SELECT * FROM ESI903 WHERE SI903_NOMCPO='BROWSE" & _
                        "' AND SI903_NOMENT = '" & classEntidade & _
                        "' AND SI903_CODUSU = " & nCod_Login.ToString
        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cSql, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtSI903)
        End Using
        If dtSI903.Rows.Count() > 0 Then
            Me.Size = New Size(dtSI903.Rows(0).Item("SI903_TAMBRW"), dtSI903.Rows(0).Item("SI903_POSBRW"))
        Else
            cSql = "INSERT INTO ESI903 (SI903_CODUSU, SI903_NOMENT, SI903_NOMCPO, SI903_TAMBRW, SI903_POSBRW, " & _
                        "SI903_SQLBRW, SI903_SQLORD) VALUES (" & nCod_Login.ToString & ",'" & Trim(classEntidade) & "','BROWSE'," & _
                        "878,384,'Tamanho do form (x,y)','')"
            cmdInsert = New OleDbCommand(cSql, g_ConnectBanco)
            Try
                cmdInsert.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
        End If

        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default

        ListView_Browse.Size = New Size(Me.Size.Width - 34, Me.Size.Height - 146) '112

    End Sub

    Private Function TrataCondicao(fFiltro As String) As String
        Dim nX1, nX2 As Integer
        Dim nStart As Integer = 8

        TrataCondicao = ""
        Do Until nStart = 0
            '12345678901234567890123456789012345678901234567890123456789012
            ' WHERE UN000_NOMUNI like '%C.M.%'         and UN000_NOMUNI like '%BA%'"
            'nx1=20 34 51 62 0 nx2=25 56   nstart=8 39 0  Condicao='%C.M.%' '%BA%'
            'nx1=20 0 33 nx2=25    nstart=8 0  Condicao='%C.M.%' '%BA%'
            '****************
            nX1 = InStr(nStart, fFiltro, " ", vbTextCompare) 'Localizar o primeiro espaço apos o "Campo"
            nX2 = InStr(nX1 + 1, fFiltro, " ", vbTextCompare) 'Loalizar o espaço após o operador

            nX1 = InStr(nStart, UCase(fFiltro), " AND ", vbTextCompare) 'Verificar se tem mais expressões
            If nX1 = 0 Then
                nX1 = Len(fFiltro)
                nStart = 0
            Else
                nStart = nX1 + 5
            End If
            TrataCondicao += LTrim(Mid(fFiltro, nX2 + 1, nX1 - nX2))

        Loop
        TrataCondicao = RTrim(TrataCondicao)

    End Function

    Private Sub ListView_Browse_DoubleClick(sender As Object, e As EventArgs) Handles ListView_Browse.DoubleClick
        Dim i As Integer
        Dim iParam As Integer = 1

        'Inicializar os Parâmetros da Chave (até 3 chaves)
        g_Param(1) = ""
        g_Param(2) = ""
        g_Param(3) = ""
        g_Comando = "ver"

        'Verificar os campos com { e carregar os Parametros
        For i = 0 To Me.ListView_Browse.Columns.Count - 1
            If Me.ListView_Browse.Columns(i).Text.StartsWith("{") Then
                g_Param(iParam) = ListView_Browse.FocusedItem.SubItems(i).Text
                iParam += 1
            End If
        Next

        Call CarregarFormCadastro()

    End Sub

    Private Sub CarregarFormCadastro()
        'Carregar o Formulário do Cadastro
        Dim frmType As Type = Type.GetType(sModulo & "." & Formulario)
        Dim frm As Object = Activator.CreateInstance(frmType)

        CType(frm, Form).Tag = Me.Tag 'nível de acesso
        CType(frm, Form).Text = "FORMULÁRIO - " & Me.Text
        frmType.InvokeMember("Show", Reflection.BindingFlags.InvokeMethod, Nothing, frm, Nothing)

        timerRefresh.Enabled = True

        'CType(frm, Form).MdiParent = mdiDesktop
        'CType(frm, Form).Modal = True

        'Cadastro_Form.MdiParent = mdiPrincipal
        'Cadastro_Form.Tag = Me.Tag 'nível de acesso
        'Cadastro_Form.Show()
    End Sub

    Private Sub btnSalvar_Click(sender As Object, e As EventArgs) Handles btnSalvar.Click, btnClear.Click
        Dim i As Integer
        Dim cSql As String
        Dim cCondicao As String = ""
        Dim cCondicaoTexto As String = ""
        Dim sTitCampo As String
        Dim cmd As OleDbCommand
        Dim AtualizarBtn As Button
        Dim sNomeCampo As String

        Dim nPos As Integer = 99
        Dim nSeq As Integer = 0
        Dim nStart As Integer = 1
        Dim sNome(10) As String

        Dim dtSI902 As DataTable = New DataTable("ESI902")

        System.Windows.Forms.Cursor.Current = Cursors.WaitCursor

        If IsNothing(sender) Then
            AtualizarBtn = Me.btnSalvar
        Else
            AtualizarBtn = sender
        End If

        'Limpar o Filtro
        If AtualizarBtn.Name = "btnClear" Then
            cbCampo.Text = ""
            cbCondicao.Text = ""
            txtValorCondicao.Text = ""
        End If

        If Not cbCampo.Text = "" Then
            cSql = "SELECT * FROM ESI902 WHERE SI902_DESCPO='" & cbCampo.Text & _
                "' AND SI902_NOMENT = '" & classEntidade & "'"
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cSql, g_ConnectBanco)

                ' Preencher o DataTable 
                da.Fill(dtSI902)
            End Using
            If dtSI902.Rows.Count() > 0 Then
                'CCondicao = " where " & dtSI902.Rows(0).Item("SI902_NOMCPO").ToString
                cCondicao = dtSI902.Rows(0).Item("SI902_NOMCPO").ToString
                If Trim(cbCondicao.Text) = "Igual a" Then
                    CCondicao += " = "
                ElseIf Trim(cbCondicao.Text) = "Maior que" Then
                    CCondicao += " > "
                ElseIf Trim(cbCondicao.Text) = "Menor que" Then
                    CCondicao += " < "
                ElseIf Trim(cbCondicao.Text) = "Diferente de" Then
                    CCondicao = CCondicao & " <> "
                ElseIf Trim(cbCondicao.Text) = "Contenha" Then
                    CCondicao += " like "
                Else
                    CCondicao = ""
                End If

                If CCondicao <> "" And Trim(txtValorCondicao.Text) <> "" Then
                    If dtSI902.Rows(0).Item("SI902_TIPCPO") = "T" Then 'Texto
                        CCondicaoTexto = txtValorCondicao.Text
                        If Trim(cbCondicao.Text) = "Contenha" Then
                            'Montar a condição com mais de uma palavra
                            nStart = 1
                            nSeq = 0
                            Do Until nPos = 0
                                nPos = InStr(nStart, txtValorCondicao.Text, " ")
                                If nPos > 0 Then
                                    sNome(nSeq) = Mid(txtValorCondicao.Text, nStart, nPos - nStart)
                                    nStart = nPos + 1
                                    nSeq += 1
                                Else
                                    sNome(nSeq) = Mid(txtValorCondicao.Text, nStart, Len(txtValorCondicao.Text) - nStart + 1)
                                End If
                            Loop
                            'Montar a Condição LIKE
                            'Limpar a Condição
                            CCondicaoTexto = ""
                            For nPos = 0 To nSeq
                                If nPos > 0 Then
                                    CCondicaoTexto += " and "
                                End If
                                CCondicaoTexto += CCondicao & "'%" & sNome(nPos) & "%'"
                            Next nPos
                            CCondicao = ""
                            'CCondicaoTexto = "'%" & CCondicaoTexto & "%" 
                        Else
                            CCondicaoTexto = "'" & CCondicaoTexto & "'"
                        End If

                    ElseIf dtSI902.Rows(0).Item("SI902_TIPCPO") = "N" And cbCondicao.Text <> "Contenha " Then 'Texto
                        If IsNumeric(txtValorCondicao.Text) Then
                            CCondicaoTexto = txtValorCondicao.Text
                        End If
                    ElseIf dtSI902.Rows(0).Item("SI902_TIPCPO") = "D" And cbCondicao.Text = "Contenha " Then 'Data
                        If IsDate(txtValorCondicao.Text) Then
                            CCondicaoTexto = Format(txtValorCondicao.Text, "dd/mm/yyyy")
                        End If
                    ElseIf dtSI902.Rows(0).Item("SI902_TIPCPO") = "B" And cbCondicao.Text = "Contenha " Then 'Verdadeiro/Falso
                        If txtValorCondicao.Text.StartsWith("V") Or txtValorCondicao.Text.StartsWith("T") Then
                            CCondicaoTexto = "True"
                        Else
                            CCondicaoTexto = "False"
                        End If
                    End If
                    CCondicao = " WHERE " & CCondicao & CCondicaoTexto
                    CCondicao = CCondicao.Replace("'", "\/")
                End If
            End If
        End If

        If ConectarBanco() Then
            For i = 0 To Me.ListView_Browse.Columns.Count - 1
                'Retirar as {} do cabecalho para comparar com a tabela
                sTitCampo = Replace(Me.ListView_Browse.Columns(i).Text, "{", "")
                sTitCampo = Replace(sTitCampo, "}", "")

                sNomeCampo = LerNomeCampo(Trim(classEntidade), sTitCampo)

                cSql = "UPDATE ESI903 set SI903_TAMBRW=" & Me.ListView_Browse.Columns(i).Width.ToString & _
                            ", SI903_SQLBRW='" & IIf(sTitCampo = cbCampo.Text, cCondicao, "") & "'"
                'Verificar a ordenação
                cSql += ", SI903_SQLORD='"
                If sTitCampo = cbCampo.Text Then 'Ordenacao pelo campo do Filtro
                    cSql += "ASC'"
                ElseIf cbCampo.Text = "" Then 'Não tem filtro
                    'Verificar se é o campo Chave
                    If InStr(Me.ListView_Browse.Columns(i).Text, "{", CompareMethod.Text) > 0 Then
                        cSql += "ASC'"
                    Else
                        cSql += "'"
                    End If
                Else
                    cSql += "'"
                End If
                cSql += " where SI903_NOMENT='" & Trim(classEntidade) & "' and SI903_NOMCPO='" & sNomeCampo & _
                        "' and SI903_CODUSU=" & nCod_Login.ToString
                cmd = New OleDbCommand(cSql, g_ConnectBanco)

                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    MsgBox(ex.ToString())
                End Try
            Next

            'Gravar o tamanho do Formulário
            cSql = "UPDATE ESI903 set SI903_TAMBRW=" & Me.Size.Width.ToString & _
                            ", SI903_POSBRW=" & Me.Size.Height.ToString & _
                            " where SI903_NOMENT='" & Trim(classEntidade) & "' and SI903_NOMCPO='BROWSE'" & _
                            " AND SI903_CODUSU=" & nCod_Login.ToString
            cmd = New OleDbCommand(cSql, g_ConnectBanco)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
            '***********************************
            nBlocoReg = 0
            'sFiltroInicial = classLoadListView.loadlistview(classEntidade, ListView_Browse, , , nBlocoReg)
            sFiltroInicial = classLoadListView.loadlistview(classEntidade, ListView_Browse, sJoin, sWhere, nBlocoReg)
            PosicionarListView(ListView_Browse)
            Call TratarPaginacao()

            'If CCondicao <> "" Then
            'classLoadListView.loadlistview(classEntidade, ListView_Browse)
            'End If
            ListView_Browse.Size = New Size(Me.Size.Width - 34, Me.Size.Height - 146) '112
        Else
            MsgBox("Erro ao acessar o banco de dados !!!")
            Me.Close()
        End If

        System.Windows.Forms.Cursor.Current = Cursors.Default

    End Sub

    Private Sub ListView_Browse_MouseDown(sender As Object, e As MouseEventArgs) Handles ListView_Browse.MouseDown
        If ListView_Browse.Items.Count = 0 Then
            If MsgBox("Nenhum registro existe para ser exibido, deseja incluir um novo?", MsgBoxStyle.Question + MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                g_Param(1) = "INSERT" 'Abrir o form do Cadastro no Modo Inserir
                Call CarregarFormCadastro()
            End If
        End If
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
        AddHandler pd.PrintPage, AddressOf Me.rptBrowse

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

    Private Sub rptBrowse(ByVal sender As Object, ByVal Relatorio As System.Drawing.Printing.PrintPageEventArgs)
        Dim LinhasPorPagina As Integer
        Dim LinhaAtual As Integer
        Dim posicaoDaLinha As Integer
        Dim x, y As Integer
        Dim EspacoGap As Integer

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
        LinhaAtual = 9999 'Forçar a primeira quebra 

        Dim Column_Cabec(ListView_Browse.Columns.Count) As String
        Dim Column_Detail(ListView_Browse.Columns.Count) As String

        'Montar os Titulos das Colunas
        For x = 0 To Me.ListView_Browse.Columns.Count - 1
            If Me.ListView_Browse.Columns(x).Text.StartsWith("{") Then
                Column_Cabec(x) = Replace(Me.ListView_Browse.Columns(x).Text, "{", "")
                Column_Cabec(x) = Replace(Column_Cabec(x), "}", "")
            Else
                Column_Cabec(x) = Me.ListView_Browse.Columns(x).Text
            End If
        Next

        'Carregar os dados da Lista
        For y = 0 To Me.ListView_Browse.Items.Count - 1

            'Verificar a Quebra
            If (LinhaAtual > LinhasPorPagina) Then
                'Imprimir o Rodapé
                If LinhaAtual <> 9999 Then
                    'imprime o rodape no relatorio
                    Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, margemInf, margemDir, margemInf)
                    Relatorio.Graphics.DrawString(System.DateTime.Now, fonteRodape, Brushes.Black, margemEsq, margemInf, New StringFormat())
                    Relatorio.Graphics.DrawString("Pag. " & paginaAtual.ToString, fonteRodape, Brushes.Black, margemDir - 50, margemInf, New StringFormat())
                    'incrementa a página atual
                    paginaAtual += 1
                    Relatorio.HasMorePages = True
                End If

                margemEsq = 10
                'Imprimir o Cabeçalho
                Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, 40, margemDir, 40)
                Relatorio.Graphics.DrawImage(Image.FromFile("logo.png"), 10, 48)
                Relatorio.Graphics.DrawString(Me.Text, fonteTitulo, Brushes.Blue, margemEsq + 275, 80, New StringFormat())
                Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, 145, margemDir, 145)

                'impressão do titulo das colunas
                For x = 0 To Me.ListView_Browse.Columns.Count - 1
                    If x > 0 Then
                        margemEsq += Me.ListView_Browse.Columns(x - 1).Width
                    End If
                    Relatorio.Graphics.DrawString(Column_Cabec(x), fonteNormal, Brushes.Red, margemEsq, 160, New StringFormat())
                Next
                Relatorio.Graphics.DrawLine(Pens.Black, 5, 190, margemDir, 190)
                LinhaAtual = 6
            End If

            posicaoDaLinha = margemSup + (LinhaAtual * fonteNormal.GetHeight(Relatorio.Graphics))

            'Carregar as colunas para o Detalhe 
            margemEsq = 10
            For cont = 0 To x - 1
                If (IsNumeric(Me.ListView_Browse.Items(y).Text) And cont = 0) Or _
                        (IsNumeric(Me.ListView_Browse.Items(y).SubItems(cont).Text) And cont > 0) Then
                    EspacoGap = CInt((Me.ListView_Browse.Columns(cont).Width - Len(Trim(Me.ListView_Browse.Items(y).Text))) / 2)
                Else
                    EspacoGap = 0
                End If
                If cont > 0 Then
                    margemEsq += Me.ListView_Browse.Columns(cont - 1).Width
                    Relatorio.Graphics.DrawString(Me.ListView_Browse.Items(y).SubItems(cont).Text, fonteNormal, Brushes.Black, margemEsq + EspacoGap, posicaoDaLinha, New StringFormat())
                Else
                    Relatorio.Graphics.DrawString(Me.ListView_Browse.Items(y).Text, fonteNormal, Brushes.Black, margemEsq + EspacoGap, posicaoDaLinha, New StringFormat())
                End If
            Next
            LinhaAtual += 1
        Next

        'imprime o rodape no relatorio
        Relatorio.Graphics.DrawLine(Pens.Black, margemEsq, margemInf, margemDir, margemInf)
        Relatorio.Graphics.DrawString(System.DateTime.Now, fonteRodape, Brushes.Black, margemEsq, margemInf, New StringFormat())
        Relatorio.Graphics.DrawString("Pag. " & paginaAtual.ToString, fonteRodape, Brushes.Black, margemDir - 50, margemInf, New StringFormat())

        Relatorio.HasMorePages = False

    End Sub

    Private Sub PosicionarListView(ByRef Lista As ListView)
        Dim bAchou As Boolean

        'Localizar até 3 parâmetros, qdo não vazio
        For Each LVI As ListViewItem In Lista.Items
            If LVI.SubItems(0).Text = g_Param(1) Then
                If g_Param(2) <> "" Then
                    If LVI.SubItems(1).Text = g_Param(2) Then
                        If g_Param(3) = "" Then
                            If LVI.SubItems(2).Text = g_Param(3) Then
                                bAchou = True
                            End If
                        Else
                            bAchou = True
                        End If
                    End If
                Else
                    bAchou = True
                End If
            End If

            If bAchou Then
                LVI.Selected = True
                LVI.EnsureVisible()

                Exit Sub
            End If
        Next

    End Sub

    Private Sub frmBrowse_SizeChanged(sender As Object, e As EventArgs) Handles Me.SizeChanged

        ListView_Browse.Size = New Size(Me.Size.Width - 34, Me.Size.Height - 146) '112

        btnAnterior.Location = New System.Drawing.Point(ListView_Browse.Size.Width - 245, ListView_Browse.Size.Height + 58)
        lblRegistros.Location = New System.Drawing.Point(btnAnterior.Location.X + 62, btnAnterior.Location.Y + 3)
        btnProximo.Location = New System.Drawing.Point(btnAnterior.Location.X + 188, btnAnterior.Location.Y)

    End Sub

    Private Sub cbCampo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbCampo.SelectedIndexChanged
        If Trim(cbCampo.Text) = "" Then
            cbCondicao.Text = ""
            txtValorCondicao.Text = ""
        End If
    End Sub

    Private Sub btnIncluir_Click(sender As Object, e As EventArgs) Handles btnIncluir.Click, btnAlterar.Click, btnExcluir.Click
        Dim fBotao As Button
        Dim i As Integer
        Dim iParam As Integer

        fBotao = sender

        'Inicializar os Parâmetros da Chave (até 3 chaves)
        g_Param(0) = ""
        g_Param(1) = ""
        g_Param(2) = ""
        g_Param(3) = ""
        g_Param(4) = ""
        g_Param(5) = ""
        
        If fBotao.Name = "btnIncluir" Then
            g_Comando = "incluir"
        ElseIf fBotao.Name = "btnExcluir" Then
            g_Comando = "excluir"
        ElseIf fBotao.Name = "btnAlterar" Then
            g_Comando = "alterar"
        Else
            g_Comando = "ver"
        End If

        If g_Comando <> "incluir" Then
            On Error GoTo erro_comandos

            'Verificar os campos com { e carregar os Parametros
            iParam = 1
            For i = 0 To Me.ListView_Browse.Columns.Count - 1
                If Me.ListView_Browse.Columns(i).Text.StartsWith("{") Then
                    g_Param(iparam) = ListView_Browse.FocusedItem.SubItems(i).Text
                    iParam += 1
                End If
            Next
        Else
            g_Param(1) = "" 'Para não localizar nenhum registro
        End If

        Call CarregarFormCadastro()

        'If g_Comando = "excluir" Then
        'For Each x As ListViewItem In ListView_Browse.SelectedItems
        'ListView_Browse.Items.Remove(x)
        'Next
        'End If

        Exit Sub

erro_comandos:
        'Qdo houver erro, não executar nenhuma ação
    End Sub

    Private Sub frmBrowse_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim toolTipText As New ToolTip()

        'Colocar os Helps nos Botões
        toolTipText.AutoPopDelay = 5000
        toolTipText.InitialDelay = 1000
        toolTipText.AutomaticDelay = 300
        toolTipText.ShowAlways = True

        toolTipText.SetToolTip(Me.btnAlterar, "Alterar")
        toolTipText.SetToolTip(Me.btnIncluir, "Incluir")
        toolTipText.SetToolTip(Me.btnExcluir, "Excluir")
        toolTipText.SetToolTip(Me.btnSalvar, "Aplicar filtro / Gravar configuração da tela")
        toolTipText.SetToolTip(Me.btnImprimir, "Imprimir a lista")
        toolTipText.SetToolTip(Me.btnClear, "Limpar o filtro")
        toolTipText.SetToolTip(Me.btnAnterior, "Página Anterior")
        toolTipText.SetToolTip(Me.btnProximo, "Próxima Página")

    End Sub

    Private Sub TratarPaginacao()
        lblRegistros.Text = Str(nBlocoReg + 1) & " - " & CStr(ListView_Browse.Items.Count() + nBlocoReg)

        btnAnterior.Enabled = nBlocoReg > 0
        btnProximo.Enabled = (ListView_Browse.Items.Count() >= LerDadosINI(nomeArquivoINI(), "BROWSE", "Bloco_Browse", 50))

    End Sub

    Private Sub btnProximo_Click(sender As Object, e As EventArgs) Handles btnProximo.Click

        Me.Cursor = Cursors.WaitCursor
        nBlocoReg += LerDadosINI(nomeArquivoINI(), "BROWSE", "Bloco_Browse", 50)
        sFiltroInicial = classLoadListView.loadlistview(classEntidade, ListView_Browse, sJoin, sWhere, nBlocoReg)

        PosicionarListView(ListView_Browse)
        Call TratarPaginacao()
        Me.Cursor = Cursors.Default

    End Sub

    Private Sub btnAnterior_Click(sender As Object, e As EventArgs) Handles btnAnterior.Click

        Me.Cursor = Cursors.WaitCursor
        If nBlocoReg > 0 Then
            nBlocoReg -= LerDadosINI(nomeArquivoINI(), "BROWSE", "Bloco_Browse", 50)

            sFiltroInicial = classLoadListView.loadlistview(classEntidade, ListView_Browse, sJoin, sWhere, nBlocoReg)
            PosicionarListView(ListView_Browse)
            Call TratarPaginacao()
        End If
        Me.Cursor = Cursors.Default

    End Sub

    Public Sub RefreshBrowse()
        If g_Comando = "REFRESH" Then
            g_Comando = ""
            timerRefresh.Enabled = False
            sFiltroInicial = classLoadListView.loadlistview(classEntidade, ListView_Browse, sJoin, sWhere, nBlocoReg)
            Application.DoEvents()
            PosicionarListView(ListView_Browse)
            Call TratarPaginacao()
        End If

    End Sub

    Private Sub timerRefresh_Tick(sender As Object, e As EventArgs) Handles timerRefresh.Tick
        Call RefreshBrowse()
    End Sub

    Private Sub txtValorCondicao_KeyDown(sender As Object, e As KeyEventArgs) Handles txtValorCondicao.KeyDown

        If e.KeyCode = Keys.Enter Then
            Call btnSalvar_Click(Nothing, New System.EventArgs())
            e.Handled = True
        End If

    End Sub

End Class