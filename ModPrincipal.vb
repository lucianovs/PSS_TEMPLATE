'Módulo Principal com as funções Globais
Imports System.Data.OleDb
Imports System.IO
Imports System.Text

Module ModPrincipal
    'Declaração das variáveis Globais
    Public g_ConnectString As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=SSVP.mdb"
    Public g_ConnectBanco = New OleDbConnection()
    'Variaveis Controle Login
    Public g_Login As String
    Public g_Nivel As Integer

    'Guarda os parâmetros para passar do Browse para aos cadastros
    Public g_Param(6) As String
    'Definição do Projeto (Deve ser incializado no Load do Form Principal e ter o mesmo nome do projeto (Sensitivo)
    Public g_Modulo As String
    'Definição do comando que o browse envio para o formulário de cadastro
    Public g_Comando As String
    'Variável para Identificar a Estrutura para Inserir Unidades
    Public g_Unidade As String
    'Variavel para o DialogBox da Impressora


    'Variáveis utilizados na Pesquisa de registros
    Friend g_Pesq1, g_Pesq2, g_Pesq3 As String

    Private Declare Auto Function GetPrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Auto Function WritePrivateProfileString Lib "Kernel32" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer

    'Usa a função GetPrivateProfileString para obter os valores do Arquivo Ini
    Public Function LerDadosINI(ByVal file_name As String, ByVal section_name As String, ByVal key_name As String, ByVal default_value As String) As String
        Const MAX_LENGTH As Integer = 500

        Dim string_builder As New StringBuilder(MAX_LENGTH)

        GetPrivateProfileString(section_name, key_name, default_value, string_builder, MAX_LENGTH, file_name)

        Return string_builder.ToString()

    End Function

    ' Retorna o nome do arquivo INI 
    Public Function nomeArquivoINI() As String
        Dim nome_arquivo_ini As String = Application.StartupPath

        'Return nome_arquivo_ini & "\" & Application.ProductName & ".ini"
        Return nome_arquivo_ini & "\SSVP.ini"

    End Function

    Public Sub GravarDadosIni(ByVal section_name As String, ByVal key_name As String, cValor As String)
        Dim nome_arquivo_ini As String = nomeArquivoINI()

        WritePrivateProfileString(section_name, key_name, cValor, nome_arquivo_ini)

    End Sub

    Public Function ConectarBanco() As Boolean
        Dim bConnect As Boolean

        'Descriptografar o String de Conexao
        g_ConnectString = ClassCrypt.Decrypt(g_ConnectString)
        '???
        'g_ConnectString = Replace(g_ConnectString, "C:\", "D:\")
        '???

        If Not g_ConnectBanco.state = 1 Then
            Try
                g_ConnectBanco.ConnectionString = g_ConnectString
                g_ConnectBanco.Open()
                bConnect = True
            Catch ex As Exception
                MsgBox("Erro ao conectar ao Banco de Dados do Sistema." & vbCrLf & ex.Message)
                bConnect = False
            Finally

            End Try
        Else
            bConnect = True
        End If
        'Criptografar Novamente
        g_ConnectString = ClassCrypt.Encrypt(g_ConnectString)

        Return bConnect

    End Function

    Public Function LerUsuario(ByVal fLgiUsu As String, Optional fPswUsu As String = "") As String
        Dim cQuery As String
        Dim dtUsuario As DataTable = New DataTable("ESI000")

        cQuery = "SELECT SI000_LGIUSU, SI000_NOMUSU, SI000_PASLGI, SI000_DATEXP, SI000_ALTPAS " & _
            "FROM ESI000 where SI000_LGIUSU = '" & fLgiUsu & "'"

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtUsuario)

            If dtUsuario.Rows.Count() > 0 Then
                If dtUsuario.Rows(0).Item("SI000_DATEXP") < Today() And Format(dtUsuario.Rows(0).Item("SI000_DATEXP"), "yyyy") <> 1900 Then
                    MsgBox("Usuário expirado. Favor contactar o administrador.")
                    LerUsuario = ""
                Else
                    LerUsuario = dtUsuario.Rows(0).Item("SI000_LGIUSU")
                    If Not IsNothing(fPswUsu) Then
                        If fPswUsu <> ClassCrypt.Decrypt(dtUsuario.Rows(0).Item("SI000_PASLGI")) Then
                            LerUsuario = ""
                        Else
                            If dtUsuario.Rows(0).Item("SI000_ALTPAS") = 1 Then
                                g_Param(2) = dtUsuario.Rows(0).Item("SI000_LGIUSU")
                            Else
                                g_Param(2) = ""
                            End If
                        End If
                    End If
                End If
            Else
                LerUsuario = ""
            End If
        End Using
        dtUsuario.Clear()

    End Function

    Public Function Usuario_GerAgregacao(ByVal fLgiUsu As String) As Integer
        'Retorna 0 - Não gerencia
        'Retorna 1 - Gerencia 
        Dim cQuery As String
        Dim dtUsuario As DataTable = New DataTable("ESI000")

        cQuery = "SELECT SI000_LGIUSU, SI000_NOMUSU, SI000_PASLGI, SI000_DATEXP, SI000_ALTPAS " & _
            "FROM ESI000 where SI000_LGIUSU = '" & fLgiUsu & "'"

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtUsuario)

            If dtUsuario.Rows.Count() > 0 Then
                Usuario_GerAgregacao = IIf(IsDBNull(dtUsuario.Rows(0).Item("SI000_GERAGR")), 0, dtUsuario.Rows(0).Item("SI000_GERAGR"))
            Else
                Usuario_GerAgregacao = 0
            End If
        End Using
        dtUsuario.Clear()
    End Function

    Public Function NivelAcesso(fCodusuario As Integer, fCodModulo As Integer, fNomeOpcao As String, fPrincOpcao As String) As Integer
        Dim cQuery As String
        Dim dtAcesso As DataTable = New DataTable("ESI002")
        'Dim dtGrupoAcesso As DataTable = New DataTable("ESI000")
        Dim nCodOpcao As Integer
        Dim nNivelAcesso As Integer

        nCodOpcao = getCodOpcao(fCodModulo, fNomeOpcao, fPrincOpcao)

        cQuery = "SELECT SI002_NIVACE FROM ESI002 where SI002_CODUSU = " & fCodusuario & _
                " and SI002_CODOPC = " & nCodOpcao

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtAcesso)
        End Using

        If dtAcesso.Rows.Count() > 0 Then
            nNivelAcesso = dtAcesso.Rows(0).Item("SI002_NIVACE")
        Else
            nNivelAcesso = 0
        End If

        If nNivelAcesso < 4 Then
            'Verificar a permissão pelo grupo é Superior
            cQuery = "SELECT SI000_CODUSU, max(SI002_NIVACE) as NIVACE FROM ESI002, ESI000, ESI006 " & _
                " where SI006_CODUSU=SI000_CODUSU and SI002_CODGRU=SI006_CODGRU " & _
                " and SI000_CODUSU = " & fCodusuario & " and SI002_CODOPC = " & nCodOpcao & _
                " group by SI000_CODUSU"

            dtAcesso.Reset()

            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

                ' Preencher o DataTable 
                da.Fill(dtAcesso)
            End Using

            If dtAcesso.Rows.Count() > 0 Then
                If dtAcesso.Rows(0).Item("NIVACE") > nNivelAcesso Then
                    nNivelAcesso = dtAcesso.Rows(0).Item("NIVACE")
                End If
            End If
        End If
        Return nNivelAcesso

    End Function

    Public Function getCodUsuario(ByVal fLoginUsuario As String) As Integer
        Dim cQuery As String
        Dim dtUsuario As DataTable = New DataTable("ESI000")

        cQuery = "SELECT SI000_CODUSU FROM ESI000 where SI000_LGIUSU = '" & Trim(fLoginUsuario) & "'"

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtUsuario)

            If dtUsuario.Rows.Count() > 0 Then
                Return dtUsuario.Rows(0).Item("SI000_CODUSU")
            Else
                Return 0
            End If
        End Using

    End Function

    Public Function getLoginUsuario(ByVal fCodUsuario As Integer) As String
        Dim cQuery As String
        Dim dtUsuario As DataTable = New DataTable("ESI000")

        cQuery = "SELECT SI000_LGIUSU FROM ESI000 where SI000_CODUSU  = " & fCodUsuario.ToString

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtUsuario)

            If dtUsuario.Rows.Count() > 0 Then
                Return dtUsuario.Rows(0).Item("SI000_LGIUSU")
            Else
                Return ""
            End If
        End Using

    End Function

    Public Function getCodGrupo(ByVal fDescrGrupo As String) As Integer
        Dim cQuery As String
        Dim dtGrupo As DataTable = New DataTable("ESI001")

        cQuery = "SELECT SI001_CODGRU FROM ESI001 where SI001_DESGRU = '" & Trim(fDescrGrupo) & "'"

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtGrupo)

            If dtGrupo.Rows.Count() > 0 Then
                Return dtGrupo.Rows(0).Item("SI001_CODGRU")
            Else
                Return 0
            End If
        End Using
        dtGrupo.Clear()

    End Function

    Public Function getDescrGrupo(ByVal fCodGrupo As Integer) As String
        Dim cQuery As String
        Dim dtGrupo As DataTable = New DataTable("ESI001")

        cQuery = "SELECT SI001_DESGRU FROM ESI001 where SI001_CODGRU = " & Trim(fCodGrupo)

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtGrupo)

            If dtGrupo.Rows.Count() > 0 Then
                Return dtGrupo.Rows(0).Item("SI001_DESGRU")
            Else
                Return ""
            End If
        End Using
        dtGrupo.Clear()

    End Function

    Public Function getCodModulo(ByVal fNomeModulo As String) As Integer
        Dim sQuery As String
        Dim dtModulo As DataTable = New DataTable("ESI004")

        sQuery = "SELECT SI004_CODMOD FROM ESI004 where SI004_DESMOD = '" & Trim(fNomeModulo) & "'"

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(sQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtModulo)

            If dtModulo.Rows.Count() > 0 Then
                Return dtModulo.Rows(0).Item("SI004_CODMOD")
            Else
                Return 0
            End If
        End Using

    End Function

    Public Function getCodOpcao(fCodModulo As Integer, fNomeOpcao As String, fPrincOpcao As String) As Integer
        Dim sQuery As String
        Dim dtModulo As DataTable = New DataTable("ESI003")

        sQuery = "SELECT SI003_CODOPC FROM ESI003 where SI003_NOMOPC = '" & Trim(fNomeOpcao) & _
                "' AND SI003_CODMOD = " & fCodModulo.ToString

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(sQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtModulo)

            If dtModulo.Rows.Count() > 0 Then
                Return dtModulo.Rows(0).Item("SI003_CODOPC")
            Else
                'Se a Função não estiver cadastrada, cadastrar
                Return InserirOpcao(fCodModulo, fNomeOpcao, fPrincOpcao)
            End If
        End Using

    End Function

    Private Function InserirOpcao(fCodModulo As Integer, fNomeOpcao As String, fPrincOpcao As String) As Integer
        Dim cSql As String = ""
        Dim cmd As OleDbCommand
        Dim nCodOpcao As Integer

        nCodOpcao = ProxCodOpcao() 'Buscar o próximo código da opção

        'Gravar a Opcao no Banco
        cSql = "INSERT INTO ESI003 (SI003_CODOPC, SI003_DESOPC, SI003_NOMOPC, SI003_PRIOPC, SI003_CODMOD)"
        cSql += " values (" & nCodOpcao & ", '" & fNomeOpcao & "','" & fNomeOpcao & _
                "','" & fPrincOpcao & "'," & fCodModulo & ")"

        cmd = New OleDbCommand(cSql, g_ConnectBanco)

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString())
        Finally
            'Se não ocorrer erros, gravar o acesso para o grupo de administradores
            cSql = "INSERT INTO ESI002 (SI002_CODUSU, SI002_CODGRU, SI002_CODOPC, SI002_NIVACE)"
            cSql += " values (0,1," & nCodOpcao & ",4)"
            'Constantes no Insert ESI002
            'Codigo do usuario = 0 -> Valer a permissão do grupo
            'Codigo Grupo = 1 -> Grupo de Administradores (Padrão)
            'Permissão = 4 -> Acesso Total
            cmd = New OleDbCommand(cSql, g_ConnectBanco)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                MsgBox(ex.ToString())
            End Try
        End Try

        Return nCodOpcao

    End Function

    Public Function ProxCodOpcao() As Double
        Dim cQuery As String
        Dim dtCodOpcao As DataTable = New DataTable("ESI003")

        cQuery = "SELECT max(SI003_CODOPC) as UltCod FROM ESI003"

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtCodOpcao)
            If IsDBNull(dtCodOpcao.Rows(0).Item("UltCod")) Then
                Return 1
            Else
                Return dtCodOpcao.Rows(0).Item("UltCod") + 1
            End If

        End Using
    End Function

    Public Function ProxCodChave(fEntidade As String, fChave As String, Optional fValorChave As String = "") As Double
        Dim cQuery As String
        Dim cCampoMax As String = ""
        Dim cGroupBy As String = ""
        Dim cWhere As String = ""
        Dim x As Integer
        Dim dtCodOpcao As DataTable = New DataTable(fEntidade)
        Dim cCampos() As String = fChave.Split(",")
        Dim cValores() As String = fValorChave.Split(",")

        For x = 0 To cCampos.Length - 1
            If x = cCampos.Length - 1 Then
                cCampoMax = cCampos(x)
            Else
                If x > 0 Then
                    cGroupBy += ", "
                    cWhere += " and "
                End If
                cGroupBy += cCampos(x)
                cWhere += cCampos(x) & "=" & cValores(x)
            End If
        Next
        'If cGroupBy <> "" Then cGroupBy += ", "

        cQuery = "SELECT " & cGroupBy & IIf(Trim(cGroupBy) = "", "", ", ") & "max(" & cCampoMax & ") as UltCod FROM " & fEntidade
        If Not Trim(cWhere) = "" Then cWhere = " Where " & cWhere
        If Not Trim(cGroupBy) = "" Then cGroupBy = " group by " & cGroupBy
        cQuery += cWhere & " " & cGroupBy

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtCodOpcao)
            If dtCodOpcao.Rows.Count() > 0 Then
                If IsDBNull(dtCodOpcao.Rows(0).Item("UltCod")) Then
                    Return 1
                Else
                    Return dtCodOpcao.Rows(0).Item("UltCod") + 1
                End If
            Else
                Return 1
            End If

        End Using
        dtCodOpcao.Clear()

    End Function

    Public Function LerCod_Unidade(fClass_Uni As String, Optional ByRef fNomeUnidade As String = "") As Double
        Dim cQuery As String
        Dim dtLerCodUni As DataTable = New DataTable("EUN000")

        cQuery = "SELECT UN000_CODRED, UN000_NOMUNI FROM EUN000 where UN000_CLAUNI = '" & fClass_Uni & "'"

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtLerCodUni)
            If dtLerCodUni.Rows.Count > 0 Then
                fNomeUnidade = dtLerCodUni.Rows(0).Item("UN000_NOMUNI")
                Return dtLerCodUni.Rows(0).Item("UN000_CODRED")
            Else
                Return 0
            End If
        End Using

        dtLerCodUni.Clear()

    End Function

    Public Function LerClasse_Unidade(fCodReduzido As Integer, Optional ByRef fNomeUnidade As String = "") As String
        Dim cQuery As String
        Dim dtLerCodUni As DataTable = New DataTable("EUN000")

        cQuery = "SELECT UN000_CLAUNI, UN000_NOMUNI FROM EUN000 where UN000_CODRED = " & fCodReduzido

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtLerCodUni)
            If dtLerCodUni.Rows.Count > 0 Then
                fNomeUnidade = dtLerCodUni.Rows(0).Item("UN000_NOMUNI")
                Return dtLerCodUni.Rows(0).Item("UN000_CLAUNI")
            Else
                Return ""
            End If
        End Using

        dtLerCodUni.Clear()

    End Function

    Public Function ProxSeq_Unidade(fClaUnidade As String) As String
        Dim cQuery As String = ""
        Dim dtProxSeqUni As DataTable = New DataTable("EUN000")
        Dim nPos As Integer
        Dim sProxSeq As String = ""

        If fClaUnidade = "00.00.00.00" Then 'CM
            cQuery = "right(UN000_CLAUNI,8)='00.00.00'"
            nPos = 1
        ElseIf Microsoft.VisualBasic.Right(fClaUnidade, 8) = "00.00.00" Then 'CC
            cQuery = "right(UN000_CLAUNI,5)='00.00' and left(UN000_CLAUNI,2) = '" & Microsoft.VisualBasic.Left(fClaUnidade, 2) & "'"
            nPos = 4
        ElseIf Microsoft.VisualBasic.Right(fClaUnidade, 5) = "00.00" Then 'CP
            cQuery = "right(UN000_CLAUNI,2)='00' and left(UN000_CLAUNI,5) = '" & Microsoft.VisualBasic.Left(fClaUnidade, 5) & "'"
            nPos = 7
        Else 'If Microsoft.VisualBasic.Right(fClaUnidade, 2) = "00" Then
            cQuery = "right(UN000_CLAUNI,2)<>'00' and left(UN000_CLAUNI,8) = '" & Microsoft.VisualBasic.Left(fClaUnidade, 8) & "'"
            nPos = 10
        End If

        cQuery = "SELECT max(UN000_CLAUNI) as UltSeq FROM EUN000 " & IIf(Trim(cQuery) = "", "", "where " & cQuery)

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtProxSeqUni)
            If dtProxSeqUni.Rows.Count > 0 Then
                If Not IsDBNull(dtProxSeqUni.Rows(0).Item("UltSeq")) Then
                    sProxSeq = Microsoft.VisualBasic.Mid(dtProxSeqUni.Rows(0).Item("UltSeq"), nPos, 2)
                Else
                    sProxSeq = "1"
                End If
                sProxSeq = Microsoft.VisualBasic.Mid(fClaUnidade, nPos, 2)
                sProxSeq = Microsoft.VisualBasic.Left(fClaUnidade, nPos - 1) & _
                        (Val(sProxSeq) + 1).ToString("00")
                If nPos < 10 Then
                    sProxSeq += CStr((CInt(sProxSeq) + 1)) & Microsoft.VisualBasic.Right(dtProxSeqUni.Rows(0).Item("UltSeq"), 11 - CInt(nPos) - 1)
                End If
                Return sProxSeq
            Else
                Return ""
            End If
        End Using

        dtProxSeqUni.Clear()

    End Function

    Public Function ConfigIcones(fLogin As String, fNomeOpcao As String, ByRef fLocalX As Integer, _
                ByRef fLocalY As Integer, ByRef fCorFundo As Integer) As Integer
        Dim sQuery As String
        Dim dtConfig As DataTable = New DataTable("ESI005")

        sQuery = "SELECT SI005_TAMOBJ, SI005_LOCALX, SI005_LOCALY, SI005_CORFDO FROM ESI005 where " & _
                "SI005_CODUSU = " & getCodUsuario(ClassCrypt.Decrypt(fLogin)) & _
                " and SI005_CODOPC = " & getCodOpcao(0, Trim(fNomeOpcao), "")

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(sQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtConfig)

            If dtConfig.Rows.Count() > 0 Then
                fLocalX = dtConfig.Rows(0).Item("SI005_LOCALX")
                fLocalY = dtConfig.Rows(0).Item("SI005_LOCALY")
                fCorFundo = dtConfig.Rows(0).Item("SI005_CORFDO")
                Return dtConfig.Rows(0).Item("SI005_TAMOBJ")
            Else
                'Se a Função não estiver cadastrada, cadastrar
                If InserirConfigIcones(getCodUsuario(ClassCrypt.Decrypt(fLogin)), getCodOpcao(0, Trim(fNomeOpcao), "")) Then
                    fLocalX = 400
                    fLocalY = 13
                    fCorFundo = 0
                    Return 0 'Tamanho pequeno
                Else
                    fLocalX = 0
                    fLocalY = 0
                    fCorFundo = 0
                    Return 0 'Tamanho pequeno
                End If
            End If
        End Using

    End Function

    Private Function InserirConfigIcones(fCodUsu As Integer, fCodOpcao As Integer) As Boolean
        Dim cSql As String = ""
        Dim cmd As OleDbCommand
        Dim pReturn As Boolean

        'Gravar a Configuração do Icone no Banco para o Usuário
        cSql = "INSERT INTO ESI005 (SI005_CODUSU, SI005_CODOPC, SI005_TAMOBJ, SI005_LOCALX, SI005_LOCALY, SI005_CORFDO)"
        cSql += " values (" & fCodUsu & ", " & fCodOpcao & ",0,400,13,0)"

        cmd = New OleDbCommand(cSql, g_ConnectBanco)

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString())
            pReturn = False
        Finally
            pReturn = True
        End Try

        Return pReturn

    End Function

    Public Function GravarConfigIcones(fCodUsu As Integer, fCodOpcao As Integer, fTamanho As Integer, _
                    fLocalX As Integer, fLocalY As Integer, fCorFundo As Integer) As Boolean
        Dim cSql As String = ""
        Dim cmd As OleDbCommand
        Dim pReturn As Boolean

        'Gravar a Configuração do Icone no Banco para o Usuário
        cSql = "UPDATE ESI005 SET SI005_TAMOBJ=" & fTamanho.ToString & ", SI005_LOCALX=" & _
            fLocalX.ToString & ", SI005_LOCALY=" & fLocalY.ToString & ", SI005_CORFDO=" & _
            fCorFundo.ToString & " where SI005_CODUSU=" & fCodUsu.ToString & _
            " and SI005_CODOPC=" & fCodOpcao.ToString

        cmd = New OleDbCommand(cSql, g_ConnectBanco)

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString())
            pReturn = False
        Finally
            pReturn = True
        End Try

        Return pReturn

    End Function

    Public Function Gravar_LogUnidade(fCodUsuario As Integer, fCodUnidade As Integer, fCampoAlt As String, fValorOld_Alt As String, fValorNew_Alt As String, ByVal fMsgRetorno As String) As Boolean
        Dim sSqlLog As String
        Dim cmd As OleDbCommand
        Dim bResult As Boolean

        'Gravar o Log da Exclusão
        sSqlLog = "INSERT INTO EUN014 (UN014_CODUSU, UN014_CODUNI, UN014_SEQALT, UN014_DATALT, UN014_CPOALT, UN014_VALOLD, UN014_VALNEW)"
        sSqlLog += " values (" & fCodUsuario.ToString & "," & fCodUnidade.ToString & "," & _
            ProxCodChave("EUN014", "UN014_CODUSU,UN014_CODUNI,UN014_SEQALT", fCodUsuario.ToString & "," & fCodUnidade.ToString).ToString & ",'" & FormatarData(Today()) & "'," & _
            "'" & fCampoAlt & "','" & fValorOld_Alt & "','" & fValorNew_Alt & "')"
        cmd = New OleDbCommand(sSqlLog, g_ConnectBanco)

        Try
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            bResult = False
            fMsgRetorno = ex.ToString
        Finally
            bResult = True
            fMsgRetorno = ""
        End Try

        Return bResult

    End Function

    Public Function Gravar_Mural(fMensagem As String, fNivel As Integer, fCodUsuario As Integer, Optional fCodMural As Integer = 0) As Integer
        Dim sSqlMural As String
        Dim cmd As OleDbCommand
        Dim bResult As Boolean = True
        Dim nCodMural As Integer

        'O código do Mural será passado qdo for para incluir Visualizador
        'Gravar o Mural
        If fCodMural = 0 Then
            'Gravar o Mural
            nCodMural = ProxCodChave("ESI007", "SI007_CODMUR")
            sSqlMural = "INSERT INTO ESI007 (SI007_CODMUR, SI007_DATPUB, SI007_MENMUR, SI007_STAMUR) "
            sSqlMural += "VALUES ("
            sSqlMural += nCodMural.ToString()
            sSqlMural += ",'" & FormatarData(Today()) & "'"
            sSqlMural += ",'" & Microsoft.VisualBasic.Left(fMensagem, 100) & "'"
            sSqlMural += ",'PUBLICADA')"
            cmd = New OleDbCommand(sSqlMural, g_ConnectBanco)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                bResult = False
                MsgBox("ModPrincipal.Gravar_Mural: " & ex.ToString & Chr(13) & sSqlMural)
            Finally
                bResult = True
            End Try
        End If

        If bResult Then
            fCodMural = nCodMural
            'Gravar os Acessos a este Mural
            sSqlMural = "INSERT INTO ESI008 (SI008_CODUSU, SI008_CODMUR, SI008_NIVACE)"
            sSqlMural += " values ("
            sSqlMural += fCodUsuario.ToString
            sSqlMural += " ," & fCodMural.ToString
            sSqlMural += " ," & fNivel.ToString
            sSqlMural += ")"
            cmd = New OleDbCommand(sSqlMural, g_ConnectBanco)
            Try
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                bResult = False
                MsgBox("ModPrincipal.Gravar_Mural: " & ex.ToString & Chr(13) & sSqlMural)
            Finally
                bResult = True
            End Try
        End If

        Return nCodMural

    End Function

    Public Function Permissao_Unidade(fClaUnidade As String, fCodUsu As Integer) As Integer
        Dim cQuery As String = ""
        Dim cWhere As String = ""
        Dim dtNivelAcesso As DataTable = New DataTable("EUN013")
        Dim nNivelAcesso As Integer = 0
        '***********************************
        ' 0 - Nenhum Acesso
        ' 1 - Leitura
        ' 2 - Alteração
        ' 3 - Gerenciamento
        '***********************************

        'Ler os Níveis de Acesso
        cQuery = "Select EUN013.UN013_PERACE FROM (EUN013 INNER JOIN EUN000 ON EUN013.UN013_CODUNI=EUN000.UN000_CODRED) " & _
                "WHERE EUN013.UN013_CODUSU=" & fCodUsu.ToString

        'Ver acesso ao CM
        'If Microsoft.VisualBasic.Right(fClaUnidade, 2) <> "00" Then
        cWhere = " and EUN000.UN000_CLAUNI = '" & Microsoft.VisualBasic.Left(fClaUnidade, 2) & ".00.00.00'"
        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery & cWhere, g_ConnectBanco)

            'Preencher o Data Table
            da.Fill(dtNivelAcesso)
            If dtNivelAcesso.Rows.Count() > 0 Then
                nNivelAcesso = dtNivelAcesso.Rows(0).Item("UN013_PERACE")
            End If
        End Using
        dtNivelAcesso.Clear()
        'End If

        'Ver acesso ao CP
        If nNivelAcesso > 0 And Microsoft.VisualBasic.Mid(fClaUnidade, 4, 2) <> "00" Then
            cWhere = " and EUN000.UN000_CLAUNI = '" & Microsoft.VisualBasic.Left(fClaUnidade, 5) & ".00.00'"
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery & cWhere, g_ConnectBanco)

                'Preencher o Data Table
                da.Fill(dtNivelAcesso)
                If dtNivelAcesso.Rows.Count() > 0 Then
                    nNivelAcesso = dtNivelAcesso.Rows(0).Item("UN013_PERACE")
                End If
            End Using
            dtNivelAcesso.Clear()
        End If

        'Ver acesso ao CC
        If nNivelAcesso > 0 And Microsoft.VisualBasic.Mid(fClaUnidade, 7, 2) <> "00" Then
            cWhere = " and EUN000.UN000_CLAUNI = '" & Microsoft.VisualBasic.Left(fClaUnidade, 5) & ".00.00'"
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery & cWhere, g_ConnectBanco)

                'Preencher o Data Table
                da.Fill(dtNivelAcesso)
                If dtNivelAcesso.Rows.Count() > 0 Then
                    nNivelAcesso = dtNivelAcesso.Rows(0).Item("UN013_PERACE")
                End If
            End Using
            dtNivelAcesso.Clear()
        End If

        'Ver acesso ao CF
        If nNivelAcesso > 0 And Microsoft.VisualBasic.Mid(fClaUnidade, 10, 2) <> "00" Then
            cWhere = " and EUN000.UN000_CLAUNI = '" & fClaUnidade & "'"
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery & cWhere, g_ConnectBanco)

                'Preencher o Data Table
                da.Fill(dtNivelAcesso)
                If dtNivelAcesso.Rows.Count() > 0 Then
                    nNivelAcesso = dtNivelAcesso.Rows(0).Item("UN013_PERACE")
                End If
            End Using
            dtNivelAcesso.Clear()
        End If

        Permissao_Unidade = nNivelAcesso

    End Function

    Public Function ExisteRelacao_Grupo(fCodGrupo As Double) As Boolean
        Dim dtExisteGrupo As DataTable = New DataTable("ESI006")
        Dim fQuery As String

        fQuery = "SELECT * FROM ESI006 WHERE ESI006.SI006_CODGRU=" & fCodGrupo.ToString
        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(fQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtExisteGrupo)
            If dtExisteGrupo.Rows.Count() > 0 Then
                ExisteRelacao_Grupo = True
            Else
                ExisteRelacao_Grupo = False
            End If
        End Using
        dtExisteGrupo.Clear()

    End Function

    Public Function FormatarData(fData As Date) As String
        Dim sConexao As String

        sConexao = ClassCrypt.Decrypt(g_ConnectString)

        If InStr(sConexao, "aacd", CompareMethod.Text) > 0 Then
            FormatarData = Format(fData, "dd/MM/yyyy")
        Else
            FormatarData = Format(fData, "yyyy/MM/dd")
        End If

    End Function

    Public Function ChaveConexao(fChave As String) As String
        'Provider=SQLOLEDB;
        'Data Source=192.168.2.1;
        'Initial Catalog=SSVP_CNB;
        'User ID=SSVP_User;
        'Password=ssvp@00;
        'Persist Security Info=true;
        'Connect Timeout=300;

        Dim nContIni As Integer
        Dim nContFin As Integer

        nContIni = InStr(0, ClassCrypt.Decrypt(g_ConnectString), fChave) + Len(fChave) + 1
        nContFin = InStr(nContIni, ClassCrypt.Decrypt(g_ConnectString), ";")

        ChaveConexao = Microsoft.VisualBasic.Mid(ClassCrypt.Decrypt(g_ConnectString), nContIni, nContFin - nContIni)
    End Function

    Public Function LerNomeCampo(sNome_Entidade As String, sDesc_Campo As String) As String
        Dim cQuery As String
        Dim dtLerNomeCampo As DataTable = New DataTable("ESI902")

        cQuery = "SELECT SI902_NOMCPO FROM ESI902 where SI902_DESCPO = '" & sDesc_Campo & "' and SI902_NOMENT='" & sNome_Entidade & "'"

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtLerNomeCampo)
            If dtLerNomeCampo.Rows.Count > 0 Then
                Return dtLerNomeCampo.Rows(0).Item("SI902_NOMCPO")
            Else
                Return 0
            End If
        End Using

        dtLerNomeCampo.Clear()

    End Function

    Public Function FormatarValor_SQL(fValor_String As String) As String
        fValor_String = Replace(fValor_String, ".", "")
        fValor_String = Replace(fValor_String, ",", ".")
        FormatarValor_SQL = fValor_String
    End Function

    Public Function LerNome_Colaborador(fCodigo_Col As Double) As String
        Dim cQuery As String
        Dim dtLerCol As DataTable = New DataTable("EUN003")

        cQuery = "SELECT UN003_NOMCOL FROM EUN003 where UN003_CODCOL = " & fCodigo_Col.ToString

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtLerCol)
            If dtLerCol.Rows.Count > 0 Then
                Return dtLerCol.Rows(0).Item("UN003_NOMCOL")
            Else
                Return 0
            End If
        End Using

        dtLerCol.Clear()

    End Function

    Public Function LerCargo_Colaborador(fCodigo_Col As Double, Optional ByRef sCodUnidade As Double = 0) As String
        Dim cQuery As String
        Dim dtLerCargo As DataTable = New DataTable("EUN011")

        cQuery = "SELECT EUN003.UN003_CODCOL, EUN003.UN003_NOMCOL, EUN011.UN011_DESOCP, EUN016.UN016_DESMDT, " & _
            "EUN016.UN016_CODRED FROM ((EUN003 INNER JOIN EUN012 ON EUN003.UN003_CODCOL = EUN012.UN012_CODCOL) " & _
            "INNER JOIN EUN011 ON EUN012.UN012_CODOCP = EUN011.UN011_CODOCP) INNER JOIN EUN016 ON " & _
            "(EUN012.UN012_CODRED = EUN016.UN016_CODRED) AND (EUN012.UN012_CODMDT = EUN016.UN016_CODMDT) " & _
            "where EUN016.UN016_DATINI <= DATE() AND EUN016.UN016_DATFIN >= DATE() AND " & _
            "UN003_CODCOL = " & fCodigo_Col.ToString

        Using daTabela As New OleDbDataAdapter()
            daTabela.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            daTabela.Fill(dtLerCargo)
            If dtLerCargo.Rows.Count > 0 Then
                sCodUnidade = dtLerCargo.Rows(0).Item("UN016_CODRED")
                Return dtLerCargo.Rows(0).Item("UN011_DESOCP")
            Else
                Return ""
            End If
        End Using

        dtLerCargo.Clear()

    End Function

    Public Function LerCodColaborador(ByVal fLoginUsuario As String) As Integer
        Dim cQuery As String
        Dim dtUsuario As DataTable = New DataTable("ESI000")

        cQuery = "SELECT SI000_CODASS FROM ESI000 where SI000_LGIUSU = '" & Trim(fLoginUsuario) & "'"

        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

            ' Preencher o DataTable 
            da.Fill(dtUsuario)

            If dtUsuario.Rows.Count() > 0 Then
                Return dtUsuario.Rows(0).Item("SI000_CODASS")
            Else
                Return 0
            End If
        End Using

    End Function

    Public Function InserirMandato(fCodUnidade As Double) As Integer
        Dim cmdMandato As OleDbCommand
        Dim sSQL As String
        Dim nProxCodigo As Integer = ProxCodChave("EUN016", "UN016_CODRED,UN016_CODMDT", fCodUnidade)

        sSQL = "INSERT INTO EUN016 (UN016_CODMDT, UN016_CODRED, UN016_DESMDT, UN016_DATINI, UN016_DATFIN) VALUES " & _
        "(" & nProxCodigo.ToString & _
        "," & fCodUnidade & ",'MANDATO " & Year(Today).ToString & "','" & _
        FormatarData(Today) & "','" & FormatarData(Today) & "')"
        cmdMandato = New OleDbCommand(sSQL, g_ConnectBanco)

        Try
            cmdMandato.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.ToString())
            InserirMandato = 0
        Finally
            InserirMandato = nProxCodigo
        End Try

    End Function

    Public Function getParametro(fListaParametro As String, fSeparador As String, fItem As Integer) As String

        If Trim(fListaParametro) <> "" And Trim(fSeparador) <> "" And fItem > 0 Then
            Dim sParametros() As String = fListaParametro.Split("|")

            getParametro = Trim(sParametros(fItem - 1))
        Else
            getParametro = ""
        End If

    End Function

End Module
