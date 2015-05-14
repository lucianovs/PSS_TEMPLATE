Imports System.Data.OleDb

Public Class classLoadListView

    Public Shared Function loadlistview(ByVal fTabela As String, ByVal listviewName As ListView, Optional sJoin As String = "", Optional sWhere As String = "", Optional nBloco As Double = 50) As String

        'Dim command As New SqlCommand
        Dim cQuery, sSQL_Where, sSQL_Select, sSQL_Order As String
        Dim dtSI903 As DataTable = New DataTable("ESI903")
        Dim dtSI902 As DataTable = New DataTable("ESI902")
        Dim cmdSI903 As OleDbCommand
        Dim dtBrowse As DataTable = New DataTable(fTabela)

        Dim sCabec As String
        Dim nCodUsuario As Integer

        Dim i As Integer
        Dim ix As Integer
        Dim iy As Integer
        Dim IColuna As Integer = 0

        Dim ix_fin As Double = nBloco + LerDadosINI(nomeArquivoINI(), "BROWSE", "Bloco_Browse", 50)

        With listviewName
            .BeginUpdate()
            .Clear()
            '.Width = 844
            '.Height = 283
            .View = View.Details
            .GridLines = True
            .FullRowSelect = True
            .HideSelection = False
            .MultiSelect = False
        End With

        nCodUsuario = getCodUsuario(ClassCrypt.Decrypt(g_Login))

        If Not ConectarBanco() Then MsgBox("erro ao conectar com o bd")

        'Verificar se existe cadastro dos campos no browse
        cQuery = "Select * from ESI903 where SI903_NOMENT = '" & fTabela & _
                "' and SI903_CODUSU=" & nCodUsuario.ToString & " and SI903_NOMCPO <> 'BROWSE'"
        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)
            da.Fill(dtSI903)
        End Using

        If dtSI903.Rows.Count() = 0 Then 'Não tem Parâmetros do browse cadastrado, inserir
            'carregar os campos
            cQuery = "Select * from ESI902 where SI902_NOMENT = '" & fTabela & "' order by SI902_SEQCPO"
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)
                da.Fill(dtSI902)

                For i = 0 To dtSI902.Rows.Count() - 1
                    If dtSI902.Rows(i).Item("SI902_STABRW") < 2 Then
                        cQuery = "Insert Into ESI903 (SI903_CODUSU, SI903_NOMENT, SI903_NOMCPO, SI903_TAMBRW, " & _
                            "SI903_POSBRW, SI903_SQLBRW, SI903_SQLORD, SI903_VERCPO) VALUES (" & _
                            nCodUsuario.ToString & ",'" & fTabela & "','" & dtSI902.Rows(i).Item("SI902_NOMCPO") & _
                            "'," & dtSI902.Rows(i).Item("SI902_TAMCPO") & "," & dtSI902.Rows(i).Item("SI902_SEQCPO").ToString & ",''," & _
                            IIf(dtSI902.Rows(i).Item("SI902_CPOKEY") = 1, "'ASC'", "''") & "," & _
                            IIf(dtSI902.Rows(i).Item("SI902_STABRW") = 0, 1, 0) & ")"
                        cmdSI903 = New OleDbCommand(cQuery, g_ConnectBanco)
                        Try
                            cmdSI903.ExecuteNonQuery()
                        Catch ex As Exception
                            MsgBox(ex.ToString())
                            Exit For
                        Finally
                        End Try
                    End If
                Next
            End Using
        End If
        dtSI903.Clear()

        sSQL_Order = ""
        sSQL_Select = ""
        sSQL_Where = ""
        'Montar o Comando SQL
        '* Ler a tabela e resgatar os campos e filtros. Montar a estrutura do listView
        cQuery = "Select SI903_NOMCPO, SI903_TAMBRW, SI903_POSBRW, SI903_SQLBRW, SI903_SQLORD, " & _
                "SI902_CPOKEY, SI902_DESCPO, SI902_STABRW, SI903_VERCPO from ESI903, ESI902 where SI902_NOMCPO=SI903_NOMCPO and " & _
                "SI902_NOMENT=SI903_NOMENT and SI903_NOMENT = '" & fTabela & _
                "' and SI903_CODUSU=" & nCodUsuario.ToString & _
                " order by SI903_POSBRW"
        Using da As New OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)
            da.Fill(dtSI903)
        End Using

        For i = 0 To dtSI903.Rows.Count() - 1
            'Verificar se o campo é permitido
            If dtSI903.Rows(i).Item("SI903_VERCPO") = 1 Then
                'Estrutura
                If dtSI903.Rows(i).Item("SI902_CPOKEY") = 1 Then
                    sCabec = "{" & dtSI903.Rows(i).Item("SI902_DESCPO") & "}"
                Else
                    sCabec = dtSI903.Rows(i).Item("SI902_DESCPO")
                End If
                listviewName.Columns.Add(sCabec, _
                                        dtSI903.Rows(i).Item("SI903_TAMBRW"), HorizontalAlignment.Left)

                'SQL
                If Not sSQL_Select = "" Then sSQL_Select += ", "
                sSQL_Select += dtSI903.Rows(i).Item("SI903_NOMCPO")
                IColuna += 1
                'Where
                If Not IsDBNull(dtSI903.Rows(i).Item("SI903_SQLBRW")) Then
                    If Not dtSI903.Rows(i).Item("SI903_SQLBRW") = "" Then
                        If Not sSQL_Where = "" Then sSQL_Where = sSQL_Where & " and "
                        sSQL_Where = sSQL_Where & dtSI903.Rows(i).Item("SI903_SQLBRW")
                        sSQL_Where = sSQL_Where.Replace("\/", "'")
                    End If
                End If
                'sSQL_Where = sSQL_Where.Replace("\/", "'")

                If Not IsDBNull(dtSI903.Rows(i).Item("SI903_SQLORD")) Then
                    If dtSI903.Rows(i).Item("SI903_SQLORD") <> "" Then
                        sSQL_Order = dtSI903.Rows(i).Item("SI903_NOMCPO") & " " & dtSI903.Rows(i).Item("SI903_SQLORD")
                    End If
                End If
            End If
        Next

        cQuery = "SELECT " & sSQL_Select & " from " & IIf(sJoin <> "", "(", "") & fTabela & " " & sJoin & IIf(sJoin <> "", ")", "")
        If Not sSQL_Where = "" Then
            cQuery += sSQL_Where & IIf(sWhere <> "", " and ", "")
        ElseIf sWhere <> "" Then
            cQuery += " where "
        End If
        cQuery += sWhere
        If Not sSQL_Order = "" Then
            cQuery += " order by " & sSQL_Order
        End If

        'Carregar os Dados 
        Try
            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)
                da.Fill(dtBrowse)
            End Using

            'Carregar os dados para a Lista
            'For ix = 0 To dtBrowse.Rows.Count() - 1
            For ix = nBloco To dtBrowse.Rows.Count() - 1
                Dim lvi As New ListViewItem
                If Not IsDBNull(dtBrowse.Rows(ix).Item(0)) Then
                    lvi.Text = dtBrowse.Rows(ix).Item(0)
                Else
                    lvi.Text = "null"
                End If

                For iy = 1 To IColuna - 1
                    If Not IsDBNull(dtBrowse.Rows(ix).Item(iy)) Then
                        lvi.SubItems.Add(dtBrowse.Rows(ix).Item(iy))
                    Else
                        lvi.SubItems.Add("")
                    End If
                Next iy
                listviewName.Items.Add(lvi)

                'Parar se registro sair do Bloco
                If ix = (ix_fin - 1) Then
                    Exit For
                End If

            Next ix
            'listviewName.Items.Add(lvi)
            'fTabela = sSQL_Where
            listviewName.EndUpdate()
            Return sSQL_Where
        Catch ex As Exception
            'fTabela = sSQL_Where
            Return ex.Message

        End Try

    End Function

End Class
