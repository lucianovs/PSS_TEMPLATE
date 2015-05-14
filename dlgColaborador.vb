Imports System.Windows.Forms
Imports System.Data.OleDb

Public Class dlgColaborador

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click, lstPesquisa.DoubleClick
        If lstPesquisa.SelectedItem.ToString <> "" Then
            'Foi selecionado um Colaborador
            txtPesquisa.Text = lstPesquisa.SelectedItem
            'Microsoft.VisualBasic.Left(lstPesquisa.SelectedItem, 6)
            Me.DialogResult = System.Windows.Forms.DialogResult.OK
            Me.Close()
        Else
            MsgBox("Nenhum colaborador foi selecionado!!")
        End If
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub btnPesquisa_Click(sender As Object, e As EventArgs) Handles btnPesquisa.Click
        Dim nPos As Integer = 99
        Dim nSeq As Integer = 0
        Dim nStart As Integer = 1
        Dim sNome(10) As String
        Dim sItemLista As String
        Dim dtPesquisar As DataTable = New DataTable("EUN003")
        Dim cQuery As String

        If Not Trim(txtPesquisa.Text) = "" Then
            cQuery = "Select EUN003.UN003_CODCOL, EUN003.UN003_NOMCOL, EUN003.UN003_BAIRRO, " & _
                "EUN003.UN003_CIDADE, EUN003.UN003_SIGEST, EUN003.UN003_DTNASC from EUN003 " & _
                "where EUN003.UN003_SITCOL<>'I' AND EUN003.UN003_CODUNI=0"
            nStart = 1

            lstPesquisa.Items.Clear()

            Do Until nPos = 0
                nPos = InStr(nStart, txtPesquisa.Text, " ")
                If nPos > 0 Then
                    sNome(nSeq) = Mid(txtPesquisa.Text, nStart, nPos - nStart)
                    nStart = nPos + 1
                    nSeq += 1
                Else
                    sNome(nSeq) = Mid(txtPesquisa.Text, nStart, Len(txtPesquisa.Text) - nStart + 1)
                End If
            Loop
            'Montar a Condição
            For nPos = 0 To nSeq
                cQuery += " and EUN003.UN003_NOMCOL LIKE '%" & sNome(nPos) & "%'"
            Next nPos
            cQuery += " order by EUN003.UN003_NOMCOL"

            Using da As New OleDbDataAdapter()
                da.SelectCommand = New OleDbCommand(cQuery, g_ConnectBanco)

                ' Preencher o DataTable 
                da.Fill(dtPesquisar)
            End Using

            For nSeq = 0 To dtPesquisar.Rows.Count - 1
                sItemLista = Format(dtPesquisar.Rows(nSeq).Item("UN003_CODCOL"), "000000") & " | "
                sItemLista += IIf(IsDBNull(dtPesquisar.Rows(nSeq).Item("UN003_NOMCOL")), "", dtPesquisar.Rows(nSeq).Item("UN003_NOMCOL"))

                If Not IsDBNull(dtPesquisar.Rows(nSeq).Item("UN003_DTNASC")) Then
                    sItemLista += " | " & Format(dtPesquisar.Rows(nSeq).Item("UN003_DTNASC"), "dd/MM/yy")
                End If
                If Not IsDBNull(dtPesquisar.Rows(nSeq).Item("UN003_SIGEST")) Then
                    sItemLista += " | " & dtPesquisar.Rows(nSeq).Item("UN003_SIGEST")
                End If
                If Not IsDBNull(dtPesquisar.Rows(nSeq).Item("UN003_CIDADE")) Then
                    sItemLista += " | " & dtPesquisar.Rows(nSeq).Item("UN003_CIDADE")
                End If
                If Not IsDBNull(dtPesquisar.Rows(nSeq).Item("UN003_BAIRRO")) Then
                    sItemLista += " | " & dtPesquisar.Rows(nSeq).Item("UN003_BAIRRO")
                End If

                lstPesquisa.Items.Add(sItemLista)
            Next nSeq

            dtPesquisar.Clear()
        Else
            MsgBox("Digite um nome para pesquisar. Para melhor performance, digitar nome e sobrenome.'")
        End If

    End Sub

    Private Sub lstPesquisa_DoubleClick(sender As Object, e As EventArgs) Handles lstPesquisa.DoubleClick

    End Sub

    Private Sub lstPesquisa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstPesquisa.SelectedIndexChanged

    End Sub
End Class
