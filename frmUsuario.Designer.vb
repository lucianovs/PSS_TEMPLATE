<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUsuario
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmUsuario))
        Me.ToolStrip1 = New System.Windows.Forms.ToolStrip()
        Me.btnIncluir = New System.Windows.Forms.ToolStripButton()
        Me.btnAlterar = New System.Windows.Forms.ToolStripButton()
        Me.btnExcluir = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnGravar = New System.Windows.Forms.ToolStripButton()
        Me.btnCancelar = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator3 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnAnterior = New System.Windows.Forms.ToolStripButton()
        Me.btnProximo = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.btnLocalizar = New System.Windows.Forms.ToolStripButton()
        Me.btnImprimir = New System.Windows.Forms.ToolStripButton()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.tssContReg = New System.Windows.Forms.ToolStripStatusLabel()
        Me.lblSenha2 = New System.Windows.Forms.Label()
        Me.txtSenha2 = New System.Windows.Forms.TextBox()
        Me.lblSenha = New System.Windows.Forms.Label()
        Me.txtSenha = New System.Windows.Forms.TextBox()
        Me.lblNmUsuario = New System.Windows.Forms.Label()
        Me.txtNmUsuario = New System.Windows.Forms.TextBox()
        Me.lblLogin = New System.Windows.Forms.Label()
        Me.txtLogin = New System.Windows.Forms.TextBox()
        Me.lblGrupo = New System.Windows.Forms.Label()
        Me.cbGrupoPrincipal = New System.Windows.Forms.ComboBox()
        Me.lstGrupoAssoc = New System.Windows.Forms.ListBox()
        Me.lblGrupoAssoc = New System.Windows.Forms.Label()
        Me.lblDtpExpira = New System.Windows.Forms.Label()
        Me.dtpExpira = New System.Windows.Forms.DateTimePicker()
        Me.chkValidade = New System.Windows.Forms.CheckBox()
        Me.chkAlterarSenha = New System.Windows.Forms.CheckBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.chkGerAgr = New System.Windows.Forms.CheckBox()
        Me.lblColaborador = New System.Windows.Forms.Label()
        Me.txtColaborador = New System.Windows.Forms.TextBox()
        Me.btnLocColaborador = New System.Windows.Forms.Button()
        Me.ToolStrip1.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ToolStrip1
        '
        Me.ToolStrip1.ImageScalingSize = New System.Drawing.Size(32, 32)
        Me.ToolStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.btnIncluir, Me.btnAlterar, Me.btnExcluir, Me.ToolStripSeparator1, Me.btnGravar, Me.btnCancelar, Me.ToolStripSeparator3, Me.btnAnterior, Me.btnProximo, Me.ToolStripSeparator2, Me.btnLocalizar, Me.btnImprimir})
        Me.ToolStrip1.Location = New System.Drawing.Point(0, 0)
        Me.ToolStrip1.Name = "ToolStrip1"
        Me.ToolStrip1.RenderMode = System.Windows.Forms.ToolStripRenderMode.System
        Me.ToolStrip1.Size = New System.Drawing.Size(759, 39)
        Me.ToolStrip1.Stretch = True
        Me.ToolStrip1.TabIndex = 1
        Me.ToolStrip1.Text = "ToolStrip1"
        '
        'btnIncluir
        '
        Me.btnIncluir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnIncluir.Image = CType(resources.GetObject("btnIncluir.Image"), System.Drawing.Image)
        Me.btnIncluir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnIncluir.Name = "btnIncluir"
        Me.btnIncluir.Size = New System.Drawing.Size(36, 36)
        Me.btnIncluir.Text = "Incluir"
        '
        'btnAlterar
        '
        Me.btnAlterar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnAlterar.Image = CType(resources.GetObject("btnAlterar.Image"), System.Drawing.Image)
        Me.btnAlterar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnAlterar.Name = "btnAlterar"
        Me.btnAlterar.Size = New System.Drawing.Size(36, 36)
        Me.btnAlterar.Text = "Alterar"
        '
        'btnExcluir
        '
        Me.btnExcluir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnExcluir.Image = CType(resources.GetObject("btnExcluir.Image"), System.Drawing.Image)
        Me.btnExcluir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnExcluir.Name = "btnExcluir"
        Me.btnExcluir.Size = New System.Drawing.Size(36, 36)
        Me.btnExcluir.Text = "Excluir"
        '
        'ToolStripSeparator1
        '
        Me.ToolStripSeparator1.Name = "ToolStripSeparator1"
        Me.ToolStripSeparator1.Size = New System.Drawing.Size(6, 39)
        '
        'btnGravar
        '
        Me.btnGravar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnGravar.Image = CType(resources.GetObject("btnGravar.Image"), System.Drawing.Image)
        Me.btnGravar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnGravar.Name = "btnGravar"
        Me.btnGravar.Size = New System.Drawing.Size(36, 36)
        Me.btnGravar.Text = "Gravar"
        '
        'btnCancelar
        '
        Me.btnCancelar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnCancelar.Image = CType(resources.GetObject("btnCancelar.Image"), System.Drawing.Image)
        Me.btnCancelar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnCancelar.Name = "btnCancelar"
        Me.btnCancelar.Size = New System.Drawing.Size(36, 36)
        Me.btnCancelar.Text = "Cancelar"
        '
        'ToolStripSeparator3
        '
        Me.ToolStripSeparator3.Name = "ToolStripSeparator3"
        Me.ToolStripSeparator3.Size = New System.Drawing.Size(6, 39)
        '
        'btnAnterior
        '
        Me.btnAnterior.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnAnterior.Image = CType(resources.GetObject("btnAnterior.Image"), System.Drawing.Image)
        Me.btnAnterior.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnAnterior.Name = "btnAnterior"
        Me.btnAnterior.Size = New System.Drawing.Size(36, 36)
        Me.btnAnterior.Text = "Anterior"
        '
        'btnProximo
        '
        Me.btnProximo.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnProximo.Image = CType(resources.GetObject("btnProximo.Image"), System.Drawing.Image)
        Me.btnProximo.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnProximo.Name = "btnProximo"
        Me.btnProximo.Size = New System.Drawing.Size(36, 36)
        Me.btnProximo.Text = "Próximo"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 39)
        '
        'btnLocalizar
        '
        Me.btnLocalizar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnLocalizar.Image = CType(resources.GetObject("btnLocalizar.Image"), System.Drawing.Image)
        Me.btnLocalizar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnLocalizar.Name = "btnLocalizar"
        Me.btnLocalizar.Size = New System.Drawing.Size(36, 36)
        Me.btnLocalizar.Text = "Localizar"
        '
        'btnImprimir
        '
        Me.btnImprimir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.btnImprimir.Image = CType(resources.GetObject("btnImprimir.Image"), System.Drawing.Image)
        Me.btnImprimir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(36, 36)
        Me.btnImprimir.Text = "Imprimir"
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tssContReg})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 321)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip1.Size = New System.Drawing.Size(759, 25)
        Me.StatusStrip1.TabIndex = 9
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'tssContReg
        '
        Me.tssContReg.Name = "tssContReg"
        Me.tssContReg.Size = New System.Drawing.Size(98, 20)
        Me.tssContReg.Text = "Registro n / n"
        '
        'lblSenha2
        '
        Me.lblSenha2.AutoSize = True
        Me.lblSenha2.Location = New System.Drawing.Point(221, 217)
        Me.lblSenha2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSenha2.Name = "lblSenha2"
        Me.lblSenha2.Size = New System.Drawing.Size(125, 17)
        Me.lblSenha2.TabIndex = 5
        Me.lblSenha2.Text = "Confirme a Senha:"
        '
        'txtSenha2
        '
        Me.txtSenha2.Location = New System.Drawing.Point(346, 212)
        Me.txtSenha2.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSenha2.MaxLength = 8
        Me.txtSenha2.Name = "txtSenha2"
        Me.txtSenha2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSenha2.Size = New System.Drawing.Size(132, 22)
        Me.txtSenha2.TabIndex = 5
        '
        'lblSenha
        '
        Me.lblSenha.AutoSize = True
        Me.lblSenha.Location = New System.Drawing.Point(17, 217)
        Me.lblSenha.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSenha.Name = "lblSenha"
        Me.lblSenha.Size = New System.Drawing.Size(53, 17)
        Me.lblSenha.TabIndex = 4
        Me.lblSenha.Text = "Senha:"
        '
        'txtSenha
        '
        Me.txtSenha.Location = New System.Drawing.Point(70, 212)
        Me.txtSenha.Margin = New System.Windows.Forms.Padding(4)
        Me.txtSenha.MaxLength = 8
        Me.txtSenha.Name = "txtSenha"
        Me.txtSenha.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtSenha.Size = New System.Drawing.Size(132, 22)
        Me.txtSenha.TabIndex = 4
        '
        'lblNmUsuario
        '
        Me.lblNmUsuario.AutoSize = True
        Me.lblNmUsuario.Location = New System.Drawing.Point(17, 105)
        Me.lblNmUsuario.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblNmUsuario.Name = "lblNmUsuario"
        Me.lblNmUsuario.Size = New System.Drawing.Size(122, 17)
        Me.lblNmUsuario.TabIndex = 2
        Me.lblNmUsuario.Text = "Nome do Usuário:"
        '
        'txtNmUsuario
        '
        Me.txtNmUsuario.Location = New System.Drawing.Point(147, 105)
        Me.txtNmUsuario.Margin = New System.Windows.Forms.Padding(4)
        Me.txtNmUsuario.MaxLength = 40
        Me.txtNmUsuario.Name = "txtNmUsuario"
        Me.txtNmUsuario.Size = New System.Drawing.Size(340, 22)
        Me.txtNmUsuario.TabIndex = 2
        '
        'lblLogin
        '
        Me.lblLogin.AutoSize = True
        Me.lblLogin.Location = New System.Drawing.Point(17, 65)
        Me.lblLogin.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblLogin.Name = "lblLogin"
        Me.lblLogin.Size = New System.Drawing.Size(47, 17)
        Me.lblLogin.TabIndex = 1
        Me.lblLogin.Text = "Login:"
        '
        'txtLogin
        '
        Me.txtLogin.Location = New System.Drawing.Point(68, 65)
        Me.txtLogin.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLogin.MaxLength = 15
        Me.txtLogin.Name = "txtLogin"
        Me.txtLogin.Size = New System.Drawing.Size(132, 22)
        Me.txtLogin.TabIndex = 1
        '
        'lblGrupo
        '
        Me.lblGrupo.AutoSize = True
        Me.lblGrupo.Location = New System.Drawing.Point(17, 144)
        Me.lblGrupo.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblGrupo.Name = "lblGrupo"
        Me.lblGrupo.Size = New System.Drawing.Size(110, 17)
        Me.lblGrupo.TabIndex = 3
        Me.lblGrupo.Text = "Grupo Principal:"
        '
        'cbGrupoPrincipal
        '
        Me.cbGrupoPrincipal.FormattingEnabled = True
        Me.cbGrupoPrincipal.Location = New System.Drawing.Point(147, 137)
        Me.cbGrupoPrincipal.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cbGrupoPrincipal.Name = "cbGrupoPrincipal"
        Me.cbGrupoPrincipal.Size = New System.Drawing.Size(340, 24)
        Me.cbGrupoPrincipal.TabIndex = 3
        '
        'lstGrupoAssoc
        '
        Me.lstGrupoAssoc.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.lstGrupoAssoc.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lstGrupoAssoc.FormattingEnabled = True
        Me.lstGrupoAssoc.ItemHeight = 16
        Me.lstGrupoAssoc.Location = New System.Drawing.Point(555, 62)
        Me.lstGrupoAssoc.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.lstGrupoAssoc.Name = "lstGrupoAssoc"
        Me.lstGrupoAssoc.Size = New System.Drawing.Size(192, 196)
        Me.lstGrupoAssoc.TabIndex = 11
        '
        'lblGrupoAssoc
        '
        Me.lblGrupoAssoc.AutoSize = True
        Me.lblGrupoAssoc.Location = New System.Drawing.Point(557, 44)
        Me.lblGrupoAssoc.Name = "lblGrupoAssoc"
        Me.lblGrupoAssoc.Size = New System.Drawing.Size(68, 17)
        Me.lblGrupoAssoc.TabIndex = 11
        Me.lblGrupoAssoc.Text = "GRUPOS"
        '
        'lblDtpExpira
        '
        Me.lblDtpExpira.AutoSize = True
        Me.lblDtpExpira.Location = New System.Drawing.Point(19, 179)
        Me.lblDtpExpira.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDtpExpira.Name = "lblDtpExpira"
        Me.lblDtpExpira.Size = New System.Drawing.Size(134, 17)
        Me.lblDtpExpira.TabIndex = 10
        Me.lblDtpExpira.Text = "Cadastro expira em:"
        '
        'dtpExpira
        '
        Me.dtpExpira.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpExpira.Location = New System.Drawing.Point(152, 178)
        Me.dtpExpira.Name = "dtpExpira"
        Me.dtpExpira.Size = New System.Drawing.Size(91, 22)
        Me.dtpExpira.TabIndex = 10
        '
        'chkValidade
        '
        Me.chkValidade.AutoSize = True
        Me.chkValidade.Location = New System.Drawing.Point(251, 179)
        Me.chkValidade.Name = "chkValidade"
        Me.chkValidade.Size = New System.Drawing.Size(114, 21)
        Me.chkValidade.TabIndex = 12
        Me.chkValidade.Text = "Nunca Expira"
        Me.chkValidade.UseVisualStyleBackColor = True
        '
        'chkAlterarSenha
        '
        Me.chkAlterarSenha.AutoSize = True
        Me.chkAlterarSenha.Location = New System.Drawing.Point(23, 244)
        Me.chkAlterarSenha.Name = "chkAlterarSenha"
        Me.chkAlterarSenha.Size = New System.Drawing.Size(236, 21)
        Me.chkAlterarSenha.TabIndex = 13
        Me.chkAlterarSenha.Text = "Alterar a Senha no próximo login"
        Me.chkAlterarSenha.UseVisualStyleBackColor = True
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 200
        '
        'chkGerAgr
        '
        Me.chkGerAgr.AutoSize = True
        Me.chkGerAgr.Location = New System.Drawing.Point(314, 244)
        Me.chkGerAgr.Name = "chkGerAgr"
        Me.chkGerAgr.Size = New System.Drawing.Size(173, 21)
        Me.chkGerAgr.TabIndex = 14
        Me.chkGerAgr.Text = "Gerenciar Agregações"
        Me.chkGerAgr.UseVisualStyleBackColor = True
        '
        'lblColaborador
        '
        Me.lblColaborador.AutoSize = True
        Me.lblColaborador.Location = New System.Drawing.Point(20, 283)
        Me.lblColaborador.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblColaborador.Name = "lblColaborador"
        Me.lblColaborador.Size = New System.Drawing.Size(90, 17)
        Me.lblColaborador.TabIndex = 15
        Me.lblColaborador.Text = "Colaborador:"
        '
        'txtColaborador
        '
        Me.txtColaborador.Location = New System.Drawing.Point(111, 281)
        Me.txtColaborador.Margin = New System.Windows.Forms.Padding(4)
        Me.txtColaborador.MaxLength = 150
        Me.txtColaborador.Name = "txtColaborador"
        Me.txtColaborador.Size = New System.Drawing.Size(367, 22)
        Me.txtColaborador.TabIndex = 16
        '
        'btnLocColaborador
        '
        Me.btnLocColaborador.BackgroundImage = Global.UNIDADES.My.Resources.Resources.AllDay_ru_Search
        Me.btnLocColaborador.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnLocColaborador.Location = New System.Drawing.Point(477, 280)
        Me.btnLocColaborador.Name = "btnLocColaborador"
        Me.btnLocColaborador.Size = New System.Drawing.Size(27, 23)
        Me.btnLocColaborador.TabIndex = 17
        Me.btnLocColaborador.UseVisualStyleBackColor = True
        '
        'frmUsuario
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(759, 346)
        Me.Controls.Add(Me.btnLocColaborador)
        Me.Controls.Add(Me.txtColaborador)
        Me.Controls.Add(Me.lblColaborador)
        Me.Controls.Add(Me.chkGerAgr)
        Me.Controls.Add(Me.chkAlterarSenha)
        Me.Controls.Add(Me.chkValidade)
        Me.Controls.Add(Me.dtpExpira)
        Me.Controls.Add(Me.lblDtpExpira)
        Me.Controls.Add(Me.lblGrupoAssoc)
        Me.Controls.Add(Me.lstGrupoAssoc)
        Me.Controls.Add(Me.cbGrupoPrincipal)
        Me.Controls.Add(Me.lblGrupo)
        Me.Controls.Add(Me.lblSenha2)
        Me.Controls.Add(Me.txtSenha2)
        Me.Controls.Add(Me.lblSenha)
        Me.Controls.Add(Me.txtSenha)
        Me.Controls.Add(Me.lblNmUsuario)
        Me.Controls.Add(Me.txtNmUsuario)
        Me.Controls.Add(Me.lblLogin)
        Me.Controls.Add(Me.txtLogin)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.ToolStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "frmUsuario"
        Me.Text = "frmUsuario"
        Me.ToolStrip1.ResumeLayout(False)
        Me.ToolStrip1.PerformLayout()
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolStrip1 As System.Windows.Forms.ToolStrip
    Friend WithEvents btnIncluir As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnAlterar As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnExcluir As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnGravar As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnCancelar As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator3 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnAnterior As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnProximo As System.Windows.Forms.ToolStripButton
    Friend WithEvents ToolStripSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents btnLocalizar As System.Windows.Forms.ToolStripButton
    Friend WithEvents btnImprimir As System.Windows.Forms.ToolStripButton
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents tssContReg As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents lblSenha2 As System.Windows.Forms.Label
    Friend WithEvents txtSenha2 As System.Windows.Forms.TextBox
    Friend WithEvents lblSenha As System.Windows.Forms.Label
    Friend WithEvents txtSenha As System.Windows.Forms.TextBox
    Friend WithEvents lblNmUsuario As System.Windows.Forms.Label
    Friend WithEvents txtNmUsuario As System.Windows.Forms.TextBox
    Friend WithEvents lblLogin As System.Windows.Forms.Label
    Friend WithEvents txtLogin As System.Windows.Forms.TextBox
    Friend WithEvents lblGrupo As System.Windows.Forms.Label
    Friend WithEvents cbGrupoPrincipal As System.Windows.Forms.ComboBox
    Friend WithEvents lstGrupoAssoc As System.Windows.Forms.ListBox
    Friend WithEvents lblGrupoAssoc As System.Windows.Forms.Label
    Friend WithEvents lblDtpExpira As System.Windows.Forms.Label
    Friend WithEvents chkValidade As System.Windows.Forms.CheckBox
    Friend WithEvents chkAlterarSenha As System.Windows.Forms.CheckBox
    Private WithEvents dtpExpira As System.Windows.Forms.DateTimePicker
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents chkGerAgr As System.Windows.Forms.CheckBox
    Friend WithEvents lblColaborador As System.Windows.Forms.Label
    Friend WithEvents txtColaborador As System.Windows.Forms.TextBox
    Friend WithEvents btnLocColaborador As System.Windows.Forms.Button
End Class
