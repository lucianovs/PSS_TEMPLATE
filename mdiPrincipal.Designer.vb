<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mdiPrincipal
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(mdiPrincipal))
        Me.MenuStrip = New System.Windows.Forms.MenuStrip()
        Me.menuCadastro = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator5 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuSair = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuProcessos = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuConsultas = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuRelatorios = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuRelUsuario = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuSistema = New System.Windows.Forms.ToolStripMenuItem()
        Me.menuSisConfiguracoes = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.menuSisUsuarios = New System.Windows.Forms.ToolStripMenuItem()
        Me.Ajuda = New System.Windows.Forms.ToolStripMenuItem()
        Me.ContentsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.IndexToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SearchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator8 = New System.Windows.Forms.ToolStripSeparator()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.CarregarUnidadesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip
        '
        Me.MenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuCadastro, Me.menuProcessos, Me.menuConsultas, Me.menuRelatorios, Me.menuSistema, Me.Ajuda})
        Me.MenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip.MdiWindowListItem = Me.menuRelatorios
        Me.MenuStrip.Name = "MenuStrip"
        Me.MenuStrip.Padding = New System.Windows.Forms.Padding(8, 2, 0, 2)
        Me.MenuStrip.Size = New System.Drawing.Size(933, 28)
        Me.MenuStrip.TabIndex = 5
        Me.MenuStrip.Text = "MenuStrip"
        '
        'menuCadastro
        '
        Me.menuCadastro.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripSeparator5, Me.mnuSair})
        Me.menuCadastro.ImageTransparentColor = System.Drawing.SystemColors.ActiveBorder
        Me.menuCadastro.Name = "menuCadastro"
        Me.menuCadastro.Size = New System.Drawing.Size(86, 24)
        Me.menuCadastro.Text = "&Cadastros"
        '
        'ToolStripSeparator5
        '
        Me.ToolStripSeparator5.Name = "ToolStripSeparator5"
        Me.ToolStripSeparator5.Size = New System.Drawing.Size(100, 6)
        '
        'mnuSair
        '
        Me.mnuSair.Name = "mnuSair"
        Me.mnuSair.Size = New System.Drawing.Size(103, 24)
        Me.mnuSair.Text = "Sair"
        '
        'menuProcessos
        '
        Me.menuProcessos.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CarregarUnidadesToolStripMenuItem})
        Me.menuProcessos.Name = "menuProcessos"
        Me.menuProcessos.Size = New System.Drawing.Size(85, 24)
        Me.menuProcessos.Text = "Processos"
        '
        'menuConsultas
        '
        Me.menuConsultas.Name = "menuConsultas"
        Me.menuConsultas.Size = New System.Drawing.Size(84, 24)
        Me.menuConsultas.Text = "Consultas"
        '
        'menuRelatorios
        '
        Me.menuRelatorios.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuRelUsuario})
        Me.menuRelatorios.Name = "menuRelatorios"
        Me.menuRelatorios.Size = New System.Drawing.Size(88, 24)
        Me.menuRelatorios.Text = "&Relatórios"
        '
        'menuRelUsuario
        '
        Me.menuRelUsuario.Name = "menuRelUsuario"
        Me.menuRelUsuario.Size = New System.Drawing.Size(220, 24)
        Me.menuRelUsuario.Text = "Impressao do usuario"
        '
        'menuSistema
        '
        Me.menuSistema.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.menuSisConfiguracoes, Me.ToolStripMenuItem1, Me.menuSisUsuarios})
        Me.menuSistema.Name = "menuSistema"
        Me.menuSistema.Size = New System.Drawing.Size(73, 24)
        Me.menuSistema.Text = "&Sistema"
        '
        'menuSisConfiguracoes
        '
        Me.menuSisConfiguracoes.Name = "menuSisConfiguracoes"
        Me.menuSisConfiguracoes.Size = New System.Drawing.Size(173, 24)
        Me.menuSisConfiguracoes.Text = "Configurações"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(170, 6)
        '
        'menuSisUsuarios
        '
        Me.menuSisUsuarios.Name = "menuSisUsuarios"
        Me.menuSisUsuarios.Size = New System.Drawing.Size(173, 24)
        Me.menuSisUsuarios.Text = "Usuários"
        '
        'Ajuda
        '
        Me.Ajuda.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ContentsToolStripMenuItem, Me.IndexToolStripMenuItem, Me.SearchToolStripMenuItem, Me.ToolStripSeparator8, Me.AboutToolStripMenuItem})
        Me.Ajuda.Name = "Ajuda"
        Me.Ajuda.Size = New System.Drawing.Size(60, 24)
        Me.Ajuda.Text = "&Ajuda"
        '
        'ContentsToolStripMenuItem
        '
        Me.ContentsToolStripMenuItem.Name = "ContentsToolStripMenuItem"
        Me.ContentsToolStripMenuItem.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.F1), System.Windows.Forms.Keys)
        Me.ContentsToolStripMenuItem.Size = New System.Drawing.Size(193, 24)
        Me.ContentsToolStripMenuItem.Text = "&Contents"
        '
        'IndexToolStripMenuItem
        '
        Me.IndexToolStripMenuItem.Image = CType(resources.GetObject("IndexToolStripMenuItem.Image"), System.Drawing.Image)
        Me.IndexToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Black
        Me.IndexToolStripMenuItem.Name = "IndexToolStripMenuItem"
        Me.IndexToolStripMenuItem.Size = New System.Drawing.Size(193, 24)
        Me.IndexToolStripMenuItem.Text = "&Index"
        '
        'SearchToolStripMenuItem
        '
        Me.SearchToolStripMenuItem.Image = CType(resources.GetObject("SearchToolStripMenuItem.Image"), System.Drawing.Image)
        Me.SearchToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Black
        Me.SearchToolStripMenuItem.Name = "SearchToolStripMenuItem"
        Me.SearchToolStripMenuItem.Size = New System.Drawing.Size(193, 24)
        Me.SearchToolStripMenuItem.Text = "&Search"
        '
        'ToolStripSeparator8
        '
        Me.ToolStripSeparator8.Name = "ToolStripSeparator8"
        Me.ToolStripSeparator8.Size = New System.Drawing.Size(190, 6)
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(193, 24)
        Me.AboutToolStripMenuItem.Text = "&About ..."
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 533)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Padding = New System.Windows.Forms.Padding(1, 0, 19, 0)
        Me.StatusStrip.Size = New System.Drawing.Size(933, 25)
        Me.StatusStrip.TabIndex = 7
        Me.StatusStrip.Text = "StatusStrip"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(49, 20)
        Me.ToolStripStatusLabel.Text = "Status"
        '
        'CarregarUnidadesToolStripMenuItem
        '
        Me.CarregarUnidadesToolStripMenuItem.Name = "CarregarUnidadesToolStripMenuItem"
        Me.CarregarUnidadesToolStripMenuItem.Size = New System.Drawing.Size(201, 24)
        Me.CarregarUnidadesToolStripMenuItem.Text = "Carregar Unidades"
        '
        'mdiPrincipal
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(933, 558)
        Me.Controls.Add(Me.MenuStrip)
        Me.Controls.Add(Me.StatusStrip)
        Me.IsMdiContainer = True
        Me.MainMenuStrip = Me.MenuStrip
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.Name = "mdiPrincipal"
        Me.Text = "TEMPLATE"
        Me.MenuStrip.ResumeLayout(False)
        Me.MenuStrip.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ContentsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Ajuda As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents IndexToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SearchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripSeparator8 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuRelatorios As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolTip As System.Windows.Forms.ToolTip
    Friend WithEvents ToolStripStatusLabel As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents ToolStripSeparator5 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents menuCadastro As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents menuSistema As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuProcessos As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuConsultas As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuSisConfiguracoes As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents mnuSair As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuSisUsuarios As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents menuRelUsuario As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CarregarUnidadesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
