<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmBrowse
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmBrowse))
        Me.ListView_Browse = New System.Windows.Forms.ListView()
        Me.btnSalvar = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cbCampo = New System.Windows.Forms.ComboBox()
        Me.cbCondicao = New System.Windows.Forms.ComboBox()
        Me.txtValorCondicao = New System.Windows.Forms.TextBox()
        Me.btnImprimir = New System.Windows.Forms.Button()
        Me.grbFiltro = New System.Windows.Forms.GroupBox()
        Me.btnIncluir = New System.Windows.Forms.Button()
        Me.btnExcluir = New System.Windows.Forms.Button()
        Me.btnAlterar = New System.Windows.Forms.Button()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnAnterior = New System.Windows.Forms.Button()
        Me.btnProximo = New System.Windows.Forms.Button()
        Me.lblRegistros = New System.Windows.Forms.Label()
        Me.timerRefresh = New System.Windows.Forms.Timer(Me.components)
        Me.grbFiltro.SuspendLayout()
        Me.SuspendLayout()
        '
        'ListView_Browse
        '
        Me.ListView_Browse.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ListView_Browse.FullRowSelect = True
        Me.ListView_Browse.GridLines = True
        Me.ListView_Browse.Location = New System.Drawing.Point(4, 54)
        Me.ListView_Browse.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.ListView_Browse.MultiSelect = False
        Me.ListView_Browse.Name = "ListView_Browse"
        Me.ListView_Browse.Size = New System.Drawing.Size(942, 247)
        Me.ListView_Browse.TabIndex = 0
        Me.ListView_Browse.UseCompatibleStateImageBehavior = False
        '
        'btnSalvar
        '
        Me.btnSalvar.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnSalvar.BackgroundImage = CType(resources.GetObject("btnSalvar.BackgroundImage"), System.Drawing.Image)
        Me.btnSalvar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSalvar.Location = New System.Drawing.Point(678, 13)
        Me.btnSalvar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnSalvar.Name = "btnSalvar"
        Me.btnSalvar.Size = New System.Drawing.Size(36, 34)
        Me.btnSalvar.TabIndex = 1
        Me.btnSalvar.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(13, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(43, 17)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Filtro:"
        '
        'cbCampo
        '
        Me.cbCampo.BackColor = System.Drawing.Color.White
        Me.cbCampo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbCampo.FormattingEnabled = True
        Me.cbCampo.Location = New System.Drawing.Point(6, 15)
        Me.cbCampo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cbCampo.Name = "cbCampo"
        Me.cbCampo.Size = New System.Drawing.Size(201, 24)
        Me.cbCampo.TabIndex = 3
        '
        'cbCondicao
        '
        Me.cbCondicao.BackColor = System.Drawing.Color.White
        Me.cbCondicao.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cbCondicao.FormattingEnabled = True
        Me.cbCondicao.Items.AddRange(New Object() {"Igual a", "Contenha", "Maior que", "Menor que", "Diferente de"})
        Me.cbCondicao.Location = New System.Drawing.Point(224, 15)
        Me.cbCondicao.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.cbCondicao.Name = "cbCondicao"
        Me.cbCondicao.Size = New System.Drawing.Size(145, 24)
        Me.cbCondicao.TabIndex = 4
        '
        'txtValorCondicao
        '
        Me.txtValorCondicao.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtValorCondicao.Location = New System.Drawing.Point(380, 15)
        Me.txtValorCondicao.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.txtValorCondicao.Name = "txtValorCondicao"
        Me.txtValorCondicao.Size = New System.Drawing.Size(238, 22)
        Me.txtValorCondicao.TabIndex = 5
        '
        'btnImprimir
        '
        Me.btnImprimir.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnImprimir.BackgroundImage = CType(resources.GetObject("btnImprimir.BackgroundImage"), System.Drawing.Image)
        Me.btnImprimir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnImprimir.Location = New System.Drawing.Point(902, 10)
        Me.btnImprimir.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnImprimir.Name = "btnImprimir"
        Me.btnImprimir.Size = New System.Drawing.Size(43, 36)
        Me.btnImprimir.TabIndex = 6
        Me.btnImprimir.UseVisualStyleBackColor = False
        '
        'grbFiltro
        '
        Me.grbFiltro.Controls.Add(Me.txtValorCondicao)
        Me.grbFiltro.Controls.Add(Me.cbCondicao)
        Me.grbFiltro.Controls.Add(Me.cbCampo)
        Me.grbFiltro.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.grbFiltro.Location = New System.Drawing.Point(57, 5)
        Me.grbFiltro.Name = "grbFiltro"
        Me.grbFiltro.Size = New System.Drawing.Size(701, 44)
        Me.grbFiltro.TabIndex = 7
        Me.grbFiltro.TabStop = False
        '
        'btnIncluir
        '
        Me.btnIncluir.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnIncluir.BackgroundImage = CType(resources.GetObject("btnIncluir.BackgroundImage"), System.Drawing.Image)
        Me.btnIncluir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnIncluir.Location = New System.Drawing.Point(764, 10)
        Me.btnIncluir.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnIncluir.Name = "btnIncluir"
        Me.btnIncluir.Size = New System.Drawing.Size(42, 36)
        Me.btnIncluir.TabIndex = 8
        Me.btnIncluir.UseVisualStyleBackColor = False
        '
        'btnExcluir
        '
        Me.btnExcluir.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnExcluir.BackgroundImage = CType(resources.GetObject("btnExcluir.BackgroundImage"), System.Drawing.Image)
        Me.btnExcluir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnExcluir.Location = New System.Drawing.Point(806, 10)
        Me.btnExcluir.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnExcluir.Name = "btnExcluir"
        Me.btnExcluir.Size = New System.Drawing.Size(42, 36)
        Me.btnExcluir.TabIndex = 9
        Me.btnExcluir.UseVisualStyleBackColor = False
        '
        'btnAlterar
        '
        Me.btnAlterar.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnAlterar.BackgroundImage = CType(resources.GetObject("btnAlterar.BackgroundImage"), System.Drawing.Image)
        Me.btnAlterar.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAlterar.Location = New System.Drawing.Point(849, 10)
        Me.btnAlterar.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnAlterar.Name = "btnAlterar"
        Me.btnAlterar.Size = New System.Drawing.Size(42, 36)
        Me.btnAlterar.TabIndex = 10
        Me.btnAlterar.UseVisualStyleBackColor = False
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnClear.BackgroundImage = CType(resources.GetObject("btnClear.BackgroundImage"), System.Drawing.Image)
        Me.btnClear.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnClear.Location = New System.Drawing.Point(717, 14)
        Me.btnClear.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(36, 34)
        Me.btnClear.TabIndex = 11
        Me.btnClear.UseVisualStyleBackColor = False
        '
        'btnAnterior
        '
        Me.btnAnterior.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnAnterior.BackgroundImage = CType(resources.GetObject("btnAnterior.BackgroundImage"), System.Drawing.Image)
        Me.btnAnterior.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnAnterior.Location = New System.Drawing.Point(697, 305)
        Me.btnAnterior.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnAnterior.Name = "btnAnterior"
        Me.btnAnterior.Size = New System.Drawing.Size(60, 34)
        Me.btnAnterior.TabIndex = 12
        Me.btnAnterior.UseVisualStyleBackColor = False
        '
        'btnProximo
        '
        Me.btnProximo.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnProximo.BackgroundImage = CType(resources.GetObject("btnProximo.BackgroundImage"), System.Drawing.Image)
        Me.btnProximo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnProximo.Location = New System.Drawing.Point(885, 305)
        Me.btnProximo.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.btnProximo.Name = "btnProximo"
        Me.btnProximo.Size = New System.Drawing.Size(60, 34)
        Me.btnProximo.TabIndex = 13
        Me.btnProximo.UseVisualStyleBackColor = False
        '
        'lblRegistros
        '
        Me.lblRegistros.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.lblRegistros.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRegistros.Location = New System.Drawing.Point(759, 308)
        Me.lblRegistros.Name = "lblRegistros"
        Me.lblRegistros.Size = New System.Drawing.Size(126, 29)
        Me.lblRegistros.TabIndex = 14
        Me.lblRegistros.Text = "1-50"
        Me.lblRegistros.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'timerRefresh
        '
        '
        'frmBrowse
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(958, 346)
        Me.Controls.Add(Me.lblRegistros)
        Me.Controls.Add(Me.btnProximo)
        Me.Controls.Add(Me.btnAnterior)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.btnAlterar)
        Me.Controls.Add(Me.btnExcluir)
        Me.Controls.Add(Me.btnIncluir)
        Me.Controls.Add(Me.btnImprimir)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnSalvar)
        Me.Controls.Add(Me.ListView_Browse)
        Me.Controls.Add(Me.grbFiltro)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(3, 2, 3, 2)
        Me.Name = "frmBrowse"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Browse de tabelas"
        Me.grbFiltro.ResumeLayout(False)
        Me.grbFiltro.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnSalvar As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cbCampo As System.Windows.Forms.ComboBox
    Friend WithEvents cbCondicao As System.Windows.Forms.ComboBox
    Friend WithEvents txtValorCondicao As System.Windows.Forms.TextBox
    Friend WithEvents btnImprimir As System.Windows.Forms.Button
    Public WithEvents ListView_Browse As System.Windows.Forms.ListView
    Friend WithEvents grbFiltro As System.Windows.Forms.GroupBox
    Friend WithEvents btnIncluir As System.Windows.Forms.Button
    Friend WithEvents btnExcluir As System.Windows.Forms.Button
    Friend WithEvents btnAlterar As System.Windows.Forms.Button
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnAnterior As System.Windows.Forms.Button
    Friend WithEvents btnProximo As System.Windows.Forms.Button
    Friend WithEvents lblRegistros As System.Windows.Forms.Label
    Friend WithEvents timerRefresh As System.Windows.Forms.Timer
End Class
