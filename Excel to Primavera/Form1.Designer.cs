namespace Excel_to_Primavera
{
	partial class Form1
	{
		/// <summary>
		/// Variável de designer necessária.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Limpar os recursos que estão sendo usados.
		/// </summary>
		/// <param name="disposing">true se for necessário descartar os recursos gerenciados; caso contrário, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Código gerado pelo Windows Form Designer

		/// <summary>
		/// Método necessário para suporte ao Designer - não modifique 
		/// o conteúdo deste método com o editor de código.
		/// </summary>
		private void InitializeComponent()
		{
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.radioArtigos = new System.Windows.Forms.RadioButton();
			this.radioFornecedores = new System.Windows.Forms.RadioButton();
			this.radioClientes = new System.Windows.Forms.RadioButton();
			this.label1 = new System.Windows.Forms.Label();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.tabela = new System.Windows.Forms.DataGridView();
			this.cbofolha = new System.Windows.Forms.ComboBox();
			this.nomeficheiro = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.btnescolher = new System.Windows.Forms.Button();
			this.btnexportar = new System.Windows.Forms.Button();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.tabela)).BeginInit();
			this.SuspendLayout();
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.radioArtigos);
			this.groupBox1.Controls.Add(this.radioFornecedores);
			this.groupBox1.Controls.Add(this.radioClientes);
			this.groupBox1.Location = new System.Drawing.Point(15, 36);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(723, 61);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Exportação de Dados:";
			// 
			// radioArtigos
			// 
			this.radioArtigos.AutoSize = true;
			this.radioArtigos.Location = new System.Drawing.Point(213, 25);
			this.radioArtigos.Name = "radioArtigos";
			this.radioArtigos.Size = new System.Drawing.Size(57, 17);
			this.radioArtigos.TabIndex = 2;
			this.radioArtigos.Text = "Artigos";
			this.radioArtigos.UseVisualStyleBackColor = true;
			// 
			// radioFornecedores
			// 
			this.radioFornecedores.AutoSize = true;
			this.radioFornecedores.Location = new System.Drawing.Point(108, 25);
			this.radioFornecedores.Name = "radioFornecedores";
			this.radioFornecedores.Size = new System.Drawing.Size(90, 17);
			this.radioFornecedores.TabIndex = 1;
			this.radioFornecedores.Text = "Fornecedores";
			this.radioFornecedores.UseVisualStyleBackColor = true;
			// 
			// radioClientes
			// 
			this.radioClientes.AutoSize = true;
			this.radioClientes.Checked = true;
			this.radioClientes.Location = new System.Drawing.Point(20, 25);
			this.radioClientes.Name = "radioClientes";
			this.radioClientes.Size = new System.Drawing.Size(62, 17);
			this.radioClientes.TabIndex = 0;
			this.radioClientes.TabStop = true;
			this.radioClientes.Text = "Clientes";
			this.radioClientes.UseVisualStyleBackColor = true;
			this.radioClientes.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(309, 14);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(95, 13);
			this.label1.TabIndex = 1;
			this.label1.Text = "Excel to Primavera";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.tabela);
			this.groupBox2.Controls.Add(this.cbofolha);
			this.groupBox2.Controls.Add(this.nomeficheiro);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.btnescolher);
			this.groupBox2.Location = new System.Drawing.Point(14, 103);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(724, 422);
			this.groupBox2.TabIndex = 2;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Dados:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(12, 99);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(87, 13);
			this.label4.TabIndex = 7;
			this.label4.Text = "Pré-visualização:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(20, 63);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(106, 13);
			this.label3.TabIndex = 6;
			this.label3.Text = "Escolher folha Excel:";
			// 
			// tabela
			// 
			this.tabela.AllowUserToAddRows = false;
			this.tabela.AllowUserToDeleteRows = false;
			this.tabela.AllowUserToResizeColumns = false;
			this.tabela.AllowUserToResizeRows = false;
			this.tabela.BackgroundColor = System.Drawing.SystemColors.ScrollBar;
			this.tabela.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.tabela.Location = new System.Drawing.Point(9, 118);
			this.tabela.Name = "tabela";
			this.tabela.ReadOnly = true;
			this.tabela.RowHeadersVisible = false;
			this.tabela.Size = new System.Drawing.Size(705, 293);
			this.tabela.TabIndex = 5;
			// 
			// cbofolha
			// 
			this.cbofolha.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cbofolha.FormattingEnabled = true;
			this.cbofolha.Location = new System.Drawing.Point(134, 60);
			this.cbofolha.Name = "cbofolha";
			this.cbofolha.Size = new System.Drawing.Size(121, 21);
			this.cbofolha.TabIndex = 4;
			this.cbofolha.SelectedIndexChanged += new System.EventHandler(this.cbofolha_SelectedIndexChanged);
			// 
			// nomeficheiro
			// 
			this.nomeficheiro.Location = new System.Drawing.Point(157, 24);
			this.nomeficheiro.Name = "nomeficheiro";
			this.nomeficheiro.ReadOnly = true;
			this.nomeficheiro.Size = new System.Drawing.Size(381, 20);
			this.nomeficheiro.TabIndex = 3;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(20, 27);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(138, 13);
			this.label2.TabIndex = 1;
			this.label2.Text = "Ficheiro Excel (.xls ou .xlsx):";
			// 
			// btnescolher
			// 
			this.btnescolher.Location = new System.Drawing.Point(543, 24);
			this.btnescolher.Name = "btnescolher";
			this.btnescolher.Size = new System.Drawing.Size(62, 21);
			this.btnescolher.TabIndex = 0;
			this.btnescolher.Text = "Escolher";
			this.btnescolher.UseVisualStyleBackColor = true;
			this.btnescolher.Click += new System.EventHandler(this.btnescolher_Click);
			// 
			// btnexportar
			// 
			this.btnexportar.Location = new System.Drawing.Point(294, 534);
			this.btnexportar.Name = "btnexportar";
			this.btnexportar.Size = new System.Drawing.Size(153, 23);
			this.btnexportar.TabIndex = 3;
			this.btnexportar.Text = "Exportar";
			this.btnexportar.UseVisualStyleBackColor = true;
			this.btnexportar.Click += new System.EventHandler(this.btnexportar_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(750, 570);
			this.Controls.Add(this.btnexportar);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.groupBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Name = "Form1";
			this.Text = "Excel to Primavera";
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.groupBox2.ResumeLayout(false);
			this.groupBox2.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.tabela)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton radioFornecedores;
		private System.Windows.Forms.RadioButton radioClientes;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button btnescolher;
		private System.Windows.Forms.Button btnexportar;
		private System.Windows.Forms.TextBox nomeficheiro;
		private System.Windows.Forms.ComboBox cbofolha;
		private System.Windows.Forms.DataGridView tabela;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.RadioButton radioArtigos;
		private System.Windows.Forms.Label label4;
	}
}

