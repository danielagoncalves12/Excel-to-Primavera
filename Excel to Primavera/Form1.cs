using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using ExcelDataReader;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;


namespace Excel_to_Primavera
{

	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			
		}

		DataTableCollection tableCollection;

		private void btnescolher_Click(object sender, EventArgs e)
		{
			using (OpenFileDialog openfiledialog = new OpenFileDialog(){ Filter= "Excel Files|*.xls;*.xlsx;*.xlsm" })
			{
				if(openfiledialog.ShowDialog() == DialogResult.OK)
				{
					nomeficheiro.Text = openfiledialog.FileName;
					using (var stream = File.Open(openfiledialog.FileName, FileMode.Open, FileAccess.Read))
					{
						using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
						{
							DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
							{
								ConfigureDataTable = (_) => new ExcelDataTableConfiguration { UseHeaderRow = true }
							});

							tableCollection = result.Tables;
							cbofolha.Items.Clear();

							foreach (System.Data.DataTable table in tableCollection)
							{ 
							cbofolha.Items.Add(table.TableName);
						}
							


						}
					}
				}
			}
		}

		private void cbofolha_SelectedIndexChanged(object sender, EventArgs e)
		{
			
			System.Data.DataTable dt = tableCollection[cbofolha.SelectedItem.ToString()];
			tabela.DataSource = dt;

				}

		private void btnexportar_Click(object sender, EventArgs e)
		{

			//Se o radio Clientes tiver selecionado irá exportar para os clientes

			if (radioClientes.Checked)
			{

				DialogResult resposta = MessageBox.Show("O processo seguinte irá exportar a folha de Excel selecionada para a tabela Clientes. Deseja continuar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (resposta == DialogResult.Yes)
				{

					if (tabela.Rows.Count != 0)
					{
						SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-5N8QSIB\PRIMAVERA;Initial Catalog=PRIDEMO;Integrated Security=True;MultipleActiveResultSets=true");

						for (int i = 0; i < tabela.Rows.Count; i++)
						{

							SqlCommand query = new SqlCommand(@"SELECT * from Clientes Where Cliente LIKE '" + tabela.Rows[i].Cells[0].Value + "';", con);
							con.Open();
							query.ExecuteNonQuery();
							SqlDataReader reader = query.ExecuteReader();


							if (reader.Read())
							{

								con.Close();
							}
							else
							{
								SqlCommand cmd = new SqlCommand(@"INSERT INTO Clientes (Cliente, Nome, Fac_Mor, Fac_Local, Fac_Cp, Fac_Cploc, Fac_Tel, Fac_Fax, Desconto, TipoPrec, TipoCred, LimiteCred, NumContrib, Pais, TipoCli, AvisosVenc, ModoPag, CondPag, Moeda) VALUES ('" + tabela.Rows[i].Cells[0].Value + "','" + tabela.Rows[i].Cells[1].Value + "','" + tabela.Rows[i].Cells[2].Value + "','" + tabela.Rows[i].Cells[3].Value + "','" + tabela.Rows[i].Cells[4].Value + "','" + tabela.Rows[i].Cells[5].Value + "','" + tabela.Rows[i].Cells[6].Value + "','" + tabela.Rows[i].Cells[7].Value + "','" + tabela.Rows[i].Cells[8].Value + "','" + tabela.Rows[i].Cells[9].Value + "','" + tabela.Rows[i].Cells[10].Value + "','" + tabela.Rows[i].Cells[11].Value + "','" + tabela.Rows[i].Cells[12].Value + "','" + tabela.Rows[i].Cells[13].Value + "','" + tabela.Rows[i].Cells[14].Value + "','" + tabela.Rows[i].Cells[15].Value + "','" + tabela.Rows[i].Cells[16].Value + "','" + tabela.Rows[i].Cells[17].Value + "','" + tabela.Rows[i].Cells[18].Value + "')", con);


								cmd.ExecuteNonQuery();
								con.Close();
							}
						}
						MessageBox.Show("Dados exportados com sucesso. Chaves primárias duplicadas foram ignoradas.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
						nomeficheiro.Clear();
						cbofolha.Items.Clear();
						tabela.DataSource = null;

					}
					else
					{
						MessageBox.Show("É necessário escolher um ficheiro Excel e a respetiva folha para exportar.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}

				}

			}

			//Se o radio Fornecedores tiver selecionado irá exportar para os fornecedores


			if (radioFornecedores.Checked)
			{

				DialogResult resposta = MessageBox.Show("O processo seguinte irá exportar a folha de Excel selecionada para a tabela Fornecedores. Deseja continuar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (resposta == DialogResult.Yes)
				{

					if (tabela.Rows.Count != 0)
					{
						SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-5N8QSIB\PRIMAVERA;Initial Catalog=PRIDEMO;Integrated Security=True;MultipleActiveResultSets=true");

						for (int i = 0; i < tabela.Rows.Count; i++)
						{
							SqlCommand query = new SqlCommand(@"SELECT * from Fornecedores Where Fornecedor LIKE '" + tabela.Rows[i].Cells[0].Value + "';", con);
							con.Open();
							query.ExecuteNonQuery();
							SqlDataReader reader = query.ExecuteReader();

							if (reader.Read())
							{
								con.Close();
							}
							else
							{

								SqlCommand cmd = new SqlCommand(@"INSERT INTO Fornecedores (Fornecedor, Nome, Morada, Local, Cp, Cploc, Tel, Fax, NumContrib, Pais, TipoFor, CondPag, ModoPag, Moeda) VALUES ('" + tabela.Rows[i].Cells[0].Value + "','" + tabela.Rows[i].Cells[1].Value + "','" + tabela.Rows[i].Cells[2].Value + "','" + tabela.Rows[i].Cells[3].Value + "','" + tabela.Rows[i].Cells[4].Value + "','" + tabela.Rows[i].Cells[5].Value + "','" + tabela.Rows[i].Cells[6].Value + "','" + tabela.Rows[i].Cells[7].Value + "','" + tabela.Rows[i].Cells[8].Value + "','" + tabela.Rows[i].Cells[9].Value + "','" + tabela.Rows[i].Cells[10].Value + "','" + tabela.Rows[i].Cells[11].Value + "','" + tabela.Rows[i].Cells[12].Value + "','" + tabela.Rows[i].Cells[13].Value + "')", con);

								
								cmd.ExecuteNonQuery();
								con.Close();
							}
						}
						MessageBox.Show("Dados exportados com sucesso. Chaves primárias duplicadas foram ignoradas.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
						nomeficheiro.Clear();
						cbofolha.Items.Clear();
						tabela.DataSource = null;
					}
					else
					{
						MessageBox.Show("É necessário escolher um ficheiro Excel e a respetiva folha para exportar.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}

				}

			}

			//Se o radio Artigos tiver selecionado irá exportar para os artigos


			if (radioArtigos.Checked)
			{

				DialogResult resposta = MessageBox.Show("O processo seguinte irá exportar a folha de Excel selecionada para a tabela Artigos. Deseja continuar?", "Aviso", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (resposta == DialogResult.Yes)
				{

					if (tabela.Rows.Count != 0)
					{

						if (tabela.Rows.Count != 0)
						{
							SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-5N8QSIB\PRIMAVERA;Initial Catalog=PRIDEMO;Integrated Security=True;MultipleActiveResultSets=true");

							for (int i = 0; i < tabela.Rows.Count; i++)
							{
								SqlCommand query = new SqlCommand(@"SELECT * from artigo Where artigo LIKE '" + tabela.Rows[i].Cells[0].Value + "';", con);
								con.Open();
								query.ExecuteNonQuery();
								SqlDataReader reader = query.ExecuteReader();

								if (reader.Read())
								{
									con.Close();
								}
								else
								{

									SqlCommand cmd = new SqlCommand(@"INSERT INTO artigo (Artigo, Descricao, UnidadeBase, Iva, MovStock, Familia, ArmazemSugestao, TipoArtigo, SubFamilia, PermiteDevolucao) VALUES ('" + tabela.Rows[i].Cells[0].Value + "','" + tabela.Rows[i].Cells[1].Value + "','" + tabela.Rows[i].Cells[2].Value + "','" + tabela.Rows[i].Cells[3].Value + "','" + tabela.Rows[i].Cells[4].Value + "','" + tabela.Rows[i].Cells[5].Value + "','" + tabela.Rows[i].Cells[6].Value + "','" + tabela.Rows[i].Cells[7].Value + "','" + tabela.Rows[i].Cells[8].Value + "','" + tabela.Rows[i].Cells[9].Value + "')", con);
									
									cmd.ExecuteNonQuery();
									con.Close();
								}
							}
							MessageBox.Show("Dados exportados com sucesso. Chaves primárias duplicadas foram ignoradas.", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
							nomeficheiro.Clear();
							cbofolha.Items.Clear();
							tabela.DataSource = null;
						}
						else
						{
							MessageBox.Show("É necessário escolher um ficheiro Excel e a respetiva folha para exportar.", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}

					}

				}
			}
		}


		private void radioButton1_CheckedChanged(object sender, EventArgs e)
		{
		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{
		}
	}
}
