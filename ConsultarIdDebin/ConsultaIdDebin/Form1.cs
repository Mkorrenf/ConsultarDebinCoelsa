using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Configuration;
using System.Threading;

namespace ConsultaIdDebin
{
	public partial class Form1 : Form
	{
		public string debinId { get; set; }
		public System.Data.DataTable dt { get; set; }

		Debin debin;

		DataColumn column;
		DataRow row;
		public class Respuesta
		{
			public string codigo { get; set; }
			public string descripcion { get; set; }
		}

		public class Evaluacion
		{
			public string puntaje { get; set; }
		}

		public class Cuenta
		{
			public string banco { get; set; }
			public string sucursal { get; set; }
			public string cbu { get; set; }
			public string esTitular { get; set; }
			public string moneda { get; set; }
			public string tipo { get; set; }
			public string endpointId { get; set; }
		}

		public class EstadoComprador
		{
			public string codigo { get; set; }
			public string descripcion { get; set; }
		}

		public class Comprador
		{
			public string titular { get; set; }
			public string cuit { get; set; }
			public Cuenta cuenta { get; set; }
			public EstadoComprador estadoComprador { get; set; }
		}

		public class Detalle
		{
			public string concepto { get; set; }
			public string idUsuario { get; set; }
			public string idComprobante { get; set; }
			public string moneda { get; set; }
			public string importe { get; set; }
			public string devolucion { get; set; }
			public string importeComision { get; set; }
			public string comision { get; set; }
			public string descripcion { get; set; }
			public string idOperacionOriginal { get; set; }
			public string fecha { get; set; }
			public string fechaExpiracion { get; set; }
		}

		public class Vendedor
		{
			public string titular { get; set; }
			public string cuit { get; set; }
			public Cuenta cuenta { get; set; }
		}

		public class Estado
		{
			public string codigo { get; set; }
			public string descripcion { get; set; }
		}

		public class Operacion
		{
			public string id { get; set; }
			public Comprador comprador { get; set; }
			public Detalle detalle { get; set; }
			public Vendedor vendedor { get; set; }
			public Estado estado { get; set; }
			public string garantiaOk { get; set; }
			public string tipo { get; set; }
			public string loteId { get; set; }
			public string fechaNegocio { get; set; }
		}

		public class Debin
		{
			public Respuesta respuesta { get; set; }
			public Evaluacion evaluacion { get; set; }
			public string preautorizado { get; set; }
			public Operacion operacion { get; set; }

		}
		public Form1()
		{
			InitializeComponent();
		}

		private void button2_Click(object sender, EventArgs e)
		{
            if (txtDebin.Text != "")
            {
				ConsultarDebinAsync(txtDebin.Text);
			}
            else
            {
				MessageBox.Show("Por favor ingresar Id Debin para consultar");
            }
		}

		async Task ConsultarDebinAsync(string pdebin) 
		{
            try
            {
				debinId = pdebin;

				var client = new HttpClient();

				client.BaseAddress = new Uri(ConfigurationManager.AppSettings.Get("baseAddress"));
				client.DefaultRequestHeaders.Add("User-Agent", "C# console program");
				client.DefaultRequestHeaders.Accept.Add(
						new MediaTypeWithQualityHeaderValue("application/json"));

				var url = ConfigurationManager.AppSettings.Get("urlConsultaDebin") + debinId;
				HttpResponseMessage response = await client.GetAsync(url);

				await Task.Delay(250);

				response.EnsureSuccessStatusCode();
				var resp = await response.Content.ReadAsStringAsync();

				debin = new Debin();
				debin = JsonConvert.DeserializeObject<Debin>(resp);

                if (debin.operacion.id is null)
                {
					AgregarRowNoEncontrado(pdebin);
				}
                else
                {
					AgregarRow();
				}

					dgvDebin.DataSource = dt;
				
			}
            catch (Exception Ex)
            {
				MessageBox.Show(Ex.Message);
            }			
		}

        private void button1_Click(object sender, EventArgs e)
        {
			this.Close();
        }

		private void ExportDataTableToExcel(System.Data.DataTable table, string Xlfile)
		{

			Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
			Workbook book = excel.Application.Workbooks.Add(Type.Missing);
			excel.Visible = false;
			excel.DisplayAlerts = false;
			Worksheet excelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)book.ActiveSheet;
			excelWorkSheet.Name = table.TableName;
			excelWorkSheet.Cells.NumberFormat = "@";

			progressBar1.Maximum = table.Columns.Count;
			for (int i = 1; i < table.Columns.Count + 1; i++) // Creating Header Column In Excel
			{
				excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;
				if (progressBar1.Value < progressBar1.Maximum)
				{
					progressBar1.Value++;
					int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
					progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
					System.Windows.Forms.Application.DoEvents();
				}
			}


			progressBar1.Maximum = table.Rows.Count;
			for (int j = 0; j < table.Rows.Count; j++) // Exporting Rows in Excel
			{
				for (int k = 0; k < table.Columns.Count; k++)
				{
					excelWorkSheet.Cells[j + 2, k + 1] = table.Rows[j].ItemArray[k].ToString();
				}

				if (progressBar1.Value < progressBar1.Maximum)
				{
					progressBar1.Value++;
					int percent = (int)(((double)progressBar1.Value / (double)progressBar1.Maximum) * 100);
					progressBar1.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar1.Width / 2 - 10, progressBar1.Height / 2 - 7));
					System.Windows.Forms.Application.DoEvents();
				}
			}


			book.SaveAs(Xlfile);
			book.Close(true);
			excel.Quit();

			Marshal.ReleaseComObject(book);
			Marshal.ReleaseComObject(book);
			Marshal.ReleaseComObject(excel);

		}

        private void txtXmlFilePath_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnBrowseFolder_Click(object sender, EventArgs e)
        {
			DialogResult drResult = SFD.ShowDialog();
			if (drResult == System.Windows.Forms.DialogResult.OK)
				txtXmlFilePath.Text = SFD.FileName;
		}

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
				progressBar1.Value = 0;
				ExportDataTableToExcel(dt, txtXmlFilePath.Text);

				MessageBox.Show("Export Completed!");
			}
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dgvDebin.AllowUserToAddRows = false;
            dgvDebin.AllowUserToDeleteRows = false;
            dgvDebin.EditMode = DataGridViewEditMode.EditProgrammatically;
            dgvDebin.MultiSelect = true;

			txtXmlFilePath.Text = ConfigurationManager.AppSettings.Get("rutaDefectoGuardarArchivo");
			txtLoadFilePath.Text = ConfigurationManager.AppSettings.Get("rutaDefectoCardarArchivo");

			GenerarDataTable();
		}

		public void GenerarDataTable()
        {
			dt = new System.Data.DataTable();
			dt.TableName = "ExportDebin";

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "RESPUESTA";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "ID";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "COMPRADOR";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "CUIT_COMPRADOR";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "CBU_COMPRADOR";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "IMPORTE";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "ID_OPERACION_ORIGINAL";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "FECHA";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "VENDEDOR";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "CUIT_VENDEDOR";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "CBU_VENDEDOR";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "ESTADO";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "FECHA_NEGOCIO";
			dt.Columns.Add(column);

			column = new DataColumn();
			column.DataType = System.Type.GetType("System.String");
			column.ColumnName = "TIPO";
			dt.Columns.Add(column);
		}

		public void AgregarRow()
        {
			DataRow row = dt.NewRow();
			row["RESPUESTA"] = debin.respuesta.descripcion.ToString();
			row["ID"] = debin.operacion.id.ToString();
			row["COMPRADOR"] = debin.operacion.comprador.titular.ToString();
			row["CUIT_COMPRADOR"] = debin.operacion.comprador.cuit.ToString();
			row["CBU_COMPRADOR"] = debin.operacion.comprador.cuenta.cbu.ToString();
			row["IMPORTE"] = debin.operacion.detalle.importe.ToString();
			row["ID_OPERACION_ORIGINAL"] = debin.operacion.detalle.idOperacionOriginal.ToString();
			row["FECHA"] = debin.operacion.detalle.fecha.ToString();
			row["VENDEDOR"] = debin.operacion.vendedor.titular.ToString();
			row["CUIT_VENDEDOR"] = debin.operacion.vendedor.cuit.ToString();
			row["CBU_VENDEDOR"] = debin.operacion.vendedor.cuenta.cbu.ToString();
			row["ESTADO"] = debin.operacion.estado.codigo.ToString();
			row["FECHA_NEGOCIO"] = debin.operacion.fechaNegocio.ToString();
			row["TIPO"] = debin.operacion.tipo.ToString();

			dt.Rows.Add(row);
		}

		public void AgregarRowNoEncontrado(string debinIngresado)
		{
			DataRow row = dt.NewRow();
			row["RESPUESTA"] = "DEBIN NO ENCONTRADO";
			row["ID"] = debinIngresado;
			row["COMPRADOR"] = "DEBIN NO ENCONTRADO";
			row["CUIT_COMPRADOR"] = "DEBIN NO ENCONTRADO";
			row["CBU_COMPRADOR"] = "DEBIN NO ENCONTRADO";
			row["IMPORTE"] = "DEBIN NO ENCONTRADO";
			row["ID_OPERACION_ORIGINAL"] = "DEBIN NO ENCONTRADO";
			row["FECHA"] = "DEBIN NO ENCONTRADO";
			row["VENDEDOR"] = "DEBIN NO ENCONTRADO";
			row["CUIT_VENDEDOR"] = "DEBIN NO ENCONTRADO";
			row["CBU_VENDEDOR"] = "DEBIN NO ENCONTRADO";
			row["ESTADO"] = "DEBIN NO ENCONTRADO";
			row["FECHA_NEGOCIO"] = "DEBIN NO ENCONTRADO";
			row["TIPO"] = "DEBIN NO ENCONTRADO";

			dt.Rows.Add(row);
		}

		private void button3_Click(object sender, EventArgs e)
        {
			GenerarDataTable();
			dgvDebin.DataSource = null;
        }

        private void button4_Click(object sender, EventArgs e)
        {
			DialogResult drResult = OFD.ShowDialog();
			if (drResult == System.Windows.Forms.DialogResult.OK)
				txtLoadFilePath.Text = OFD.FileName;
		}

        private void btnConsulta_Click(object sender, EventArgs e)
        {
			if (txtLoadFilePath.Text != "")
			{
				try
				{
					progressBar2.Value = 0;

					string FileToRead = txtLoadFilePath.Text;
					int totalLineas = File.ReadAllLines(FileToRead).Count();

					using (StreamReader ReaderObject = new StreamReader(FileToRead))
					{

						progressBar2.Maximum = totalLineas;

						string Line;

						while ((Line = ReaderObject.ReadLine()) != null)
						{
							ConsultarDebinAsync(Line);
							
							if (progressBar2.Value < progressBar2.Maximum)
							{
								progressBar2.Value++;
								int percent = (int)(((double)progressBar2.Value / (double)progressBar2.Maximum) * 100);
								progressBar2.CreateGraphics().DrawString(percent.ToString() + "%", new System.Drawing.Font("Arial", (float)8.25, FontStyle.Regular), Brushes.Black, new PointF(progressBar2.Width / 2 - 10, progressBar2.Height / 2 - 7));
								System.Windows.Forms.Application.DoEvents();

								Thread.Sleep(1000);
							}
						}
					}
					
					MessageBox.Show("Consulta Completada!");
				}
				catch (Exception Ex)
				{
					MessageBox.Show(Ex.Message);
				}
			}
			else
			{
				MessageBox.Show("Por favor seleccione un archivo!");
			}
		}
    }
}
