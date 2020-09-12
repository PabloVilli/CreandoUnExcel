using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Sql;
using System.Data.SqlClient;
using excel = Microsoft.Office.Interop.Excel;

namespace LecturaSql
{
    public partial class Form1 : Form
    {
        excel.Application aplicacion;
        excel.Workbook libroTrabajo;
        excel.Worksheet hojaTrabajo;
        excel.Range rango;

        public Form1()
        {
            InitializeComponent();
        }

        public void LeerDB()
        {
            try
            {
                using (SqlConnection cn = new SqlConnection("Data Source=DSA6-PC\\COMPAC;Initial Catalog=ArconNet;Persist Security Info=True;User ID=sa;Password=DSA"))
                {
                    cn.Open();
                    string select = "SELECT dbo.Clientes.clId, dbo.Clientes.clCodigo, dbo.Clientes.clRazonSocial, dbo.Clientes.clRFC, dbo.Clientes.clCURP, dbo.Clientes.clTel1, dbo.Clientes.clEmail,  dbo.Clientes.clTel2, dbo.Clientes.rfId, dbo.Clientes.clClausula, dbo.Clientes.clFecha, dbo.Clientes.clStatus, dbo.Domicilios.doTipo, dbo.Domicilios.doCalle, dbo.Domicilios.doNoExt, dbo.Domicilios.doNoInt, dbo.Domicilios.doColonia, dbo.Domicilios.doMunicipio, dbo.Domicilios.doEstado, dbo.Domicilios.doPais, dbo.Domicilios.doCP, dbo.Domicilios.doCiudad FROM dbo.Domicilios INNER JOIN dbo.Localidades ON dbo.Domicilios.loId = dbo.Localidades.loId INNER JOIN dbo.Clientes ON dbo.Domicilios.IdRelacionado = dbo.Clientes.clId AND dbo.Domicilios.doTipo = 1";
                    SqlDataAdapter adaptador = new SqlDataAdapter(select, cn);
                    SqlCommandBuilder commandBuilder = new SqlCommandBuilder(adaptador);
                    DataSet ds = new DataSet();
                    adaptador.Fill(ds);
                    misDatosDataGridView1.ReadOnly = true;
                    misDatosDataGridView1.DataSource = ds.Tables[0];
                    MessageBox.Show("Importacion Finalizada");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void EstablecerCabecera()
        {
            try
            {
                //Asigno el titulo de documento excel.
                hojaTrabajo.Cells[1, 2] = "MI Primer excel con formato";

                //Asignos las cabeceras de cada linea.
                hojaTrabajo.Cells[3, 1] = "clId";
                hojaTrabajo.Cells[3, 2] = "clCodigo";
                hojaTrabajo.Cells[3, 3] = "clRazonSocial";
                hojaTrabajo.Cells[3, 4] = "clRFC";
                hojaTrabajo.Cells[3, 5] = "clCurp";
                hojaTrabajo.Cells[3, 6] = "clTel1";
                hojaTrabajo.Cells[3, 7] = "clEmail";
                hojaTrabajo.Cells[3, 8] = "clTel2";
                hojaTrabajo.Cells[3, 9] = "rfId";
                hojaTrabajo.Cells[3, 10] = "clClausula";
                hojaTrabajo.Cells[3, 11] = "clFecha";
                hojaTrabajo.Cells[3, 12] = "clFecha";
                hojaTrabajo.Cells[3, 13] = "doTipo";
                hojaTrabajo.Cells[3, 14] = "doCalle";
                hojaTrabajo.Cells[3, 15] = "doNoExt";
                hojaTrabajo.Cells[3, 16] = "doNoInt";
                hojaTrabajo.Cells[3, 17] = "doColonia";
                hojaTrabajo.Cells[3, 18] = "doMunicipio";
                hojaTrabajo.Cells[3, 19] = "doEstado";
                hojaTrabajo.Cells[3, 20] = "doPais";
                hojaTrabajo.Cells[3, 21] = "doCP";
                hojaTrabajo.Cells[3, 22] = "doCiudad";

                //Asigno el borde de las cabeceras.
                rango = hojaTrabajo.Range["A3", "V3"];
                rango.Borders.LineStyle = excel.XlLineStyle.xlContinuous;

                //Centro los textos
                rango = hojaTrabajo.Rows[3];
                rango.HorizontalAlignment = excel.XlHAlign.xlHAlignCenter;

                rango = hojaTrabajo.Columns[1];
                rango.ColumnWidth = 5;

                rango = hojaTrabajo.Columns[2];
                rango.ColumnWidth = 5;

                rango = hojaTrabajo.Columns[3];
                rango.ColumnWidth = 20;

                rango = hojaTrabajo.Columns[4];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[5];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[6];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[7];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[8];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[9];
                rango.ColumnWidth = 3;

                rango = hojaTrabajo.Columns[10];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[11];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[12];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[13];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[14];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[15];
                rango.ColumnWidth = 15;

                rango = hojaTrabajo.Columns[16];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[17];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[18];
                rango.ColumnWidth = 5;

                rango = hojaTrabajo.Columns[19];
                rango.ColumnWidth = 5;

                rango = hojaTrabajo.Columns[20];
                rango.ColumnWidth = 10;

                rango = hojaTrabajo.Columns[21];
                rango.ColumnWidth = 1;

                rango = hojaTrabajo.Columns[22];
                rango.ColumnWidth = 10;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExportarAExcel()
        {
            try
            {
                SaveFileDialog fichero = new SaveFileDialog();
                fichero.Filter = "Excel (*.xls)|*.xls";
                fichero.FileName = "Exportacion Excel";
                if (fichero.ShowDialog() == DialogResult.OK)
                {

                    aplicacion = new excel.Application();
                    libroTrabajo = aplicacion.Workbooks.Add(excel.XlWBATemplate.xlWBATWorksheet);
                    hojaTrabajo = (excel.Worksheet)libroTrabajo.Worksheets.get_Item(1);

                    EstablecerCabecera();

                    for (int i = 4; i < misDatosDataGridView1.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < misDatosDataGridView1.Columns.Count; j++)
                        {
                            if ((misDatosDataGridView1.Rows[i].Cells[j].Value == null) == false)
                            {
                                hojaTrabajo.Cells[i + 1, j + 1] = misDatosDataGridView1.Rows[i].Cells[j].Value.ToString();
                                
                            }
                        }
                    }
                    libroTrabajo.SaveAs(fichero.FileName, excel.XlFileFormat.xlWorkbookNormal);
                    libroTrabajo.Close(true);
                    aplicacion.Quit();
                    MessageBox.Show("Exportacion Finalizada");
                    Application.Exit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        

        //public void AvanceBar(int valMax)
        //{
        //    AvanceProgressBar.Maximum = valMax;

        //    if (AvanceProgressBar.Value != AvanceProgressBar.Maximum)
        //    {
        //        AvanceProgressBar.Value += 1;

        //        ResToolStripLabel.Text = "No de Reistros = " + AvanceProgressBar.Value.ToString();
        //    }
        //}

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            LeerDB();
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            ExportarAExcel();
        }
    }
}
