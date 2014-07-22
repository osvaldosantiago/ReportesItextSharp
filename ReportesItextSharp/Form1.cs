using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using MySql.Data;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Diagnostics;
using System.IO;
namespace ReportesItextSharp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
       
        private static string CadenaCon()
        {
            try
            {
                return "server=localhost; database=puntoventa; user id=root;password=PWD56XZ ";
            }
            catch (Exception ex)
            {
                return "";
            }
        }
        public void CargarGrid() {
            string cadena = CadenaCon();
            MySqlConnection Conexion = new MySqlConnection(cadena);
            DataSet ds = new DataSet();
            MySqlDataAdapter data = new MySqlDataAdapter("Select * from usuarios",Conexion);
            Conexion.Open();
            data.Fill(ds, "usuarios");
            Conexion.Close();
            dataGridView1.DataSource = ds.Tables["usuarios"].DefaultView;
            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToOrderColumns = false;
            dataGridView1.AllowUserToResizeColumns = false;
            dataGridView1.AllowUserToResizeRows = false;
            dataGridView1.AutoResizeColumns();
            dataGridView1.BorderStyle = BorderStyle.None;
        }
        #region crearPDF
        private void To_pdf()
        {
                    Document doc = new Document(PageSize.A4.Rotate(), 10, 10, 10, 10);
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.InitialDirectory = @"C:";
                    saveFileDialog1.Title = "Guardar Reporte";
                    saveFileDialog1.DefaultExt = "pdf";
                    saveFileDialog1.Filter = "pdf Files (*.pdf)|*.pdf| All Files (*.*)|*.*";
                    saveFileDialog1.FilterIndex = 2;
                    saveFileDialog1.RestoreDirectory = true;
                    string filename = "";
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        filename = saveFileDialog1.FileName;
                    }

                    if (filename.Trim() != "")
                    {
                        FileStream file = new FileStream(filename,
                        FileMode.OpenOrCreate,
                        FileAccess.ReadWrite,
                        FileShare.ReadWrite);
                        PdfWriter.GetInstance(doc, file);
                        doc.Open();
                        string remito = "Autorizo: OSVALDO SANTIAGO ESTRADA";
                        string envio = "Fecha:" + DateTime.Now.ToString();

                        Chunk chunk = new Chunk("Reporte de General Usuarios", FontFactory.GetFont("ARIAL", 20, iTextSharp.text.Font.BOLD));
                        doc.Add(new Paragraph(chunk));
                        doc.Add(new Paragraph("                       "));
                        doc.Add(new Paragraph("                       "));
                        doc.Add(new Paragraph("------------------------------------------------------------------------------------------"));
                        doc.Add(new Paragraph("Lagos de moreno Jalisco"));
                        doc.Add(new Paragraph(remito));
                        doc.Add(new Paragraph(envio));
                        doc.Add(new Paragraph("------------------------------------------------------------------------------------------"));
                        doc.Add(new Paragraph("                       "));
                        doc.Add(new Paragraph("                       "));
                        doc.Add(new Paragraph("                       "));
                        GenerarDocumento(doc);
                        doc.AddCreationDate();
                        doc.Add(new Paragraph("______________________________________________", FontFactory.GetFont("ARIAL", 20, iTextSharp.text.Font.BOLD)));
                        doc.Add(new Paragraph("Firma", FontFactory.GetFont("ARIAL", 20, iTextSharp.text.Font.BOLD)));
                        doc.Close();
                        Process.Start(filename);//Esta parte se puede omitir, si solo se desea guardar el archivo, y que este no se ejecute al instante
                    }
              
        }
        public void GenerarDocumento(Document document)
        {
            int i, j;
            PdfPTable datatable = new PdfPTable(dataGridView1.ColumnCount);
            datatable.DefaultCell.Padding = 3;
            float[] headerwidths = GetTamañoColumnas(dataGridView1);
            datatable.SetWidths(headerwidths);
            datatable.WidthPercentage = 100;
            datatable.DefaultCell.BorderWidth = 2;
            datatable.DefaultCell.HorizontalAlignment = Element.ALIGN_CENTER;
            for (i = 0; i < dataGridView1.ColumnCount; i++)
            {
                datatable.AddCell(dataGridView1.Columns[i].HeaderText);
            }
            datatable.HeaderRows = 1;
            datatable.DefaultCell.BorderWidth = 1;
            for (i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (dataGridView1[j, i].Value != null)
                    {
                        datatable.AddCell(new Phrase(dataGridView1[j, i].Value.ToString()));//En esta parte, se esta agregando un renglon por cada registro en el datagrid
                    }
                }
                datatable.CompleteRow();
            }
            document.Add(datatable);
        }
        public float[] GetTamañoColumnas(DataGridView dg)
        {
            float[] values = new float[dg.ColumnCount];
            for (int i = 0; i < dg.ColumnCount; i++)
            {
                values[i] = (float)dg.Columns[i].Width;
            }
            return values;

        }
        #endregion
        private void button2_Click(object sender, EventArgs e)
        {
            To_pdf();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CargarGrid();
        }
        /*
         *Eso es todo por el momento!! Cualquier duda, dejenme un comentario!!
         *Gracias
         */
    }
}
