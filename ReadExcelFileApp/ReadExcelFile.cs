using System;
using System.Data;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelFileApp
{
    public partial class ReadExcelFile : Form
    {
        public ReadExcelFile()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        DataTable ReadExcel(string path)
        {
            try
            {
                Excel.Application app = new Excel.Application();
                Excel.Workbook workbook = app.Workbooks.Open(path, true, true);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                DataTable dt = new DataTable();

                // read header
                int maxCol = 1;
                while (true)
                {
                    Excel.Range cell = worksheet.Cells[1, maxCol];
                    if (cell.Value2 == null)
                    {
                        break;
                    }

                    dt.Columns.Add(Convert.ToString(cell.Value2));
                    maxCol++;
                }


                // read body
                DataRow row;
                int r = 2;
                while (true)
                {
                    Excel.Range cell = worksheet.Cells[r, 1];
                    if (cell.Value2 == null)
                    {
                        break;
                    }
                    row = dt.NewRow();
                    row[0] = Convert.ToString(cell.Value2);

                    for (int i = 2; i < maxCol; ++i)
                    {
                        cell = worksheet.Cells[r, i];
                        row[i - 1] = Convert.ToString(cell.Value2);
                    }

                    dt.Rows.Add(row);
                    r++;
                }

                workbook.Close(false);
                app.Workbooks.Close();
                app.Quit();

                return dt;
            }
            catch (Exception)
            {
                return null;
            }
        }

        private void ChooseAndReadFileButton_Click(object sender, EventArgs e)
        {
            string path = string.Empty;
            string ext = string.Empty;
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    DataTable dt = ReadExcel(file.FileName);
                    dataGridView1.Visible = true;
                    dataGridView1.DataSource = dt;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }

        private void ReadExcelFile_Load(object sender, EventArgs e)
        {

        }
    }
}
