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
                Excel.Range range = worksheet.UsedRange;

                DataTable table = new DataTable();

                object[,] values = (object[,])range.get_Value(
                    Excel.XlRangeValueDataType.xlRangeValueDefault);
                int numRows = values.GetLength(0);
                int numCols = values.GetLength(1);

                for (int j = 1; j <= numCols; ++j)
                {
                    object value = values[1, j];
                    table.Columns.Add(Convert.ToString(value));
                }

                DataRow row;
                for (int i = 2; i <= numRows; ++i)
                {
                    row = table.NewRow();
                    for (int j = 1; j <= numCols; ++j)
                    {
                        object value = values[i, j];
                        row[j - 1] = value;
                    }
                    table.Rows.Add(row);
                }

                workbook.Close(false);
                app.Workbooks.Close();
                app.Quit();

                return table;
            }
            catch (Exception e)
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
