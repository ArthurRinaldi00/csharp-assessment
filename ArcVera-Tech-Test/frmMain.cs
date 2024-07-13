using Parquet.Schema;
using Parquet;
using System.Data;
using OxyPlot;
using OxyPlot.Series;
using OxyPlot.WindowsForms;
using DataColumn = System.Data.DataColumn;
using OxyPlot.Axes;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ArcVera_Tech_Test
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private async void btnImportEra5_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Parquet files (*.parquet)|*.parquet|All files (*.*)|*.*";
                openFileDialog.Title = "Select a Parquet File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    DataTable dataTable = await ReadParquetFileAsync(filePath);
                    dgImportedEra5.DataSource = dataTable;
                    PlotU10DailyValues(dataTable);
                }
            }
        }

        private async Task<DataTable> ReadParquetFileAsync(string filePath)
        {
            using (Stream fileStream = File.OpenRead(filePath))
            {
                using (var parquetReader = await ParquetReader.CreateAsync(fileStream))
                {
                    DataTable dataTable = new DataTable();

                    for (int i = 0; i < parquetReader.RowGroupCount; i++)
                    {
                        using (ParquetRowGroupReader groupReader = parquetReader.OpenRowGroupReader(i))
                        {
                            // Create columns
                            foreach (DataField field in parquetReader.Schema.GetDataFields())
                            {
                                if (!dataTable.Columns.Contains(field.Name))
                                {
                                    Type columnType = field.HasNulls ? typeof(object) : field.ClrType;
                                    dataTable.Columns.Add(field.Name, columnType);
                                }

                                // Read values from Parquet column
                                DataColumn column = dataTable.Columns[field.Name];
                                Array values = (await groupReader.ReadColumnAsync(field)).Data;
                                for (int j = 0; j < values.Length; j++)
                                {
                                    if (dataTable.Rows.Count <= j)
                                    {
                                        dataTable.Rows.Add(dataTable.NewRow());
                                    }
                                    dataTable.Rows[j][field.Name] = values.GetValue(j);
                                }
                            }
                        }
                    }

                    return dataTable;
                }
            }
        }

        private void PlotU10DailyValues(DataTable dataTable)
        {
            var plotModel = new PlotModel { Title = "Daily u10 Values" };
            var lineSeries = new LineSeries { Title = "u10" };

            var groupedData = dataTable.AsEnumerable()
                .GroupBy(row => DateTime.Parse(row["date"].ToString()))
                .Select(g => new
                {
                    Date = g.Key,
                    U10Average = g.Average(row => Convert.ToDouble(row["u10"]))
                })
                .OrderBy(data => data.Date);

            foreach (var data in groupedData)
            {
                lineSeries.Points.Add(new DataPoint(DateTimeAxis.ToDouble(data.Date), data.U10Average));
            }

            plotModel.Series.Add(lineSeries);
            plotView1.Model = plotModel;
        }

        private DataTable GetDataTableFromDataGridView(DataGridView dgv)
        {
            if (dgv.DataSource is DataTable dataTable)
            {
                return dataTable;
            }
            else
            {
                throw new InvalidOperationException("DataGrid Vazio.");
            }
        }

        private void btnExportCsv_Click(object sender, EventArgs e)
        {
            ExportDataGridViewToCSV();
        }

        private void ExportDataGridViewToCSV()
        {
            try
            {
                DataTable dt = GetDataTableFromDataGridView(dgImportedEra5);
                ExportToCSV.Export(dt, "C:\\Users\\arthur\\Downloads\\arquivo.csv", 50);
                MessageBox.Show("Csv Exportado com Sucesso");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro: " + ex.Message);
            }
        }

        public static class ExportToCSV
        {
            public static void Export(DataTable dt, string filePath, int maxRows = 50)
            {
                using (StreamWriter sw = new StreamWriter(filePath))
                {
                    var columnNames = dt.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
                    sw.WriteLine(string.Join(",", columnNames));

                    int rowsToExport = Math.Min(dt.Rows.Count, maxRows);

                    for (int i = 0; i < rowsToExport; i++)
                    {
                        var fields = dt.Rows[i].ItemArray.Select(field => field.ToString());
                        sw.WriteLine(string.Join(",", fields));
                    }
                }
            }
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            ExportDataGridViewToExcel();
        }

        private void ExportDataGridViewToExcel()
        {
            try
            {
                DataTable dt = GetDataTableFromDataGridView(dgImportedEra5);
                ExportToExcel.Export(dt, "C:\\Users\\arthur\\Downloads\\arquivo.xlsx");
                MessageBox.Show("Excel Exportado com Sucesso");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        public static class ExportToExcel
        {
            public static void Export(DataTable dt, string filePath)
            {
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = workbook.Sheets[1];
                worksheet = workbook.ActiveSheet;
                worksheet.Name = "January-era5";

                for (int i = 1; i < dt.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dt.Columns[i - 1].ColumnName;
                }

                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    for (int j = 0; j < dt.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dt.Rows[i][j].ToString();
                    }
                }

                workbook.SaveAs(filePath);
                workbook.Close();
                excelApp.Quit();

                ReleaseObject(worksheet);
                ReleaseObject(workbook);
                ReleaseObject(excelApp);
            }

            private static void ReleaseObject(object obj)
            {
                try
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                    obj = null;
                }
                catch (Exception ex)
                {
                    obj = null;
                    throw ex;
                }
                finally
                {
                    GC.Collect();
                }
            }
        }
    }
}
