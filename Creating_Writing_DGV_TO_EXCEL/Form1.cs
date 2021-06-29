using MySql.Data.MySqlClient;
using OfficeOpenXml;
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

namespace Creating_Writing_DGV_TO_EXCEL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            // first, Adding EPPPlus from nugat packages
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private DataTable dt = new DataTable();

        private async Task<DataTable> loadData()
        {
            MySqlConnection conn = new MySqlConnection("server=192.168.0.2;userid=root;password=boom123;");
            MySqlDataAdapter da = new MySqlDataAdapter(txtSQL.Text, conn);
            conn.Open();

            await Task.Run(() => da.Fill(dt));

            return dt;
        }

        private void btnExportExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Workbook|*.xlsx";
            try
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    var fileInfo = new FileInfo(sfd.FileName);
                    ExcelPackage excelPackage = new ExcelPackage(fileInfo);
                    ExcelWorksheet workSheet = excelPackage.Workbook.Worksheets.Add("sheet1");
                    workSheet.Cells.LoadFromDataTable(dt, true, OfficeOpenXml.Table.TableStyles.None);
                    excelPackage.Save();
                }
            }
            catch (Exception)
            {
                throw;
            }
            MessageBox.Show("เรียบร้อยครับ");
        }

        private async void btnRun_Click(object sender, EventArgs e)
        {
            if (dt.Rows.Count < 0 || dt == null)
            {
                MessageBox.Show("ตรวจสอบ sql string อีกครั้ง");
            }
            else
            {
                dgvData.DataSource = await loadData();
            }
        }
    }
}