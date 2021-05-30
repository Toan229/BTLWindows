using Data;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BTLWin
{
    public partial class TrangChu_Form : Form
    {
        private string MaGV;
        int accountType;
        public TrangChu_Form()
        {
            InitializeComponent();
        }

        public TrangChu_Form(string ID, int accountType)
        {
            MaGV = ID;
            this.accountType = accountType;
            InitializeComponent();
        }

        private void TrangChu_Form_Load(object sender, EventArgs e)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt = new Database().SelectData("SELECT TenMonHoc FROM MONHOC");
            if (dt.Rows.Count != 0)
            {
                for (int i = 0; i < dt.Rows.Count; i++)
                {
                    comboBox1.Items.Add(dt.Rows[i][0].ToString());
                }
                comboBox1.SelectedIndex = 0;
                thongKe(comboBox1.Text);
            }
        }

        public void thongKe(String TenMH)
        {
            dataGridView1.DataSource = new Database().SelectData("SELECT * FROM Fn_ThongKe (N'" + TenMH + "')");
            if (dataGridView1.DataSource != null)
            {
                dataGridView1.Columns[0].HeaderText = "Mã lớp";
                dataGridView1.Columns[1].HeaderText = "Số lượng sinh viên";
                dataGridView1.Columns[2].HeaderText = "0 - 3.9 (F)";
                dataGridView1.Columns[3].HeaderText = "4 - 4.6 (D)";
                dataGridView1.Columns[4].HeaderText = "4.6 - 5.4 (D+)";
                dataGridView1.Columns[5].HeaderText = "5.5 - 6.1 (C)";
                dataGridView1.Columns[6].HeaderText = "6.2 - 6.9 (C+)";
                dataGridView1.Columns[7].HeaderText = "7 - 7.6 (B)";
                dataGridView1.Columns[8].HeaderText = "7.7 - 8.4 (B+)";
                dataGridView1.Columns[9].HeaderText = "8.5 - 10 (A)";
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            thongKe(comboBox1.Text);
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1)
            {
                sapXep(comboBox2.SelectedIndex);
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1)
            {
                sapXep(comboBox2.SelectedIndex);
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex != -1)
            {
                sapXep(comboBox2.SelectedIndex);
            }
        }

        private void sapXep(int index)
        {
            if (radioButton1.Checked)
                dataGridView1.Sort(dataGridView1.Columns[index], ListSortDirection.Ascending);
            else
                dataGridView1.Sort(dataGridView1.Columns[index], ListSortDirection.Descending);
        }

        private void btnXuatTK_Click(object sender, EventArgs e)
        {
           
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel File(*.xlsx)|*.xlsx";
            DialogResult dialogResult =  saveFile.ShowDialog();
            if (dialogResult != DialogResult.Cancel)
            {


                string fileName = saveFile.FileName;
                object missingValue = System.Reflection.Missing.Value;
                this.Cursor = Cursors.WaitCursor;
                Microsoft.Office.Interop.Excel._Application application = new Microsoft.Office.Interop.Excel.Application();

                //Khoi tao WorkBook
                Microsoft.Office.Interop.Excel._Workbook workbook = application.Workbooks.Add(Type.Missing);

                //Khoi tao WorkSheet

                Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.ActiveSheet;
                try
                {

                    worksheet.Cells[1, 1] = "BẢNG ĐIỂM THỐNG KÊ ĐIỂM THEO MÔN ";
                    worksheet.Cells[3, 2] = "Môn học:  " + comboBox1.Text;


                    worksheet.Cells[9, 1] = "STT";
                    worksheet.Cells[9, 2] = "Mã lớp";
                    worksheet.Cells[9, 3] = "Số lượng sinh viên";
                    worksheet.Cells[9, 4] = "Điểm F";
                    worksheet.Cells[9, 5] = "Điểm D";
                    worksheet.Cells[9, 6] = "Điểm D+";
                    worksheet.Cells[9, 7] = "Điểm C";
                    worksheet.Cells[9, 8] = "Điểm C+";
                    worksheet.Cells[9, 9] = "Điểm B";
                    worksheet.Cells[9, 10] = "Điểm B+";
                    worksheet.Cells[9, 11] = "Điểm A";

                    for (int i = 0; i < dataGridView1.RowCount; i++)
                    {
                        worksheet.Cells[i + 10, 1] = i + 1;

                        for (int j = 0; j < 10; j++)
                        {
                            worksheet.Cells[i + 10, j + 2] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                        }

                    }


                    int diem = dataGridView1.RowCount;
                    worksheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                    worksheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA3;
                    worksheet.PageSetup.LeftMargin = 0;
                    worksheet.PageSetup.RightMargin = 0;
                    worksheet.PageSetup.TopMargin = 0;
                    worksheet.PageSetup.BottomMargin = 0;
                    worksheet.Range["A1"].ColumnWidth = 5;
                    worksheet.Range["B1"].ColumnWidth = 15;
                    worksheet.Range["C1"].ColumnWidth = 15;
                    worksheet.Range["D1"].ColumnWidth = 15;
                    worksheet.Range["E1"].ColumnWidth = 15;
                    worksheet.Range["F1"].ColumnWidth = 15;
                    worksheet.Range["G1"].ColumnWidth = 15;
                    worksheet.Range["H1"].ColumnWidth = 15;
                    worksheet.Range["I1"].ColumnWidth = 15;
                    worksheet.Range["J1"].ColumnWidth = 15;
                    worksheet.Range["K1"].ColumnWidth = 15;




                    worksheet.Range["A1", "K100"].Font.Name = "Times New Roman";
                    worksheet.Range["A1", "K1"].MergeCells = true;
                    worksheet.Range["A1", "K1"].Font.Bold = true;
                    worksheet.Range["A9", "K" + (diem + 10)].Borders.LineStyle = 1;

                    worksheet.Range["A1", "K1"].HorizontalAlignment = 3;
                    worksheet.Range["A9", "K9"].HorizontalAlignment = 3;
                    worksheet.Range["A10", "A" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["B10", "B" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["C10", "C" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["D10", "D" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["E10", "E" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["F10", "F" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["G10", "G" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["H10", "H" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["I10", "I" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["J10", "J" + (diem + 10)].HorizontalAlignment = 3;
                    worksheet.Range["K10", "K" + (diem + 10)].HorizontalAlignment = 3;

                    workbook.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, missingValue, missingValue, true, false, XlSaveAsAccessMode.xlNoChange,
                XlSaveConflictResolution.xlLocalSessionChanges, missingValue, missingValue);

                    workbook.Close(true, missingValue, missingValue);
                    application.Quit();
                    this.Cursor = Cursors.Default;

                    DialogResult result = MessageBox.Show("Bảng thống kê đã được lưu vào file.\nBạn có muốn mở file hay không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.Yes)
                    {
                        System.Diagnostics.Process.Start("Excel.exe", fileName);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex.Message + "\n Không thể lưu file!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {

                    //Đối tượng COM là đối tượng không được quản lý, chúng sẽ không được giải phóng bộ nhớ khi kết thúc, 
                    //vì vậy phải giải phóng chúng
                    QuanLyDiem ql = new QuanLyDiem();
                    ql.releaseObject(worksheet);
                    ql.releaseObject(workbook);
                    ql.releaseObject(application);
                }
            }
        }
    }
}

