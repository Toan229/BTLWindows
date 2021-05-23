using Data;
using Microsoft.Office.Interop.Excel;
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
using ExcelDataReader;

namespace BTLWin
{
    struct xoaDiem
    {
        public string MH, SV;
        public int index;
        public xoaDiem(int index, string MH, string SV)
        {
            this.index = index;
            this.MH = MH;
            this.SV = SV;
        }
    }
    public partial class QuanLyDiem : Form
    {
        string MaGV, MaMH, MaMHCu;
        bool Saved;

        List<DiemSV> diemUpdate, diemInsert;
        List<string> diemTrongCSDL;
        public QuanLyDiem()
        {
            InitializeComponent();
        }

        public QuanLyDiem(string username)
        {
            InitializeComponent();
            this.MaGV = username;
            Saved = true;
            diemUpdate = new List<DiemSV>();
            diemInsert = new List<DiemSV>();
            diemTrongCSDL = new List<string>();
        }

        private void QuanLyDiem_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = new Database().SelectData("EXEC TimKiemMonHoc_TheoMaGV '" + MaGV + "'");
            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + dataGridView1.Rows[0].Cells[0].Value + "'");
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                diemTrongCSDL.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
            }
            MaMH = dataGridView1.Rows[0].Cells[0].Value.ToString();
            MaMHCu = MaMH;
            lblTongSV.Text = dataGridView2.RowCount + " sinh viên";
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                MaMH = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                if (!Saved)
                {
                    if (!MaMHCu.Equals(MaMH))//Sau khi chỉnh sửa điểm các sinh viên ở môn học cũ thì chưa lưu
                    {
                        DialogResult result = MessageBox.Show("Thông tin bạn vừa cập nhập chưa được lưu. \nBạn có muốn lưu chúng không ?", "Warning",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (result == DialogResult.Yes)
                        {//Cập nhập thông tin

                        }
                        dataGridView2.ReadOnly = true;
                        dataGridView2.AllowUserToAddRows = false;
                        dataGridView2.AllowUserToDeleteRows = false;
                        Saved = true;
                    }
                }

                dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '"
                    + dataGridView1.Rows[e.RowIndex].Cells[0].Value + "'");
                diemTrongCSDL.Clear();
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    diemTrongCSDL.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
                }
                lblTongSV.Text = dataGridView2.RowCount.ToString() + " sinh viên";
                MaMHCu = MaMH;
            }
        }

        private void QuanLyDiem_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (!Saved)
            {
                DialogResult result = MessageBox.Show("Thông tin bạn vừa cập nhập chưa được lưu. \nBạn có muốn lưu chúng không ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {//Cập nhập thông tin
                    btnLuu_Click(null, null);
                }
            }
        }

        private void btnXuatFile_Click(object sender, EventArgs e)
        {
            if (!Saved)
            {
                DialogResult result = MessageBox.Show("Thông tin bạn vừa cập nhập chưa được lưu. \nBạn có muốn lưu chúng không ?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {//Cập nhập thông tin
                    btnLuu_Click(null, null);
                }
            }
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel File(*.xlsx)|*.xlsx";
            saveFile.ShowDialog();
            string fileName = saveFile.FileName;
            xuatFileExcel(fileName);
        }

        private void xuatFileExcel(string fileName)
        {
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();

            //Biểu diễn cho các giá trị bị thiếu
            //Use this instance of the Missing class to represent missing values, for example, 
            //when you invoke methods that have default parameter values. For a code example, see the Missing class.
            object missingValue = System.Reflection.Missing.Value;

            Workbook workbook = application.Workbooks.Add(missingValue);
            Worksheet worksheet = workbook.ActiveSheet;
            try
            {
                //Excel sẽ tự động định dạng kiểu cho dữ liệu, chuyển cột mã sinh viên từ kiểu số sang chuỗi
                worksheet.Columns[2].NumberFormat = "@";

                worksheet.Cells[1, 1] = "BẢNG ĐIỂM SINH VIÊN THEO MÔN HỌC ";
                worksheet.Cells[3, 2] = "Mã giảng viên: " + MaGV;
                worksheet.Cells[4, 2] = "Mã môn:  " + MaMH;
                worksheet.Cells[5, 2] = "Tên Môn:  " + dataGridView1.CurrentRow.Cells[1].Value;
                worksheet.Cells[6, 2] = "Số lượng sinh viên:  " + lblTongSV.Text;

                //Chỉnh header text cho từng cột
                worksheet.Cells[9, 1] = "STT";
                worksheet.Cells[9, 2] = "Mã sinh viên";
                worksheet.Cells[9, 3] = "Điểm thường xuyên";
                worksheet.Cells[9, 4] = "Điểm thi kết thúc học phần";
                worksheet.Cells[9, 5] = "Điểm trung bình";
                worksheet.Cells[9, 6] = "Điểm chữ";

                worksheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlPortrait;
                worksheet.PageSetup.PaperSize = Microsoft.Office.Interop.Excel.XlPaperSize.xlPaperA3;
                worksheet.PageSetup.LeftMargin = 0;
                worksheet.PageSetup.RightMargin = 0;
                worksheet.PageSetup.TopMargin = 0;
                worksheet.PageSetup.BottomMargin = 0;
                worksheet.Range["A1"].ColumnWidth = 5;
                worksheet.Range["B1"].ColumnWidth = 15;
                worksheet.Range["C1"].ColumnWidth = 30;
                worksheet.Range["D1"].ColumnWidth = 30;
                worksheet.Range["E1"].ColumnWidth = 30;
                worksheet.Range["F1"].ColumnWidth = 10;

                int mon = dataGridView2.RowCount;
                worksheet.Range["A1", "E100"].Font.Name = "Times New Roman";
                worksheet.Range["A1", "F1"].MergeCells = true;
                worksheet.Range["A1", "F1"].Font.Bold = true;
                worksheet.Range["A9", "F" + (mon + 9)].Borders.LineStyle = 1;

                worksheet.Range["A1", "G1"].HorizontalAlignment = 3;
                worksheet.Range["A9", "F9"].HorizontalAlignment = 3;
                worksheet.Range["A10", "A" + (mon + 9)].HorizontalAlignment = 3;
                worksheet.Range["B10", "B" + (mon + 9)].HorizontalAlignment = 3;
                worksheet.Range["C10", "C" + (mon + 9)].HorizontalAlignment = 3;
                worksheet.Range["D10", "D" + (mon + 9)].HorizontalAlignment = 3;
                worksheet.Range["E10", "E" + (mon + 9)].HorizontalAlignment = 3;
                worksheet.Range["F10", "F" + (mon + 9)].HorizontalAlignment = 3;

                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    worksheet.Cells[1][i + 10] = i + 1;
                    for (int j = 0; j < dataGridView2.ColumnCount; j++)
                    {
                        DataGridViewCell cell = dataGridView2.Rows[i].Cells[j];
                        worksheet.Cells[j + 2][i + 10] = cell.Value.ToString();
                    }
                }
                workbook.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, missingValue, missingValue, true, false, XlSaveAsAccessMode.xlNoChange,
                    XlSaveConflictResolution.xlLocalSessionChanges, missingValue, missingValue);
                MessageBox.Show("Lưu thành công", "Thông báo");
                workbook.Close(true, missingValue, missingValue);
                application.Quit();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.Message + "\n Không thể lưu file!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {

                //Đối tượng COM là đối tượng không được quản lý, chúng sẽ không được giải phóng bộ nhớ khi kết thúc, 
                //vì vậy phải giải phóng chúng
                releaseObject(worksheet);
                releaseObject(workbook);
                releaseObject(application);
            }
        }

        private void releaseObject(object obj)
        {
            //This method is used to explicitly control the lifetime of a COM object used from managed code. 
            //You should use this method to free the underlying COM object that holds references to resources in a timely manner 
            //or when objects must be freed in a specific order
            try
            {
                //System.Runtime.InteropServices.Marshal.ReleaseComObject () có thể được sử dụng 
                //để giải phóng đối tượng COM trước khi chúng được hoàn thiện (trong bộ thu gom rác)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception)
            {
                obj = null;
            }
            finally
            {
                //Forces an immediate garbage collection of all generations.
                //Thu dọn đối tượng obj khi nó không tham  chiếu đến vùng bộ nhớ nào
                GC.Collect();
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            dataGridView2.ReadOnly = true;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.AllowUserToDeleteRows = false;
            Saved = true;

            Data.Database data = new Data.Database();
            if(diemUpdate.Count != 0)
            {
                foreach (DiemSV item in diemUpdate)
                {
                    data.ExecCmd("EXEC Update_Diem '" + item.MaSV + "', '" + item.MaMH + "', " + item.DiemTX + ", " + item.DiemKTHP + ", " + item.DiemTB + ", '" + item.DiemChu + "'");
                }
                diemUpdate.Clear();
            }


            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '"
                + MaMH+ "'");
            //Xóa điểm ở csdl

        }

        private void btnChinhSua_Click(object sender, EventArgs e)
        {
            dataGridView2.ReadOnly = false;
            dataGridView2.AllowUserToAddRows = !btnHuyKQ.Visible;
            dataGridView2.AllowUserToDeleteRows = false;
            dataGridView2.Columns[3].ReadOnly = true;
            dataGridView2.Columns[4].ReadOnly = true;
            Saved = false;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("Bạn chưa chọn dòng nào để xóa", "Thông báo");
            }
            else
            {
                
            }
        }

        private void btnHuyKQ_Click(object sender, EventArgs e)
        {
            txtTimKiem.Text = "";
            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
            btnHuyKQ.Visible = false;
            lblTongSV.Text = dataGridView2.RowCount.ToString() + " sinh viên";
        }

        private void btnNhapExcel_Click(object sender, EventArgs e)
        {
            //try
            //{

            System.Data.DataTable dt = new System.Data.DataTable();
            dt = Import();
            if (dt != null)
            {
                //checkExcelData();
                dataGridView2.DataSource = dt;
                Saved = false;
            }
            lblTongSV.Text = dataGridView2.RowCount + " sinh viên";
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Lỗi : " + ex.Message + "\n Không thể import file", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //}
        }

        private void hienThiSoLuongSinhVien()
        {
            if (dataGridView2.AllowUserToAddRows == false)
            {
                lblTongSV.Text = dataGridView2.RowCount.ToString() + " sinh viên";
            }
            else
            {
                lblTongSV.Text = (dataGridView2.RowCount - 1).ToString() + " sinh viên";
            }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < dataGridView2.RowCount - 1)
            {
                if(string.IsNullOrEmpty(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString()))
                {                   
                    MessageBox.Show("Mã sinh viên đang bị bỏ trống.\nGiá trị của mã sinh viên sẽ được đặt mặc định.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    dataGridView2.Rows[e.RowIndex].Cells[0].Value = e.RowIndex;
                }
                try
                {
                    double diemtx = double.NaN, diemkt = double.NaN, diemtb = double.NaN;
                    string diemChu = string.Empty;
                    bool check = true;
                    if ( !string.IsNullOrEmpty(dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString()))
                    {
                        diemtx = double.Parse(dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString());
                        if (diemtx < 0 || diemtx > 10)
                        {
                            diemtx = 0;
                            MessageBox.Show("Giá trị của điểm thường xuyên chỉ ở trong khoảng 0-10.\nĐiểm thường xuyên sẽ được đặt giá trị bằng 0.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            dataGridView2.Rows[e.RowIndex].Cells[1].Value = "0";
                        }
                    }
                    else
                    {
                        check = false;
                    }

                    if (!string.IsNullOrEmpty(dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString() ))
                    {
                        diemkt = double.Parse(dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString());
                        if (diemkt < 0 || diemkt > 10)
                        {
                            diemkt = 0;
                            MessageBox.Show("Giá trị của điểm kết thúc chỉ ở trong khoảng 0-10.\nĐiểm kết thúc sẽ được đặt giá trị bằng 0.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            dataGridView2.Rows[e.RowIndex].Cells[2].Value = "0";
                        }
                    }
                    else
                    {
                        check = false;
                    }
                    if (check)
                    {

                        diemtb = Math.Round((diemtx + diemkt * 2) / 3, 2);
                        dataGridView2.Rows[e.RowIndex].Cells[3].Value = diemtb.ToString();
                        if (diemtb >= 8.5) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "A";
                        else if (diemtb >= 7.7) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "B+";
                        else if (diemtb >= 7) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "B";
                        else if (diemtb >= 6.2) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "C+";
                        else if (diemtb >= 5.5) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "C";
                        else if (diemtb >= 4.7) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "D+";
                        else if (diemtb >= 4.0) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "D";
                        else if (diemtb < 4) dataGridView2.Rows[e.RowIndex].Cells[4].Value = "F";
                        diemChu = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
                    }                 
                    if(diemTrongCSDL.Contains(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString().Trim()))
                    {
                        diemUpdate.Add(new DiemSV(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString().Trim(), MaMH, diemtx, diemkt, diemtb, diemChu));
                    }
                    else
                    {
                        diemInsert.Add(new DiemSV(dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString().Trim(), MaMH, diemtx, diemkt, diemtb, diemChu));
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex);
                }
            }
        }

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Lỗi : " + e.Exception.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            if (txtTimKiem.Text != "")
            {
                dataGridView2.DataSource = new Data.Database().SelectData("EXEC TimKiem_Diem '" + txtTimKiem.Text + "', '" + MaMH + "'");
                btnHuyKQ.Visible = true;
                lblTongSV.Text = dataGridView2.RowCount.ToString() + " sinh viên";
            }
            else
            {
                MessageBox.Show("Thông tin tìm kiếm đang bị bỏ trống", "Thông báo");
            }
        }

        public System.Data.DataTable Import()
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*xlsx", ValidateNames = true })
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                if (ofd.ShowDialog() == DialogResult.OK)
                {

                    using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        IExcelDataReader reader;
                        if (ofd.FilterIndex == 2)
                        {
                            reader = ExcelReaderFactory.CreateBinaryReader(stream);
                        }
                        else
                        {
                            reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                        }
                        DataSet ds = new DataSet();
                        ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                        {
                            ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                            {
                                UseHeaderRow = true
                            }
                        });
                        foreach (System.Data.DataTable item in ds.Tables)
                        {
                            dt = item;
                        }
                        reader.Close();

                    }
                    return dt;
                }
            }
            return null;
        }
    }
}
