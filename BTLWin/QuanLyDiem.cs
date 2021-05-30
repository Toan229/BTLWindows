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
using System.Threading;

namespace BTLWin
{
    struct DsDiem
    {
        public DiemSV diemSV;
        public int index;
        public DsDiem(DiemSV diemSV, int index)
        {
            this.diemSV = diemSV;
            this.index = index;
        }
    }

    public partial class QuanLyDiem : Form
    {
        string MaGV, MaMH;

        List<DsDiem> diemUpdate, diemInsert;
        List<string> diemTrongCSDL;
        System.Data.DataTable checkUpdateTable;
        bool escKey, xoabtn;

        public QuanLyDiem()
        {
            InitializeComponent();
        }

        public QuanLyDiem(string username)
        {
            InitializeComponent();
            this.MaGV = username;
            diemUpdate = new List<DsDiem>();
            diemInsert = new List<DsDiem>();
            diemTrongCSDL = new List<string>();
            escKey = false;
            xoabtn = false;
        }

        private void QuanLyDiem_Load(object sender, EventArgs e)
        {
            dataGridView1.DataSource = new Database().SelectData("EXEC TimKiemMonHoc_TheoMaGV '" + MaGV + "'");
            checkUpdateTable = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + dataGridView1.Rows[0].Cells[0].Value + "'");
            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + dataGridView1.Rows[0].Cells[0].Value + "'"); ;
            foreach (DataRow item in checkUpdateTable.Rows)
            {
                diemTrongCSDL.Add(item.ItemArray[0].ToString());
            }
            MaMH = dataGridView1.Rows[0].Cells[0].Value.ToString();
            hienThiSoLuongSinhVien();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > -1)
            {
                dataGridView2.ReadOnly = true;
                dataGridView2.AllowUserToAddRows = false;
                dataGridView2.AllowUserToDeleteRows = false;

                checkUpdateTable = null;
                checkUpdateTable = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '"
                    + dataGridView1.Rows[e.RowIndex].Cells[0].Value + "'");

                dataGridView2.DataSource = null;
                dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '"
                    + dataGridView1.Rows[e.RowIndex].Cells[0].Value + "'");
                diemTrongCSDL.Clear();
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    diemTrongCSDL.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
                }
                hienThiSoLuongSinhVien();
                MaMH = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
                txtTimKiem.Text = string.Empty;
                btnHuyKQ.Visible = false;
                btnChinhSua.BackColor = SystemColors.ControlDark;
            }
        }

        private void btnXuatFile_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Excel File(*.xlsx)|*.xlsx";
            saveFile.ShowDialog();
            string fileName = saveFile.FileName;
            dataGridView2.AllowUserToAddRows = false;
            btnChinhSua.BackColor = SystemColors.ControlDark;
            xuatFileExcel(fileName);

        }

        private void xuatFileExcel(string fileName)
        {
            this.Cursor = Cursors.WaitCursor;
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

                workbook.Close(true, missingValue, missingValue);
                application.Quit();
                this.Cursor = Cursors.Default;

                DialogResult result = MessageBox.Show("Bảng điểm đã được lưu vào file.\nBạn có muốn mở file hay không ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                releaseObject(worksheet);
                releaseObject(workbook);
                releaseObject(application);
            }
        }

        public void releaseObject(object obj)
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

        private void btnChinhSua_Click(object sender, EventArgs e)
        {
            dataGridView2.ReadOnly = false;
            dataGridView2.AllowUserToAddRows = !btnHuyKQ.Visible;
            dataGridView2.Columns[0].ReadOnly = btnHuyKQ.Visible;
            dataGridView2.AllowUserToDeleteRows = false;
            dataGridView2.Columns[3].ReadOnly = true;
            dataGridView2.Columns[4].ReadOnly = true;
            btnChinhSua.BackColor = SystemColors.ControlDarkDark;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count == 0)
            {
                MessageBox.Show("Bạn chưa chọn dòng nào để xóa", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                DialogResult result = MessageBox.Show("Dữ liệu bị xóa sẽ không thể khôi phục.\nBạn có chắc muốn xóa không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    xoabtn = true;
                    xoa();
                    xoabtn = false;
                }
            }
        }

        private void xoa()
        {
            Database data = new Database();
            List<int> position = new List<int>();
            foreach (DataGridViewRow item in dataGridView2.SelectedRows)
            {
                if (!dataGridView2.AllowUserToAddRows || (dataGridView2.AllowUserToAddRows && item.Index < dataGridView2.RowCount - 1))
                {
                    if (dataGridView2.Rows[item.Index].ErrorText == "" && diemTrongCSDL.Contains(item.Cells[0].Value.ToString()))
                    {
                        diemTrongCSDL.Remove(item.Cells[0].Value.ToString());
                        data.ExecCmd("EXEC Delete_Diem '" + item.Cells[0].Value.ToString() + "', '" + MaMH + "'");
                    }
                    position.Add(item.Index);
                }
            }
            position.Sort();
            for (int i = position.Count - 1; i >= 0; i--)
            {
                dataGridView2.Rows.RemoveAt(position[i]);
            }
            checkUpdateTable = null;
            checkUpdateTable = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");

            dataGridView2.CurrentCell = null;
            dataGridView2.DataSource = null;
            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
            diemTrongCSDL.Clear();
            for (int i = 0; i < checkUpdateTable.Rows.Count; i++)
            {
                diemTrongCSDL.Add(checkUpdateTable.Rows[i].ItemArray[0].ToString());
            }
            btnHuyKQ.Visible = false;
            txtTimKiem.Text = string.Empty;
            hienThiSoLuongSinhVien();
        }

        private void btnHuyKQ_Click(object sender, EventArgs e)
        {
            txtTimKiem.Text = string.Empty;
            dataGridView2.DataSource = null;
            btnChinhSua.BackColor = SystemColors.ControlDark;
            dataGridView2.ReadOnly = true;
            checkUpdateTable = null;
            checkUpdateTable = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");

            dataGridView2.DataSource = null;
            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
            diemTrongCSDL.Clear();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                diemTrongCSDL.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
            }
            btnHuyKQ.Visible = false;
            hienThiSoLuongSinhVien();
        }

        private void btnNhapExcel_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Dữ liệu đọc trong file sẽ ghi đè lên dữ liệu hiện có.\nCác dòng bị lỗi dữ liệu sẽ bị bỏ qua.\nBạn có chắc muốn nhập dữ liệu từ file Excel ?", "Thông báo",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                try
                {

                    System.Data.DataTable dt = new System.Data.DataTable();
                    dt = Import();
                    if (dt != null)
                    {
                        new Database().ExecCmd("DELETE FROM DIEM WHERE MaMH = '" + MaMH + "'");
                        int rowsEffected = 0;
                        List<string> result1 = new List<string>();
                        double diemtx, diemkt;
                        foreach (DataRow item in dt.Rows)
                        {
                            if (item.ItemArray[0] != null && item.ItemArray[1] != null && item.ItemArray[2] != null)
                            {
                                if (double.TryParse(item.ItemArray[1].ToString(), out diemtx) && double.TryParse(item.ItemArray[2].ToString(), out diemkt) && diemtx > 0 && diemtx < 10
                                    && diemkt > 0 && diemkt < 10)
                                {
                                    if (diemTrongCSDL.Contains(item.ItemArray[0].ToString()))
                                    {
                                        result1 = new Database().ExecCmd("EXEC Update_Diem '" + item.ItemArray[0] + "', '" + MaMH + "', " + diemtx + ", " + diemkt + ", " + tinhDiem(diemtx, diemkt)[0] + ", '" + tinhDiem(diemtx, diemkt)[1] + "', '" + item.ItemArray[0] + "'");
                                    }
                                    //else if ()
                                    //{
                                    //    rowEffected = new Database().ExecCmd("INSERT INTO DIEM VALUES('" + item.ItemArray[0] + "', '" + MaMH + "', " + diemtx + ", " + diemkt + ", " + tinhDiem(diemtx, diemkt)[0] + ", '" + tinhDiem(diemtx, diemkt)[1] + "')");
                                    //    if (rowEffected != 0)
                                    //    {
                                    //        diemTrongCSDL.Add(item.ItemArray[0].ToString());
                                    //    }
                                    //}
                                    rowsEffected += int.Parse(result1[0]);
                                }
                            }
                        }


                        dataGridView2.DataSource = null;
                        dataGridView2.ReadOnly = true;
                        dataGridView2.AllowUserToAddRows = false;
                        dataGridView2.CurrentCell = null;
                        dataGridView2.DataSource = null;
                        dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
                        if (rowsEffected == dt.Rows.Count)
                        {
                            MessageBox.Show("Nhập dữ liệu thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("Số dòng nhập thành công : " + rowsEffected + "/" + dt.Rows.Count, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        diemTrongCSDL.Clear();
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            diemTrongCSDL.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
                        }
                        hienThiSoLuongSinhVien();
                    }
                    hienThiSoLuongSinhVien();
                    btnChinhSua.BackColor = SystemColors.ControlDark;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Lỗi : " + ex.Message + "\n Không thể import file", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private List<string> tinhDiem(double diemtx, double diemkt)
        {
            List<string> result = new List<string>(2);
            string diemchu;
            double diemtb = Math.Round((diemtx + diemkt * 2) / 3, 2);
            if (diemtb >= 8.5) diemchu = "A";
            else if (diemtb >= 7.7) diemchu = "B+";
            else if (diemtb >= 7) diemchu = "B";
            else if (diemtb >= 6.2) diemchu = "C+";
            else if (diemtb >= 5.5) diemchu = "C";
            else if (diemtb >= 4.7) diemchu = "D+";
            else if (diemtb >= 4.0) diemchu = "D";
            else diemchu = "F";
            result.Add(diemtb.ToString());
            result.Add(diemchu);
            return result;
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

        private void dataGridView2_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            if (dataGridView2.AllowUserToAddRows && e.RowIndex < dataGridView2.RowCount - 1)
            {
                MessageBox.Show("Lỗi : " + e.Exception.Message + " ở dòng " + (e.RowIndex + 1), "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnReloadData_Click(object sender, EventArgs e)
        {
            dataGridView2.ReadOnly = true;
            dataGridView2.AllowUserToAddRows = false;
            dataGridView2.AllowUserToDeleteRows = false;
            btnHuyKQ.Visible = false;
            txtTimKiem.Text = string.Empty;
            dataGridView2.DataSource = null;
            dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
            btnChinhSua.BackColor = SystemColors.ControlDark;
            diemTrongCSDL.Clear();
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                diemTrongCSDL.Add(dataGridView2.Rows[i].Cells[0].Value.ToString());
            }
            hienThiSoLuongSinhVien();
        }

        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (dataGridView2.ReadOnly == false)
            {
                dataGridView2.Rows[e.RowIndex].ErrorText = "";
                if (dataGridView2.Rows[e.RowIndex].IsNewRow) { return; }
                if (e.ColumnIndex == 0)
                {
                    string sv = e.FormattedValue.ToString().Trim();
                    if (sv.Length > 10)
                    {
                        e.Cancel = true;
                        dataGridView2.Rows[e.RowIndex].ErrorText = "Độ dài tối đa của mã sinh viên là 10";
                    }
                }
                else if (e.ColumnIndex == 1 || e.ColumnIndex == 2)
                {
                    double value;
                    if (!double.TryParse(e.FormattedValue.ToString(), out value) || value < 0 || value > 10)
                    {
                        e.Cancel = true;
                        dataGridView2.Rows[e.RowIndex].ErrorText = "Giá trị của " + dataGridView2.Columns[e.ColumnIndex].HeaderText.ToLower() + " chỉ trong khoảng 0 đến 10 và chỉ chứa các kì từ 0-9";
                    }
                }
            }
        }

        private void dataGridView2_RowValidating(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (!escKey)
            {
                if (dataGridView2.AllowUserToAddRows && e.RowIndex < dataGridView2.RowCount - 1)
                {
                    DataGridViewRow row = dataGridView2.Rows[e.RowIndex];
                    e.Cancel = !checkCell(row.Cells[0]) || !checkCell(row.Cells[1]) || !checkCell(row.Cells[2]);
                    if (!e.Cancel)
                    {
                        capNhapDiem(row, e);
                    }
                }
            }
        }

        private void capNhapDiem(DataGridViewRow row, DataGridViewCellCancelEventArgs e)
        {
            try
            {
                bool updateTable = false;
                List<string> result = new List<string>();
                if (row.Index > checkUpdateTable.Rows.Count - 1)
                {
                    result = new Database().ExecCmd("INSERT INTO DIEM VALUES('" + row.Cells[0].Value + "', '" + MaMH + "', " + row.Cells[1].Value + ", " + row.Cells[2].Value + ", " + row.Cells[3].Value + ", '" + row.Cells[4].Value + "')");
                    updateTable = true;
                }
                else
                {
                    if (!xoabtn)
                    {
                        if (row.Cells[0].Value.ToString() != checkUpdateTable.Rows[row.Index].ItemArray[0].ToString() || row.Cells[1].Value.ToString() != checkUpdateTable.Rows[row.Index].ItemArray[1].ToString() || row.Cells[2].Value.ToString() != checkUpdateTable.Rows[row.Index].ItemArray[2].ToString())
                        {
                            result = new Database().ExecCmd("EXEC Update_Diem '" + checkUpdateTable.Rows[row.Index].ItemArray[0] + "', '" + MaMH + "', " + row.Cells[1].Value + ", " + row.Cells[2].Value + ", " + row.Cells[3].Value + ", '" + row.Cells[4].Value + "', '" + row.Cells[0].Value + "'");
                            updateTable = true;
                        }
                    }
                }

                if (updateTable)
                {
                    if (int.Parse(result[0]) == 0)
                    {
                        MessageBox.Show("Không thể cập nhập điểm.\nLỗi : " + result[1] + "\nNhấn ESC để thoát chế độ chỉnh sửa và loại bỏ dòng lỗi", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        e.Cancel = true;
                    }
                    else
                    {
                        checkUpdateTable = null;
                        checkUpdateTable = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
                        foreach (DataGridViewCell item in row.Cells)
                        {
                            item.ErrorText = string.Empty;
                        }
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        private bool checkCell(DataGridViewCell cell)
        {
            if (cell.Value != null && cell.Value.ToString() == string.Empty)
            {
                MessageBox.Show("Mã sinh viên, điểm thường xuyên và điểm kết thúc học phần không cho phép giá trị NULL", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex != 3 && e.ColumnIndex != 4)
            {
                if (e.RowIndex > checkUpdateTable.Rows.Count - 1)
                {
                    dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "Giá trị đã được thay đổi";
                }
                else
                {
                    if (dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() != checkUpdateTable.Rows[e.RowIndex].ItemArray[e.ColumnIndex].ToString())
                    {
                        dataGridView2.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "Giá trị đã được thay đổi";
                    }
                }
            }
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                escKey = true;
                dataGridView2.CurrentCell = null;
                dataGridView2.DataSource = null;
                dataGridView2.DataSource = new Database().SelectData("EXEC TimKiem_Diem_TheoMaMH '" + MaMH + "'");
                escKey = false;
                hienThiSoLuongSinhVien();
            }
        }

        private void dataGridView2_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            double diemtx, diemkt;
            if (dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString() != string.Empty && double.TryParse(dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString(), out diemtx) && double.TryParse(dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString(), out diemkt))
            {
                dataGridView2.Rows[e.RowIndex].Cells[3].Value = tinhDiem(diemtx, diemkt)[0];
                dataGridView2.Rows[e.RowIndex].Cells[4].Value = tinhDiem(diemtx, diemkt)[1];
                hienThiSoLuongSinhVien();
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            if (txtTimKiem.Text != "")
            {
                dataGridView2.DataSource = null;
                dataGridView2.DataSource = new Data.Database().SelectData("EXEC TimKiem_Diem '" + txtTimKiem.Text + "', '" + MaMH + "'");
                dataGridView2.AllowUserToAddRows = false;
                btnHuyKQ.Visible = true;
                hienThiSoLuongSinhVien();
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
