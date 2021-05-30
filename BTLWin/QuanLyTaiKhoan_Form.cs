﻿using Data;
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
    public partial class QuanLyTaiKhoan_Form : Form
    {
        List<string> tdnCSDL, maCSDL;
        string userName;
        public QuanLyTaiKhoan_Form(string userName)
        {
            InitializeComponent();
            tdnCSDL = new List<string>();
            maCSDL = new List<string>();
            this.userName = userName;
        }
        /*
         * Phần này làm gần giống phần QuanLyDiem nhưng đơn giản hơn
         * isSaved = true khi dữ liệu đã được lưu
         * isSaved = false khi dữ liệu chưa được lưu
         */
        private bool isSaved = true;
        /*
         * accountType lưu loại tài khoản được chọn
         * 1: Giảng viên
         * 0: Sinh viên
         * 2: Quản trị viên
         */
        private int accountType;
        public void loadDuLieu(int index)
        {
            try
            {
                accountType = index;
                if (accountType == 0)
                {
                    dataGridView2.DataSource = new Database().SelectData("SELECT * FROM TAIKHOANSV");
                    dataGridView2.Columns[0].HeaderText = "Tên đăng nhập";
                    dataGridView2.Columns[1].HeaderText = "Mật khẩu";
                    dataGridView2.Columns[2].HeaderText = "Mã sinh viên";

                }
                else if (accountType == 1)
                {
                    dataGridView2.DataSource = new Database().SelectData("SELECT * FROM TAIKHOANGV");
                    dataGridView2.Columns[0].HeaderText = "Tên đăng nhập";
                    dataGridView2.Columns[1].HeaderText = "Mật khẩu";
                    dataGridView2.Columns[2].HeaderText = "Mã giảng viên";
                }
                else if (accountType == 2)
                {
                    dataGridView2.DataSource = new Database().SelectData("SELECT * FROM TAIKHOANQTV WHERE TaiKhoan != '" + this.userName + "'");
                    dataGridView2.Columns[0].HeaderText = "Tên đăng nhập";
                    dataGridView2.Columns[1].HeaderText = "Mật khẩu";
                }
                dataGridView2.ReadOnly = true;
                btnXoa.Enabled = true;
                btnLuu.Enabled = true;
                isSaved = true;
                btnChinhSua.BackColor = SystemColors.ControlDark;
                txtTimKiem.Clear();
                btnHuyKQ.Visible = false;
                tdnCSDL.Clear();
                maCSDL.Clear();
                foreach (DataGridViewRow item in dataGridView2.Rows)
                {
                    tdnCSDL.Add(item.Cells[0].Value.ToString().Trim());
                    if (accountType != 2)
                    {
                        maCSDL.Add(item.Cells[2].Value.ToString().Trim());
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        public int KiemTraDuLieu()
        {
            try
            {
                //Error là chuỗi thông báo lỗi, thông báo ra các dòng dữ liệu bị lỗi
                String error = "";
                //List er dùng để lưu các dòng datagridview có dữ liệu không hợp lệ
                List<int> er = new List<int>();
                switch (accountType)
                {
                    //TH: Tài khoản sinh viên
                    case 0:
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            //Lấy dữ liệu
                            String passWord1 = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            String MaSV = dataGridView2.Rows[i].Cells[2].Value.ToString();
                            if (passWord1.Length == 0 || passWord1.Length > 10 || MaSV.Length == 0 || MaSV.Length > 10)
                            {
                                error += " " + (i + 1);
                                er.Add(i);
                                continue;
                            }
                            else
                            {
                                //Ô Tài khoản không được phép nhập mà lấy trực tiếp từ MaSV làm tài khoản
                                dataGridView2.Rows[i].Cells[0].Value = MaSV;
                            }
                            DataTable dt = new Database().SelectData("SELECT * FROM SINHVIEN WHERE MaSV= '" + MaSV + "'");
                            if (dt.Rows.Count == 0)
                            {
                                error += " " + (i + 1);
                                er.Add(i);
                                continue;
                            }

                        }
                        break;
                    //TH: Tài khoản giảng viên
                    case 1:
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            //Lấy dữ liệu
                            //Ô Tài khoản không được phép nhập mà lấy trực tiếp từ MaGV làm tài khoản
                            String passWord = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            String MaGV = dataGridView2.Rows[i].Cells[2].Value.ToString();
                            /*
                             * Kiểm tra dữ liệu
                             * Các ô không được để trống .length phải khác 0
                             * password và MaGV chỉ từ 0 đến 10 ký tự
                             */
                            if (passWord.Length == 0 || passWord.Length > 10 || MaGV.Length == 0 || MaGV.Length > 10)
                            {
                                error += " " + (i + 1);
                                er.Add(i);
                                continue;
                            }
                            else
                            {
                                dataGridView2.Rows[i].Cells[0].Value = MaGV;
                            }
                            DataTable dt = new Database().SelectData("SELECT * FROM GIANGVIEN WHERE MaGV= '" + MaGV + "'");
                            if (dt.Rows.Count == 0)
                            {
                                error += " " + (i + 1);
                                er.Add(i);
                                continue;
                            }
                        }
                        break;
                    case 2:
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            String userName2 = dataGridView2.Rows[i].Cells[0].Value.ToString();
                            String passWord2 = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            if (userName2.Length == 0 || passWord2.Length == 0 || passWord2.Length > 10)
                            {
                                error += " " + (i + 1);
                                er.Add(i);
                                continue;
                            }

                        }
                        break;
                    default:
                        break;
                }
                if (er.Count != 0)
                {
                    DialogResult rsl = MessageBox.Show("Dữ liệu lỗi tại các dòng " + error + "\nNếu tiếp tục sẽ bỏ qua các dòng lỗi, bạn có muốn tiếp tục?", "Cảnh báo", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (rsl == DialogResult.Yes)
                    {
                        for (int i = er.Count - 1; i >= 0; i--)
                        {
                            dataGridView2.Rows.RemoveAt(er[i]);
                        }
                        return 1;
                    }
                    else
                    {
                        return 0;
                    }
                }
                else return 1;

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return 0;
            }
        }

        public void update()
        {
            try
            {
                switch (accountType)
                {
                    case 0:
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            String userName = dataGridView2.Rows[i].Cells[0].Value.ToString();
                            String passWord = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            String MaSV = dataGridView2.Rows[i].Cells[2].Value.ToString();
                            new Database().ExecCmd("EXEC Update_TaiKhoanSV '" + userName + "', '" + passWord + "', '" + MaSV + "'");
                        }
                        break;
                    case 1:
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            String userName = dataGridView2.Rows[i].Cells[0].Value.ToString();
                            String passWord = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            String MaGV = dataGridView2.Rows[i].Cells[2].Value.ToString();
                            new Database().ExecCmd("EXEC Update_TaiKhoanGV '" + userName + "', '" + passWord + "', '" + MaGV + "'");
                        }
                        break;
                    case 2:
                        for (int i = 0; i < dataGridView2.RowCount - 1; i++)
                        {
                            String userName = dataGridView2.Rows[i].Cells[0].Value.ToString();
                            String passWord = dataGridView2.Rows[i].Cells[1].Value.ToString();
                            new Database().ExecCmd("EXEC Update_TaiKhoanQTV '" + userName + "', '" + passWord + "'");
                        }
                        break;
                    default:
                        break;
                }
                MessageBox.Show("Cập nhật thành công");

            }
            catch (Exception)
            {
                MessageBox.Show("Cập nhật thất bại");
            }
        }

        public void TimKiem()
        {
            try
            {
                if (txtTimKiem.Text != "")
                {
                    for (int i = dataGridView2.RowCount - 2; i >= 0; i--)
                    {
                        if (dataGridView2.Rows[i].Cells[0].Value.ToString() != txtTimKiem.Text)
                        {
                            dataGridView2.Rows.RemoveAt(i);
                        }
                    }
                }
                btnHuyKQ.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void xoa(string TaiKhoan, string MatKhau)
        {
            try
            {
                if (TaiKhoan.Length != 0 && MatKhau.Length != 0)
                {
                    switch (accountType)
                    {
                        case 0:
                            new Database().ExecCmd("DELETE TAIKHOANSV WHERE TaiKhoan= '" + TaiKhoan + "'");
                            break;
                        case 1:
                            new Database().ExecCmd("DELETE TAIKHOANGV WHERE TaiKhoan= '" + TaiKhoan + "'");
                            break;
                        case 2:
                            new Database().ExecCmd("DELETE TAIKHOANQTV WHERE TaiKhoan= '" + TaiKhoan + "'");
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private void QuanLyTaiKhoan_Form_Load(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Rows.Add(3);
            dataGridView1.Rows[0].Cells[0].Value = "Sinh viên";
            dataGridView1.Rows[1].Cells[0].Value = "Giảng viên";
            dataGridView1.Rows[2].Cells[0].Value = "Quản trị viên";
            dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.NotSortable;
            loadDuLieu(0);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.Message + ".\nKhông thể tải dữ liệu.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (!isSaved)
                {
                    DialogResult rsl = MessageBox.Show("Dữ liệu bạn vừa thay đổi chưa được lưu, bạn có muốn lưu không?", "Hỏi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (rsl == DialogResult.Yes)
                    {
                        btnLuu_Click(sender, e);
                    }
                }
                loadDuLieu(e.RowIndex);
                isSaved = true;
            }
            catch (Exception) { }
        }

        private void btnChinhSua_Click(object sender, EventArgs e)
        {
            dataGridView2.ReadOnly = false;
            btnXoa.Enabled = true;
            btnLuu.Enabled = true;
            isSaved = false;
            btnChinhSua.BackColor = SystemColors.ControlDarkDark;
            dataGridView2.Columns[0].ReadOnly = true;
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            if (!isSaved)
            {
                int check = KiemTraDuLieu();
                if (check == 1)
                {
                    update();
                    loadDuLieu(dataGridView1.CurrentRow.Index);
                }
                else
                {
                    MessageBox.Show("Cập nhật thất bại");
                }
            }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (accountType != 2)
                {
                    String Ma = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
                    dataGridView2.Rows[e.RowIndex].Cells[0].Value = Ma;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            try
            {
                if (txtTimKiem.Text.ToString() == "")
                {
                    MessageBox.Show("Thông tin tìm kiếm đang bị bỏ trống", "Thông báo");
                }
                else
                {
                    TimKiem();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void btnHuyKQ_Click(object sender, EventArgs e)
        {
            loadDuLieu(dataGridView1.CurrentRow.Index);
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            try
            {
                if (isSaved)
                {
                    if(dataGridView2.CurrentCell != null)
                    {
                        DialogResult result = MessageBox.Show("Dữ liệu bị xóa sẽ không thể khôi phục được.\nBạn có chắc muốn xóa?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result == DialogResult.Yes)
                        {
                            string TaiKhoan = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[0].Value.ToString();
                            string MatKhau = dataGridView2.Rows[dataGridView2.CurrentCell.RowIndex].Cells[1].Value.ToString();
                            xoa(TaiKhoan, MatKhau);
                            loadDuLieu(dataGridView1.CurrentRow.Index);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Bạn phải lưu thay đổi trước khi thực hiện xóa");
                    btnLuu_Click(sender, e);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi : " + ex.Message + ".\nXóa thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void themTaiKhoan()
        {
            loadDuLieu(dataGridView1.CurrentRow.Index);
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            using (ThemTaiKhoan_Form themtk = new ThemTaiKhoan_Form(accountType, tdnCSDL, maCSDL))
            {
                themtk.ShowDialog(this);
            }
        }
    }
}
