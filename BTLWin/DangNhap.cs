using Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BTLWin
{
    public partial class DangNhap : Form
    {
        bool mov, keyFlag;
        int movX, movY;
        public DangNhap()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnMinimized_Click(object sender, EventArgs e)
        {
            //Ẩn form xuống thanh taskbar
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            //Khi ấn chuột xuống lấy tọa độ của chuột so với form
            movX = e.X;
            movY = e.Y;
            mov = true;
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            //Nhả chuột ra
            mov = false;
        }

        private void DangNhap_Load(object sender, EventArgs e)
        {
            //Hiển thị mật khẩu
            lblSaiTKorMK.Visible = false;//Ẩn thông báo sai mk hoặc tên đăng nhập
            cmbLoaiTK.SelectedIndex = 0;
            keyFlag = false;
        }

        private void txtDangNhap_Validating(object sender, CancelEventArgs e)
        {
            if (lblSaiTKorMK.Visible == true)
            {
                lblSaiTKorMK.Visible = false;
            }
            if (txtDangNhap.Text == "")
            {
                errorProvider1.SetError(txtDangNhap, "Tên đăng nhập đang bị bỏ trống");
                txtDangNhap.Focus();
            }
            else
            {
                errorProvider1.SetError(txtDangNhap, "");
                e.Cancel = false;
            }
        }

        private void txtMatKhau_Validating(object sender, CancelEventArgs e)
        {
            if (lblSaiTKorMK.Visible == true)
            {
                lblSaiTKorMK.Visible = false;
            }
            if (txtMatKhau.Text == "")
            {
                errorProvider2.SetError(txtMatKhau, "Mật khẩu đang bị bỏ trống");
                txtMatKhau.Focus();
            }
            else
            {
                errorProvider2.SetError(txtMatKhau, "");
                e.Cancel = false;
            }
        }

        private void chkHienThiMK_CheckedChanged(object sender, EventArgs e)
        {
            if (chkHienThiMK.Checked)
            {
                txtMatKhau.PasswordChar = '\0';
            }
            else
            {
                txtMatKhau.PasswordChar = '*';
            }
        }

        private void btnDangNhap_Click(object sender, EventArgs e)
        {
            DataTable table = new DataTable();
            switch (cmbLoaiTK.SelectedIndex)
            {
                case 0:
                    table = new Database().SelectData("EXEC DangNhap_SV '" + txtDangNhap.Text + "', '" + txtMatKhau.Text + "'");
                    break;
                case 1:
                    table = new Database().SelectData("EXEC DangNhap_GV '" + txtDangNhap.Text + "', '" + txtMatKhau.Text + "'");
                    break;
                case 2:
                    table = new Database().SelectData("SELECT * FROM TAIKHOANQTV WHERE TaiKhoan = '" + txtDangNhap.Text + "' AND MatKhau = '" + txtMatKhau.Text + "'");
                    break;
            }
            if (table != null)
            {
                if (table.Rows.Count != 0)
                {
                    this.Hide();
                    //Truyền tên đăng nhập, mật khẩu và mã giáo viên vào main form
                    if (cmbLoaiTK.SelectedIndex == 2)
                    {
                        new MainForm(txtDangNhap.Text, txtMatKhau.Text, null, cmbLoaiTK.SelectedIndex).ShowDialog();
                    }
                    else
                    {
                        new MainForm(txtDangNhap.Text, txtMatKhau.Text, table.Rows[0][2].ToString(), cmbLoaiTK.SelectedIndex).ShowDialog();
                    }
                    this.Close();
                }
                else
                {
                    lblSaiTKorMK.Visible = true;
                }
            }
        }

        private void keyDownCtrlBack(TextBox textBox, KeyEventArgs e)
        {
            if (e.Modifiers == Keys.Control && e.KeyCode == Keys.Back)
            {
                keyFlag = true;
                textBox.Text = "";
            }
            else
            {
                keyFlag = false;
            }
        }

        private void keyPressCtrlBack(TextBox textBox, KeyPressEventArgs e)
        {
            if (keyFlag)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }

        private void txtDangNhap_KeyDown(object sender, KeyEventArgs e)
        {
            keyDownCtrlBack(txtDangNhap, e);
        }

        private void txtMatKhau_KeyDown(object sender, KeyEventArgs e)
        {
            keyDownCtrlBack(txtMatKhau, e);
        }

        private void txtMatKhau_KeyPress(object sender, KeyPressEventArgs e)
        {
            keyPressCtrlBack(txtMatKhau, e);
        }

        private void txtDangNhap_KeyPress(object sender, KeyPressEventArgs e)
        {
            keyPressCtrlBack(txtDangNhap, e);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (mov)// Nêu chuột còn đang được ấn xuống và di chuyển
            {
                this.SetDesktopLocation(MousePosition.X - movX, MousePosition.Y - movY);
                //MousePosition là vị trí của chuột so với màn hình 
            }
        }
    }
}
