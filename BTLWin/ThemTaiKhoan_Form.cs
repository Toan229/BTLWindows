﻿using System;
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
    public partial class ThemTaiKhoan_Form : Form
    {
        int accountType;
        List<string> tdnCSDL, maCSDL;
        QuanLyTaiKhoan_Form parent;
        public ThemTaiKhoan_Form(int accountType, List<string> tdnCSDL, List<string> maCSDL)
        {
            InitializeComponent();
            this.accountType = accountType;
            this.tdnCSDL = tdnCSDL;
            if (accountType != 2)
            {
                this.maCSDL = maCSDL;
            }
        }

        private void them(Form parent)
        {
            this.parent = (QuanLyTaiKhoan_Form)parent;
            int rowEffected = 0;
            string tdn1 = txtTenDN.Text.ToString().Trim();
            string mk1 = txtMatKhau.Text.ToString().Trim();
            string ma1 = txtMa.Text.ToString().Trim();
            switch (accountType)
            {
                case 0:
                    rowEffected = new Data.Database().ExecCmd("INSERT INTO TAIKHOANSV VALUES('" + tdn1 + "', '" + mk1 + "', '" + ma1 + "')");
                    break;
                case 1:
                    rowEffected = new Data.Database().ExecCmd("INSERT INTO TAIKHOANGV VALUES('" + tdn1 + "', '" + mk1 + "', '" + ma1 + "')");
                    break;
                case 2:
                    rowEffected = new Data.Database().ExecCmd("INSERT INTO TAIKHOANQTV VALUES('" + tdn1 + "', '" + mk1 + "')");
                    break;
            }
            if (rowEffected != 0)
            {
                this.parent.themTaiKhoan();
                tdnCSDL.Add(tdn1);
                if (txtMa.Visible == true)
                {
                    maCSDL.Add(ma1);
                }
                txtTenDN.Text = string.Empty;
                txtMatKhau.Text = string.Empty;
                txtMa.Text = string.Empty;
            }
            else
            {
                MessageBox.Show("Thêm tài khoản thất bại.\nHãy chắc chắn mã người dùng bạn nhập là đúng", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            them(((MainForm)this.Owner).currentChildForm);
        }

        private void ThemTaiKhoan_Form_Load(object sender, EventArgs e)
        {
            if (accountType == 2)
            {
                lblMa.Visible = false;
                txtMa.Visible = false;
                panel3.Visible = false;
            }
            if(accountType == 0)
            {
                lblMa.Text = "Mã sinh viên";
            }    
        }

        private void button1_Click(object sender, EventArgs e)
        {
            txtTenDN.Text = string.Empty;
            txtMatKhau.Text = string.Empty;
            txtMa.Text = string.Empty;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void txtTenDN_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtTenDN.Text.ToString()))
            {
                txtTenDN.Focus();
                errorProvider1.SetError(txtTenDN, "Tên đăng nhập đang bị bỏ trống");
            }
            else
            {
                if (tdnCSDL.Contains(txtTenDN.Text.ToString().Trim()))
                {
                    txtTenDN.Focus();
                    errorProvider1.SetError(txtTenDN, "Tên đăng nhập này đã tồn tại");
                }
                else
                {
                    errorProvider1.SetError(txtTenDN, string.Empty);
                    e.Cancel = false;
                }

            }
        }

        private void txtMatKhau_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtMatKhau.Text.ToString()))
            {
                txtMatKhau.Focus();
                errorProvider2.SetError(txtMatKhau, "Mật khẩu đang bị bỏ trống");
            }
            else
            {
                errorProvider2.SetError(txtMatKhau, string.Empty);
                e.Cancel = false;
            }
        }

        private void txtMa_Validating(object sender, CancelEventArgs e)
        {
            if (string.IsNullOrEmpty(txtMa.Text.ToString()))
            {
                txtMa.Focus();
                errorProvider3.SetError(txtMa, "Mã người dùng đang bị bỏ trống");
            }
            else
            {
                if (maCSDL.Contains(txtMa.Text.ToString().Trim()))
                {
                    txtMa.Focus();
                    errorProvider3.SetError(txtMa, "Người dùng này đã có tài khoản");
                }
                else
                {
                    errorProvider3.SetError(txtMa, string.Empty);
                    e.Cancel = false;
                }

            }
        }
    }
}
