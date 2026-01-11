using projectQLTV.View;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace projectQLTV
{
    public partial class TrangChu : Form
    {
        public TrangChu()
        {
            InitializeComponent();
        }

        private void guna2Button1_Click(object sender, EventArgs e)
        {

        }


        private void TrangChu_Load(object sender, EventArgs e)
        {

        }
        private Form currentFormChild;
        private void OpenChildForm(Form childForm)
        {
            if (currentFormChild != null)
            {
                currentFormChild.Close();
            }
            currentFormChild = childForm;
            childForm.TopLevel = false;
            pntlContent.Controls.Add(childForm);
            pntlContent.Tag = childForm;
            childForm.BringToFront();
            childForm.Show();

        }

        private void btnQLSach_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormSach());
        }

        private void btnQLTheLoai_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormTheLoai());
        }

        private void btnQLTacGia_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormTacGia());
        }

        private void btnQLNXB_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormNXB());

        }

        private void btnQLNgonNgu_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormNgonNgu());

        }

        private void btnQLDocGia_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormDocGia());
        }

        private void btnQLNhanVien_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormNhanVien());
        }

        private void btnQLMuonTra_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormMuonTra());
        }

        private void btnQLKeSach_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormKeSach());
        }

        private void btnQLKhoa_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormKhoa());
        }

        private void btnQLLop_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormLop());
        }

        private void btnQLTheThuVien_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormTheThuVien());

        }

        private void btnThongKe_Click(object sender, EventArgs e)
        {
            OpenChildForm(new FormThongKe());

        }

        private void guna2Button1_Click_1(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Bạn có chắc muốn đăng xuất?", "Xác nhận",MessageBoxButtons.YesNo,MessageBoxIcon.Question);

            if (dr == DialogResult.Yes)
            {
                
                DangNhap formdangnhap = new DangNhap();
                formdangnhap.Show();
                this.Close();
            }
        }
    }
}
