using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ex_cel = Microsoft.Office.Interop.Excel;
using xls = Microsoft.Office.Interop.Excel;

namespace projectQLTV.View
{
    public partial class FormDocGia : Form
    {
        public FormDocGia()
        {
            InitializeComponent();
            this.Load += QuanLyDocGia_Load;
        }

        private void load_DocGia()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "SELECT DocGia.*, Khoa.TenKhoa, Lop.TenLop FROM DocGia " +
                         "JOIN Khoa ON DocGia.MaKhoa = Khoa.MaKhoa " +
                         "JOIN Lop ON DocGia.MaLop = Lop.MaLop";
            SetTrangThaiButton(false);  
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            txtMaDG.Text = TaoMaDGMoi();  
            thuvien.Hienthi(dgvQuanLyDocGia, cmd); 
        }
        private void QuanLyDocGia_Load(object sender, EventArgs e)
        {
            thuvien.hienthicbo(cboKhoa, "Khoa", "MaKhoa", "TenKhoa");
            thuvien.hienthicbo(cboLop, "Lop", "MaLop", "TenLop");

            dgvQuanLyDocGia.AutoGenerateColumns = false;
            load_DocGia();
        }
        private bool KiemTraDuLieu()
        {
            if (string.IsNullOrWhiteSpace(txtTenDG.Text))
            {
                MessageBox.Show("Vui lòng nhập tên độc giả");
                txtTenDG.Focus();
                return false;
            }

            if (cboKhoa.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn khoa");
                cboKhoa.Focus();
                return false;
            }

            if (cboLop.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn lớp");
                cboLop.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtEmail.Text))
            {
                MessageBox.Show("Vui lòng nhập email");
                txtEmail.Focus();
                return false;
            }

            if (!txtEmail.Text.Contains("@"))
            {
                MessageBox.Show("Email không hợp lệ");
                txtEmail.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSdt.Text))
            {
                MessageBox.Show("Vui lòng nhập số điện thoại");
                txtSdt.Focus();
                return false;
            }

            if (txtSdt.Text.Length != 10 || !long.TryParse(txtSdt.Text.Trim(), out _))
            {
                MessageBox.Show("Số điện thoại phải có 10 chữ số");
                txtSdt.Focus();
                return false;
            }

            if (dtpNgaySinh.Value == null)
            {
                MessageBox.Show("Vui lòng nhập ngày sinh");
                dtpNgaySinh.Focus();
                return false;
            }

            if (dtpNgaySinh.Value.Year < 1900 || dtpNgaySinh.Value.Year > DateTime.Now.Year)
            {
                MessageBox.Show("Năm sinh không hợp lệ");
                dtpNgaySinh.Focus();
                return false;
            }


            return true;
        }
        public static string TaoMaDGMoi()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "Select top 1 MaDG from DocGia order by MaDG DESC";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            object result = cmd.ExecuteScalar();
            thuvien.con.Close();

            if (result == null)
            {
                return "DG001";
            }
            else
            {
                string macu = result.ToString();
                int id = int.Parse(macu.Substring(2)) + 1;
                return "DG" + id.ToString("D3");
            }
        }


        private void btnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string trungma = "SELECT Count(*) FROM DocGia WHERE MaDG LIKE @madg";
            SqlCommand kiemtra = new SqlCommand(trungma, thuvien.con);
            kiemtra.Parameters.AddWithValue("@madg", txtMaDG.Text.Trim());
            int kq = (int)kiemtra.ExecuteScalar();

            if (kq > 0)
            {
                MessageBox.Show("Trùng mã độc giả, vui lòng nhập mã khác!");
                return;
            }

            string sql = @"INSERT INTO DocGia(MaDG, TenDG, MaKhoa, MaLop, GioiTinh, DiaChi, Email, Sdt, NgaySinh)
                           VALUES (@madg, @tendg, @makhoa, @malop, @gioitinh, @diachi, @email, @sdt, @ngaysinh)";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@madg", txtMaDG.Text.Trim());
            cmd.Parameters.AddWithValue("@tendg", txtTenDG.Text.Trim());
            cmd.Parameters.AddWithValue("@makhoa", cboKhoa.SelectedValue);
            cmd.Parameters.AddWithValue("@malop", cboLop.SelectedValue);
            cmd.Parameters.AddWithValue("@gioitinh", cboGioiTinh.SelectedItem.ToString());
            cmd.Parameters.AddWithValue("@diachi", txtDiaChi.Text.Trim());
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim());
            cmd.Parameters.AddWithValue("@sdt", txtSdt.Text.Trim());
            cmd.Parameters.AddWithValue("@ngaysinh", dtpNgaySinh.Value); // Lấy giá trị từ DateTimePicker

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Thêm độc giả thành công");
            reset();
            load_DocGia();
        }

        private void dgvQuanLyDocGia_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            int i = e.RowIndex;
            txtMaDG.Text = dgvQuanLyDocGia.Rows[i].Cells["MaDG"].Value.ToString();
            txtTenDG.Text = dgvQuanLyDocGia.Rows[i].Cells["TenDG"].Value.ToString();
            cboKhoa.SelectedValue = dgvQuanLyDocGia.Rows[i].Cells["MaKhoa"].Value.ToString();
            cboLop.SelectedValue = dgvQuanLyDocGia.Rows[i].Cells["MaLop"].Value.ToString();
            cboGioiTinh.SelectedItem = dgvQuanLyDocGia.Rows[i].Cells["GioiTinh"].Value.ToString();
            txtDiaChi.Text = dgvQuanLyDocGia.Rows[i].Cells["DiaChi"].Value.ToString();
            txtEmail.Text = dgvQuanLyDocGia.Rows[i].Cells["Email"].Value.ToString();
            txtSdt.Text = dgvQuanLyDocGia.Rows[i].Cells["Sdt"].Value.ToString();
            dtpNgaySinh.Value = Convert.ToDateTime(dgvQuanLyDocGia.Rows[i].Cells["NgaySinh"].Value); 
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaDG.Text))
            {
                MessageBox.Show("Vui lòng chọn độc giả cần xóa.");
                return;
            }

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string madg = txtMaDG.Text.Trim();
            string sql = "DELETE FROM DocGia WHERE MaDG = @madg";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@madg", madg);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Xóa độc giả thành công.");
            reset();
            load_DocGia();
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaDG.Text))
            {
                MessageBox.Show("Vui lòng chọn độc giả cần sửa");
                return;
            }

            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = @"UPDATE DocGia SET
                           TenDG = @tendg,
                           MaKhoa = @makhoa,
                           MaLop = @malop,
                           GioiTinh = @gioitinh,
                           DiaChi = @diachi,
                           Email = @email,
                           Sdt = @sdt,
                           NgaySinh = @ngaysinh
                           WHERE MaDG = @madg";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@madg", txtMaDG.Text.Trim());
            cmd.Parameters.AddWithValue("@tendg", txtTenDG.Text.Trim());
            cmd.Parameters.AddWithValue("@makhoa", cboKhoa.SelectedValue);
            cmd.Parameters.AddWithValue("@malop", cboLop.SelectedValue);
            cmd.Parameters.AddWithValue("@gioitinh", cboGioiTinh.SelectedItem.ToString());
            cmd.Parameters.AddWithValue("@diachi", txtDiaChi.Text.Trim());
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim());
            cmd.Parameters.AddWithValue("@sdt", txtSdt.Text.Trim());
            cmd.Parameters.AddWithValue("@ngaysinh", dtpNgaySinh.Value);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Cập nhật độc giả thành công");
            reset();
            load_DocGia();
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            load_DocGia();
            reset();
            SetTrangThaiButton(false);
        }
        private void reset()
        {
            txtMaDG.Text = TaoMaDGMoi();
            txtTenDG.Clear();
            txtDiaChi.Clear();
            txtEmail.Clear();
            txtSdt.Clear();
            dtpNgaySinh.Value = DateTime.Now; 
            txtTimKiem.Clear();
            cboKhoa.SelectedIndex = -1;
            cboLop.SelectedIndex = -1;
            cboGioiTinh.SelectedIndex = -1;

            dgvQuanLyDocGia.ClearSelection();
            dgvQuanLyDocGia.CurrentCell = null;

            SetTrangThaiButton(false);
        }
        private void SetTrangThaiButton(bool enabled)
        {
            btnSua.Enabled = enabled;
            btnXoa.Enabled = enabled;
        }

        private void btnNhap_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.Multiselect = false;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                ReadExcel(filePath);
                load_DocGia();
            }
        }
        private void ReadExcel(string filename)
        {
            if (filename == null)
            {
                MessageBox.Show("Chưa chọn file");
            }
            else
            {
                xls.Application excelApp = new xls.Application();
                excelApp.Workbooks.Open(filename);

                foreach (xls.Worksheet wsheet in excelApp.Worksheets)
                {
                    int row = 4;
                    do
                    {
                        if (wsheet.Cells[row, 1].Value == null && wsheet.Cells[row, 2].Value == null && wsheet.Cells[row, 3].Value == null)
                        {
                            break;
                        }
                        else
                        {
                            ThemDocGia(
                                wsheet.Cells[row, 1].Value.ToString(),
                                wsheet.Cells[row, 2].Value.ToString(),
                                wsheet.Cells[row, 3].Value.ToString(),
                                wsheet.Cells[row, 4].Value.ToString(),
                                wsheet.Cells[row, 5].Value.ToString(),
                                wsheet.Cells[row, 6].Value.ToString(),
                                wsheet.Cells[row, 7].Value.ToString(),
                                wsheet.Cells[row, 8].Value.ToString(),
                                Convert.ToDateTime(wsheet.Cells[row, 9].Value)
                            );
                            row++;
                        }
                    }
                    while (true);
                }
                MessageBox.Show("Nhập Excel thành công");
            }
        }
        private void ThemDocGia(string madg, string tendg, string makhoa, string malop, string gioitinh, string diachi, string email, string sdt, DateTime ngaysinh)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = @"INSERT INTO DocGia(MaDG, TenDG, MaKhoa, MaLop, GioiTinh, DiaChi, Email, Sdt, NgaySinh) 
                   VALUES (@madg, @tendg, @makhoa, @malop, @gioitinh, @diachi, @email, @sdt, @ngaysinh)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@madg", madg);
            cmd.Parameters.AddWithValue("@tendg", tendg);
            cmd.Parameters.AddWithValue("@makhoa", makhoa);
            cmd.Parameters.AddWithValue("@malop", malop);
            cmd.Parameters.AddWithValue("@gioitinh", gioitinh);
            cmd.Parameters.AddWithValue("@diachi", diachi);
            cmd.Parameters.AddWithValue("@email", email);
            cmd.Parameters.AddWithValue("@sdt", sdt);
            cmd.Parameters.AddWithValue("@ngaysinh", ngaysinh);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();




            string sql = "SELECT dg.MaDG, dg.TenDG, k.TenKhoa, l.TenLop, dg.NgaySinh, dg.GioiTinh, dg.DiaChi, dg.Email, dg.Sdt " +
                         "FROM DocGia dg " +
                         "JOIN Khoa k ON dg.MaKhoa = k.MaKhoa " +
                         "JOIN Lop l ON dg.MaLop = l.MaLop";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;

            DataTable tb = new DataTable();
            da.Fill(tb);

            cmd.Dispose();
            thuvien.con.Close();

            ExportExcel(tb, "Danh Sách Độc Giả");
        }
        public void ExportExcel(DataTable tb, string sheetname)
        {

        }





        private void txtTimKiem_TextChanged(object sender, EventArgs e)
        {

        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string keyword = txtTimKiem.Text.Trim();
            string sql = "SELECT * FROM DocGia WHERE MaDG LIKE @keyword OR TenDG LIKE @keyword OR TenKhoa LIKE @keyword OR TenLop LIKE @keyword OR Email LIKE @keyword";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);

            dgvQuanLyDocGia.DataSource = dt;
            thuvien.con.Close();
        }
    }
}
