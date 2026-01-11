using System;
using System.Data;
using System.Data.SqlClient;
using System.Text;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using nhanvien = Microsoft.Office.Interop.Excel;
using ex_cel = Microsoft.Office.Interop.Excel;


namespace projectQLTV.View
{
    public partial class FormNhanVien : Form
    {
        public FormNhanVien()
        {
            InitializeComponent();
            load_NhanVien();
        }
        SqlConnection con = thuvien.con;
        public void connect()
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
        }

        private void load_NhanVien()
        {
            connect();
            string sql = "SELECT * FROM NhanVien";
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            cmd.Dispose();
            con.Close();
            dgvNhanVien.DataSource = dt;
        }
        private bool checktrungmanhanvien(string manv)
        {
            connect();
            string sql = "select count(*) From NhanVien where  MaNV ='" + manv + "'";
            SqlCommand cmd = new SqlCommand(sql, con);
            int kq = (int)cmd.ExecuteScalar();
            cmd.Dispose();
            con.Close();
            if (kq > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        string MaHoa(string mk)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(mk));
        }

        

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            string MaNV = txtMaNV.Text.Trim();
            string TenNV = txtTenNV.Text.Trim();
            string QueQuan = txtQueQuan.Text.Trim();
            string GioiTinh = cboGioiTinh.SelectedItem.ToString();
            string VaiTro = cboVaiTro.SelectedItem.ToString();
            string DiaChi = cboDiaChi.SelectedItem.ToString();
            string Email = txtEmail.Text.Trim();
            string Sdt = txtSdt.Text.Trim();
            string Tendangnhap = txtTendangnhap.Text.Trim();
            string Matkhau = MaHoa(txtMatkhau.Text.Trim());

            if (checktrungmanhanvien(MaNV))
            {
                txtMaNV.Focus();
                MessageBox.Show("Mã nhân viên đã tồn tại, vui lòng nhập mã khác!");
                return;
            }

            connect();
            string sql = "INSERT INTO NhanVien (MaNV, TenNV, QueQuan, GioiTinh, VaiTro, DiaChi, Email, Sdt, Tendangnhap, Matkhau) " +
                         "VALUES ('" + MaNV + "', N'" + TenNV + "', N'" + QueQuan + "', N'" + GioiTinh + "', N'" + VaiTro + "', N'" + DiaChi + "', '" + Email + "', '" + Sdt + "', N'" + Tendangnhap + "', N'" + Matkhau + "')";
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            con.Close();
            MessageBox.Show("Thêm mới thành công!");
            load_NhanVien();
        }

        private void dgvnhanvien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dgvNhanVien.CurrentCell.RowIndex;

            txtMaNV.Text = dgvNhanVien.Rows[i].Cells[0].Value.ToString();
            txtTenNV.Text = dgvNhanVien.Rows[i].Cells[1].Value.ToString();
            txtQueQuan.Text = dgvNhanVien.Rows[i].Cells[2].Value.ToString();
            cboGioiTinh.Text = dgvNhanVien.Rows[i].Cells[3].Value.ToString();
            cboVaiTro.Text = dgvNhanVien.Rows[i].Cells[4].Value.ToString();
            cboDiaChi.Text = dgvNhanVien.Rows[i].Cells[5].Value.ToString();
            txtEmail.Text = dgvNhanVien.Rows[i].Cells[6].Value.ToString();
            txtSdt.Text = dgvNhanVien.Rows[i].Cells[7].Value.ToString();
            txtTendangnhap.Text = dgvNhanVien.Rows[i].Cells[8].Value.ToString();
            txtMatkhau.Text = dgvNhanVien.Rows[i].Cells[9].Value.ToString();
        }

        private void btn_xoa_Click(object sender, EventArgs e)
        {
            string MaNV =txtMaNV.Text.Trim();
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            string sql = "DELETE FROM NhanVien WHERE MaNV='" + MaNV + "'";
            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            con.Close();
            MessageBox.Show("Xóa thành công");
            load_NhanVien();

        }

        private void btn_sua_Click(object sender, EventArgs e)
        {
            string MaNV = txtMaNV.Text.Trim();
            string TenNV = txtTenNV.Text.Trim();
            string QueQuan = txtQueQuan.Text.Trim();
            string GioiTinh = cboGioiTinh.SelectedItem.ToString();
            string VaiTro = cboVaiTro.SelectedItem.ToString();
            string DiaChi = cboDiaChi.SelectedItem.ToString();
            string Email = txtEmail.Text.Trim();
            string Sdt = txtSdt.Text.Trim();
            string Tendangnhap = txtTendangnhap.Text.Trim();
            string Matkhau = MaHoa(txtMatkhau.Text.Trim());

            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }

            string sql = "UPDATE NhanVien SET " +"TenNV = N'" + TenNV +"', QueQuan = N'" + QueQuan +"', GioiTinh = N'" + GioiTinh +"', VaiTro = N'" + VaiTro +"', DiaChi = N'" + DiaChi +"', Email = N'" + Email +"', Sdt = '" + Sdt +"', Tendangnhap = N'" + Tendangnhap +"', Matkhau = N'" + Matkhau +"' WHERE MaNV = '" + MaNV + "'";

            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            con.Close();
            MessageBox.Show("Sửa thành công");
            load_NhanVien();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            string tim=txtTimkiem.Text.Trim();
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            string sql = "SELECT * FROM NhanVien WHERE MaNV LIKE '%" + tim + "%' OR TenNV LIKE N'%" + tim + "%' OR QueQuan LIKE N'%" + tim + "%' OR GioiTinh LIKE N'%" + tim + "%' OR VaiTro LIKE N'%" + tim + "%' OR DiaChi LIKE N'%" + tim + "%' OR Email LIKE '%" + tim + "%' OR Sdt LIKE '%" + tim + "%' OR Tendangnhap LIKE N'%" + tim + "%'";
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            cmd.Dispose();
            con.Close();
            dgvNhanVien.DataSource = dt;
        }
        public string filename = "";
        private void guna2Button7_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls";

            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Chưa chọn file");
                return;
            }

            filename = ofd.FileName;

            nhanvien.Application Excel = new nhanvien.Application();
            Excel.Workbooks.Open(filename);

            foreach (nhanvien.Worksheet wsheet in Excel.Worksheets)
            {
                int i = 2;
                do
                {
                    if (wsheet.Cells[i, 1].Value == null)
                        break;

                    string MaNV = wsheet.Cells[i, 1].Value.ToString();
                    string TenNV = wsheet.Cells[i, 2].Value.ToString();
                    string QueQuan = wsheet.Cells[i, 3].Value.ToString();
                    string GioiTinh = wsheet.Cells[i, 4].Value.ToString();
                    string VaiTro = wsheet.Cells[i, 5].Value.ToString();
                    string DiaChi = wsheet.Cells[i, 6].Value.ToString();
                    string Email = wsheet.Cells[i, 7].Value.ToString();
                    string Sdt = wsheet.Cells[i, 8].Value.ToString();
                    string Tendangnhap = wsheet.Cells[i, 9].Value.ToString();
                    string Matkhau = wsheet.Cells[i, 10].Value.ToString();

                    ThemmoiNhanVien(MaNV, TenNV, QueQuan, GioiTinh, VaiTro, DiaChi, Email, Sdt, Tendangnhap, Matkhau);
                    i++;
                }
                while (true);
            }
        }
        private void ThemmoiNhanVien(string MaNV, string TenNV, string QueQuan, string GioiTinh, string VaiTro, string DiaChi, string Email, string Sdt, string Tendangnhap, string Matkhau)
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
            try
            {
                string sql = @"INSERT INTO NhanVien
                (MaNV, TenNV, QueQuan, GioiTinh, VaiTro, DiaChi, Email, Sdt, Tendangnhap, Matkhau)
                VALUES (@MaNV, @TenNV, @QueQuan, @GioiTinh, @VaiTro, @DiaChi, @Email, @Sdt, @Tendangnhap, @Matkhau)";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddWithValue("@MaNV", MaNV);
                cmd.Parameters.AddWithValue("@TenNV", TenNV);
                cmd.Parameters.AddWithValue("@QueQuan", QueQuan);
                cmd.Parameters.AddWithValue("@GioiTinh", GioiTinh);
                cmd.Parameters.AddWithValue("@VaiTro", VaiTro);
                cmd.Parameters.AddWithValue("@DiaChi", DiaChi);
                cmd.Parameters.AddWithValue("@Email", Email);
                cmd.Parameters.AddWithValue("@Sdt", Sdt);
                cmd.Parameters.AddWithValue("@Tendangnhap", Tendangnhap);
                cmd.Parameters.AddWithValue("@Matkhau", Matkhau);

                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();
                MessageBox.Show("Thêm nhân viên từ file Excel thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm nhân viên: " + ex.Message);
            }
        }
        public void ExportExcel(DataTable tb, string sheetname)
        {
            ex_cel.Application oExcel = new ex_cel.Application();
            ex_cel.Workbooks oBooks;
            ex_cel.Sheets oSheets;
            ex_cel.Workbook oBook;
            ex_cel.Worksheet oSheet;

            oExcel.Visible = true;
            oExcel.DisplayAlerts = false;
            oExcel.Application.SheetsInNewWorkbook = 1;
            oBooks = oExcel.Workbooks;
            oBook = (ex_cel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
            oSheets = oBook.Worksheets;
            oSheet = (ex_cel.Worksheet)oSheets.get_Item(1);
            oSheet.Name = sheetname;

            oSheet.get_Range("A3").Value2 = "Mã nhân viên";
            oSheet.get_Range("B3").Value2 = "Tên nhân viên";
            oSheet.get_Range("C3").Value2 = "Quê quán";
            oSheet.get_Range("D3").Value2 = "Giới tính";
            oSheet.get_Range("E3").Value2 = "Vai trò";
            oSheet.get_Range("F3").Value2 = "Địa chỉ";
            oSheet.get_Range("G3").Value2 = "Email";
            oSheet.get_Range("H3").Value2 = "Số điện thoại";
            oSheet.get_Range("I3").Value2 = "Tên đăng nhập";
            oSheet.get_Range("J3").Value2 = "Mật khẩu";

            oSheet.Columns[1].ColumnWidth = 15;
            oSheet.Columns[2].ColumnWidth = 25;
            oSheet.Columns[3].ColumnWidth = 20;
            oSheet.Columns[4].ColumnWidth = 15;
            oSheet.Columns[5].ColumnWidth = 20;
            oSheet.Columns[6].ColumnWidth = 30;
            oSheet.Columns[7].ColumnWidth = 30;
            oSheet.Columns[8].ColumnWidth = 20;
            oSheet.Columns[9].ColumnWidth = 25;
            oSheet.Columns[10].ColumnWidth = 25;

            object[,] arr = new object[tb.Rows.Count, tb.Columns.Count];
            for (int r = 0; r < tb.Rows.Count; r++)
            {
                DataRow dr = tb.Rows[r];
                for (int c = 0; c < tb.Columns.Count; c++)
                {
                    arr[r, c] = dr[c];
                }
            }

            int rowStart = 4;
            int columnStart = 1;
            int rowEnd = rowStart + tb.Rows.Count - 1;
            int columnEnd = tb.Columns.Count;

            ex_cel.Range c1 = (ex_cel.Range)oSheet.Cells[rowStart, columnStart];
            ex_cel.Range c2 = (ex_cel.Range)oSheet.Cells[rowEnd, columnEnd];
            ex_cel.Range range = oSheet.get_Range(c1, c2);

            range.Value2 = arr;
            range.Borders.LineStyle = ex_cel.Constants.xlSolid;
        }

        private void btnxuatexcel_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dgvNhanVien.DataSource;
            ExportExcel(dt, "NhanVienExcel");
            load_NhanVien();
        }
    }
}