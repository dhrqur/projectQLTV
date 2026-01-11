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
using tacgia = Microsoft.Office.Interop.Excel;
using ex_cel = Microsoft.Office.Interop.Excel;


namespace projectQLTV.View
{
    public partial class FormTacGia : Form
    {
        public FormTacGia()
        {
            InitializeComponent();
            load_TacGia();
            btnxoa.Enabled = false;
            btnsua.Enabled = false;
            txtMaTG.Enabled = false;
            txtMaTG.Text = TaoMaTG();
        }
        SqlConnection con = thuvien.con;
        public void connect()
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
        }
        private void load_TacGia()
        {
            connect();
            string sql = "SELECT * FROM TacGia";
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            cmd.Dispose();
            con.Close();
            dgvTacGia.DataSource = dt;
        }
        private bool checktrungmatacgia(string matg)
        {
            connect();
            string sql = "select count(*) From TacGia where MaTG ='" + matg + "'";
            SqlCommand cmd = new SqlCommand(sql, con);
            int kq = (int)cmd.ExecuteScalar();
            cmd.Dispose();
            con.Close();
            if (kq > 0)
                return true;
            else
                return false;
        }


        string TaoMaTG()
        {
            connect();
            string sql = "SELECT TOP 1 MaTG FROM TacGia ORDER BY MaTG DESC";
            SqlCommand cmd = new SqlCommand(sql, con);
            object kq = cmd.ExecuteScalar();
            cmd.Dispose();
            con.Close();

            if (kq == null)
                return "TG001";

            string maCu = kq.ToString();
            int so = int.Parse(maCu.Substring(2));
            so++;
            return "TG" + so.ToString("000");
        }

        private void guna2Button3_Click(object sender, EventArgs e)
        {
            string MaTG = TaoMaTG();
            txtMaTG.Text = MaTG;
            string TenTG = txtTenTG.Text.Trim();
            string NamSinh = txtNamSinh.Text.Trim();
            string GioiTinh = cboGioiTinh.SelectedItem.ToString();
            string QuocTich = txtQuocTich.Text.Trim();

            if (checktrungmatacgia(MaTG))
            {
                txtMaTG.Focus();
                MessageBox.Show("Mã tác giả đã tồn tại!");
                return;
            }
            try
            {
                int ns = int.Parse(NamSinh);
                connect();
                string sql = "INSERT INTO TacGia (MaTG, TenTG, NamSinh, GioiTinh, QuocTich) " +
                             "VALUES ('" + MaTG + "', N'" + TenTG + "', " + NamSinh + ", N'" + GioiTinh + "', N'" + QuocTich + "')";
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();
                MessageBox.Show("Thêm thành công!");
                load_TacGia();
            }
            catch
            {
                txtNamSinh.Focus();
                MessageBox.Show("Năm sinh phải là số nguyên!");
                return;
            }
            
        }

        private void dgvTacGia_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int i = dgvTacGia.CurrentCell.RowIndex;
            btnthem.Enabled = false;
            btnxoa.Enabled = true;
            btnsua.Enabled = true;
            txtMaTG.Enabled = false;
            txtMaTG.Text = dgvTacGia.Rows[i].Cells[0].Value.ToString();
            txtTenTG.Text = dgvTacGia.Rows[i].Cells[1].Value.ToString();
            txtNamSinh.Text = dgvTacGia.Rows[i].Cells[2].Value.ToString();
            cboGioiTinh.Text = dgvTacGia.Rows[i].Cells[3].Value.ToString();
            txtQuocTich.Text = dgvTacGia.Rows[i].Cells[4].Value.ToString();
        }

        private void guna2Button5_Click(object sender, EventArgs e)
        {
            string MaTG = txtMaTG.Text.Trim();
            if (con.State == ConnectionState.Closed)
                con.Open();
            try {
                string sql = "DELETE FROM TacGia WHERE MaTG='" + MaTG + "'";
                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();
                MessageBox.Show("Xóa thành công");
                load_TacGia();
            }
            catch
            {
                MessageBox.Show("Lỗi xóa dữ liệu!");
                return;
            }
            
        }

        private void btnlammoi_Click(object sender, EventArgs e)
        {
            txtMaTG.Enabled = false;
            txtMaTG.Enabled = false;
            btnthem.Enabled = true;
            btnxoa.Enabled = false;
            btnsua.Enabled = false;
            txtMaTG.Text = TaoMaTG();
            txtTenTG.Clear();
            txtNamSinh.Clear();
            txtQuocTich.Clear();
            cboGioiTinh.SelectedIndex = -1;
            load_TacGia();
        }

        private void btnsua_Click(object sender, EventArgs e)
        {
            txtMaTG.Enabled = false;
            string MaTG = txtMaTG.Text.Trim();
            string TenTG = txtTenTG.Text.Trim();
            string NamSinh = txtNamSinh.Text.Trim();
            string GioiTinh = cboGioiTinh.SelectedItem.ToString();
            string QuocTich = txtQuocTich.Text.Trim();

            if (con.State == ConnectionState.Closed)
                con.Open();

            string sql = "UPDATE TacGia SET " +" TenTG = N'" + TenTG + "', " +" NamSinh = " + NamSinh + ", " +" GioiTinh = N'" + GioiTinh + "', " +" QuocTich = N'" + QuocTich + "' " +" WHERE MaTG = '" + MaTG + "'";

            SqlCommand cmd = new SqlCommand(sql, con);
            cmd.ExecuteNonQuery();
            cmd.Dispose();
            con.Close();
            MessageBox.Show("Sửa thành công");
            load_TacGia();
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            string tim = txttim.Text.Trim();
            if (con.State == ConnectionState.Closed)
                con.Open();

            string sql = "SELECT * FROM TacGia WHERE MaTG LIKE '%" + tim + "%' OR TenTG LIKE N'%" + tim + "%' OR GioiTinh LIKE N'%" + tim + "%' OR QuocTich LIKE N'%" + tim + "%'";
            SqlCommand cmd = new SqlCommand(sql, con);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable dt = new DataTable();
            da.Fill(dt);
            cmd.Dispose();
            con.Close();
            dgvTacGia.DataSource = dt;
        }
        public string filename = "";

        private void btnnhapexcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Excel Files|*.xlsx;*.xls";

            if (ofd.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Chưa chọn file");
                return;
            }

            filename = ofd.FileName;

            tacgia.Application Excel = new tacgia.Application();
            Excel.Workbooks.Open(filename);

            foreach (tacgia.Worksheet wsheet in Excel.Worksheets)
            {
                int i = 2;
                do
                {
                    if (wsheet.Cells[i, 1].Value == null)
                        break;

                    string MaTG = wsheet.Cells[i, 1].Value.ToString();
                    string TenTG = wsheet.Cells[i, 2].Value.ToString();
                    string NamSinh = wsheet.Cells[i, 3].Value.ToString();
                    string GioiTinh = wsheet.Cells[i, 4].Value.ToString();
                    string QuocTich = wsheet.Cells[i, 5].Value.ToString();

                    ThemmoiTacGia(MaTG, TenTG, NamSinh, GioiTinh, QuocTich);
                    i++;
                }
                while (true);
            }
            Excel.Quit();
            load_TacGia();
        }
        private void ThemmoiTacGia(string MaTG, string TenTG, string NamSinh, string GioiTinh, string QuocTich)
        {
            if (con.State == ConnectionState.Closed)
                con.Open();
            try
            {
                string sql = @"INSERT INTO TacGia(MaTG, TenTG, NamSinh, GioiTinh, QuocTich) VALUES (@MaTG, @TenTG, @NamSinh, @GioiTinh, @QuocTich)";

                SqlCommand cmd = new SqlCommand(sql, con);
                cmd.Parameters.AddWithValue("@MaTG", MaTG);
                cmd.Parameters.AddWithValue("@TenTG", TenTG);
                cmd.Parameters.AddWithValue("@NamSinh", NamSinh);
                cmd.Parameters.AddWithValue("@GioiTinh", GioiTinh);
                cmd.Parameters.AddWithValue("@QuocTich", QuocTich);

                cmd.ExecuteNonQuery();
                cmd.Dispose();
                con.Close();
                MessageBox.Show("Thêm tác giả từ Excel thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi: " + ex.Message);
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

            oSheet.get_Range("A3").Value2 = "Mã tác giả";
            oSheet.get_Range("B3").Value2 = "Tên tác giả";
            oSheet.get_Range("C3").Value2 = "Năm sinh";
            oSheet.get_Range("D3").Value2 = "Giới tính";
            oSheet.get_Range("E3").Value2 = "Quốc tịch";

            object[,] arr = new object[tb.Rows.Count, tb.Columns.Count];
            for (int r = 0; r < tb.Rows.Count; r++)
            {
                DataRow dr = tb.Rows[r];
                for (int c = 0; c < tb.Columns.Count; c++)
                    arr[r, c] = dr[c];
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
            DataTable dt = (DataTable)dgvTacGia.DataSource;
            ExportExcel(dt, "TacGiaExcel");
        }
    }
}
