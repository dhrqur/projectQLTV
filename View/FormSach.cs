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
    public partial class FormSach : Form
    {
        public FormSach()
        {
            InitializeComponent();
            this.Load += QuanLySach_Load;
        }

        private void load_Sach()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }
            string sql = "SELECT Sach.*, TheLoai.TenTL, NhaXuatBan.TenNXB, TacGia.TenTG, KeSach.TenKe, NgonNgu.TenNN " +
              "FROM Sach " +
              "JOIN TheLoai ON Sach.MaTL = TheLoai.MaTL " +
              "JOIN NhaXuatBan ON Sach.MaNXB = NhaXuatBan.MaNXB " +
              "JOIN TacGia ON Sach.MaTG = TacGia.MaTG " +
              "JOIN KeSach ON Sach.MaKe = KeSach.MaKe " +
              "JOIN NgonNgu ON Sach.MaNN = NgonNgu.MaNN";
            SetTrangThaiButton(false);
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            txtMaSach.Text = TaoMaSachMoi();
            thuvien.Hienthi(dgvQuanLySach, cmd);
        }

        private void QuanLySach_Load(object sender, EventArgs e)
        {

            thuvien.hienthicbo(cboTacGia, "TacGia", "maTG", "tenTG");

            thuvien.hienthicbo(cboTheLoai, "TheLoai", "MaTL", "TenTL");
            thuvien.hienthicbo(cboNhaXuatBan, "NhaXuatBan", "MaNXB", "TenNXB");
            thuvien.hienthicbo(cboNgonNgu, "NgonNgu", "MaNN", "TenNN");
            thuvien.hienthicbo(cboKeSach, "KeSach", "MaKe", "TenKe");
            dgvQuanLySach.AutoGenerateColumns = false;
            load_Sach();
        }
        private bool KiemTraDuLieu()
        {
            if (string.IsNullOrWhiteSpace(txtTenSach.Text))
            {
                MessageBox.Show("Vui lòng nhập tên sách");
                txtTenSach.Focus();
                return false;
            }

            if (cboTacGia.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn tác giả");
                cboTacGia.Focus();
                return false;
            }

            if (cboNgonNgu.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn ngôn ngữ");
                cboNgonNgu.Focus();
                return false;
            }

            if (cboKeSach.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn vị trí");
                cboKeSach.Focus();
                return false;
            }

            if (cboNhaXuatBan.SelectedIndex == -1)
            {
                MessageBox.Show("Vui lòng chọn nhà xuất bản");
                cboNhaXuatBan.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtNamXuatBan.Text))
            {
                MessageBox.Show("Vui lòng nhập năm xuất bản");
                txtNamXuatBan.Focus();
                return false;
            }

            if (!int.TryParse(txtNamXuatBan.Text.Trim(), out int namXB))
            {
                MessageBox.Show("Năm xuất bản phải là số");
                txtNamXuatBan.Focus();
                return false;
            }

            if (namXB < 1000 || namXB > DateTime.Now.Year)
            {
                MessageBox.Show("Năm xuất bản không hợp lệ");
                txtNamXuatBan.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSoLuong.Text))
            {
                MessageBox.Show("Vui lòng nhập số lượng");
                txtSoLuong.Focus();
                return false;
            }

            if (!int.TryParse(txtSoLuong.Text.Trim(), out int soLuong))
            {
                MessageBox.Show("Số lượng phải là số");
                txtSoLuong.Focus();
                return false;
            }

            if (soLuong < 0)
            {
                MessageBox.Show("Số lượng phải lớn hơn hoặc bằng 0");
                txtSoLuong.Focus();
                return false;
            }

            return true;
        }



        public static string TaoMaSachMoi()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }
            string sql = "Select top 1 MaSach from Sach order by MaSach DESC";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            object result = cmd.ExecuteScalar();
            thuvien.con.Close();
            if (result == null)
            {
                return "S001";
            }
            else
            {
                string macu = result.ToString();
                int id = int.Parse(macu.Substring(1)) + 1;
                return "S" + id.ToString("D3");
            }
        }
        private void guna2HtmlLabel3_Click(object sender, EventArgs e)
        {

        }

        private void guna2HtmlLabel7_Click(object sender, EventArgs e)
        {

        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string trungma = "Select Count(*) from Sach Where MaSach LIKE @masach";
            SqlCommand kiemtra = new SqlCommand(trungma, thuvien.con);
            kiemtra.Parameters.AddWithValue("@masach", txtMaSach.Text.Trim());
            int kq = (int)kiemtra.ExecuteScalar();
            if (kq > 0)
            {
                MessageBox.Show("Trùng mã sách, vui lòng bấm nút làm mới và nhập lại");
                return;
            }
            
            string sql = @"INSERT INTO Sach(MaSach, MaTG, MaNXB, MaTL, TenSach, NamXB, SoLuong, MoTa, MaNN, MaKe)
                   VALUES (@masach, @matg, @manxb, @matl, @tensach, @namxb, @soluong, @mota, @mann, @make)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@masach", txtMaSach.Text.Trim());
            cmd.Parameters.AddWithValue("@tensach", txtTenSach.Text.Trim());
            cmd.Parameters.AddWithValue("@matg", cboTacGia.SelectedValue);
            cmd.Parameters.AddWithValue("@manxb", cboNhaXuatBan.SelectedValue);
            cmd.Parameters.AddWithValue("@matl", cboTheLoai.SelectedValue);
            cmd.Parameters.AddWithValue("@namxb", int.Parse(txtNamXuatBan.Text));
            cmd.Parameters.AddWithValue("@soluong", int.Parse(txtSoLuong.Text));
            cmd.Parameters.AddWithValue("@mota", txtMoTa.Text.Trim());
            cmd.Parameters.AddWithValue("@mann", cboNgonNgu.SelectedValue);
            cmd.Parameters.AddWithValue("@make", cboKeSach.SelectedValue);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Thêm sách thành công");
            reset();
            load_Sach();
        }

        private void dgvQuanLySach_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            int i = e.RowIndex;
            txtMaSach.Text = dgvQuanLySach.Rows[i].Cells["MaSach"].Value.ToString();
            txtTenSach.Text = dgvQuanLySach.Rows[i].Cells["TenSach"].Value.ToString();
            cboTacGia.SelectedValue = dgvQuanLySach.Rows[i].Cells["MaTG"].Value.ToString();
            cboTheLoai.SelectedValue = dgvQuanLySach.Rows[i].Cells["MaTL"].Value.ToString();
            cboNhaXuatBan.SelectedValue = dgvQuanLySach.Rows[i].Cells["MaNXB"].Value.ToString();
            txtNamXuatBan.Text = dgvQuanLySach.Rows[i].Cells["NamXB"].Value.ToString();
            txtSoLuong.Text = dgvQuanLySach.Rows[i].Cells["SoLuong"].Value.ToString();
            txtMoTa.Text = dgvQuanLySach.Rows[i].Cells["MoTa"].Value.ToString();
            cboKeSach.SelectedValue = dgvQuanLySach.Rows[i].Cells["MaKe"].Value.ToString();
            cboNgonNgu.SelectedValue = dgvQuanLySach.Rows[i].Cells["MaNN"].Value.ToString();
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaSach.Text))
            {
                MessageBox.Show("Vui lòng chọn sách cần xóa.");
                return;
            }

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string masach = txtMaSach.Text.Trim();

            string sql = "SELECT COUNT(*) FROM ChiTietMuonTra WHERE MaSach = @masach";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@masach", masach);

            int dangMuon = (int)cmd.ExecuteScalar();

            if (dangMuon > 0)
            {
                MessageBox.Show(
                    "Không thể xóa sách vì sách đang được mượn.",
                    "Cảnh báo",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                thuvien.con.Close();
                return;
            }

            DialogResult kq = MessageBox.Show(
                "Bạn có chắc muốn xóa sách này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (kq == DialogResult.No)
            {
                thuvien.con.Close();
                return;
            }

            string sql1 = "DELETE FROM Sach WHERE MaSach = @masach";
            SqlCommand cmd1 = new SqlCommand(sql1, thuvien.con);
            cmd1.Parameters.AddWithValue("@masach", masach);
            cmd1.ExecuteNonQuery();
            cmd1.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Xóa sách thành công.");
            reset();
            load_Sach();
        }


        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaSach.Text))
            {
                MessageBox.Show("Vui lòng chọn sách cần sửa");
                return;
            }

            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = @"UPDATE Sach SET
                    TenSach = @tensach,
                    MaTG = @matg,
                    MaNXB = @manxb,
                    MaTL = @matl,
                    NamXB = @namxb,
                    SoLuong = @soluong,
                    MoTa = @mota,
                    MaNN = @mann,
                    MaKe = @make
                   WHERE MaSach = @masach";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@masach", txtMaSach.Text.Trim());
            cmd.Parameters.AddWithValue("@tensach", txtTenSach.Text.Trim());
            cmd.Parameters.AddWithValue("@matg", cboTacGia.SelectedValue);
            cmd.Parameters.AddWithValue("@manxb", cboNhaXuatBan.SelectedValue);
            cmd.Parameters.AddWithValue("@matl", cboTheLoai.SelectedValue);
            cmd.Parameters.AddWithValue("@namxb", int.Parse(txtNamXuatBan.Text));
            cmd.Parameters.AddWithValue("@soluong", int.Parse(txtSoLuong.Text));
            cmd.Parameters.AddWithValue("@mota", txtMoTa.Text.Trim());
            cmd.Parameters.AddWithValue("@mann", cboNgonNgu.SelectedValue);
            cmd.Parameters.AddWithValue("@make", cboKeSach.SelectedValue);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Cập nhật sách thành công");
            reset();
            load_Sach();
        }
        private void reset()
        {
            txtMaSach.Text = TaoMaSachMoi();
            txtTenSach.Clear();
            txtNamXuatBan.Clear();
            txtSoLuong.Clear();
            txtMoTa.Clear();
            txttimkiem.Clear();
            txtNamXuatBan.Clear();
            txtSoLuong.Clear();
            cboNhaXuatBan.SelectedIndex = -1;
            cboTacGia.SelectedIndex = -1;
            cboTheLoai.SelectedIndex = -1;
            cboKeSach.SelectedIndex = -1;
            cboNgonNgu.SelectedIndex = -1;

            dgvQuanLySach.ClearSelection();
            dgvQuanLySach.CurrentCell = null;

            SetTrangThaiButton(false);
        }
        private void btnlammoi_Click(object sender, EventArgs e)
        {
            load_Sach();
            reset();
            SetTrangThaiButton(false);
        }
        private void SetTrangThaiButton(bool enabled)
        {
            btnSua.Enabled = enabled;
            btnXoa.Enabled = enabled;
        }
        private void dgvQuanLySach_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        
        public void ExportExcel(DataTable tb, string sheetname)
        {
            //Tạo các đối tượng Excel
            ex_cel.Application oExcel = new ex_cel.Application();
            ex_cel.Workbooks oBooks;
            ex_cel.Sheets oSheets;
            ex_cel.Workbook oBook;
            ex_cel.Worksheet oSheet;

            //Tạo mới một Excel WorkBook 
            oExcel.Visible = true;
            oExcel.DisplayAlerts = false;
            oExcel.Application.SheetsInNewWorkbook = 1;
            oBooks = oExcel.Workbooks;
            oBook = (ex_cel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
            oSheets = oBook.Worksheets;
            oSheet = (ex_cel.Worksheet)oSheets.get_Item(1);
            oSheet.Name = sheetname;


            ex_cel.Range head = oSheet.get_Range("A1", "O1");
            head.MergeCells = true;
            head.Value2 = "DANH SÁCH SÁCH";
            head.Font.Bold = true;
            head.Font.Name = "Tahoma";
            head.Font.Size = "16";
            head.HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;

            // ====== HEADER CỘT (HÀNG 3) ======
            // A: MaSach
            ex_cel.Range cl1 = oSheet.get_Range("A3", "A3");
            cl1.Value2 = "MÃ SÁCH";
            cl1.ColumnWidth = 12;

            // B: TenSach
            ex_cel.Range cl2 = oSheet.get_Range("B3", "B3");
            cl2.Value2 = "TÊN SÁCH";
            cl2.ColumnWidth = 30;

            // C: NamXB
            ex_cel.Range cl3 = oSheet.get_Range("C3", "C3");
            cl3.Value2 = "NĂM XB";
            cl3.ColumnWidth = 10;

            // D: SoLuong
            ex_cel.Range cl4 = oSheet.get_Range("D3", "D3");
            cl4.Value2 = "SỐ LƯỢNG";
            cl4.ColumnWidth = 10;

            // E: MoTa
            ex_cel.Range cl5 = oSheet.get_Range("E3", "E3");
            cl5.Value2 = "MÔ TẢ";
            cl5.ColumnWidth = 35;

            // F: MaTG
            ex_cel.Range cl6 = oSheet.get_Range("F3", "F3");
            cl6.Value2 = "MÃ TG";
            cl6.ColumnWidth = 12;

            // G: TenTG (hiển thị)
            ex_cel.Range cl7 = oSheet.get_Range("G3", "G3");
            cl7.Value2 = "TÊN TÁC GIẢ";
            cl7.ColumnWidth = 20;

            // H: MaTL
            ex_cel.Range cl8 = oSheet.get_Range("H3", "H3");
            cl8.Value2 = "MÃ TL";
            cl8.ColumnWidth = 12;

            // I: TenTL (hiển thị)
            ex_cel.Range cl9 = oSheet.get_Range("I3", "I3");
            cl9.Value2 = "TÊN THỂ LOẠI";
            cl9.ColumnWidth = 18;

            // J: MaNXB / TenNXB (tùy bạn chọn cột nào trong DataTable)
            ex_cel.Range cl10 = oSheet.get_Range("J3", "J3");
            cl10.Value2 = "MÃ NXB";
            cl10.ColumnWidth = 22;

            // K: TenNN
            ex_cel.Range cl11 = oSheet.get_Range("K3", "K3");
            cl11.Value2 = "TÊN NXB";
            cl11.ColumnWidth = 15;

            // L: TenKe
            ex_cel.Range cl12 = oSheet.get_Range("L3", "L3");
            cl12.Value2 = "MÃ NN";
            cl12.ColumnWidth = 15;

            ex_cel.Range cl13 = oSheet.get_Range("M3", "M3");
            cl13.Value2 = "TÊN NGÔN NGỮ";
            cl13.ColumnWidth = 12;

            ex_cel.Range cl14 = oSheet.get_Range("N3", "N3");
            cl14.Value2 = "MÃ KỆ SÁCH";
            cl14.ColumnWidth = 15;

            ex_cel.Range cl15 = oSheet.get_Range("O3", "O3");
            cl15.Value2 = "TÊN KỆ";
            cl15.ColumnWidth = 15;

            // Format header row
            ex_cel.Range rowHead = oSheet.get_Range("A3", "O3");
            rowHead.Font.Bold = true;
            rowHead.Borders.LineStyle = ex_cel.Constants.xlSolid;
            rowHead.Interior.ColorIndex = 15;
            rowHead.HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;

            // ====== ĐỔ DỮ LIỆU ======
            object[,] arr = new object[tb.Rows.Count, tb.Columns.Count];

            for (int r = 0; r < tb.Rows.Count; r++)
            {
                DataRow dr = tb.Rows[r];
                for (int c = 0; c < tb.Columns.Count; c++)
                {
                    // Cột mã thường là nvarchar, nếu bạn muốn giữ nguyên dạng text thì có thể thêm "'" cho các cột mã
                    // Ở đây mình giữ y logic cũ của bạn: chỉ thêm ' cho cột thứ 6 (index 5) nếu cần.
                    if (c == 5) arr[r, c] = "'" + dr[c];
                    else arr[r, c] = dr[c];
                }
            }

            int rowStart = 4;
            int columnStart = 1;
            int rowEnd = rowStart + tb.Rows.Count - 1;
            int columnEnd = tb.Columns.Count; // lấy đúng số cột DataTable

            ex_cel.Range c1 = (ex_cel.Range)oSheet.Cells[rowStart, columnStart];
            ex_cel.Range c2 = (ex_cel.Range)oSheet.Cells[rowEnd, columnEnd];
            ex_cel.Range range = oSheet.get_Range(c1, c2);
            range.Value2 = arr;

            // Kẻ viền
            range.Borders.LineStyle = ex_cel.Constants.xlSolid;

            // Căn giữa cột A (Mã sách)
            ex_cel.Range colA1 = (ex_cel.Range)oSheet.Cells[rowStart, 1];
            ex_cel.Range colA2 = (ex_cel.Range)oSheet.Cells[rowEnd, 1];
            oSheet.get_Range(colA1, colA2).HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;
        }

        private void btnxuat_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string masach = txtMaSach.Text.Trim();

            string sql =
                "SELECT " +
                " s.MaSach, s.TenSach, s.NamXB, s.SoLuong, s.MoTa, " +
                " s.MaTG, tg.TenTG, " +
                " s.MaTL, tl.TenTL, " +
                " s.MaNXB, nxb.TenNXB, " +
                " s.MaNN, nn.TenNN, " +
                " s.MaKe, ks.TenKe " +
                "FROM Sach s " +
                "JOIN TacGia tg ON s.MaTG = tg.MaTG " +
                "JOIN TheLoai tl ON s.MaTL = tl.MaTL " +
                "JOIN NhaXuatBan nxb ON s.MaNXB = nxb.MaNXB " +
                "JOIN NgonNgu nn ON s.MaNN = nn.MaNN " +
                "JOIN KeSach ks ON s.MaKe = ks.MaKe ";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;

            DataTable tb = new DataTable();
            da.Fill(tb);

            cmd.Dispose();
            thuvien.con.Close();

            ExportExcel(tb, "Danh Sách Sách");
        }

        private void ReadExcel(string filename)
        {
            //kiểm tra xem filename đã có dữ liệu chưa
            if (filename == null)
            {
                MessageBox.Show("Chưa chọn file");
            }
            else
            {
                xls.Application Excel = new xls.Application();

                Excel.Workbooks.Open(filename);

                foreach (xls.Worksheet wsheet in Excel.Worksheets)
                {
                    int i = 4;
                    do
                    {
                        if (wsheet.Cells[i, 1].Value == null && wsheet.Cells[i, 2].Value == null && wsheet.Cells[i, 3].Value == null)
                        {
                            break;
                        }
                        else
                        {

                            ThemSach(
                                wsheet.Cells[i, 1].Value.ToString(),
                                wsheet.Cells[i, 6].Value.ToString(),
                                wsheet.Cells[i, 10].Value.ToString(),
                                wsheet.Cells[i, 8].Value.ToString(),
                                wsheet.Cells[i, 2].Value.ToString(),
                                int.Parse(wsheet.Cells[i, 3].Value.ToString()),
                                int.Parse(wsheet.Cells[i, 4].Value.ToString()),
                                wsheet.Cells[i, 5].Value.ToString(),
                                wsheet.Cells[i, 12].Value.ToString(),
                                wsheet.Cells[i, 14].Value.ToString()
                            );

                            i++;
                        }
                    }
                    while (true);
                }
            }
            MessageBox.Show("Nhập Excel thành công");
        }

        private void ThemSach(string masach, string matg, string manxb, string matl,
                      string tensach, int namxb, int soluong, string mota,
                      string mann, string make)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "INSERT INTO Sach(MaSach, MaTG, MaNXB, MaTL, TenSach, NamXB, SoLuong, MoTa, MaNN, MaKe) " +
                         "VALUES (@masach, @matg, @manxb, @matl, @tensach, @namxb, @soluong, @mota, @mann, @make)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@masach", masach);
            cmd.Parameters.AddWithValue("@matg", matg);
            cmd.Parameters.AddWithValue("@manxb", manxb);
            cmd.Parameters.AddWithValue("@matl", matl);
            cmd.Parameters.AddWithValue("@tensach", tensach);
            cmd.Parameters.AddWithValue("@namxb", namxb);
            cmd.Parameters.AddWithValue("@soluong", soluong);
            cmd.Parameters.AddWithValue("@mota", mota);
            cmd.Parameters.AddWithValue("@mann", mann);
            cmd.Parameters.AddWithValue("@make", make);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();
        }
        private void btnnhap_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FilterIndex = 1; //Trỏ vào vị trí đầu tiên của bộ lọc
            openFileDialog1.RestoreDirectory = true;//Nhớ đường dẫn của lần truy cập trước
            openFileDialog1.Multiselect = false;//Không cho chọn nhiều file
            DialogResult kq = openFileDialog1.ShowDialog();
            if (kq == DialogResult.OK)
            {
                ReadExcel(openFileDialog1.FileName);
                load_Sach();
            }
        }

        private void btntimkiem_Click_1(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }
            string keyword = txttimkiem.Text.Trim();
            string sql = "SELECT s.MaSach, s.MaTG, s.MaNXB, s.MaTL, s.TenSach, s.NamXB, s.SoLuong, s.MaNN, s.MaKe, " +
                "tg.TenTG, nxb.TenNXB, tl.TenTL, nn.TenNN, ks.TenKe, s.MoTa " +
                "FROM Sach s " +
                "JOIN TacGia tg ON s.MaTG = tg.MaTG " +
                "JOIN NhaXuatBan nxb ON s.MaNXB = nxb.MaNXB " +
                "JOIN TheLoai tl ON s.MaTL = tl.MaTL " +
                "JOIN NgonNgu nn ON s.MaNN = nn.MaNN " +
                "JOIN KeSach ks ON s.MaKe = ks.MaKe " +
                "WHERE s.MaSach LIKE @keyword OR s.TenSach LIKE @keyword OR tg.TenTG LIKE @keyword OR nxb.TenNXB LIKE @keyword " +
                "OR tl.TenTL LIKE @keyword OR nn.TenNN LIKE @keyword OR ks.TenKe LIKE @keyword ";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");
            bool isNumber = int.TryParse(keyword, out int nam);
            if (isNumber)
            {
                cmd.CommandText += " OR s.NamXB = @nam";
                cmd.Parameters.AddWithValue("@nam", nam);
            }

            thuvien.Hienthi(dgvQuanLySach, cmd);
        }

        private void txttimkiem_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
