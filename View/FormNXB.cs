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
    public partial class FormNXB : Form
    {
        public FormNXB()
        {
            InitializeComponent();
            this.Load += QuanLyNhaXuatBan_Load;
        }

        private void load_NhaXuatBan()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "SELECT * FROM NhaXuatBan";
            SetTrangThaiButton(false); 
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            txtMaNXB.Text = TaoMaNXBMoi();  
            thuvien.Hienthi(dgvQuanLyNhaXuatBan, cmd);  
        }

        private void QuanLyNhaXuatBan_Load(object sender, EventArgs e)
        {
            load_NhaXuatBan();  
        }
        private void SetTrangThaiButton(bool enabled)
        {
            btnSua.Enabled = enabled;
            btnXoa.Enabled = enabled;
        }
        private void reset()
        {
            txtMaNXB.Text = TaoMaNXBMoi();
            txtTenNXB.Clear();
            txtDiaChi.Clear();
            txtEmail.Clear();
            txtSdt.Clear();

            dgvQuanLyNhaXuatBan.ClearSelection();
            dgvQuanLyNhaXuatBan.CurrentCell = null;

            SetTrangThaiButton(false);
        }

        private bool KiemTraDuLieu()
        {
            if (string.IsNullOrWhiteSpace(txtTenNXB.Text)) 
            {
                MessageBox.Show("Vui lòng nhập tên nhà xuất bản");
                txtTenNXB.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtDiaChi.Text))
            {
                MessageBox.Show("Vui lòng nhập địa chỉ nhà xuất bản");
                txtDiaChi.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtEmail.Text) || !txtEmail.Text.Contains("@"))
            {
                MessageBox.Show("Vui lòng nhập email hợp lệ (có chứa '@')");
                txtEmail.Focus();
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtSdt.Text) || !txtSdt.Text.All(char.IsDigit) || (txtSdt.Text.Length < 10 || txtSdt.Text.Length > 11))
            {
                MessageBox.Show("Vui lòng nhập số điện thoại hợp lệ (chỉ chứa số và dài từ 10 đến 11 ký tự)");
                txtSdt.Focus();
                return false;
            }

            return true;
        }
        public static string TaoMaNXBMoi()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "SELECT TOP 1 MaNXB FROM NhaXuatBan ORDER BY MaNXB DESC";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            object result = cmd.ExecuteScalar();
            thuvien.con.Close();

            if (result == null)
            {
                return "NXB001";
            }
            else
            {
                string macu = result.ToString();
                int id = int.Parse(macu.Substring(3)) + 1;  
                return "NXB" + id.ToString("D3");
            }
        }

        private void dgvQuanLyNhaXuatBan_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaNXB.Text = dgvQuanLyNhaXuatBan.CurrentRow.Cells[0].Value.ToString();
            txtTenNXB.Text = dgvQuanLyNhaXuatBan.CurrentRow.Cells[1].Value.ToString();
            txtDiaChi.Text = dgvQuanLyNhaXuatBan.CurrentRow.Cells[2].Value.ToString();
            txtEmail.Text = dgvQuanLyNhaXuatBan.CurrentRow.Cells[3].Value.ToString();
            txtSdt.Text = dgvQuanLyNhaXuatBan.CurrentRow.Cells[4].Value.ToString();
            txtMaNXB.Enabled = false;
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string trungma = "SELECT COUNT(*) FROM NhaXuatBan WHERE MaNXB LIKE @manxb";
            SqlCommand kiemtra = new SqlCommand(trungma, thuvien.con);
            kiemtra.Parameters.AddWithValue("@manxb", txtMaNXB.Text.Trim());
            int kq = (int)kiemtra.ExecuteScalar();

            if (kq > 0)
            {
                MessageBox.Show("Trùng mã nhà xuất bản, vui lòng bấm nút làm mới và nhập lại");
                return;
            }

            string sql = @"INSERT INTO NhaXuatBan (MaNXB, TenNXB, DiaChi, Email, Sdt)
                   VALUES (@manxb, @tennxb, @diachi, @email, @sdt)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@manxb", txtMaNXB.Text.Trim());
            cmd.Parameters.AddWithValue("@tennxb", txtTenNXB.Text.Trim());
            cmd.Parameters.AddWithValue("@diachi", txtDiaChi.Text.Trim());
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim());
            cmd.Parameters.AddWithValue("@sdt", txtSdt.Text.Trim());

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Thêm nhà xuất bản thành công");
            reset();
            load_NhaXuatBan();
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaNXB.Text))
            {
                MessageBox.Show("Vui lòng chọn nhà xuất bản cần xóa.");
                return;
            }

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string manxb = txtMaNXB.Text.Trim();

            string sql = "SELECT COUNT(*) FROM Sach WHERE MaNXB = @manxb";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@manxb", manxb);

            int countSach = (int)cmd.ExecuteScalar();

            if (countSach > 0)
            {
                MessageBox.Show("Không thể xóa nhà xuất bản vì đã có sách liên quan.");
                thuvien.con.Close();
                return;
            }

            DialogResult kq = MessageBox.Show(
                "Bạn có chắc muốn xóa nhà xuất bản này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (kq == DialogResult.No)
            {
                thuvien.con.Close();
                return;
            }

            string sql1 = "DELETE FROM NhaXuatBan WHERE MaNXB = @manxb";
            SqlCommand cmd1 = new SqlCommand(sql1, thuvien.con);
            cmd1.Parameters.AddWithValue("@manxb", manxb);
            cmd1.ExecuteNonQuery();
            cmd1.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Xóa nhà xuất bản thành công.");
            reset();
            load_NhaXuatBan();
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaNXB.Text))
            {
                MessageBox.Show("Vui lòng chọn nhà xuất bản cần sửa");
                return;
            }

            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = @"UPDATE NhaXuatBan SET
                   TenNXB = @tennxb, DiaChi = @diachi, Email = @email, Sdt = @sdt
                   WHERE MaNXB = @manxb";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@manxb", txtMaNXB.Text.Trim());
            cmd.Parameters.AddWithValue("@tennxb", txtTenNXB.Text.Trim());
            cmd.Parameters.AddWithValue("@diachi", txtDiaChi.Text.Trim());
            cmd.Parameters.AddWithValue("@email", txtEmail.Text.Trim());
            cmd.Parameters.AddWithValue("@sdt", txtSdt.Text.Trim());

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Cập nhật nhà xuất bản thành công");
            reset();
            load_NhaXuatBan();
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            reset(); 
            load_NhaXuatBan(); 
            SetTrangThaiButton(false);
        }

        public void ExportExcel(DataTable tb, string sheetname)
        {
            ex_cel.Application oExcel = new ex_cel.Application();
            ex_cel.Workbooks oBooks;
            ex_cel.Sheets oSheets;
            ex_cel.Workbook oBook;
            ex_cel.Worksheet oSheet;

            // Tạo mới một Excel WorkBook 
            oExcel.Visible = true;
            oExcel.DisplayAlerts = false;
            oExcel.Application.SheetsInNewWorkbook = 1;
            oBooks = oExcel.Workbooks;
            oBook = (ex_cel.Workbook)(oExcel.Workbooks.Add(Type.Missing));
            oSheets = oBook.Worksheets;
            oSheet = (ex_cel.Worksheet)oSheets.get_Item(1);
            oSheet.Name = sheetname;

            ex_cel.Range head = oSheet.get_Range("A1", "E2");
            head.MergeCells = true;
            head.Value2 = "DANH SÁCH NHÀ XUẤT BẢN";
            head.Font.Bold = true;
            head.Font.Name = "Tahoma";
            head.Font.Size = 16;
            head.HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;
            // ====== HEADER CỘT (HÀNG 3) ======
            // A: MaNXB
            ex_cel.Range cl1 = oSheet.get_Range("A3", "A3");
            cl1.Value2 = "MÃ NHÀ XUẤT BẢN";
            cl1.ColumnWidth = 12;

            // B: TenNXB
            ex_cel.Range cl2 = oSheet.get_Range("B3", "B3");
            cl2.Value2 = "TÊN NHÀ XUẤT BẢN";
            cl2.ColumnWidth = 30;

            // C: DiaChi
            ex_cel.Range cl3 = oSheet.get_Range("C3", "C3");
            cl3.Value2 = "ĐỊA CHỈ";
            cl3.ColumnWidth = 25;

            // D: Email
            ex_cel.Range cl4 = oSheet.get_Range("D3", "D3");
            cl4.Value2 = "EMAIL";
            cl4.ColumnWidth = 25;

            // E: Sdt
            ex_cel.Range cl5 = oSheet.get_Range("E3", "E3");
            cl5.Value2 = "SĐT";
            cl5.ColumnWidth = 15;

            // Format header row
            ex_cel.Range rowHead = oSheet.get_Range("A3", "E3");
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
                    if (c == 4) // Cột SĐT
                    {
                        arr[r, c] = "'" + dr[c].ToString();  
                    }
                    else
                    {
                        arr[r, c] = dr[c];
                    }
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

            // Kẻ viền
            range.Borders.LineStyle = ex_cel.Constants.xlSolid;

            // Căn giữa cột A (Mã nhà xuất bản)
            ex_cel.Range colA1 = (ex_cel.Range)oSheet.Cells[rowStart, 1];
            ex_cel.Range colA2 = (ex_cel.Range)oSheet.Cells[rowEnd, 1];
            oSheet.get_Range(colA1, colA2).HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;
        }
        private void ReadExcel(string filename)
        {
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

                            ThemNhaXuatBan(
                                wsheet.Cells[i, 1].Value.ToString(),
                                wsheet.Cells[i, 2].Value.ToString(),
                                wsheet.Cells[i, 3].Value.ToString(),
                                wsheet.Cells[i, 4].Value.ToString(),
                                wsheet.Cells[i, 5].Value.ToString()
                            );

                            i++;
                        }
                    }
                    while (true);
                }
            }
            MessageBox.Show("Nhập Excel thành công");
        }
        private void ThemNhaXuatBan(string maNXB, string tenNXB, string diaChi, string email, string sdt)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "INSERT INTO NhaXuatBan(MaNXB, TenNXB, DiaChi, Email, Sdt) VALUES (@manxb, @tennxb, @diachi, @email, @sdt)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@manxb", maNXB);
            cmd.Parameters.AddWithValue("@tennxb", tenNXB);
            cmd.Parameters.AddWithValue("@diachi", diaChi);
            cmd.Parameters.AddWithValue("@email", email);
            cmd.Parameters.AddWithValue("@sdt", sdt);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();
        }

        private void btnNhap_Click(object sender, EventArgs e)
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
                load_NhaXuatBan();
            }
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "SELECT MaNXB, TenNXB, DiaChi, Email, Sdt FROM NhaXuatBan"; 
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;

            DataTable tb = new DataTable();
            da.Fill(tb);

            cmd.Dispose();
            thuvien.con.Close();

            ExportExcel(tb, "Danh Sách Nhà Xuất Bản");
        }

        private void btnTimKiem_Click(object sender, EventArgs e)
        {
            string keyword = txtTimKiem.Text.Trim();

            if (string.IsNullOrEmpty(keyword))
            {
                MessageBox.Show("Vui lòng nhập từ khóa tìm kiếm.");
                return;
            }

            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "SELECT * FROM NhaXuatBan WHERE MaNXB LIKE @keyword OR TenNXB LIKE @keyword OR DiaChi LIKE @keyword OR Email LIKE @keyword OR Sdt LIKE @keyword";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable tb = new DataTable();
            da.Fill(tb);

            thuvien.con.Close();

            dgvQuanLyNhaXuatBan.DataSource = tb;
        }
    }
}
