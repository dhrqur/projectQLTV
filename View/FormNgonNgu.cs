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
    public partial class FormNgonNgu : Form
    {
        public FormNgonNgu()
        {
            InitializeComponent();
            this.Load += QuanLyNgonNgu_Load;
        }
        private void load_NgonNgu()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            // Truy vấn SQL để lấy dữ liệu từ bảng NgonNgu
            string sql = "SELECT * FROM NgonNgu";
            SetTrangThaiButton(false); // Nếu cần thiết, có thể bỏ qua nếu không cần trạng thái nút
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            txtMaNN.Text = TaoMaNNMoi();  // Gọi hàm tạo mã tự động
            thuvien.Hienthi(dgvQuanLyNgonNgu, cmd);  // Cập nhật DataGridView
        }
        private void QuanLyNgonNgu_Load(object sender, EventArgs e)
        {
            load_NgonNgu();  // Hiển thị dữ liệu của bảng NgonNgu
        }
        private bool KiemTraDuLieu()
        {
            if (string.IsNullOrWhiteSpace(txtTenNN.Text)) // Tên ngôn ngữ
            {
                MessageBox.Show("Vui lòng nhập tên ngôn ngữ");
                txtTenNN.Focus();
                return false;
            }

            return true;
        }
        public static string TaoMaNNMoi()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "SELECT TOP 1 MaNN FROM NgonNgu ORDER BY MaNN DESC";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            object result = cmd.ExecuteScalar();
            thuvien.con.Close();

            if (result == null)
            {
                return "NN001";
            }
            else
            {
                string macu = result.ToString();
                int id = int.Parse(macu.Substring(2)) + 1;  // Tăng số sau "NN"
                return "NN" + id.ToString("D3");
            }
        }
        private void btnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string trungma = "SELECT COUNT(*) FROM NgonNgu WHERE MaNN LIKE @mann";
            SqlCommand kiemtra = new SqlCommand(trungma, thuvien.con);
            kiemtra.Parameters.AddWithValue("@mann", txtMaNN.Text.Trim());
            int kq = (int)kiemtra.ExecuteScalar();

            if (kq > 0)
            {
                MessageBox.Show("Trùng mã ngôn ngữ, vui lòng bấm nút làm mới và nhập lại");
                return;
            }

            string sql = @"INSERT INTO NgonNgu (MaNN, TenNN)
                   VALUES (@mann, @tennn)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@mann", txtMaNN.Text.Trim());
            cmd.Parameters.AddWithValue("@tennn", txtTenNN.Text.Trim());

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Thêm ngôn ngữ thành công");
            reset();
            load_NgonNgu();

        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaNN.Text))
            {
                MessageBox.Show("Vui lòng chọn ngôn ngữ cần xóa.");
                return;
            }

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string mann = txtMaNN.Text.Trim();

            string sql = "SELECT COUNT(*) FROM Sach WHERE MaNN = @mann";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@mann", mann);

            int countSach = (int)cmd.ExecuteScalar();

            if (countSach > 0)
            {
                MessageBox.Show("Không thể xóa ngôn ngữ vì đã có sách liên quan.");
                thuvien.con.Close();
                return;
            }

            DialogResult kq = MessageBox.Show(
                "Bạn có chắc muốn xóa ngôn ngữ này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (kq == DialogResult.No)
            {
                thuvien.con.Close();
                return;
            }

            string sql1 = "DELETE FROM NgonNgu WHERE MaNN = @mann";
            SqlCommand cmd1 = new SqlCommand(sql1, thuvien.con);
            cmd1.Parameters.AddWithValue("@mann", mann);
            cmd1.ExecuteNonQuery();
            cmd1.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Xóa ngôn ngữ thành công.");
            reset();
            load_NgonNgu();
        }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaNN.Text))
            {
                MessageBox.Show("Vui lòng chọn ngôn ngữ cần sửa");
                return;
            }

            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = @"UPDATE NgonNgu SET
                   TenNN = @tennn
                   WHERE MaNN = @mann";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@mann", txtMaNN.Text.Trim());
            cmd.Parameters.AddWithValue("@tennn", txtTenNN.Text.Trim());

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Cập nhật ngôn ngữ thành công");
            reset();
            load_NgonNgu();
        }

        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            reset();  // Reset các trường dữ liệu trong form
            load_NgonNgu();  // Làm mới DataGridView
            SetTrangThaiButton(false);  // Tắt nút sửa, xóa
        }

        private void SetTrangThaiButton(bool enabled)
        {
            btnSua.Enabled = enabled;
            btnXoa.Enabled = enabled;
        }

        private void reset()
        {
            txtMaNN.Text = TaoMaNNMoi();
            txtTenNN.Clear();

            dgvQuanLyNgonNgu.ClearSelection();
            dgvQuanLyNgonNgu.CurrentCell = null;

            SetTrangThaiButton(false);
        }

        private void dgvQuanLyNgonNgu_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaNN.Text = dgvQuanLyNgonNgu.CurrentRow.Cells[0].Value.ToString();
            txtTenNN.Text = dgvQuanLyNgonNgu.CurrentRow.Cells[1].Value.ToString();
            txtMaNN.Enabled = false;
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
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

            string sql = "SELECT * FROM NgonNgu WHERE MaNN LIKE @keyword OR TenNN LIKE @keyword";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable tb = new DataTable();
            da.Fill(tb);

            thuvien.con.Close();

            dgvQuanLyNgonNgu.DataSource = tb;
        }

        public void ExportExcel(DataTable tb, string sheetname)
        {
            // Tạo các đối tượng Excel
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

            ex_cel.Range head = oSheet.get_Range("A1", "B2");
            head.MergeCells = true;
            head.Value2 = "DANH SÁCH NGÔN NGỮ";
            head.Font.Bold = true;
            head.Font.Name = "Tahoma";
            head.Font.Size = 16;
            head.HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;

            // ====== HEADER CỘT (HÀNG 3) ======
            // A: MaNN
            ex_cel.Range cl1 = oSheet.get_Range("A3", "A3");
            cl1.Value2 = "MÃ NGÔN NGỮ";
            cl1.ColumnWidth = 12;

            // B: TenNN
            ex_cel.Range cl2 = oSheet.get_Range("B3", "B3");
            cl2.Value2 = "TÊN NGÔN NGỮ";
            cl2.ColumnWidth = 30;

            // Format header row
            ex_cel.Range rowHead = oSheet.get_Range("A3", "B3");
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

            // Kẻ viền
            range.Borders.LineStyle = ex_cel.Constants.xlSolid;

            // Căn giữa cột A (Mã ngôn ngữ)
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
                        if (wsheet.Cells[i, 1].Value == null && wsheet.Cells[i, 2].Value == null)
                        {
                            break;
                        }
                        else
                        {
                            // Nhập dữ liệu từ Excel vào bảng NgonNgu
                            ThemNgonNgu(
                                wsheet.Cells[i, 1].Value.ToString(),
                                wsheet.Cells[i, 2].Value.ToString()
                            );

                            i++;
                        }
                    }
                    while (true);
                }
            }
            MessageBox.Show("Nhập Excel thành công");
        }

        private void ThemNgonNgu(string maNN, string tenNN)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "INSERT INTO NgonNgu(MaNN, TenNN) VALUES (@mann, @tennn)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@mann", maNN);
            cmd.Parameters.AddWithValue("@tennn", tenNN);

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();
        }

        private void btnNhap_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog1.FilterIndex = 1; //Trỏ vào vị trí đầu tiên của bộ lọc
            openFileDialog1.RestoreDirectory = true; //Nhớ đường dẫn của lần truy cập trước
            openFileDialog1.Multiselect = false; //Không cho chọn nhiều file
            DialogResult kq = openFileDialog1.ShowDialog();
            if (kq == DialogResult.OK)
            {
                ReadExcel(openFileDialog1.FileName);
                load_NgonNgu();
            }
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "SELECT MaNN, TenNN FROM NgonNgu";  // Lấy dữ liệu ngôn ngữ từ bảng NgonNgu
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;

            DataTable tb = new DataTable();
            da.Fill(tb);

            cmd.Dispose();
            thuvien.con.Close();

            ExportExcel(tb, "Danh Sách Ngôn Ngữ");
        }
    }
}
