using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ex_cel = Microsoft.Office.Interop.Excel;
using xls = Microsoft.Office.Interop.Excel;


namespace projectQLTV.View
{
    public partial class FormTheLoai : Form
    {
        public FormTheLoai()
        {
            InitializeComponent();
            this.Load += QuanLyTheLoai_Load;

        }
        private void load_TheLoai()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }
            string sql = "SELECT * FROM TheLoai";
            SetTrangThaiButton(false); 
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            txtMaTL.Text = TaoMaTLMoi();  
            thuvien.Hienthi(dgvQuanLyTheLoai, cmd); 
        }
        private void QuanLyTheLoai_Load(object sender, EventArgs e)
        {

            load_TheLoai();
        }
        private void SetTrangThaiButton(bool enabled)
        {
            btnSua.Enabled = enabled;
            btnXoa.Enabled = enabled;
        }

        private void reset()
        {
            txtMaTL.Text = TaoMaTLMoi();
            txtTenTL.Clear();


            dgvQuanLyTheLoai.ClearSelection();
            dgvQuanLyTheLoai.CurrentCell = null;

            SetTrangThaiButton(false);
        }

        private void dgvQuanLyTheLoai_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            txtMaTL.Text = dgvQuanLyTheLoai.CurrentRow.Cells[0].Value.ToString();
            txtTenTL.Text = dgvQuanLyTheLoai.CurrentRow.Cells[1].Value.ToString();
            txtMaTL.Enabled = false;
            btnXoa.Enabled = true;
            btnSua.Enabled = true;
        }


        private bool KiemTraDuLieu()
        {
            if (string.IsNullOrWhiteSpace(txtTenTL.Text)) 
            {
                MessageBox.Show("Vui lòng nhập tên thể loại");
                txtTenTL.Focus();
                return false;
            }

            return true;
        }

        public static string TaoMaTLMoi()
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string sql = "SELECT TOP 1 MaTL FROM TheLoai ORDER BY MaTL DESC";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            object result = cmd.ExecuteScalar();
            thuvien.con.Close();

            if (result == null)
            {
                return "TL001";
            }
            else
            {
                string macu = result.ToString();
                int id = int.Parse(macu.Substring(2)) + 1;  
                return "TL" + id.ToString("D3");
            }
        }

        private void btnThem_Click(object sender, EventArgs e)
        {
            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string trungma = "SELECT COUNT(*) FROM TheLoai WHERE MaTL LIKE @matl";
            SqlCommand kiemtra = new SqlCommand(trungma, thuvien.con);
            kiemtra.Parameters.AddWithValue("@matl", txtMaTL.Text.Trim());
            int kq = (int)kiemtra.ExecuteScalar();

            if (kq > 0)
            {
                MessageBox.Show("Trùng mã thể loại, vui lòng bấm nút làm mới và nhập lại");
                return;
            }

            string sql = @"INSERT INTO TheLoai (MaTL, TenTL)
                   VALUES (@matl, @tentl)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@matl", txtMaTL.Text.Trim());
            cmd.Parameters.AddWithValue("@tentl", txtTenTL.Text.Trim());

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Thêm thể loại thành công");
            reset();
            load_TheLoai();
        }

        private void btnXoa_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaTL.Text))
            {
                MessageBox.Show("Vui lòng chọn thể loại cần xóa.");
                return;
            }

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string matl = txtMaTL.Text.Trim();

            string sql = "SELECT COUNT(*) FROM Sach WHERE MaTL = @matl";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@matl", matl);

            int countSach = (int)cmd.ExecuteScalar();

            if (countSach > 0)
            {
                MessageBox.Show("Không thể xóa thể loại vì đã có sách liên quan.");
                thuvien.con.Close();
                return;
            }

            DialogResult kq = MessageBox.Show(
                "Bạn có chắc muốn xóa thể loại này không?",
                "Xác nhận",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (kq == DialogResult.No)
            {
                thuvien.con.Close();
                return;
            }

            string sql1 = "DELETE FROM TheLoai WHERE MaTL = @matl";
            SqlCommand cmd1 = new SqlCommand(sql1, thuvien.con);
            cmd1.Parameters.AddWithValue("@matl", matl);
            cmd1.ExecuteNonQuery();
            cmd1.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Xóa thể loại thành công.");
            reset();
            load_TheLoai();
        

    }

        private void btnSua_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtMaTL.Text))
            {
                MessageBox.Show("Vui lòng chọn thể loại cần sửa");
                return;
            }

            if (!KiemTraDuLieu()) return;

            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = @"UPDATE TheLoai SET
                   TenTL = @tentl
                   WHERE MaTL = @matl";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@matl", txtMaTL.Text.Trim());
            cmd.Parameters.AddWithValue("@tentl", txtTenTL.Text.Trim());

            cmd.ExecuteNonQuery();
            cmd.Dispose();
            thuvien.con.Close();

            MessageBox.Show("Cập nhật thể loại thành công");
            reset();
            load_TheLoai();
        }


        private void btnLamMoi_Click(object sender, EventArgs e)
        {
            reset();  // Reset các trường dữ liệu trong form
            load_TheLoai();  // Làm mới DataGridView
            SetTrangThaiButton(false);  // Tắt nút sửa, xóa
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
            head.Value2 = "DANH SÁCH THỂ LOẠI";
            head.Font.Bold = true;
            head.Font.Name = "Tahoma";
            head.Font.Size = 16;  
            head.HorizontalAlignment = ex_cel.XlHAlign.xlHAlignCenter;
            // ====== HEADER CỘT (HÀNG 3) ======
            // A: MaTL
            ex_cel.Range cl1 = oSheet.get_Range("A3", "A3");
            cl1.Value2 = "MÃ THỂ LOẠI";
            cl1.ColumnWidth = 12;

            // B: TenTL
            ex_cel.Range cl2 = oSheet.get_Range("B3", "B3");
            cl2.Value2 = "TÊN THỂ LOẠI";
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

            // Căn giữa cột A (Mã thể loại)
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

                            ThemTheLoai(
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
        private void ThemTheLoai(string maTL, string tenTL)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "INSERT INTO TheLoai(MaTL, TenTL) VALUES (@matl, @tentl)";

            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@matl", maTL);
            cmd.Parameters.AddWithValue("@tentl", tenTL);

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
                load_TheLoai();
            }
        }

        private void btnXuat_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
                thuvien.con.Open();

            string sql = "SELECT MaTL, TenTL FROM TheLoai";  
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);

            SqlDataAdapter da = new SqlDataAdapter();
            da.SelectCommand = cmd;

            DataTable tb = new DataTable();
            da.Fill(tb);

            cmd.Dispose();
            thuvien.con.Close();

            ExportExcel(tb, "Danh Sách Thể Loại");
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            if (thuvien.con.State == ConnectionState.Closed)
            {
                thuvien.con.Open();
            }

            string keyword = txtTimKiem.Text.Trim();
            string sql = "SELECT * FROM TheLoai WHERE MaTL LIKE @keyword OR TenTL LIKE @keyword";
            SqlCommand cmd = new SqlCommand(sql, thuvien.con);
            cmd.Parameters.AddWithValue("@keyword", "%" + keyword + "%");

            SqlDataAdapter da = new SqlDataAdapter(cmd);
            DataTable tb = new DataTable();
            da.Fill(tb);

            thuvien.con.Close();

            dgvQuanLyTheLoai.DataSource = tb;
        }
    }
}
