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
using projectQLTV.Helper;
using OfficeOpenXml;

namespace projectQLTV.View
{
    public partial class FormLop : Form
    {
        private string currentMaLop = "";

        public FormLop()
        {
            InitializeComponent();
            WireUpEvents();
            LoadKhoaComboBox();
            LoadData();
        }

        private void WireUpEvents()
        {
            btnThem.Click += btnThem_Click; // Thêm
            btnSua.Click += btnSua_Click; // Sửa
            btnXoa.Click += btnXoa_Click; // Xóa
            btnLamMoi.Click += btnLamMoi_Click; // Làm mới
            btnTimKiem.Click += btnTimKiem_Click; // Tìm kiếm
            btnNhapFile.Click += btnNhapFile_Click; // Nhập file
            btnXuatFile.Click += btnXuatFile_Click; // Xuất file
            dgvLop.CellClick += dgvLop_CellClick;
        }

        private void LoadKhoaComboBox()
        {
            try
            {
                string query = "SELECT MaKhoa, TenKhoa FROM Khoa";
                DataTable dt = DatabaseHelper.ExecuteQuery(query);
                cboKhoa.Items.Clear();
                cboKhoa.Items.Add("-- Chọn Khoa --");
                foreach (DataRow row in dt.Rows)
                {
                    cboKhoa.Items.Add(row["MaKhoa"].ToString() + " - " + row["TenKhoa"].ToString());
                }
                if (cboKhoa.Items.Count > 0)
                    cboKhoa.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải danh sách khoa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadData()
        {
            try
            {
                string query = @"SELECT l.MaLop, l.TenLop, l.MaKhoa, k.TenKhoa 
                                 FROM Lop l 
                                 LEFT JOIN Khoa k ON l.MaKhoa = k.MaKhoa";
                DataTable dt = DatabaseHelper.ExecuteQuery(query);
                dgvLop.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GenerateMaLop()
        {
            try
            {
                string query = "SELECT MAX(CAST(SUBSTRING(MaLop, 2, LEN(MaLop)) AS INT)) FROM Lop WHERE MaLop LIKE 'L%'";
                object result = DatabaseHelper.ExecuteScalar(query);
                int nextNumber = 1;
                if (result != null && result != DBNull.Value)
                {
                    nextNumber = Convert.ToInt32(result) + 1;
                }
                return "L" + nextNumber.ToString("D03");
            }
            catch
            {
                return "L001";
            }
        }

        private void ClearFields()
        {
            txtMaLop.Text = ""; // MaLop
            txtTenLop.Text = ""; // TenLop
            cboKhoa.SelectedIndex = 0;
            currentMaLop = "";
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtTenLop.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên Lớp!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenLop.Focus();
                return false;
            }

            if (cboKhoa.SelectedIndex <= 0 || cboKhoa.SelectedItem == null)
            {
                MessageBox.Show("Vui lòng chọn Khoa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboKhoa.Focus();
                return false;
            }

            return true;
        }

        private string GetMaKhoaFromComboBox()
        {
            if (cboKhoa.SelectedItem != null)
            {
                string selected = cboKhoa.SelectedItem.ToString();
                if (selected.Contains(" - "))
                {
                    return selected.Split('-')[0].Trim();
                }
            }
            return "";
        }

        private bool CheckDuplicate(string maLop, string tenLop)
        {
            try
            {
                string query = "SELECT COUNT(*) FROM Lop WHERE (MaLop = @MaLop OR TenLop = @TenLop) AND MaLop != @CurrentMaLop";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaLop", maLop),
                    new SqlParameter("@TenLop", tenLop),
                    new SqlParameter("@CurrentMaLop", currentMaLop)
                };
                int count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(query, parameters));
                return count > 0;
            }
            catch
            {
                return false;
            }
        }

        private void btnThem_Click(object sender, EventArgs e) // Thêm
        {
            try
            {
                if (!ValidateInput())
                    return;

                string maLop = GenerateMaLop();
                string tenLop = txtTenLop.Text.Trim();
                string maKhoa = GetMaKhoaFromComboBox();

                if (string.IsNullOrEmpty(maKhoa))
                {
                    MessageBox.Show("Vui lòng chọn Khoa hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (CheckDuplicate(maLop, tenLop))
                {
                    MessageBox.Show("Mã Lớp hoặc Tên Lớp đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "INSERT INTO Lop (MaLop, TenLop, MaKhoa) VALUES (@MaLop, @TenLop, @MaKhoa)";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaLop", maLop),
                    new SqlParameter("@TenLop", tenLop),
                    new SqlParameter("@MaKhoa", maKhoa)
                };

                int result = DatabaseHelper.ExecuteNonQuery(query, parameters);
                if (result > 0)
                {
                    MessageBox.Show("Thêm lớp thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm lớp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSua_Click(object sender, EventArgs e) // Sửa
        {
            try
            {
                if (string.IsNullOrEmpty(currentMaLop))
                {
                    MessageBox.Show("Vui lòng chọn lớp cần sửa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!ValidateInput())
                    return;

                string tenLop = txtTenLop.Text.Trim();
                string maKhoa = GetMaKhoaFromComboBox();

                if (string.IsNullOrEmpty(maKhoa))
                {
                    MessageBox.Show("Vui lòng chọn Khoa hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (CheckDuplicate(currentMaLop, tenLop))
                {
                    MessageBox.Show("Tên Lớp đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "UPDATE Lop SET TenLop = @TenLop, MaKhoa = @MaKhoa WHERE MaLop = @MaLop";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaLop", currentMaLop),
                    new SqlParameter("@TenLop", tenLop),
                    new SqlParameter("@MaKhoa", maKhoa)
                };

                int result = DatabaseHelper.ExecuteNonQuery(query, parameters);
                if (result > 0)
                {
                    MessageBox.Show("Sửa lớp thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi sửa lớp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnXoa_Click(object sender, EventArgs e) // Xóa
        {
            try
            {
                if (string.IsNullOrEmpty(currentMaLop))
                {
                    MessageBox.Show("Vui lòng chọn lớp cần xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Kiểm tra xem lớp có đang được sử dụng không
                string checkQuery = "SELECT COUNT(*) FROM DocGia WHERE MaLop = @MaLop";
                SqlParameter[] checkParams = new SqlParameter[] { new SqlParameter("@MaLop", currentMaLop) };
                int count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkQuery, checkParams));
                
                if (count > 0)
                {
                    MessageBox.Show("Không thể xóa lớp này vì đang có độc giả sử dụng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa lớp này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    string query = "DELETE FROM Lop WHERE MaLop = @MaLop";
                    SqlParameter[] parameters = new SqlParameter[] { new SqlParameter("@MaLop", currentMaLop) };

                    int deleteResult = DatabaseHelper.ExecuteNonQuery(query, parameters);
                    if (deleteResult > 0)
                    {
                        MessageBox.Show("Xóa lớp thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData();
                        ClearFields();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa lớp: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLamMoi_Click(object sender, EventArgs e) // Làm mới
        {
            ClearFields();
            LoadKhoaComboBox();
            LoadData();
        }

        private void btnTimKiem_Click(object sender, EventArgs e) // Tìm kiếm
        {
            try
            {
                string searchText = txtTimKiem.Text.Trim();
                string query = @"SELECT l.MaLop, l.TenLop, l.MaKhoa, k.TenKhoa 
                                 FROM Lop l 
                                 LEFT JOIN Khoa k ON l.MaKhoa = k.MaKhoa
                                 WHERE l.MaLop LIKE @Search OR l.TenLop LIKE @Search OR k.TenKhoa LIKE @Search";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Search", "%" + searchText + "%")
                };
                DataTable dt = DatabaseHelper.ExecuteQuery(query, parameters);
                dgvLop.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgvLop_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvLop.Rows[e.RowIndex];
                currentMaLop = row.Cells[0].Value.ToString();
                txtMaLop.Text = currentMaLop;
                txtTenLop.Text = row.Cells[1].Value.ToString();
                
                // Set combo box
                string maKhoa = row.Cells[2].Value?.ToString() ?? "";
                for (int i = 0; i < cboKhoa.Items.Count; i++)
                {
                    if (cboKhoa.Items[i].ToString().StartsWith(maKhoa + " - "))
                    {
                        cboKhoa.SelectedIndex = i;
                        break;
                    }
                }
            }
        }

        private void btnNhapFile_Click(object sender, EventArgs e) // Nhập file
        {
            try
            {
                OpenFileDialog openFile = new OpenFileDialog();
                openFile.Filter = "Excel Files|*.xlsx;*.xls|CSV Files|*.csv|All Files|*.*";
                openFile.Title = "Chọn file Excel để nhập";

                if (openFile.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = ReadExcelFile(openFile.FileName);
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        int successCount = 0;
                        int errorCount = 0;
                        int startRow = 0;

                        // Tự động phát hiện header: nếu dòng đầu tiên chứa từ khóa header thì skip
                        if (dt.Rows.Count > 0)
                        {
                            string firstRowFirstCol = dt.Rows[0][0].ToString().Trim().ToLower();
                            if (firstRowFirstCol.Contains("tên") || firstRowFirstCol.Contains("mã") || 
                                firstRowFirstCol.Contains("ten") || firstRowFirstCol.Contains("ma") || 
                                firstRowFirstCol == "tên lớp" || firstRowFirstCol == "ten lop")
                            {
                                startRow = 1; // Skip header row
                            }
                        }

                        for (int i = startRow; i < dt.Rows.Count; i++)
                        {
                            DataRow row = dt.Rows[i];
                            try
                            {
                                // Kiểm tra số cột hợp lệ
                                if (row.ItemArray.Length < 1)
                                    continue;
                                    
                                string tenLop = row[0] != null ? row[0].ToString().Trim() : "";
                                string maKhoa = row.ItemArray.Length > 1 && row[1] != null ? row[1].ToString().Trim() : "";

                                if (string.IsNullOrWhiteSpace(tenLop) || string.IsNullOrWhiteSpace(maKhoa))
                                    continue;

                                // Kiểm tra khoa tồn tại
                                string checkKhoaQuery = "SELECT COUNT(*) FROM Khoa WHERE MaKhoa = @MaKhoa";
                                SqlParameter[] checkKhoaParams = new SqlParameter[] { new SqlParameter("@MaKhoa", maKhoa) };
                                int khoaCount = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkKhoaQuery, checkKhoaParams));

                                if (khoaCount == 0)
                                {
                                    errorCount++;
                                    continue;
                                }

                                // Kiểm tra trùng
                                string checkQuery = "SELECT COUNT(*) FROM Lop WHERE TenLop = @TenLop";
                                SqlParameter[] checkParams = new SqlParameter[] { new SqlParameter("@TenLop", tenLop) };
                                int count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkQuery, checkParams));

                                if (count == 0)
                                {
                                    string maLop = GenerateMaLop();
                                    string insertQuery = "INSERT INTO Lop (MaLop, TenLop, MaKhoa) VALUES (@MaLop, @TenLop, @MaKhoa)";
                                    SqlParameter[] parameters = new SqlParameter[]
                                    {
                                        new SqlParameter("@MaLop", maLop),
                                        new SqlParameter("@TenLop", tenLop),
                                        new SqlParameter("@MaKhoa", maKhoa)
                                    };
                                    DatabaseHelper.ExecuteNonQuery(insertQuery, parameters);
                                    successCount++;
                                }
                                else
                                {
                                    errorCount++;
                                }
                            }
                            catch
                            {
                                errorCount++;
                            }
                        }

                        MessageBox.Show($"Nhập file thành công!\nThành công: {successCount}\nLỗi/Trùng: {errorCount}", "Kết quả", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi nhập file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnXuatFile_Click(object sender, EventArgs e) // Xuất file
        {
            try
            {
                SaveFileDialog saveFile = new SaveFileDialog();
                saveFile.Filter = "Excel Files|*.xlsx|CSV Files|*.csv|All Files|*.*";
                saveFile.Title = "Lưu file Excel";
                saveFile.FileName = "DanhSachLop_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = (DataTable)dgvLop.DataSource;
                    if (dt != null)
                    {
                        ExportToExcel(dt, saveFile.FileName);
                        MessageBox.Show("Xuất file thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xuất file: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private DataTable ReadExcelFile(string filePath)
        {
            DataTable dt = new DataTable();
            try
            {
                if (filePath.EndsWith(".csv"))
                {
                    using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8))
                    {
                        string line;
                        bool isFirstLine = true;
                        while ((line = sr.ReadLine()) != null)
                        {
                            string[] values = line.Split(',');
                            if (isFirstLine)
                            {
                                for (int i = 0; i < values.Length; i++)
                                {
                                    dt.Columns.Add("Column" + i);
                                }
                                isFirstLine = false;
                            }
                            dt.Rows.Add(values);
                        }
                    }
                }
                else
                {
                    // Sử dụng EPPlus để đọc Excel
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Lấy sheet đầu tiên
                        if (worksheet == null)
                            throw new Exception("Không tìm thấy sheet trong file Excel.");

                        // Tạo cột cho DataTable
                        dt.Columns.Add("Column0"); // TenLop
                        dt.Columns.Add("Column1"); // MaKhoa

                        int rowCount = worksheet.Dimension?.End.Row ?? 0;
                        int colCount = worksheet.Dimension?.End.Column ?? 0;
                        
                        // Luôn bỏ qua cột đầu tiên (MaLop), đọc từ cột 2 (TenLop) và cột 3 (MaKhoa)
                        int startCol = 2; // Bỏ qua cột MaLop
                        
                        // Đọc dữ liệu từ dòng 1 (có thể là header)
                        for (int row = 1; row <= rowCount; row++)
                        {
                            string cell1 = startCol <= colCount ? worksheet.Cells[row, startCol].Text?.Trim() ?? "" : "";
                            string cell2 = startCol + 1 <= colCount ? worksheet.Cells[row, startCol + 1].Text?.Trim() ?? "" : "";
                            
                            // Bỏ qua dòng trống hoàn toàn
                            if (string.IsNullOrWhiteSpace(cell1) && string.IsNullOrWhiteSpace(cell2))
                                continue;
                            
                            dt.Rows.Add(cell1, cell2);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Fallback to CSV if Excel reading fails
                try
                {
                    dt = new DataTable();
                    using (StreamReader sr = new StreamReader(filePath, Encoding.UTF8))
                    {
                        string line;
                        bool isFirstLine = true;
                        while ((line = sr.ReadLine()) != null)
                        {
                            string[] values = line.Split(',');
                            if (isFirstLine)
                            {
                                for (int i = 0; i < values.Length; i++)
                                {
                                    dt.Columns.Add("Column" + i);
                                }
                                isFirstLine = false;
                            }
                            // Chỉ thêm nếu số phần tử khớp với số cột
                            if (values.Length <= dt.Columns.Count)
                            {
                                dt.Rows.Add(values);
                            }
                        }
                    }
                }
                catch
                {
                    throw new Exception("Không thể đọc file Excel. Vui lòng kiểm tra lại file hoặc lưu dưới dạng CSV. Chi tiết: " + ex.Message);
                }
            }
            return dt;
        }

        private void ExportToExcel(DataTable dt, string filePath)
        {
            try
            {
                if (filePath.EndsWith(".csv"))
                {
                    using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sw.Write(dt.Columns[i].ColumnName);
                            if (i < dt.Columns.Count - 1) sw.Write(",");
                        }
                        sw.WriteLine();

                        foreach (DataRow row in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                sw.Write(row[i].ToString());
                                if (i < dt.Columns.Count - 1) sw.Write(",");
                            }
                            sw.WriteLine();
                        }
                    }
                }
                else
                {
                    // Sử dụng EPPlus để xuất Excel
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Danh sách Lớp");
                        
                        // Ghi header
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = dt.Columns[i].ColumnName;
                            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                            worksheet.Cells[1, i + 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            worksheet.Cells[1, i + 1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(246, 145, 71));
                            worksheet.Cells[1, i + 1].Style.Font.Color.SetColor(System.Drawing.Color.White);
                        }
                        
                        // Ghi dữ liệu
                        for (int i = 0; i < dt.Rows.Count; i++)
                        {
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                worksheet.Cells[i + 2, j + 1].Value = dt.Rows[i][j].ToString();
                            }
                        }
                        
                        // Auto fit columns
                        worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                        
                        // Lưu file
                        FileInfo fileInfo = new FileInfo(filePath);
                        package.SaveAs(fileInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Lỗi khi xuất file: " + ex.Message);
            }
        }
    }
}
