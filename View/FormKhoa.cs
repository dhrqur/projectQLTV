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
    public partial class FormKhoa : Form
    {
        private string currentMaKhoa = "";

        public FormKhoa()
        {
            InitializeComponent();
            WireUpEvents();
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
            dgvKhoa.CellClick += dgvKhoa_CellClick;
        }

        private void LoadData()
        {
            try
            {
                string query = "SELECT MaKhoa, TenKhoa, MoTa FROM Khoa";
                DataTable dt = DatabaseHelper.ExecuteQuery(query);
                dgvKhoa.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GenerateMaKhoa()
        {
            try
            {
                string query = "SELECT MAX(CAST(SUBSTRING(MaKhoa, 2, LEN(MaKhoa)) AS INT)) FROM Khoa WHERE MaKhoa LIKE 'K%'";
                object result = DatabaseHelper.ExecuteScalar(query);
                int nextNumber = 1;
                if (result != null && result != DBNull.Value)
                {
                    nextNumber = Convert.ToInt32(result) + 1;
                }
                return "K" + nextNumber.ToString("D03");
            }
            catch
            {
                return "K001";
            }
        }

        private void ClearFields()
        {
            txtMaKhoa.Text = ""; // MaKhoa
            txtTenKhoa.Text = ""; // TenKhoa
            txtMoTa.Text = ""; // MoTa
            currentMaKhoa = "";
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtTenKhoa.Text))
            {
                MessageBox.Show("Vui lòng nhập Tên Khoa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTenKhoa.Focus();
                return false;
            }

            return true;
        }

        private bool CheckDuplicate(string maKhoa, string tenKhoa)
        {
            try
            {
                string query = "SELECT COUNT(*) FROM Khoa WHERE (MaKhoa = @MaKhoa OR TenKhoa = @TenKhoa) AND MaKhoa != @CurrentMaKhoa";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaKhoa", maKhoa),
                    new SqlParameter("@TenKhoa", tenKhoa),
                    new SqlParameter("@CurrentMaKhoa", currentMaKhoa)
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

                string maKhoa = GenerateMaKhoa();
                string tenKhoa = txtTenKhoa.Text.Trim();
                string moTa = txtMoTa.Text.Trim();

                if (CheckDuplicate(maKhoa, tenKhoa))
                {
                    MessageBox.Show("Mã Khoa hoặc Tên Khoa đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "INSERT INTO Khoa (MaKhoa, TenKhoa, MoTa) VALUES (@MaKhoa, @TenKhoa, @MoTa)";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaKhoa", maKhoa),
                    new SqlParameter("@TenKhoa", tenKhoa),
                    new SqlParameter("@MoTa", moTa)
                };

                int result = DatabaseHelper.ExecuteNonQuery(query, parameters);
                if (result > 0)
                {
                    MessageBox.Show("Thêm khoa thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm khoa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSua_Click(object sender, EventArgs e) // Sửa
        {
            try
            {
                if (string.IsNullOrEmpty(currentMaKhoa))
                {
                    MessageBox.Show("Vui lòng chọn khoa cần sửa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!ValidateInput())
                    return;

                string tenKhoa = txtTenKhoa.Text.Trim();
                string moTa = txtMoTa.Text.Trim();

                if (CheckDuplicate(currentMaKhoa, tenKhoa))
                {
                    MessageBox.Show("Tên Khoa đã tồn tại!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "UPDATE Khoa SET TenKhoa = @TenKhoa, MoTa = @MoTa WHERE MaKhoa = @MaKhoa";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaKhoa", currentMaKhoa),
                    new SqlParameter("@TenKhoa", tenKhoa),
                    new SqlParameter("@MoTa", moTa)
                };

                int result = DatabaseHelper.ExecuteNonQuery(query, parameters);
                if (result > 0)
                {
                    MessageBox.Show("Sửa khoa thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi sửa khoa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnXoa_Click(object sender, EventArgs e) // Xóa
        {
            try
            {
                if (string.IsNullOrEmpty(currentMaKhoa))
                {
                    MessageBox.Show("Vui lòng chọn khoa cần xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Kiểm tra xem khoa có đang được sử dụng không
                string checkQuery = "SELECT COUNT(*) FROM Lop WHERE MaKhoa = @MaKhoa";
                SqlParameter[] checkParams = new SqlParameter[] { new SqlParameter("@MaKhoa", currentMaKhoa) };
                int count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkQuery, checkParams));
                
                if (count > 0)
                {
                    MessageBox.Show("Không thể xóa khoa này vì đang có lớp sử dụng!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa khoa này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    string query = "DELETE FROM Khoa WHERE MaKhoa = @MaKhoa";
                    SqlParameter[] parameters = new SqlParameter[] { new SqlParameter("@MaKhoa", currentMaKhoa) };

                    int deleteResult = DatabaseHelper.ExecuteNonQuery(query, parameters);
                    if (deleteResult > 0)
                    {
                        MessageBox.Show("Xóa khoa thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData();
                        ClearFields();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa khoa: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLamMoi_Click(object sender, EventArgs e) // Làm mới
        {
            ClearFields();
            LoadData();
        }

        private void btnTimKiem_Click(object sender, EventArgs e) // Tìm kiếm
        {
            try
            {
                string searchText = txtTimKiem.Text.Trim();
                string query = "SELECT MaKhoa, TenKhoa, MoTa FROM Khoa WHERE MaKhoa LIKE @Search OR TenKhoa LIKE @Search OR MoTa LIKE @Search";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Search", "%" + searchText + "%")
                };
                DataTable dt = DatabaseHelper.ExecuteQuery(query, parameters);
                dgvKhoa.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgvKhoa_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvKhoa.Rows[e.RowIndex];
                currentMaKhoa = row.Cells[0].Value.ToString();
                txtMaKhoa.Text = currentMaKhoa;
                txtTenKhoa.Text = row.Cells[1].Value.ToString();
                txtMoTa.Text = row.Cells[2].Value?.ToString() ?? "";
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
                                firstRowFirstCol.Contains("mô tả") || firstRowFirstCol.Contains("ten") || 
                                firstRowFirstCol.Contains("ma") || firstRowFirstCol == "tên khoa" || 
                                firstRowFirstCol == "ten khoa")
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
                                    
                                string tenKhoa = row[0] != null ? row[0].ToString().Trim() : "";
                                string moTa = row.ItemArray.Length > 1 && row[1] != null ? row[1].ToString().Trim() : "";

                                if (string.IsNullOrWhiteSpace(tenKhoa))
                                    continue;

                                // Kiểm tra trùng
                                string checkQuery = "SELECT COUNT(*) FROM Khoa WHERE TenKhoa = @TenKhoa";
                                SqlParameter[] checkParams = new SqlParameter[] { new SqlParameter("@TenKhoa", tenKhoa) };
                                int count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkQuery, checkParams));

                                if (count == 0)
                                {
                                    string maKhoa = GenerateMaKhoa();
                                    string insertQuery = "INSERT INTO Khoa (MaKhoa, TenKhoa, MoTa) VALUES (@MaKhoa, @TenKhoa, @MoTa)";
                                    SqlParameter[] parameters = new SqlParameter[]
                                    {
                                        new SqlParameter("@MaKhoa", maKhoa),
                                        new SqlParameter("@TenKhoa", tenKhoa),
                                        new SqlParameter("@MoTa", moTa)
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
                saveFile.FileName = "DanhSachKhoa_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = (DataTable)dgvKhoa.DataSource;
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
                            // Chỉ thêm nếu số phần tử khớp với số cột
                            if (values.Length <= dt.Columns.Count)
                            {
                                dt.Rows.Add(values);
                            }
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
                        dt.Columns.Add("Column0"); // TenKhoa
                        dt.Columns.Add("Column1"); // MoTa

                        int rowCount = worksheet.Dimension?.End.Row ?? 0;
                        int colCount = worksheet.Dimension?.End.Column ?? 0;
                        
                        // Phát hiện header và cột bắt đầu
                        int startRow = 1; // Mặc định bắt đầu từ dòng 1
                        int startCol = 1; // Mặc định bắt đầu từ cột 1
                        
                        if (rowCount > 0)
                        {
                            string firstCell = worksheet.Cells[1, 1].Text?.Trim().ToLower() ?? "";
                            // Kiểm tra xem dòng đầu tiên có phải là header không
                            if (firstCell.Contains("mã khoa") || firstCell.Contains("ma khoa") || 
                                firstCell.Contains("makhoa") || firstCell == "mã" || firstCell == "ma")
                            {
                                startRow = 2; // Bỏ qua dòng header, đọc từ dòng 2
                                startCol = 2; // Bỏ qua cột mã khoa, đọc từ cột TenKhoa (cột 2)
                            }
                            else if (firstCell.StartsWith("k0") || firstCell.StartsWith("k1") || firstCell.StartsWith("k2"))
                            {
                                // Dòng đầu tiên là dữ liệu, không phải header
                                startRow = 1;
                                startCol = 2; // Bỏ qua cột mã khoa, đọc từ cột TenKhoa
                            }
                        }
                        
                        // Đọc dữ liệu từ dòng startRow
                        for (int row = startRow; row <= rowCount; row++)
                        {
                            string cell1 = worksheet.Cells[row, startCol].Text?.Trim() ?? "";
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
                    throw new Exception("Không thể đọc file Excel. Chi tiết: " + ex.Message);
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
                        // Write headers
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sw.Write(dt.Columns[i].ColumnName);
                            if (i < dt.Columns.Count - 1) sw.Write(",");
                        }
                        sw.WriteLine();

                        // Write data
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
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Danh sách Khoa");
                        
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
