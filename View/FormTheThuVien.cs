using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using projectQLTV.Helper;
using OfficeOpenXml;

namespace projectQLTV.View
{
    public partial class FormTheThuVien : Form
    {
        private string currentMaThe = "";

        public FormTheThuVien()
        {
            InitializeComponent();
            WireUpEvents();
            LoadDocGiaComboBox();
            LoadTrangThaiComboBox();
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
            dgvTheThuVien.CellClick += dgvTheThuVien_CellClick;
        }

        private void LoadDocGiaComboBox()
        {
            try
            {
                string query = "SELECT MaDG, TenDG FROM DocGia";
                DataTable dt = DatabaseHelper.ExecuteQuery(query);
                cboDocGia.Items.Clear();
                cboDocGia.Items.Add("-- Chọn Độc Giả --");
                foreach (DataRow row in dt.Rows)
                {
                    cboDocGia.Items.Add(row["MaDG"].ToString() + " - " + row["TenDG"].ToString());
                }
                if (cboDocGia.Items.Count > 0)
                    cboDocGia.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải danh sách độc giả: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadTrangThaiComboBox()
        {
            cboTrangThai.Items.Clear();
            cboTrangThai.Items.Add("Hoạt động");
            cboTrangThai.Items.Add("Hết hạn");
            cboTrangThai.Items.Add("Khóa");
            if (cboTrangThai.Items.Count > 0)
                cboTrangThai.SelectedIndex = 0;
        }

        private void LoadData()
        {
            try
            {
                string query = @"SELECT t.MaThe, t.MaDG, d.TenDG, t.NgayCap, t.NgayHetHan, t.TrangThai 
                                 FROM TheThuVien t 
                                 LEFT JOIN DocGia d ON t.MaDG = d.MaDG";
                DataTable dt = DatabaseHelper.ExecuteQuery(query);
                dgvTheThuVien.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tải dữ liệu: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GenerateMaThe()
        {
            try
            {
                string query = "SELECT MAX(CAST(SUBSTRING(MaThe, 2, LEN(MaThe)) AS INT)) FROM TheThuVien WHERE MaThe LIKE 'T%'";
                object result = DatabaseHelper.ExecuteScalar(query);
                int nextNumber = 1;
                if (result != null && result != DBNull.Value)
                {
                    nextNumber = Convert.ToInt32(result) + 1;
                }
                return "T" + nextNumber.ToString("D03");
            }
            catch
            {
                return "T001";
            }
        }

        private void ClearFields()
        {
            txtMaThe.Text = ""; // MaThe
            cboDocGia.SelectedIndex = 0;
            cboTrangThai.SelectedIndex = 0;
            dtpNgayCap.Value = DateTime.Now; // NgayCap (using DateTimePicker)
            dtpNgayHetHan.Value = DateTime.Now.AddYears(4); // NgayHetHan (default 4 years from now)
            currentMaThe = "";
        }

        private bool ValidateInput()
        {
            if (cboDocGia.SelectedIndex <= 0 || cboDocGia.SelectedItem == null)
            {
                MessageBox.Show("Vui lòng chọn Độc Giả!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cboDocGia.Focus();
                return false;
            }

            DateTime ngayCap = dtpNgayCap.Value;
            DateTime ngayHetHan = dtpNgayHetHan.Value;

            if (ngayHetHan <= ngayCap)
            {
                MessageBox.Show("Ngày Hết Hạn phải lớn hơn Ngày Cấp!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                dtpNgayHetHan.Focus();
                return false;
            }

            return true;
        }

        private string GetMaDGFromComboBox()
        {
            if (cboDocGia.SelectedItem != null)
            {
                string selected = cboDocGia.SelectedItem.ToString();
                if (selected.Contains(" - "))
                {
                    return selected.Split('-')[0].Trim();
                }
            }
            return "";
        }

        private bool CheckDuplicate(string maThe, string maDG)
        {
            try
            {
                string query = "SELECT COUNT(*) FROM TheThuVien WHERE (MaThe = @MaThe OR (MaDG = @MaDG AND TrangThai = N'Hoạt động')) AND MaThe != @CurrentMaThe";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaThe", maThe),
                    new SqlParameter("@MaDG", maDG),
                    new SqlParameter("@CurrentMaThe", currentMaThe)
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

                string maThe = GenerateMaThe();
                string maDG = GetMaDGFromComboBox();
                DateTime ngayCap = dtpNgayCap.Value;
                DateTime ngayHetHan = dtpNgayHetHan.Value;
                string trangThai = cboTrangThai.SelectedItem?.ToString() ?? "Hoạt động";

                if (string.IsNullOrEmpty(maDG))
                {
                    MessageBox.Show("Vui lòng chọn Độc Giả hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (CheckDuplicate(maThe, maDG))
                {
                    MessageBox.Show("Mã Thẻ đã tồn tại hoặc Độc Giả đã có thẻ đang hoạt động!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "INSERT INTO TheThuVien (MaThe, MaDG, NgayCap, NgayHetHan, TrangThai) VALUES (@MaThe, @MaDG, @NgayCap, @NgayHetHan, @TrangThai)";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaThe", maThe),
                    new SqlParameter("@MaDG", maDG),
                    new SqlParameter("@NgayCap", ngayCap),
                    new SqlParameter("@NgayHetHan", ngayHetHan),
                    new SqlParameter("@TrangThai", trangThai)
                };

                int result = DatabaseHelper.ExecuteNonQuery(query, parameters);
                if (result > 0)
                {
                    MessageBox.Show("Thêm thẻ thư viện thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi thêm thẻ thư viện: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSua_Click(object sender, EventArgs e) // Sửa
        {
            try
            {
                if (string.IsNullOrEmpty(currentMaThe))
                {
                    MessageBox.Show("Vui lòng chọn thẻ cần sửa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (!ValidateInput())
                    return;

                string maDG = GetMaDGFromComboBox();
                DateTime ngayCap = dtpNgayCap.Value;
                DateTime ngayHetHan = dtpNgayHetHan.Value;
                string trangThai = cboTrangThai.SelectedItem?.ToString() ?? "Hoạt động";

                if (string.IsNullOrEmpty(maDG))
                {
                    MessageBox.Show("Vui lòng chọn Độc Giả hợp lệ!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string query = "UPDATE TheThuVien SET MaDG = @MaDG, NgayCap = @NgayCap, NgayHetHan = @NgayHetHan, TrangThai = @TrangThai WHERE MaThe = @MaThe";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@MaThe", currentMaThe),
                    new SqlParameter("@MaDG", maDG),
                    new SqlParameter("@NgayCap", ngayCap),
                    new SqlParameter("@NgayHetHan", ngayHetHan),
                    new SqlParameter("@TrangThai", trangThai)
                };

                int result = DatabaseHelper.ExecuteNonQuery(query, parameters);
                if (result > 0)
                {
                    MessageBox.Show("Sửa thẻ thư viện thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadData();
                    ClearFields();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi sửa thẻ thư viện: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnXoa_Click(object sender, EventArgs e) // Xóa
        {
            try
            {
                if (string.IsNullOrEmpty(currentMaThe))
                {
                    MessageBox.Show("Vui lòng chọn thẻ cần xóa!", "Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult result = MessageBox.Show("Bạn có chắc chắn muốn xóa thẻ này?", "Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    string query = "DELETE FROM TheThuVien WHERE MaThe = @MaThe";
                    SqlParameter[] parameters = new SqlParameter[] { new SqlParameter("@MaThe", currentMaThe) };

                    int deleteResult = DatabaseHelper.ExecuteNonQuery(query, parameters);
                    if (deleteResult > 0)
                    {
                        MessageBox.Show("Xóa thẻ thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadData();
                        ClearFields();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi xóa thẻ: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnLamMoi_Click(object sender, EventArgs e) // Làm mới
        {
            ClearFields();
            LoadDocGiaComboBox();
            LoadData();
        }

        private void btnTimKiem_Click(object sender, EventArgs e) // Tìm kiếm
        {
            try
            {
                string searchText = txtTimKiem.Text.Trim();
                string query = @"SELECT t.MaThe, t.MaDG, d.TenDG, t.NgayCap, t.NgayHetHan, t.TrangThai 
                                 FROM TheThuVien t 
                                 LEFT JOIN DocGia d ON t.MaDG = d.MaDG
                                 WHERE t.MaThe LIKE @Search OR t.MaDG LIKE @Search OR d.TenDG LIKE @Search OR t.TrangThai LIKE @Search";
                SqlParameter[] parameters = new SqlParameter[]
                {
                    new SqlParameter("@Search", "%" + searchText + "%")
                };
                DataTable dt = DatabaseHelper.ExecuteQuery(query, parameters);
                dgvTheThuVien.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Lỗi khi tìm kiếm: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void dgvTheThuVien_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dgvTheThuVien.Rows[e.RowIndex];
                currentMaThe = row.Cells[0].Value.ToString();
                txtMaThe.Text = currentMaThe;
                
                // Set DocGia combo box
                string maDG = row.Cells[1].Value?.ToString() ?? "";
                for (int i = 0; i < cboDocGia.Items.Count; i++)
                {
                    if (cboDocGia.Items[i].ToString().StartsWith(maDG + " - "))
                    {
                        cboDocGia.SelectedIndex = i;
                        break;
                    }
                }

                // Set dates
                if (row.Cells[3].Value != null)
                {
                    DateTime ngayCap = Convert.ToDateTime(row.Cells[3].Value);
                    dtpNgayCap.Value = ngayCap;
                }
                if (row.Cells[4].Value != null)
                {
                    DateTime ngayHetHan = Convert.ToDateTime(row.Cells[4].Value);
                    dtpNgayHetHan.Value = ngayHetHan;
                }

                // Set TrangThai combo box
                string trangThai = row.Cells[5].Value?.ToString() ?? "";
                for (int i = 0; i < cboTrangThai.Items.Count; i++)
                {
                    if (cboTrangThai.Items[i].ToString() == trangThai)
                    {
                        cboTrangThai.SelectedIndex = i;
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
                            if (firstRowFirstCol.Contains("mã") || firstRowFirstCol.Contains("ma") || 
                                firstRowFirstCol.Contains("ngày") || firstRowFirstCol.Contains("ngay") ||
                                firstRowFirstCol.Contains("trạng") || firstRowFirstCol.Contains("trang"))
                            {
                                startRow = 1; // Skip header row
                            }
                        }

                        for (int i = startRow; i < dt.Rows.Count; i++)
                        {
                            DataRow row = dt.Rows[i];
                            string maDG = "";
                            object ngayCapObj = null;
                            object ngayHetHanObj = null;
                            string trangThai = "Hoạt động";
                            
                            try
                            {
                                // Kiểm tra số cột hợp lệ
                                if (row.ItemArray.Length < 1)
                                    continue;
                                    
                                maDG = row[0] != null ? row[0].ToString().Trim() : "";
                                ngayCapObj = row.ItemArray.Length > 1 ? row[1] : null;
                                ngayHetHanObj = row.ItemArray.Length > 2 ? row[2] : null;
                                trangThai = row.ItemArray.Length > 3 && row[3] != null ? row[3].ToString().Trim() : "Hoạt động";

                                if (string.IsNullOrWhiteSpace(maDG) || ngayCapObj == null || ngayCapObj == DBNull.Value || 
                                    ngayHetHanObj == null || ngayHetHanObj == DBNull.Value)
                                    continue;

                                // Parse ngày và chỉ lấy phần ngày (không có giờ)
                                DateTime ngayCap = DateTime.MinValue;
                                DateTime ngayHetHan = DateTime.MinValue;
                                bool parseNgayCap = false;
                                bool parseNgayHetHan = false;
                                
                                // Xử lý ngày cấp: nếu là DateTime thì dùng trực tiếp, nếu là string thì parse
                                if (ngayCapObj is DateTime dateTimeCap)
                                {
                                    ngayCap = dateTimeCap.Date;
                                    parseNgayCap = true;
                                }
                                else
                                {
                                    string ngayCapStr = ngayCapObj.ToString().Trim();
                                    // Thử parse với nhiều format và culture
                                    string[] dateFormats = { "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy", "M/d/yyyy" };
                                    CultureInfo viCulture = new CultureInfo("vi-VN");
                                    CultureInfo enCulture = new CultureInfo("en-US");
                                    
                                    if (DateTime.TryParseExact(ngayCapStr, dateFormats, viCulture, DateTimeStyles.None, out DateTime tempNgayCap) ||
                                        DateTime.TryParseExact(ngayCapStr, dateFormats, enCulture, DateTimeStyles.None, out tempNgayCap) ||
                                        DateTime.TryParse(ngayCapStr, viCulture, DateTimeStyles.None, out tempNgayCap) ||
                                        DateTime.TryParse(ngayCapStr, enCulture, DateTimeStyles.None, out tempNgayCap))
                                    {
                                        ngayCap = tempNgayCap.Date;
                                        parseNgayCap = true;
                                    }
                                }
                                
                                // Xử lý ngày hết hạn: nếu là DateTime thì dùng trực tiếp, nếu là string thì parse
                                if (ngayHetHanObj is DateTime dateTimeHetHan)
                                {
                                    ngayHetHan = dateTimeHetHan.Date;
                                    parseNgayHetHan = true;
                                }
                                else
                                {
                                    string ngayHetHanStr = ngayHetHanObj.ToString().Trim();
                                    // Thử parse với nhiều format và culture
                                    string[] dateFormats = { "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy", "M/d/yyyy" };
                                    CultureInfo viCulture = new CultureInfo("vi-VN");
                                    CultureInfo enCulture = new CultureInfo("en-US");
                                    
                                    if (DateTime.TryParseExact(ngayHetHanStr, dateFormats, viCulture, DateTimeStyles.None, out DateTime tempNgayHetHan) ||
                                        DateTime.TryParseExact(ngayHetHanStr, dateFormats, enCulture, DateTimeStyles.None, out tempNgayHetHan) ||
                                        DateTime.TryParse(ngayHetHanStr, viCulture, DateTimeStyles.None, out tempNgayHetHan) ||
                                        DateTime.TryParse(ngayHetHanStr, enCulture, DateTimeStyles.None, out tempNgayHetHan))
                                    {
                                        ngayHetHan = tempNgayHetHan.Date;
                                        parseNgayHetHan = true;
                                    }
                                }
                                
                                if (!parseNgayCap || !parseNgayHetHan)
                                {
                                    errorCount++;
                                    continue;
                                }

                                // Kiểm tra độc giả tồn tại
                                string checkDGQuery = "SELECT COUNT(*) FROM DocGia WHERE MaDG = @MaDG";
                                SqlParameter[] checkDGParams = new SqlParameter[] { new SqlParameter("@MaDG", maDG) };
                                int dgCount = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkDGQuery, checkDGParams));

                                if (dgCount == 0)
                                {
                                    errorCount++;
                                    continue;
                                }

                                // Kiểm tra trùng: chỉ kiểm tra nếu trạng thái là "Hoạt động"
                                // Nếu trạng thái khác (Khóa, Hết hạn) thì cho phép import
                                if (trangThai == "Hoạt động")
                                {
                                    string checkQuery = "SELECT COUNT(*) FROM TheThuVien WHERE MaDG = @MaDG AND TrangThai = N'Hoạt động'";
                                    SqlParameter[] checkParams = new SqlParameter[] { new SqlParameter("@MaDG", maDG) };
                                    int count = Convert.ToInt32(DatabaseHelper.ExecuteScalar(checkQuery, checkParams));

                                    if (count > 0)
                                    {
                                        errorCount++;
                                        continue;
                                    }
                                }

                                // Insert dữ liệu
                                string maThe = GenerateMaThe();
                                string insertQuery = "INSERT INTO TheThuVien (MaThe, MaDG, NgayCap, NgayHetHan, TrangThai) VALUES (@MaThe, @MaDG, @NgayCap, @NgayHetHan, @TrangThai)";
                                SqlParameter[] parameters = new SqlParameter[]
                                {
                                    new SqlParameter("@MaThe", maThe),
                                    new SqlParameter("@MaDG", maDG),
                                    new SqlParameter("@NgayCap", ngayCap),
                                    new SqlParameter("@NgayHetHan", ngayHetHan),
                                    new SqlParameter("@TrangThai", trangThai)
                                };
                                DatabaseHelper.ExecuteNonQuery(insertQuery, parameters);
                                successCount++;
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
                saveFile.FileName = "DanhSachTheThuVien_" + DateTime.Now.ToString("yyyyMMdd_HHmmss");

                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    DataTable dt = (DataTable)dgvTheThuVien.DataSource;
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
                        dt.Columns.Add("Column0"); // MaDG
                        dt.Columns.Add("Column1"); // NgayCap
                        dt.Columns.Add("Column2"); // NgayHetHan
                        dt.Columns.Add("Column3"); // TrangThai

                        int rowCount = worksheet.Dimension?.End.Row ?? 0;
                        int colCount = worksheet.Dimension?.End.Column ?? 0;
                        
                        // Phát hiện cấu trúc file Excel
                        // File có thể có: MaThe(1), MaDG(2), TenDG(3), NgayCap(4), NgayHetHan(5), TrangThai(6)
                        // Hoặc: MaDG(1), TenDG(2), NgayCap(3), NgayHetHan(4), TrangThai(5)
                        int maDGCol = 1; // Cột MaDG mặc định
                        int ngayCapCol = 3; // Cột NgayCap mặc định (bỏ qua TenDG)
                        int ngayHetHanCol = 4; // Cột NgayHetHan mặc định
                        int trangThaiCol = 5; // Cột TrangThai mặc định
                        
                        if (rowCount > 0)
                        {
                            // Kiểm tra cột đầu tiên có phải là MaThe không
                            string firstCell = worksheet.Cells[1, 1].Text?.Trim().ToLower() ?? "";
                            if (firstCell.Contains("mã thẻ") || firstCell.Contains("ma the") || 
                                firstCell.Contains("mathe") || firstCell == "mã" || firstCell == "ma" || 
                                (firstCell.StartsWith("t0") && firstCell.Length <= 4))
                            {
                                // File có MaThe, điều chỉnh các cột
                                maDGCol = 2;
                                ngayCapCol = 4; // Bỏ qua TenDG ở cột 3
                                ngayHetHanCol = 5;
                                trangThaiCol = 6;
                            }
                            else
                            {
                                // File không có MaThe
                                maDGCol = 1;
                                ngayCapCol = 3; // Bỏ qua TenDG ở cột 2
                                ngayHetHanCol = 4;
                                trangThaiCol = 5;
                            }
                        }
                        
                        // Phát hiện header và dòng bắt đầu đọc dữ liệu
                        int startRow = 1;
                        if (rowCount > 0)
                        {
                            string firstCell = worksheet.Cells[1, maDGCol].Text?.Trim().ToLower() ?? "";
                            // Kiểm tra xem dòng đầu tiên có phải là header không
                            if (firstCell.Contains("mã độc giả") || firstCell.Contains("ma doc gia") || 
                                firstCell.Contains("madg") || firstCell.Contains("ngày cấp") || 
                                firstCell.Contains("ngay cap") || firstCell.Contains("ngaycap") ||
                                firstCell == "mã" || firstCell == "ma" || firstCell.StartsWith("dg"))
                            {
                                startRow = 2; // Bỏ qua dòng header
                            }
                        }
                        
                        for (int row = startRow; row <= rowCount; row++)
                        {
                            // Đọc MaDG từ cột maDGCol
                            string cell1 = maDGCol <= colCount ? worksheet.Cells[row, maDGCol].Text?.Trim() ?? "" : "";
                            object cell2Value = null;
                            object cell3Value = null;
                            // Đọc TrangThai từ cột trangThaiCol
                            string cell4 = trangThaiCol <= colCount ? worksheet.Cells[row, trangThaiCol].Text?.Trim() ?? "" : "";
                            
                            // Đọc NgayCap từ cột ngayCapCol
                            if (ngayCapCol <= colCount)
                            {
                                var cell2 = worksheet.Cells[row, ngayCapCol];
                                if (cell2.Value != null)
                                {
                                    if (cell2.Value is DateTime dateTime2)
                                    {
                                        cell2Value = dateTime2.Date; // Chỉ lấy phần ngày
                                    }
                                    else
                                    {
                                        // Parse từ string với nhiều format
                                        string dateStr = cell2.Text?.Trim() ?? "";
                                        string[] dateFormats = { "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy", "M/d/yyyy" };
                                        CultureInfo viCulture = new CultureInfo("vi-VN");
                                        CultureInfo enCulture = new CultureInfo("en-US");
                                        
                                        if (DateTime.TryParseExact(dateStr, dateFormats, viCulture, DateTimeStyles.None, out DateTime parsedDate2) ||
                                            DateTime.TryParseExact(dateStr, dateFormats, enCulture, DateTimeStyles.None, out parsedDate2) ||
                                            DateTime.TryParse(dateStr, viCulture, DateTimeStyles.None, out parsedDate2) ||
                                            DateTime.TryParse(dateStr, enCulture, DateTimeStyles.None, out parsedDate2))
                                        {
                                            cell2Value = parsedDate2.Date; // Chỉ lấy phần ngày
                                        }
                                        else
                                        {
                                            cell2Value = dateStr; // Giữ nguyên string nếu không parse được
                                        }
                                    }
                                }
                            }
                            
                            // Đọc NgayHetHan từ cột ngayHetHanCol
                            if (ngayHetHanCol <= colCount)
                            {
                                var cell3 = worksheet.Cells[row, ngayHetHanCol];
                                if (cell3.Value != null)
                                {
                                    if (cell3.Value is DateTime dateTime3)
                                    {
                                        cell3Value = dateTime3.Date; // Chỉ lấy phần ngày
                                    }
                                    else
                                    {
                                        // Parse từ string với nhiều format
                                        string dateStr = cell3.Text?.Trim() ?? "";
                                        string[] dateFormats = { "dd/MM/yyyy", "MM/dd/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM-dd-yyyy", "d/M/yyyy", "M/d/yyyy" };
                                        CultureInfo viCulture = new CultureInfo("vi-VN");
                                        CultureInfo enCulture = new CultureInfo("en-US");
                                        
                                        if (DateTime.TryParseExact(dateStr, dateFormats, viCulture, DateTimeStyles.None, out DateTime parsedDate3) ||
                                            DateTime.TryParseExact(dateStr, dateFormats, enCulture, DateTimeStyles.None, out parsedDate3) ||
                                            DateTime.TryParse(dateStr, viCulture, DateTimeStyles.None, out parsedDate3) ||
                                            DateTime.TryParse(dateStr, enCulture, DateTimeStyles.None, out parsedDate3))
                                        {
                                            cell3Value = parsedDate3.Date; // Chỉ lấy phần ngày
                                        }
                                        else
                                        {
                                            cell3Value = dateStr; // Giữ nguyên string nếu không parse được
                                        }
                                    }
                                }
                            }
                            
                            // Bỏ qua dòng trống hoàn toàn
                            if (string.IsNullOrWhiteSpace(cell1) && (cell2Value == null || string.IsNullOrWhiteSpace(cell2Value.ToString())) && 
                                (cell3Value == null || string.IsNullOrWhiteSpace(cell3Value.ToString())) && 
                                string.IsNullOrWhiteSpace(cell4))
                                continue;
                            
                            dt.Rows.Add(cell1, cell2Value ?? "", cell3Value ?? "", cell4);
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
                                string columnName = dt.Columns[i].ColumnName;
                                object cellValue = row[i];
                                
                                // Format ngày chỉ hiển thị ngày, không có giờ
                                if ((columnName == "NgayCap" || columnName == "NgayHetHan") && cellValue != null && cellValue != DBNull.Value)
                                {
                                    if (DateTime.TryParse(cellValue.ToString(), out DateTime dateValue))
                                    {
                                        sw.Write(dateValue.ToString("dd/MM/yyyy"));
                                    }
                                    else
                                    {
                                        sw.Write(cellValue.ToString());
                                    }
                                }
                                else
                                {
                                    sw.Write(cellValue?.ToString() ?? "");
                                }
                                
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
                        ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Danh sách Thẻ Thư Viện");
                        
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
                                string columnName = dt.Columns[j].ColumnName;
                                object cellValue = dt.Rows[i][j];
                                
                                // Format ngày chỉ hiển thị ngày, không có giờ
                                if ((columnName == "NgayCap" || columnName == "NgayHetHan") && cellValue != null && cellValue != DBNull.Value)
                                {
                                    if (DateTime.TryParse(cellValue.ToString(), out DateTime dateValue))
                                    {
                                        worksheet.Cells[i + 2, j + 1].Value = dateValue.ToString("dd/MM/yyyy");
                                    }
                                    else
                                    {
                                        worksheet.Cells[i + 2, j + 1].Value = cellValue.ToString();
                                    }
                                }
                                else
                                {
                                    worksheet.Cells[i + 2, j + 1].Value = cellValue?.ToString() ?? "";
                                }
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
