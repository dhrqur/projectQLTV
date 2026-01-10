create database dbqltv
use dbqltv
Create table ChiTietMuonTra(
    MaMT nvarchar(10) not null,
    MaSach nvarchar(10) not null,
    SoLuong int not null
)
Create table DocGia(
    MaDG nvarchar(10) primary key,
    MaKhoa nvarchar (10) not null,
    MaLop nvarchar(10) not null,
    TenDG nvarchar(50) not null,
    GioiTinh nvarchar(10) not null,
    DiaChi nvarchar(100) not null,
    Email nvarchar(50) not null,
    Sdt nvarchar(13) not null check (len(Sdt) between 10 AND 13)
)

Create table MuonTra(
    MaMT nvarchar(10) primary key,
    MaDG nvarchar(10) not null,
    MaNV nvarchar(10) not null,
    NgayMuon date default cast(getdate() as date),
    HanTra date not null default DATEADD(MONTH, 3, CAST(GETDATE() AS DATE)),
    TrangThai nvarchar(20) default N'Chưa trả'
)
create table NhanVien(
    MaNV nvarchar(10) primary key,
    TenNV nvarchar(50) not null,
    QueQuan nvarchar(50) not null,
    GioiTinh nvarchar(10) not null,
    VaiTro nvarchar(30) not null,
    DiaChi nvarchar(100) not null,
    Email nvarchar(50) not null,
    Sdt nvarchar(13) check (len(Sdt) between 10 AND 13),
    Tendangnhap nvarchar(30) not null,
    Matkhau nvarchar(30) not null
    )

create table NhaXuatBan(
    MaNXB nvarchar(10) primary key,
    TenNXB nvarchar(100) not null,
    DiaChi nvarchar(100) not null,
    Email nvarchar(50) not null,
    Sdt nvarchar(13) check (Len(Sdt) between 10 and 13)
)

create table PhieuTra(
    MaPT nvarchar(10) primary key,
    MaMT nvarchar(10) not null,
    NgayTra date default cast(getdate() as date)
   )
create table Sach(
MaSach nvarchar(10) primary key,
MaTG nvarchar(10) not null,
MaNXB nvarchar(10) not null,
MaTL nvarchar(10) not null,
TenSach nvarchar(100) not null,
NamXB int not null,
SoLuong int null ,
MoTa nvarchar(max) ,
MaNN nvarchar(10) not null,
MaKe nvarchar(10) not null
)
create table TacGia(
    MaTG nvarchar(10) primary key,
    TenTG nvarchar(50) not null,
    NamSinh int,
    GioiTinh nvarchar(20) not null,
    QuocTich nvarchar(30) not null
)
create table TheLoai(
    MaTL nvarchar(10) primary key,
    TenTL nvarchar(50) not null
)
CREATE TABLE NgonNgu (
    MaNN nvarchar(10) PRIMARY KEY,
    TenNN NVARCHAR(100) NOT NULL
)
    CREATE TABLE Khoa (
    MaKhoa nvarchar(10) PRIMARY KEY,
    TenKhoa NVARCHAR(100) NOT NULL,
    MoTa NVARCHAR(255)
);
CREATE TABLE Lop (
    MaLop nvarchar(10) PRIMARY KEY,
    TenLop NVARCHAR(100) NOT NULL,
    MaKhoa nvarchar(10)
);
Create table PhieuDatTruoc(
    MaPhieuDat nvarchar(10) primary key,
    MaDG nvarchar(10) not null,
    MaSach nvarchar(10) not null,
    NgayDat date not null,
    TrangThai nvarchar (20) default N'Đang chờ',
    NgayHetHan date default DATEADD(day, 7, CAST(GETDATE() AS DATE))
)
create table TheThuVien(
    MaThe nvarchar(10) primary key,
    MaDG nvarchar(10) not null,
    NgayCap date not null,
    NgayHetHan date not null,
    TrangThai nvarchar(20) not null
)
create table KeSach(
    MaKe nvarchar(10) primary key,
    TenKe nvarchar(50) not null
)
alter table DocGia
add
constraint fk_dg_lop
foreign key(MaLop)
references Lop(MaLop),

constraint fk_dg_khoa
foreign key (MaKhoa)
references Khoa(MaKhoa)

alter table Lop
add
constraint fk_lop_khoa
foreign key (MaKhoa)
references Khoa(MaKhoa)

alter table MuonTra
add
constraint fk_mt_nv
foreign key (MaNV)
references NhanVien(MaNV),

constraint fk_mt_dg
foreign key (MaDG)
references DocGia(MaDG)

alter table PhieuDatTruoc
add
constraint fk_pdt_s
foreign key (MaSach)
references Sach(MaSach),

constraint fk_pdt_dg
foreign key (MaDG)
references DocGia(MaDG)

alter table PhieuTra
add
constraint fk_pt_mt
foreign key(MaMT)
references MuonTra(MaMT)

alter table ChiTietMuonTra
add
constraint fk_ctmt_s
foreign key (MaSach)
references Sach(MaSach),

constraint fk_ctms_mt
foreign key (MaMT)
references MuonTra(MaMT)

alter table Sach
add
constraint fk_s_tg
foreign key (MaTG)
references TacGia(MaTG),

constraint fk_s_nxb
foreign key (MaNXB)
references NhaXuatBan(MaNXB),

constraint fk_s_tl
foreign key (MaTL)
references TheLoai(MaTL),

constraint fk_s_nn
foreign key (MaNN)
references NgonNgu(MaNN),

constraint fk_s_vt
foreign key(MaKe)
references KeSach(MaKe)

alter table TheThuVien
add
constraint fk_ttv_dg_
foreign key (MaDG)
references DocGia(MaDG)

INSERT INTO NhanVien (MaNV, TenNV, QueQuan, GioiTinh, VaiTro, DiaChi, Email, Sdt, Tendangnhap, Matkhau) VALUES 
(N'NV001', N'Nguyễn Thị Hương', N'Hà Nội', N'Nữ', N'Admin', N'Cầu Giấy, Hà Nội', 'huong.nt@utt.edu.vn', '0987654321', N'admin', N'123456'), 
(N'NV002', N'Phạm Văn Dũng', N'Nam Định', N'Nam', N'Thủ thư', N'Đống Đa, Hà Nội', 'dung.pv@utt.edu.vn', '0978123456', N'thuthu01', N'123456'),
(N'NV003', N'Lê Thị Mai', N'Hải Phòng', N'Nữ', N'Thủ thư', N'Ba Đình, Hà Nội', 'mai.lt@utt.edu.vn', '0903123456', N'thuthu02', N'123456'),
(N'NV004', N'Trần Quang Huy', N'Bắc Ninh', N'Nam', N'Quản lý kho', N'Nam Từ Liêm, Hà Nội', 'huy.tq@utt.edu.vn', '0912345678', N'kho01', N'123456'), 
(N'NV005', N'Đặng Thu Trang', N'Hà Nam', N'Nữ', N'CSKH - Mượn trả', N'Thanh Xuân, Hà Nội', 'trang.dt@utt.edu.vn', '0934567890', N'muontra01', N'123456');
/* =========================
   1. KHOA
========================= */
INSERT INTO Khoa (MaKhoa, TenKhoa, MoTa) VALUES
(N'K001', N'Công nghệ thông tin', N'Đào tạo CNTT, an toàn thông tin, khoa học dữ liệu'),
(N'K002', N'Kinh tế - Vận tải', N'Kinh tế vận tải, logistics, quản trị kinh doanh'),
(N'K003', N'Cơ khí', N'Cơ khí chế tạo, cơ điện tử, kỹ thuật ô tô'),
(N'K004', N'Điện - Điện tử', N'Điện công nghiệp, tự động hóa, điện tử viễn thông'),
(N'K005', N'Xây dựng - Cầu đường', N'Xây dựng dân dụng, giao thông, cầu đường');

/* =========================
   2. LỚP
========================= */
INSERT INTO Lop (MaLop, TenLop, MaKhoa) VALUES
(N'L001', N'CNTT K63A', N'K001'),
(N'L002', N'CNTT K63B', N'K001'),
(N'L003', N'Logistics K63', N'K002'),
(N'L004', N'Kinh tế vận tải K63', N'K002'),
(N'L005', N'Cơ khí K62A', N'K003'),
(N'L006', N'Ô tô K62B', N'K003'),
(N'L007', N'Tự động hóa K63', N'K004'),
(N'L008', N'Cầu đường K63', N'K005');

/* =========================
   3. NGÔN NGỮ
========================= */
INSERT INTO NgonNgu (MaNN, TenNN) VALUES
(N'NN001', N'Tiếng Việt'),
(N'NN002', N'English'),
(N'NN003', N'日本語 (Tiếng Nhật)'),
(N'NN004', N'Français'),
(N'NN005', N'中文 (Tiếng Trung)');

/* =========================
   4. THỂ LOẠI
========================= */
INSERT INTO TheLoai (MaTL, TenTL) VALUES
(N'TL001', N'Giáo trình'),
(N'TL002', N'Tham khảo'),
(N'TL003', N'Kỹ thuật'),
(N'TL004', N'Kinh tế'),
(N'TL005', N'CNTT'),
(N'TL006', N'Xây dựng - Giao thông'),
(N'TL007', N'Quản trị - Kỹ năng'),
(N'TL008', N'Ngoại ngữ');

/* =========================
   5. TÁC GIẢ
========================= */
INSERT INTO TacGia (MaTG, TenTG, NamSinh, GioiTinh, QuocTich) VALUES
(N'TG001', N'Nguyễn Văn An', 1978, N'Nam', N'Việt Nam'),
(N'TG002', N'Trần Thị Minh', 1982, N'Nữ', N'Việt Nam'),
(N'TG003', N'Phạm Đức Hùng', 1975, N'Nam', N'Việt Nam'),
(N'TG004', N'Lê Quang Vũ', 1980, N'Nam', N'Việt Nam'),
(N'TG005', N'Đặng Thu Hà', 1986, N'Nữ', N'Việt Nam'),
(N'TG006', N'Robert C. Martin', 1952, N'Nam', N'Mỹ'),
(N'TG007', N'Andrew S. Tanenbaum', 1944, N'Nam', N'Hà Lan'),
(N'TG008', N'Philip Kotler', 1931, N'Nam', N'Mỹ'),
(N'TG009', N'Ngô Thời Nhiệm', 1969, N'Nam', N'Việt Nam'),
(N'TG010', N'Yukio Mishima', 1925, N'Nam', N'Nhật Bản');

/* =========================
   6. NHÀ XUẤT BẢN
========================= */
INSERT INTO NhaXuatBan (MaNXB, TenNXB, DiaChi, Email, Sdt) VALUES
(N'NXB001', N'NXB Giáo dục Việt Nam', N'81 Trần Hưng Đạo, Hà Nội', 'contact@nxbgiaduc.vn', '02439422222'),
(N'NXB002', N'NXB Khoa học & Kỹ thuật', N'70 Trần Hưng Đạo, Hà Nội', 'info@nxbkhkt.vn', '02438255555'),
(N'NXB003', N'NXB Giao thông Vận tải', N'80 Trần Hưng Đạo, Hà Nội', 'support@nxbgtvt.vn', '02439439999'),
(N'NXB004', N'NXB Thống kê', N'54 Nguyễn Chí Thanh, Hà Nội', 'lienhe@nxbtk.vn', '02437766666'),
(N'NXB005', N'NXB Trẻ', N'161B Lý Chính Thắng, TP.HCM', 'contact@nxbtre.vn', '02839312222');

/* =========================
   7. KỆ SÁCH
========================= */
INSERT INTO KeSach (MaKe, TenKe) VALUES
(N'KS001', N'Kệ A - CNTT'),
(N'KS002', N'Kệ B - Kinh tế'),
(N'KS003', N'Kệ C - Cơ khí'),
(N'KS004', N'Kệ D - Điện - Tự động hóa'),
(N'KS005', N'Kệ E - Xây dựng - GTVT'),
(N'KS006', N'Kệ F - Ngoại ngữ');

/* =========================
   8. SÁCH
========================= */
INSERT INTO Sach (MaSach, MaTG, MaNXB, MaTL, TenSach, NamXB, SoLuong, MoTa, MaNN, MaKe) VALUES
(N'S001', N'TG007', N'NXB002', N'TL005', N'Mạng máy tính', 2019, 8, N'Giáo trình mạng máy tính.', N'NN001', N'KS001'),
(N'S002', N'TG006', N'NXB002', N'TL005', N'Clean Code', 2020, 6, N'Mã sạch.', N'NN002', N'KS001'),
(N'S003', N'TG007', N'NXB002', N'TL005', N'Hệ điều hành hiện đại', 2018, 5, N'Hệ điều hành.', N'NN001', N'KS001'),
(N'S004', N'TG008', N'NXB004', N'TL004', N'Marketing căn bản', 2021, 10, N'Marketing.', N'NN001', N'KS002'),
(N'S005', N'TG003', N'NXB003', N'TL006', N'Kỹ thuật giao thông', 2020, 4, N'Giao thông.', N'NN001', N'KS005'),
(N'S006', N'TG002', N'NXB003', N'TL006', N'Thiết kế đường ô tô', 2019, 7, N'Đường ô tô.', N'NN001', N'KS005'),
(N'S007', N'TG004', N'NXB002', N'TL003', N'Kỹ thuật nhiệt', 2017, 3, N'Truyền nhiệt.', N'NN001', N'KS003'),
(N'S008', N'TG001', N'NXB002', N'TL003', N'Cơ học kỹ thuật', 2016, 9, N'Cơ học.', N'NN001', N'KS003'),
(N'S009', N'TG005', N'NXB001', N'TL001', N'Toán cao cấp', 2015, 12, N'Giáo trình toán.', N'NN001', N'KS001'),
(N'S010', N'TG010', N'NXB005', N'TL008', N'Tiếng Nhật căn bản', 2022, 6, N'Học tiếng Nhật.', N'NN003', N'KS006');

-- ===== DOC GIA =====
INSERT INTO DocGia (MaDG, MaKhoa, MaLop, TenDG, GioiTinh, DiaChi, Email, Sdt) VALUES
(N'DG001', N'K01', N'L01', N'Nguyễn Minh Tuấn', N'Nam', N'Phường Dịch Vọng, Cầu Giấy, Hà Nội', 'tuan.nm.k63a@sv.utt.edu.vn', '0398123456'),
(N'DG002', N'K01', N'L02', N'Trần Thùy Linh', N'Nữ', N'Phường Mỹ Đình 2, Nam Từ Liêm, Hà Nội', 'linh.tt.k63b@sv.utt.edu.vn', '0389123456'),
(N'DG003', N'K02', N'L03', N'Phạm Đức Long', N'Nam', N'Phường Mộ Lao, Hà Đông, Hà Nội', 'long.pd.log@sv.utt.edu.vn', '0977001122'),
(N'DG004', N'K02', N'L04', N'Vũ Ngọc Anh', N'Nữ', N'Phường Láng Thượng, Đống Đa, Hà Nội', 'anh.vn.ktvt@sv.utt.edu.vn', '0868123123'),
(N'DG005', N'K03', N'L05', N'Đỗ Quang Minh', N'Nam', N'Phường Ngọc Khánh, Ba Đình, Hà Nội', 'minh.dq.ck62a@sv.utt.edu.vn', '0919002233'),
(N'DG006', N'K03', N'L06', N'Lê Hoàng Nam', N'Nam', N'Phường Phú Thượng, Tây Hồ, Hà Nội', 'nam.lh.ot62b@sv.utt.edu.vn', '0933004455'),
(N'DG007', N'K04', N'L07', N'Nguyễn Thị Hạnh', N'Nữ', N'Phường Khương Trung, Thanh Xuân, Hà Nội', 'hanh.nt.tdh@sv.utt.edu.vn', '0966007788'),
(N'DG008', N'K05', N'L08', N'Trịnh Gia Bảo', N'Nam', N'Phường Yên Hòa, Cầu Giấy, Hà Nội', 'bao.tgb.cd63@sv.utt.edu.vn', '0978123987'),
(N'DG009', N'K01', N'L01', N'Hoàng Khánh Ly', N'Nữ', N'Phường Trung Hòa, Cầu Giấy, Hà Nội', 'ly.hk.cntt@sv.utt.edu.vn', '0834567812'),
(N'DG010', N'K02', N'L03', N'Ngô Đức Thành', N'Nam', N'Phường Văn Quán, Hà Đông, Hà Nội', 'thanh.nd.log@sv.utt.edu.vn', '0702345678'),
(N'DG011', N'K04', N'L07', N'Bùi Thu Hà', N'Nữ', N'Phường Đại Kim, Hoàng Mai, Hà Nội', 'ha.bt.tdh@sv.utt.edu.vn', '0356789123'),
(N'DG012', N'K05', N'L08', N'Đinh Quang Hùng', N'Nam', N'Phường Minh Khai, Hai Bà Trưng, Hà Nội', 'hung.dq.cd@sv.utt.edu.vn', '0845678901');
GO

-- ===== THE THU VIEN =====
INSERT INTO TheThuVien (MaThe, MaDG, NgayCap, NgayHetHan, TrangThai) VALUES
(N'TTV001', N'DG001', '2024-09-10', '2028-09-10', N'Đang hoạt động'),
(N'TTV002', N'DG002', '2024-09-10', '2028-09-10', N'Đang hoạt động'),
(N'TTV003', N'DG003', '2024-10-01', '2028-10-01', N'Đang hoạt động'),
(N'TTV004', N'DG004', '2024-10-01', '2028-10-01', N'Đang hoạt động'),
(N'TTV005', N'DG005', '2023-09-15', '2027-09-15', N'Đang hoạt động'),
(N'TTV006', N'DG006', '2023-09-15', '2027-09-15', N'Đang hoạt động'),
(N'TTV007', N'DG007', '2024-11-05', '2028-11-05', N'Đang hoạt động'),
(N'TTV008', N'DG008', '2024-11-05', '2028-11-05', N'Đang hoạt động'),
(N'TTV009', N'DG009', '2024-09-20', '2028-09-20', N'Đang hoạt động'),
(N'TTV010', N'DG010', '2024-09-20', '2028-09-20', N'Đang hoạt động'),
(N'TTV011', N'DG011', '2024-12-15', '2028-12-15', N'Đang hoạt động'),
(N'TTV012', N'DG012', '2024-12-15', '2028-12-15', N'Tạm khóa');
GO

-- ===== MUON TRA =====
INSERT INTO MuonTra (MaMT, MaDG, MaNV, NgayMuon, HanTra, TrangThai) VALUES
(N'MT001', N'DG001', N'NV002', '2025-10-10', '2026-01-10', N'Chưa trả'),
(N'MT002', N'DG002', N'NV002', '2025-09-20', '2025-12-20', N'Đã trả'),
(N'MT003', N'DG003', N'NV003', '2025-11-05', '2026-02-05', N'Chưa trả'),
(N'MT004', N'DG004', N'NV003', '2025-08-15', '2025-11-15', N'Quá hạn'),
(N'MT005', N'DG005', N'NV002', '2025-12-01', '2026-03-01', N'Chưa trả'),
(N'MT006', N'DG006', N'NV005', '2025-07-10', '2025-10-10', N'Đã trả'),
(N'MT007', N'DG007', N'NV005', '2025-12-20', '2026-03-20', N'Chưa trả'),
(N'MT008', N'DG008', N'NV002', '2025-09-01', '2025-12-01', N'Đã trả'),
(N'MT009', N'DG009', N'NV003', '2025-10-25', '2026-01-25', N'Chưa trả'),
(N'MT010', N'DG010', N'NV005', '2025-06-01', '2025-09-01', N'Đã trả'),
(N'MT011', N'DG011', N'NV002', '2025-09-28', '2025-12-28', N'Quá hạn'),
(N'MT012', N'DG012', N'NV003', '2025-12-28', '2026-03-28', N'Chưa trả');
GO

-- ===== CHI TIET MUON TRA =====
INSERT INTO ChiTietMuonTra (MaMT, MaSach, SoLuong) VALUES
(N'MT001', N'S001', 1),
(N'MT001', N'S003', 1),

(N'MT002', N'S002', 1),
(N'MT002', N'S020', 1),

(N'MT003', N'S004', 1),
(N'MT003', N'S016', 1),

(N'MT004', N'S005', 1),
(N'MT004', N'S006', 1),

(N'MT005', N'S008', 1),

(N'MT006', N'S007', 1),
(N'MT006', N'S009', 1),

(N'MT007', N'S014', 1),
(N'MT007', N'S015', 1),

(N'MT008', N'S011', 1),

(N'MT009', N'S018', 1),

(N'MT010', N'S010', 1),
(N'MT010', N'S012', 1),

(N'MT011', N'S013', 1),

(N'MT012', N'S019', 1);
GO

-- ===== PHIEU TRA (chỉ cho các MT đã trả) =====
INSERT INTO PhieuTra (MaPT, MaMT, NgayTra) VALUES
(N'PT001', N'MT002', '2025-12-18'),
(N'PT002', N'MT006', '2025-10-05'),
(N'PT003', N'MT008', '2025-11-28'),
(N'PT004', N'MT010', '2025-08-28');
GO

-- ===== PHIEU DAT TRUOC =====
INSERT INTO PhieuDatTruoc (MaPhieuDat, MaDG, MaSach, NgayDat, TrangThai, NgayHetHan) VALUES
(N'PD001', N'DG001', N'S018', '2026-01-03', N'Đang chờ', '2026-01-10'),
(N'PD002', N'DG004', N'S005', '2025-12-20', N'Hết hạn', '2025-12-27'),
(N'PD003', N'DG007', N'S002', '2026-01-05', N'Đang chờ', '2026-01-12'),
(N'PD004', N'DG009', N'S004', '2025-12-28', N'Đã hủy',  '2026-01-04'),
(N'PD005', N'DG010', N'S011', '2025-12-30', N'Đang chờ', '2026-01-06');
GO
