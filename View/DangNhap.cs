using Microsoft.Office.Interop.Excel;
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

namespace projectQLTV
{
    public partial class DangNhap : Form
    {
        public DangNhap()
        {
            InitializeComponent();
        }
        SqlConnection con = new SqlConnection("Data Source=DUC;Initial Catalog=dbqltv;Integrated Security=True");
        public void connect()
        {
            if (con.State == ConnectionState.Closed)
            {
                con.Open();
            }
        }


        private void guna2PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void guna2Button1_Click(object sender, EventArgs e){
            string user = txtUser.Text;
            string pass = txtPass.Text;
            try{
                connect();
                if (user == "" || pass == ""){
                    MessageBox.Show("Vui lòng nhập đầy đủ tài khoản và mật khẩu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string sql = "SELECT count(*) FROM NhanVien WHERE Tendangnhap= '"+user.Trim()+"' AND Matkhau = '"+pass.Trim()+"'";
                SqlCommand cmd = new SqlCommand(sql, con);
                int ketqua = (int)cmd.ExecuteScalar();
                cmd.Dispose();
                con.Close();
                if (ketqua > 0){
                    this.Hide();
                    TrangChu main = new TrangChu();
                    main.Show();
                }
                else{
                    MessageBox.Show("Sai tài khoản hoặc mật khẩu!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }catch (Exception ex){
                MessageBox.Show("Lỗi kết nối cơ sở dữ liệu: " + ex.Message, "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void guna2Button2_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
