using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace GardenSoft
{
    public partial class Form1 : Form
    {
        private string conStr = "Server=DESKTOP-101QR58;Database=QLKH;Trusted_Connection=True;";
        private SqlConnection con;
        private SqlDataAdapter myADapter;
        private SqlCommand comm;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnThoat_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Bạn có chắc chắn thoát?",
                "Confirm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (result == DialogResult.Yes) Close();
            else if (result == DialogResult.No) MessageBox.Show("Bạn đã hủy thoát chương trình!");
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult resultClosing = MessageBox.Show(
                "Bạn có chắc chắn thoát?",
                "Confirm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (resultClosing == DialogResult.Yes)
            {
                e.Cancel = false;
            }
            else if (resultClosing == DialogResult.No)
            {
                e.Cancel = true;
            }
        }

        private void btnLuu_Click(object sender, EventArgs e)
        {
            con = new SqlConnection(conStr);
            con.Open();

            string sqlStr = "INSERT INTO KHACHHANG (MaID, TEN) VALUES('" + txtMaID.Text + "',N'" + txtTen.Text + "')";
            comm = new SqlCommand(sqlStr, con);
            comm.ExecuteNonQuery();

            con.Close();
            MessageBox.Show("Bạn đã thêm thành công !!!");
            txtMaID.Text = "";
            txtTen.Text = "";
        }
    }
}
