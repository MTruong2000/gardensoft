using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.IO;
using GardenSoft.views;
using System.Runtime.Remoting.Messaging;

namespace GardenSoft
{
    public partial class Form1 : Form, IKH
    {
        private string conStr = "Server=DESKTOP-AH3TGNG;Database=QLKH;Trusted_Connection=True;";
        List<string> maIDList = new List<string>();

        public string MaID {get{ return txtMaID.Text; } set { txtMaID.Text = value; } }
        public string Ten { get { return txtTen.Text; } set { txtTen.Text = value; } }
        public DateTime NgaySinh { get { return dtPickerNgaySinh.Value; } set { dtPickerNgaySinh.Value = value; } }
        public string DiaChi { get { return txtDiaChi.Text; } set { txtDiaChi.Text = value; } }
        public string PassPort { get { return txtPassPort.Text; } set { txtPassPort.Text = value; } }
        public DateTime NgayCap { get { return dtPickerNgayCap.Value; } set { dtPickerNgayCap.Value = value; } }
        public string DienThoai { get { return txtDienThoai.Text; } set { txtDienThoai.Text = value; } }
        public string DiDong { get { return txtDiDong.Text; } set { txtDiDong.Text = value; } }
        public string Fax { get { return txtFax.Text; } set { txtFax.Text = value; } }
        public string Email { get { return txtEmail.Text; } set { txtEmail.Text = value; } }
        public string TaiKhoanNH { get { return txtTKNH.Text; } set { txtTKNH.Text = value; } }
        public string TenNH { get { return txtTNH.Text; } set { txtTNH.Text = value; } }
        public string LoaiKH {
            get { return cbbLKH.SelectedItem.ToString(); }
            set { cbbLKH.SelectedIndex = cbbLKH.FindString(value); }
        }
        public string HanTT { get { return txtHTT.Text; } set { txtHTT.Text = value; } }

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
            else if (result == DialogResult.No) MessageBox.Show("Xin mời bạn tiếp tục chương trình!");
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
            if (txtMaID.Text == "" || txtTen.Text == "" || txtDiDong.Text == "")
            {
                MessageBox.Show("Bạn nhập còn thiếu");
            }
            else
            {
                using (SqlConnection con = new SqlConnection(conStr))
                {
                    con.Open();

                    string sqlStrMaID = "SELECT MaID FROM KHACHHANG";
                    using (SqlCommand comm = new SqlCommand(sqlStrMaID, con))
                    {
                        using (SqlDataReader reader = comm.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string maID = reader.GetString(reader.GetOrdinal("MaID"));
                                maIDList.Add(maID);
                            }
                        }
                    }

                    if (maIDList.Exists(x => string.Equals(x, txtMaID.Text, StringComparison.OrdinalIgnoreCase)))
                    {
                        MessageBox.Show("Mã ID đã tồn tại");
                    }
                    else
                    {
                        string sqlStr = "INSERT INTO KHACHHANG (MaID, TEN, NgaySinh, DiaChi, PassPort, NgayCap, DienThoai, DiDong, Fax, Email, TaiKhoanNH, TenNH, LoaiKH, HanTT) " +
                            "VALUES(@MaID, @Ten, @NgaySinh, @DiaChi, @PassPort, @NgayCap, @DienThoai, @DiDong, @Fax, @Email, @TaiKhoanNH, @TenNH, @LoaiKH, @HanTT)";

                        using (SqlCommand comm = new SqlCommand(sqlStr, con))
                        {
                            comm.Parameters.AddWithValue("@MaID", MaID);
                            comm.Parameters.AddWithValue("@Ten", Ten);
                            comm.Parameters.AddWithValue("@NgaySinh", NgaySinh);
                            comm.Parameters.AddWithValue("@DiaChi", DiaChi);
                            comm.Parameters.AddWithValue("@PassPort", PassPort);
                            comm.Parameters.AddWithValue("@NgayCap", NgayCap);
                            comm.Parameters.AddWithValue("@DienThoai", DienThoai);
                            comm.Parameters.AddWithValue("@DiDong", DiDong);
                            comm.Parameters.AddWithValue("@Fax", Fax);
                            comm.Parameters.AddWithValue("@Email", Email);
                            comm.Parameters.AddWithValue("@TaiKhoanNH", TaiKhoanNH);
                            comm.Parameters.AddWithValue("@TenNH", TenNH);
                            comm.Parameters.AddWithValue("@LoaiKH", LoaiKH);
                            comm.Parameters.AddWithValue("@HanTT", HanTT);

                            comm.ExecuteNonQuery();
                        }


                        MessageBox.Show("Bạn đã thêm thành công!!!");
                        txtMaID.Text = "";
                        txtTen.Text = "";
                        dtPickerNgaySinh.Value = DateTime.Now;
                        txtDiaChi.Text = "";
                        txtPassPort.Text = "";
                        dtPickerNgayCap.Value = DateTime.Now;
                        txtDienThoai.Text = "";
                        txtDiDong.Text = "";
                        txtFax.Text = "";
                        txtEmail.Text = "";
                        txtTKNH.Text = "";
                        txtTNH.Text = "";
                        txtHTT.Text = "";
                    }
                }
            }
        }

        private void btnNhapLieu_Click(object sender, EventArgs e)
        {
            pnNhapLieu.Visible = true;
            pnImpFile.Visible = false;
        }

        private void btnImpFile_Click(object sender, EventArgs e)
        {
            pnNhapLieu.Visible = false;
            pnImpFile.Visible = true;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show(
                "Bạn có muốn export file",
                "Confirm",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                using (SqlConnection connection = new SqlConnection(conStr))
                {
                    connection.Open();

                    string sqlQuery = "SELECT * FROM KHACHHANG";

                    using (SqlCommand command = new SqlCommand(sqlQuery, connection))
                    {
                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            ToExcel(reader, "QLKH");
                        }
                    }
                }
            }
            else if (result == DialogResult.No)
            {
                MessageBox.Show("Xin mời bạn tiếp tục chương trình!");
            }
        }

        private void ToExcel(SqlDataReader reader, string baseFileName)
        {
            Microsoft.Office.Interop.Excel.Application excel;
            Microsoft.Office.Interop.Excel.Workbook workbook;
            Microsoft.Office.Interop.Excel.Worksheet worksheet;

            try
            {
                excel = new Microsoft.Office.Interop.Excel.Application();
                excel.Visible = false;
                excel.DisplayAlerts = false;

                string fileName = GetUniqueFileName(baseFileName);
                workbook = excel.Workbooks.Add(Type.Missing);
                worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "Quản lý khách hàng";

                for (int i = 0; i < reader.FieldCount; i++)
                {
                    worksheet.Cells[1, i + 1] = reader.GetName(i);
                }

                int row = 2;

                while (reader.Read())
                {
                    for (int col = 0; col < reader.FieldCount; col++)
                    {
                        worksheet.Cells[row, col + 1] = reader.GetValue(col).ToString();
                    }
                    row++;
                }

                workbook.SaveAs(fileName);
                workbook.Close();
                excel.Quit();
                MessageBox.Show("Xuất dữ liệu ra Excel thành công!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                workbook = null;
                worksheet = null;
            }
        }

        private string GetUniqueFileName(string baseFileName)
        {
            string fileExtension = Path.GetExtension(baseFileName);

            int count = 1;
            string uniqueFileName = baseFileName;

            string path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            path = path.Replace("Documents", "Documents\\");
            while (File.Exists(path + uniqueFileName+ ".xlsx"))
            {
                uniqueFileName = Path.Combine(fileExtension, $"{baseFileName} ({count})");
                count++;
            }

            return uniqueFileName;
        }
    }
}
