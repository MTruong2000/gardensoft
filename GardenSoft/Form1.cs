using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.IO;

namespace GardenSoft
{
    public partial class Form1 : Form
    {
        private string conStr = "Server=DESKTOP-AH3TGNG;Database=QLKH;Trusted_Connection=True;";
        List<string> maIDList = new List<string>();

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
                            comm.Parameters.AddWithValue("@MaID", txtMaID.Text);
                            comm.Parameters.AddWithValue("@Ten", txtTen.Text);
                            comm.Parameters.AddWithValue("@NgaySinh", dtPickerNgaySinh.Value);
                            comm.Parameters.AddWithValue("@DiaChi", txtDiaChi.Text);
                            comm.Parameters.AddWithValue("@PassPort", txtPassPort.Text);
                            comm.Parameters.AddWithValue("@NgayCap", dtPickerNgayCap.Value);
                            comm.Parameters.AddWithValue("@DienThoai", txtDienThoai.Text);
                            comm.Parameters.AddWithValue("@DiDong", txtDiDong.Text);
                            comm.Parameters.AddWithValue("@Fax", txtFax.Text);
                            comm.Parameters.AddWithValue("@Email", txtEmail.Text);
                            comm.Parameters.AddWithValue("@TaiKhoanNH", txtTKNH.Text);
                            comm.Parameters.AddWithValue("@TenNH", txtTNH.Text);
                            comm.Parameters.AddWithValue("@LoaiKH", cbbLKH.SelectedItem.ToString());
                            comm.Parameters.AddWithValue("@HanTT", txtHTT.Text);

                            comm.ExecuteNonQuery();
                        }


                        MessageBox.Show("Bạn đã thêm thành công!!!");
                        txtMaID.Text = "";
                        txtTen.Text = "";
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
