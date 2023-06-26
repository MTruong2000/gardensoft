using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GardenSoft.views
{
    internal interface IKH
    {
        string MaID { get; set; }
        string Ten { get; set; }
        DateTime NgaySinh { get; set; }
        string DiaChi { get; set; }
        string PassPort { get; set; }
        DateTime NgayCap { get; set; }
        string DienThoai { get; set; }
        string DiDong { get; set; }
        string Fax { get; set; }
        string Email { get; set; }
        string TaiKhoanNH { get; set; }
        string TenNH { get; set; }
        string LoaiKH { get; set; }
        string HanTT { get; set; }
    }
}
