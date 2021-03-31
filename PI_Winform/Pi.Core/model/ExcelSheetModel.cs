using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pi.Core
{
    public class ExcelSheetModel
    {
        [DisplayName("Tên sheet")]
        public string SheetName { get; set; }
        [DisplayName("Tên cột cần tính")]
        public string ColumnName { get; set; }
        [DisplayName("Chỉ STT cột cần tính")]
        public int ColumnIndex { get; set; }
    }

    public class ExcelCalculatorModel
    {
        [DisplayName("Tên sheet")]
        public string SheetName { get; set; }
        [DisplayName("Tên cột cần tính")]
        public string ColumnName { get; set; }
        [DisplayName("Chỉ STT cột cần tính")]
        public int ColumnIndex { get; set; }
        [DisplayName("Giá trị nhỏ nhất")]
        public double MinValue { get; set; }
        [DisplayName("Giá trị lớn nhất")]
        public double MaxValue { get; set; }
        [DisplayName("Giá trị lớn nhất cuối cùng")]
        public double MaxValueFinal { get; set; }
    }
}
