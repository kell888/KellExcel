using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Data;
using System.Data.OleDb;

namespace KellExcel
{
    /// <summary>
    /// Sheet索引
    /// 作    者： KELL
    /// 日    期： 2007-5-18 18:10:00
    /// </summary>
    public enum ExcelSheetIndex : uint
    {
        CurrentSheet = 0,
        Sheet1 = 1,
        Sheet2 = 2,
        Sheet3 = 3,
        Sheet4 = 4,
        Sheet5 = 5,
        Sheet6 = 6,
        Sheet7 = 7,
        Sheet8 = 8,
        Sheet9 = 9,
        Sheet10 = 10,
        Sheet11 = 11,
        Sheet12 = 12,
        Sheet13 = 13,
        Sheet14 = 14,
        Sheet15 = 15,
        Sheet16 = 16,
        Sheet17 = 17,
        Sheet18 = 18,
        Sheet19 = 19,
        Sheet20 = 20,
        Sheet21 = 21,
        Sheet22 = 22,
        Sheet23 = 23,
        Sheet24 = 24,
        Sheet25 = 25,
        Sheet26 = 26,
        Sheet27 = 27,
        Sheet28 = 28,
        Sheet29 = 29,
        Sheet30 = 30,
        Sheet31 = 31,
        Sheet32 = 32
    }
    
    /// <summary>
    /// Excel写入类型
    /// 作    者： KELL
    /// 日    期： 2007-5-18 18:10:00
    /// </summary>
    public enum ExcelWriteType//只使用了ReWrite类型，其它的类型留待以后扩展
    {
        /// <summary>
        /// 未知的写入类型
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        None = 0,
        /// <summary>
        /// 重写
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        ReWrite = 1,
        /// <summary>
        /// 追加
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        Append = 2,
        /// <summary>
        /// 插入
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        Insert = 3,
    }

    /// <summary>
    /// 单元格行号和列号索引结构
    /// </summary>
    public struct CellIndexs
    {
        /// <summary>
        /// 从1开始的行号
        /// </summary>
        public int Row;
        /// <summary>
        /// 从1开始的列号
        /// </summary>
        public int Col;
        /// <summary>
        /// 将CellIndexs结构转化为可读字符串
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return Row.ToString() + "," + Col.ToString();
        }
    }
    /*
    /// <summary>
    /// Excel写入语言
    /// 作    者： KELL
    /// 日    期： 2007-5-18 18:10:00
    /// </summary>
    public enum Language
    {
        /// <summary>
        /// 简体
        /// </summary>
        SimplifiedChinese = 0,
        /// <summary>
        /// 繁体
        /// </summary>
        TraditionalChinese = 1
    }
    
    /// <summary>
    /// 边框线样式枚举
    /// </summary>
    public enum BorderLineStyle
    {
        Continue,
        Dash,
        DashDot,
        DashDotDot,
        Dot,
        Double,
        LineStyleNone,
        SlantDashDot
    }
    /// <summary>
    /// 边框线厚度枚举
    /// </summary>
    public enum BorderBorderWeight
    {
        Hairline,
        Medium,
        Thick,
        Thin
    }
    /// <summary>
    /// 边框线颜色枚举
    /// </summary>
    public enum BorderColorIndex
    {
        ColorIndexAutoMatic,
        ColorIndexNone
    }
    /// <summary>
    /// 边框样式结构
    /// </summary>
    public struct BorderStyle
    {
        public Excel.XlLineStyle LineStyle;
        public Excel.XlBorderWeight BorderWeight;
        public Excel.XlColorIndex ColorIndex;
    }*/
    /// <summary>
    /// Excel操作类
    /// </summary>
    public class MyExcel
    {
        /// <summary>
        /// 最大Sheet数
        /// </summary>
        public const int MaxSheetCount = 32;
        private string filePath = "";
        DateTime beforeTime, afterTime;
        Excel.ApplicationClass app;
        Excel.Workbook wb;
        Excel.Worksheet ws;
        Excel.TextBox tb;
        Excel.Range rng;
        System.Drawing.Color backColor = System.Drawing.Color.White;
        System.Drawing.Color foreColor = System.Drawing.Color.Black;
        System.Drawing.Font font = new System.Drawing.Font("宋体", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(134)), false);
        bool isLink = false;
        string linkFile = "";
        string linkSheet = "Sheet1";
        string linkCell = "A1";
        string sheetName = "Sheet1";
        string backgroundImage = "";
        ExcelSheetIndex sheetIndex = ExcelSheetIndex.Sheet1;
        ExcelWriteType writeType = ExcelWriteType.ReWrite;
        //Language wordLanguage = Language.SimplifiedChinese;
        /*BorderStyle borderStyle;
        /// <summary>
        /// 设置边框线样式
        /// </summary>
        public BorderLineStyle BLineStyle
        {
            set
            {
                borderStyle.LineStyle = GetLineStyle(value);
            }
        }
        /// <summary>
        /// 根据外部索引获取内部索引（LineStyle）
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private Excel.XlLineStyle GetLineStyle(BorderLineStyle value)
        {
            Excel.XlLineStyle ret = Excel.XlLineStyle.xlContinuous;
            switch (value)
            {
                case BorderLineStyle.Continue:
                    ret = Excel.XlLineStyle.xlContinuous;
                    break;
                case BorderLineStyle.Dash:
                    ret = Excel.XlLineStyle.xlDash;
                    break;
                case BorderLineStyle.DashDot:
                    ret = Excel.XlLineStyle.xlDashDot;
                    break;
                case BorderLineStyle.DashDotDot:
                    ret = Excel.XlLineStyle.xlDashDotDot;
                    break;
                case BorderLineStyle.Dot:
                    ret = Excel.XlLineStyle.xlDot;
                    break;
                case BorderLineStyle.Double:
                    ret = Excel.XlLineStyle.xlDouble;
                    break;
                case BorderLineStyle.LineStyleNone:
                    ret = Excel.XlLineStyle.xlLineStyleNone;
                    break;
                case BorderLineStyle.SlantDashDot:
                    ret = Excel.XlLineStyle.xlSlantDashDot;
                    break;
            }
            return ret;
        }
        /// <summary>
        /// 设置边框线厚度
        /// </summary>
        public BorderBorderWeight BBorderWeight
        {
            set
            {
                borderStyle.BorderWeight = GetBorderWeight(value);
            }
        }
        /// <summary>
        /// 根据外部索引获取内部索引（BorderWeight）
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private Excel.XlBorderWeight GetBorderWeight(BorderBorderWeight value)
        {
            Excel.XlBorderWeight ret = Excel.XlBorderWeight.xlThin;
            switch (value)
            {
                case BorderBorderWeight.Hairline:
                    ret = Excel.XlBorderWeight.xlHairline;
                    break;
                case BorderBorderWeight.Medium:
                    ret = Excel.XlBorderWeight.xlMedium;
                    break;
                case BorderBorderWeight.Thick:
                    ret = Excel.XlBorderWeight.xlThick;
                    break;
                case BorderBorderWeight.Thin:
                    ret = Excel.XlBorderWeight.xlThin;
                    break;
            }
            return ret;
        }
        /// <summary>
        /// 设置边框线颜色
        /// </summary>
        public BorderColorIndex BColorIndex
        {
            set
            {
                borderStyle.ColorIndex = GetColorIndex(value);
            }
        }
        /// <summary>
        /// 根据外部索引获取内部索引（ColorIndex）
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private Excel.XlColorIndex GetColorIndex(BorderColorIndex value)
        {
            Excel.XlColorIndex ret = Excel.XlColorIndex.xlColorIndexAutomatic;
            switch (value)
            {
                case BorderColorIndex.ColorIndexAutoMatic:
                    ret = Excel.XlColorIndex.xlColorIndexAutomatic;
                    break;
                case BorderColorIndex.ColorIndexNone:
                    ret = Excel.XlColorIndex.xlColorIndexNone;
                    break;
            }
            return ret;
        }*/
        /// <summary>
        /// 构造函数
        /// </summary>
        public MyExcel()
        {
            //borderStyle.BorderWeight = Excel.XlBorderWeight.xlThin;
            //borderStyle.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            //borderStyle.LineStyle = Excel.XlLineStyle.xlContinuous;
        }
        /// <summary>
        /// 获取或设置背景颜色
        /// </summary>
        public System.Drawing.Color BackColor
        {
            get
            {
                if (backColor == System.Drawing.Color.Empty)
                    return System.Drawing.Color.White;
                return backColor;
            }
            set
            {
                backColor = value;
            }
        }
        /// <summary>
        /// 获取或设置字体颜色
        /// </summary>
        public System.Drawing.Color ForeColor
        {
            get
            {
                if (foreColor == System.Drawing.Color.Empty)
                    return System.Drawing.Color.Black;
                return foreColor;
            }
            set
            {
                foreColor = value;
            }
        }
        /// <summary>
        /// 获取或设置字体样式
        /// </summary>
        public System.Drawing.Font Font
        {
            get
            {
                return font;
            }
            set
            {
                font = value;
            }
        }
        /// <summary>
        /// 获取或设置是否为链接
        /// </summary>
        public bool IsLink
        {
            get
            {
                return isLink;
            }
            set
            {
                isLink = value;
            }
        }
        /// <summary>
        /// 获取或设置链接的目的Excel文件
        /// </summary>
        public string LinkFile
        {
            get
            {
                return linkFile;
            }
            set
            {
                linkFile = value;
            }
        }
        /// <summary>
        /// 获取或设置链接的目的工作表
        /// </summary>
        public string LinkSheet
        {
            get
            {
                return linkSheet;
            }
            set
            {
                linkSheet = value;
            }
        }
        /// <summary>
        /// 获取或设置链接的目的单元格
        /// </summary>
        public string LinkCell
        {
            get
            {
                return linkCell;
            }
            set
            {
                linkCell = value;
            }
        }
        /// <summary>
        /// 获取或设置Sheet表名称
        /// </summary>
        public string SheetName
        {
            get
            {
                return sheetName;
            }
            set//暂时还没用上此功能
            {
                if (value.Length > 32)
                {
                    throw new Exception("Sheet表名不能大于32字符，请检查！");
                }
                else
                {
                    sheetName = value;
                    ws.Name = value;
                }
            }
        }
        /// <summary>
        /// 获取或设置背景图片
        /// </summary>
        public string BackgroundImage
        {
            get
            {
                return backgroundImage;
            }
            set
            {
                if (File.Exists(value))
                {
                    ws.SetBackgroundPicture(value);
                    backgroundImage = value;
                }
                else
                {
                    throw new Exception("不存在该图片文件，请检查！");
                }
            }
        }
        /// <summary>
        /// 获取或设置Sheet表索引
        /// </summary>
        public ExcelSheetIndex SheetIndex
        {
            get
            {
                return sheetIndex;
            }
            set
            {
                if ((int)value > 32)
                {
                    throw new Exception("Sheet表索引不能大于32，请检查！");
                }
                else
                {
                    sheetIndex = value;
                }
            }
        }
        /// <summary>
        /// 获取或设置Excel文件写入类型
        /// </summary>
        public ExcelWriteType WriteType
        {
            get
            {
                return writeType;
            }
            set
            {
                writeType = value;
            }
        }
        /*
        /// <summary>
        /// 获取或设置Excel写入语言
        /// </summary>
        public Language WordLanguage
        {
            get
            {
                return wordLanguage;
            }
            set
            {
                wordLanguage = value;
            }
        }
        
        /// <summary>
        /// 获取或设置Excel边框样式
        /// </summary>
        public BorderStyle TheBorderStyle
        {
            get
            {
                return borderStyle;
            }
            set
            {
                borderStyle = value;
            }
        }*/
        /// <summary>
        /// 获取源文件路径
        /// </summary>
        public string FilePath
        {
            get
            {
                return filePath;
            }
        }
        /// <summary>
        /// 获取打开Excel应用之前的时间
        /// </summary>
        public DateTime BeforeTime
        {
            get
            {
                return beforeTime;
            }
        }
        /// <summary>
        /// 获取打开Excel应用之后的时间
        /// </summary>
        public DateTime AfterTime
        {
            get
            {
                return afterTime;
            }
        }
        /// <summary>
        /// 获取Excel应用
        /// </summary>
        public Excel.ApplicationClass Application
        {
            get
            {
                return app;
            }
        }
        /// <summary>
        /// 获取Excel工作簿
        /// </summary>
        public Excel.Workbook WorkBook
        {
            get
            {
                return wb;
            }
        }
        /// <summary>
        /// 获取或设置Excel工作表
        /// </summary>
        public Excel.Worksheet WorkSheet
        {
            get
            {
                return ws;
            }
            set
            {
                ws = value;
            }
        }
        /// <summary>
        /// 获取Excel的TextBox
        /// </summary>
        public Excel.TextBox TextBox
        {
            get
            {
                return tb;
            }
        }
        /// <summary>
        /// 获取Excel的Range
        /// </summary>
        public Excel.Range Range
        {
            get
            {
                return rng;
            }
        }
        /// <summary>
        /// 方法名称： CreateApp
        /// 内容描述： 实例化Excel对象
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public void CreateApp(bool display)
        {
            try
            {
                if (app != null)
                    Dispose();
                beforeTime = DateTime.Now;
                app = new Excel.ApplicationClass();
                app.Visible = display;
                app.DisplayAlerts = false;
                app.UserControl = true;
                afterTime = DateTime.Now;
            }
            catch
            {
                KillExcelProcess();
                throw new Exception("创建Excel应用对象出现错误，请检查你的机器！");
            }
        }
        /// <summary>
        /// 判断Excel应用是否已经创建
        /// </summary>
        public bool IsAppCreate
        {
            get
            {
                if (app != null)
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        /// <summary>
        /// 判断Excel工作簿或者工作表是否已经打开
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public bool IsOpen
        {
            get
            {
                if ((wb != null) || (ws != null))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
        /// <summary>
        /// 方法名称： FileExteCheck
        /// 内容描述： Excel文件扩展名检查
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="strpath"></param>
        private void FileExteCheck(string strpath)
        {
            //string strFile;
            string strExt;

            //strFile = strpath.Remove(strpath.Length - 4, 4);
            strExt = Path.GetExtension(strpath);

            if (strExt.ToLower() != ".xls")
            {
                KillExcelProcess();
                throw new Exception("文件格式不正确！");
            }
        }
        /// <summary>
        /// 方法名称： Open
        /// 内容描述： 无
        /// 实现流程： 打开/连接一个excel数据文档，只能指定一个存在的文件
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <returns></returns>
        public bool Open(string filepath, ExcelSheetIndex sheetInd, bool display, bool afterClear)
        {
            if (app == null)
            {
                CreateApp(display);
            }
            bool bolRetValue = false;
            filePath = filepath;
            if (!File.Exists(filepath))
            {
                KillExcelProcess();
                throw new Exception("源文件不存在，请检查！");
            }
            this.FileExteCheck(filepath);
            try
            {
                if (this.IsOpen == false)
                {
                    if (wb == null)
                        wb = app.Workbooks.Open(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//,Type.Missing,Type.Missing);
                    if (sheetInd != ExcelSheetIndex.CurrentSheet)
                        sheetIndex = sheetInd;
                    ws = (Excel.Worksheet)wb.Worksheets[(int)sheetIndex];
                    if (afterClear)
                    {
                        for (int j = 1; j <= wb.Sheets.Count; j++)
                        {
                            ((Excel.Worksheet)wb.Sheets[j]).Cells.Clear();
                            for (int i = 1; i <= ws.Shapes.Count; i++)
                            {
                                ((Excel.Worksheet)wb.Sheets[j]).Shapes.Item(i).Delete();
                            }
                            for (int i = 1; i < wb.Sheets.Count; i++)
                            {
                                Excel.Pictures pics = (Excel.Pictures)((Excel.Worksheet)wb.Sheets[j]).Pictures();
                                pics.Delete();
                            }
                        }
                    }
                    ws.Activate();
                }
                bolRetValue = true;
            }
            catch (Exception ex)
            {
                KillExcelProcess();
                throw new Exception("打开或连接Excel文档错误，请检查！\n"+ex.Message);
            }
            return bolRetValue;
        }
        /// <summary>
        /// 方法名称： OpenCreate
        /// 内容描述： 写文件时，用于文件创建及打开，可以指定不存在的文件
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <returns></returns>
        public bool OpenCreate(string filepath, ExcelSheetIndex sheetInd, bool display, bool afterClear)
        {
            bool bolRet = false;
            try
            {
                this.FileExteCheck(filepath);
                CreateApp(display);
                //如果文件不存在，则创建文件
                if (!File.Exists(filepath))
                {
                    FileStream fs = File.Create(filepath);
                    fs.Close();
                    Object filename = filepath;
                    Object missing = System.Type.Missing;
                    Object Template = System.Type.Missing;
                    wb = app.Workbooks.Add(Template);
                    ws = (Excel.Worksheet)wb.Worksheets[1];
                    wb.SaveAs(filename, missing, missing, missing, missing, missing, Excel.XlSaveAsAccessMode.xlExclusive, missing, missing, missing, missing);
                }
                bolRet = this.Open(filepath, sheetInd, display, afterClear);
            }
            catch (Exception ex)
            {
                bolRet = false;
                KillExcelProcess();
                throw new Exception("打开或创建Excel出现错误，请检查！\n" + ex.Message);
            }
            return bolRet;
        }
        /// <summary>
        /// 判断是否为数字
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static bool IsNumeric(string str)
        {
            if (str == null || str.Length == 0)
                return false;
            foreach (char c in str)
            {
                if (!Char.IsNumber(c))
                {
                    return false;
                }
            }
            return true;
        }
        /// <summary>
        /// 读写Excel时Cell合法性检查，单元格方式
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="strCell"></param>
        /// <returns></returns>
        private static bool CellCheck(string strCell)
        {
            string str = strCell.ToUpper(); //"AABB12";
            Regex r = new Regex(@"\A[A-Z]+[0-9]+\z");
            if (r.IsMatch(str))
            {
                int numIndex = GetIndexOfChrAndNum(str);
                string str1 = str.Substring(0, 1);
                string str2 = str.Substring(1, 1);
                Regex r1 = new Regex(@"\A[A-I]+\z");
                Regex r2 = new Regex(@"\A[A-V]+\z");
                if (numIndex > 2 || (numIndex == 2 && (!r1.IsMatch(str1) || !r2.IsMatch(str2))))
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            return false;
        }
        /// <summary>
        /// 读写Excel时Cell合法性检查，行，列方式
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <returns></returns>
        private static bool CellCheck(int iRow, int iCol)
        {
            return iRow > 0 && iRow <= 65536 && iCol > 0 && iCol <= 256;
        }
        /// <summary>
        /// 以单元格方式检测表示行的字符串的开始索引
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static int GetIndexOfChrAndNum(string cell)
        {
            for (int i = 0; i < cell.Length; i++)
            {
                if (IsNumeric(cell.Substring(i, 1)))
                    return i;
            }
            return 0;
        }

        /// <summary>
        /// 根据单元格字符串获取行号(1,65536)
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static int GetCellRow(string cell)
        {
            if (CellCheck(cell))
            {
                int ind = GetIndexOfChrAndNum(cell);
                string num = cell.Substring(ind);
                int iRow = Convert.ToInt32(num);
                iRow = iRow < 1 ? 1 : iRow;
                iRow = iRow > 65536 ? 65536 : iRow;
                return iRow;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 根据单元格字符串获取列号(1,256)
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="cell"></param>
        /// <returns></returns>
        public static int GetCellColume(string cell)
        {
            if (CellCheck(cell))
            {
                int ind = GetIndexOfChrAndNum(cell);
                string chr = cell.Substring(0, ind).ToUpper();
                int len = chr.Length;
                int achr1 = (int)chr.ToCharArray(0, len)[0] - 65 + 1;
                int achr2 = 0;
                int iCol = achr1;
                if (len > 1)
                {
                    achr2 = (int)chr.ToCharArray(0, len)[1] - 65 + 1;
                    iCol = 26 + (achr1 - 1) * 26 + achr2;
                }
                iCol = iCol < 1 ? 1 : iCol;
                iCol = iCol > 256 ? 256 : iCol;
                return iCol;
            }
            else
            {
                return 0;
            }
        }

        /// <summary>
        /// 根据单元格名获取单元格的行号和列号
        /// 作    者： KELL
        /// 日    期： 2008-7-21 15:05:00
        /// </summary>
        /// <param name="cell">单元格名</param>
        /// <returns></returns>
        public static CellIndexs GetCellIndexsByName(string cell)
        {
            int iRow = GetCellRow(cell);
            int iCol = GetCellColume(cell);
            CellIndexs cellIndexs = new CellIndexs();
            cellIndexs.Row = iRow;
            cellIndexs.Col = iCol;
            return cellIndexs;
        }

        /// <summary>
        /// 根据单元格的行号和列号获取单元格名
        /// 作    者： KELL
        /// 日    期： 2008-7-21 15:05:00
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <returns></returns>
        public static string GetCellNameByIndexs(int iRow, int iCol)
        {
            if (CellCheck(iRow, iCol))
            {
                string row = iRow.ToString();
                string col = "";
                if (iCol > 26)
                {
                    col = char.ConvertFromUtf32((iCol / 26 - 1) + 65) + char.ConvertFromUtf32(iCol % 26 - 1 + 65);
                }
                else
                {
                    col = char.ConvertFromUtf32(iCol - 1 + 65);
                }
                return col + row;
            }
            else
            {
                return "UnknownCell";
            }
        }

        /// <summary>
        /// 根据单元格的行号和列号索引结构获取单元格名
        /// 作    者： KELL
        /// 日    期： 2008-7-21 15:05:00
        /// </summary>
        /// <param name="cellIndexs">单元格的行号和列号索引结构</param>
        /// <returns></returns>
        public static string GetCellNameByIndexs(CellIndexs cellIndexs)
        {
            if (CellCheck(cellIndexs.Row, cellIndexs.Col))
            {
                string row = cellIndexs.Row.ToString();
                string col = "";
                if (cellIndexs.Col > 26)
                {
                    col = char.ConvertFromUtf32((cellIndexs.Col / 26 - 1) + 65) + char.ConvertFromUtf32(cellIndexs.Col % 26 - 1 + 65);
                }
                else
                {
                    col = char.ConvertFromUtf32(cellIndexs.Col - 1 + 65);
                }
                return col + row;
            }
            else
            {
                return "UnknownCell";
            }
        }

        /// <summary>
        /// 转到下一个Sheet，如果不存在则自动在当前的Sheet后添加一个新的Sheet（最多32个），并且游标的初始位置为A1，即(1, 1)
        /// 作    者： KELL
        /// 日    期： 2008-7-21 15:05:00
        /// </summary>
        /// <returns></returns>
        public void GotoNextSheet()
        {
            int current = GetCurrentSheetIndex();
            if (current < MaxSheetCount)
            {
                if (wb.Worksheets.Count < current + 1)
                {
                    wb.Worksheets.Add(Type.Missing, ws, 1, Type.Missing);
                }
                ws = (Excel.Worksheet)wb.Worksheets[current + 1];
                ws.Activate();
            }
        }

        /// <summary>
        /// 转到上一个Sheet，如果不存在则自动在当前的Sheet前插入加一个新的Sheet（最多32个），并且游标的初始位置为A1，即(1, 1)
        /// 作    者： KELL
        /// 日    期： 2011-7-19 14:09:00
        /// </summary>
        /// <returns></returns>
        public void GotoPrevSheet()
        {
            int current = GetCurrentSheetIndex();
            if (current > 0)
            {
                if (wb.Worksheets.Count < MaxSheetCount && current == 1)
                {
                    wb.Worksheets.Add(ws, Type.Missing, 1, Type.Missing);
                    ws = (Excel.Worksheet)wb.Worksheets[1];
                }
                ws = (Excel.Worksheet)wb.Worksheets[current - 1];
                ws.Activate();
            }
        }

        /// <summary>
        /// 获取当前工作表已经使用的区域（从1开始的X,Y）
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <returns></returns>
        public System.Drawing.Rectangle GetUsageCapition()
        {
            int usedRows = 0;
            int usedColumns = 0;
            if (ws.UsedRange != null && ws.UsedRange.Rows != null)
            {
                usedRows = ws.UsedRange.Rows.Count;
            }
            if (ws.UsedRange != null && ws.UsedRange.Columns != null)
            {
                usedColumns = ws.UsedRange.Columns.Count;
            }
            int startRow = 1;
            int startCol = 1;
            if (ws.UsedRange != null)
            {
                startRow = ws.UsedRange.Row;
                startCol = ws.UsedRange.Column;
            }
            System.Drawing.Rectangle usedRange = new System.Drawing.Rectangle(startCol, startRow, usedColumns, usedRows);
            return usedRange;
        }
        /// <summary>
        /// 获取已用区域的最底行索引和最右列索引(从1开始的索引)
        /// </summary>
        /// <returns></returns>
        public System.Drawing.Size GetUsageBottomAndRight()
        {
            int usedRows = 0;
            int usedColumns = 0;
            if (ws.UsedRange != null && ws.UsedRange.Rows != null)
            {
                usedRows = ws.UsedRange.Rows.Count;
            }
            if (ws.UsedRange != null && ws.UsedRange.Columns != null)
            {
                usedColumns = ws.UsedRange.Columns.Count;
            }
            return new System.Drawing.Size(ws.UsedRange.Row + usedRows, ws.UsedRange.Column + usedColumns);
        }
        /// <summary>
        /// 设置指定范围内所有单元格的字体
        /// </summary>
        /// <param name="rowBegin"></param>
        /// <param name="rowEnd"></param>
        /// <param name="colBegin"></param>
        /// <param name="colEnd"></param>
        /// <param name="font"></param>
        public void SetAllFont(int rowBegin, int rowEnd, int colBegin, int colEnd, System.Drawing.Font font)
        {
            if (font != null)
            {
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).Font.FontStyle = font.Style;
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).Font.Bold = font.Bold;
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).Font.Italic = font.Italic;
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).Font.Underline = font.Underline;
                if (font.Name != "")
                {
                    ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).Font.Name = font.Name;
                }
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).Font.Size = font.Size;
                this.WorkBook.Save();
            }
        }
        /// <summary>
        /// 设置指定范围内所有单元格的行高
        /// </summary>
        /// <param name="rowBegin"></param>
        /// <param name="rowEnd"></param>
        /// <param name="colBegin"></param>
        /// <param name="colEnd"></param>
        /// <param name="height"></param>
        public void SetAllRowHeight(int rowBegin, int rowEnd, int colBegin, int colEnd, int height)
        {
            if (height > 0)
            {
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).RowHeight = height;
                this.WorkBook.Save();
            }
        }

        /// <summary>
        /// 获取当前Workbook中已经存在了多少个Sheet
        /// </summary>
        /// <returns></returns>
        public int GetUsageSheetCount()
        {
            return this.WorkBook.Sheets.Count;
        }
        /// <summary>
        /// 获取当前工作表行高
        /// 作    者： KELL
        /// 日    期： 2007-8-2 17:10:00
        /// </summary>
        /// <returns></returns>
        public double GetCurrentSheetRowHeight()
        {
            try
            {
                double rowHeight;
                rowHeight = (double)ws.Rows.RowHeight;
                return rowHeight;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 获取当前工作表列宽
        /// 作    者： KELL
        /// 日    期： 2007-8-2 17:10:00
        /// </summary>
        /// <returns></returns>
        public double GetCurrentSheetColumnWidth()
        {
            try
            {
                double columnWidth;
                columnWidth = (double)ws.Columns.ColumnWidth;
                return columnWidth;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// 获取当前待编辑的Sheet索引
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public int GetCurrentSheetIndex()
        {
            return this.WorkSheet.Index;
        }
        /// <summary>
        /// 获取当前待编辑的Sheet名字
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        public string GetCurrentSheetName()
        {
            return this.WorkSheet.Name;
        }
        /// <summary>
        /// 获取当前Workbook中指定索引处的Sheet名字
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        public string GetSheetName(int index)
        {
            object Index = index;
            if (ExistsSheetIndex(index))
                return ((Excel.Worksheet)this.WorkBook.Sheets[Index]).Name;
            else
                throw new Exception("此索引处无Sheet！");
        }
        /// <summary>
        /// 判断当前Workbook中是否存在索引index处的Sheet
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public bool ExistsSheetIndex(int index)
        {
            return index > 0 && index <= this.WorkBook.Sheets.Count;
        }
        /// <summary>
        /// 在当前的Worksheet之后添加1个外部的Worksheet，有问题！
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="externalExcelFile">外部的Excel文件</param>
        public void AddAnExternalSheet(string externalExcelFile)
        {
            MyExcel externalExcel = new MyExcel();
            try
            {
                //if (externalExcel.OpenCreate(externalExcelFile, false))
                //{
                //externalExcel.Worksheet.Copy(Type.Missing, ws);
                //DataTable dt = Common.ReadSheet(externalExcelFile, Common.GetSheetName(externalExcelFile, 1));
                //string content = Common.ReadSheet(externalExcelFile);
                GotoNextSheet(); this.wb.MergeWorkbook(externalExcelFile);
                //WriteSheet(content);//会把原来的其他sheet内容重写掉！！！
                //WriteSheet(dt, ws.Index);
                //externalExcel.Worksheet.Copy();
                //externalExcel.Worksheet.Paste(this.wb);
                //}
            }
            catch (Exception e)
            { System.Windows.Forms.MessageBox.Show("AddAnExternalSheet:" + e.Message); }
            finally
            {
                externalExcel.Close();
            }
        }
        /// <summary>
        /// 在当前的Worksheet之前插入1个外部的Worksheet，有问题！
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="externalExcelFile">外部的Excel文件</param>
        public void InsertAnExternalSheet(string externalExcelFile)
        {
            MyExcel externalExcel = new MyExcel();
            try
            {
                //if (externalExcel.OpenCreate(externalExcelFile, false))
                //{
                //externalExcel.Worksheet.Copy(ws, Type.Missing);
                //DataTable dt = Common.ReadSheet(externalExcelFile, Common.GetSheetName(externalExcelFile, 1));
                //string content = Common.ReadSheet(externalExcelFile);
                GotoPrevSheet();this.wb.MergeWorkbook(externalExcelFile);
                //WriteSheet(content);//会把原来的其他sheet内容重写掉！！！
                //WriteSheet(dt, ws.Index);
                //externalExcel.Worksheet.Copy();
                //externalExcel.Worksheet.Paste(this.wb);
                //}
            }
            catch (Exception e)
            { System.Windows.Forms.MessageBox.Show("InsertAnExternalSheet:" + e.Message); }
            finally
            {
                externalExcel.Close();
            }
        }

        /// <summary>
        /// 读取当前Excel中的内容
        /// </summary>
        /// <returns></returns>
        public string ReadSheet()
        {
            string content = "";
            try
            {
                content = File.ReadAllText(this.filePath);
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("导出sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            return content;
        }
        /// <summary>
        /// 读取当前Excel中的指定Sheet的内容，并以DataTable的形式输出(支持Excel11.0)
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public DataTable ReadSheet(string sheetName)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + this.filePath + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
            string strExcel = "select * from [" + sheetName + "$]";
            DataTable dt = new DataTable();
            OleDbConnection conns = new OleDbConnection(strConn);
            try
            {
                conns.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, conns);
                adapter.Fill(dt);
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("导出sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            finally
            {
                conns.Close();
            }
            return dt;
        }

        /// <summary>
        /// 将指定的字符串写入到指定的excelFile文件中
        /// </summary>
        /// <param name="dt">要写入的文本</param>
        /// <returns></returns>
        public bool WriteSheet(string content)
        {
            bool flag = false;
            try
            {
                using (StreamWriter sw = File.CreateText(this.filePath))
                {
                    sw.Write(content);
                }
                flag = true;
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("写入sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            return flag;
        }

        /// <summary>
        /// 将指定的DataTable写入到当前Excel中名字为的sheetName的Sheet中
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="dt">只能写入DataTable中的文本</param>
        /// <param name="sheetIndex">从1开始的Sheet索引，默认为1</param>
        /// <param name="showColumnName">默认为false</param>
        /// <returns></returns>
        public bool WriteSheet(DataTable dt, int sheetIndex = 1, bool showColumnName = false)
        {
            bool flag = false;
            MyExcel excel = new MyExcel();
            try
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[sheetIndex];
                ws.Cells.Clear();
                int colHeight = 0;
                int col = dt.Columns.Count;
                if (showColumnName)
                {
                    colHeight = 1;
                    for (int i = 0; i < col; i++)
                    {
                        ws.Cells[1, 1 + i] = dt.Columns[i].ColumnName;
                        ws.get_Range(ws.Cells[1, 1 + i], ws.Cells[1, 1 + i]).Font.Bold = true;
                    }
                }
                if (dt.Rows.Count > 0)
                {
                    int row = dt.Rows.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = dt.Rows[i][j].ToString();
                            ws.Cells[i + 1 + colHeight, j + 1] = str;
                        }
                    }
                }
                ws.Name = dt.TableName;
                wb.Save();
                flag = true;
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("写入sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            finally
            {
                excel.Close();
            }
            return flag;
        }
        /// <summary>
        /// 在当前的Worksheet之前添加1个Sheet
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        /// <returns></returns>
        public void AddSheet()
        {
            if (this.WorkBook.Sheets.Count < MaxSheetCount)
                this.WorkBook.Sheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        }
        /// <summary>
        /// 在当前的Worksheet之前添加count个Sheet
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        /// <param name="count">添加Sheet的个数</param>
        /// <returns></returns>
        public void AddSheet(int count)
        {
            if (this.WorkBook.Sheets.Count < MaxSheetCount)
                this.WorkBook.Sheets.Add(Type.Missing, Type.Missing, count, Type.Missing);
        }
        /// <summary>
        /// 在当前的Workbook的某个Sheet后面插入count个Sheet
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        /// <param name="currentSheet">指定的Worksheet</param>
        /// <param name="count">添加Sheet的个数</param>
        /// <returns></returns>
        public void AddSheetAfter(Excel.Worksheet currentSheet, int count)
        {
            if (this.WorkBook.Sheets.Count < MaxSheetCount)
                this.WorkBook.Sheets.Add(Type.Missing, currentSheet, count, Type.Missing);
        }
        /// <summary>
        /// 在当前的Workbook的某个Sheet之前插入count个Sheet
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        /// <param name="currentSheet">指定的Worksheet</param>
        /// <param name="count">添加Sheet的个数</param>
        /// <returns></returns>
        public void AddSheetBefore(Excel.Worksheet currentSheet, int count)
        {
            if (this.WorkBook.Sheets.Count < MaxSheetCount)
                this.WorkBook.Sheets.Add(currentSheet, Type.Missing, count, Type.Missing);
        }
        /// <summary>
        /// 设置当前Worksheet的名字
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        /// <param name="name"></param>
        public void SetSheetName(string name)
        {
            this.WorkSheet.Name = name;
        }
        /// <summary>
        /// 设置指定Worksheet的名字
        /// 作    者： KELL
        /// 日    期： 2011-2-21
        /// </summary>
        /// <param name="currentSheet">指定的Worksheet</param>
        /// <param name="name"></param>
        public void SetSheetName(Excel.Worksheet currentSheet, string name)
        {
            currentSheet.Name = name;
        }
        /// <summary>
        /// 获取当前Sheet的页面设置
        /// </summary>
        public Excel.PageSetup PageSetup
        {
            get
            {
                return this.WorkSheet.PageSetup;
            }
        }
        /// <summary>
        /// 设置当前Sheet为横向打印模式
        /// </summary>
        public void SetPrintOrientationHor()
        {
            this.WorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Sheet为纵向打印模式
        /// </summary>
        public void SetPrintOrientationVer()
        {
            this.WorkSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Sheet从指定的页码开始打印
        /// </summary>
        public void SetPrintFirstPageAt(int firstPage)
        {
            this.WorkSheet.PageSetup.FirstPageNumber = firstPage;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Sheet的打印区域为缩放zoom的范围
        /// </summary>
        /// <param name="zoom">缩放zoom倍，范围：10%~400%</param>
        public void SetPrintRangeZoom(int zoom)
        {
            if (zoom < 10) zoom = 10;
            if (zoom > 400) zoom = 400;
            this.WorkSheet.PageSetup.Zoom = zoom;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Sheet的横向打印区域缩放为pageCount页
        /// </summary>
        /// <param name="pageCount">页数</param>
        public void SetPrintFitToPagesWidth(int pageCount)
        {
            this.WorkSheet.PageSetup.Zoom = false;
            this.WorkSheet.PageSetup.FitToPagesWide = pageCount;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Sheet的纵向打印区域缩放为pageCount页
        /// </summary>
        /// <param name="pageCount">页数</param>
        public void SetPrintFitToPagesHeight(int pageCount)
        {
            this.WorkSheet.PageSetup.Zoom = false;
            this.WorkSheet.PageSetup.FitToPagesTall = pageCount;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Sheet的打印区域缩放为1页(包括横向和纵向)
        /// </summary>
        public void SetPrintFitToOnePage()
        {
            this.WorkSheet.PageSetup.Zoom = false;
            this.WorkSheet.PageSetup.FitToPagesWide = 1;
            this.WorkSheet.PageSetup.FitToPagesTall = 1;
            this.WorkBook.Save();
        }
        /// <summary>
        /// 设置当前Book中所有的Sheet的打印区域缩放为1页(包括横向和纵向)
        /// </summary>
        public void SetPrintFitAllToOnePage()
        {
            for (int i = 1; i <= wb.Worksheets.Count; i++)
            {
                Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[i];
                ws.PageSetup.Zoom = false;
                ws.PageSetup.FitToPagesWide = 1;
                ws.PageSetup.FitToPagesTall = 1;
            }
            wb.Save();
        }
        /// <summary>
        /// 返回或设置纸张的大小
        /// </summary>
        public Excel.XlPaperSize PaperSize
        {
            get
            {
                return this.WorkSheet.PageSetup.PaperSize;
            }
            set
            {
                this.WorkSheet.PageSetup.PaperSize = value;
            }
        }
        /// <summary>
        /// 拷贝对象到剪贴板
        /// </summary>
        /// <param name="obj">要拷贝的对象</param>
        public static void CopyData(object obj)
        {
            System.Windows.Forms.Clipboard.SetDataObject(obj, true);
        }
        /// <summary>
        /// 拷贝html文本到剪贴板
        /// </summary>
        /// <param name="content">html文本</param>
        public static void CopyHtml(string content)
        {
            System.Windows.Forms.Clipboard.SetText(content, System.Windows.Forms.TextDataFormat.Html);
        }
        /// <summary>
        /// 将剪贴板中的对象粘贴到Excel中
        /// </summary>
        public void Paste()
        {
            this.WorkSheet.Paste();
            this.WorkBook.Save();
        }
        /// <summary>
        /// 将剪贴板中的对象带格式粘贴到Excel中
        /// </summary>
        public void PasteSpecial()
        {
            this.WorkSheet.PasteSpecial();
            this.WorkBook.Save();
        }
        /// <summary>
        /// 将剪贴板中的html文本粘贴到Excel中
        /// </summary>
        public void PasteHtml()
        {
            this.WorkSheet.PasteSpecial("HTML", Type.Missing, Type.Missing, Type.Missing, Type.Missing, false);
            this.WorkBook.Save();
        }
        /// <summary>
        /// 打印预览
        /// </summary>
        public void PrintPreview()
        {
            this.WorkSheet.PrintPreview();
        }
        /// <summary>
        /// 获取当前待编辑的Cell位置(Col, Row)
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public System.Drawing.Point GetCurrentCellPosition()
        {
            int iCol = app.ActiveCell.Column;
            int iRow = app.ActiveCell.Row;
            return new System.Drawing.Point(iCol, iRow);
        }
        /// <summary>
        /// 由数字索引获取ExcelSheetIndex枚举(Col, Row)
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="index">从1开始，最大为32，为0时就是当前Sheet</param>
        /// <returns></returns>
        public static ExcelSheetIndex GetExcelSheetIndexByIndex(int index)
        {
            ExcelSheetIndex sheetInd = ExcelSheetIndex.CurrentSheet;
            switch (index)
            {
                case 0:
                    sheetInd = ExcelSheetIndex.CurrentSheet;
                    break;
                case 1:
                    sheetInd = ExcelSheetIndex.Sheet1;
                    break;
                case 2:
                    sheetInd = ExcelSheetIndex.Sheet2;
                    break;
                case 3:
                    sheetInd = ExcelSheetIndex.Sheet3;
                    break;
                case 4:
                    sheetInd = ExcelSheetIndex.Sheet4;
                    break;
                case 5:
                    sheetInd = ExcelSheetIndex.Sheet5;
                    break;
                case 6:
                    sheetInd = ExcelSheetIndex.Sheet6;
                    break;
                case 7:
                    sheetInd = ExcelSheetIndex.Sheet7;
                    break;
                case 8:
                    sheetInd = ExcelSheetIndex.Sheet8;
                    break;
                case 9:
                    sheetInd = ExcelSheetIndex.Sheet9;
                    break;
                case 10:
                    sheetInd = ExcelSheetIndex.Sheet10;
                    break;
                case 11:
                    sheetInd = ExcelSheetIndex.Sheet11;
                    break;
                case 12:
                    sheetInd = ExcelSheetIndex.Sheet12;
                    break;
                case 13:
                    sheetInd = ExcelSheetIndex.Sheet13;
                    break;
                case 14:
                    sheetInd = ExcelSheetIndex.Sheet14;
                    break;
                case 15:
                    sheetInd = ExcelSheetIndex.Sheet15;
                    break;
                case 16:
                    sheetInd = ExcelSheetIndex.Sheet16;
                    break;
                case 17:
                    sheetInd = ExcelSheetIndex.Sheet17;
                    break;
                case 18:
                    sheetInd = ExcelSheetIndex.Sheet18;
                    break;
                case 19:
                    sheetInd = ExcelSheetIndex.Sheet19;
                    break;
                case 20:
                    sheetInd = ExcelSheetIndex.Sheet20;
                    break;
                case 21:
                    sheetInd = ExcelSheetIndex.Sheet21;
                    break;
                case 22:
                    sheetInd = ExcelSheetIndex.Sheet22;
                    break;
                case 23:
                    sheetInd = ExcelSheetIndex.Sheet23;
                    break;
                case 24:
                    sheetInd = ExcelSheetIndex.Sheet24;
                    break;
                case 25:
                    sheetInd = ExcelSheetIndex.Sheet25;
                    break;
                case 26:
                    sheetInd = ExcelSheetIndex.Sheet26;
                    break;
                case 27:
                    sheetInd = ExcelSheetIndex.Sheet27;
                    break;
                case 28:
                    sheetInd = ExcelSheetIndex.Sheet28;
                    break;
                case 29:
                    sheetInd = ExcelSheetIndex.Sheet29;
                    break;
                case 30:
                    sheetInd = ExcelSheetIndex.Sheet30;
                    break;
                case 31:
                    sheetInd = ExcelSheetIndex.Sheet31;
                    break;
                case 32:
                    sheetInd = ExcelSheetIndex.Sheet32;
                    break;
            }
            return sheetInd;
        }
        /// <summary>
        /// 设置当前待编辑的Sheet索引，并激活
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public void SetCurrentSheetAt(int index)
        {
            this.WorkSheet = (Excel.Worksheet)wb.Sheets.get_Item(index);
            this.SheetIndex = GetExcelSheetIndexByIndex(index);
            ws.Activate();
            ws.Cells.get_Range(ws.Cells[1, 1], ws.Cells[1, 1]).Select();
            ws.Cells.get_Range(ws.Cells[1, 1], ws.Cells[1, 1]).Activate();
        }
        /// <summary>
        /// 设置当前待编辑的Cell位置，并激活
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="strCell"></param>
        public void SetCurrentCellAt(string strCell)
        {
            ws.Cells.get_Range((object)strCell, Type.Missing).Select();
            ws.Cells.get_Range((object)strCell, Type.Missing).Activate();
        }
        /// <summary>
        /// 设置当前待编辑的Cell位置，并激活
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public void SetCurrentCellAt(int iRow, int iCol)
        {
            ws.Cells.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Select();
            ws.Cells.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Activate();
        }
        /// <summary>
        /// 激活Sheet表
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public void ActiveSheet()
        {
            if (app.Worksheets.Count < ws.Index)
            {
                Dispose();
                throw new Exception("Sheet表不存在，请检查！");
            }
            try
            {
                ws = (Excel.Worksheet)wb.Worksheets[ws.Index];
                ws.Activate();
            }
            catch
            {
                Dispose();
                throw new Exception("Sheet表激活错误，请检查！");
            }
            int irowcount = 0;
            app.ActiveCell.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Select();
            irowcount = app.ActiveCell.Row;
            if (writeType != ExcelWriteType.ReWrite && irowcount > 65536)
            {
                Dispose();
                throw new Exception("当前Sheet表已达存储上限，不能写入，请检查！");
            }
            int iSheetCount = wb.Sheets.Count;
            for (int i = 1; i <= iSheetCount; i++)
            {
                Excel.Worksheet wsT = (Excel.Worksheet)wb.Sheets[i];
                if (sheetName != null && sheetName != "" && sheetName.Trim() == wsT.Name.Trim() && wsT.Name != ws.Name)
                {
                    Dispose();
                    throw new Exception("当前文件存在同名Sheet表，请检查！");
                }
            }
        }/// <summary>
        /// 方法名称： ReadCell
        /// 内容描述： 读取某单元格的内容，注意输入单元格的合法性
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="strCell"></param>
        /// <returns></returns>
        public string ReadCell(string strCell)
        {
            if (ws == null)
            {
                throw new Exception("当前Sheet已经关闭！");
            }
            // 判断输入项不合法
            Excel.Range rng;
            string strValue = "";
            // Checking
            if (!CellCheck(strCell))
            {
                throw new Exception("读取单元格标识错误，请检查！");
            }
            try
            {
                rng = ws.get_Range(strCell, System.Reflection.Missing.Value);
                if (rng != null)
                {
                    if (rng.Value2 == null)
                        strValue = "";
                    else
                        strValue = rng.Value2.ToString();
                }
            }
            catch
            {
                throw new Exception("读取当前Cell错误，请检查！");
            }
            return strValue;
        }
        /// <summary>
        /// 方法名称： ReadCell
        /// 内容描述： 读取某单元格内容，按照行列参数读取
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <returns></returns>
        public string ReadCell(int iRow, int iCol)
        {
            if (ws == null)
            {
                throw new Exception("当前Sheet已经关闭！");
            }
            Excel.Range rng;
            string strValue = "";
            app.ActiveCell.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Select();
            int irowcount = 0;
            int icolcount = 0;
            irowcount = app.ActiveCell.Row;
            icolcount = app.ActiveCell.Column;
            //if(!CellCheck(iRow,iCol))
            if ((iRow > irowcount) || (iCol > icolcount))
            {
                throw new Exception("读取单元格超出范围，请检查！");
            }
            try
            {
                rng = ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]);
                if (rng != null)
                {
                    if (rng.Value2 == null)
                        strValue = "";
                    else
                        strValue = rng.Value2.ToString();
                }
            }
            catch
            {
                throw new Exception("读取当前Cell错误，请检查！");
            }
            return strValue;
        }
        private static void ReleaseAllRef(Object obj)
        {
            try
            {
                if (obj != null)
                {
                    while (System.Runtime.InteropServices.Marshal.ReleaseComObject(obj) > 1) ;
                }
            }
            finally
            {
                obj = null;
            }
        }
        /// <summary>
        /// 方法名称： WriteCell
        /// 内容描述： 写入数据到某单元格(如果是链接必须先设置好IsLink和LinkFile属性，而LinkSheet、LinkCell、wordLanguage属性则为可选属性，因为它们有默认值Sheet1、A1、SimplifiedChinese)
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="strValue"></param>
        /// <returns></returns>
        public bool WriteCell(int iRow, int iCol, string strValue)
        {
            bool bolRet = false;
            if (ws == null)
            {
                bolRet = false;
                //throw new Exception("当前Sheet已经关闭！");
            }
            if (!CellCheck(iRow, iCol))
            {
                bolRet = false;
                //throw new Exception("写入单元格超出范围，请检查！");
            }
            if (strValue.Length > 255)
            {
                bolRet = false;
                //throw new Exception("单元格值长度超过255个字符，请检查！");
            }
            try
            {
                if (strValue != "")
                {
                    /*
                    Word.ApplicationClass wapp = new Word.ApplicationClass();
                    object Template = Type.Missing;
                    object Visible = false;
                    Word.Documents docs = wapp.Documents;
                    Word.Document doc = docs.Add(ref Template, ref Template, ref Template, ref Visible);
                    object start = 0;
                    object end = 0;
                    Word.Range rng = doc.Range(ref start, ref end);
                    rng = doc.Range(ref start, ref end);
                    rng.InsertBefore(strValue);
                    rng.SpellingChecked = false;
                    rng.Select();
                    switch (this.wordLanguage)
                    {
                        case Language.SimplifiedChinese:
                            rng.LanguageID = Word.WdLanguageID.wdSimplifiedChinese;
                            break;
                        case Language.TraditionalChinese:
                            rng.LanguageID = Word.WdLanguageID.wdTraditionalChinese;
                            break;
                    }
                    rng.LanguageDetected = true;
                    strValue = rng.Text;
                    WordRelease(rng, doc, docs, wapp);
                    */
                    //ws.Cells.AutoFit();//有问题！
                    //ws.Cells.ColumnWidth = strValue.Length;
                    ws.Cells[iRow, iCol] = strValue;
                    Excel.Range rng = (Excel.Range)ws.Cells[iRow, iCol];
                    //rng.Columns.AutoFit();

                    Excel.XlLineStyle ls = Excel.XlLineStyle.xlContinuous;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).LineStyle = ls;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = ls;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = ls;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).LineStyle = ls;
                    
                    Excel.XlColorIndex ci = Excel.XlColorIndex.xlColorIndexAutomatic;//(Excel.XlColorIndex)ws.Cells.Borders.ColorIndex;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeTop).ColorIndex = ci;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeBottom).ColorIndex = ci;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeLeft).ColorIndex = ci;
                    rng.Borders.get_Item(Excel.XlBordersIndex.xlEdgeRight).ColorIndex = ci;
                    //设置具体某个单元格的边框
                    /*Excel.XlLineStyle ls = this.TheBorderStyle.LineStyle;
                    Excel.XlBorderWeight bw = this.TheBorderStyle.BorderWeight;
                    Excel.XlColorIndex ci = this.TheBorderStyle.ColorIndex;
                    rng.BorderAround(ls, bw, ci, null);
                    rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].ColorIndex = ci;
                    rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = ls;
                    rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].Weight = bw;
                    rng.Borders[Excel.XlBordersIndex.xlInsideVertical].ColorIndex = ci;
                    rng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = ls;
                    rng.Borders[Excel.XlBordersIndex.xlInsideVertical].Weight = bw;*/
                    //System.Diagnostics.Process.Start(@"C:\2.bmp");
                    if (this.BackColor != System.Drawing.Color.Empty)
                    {
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Interior.Color = System.Drawing.ColorTranslator.ToOle(this.BackColor);
                    }
                    if (this.ForeColor != System.Drawing.Color.Empty)
                    {
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.Color = System.Drawing.ColorTranslator.ToOle(this.ForeColor);
                    }
                    if (this.Font != null)
                    {
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.FontStyle = this.Font.Style;
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.Bold = this.Font.Bold;
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.Italic = this.Font.Italic;
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.Underline = this.Font.Underline;
                        if (this.Font.Name != "")
                        {
                            ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.Name = this.Font.Name;
                        }
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Font.Size = this.Font.Size;
                    }
                    if (this.IsLink)
                    {//加上链接
                        object showText = strValue;
                        object miss = System.Reflection.Missing.Value;
                        ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]).Hyperlinks.Add(ws.Cells[iRow, iCol], this.LinkFile, this.LinkSheet + "!" + this.LinkCell, miss, showText);
                    }
                }
                wb.Save();
                bolRet = true;
            }
            catch
            {
                bolRet = false;
                //throw new Exception("写入单元格错误，请检查！");
            }
            return bolRet;
        }
        /// <summary>
        /// 方法名称： SetPictureToRange
        /// 内容描述： 写入图片到Range
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="strCell"></param>
        /// <param name="picFilePath"></param>
        public void SetPictureToRange(string strCell, string picFilePath)
        {
            if (File.Exists(picFilePath))
            {
                try
                {
                    Excel.Pictures v_Pictures = (Excel.Pictures)ws.Pictures(Type.Missing);
                    Excel.Picture v_Picture = v_Pictures.Insert(picFilePath, Type.Missing);// + ".jpeg", Type.Missing);
                    // Excel的get_Range方法可以得到Excel的单元格，可以用来设置图片显示的位置
                    Excel.Range v_Range = ws.get_Range((object)strCell, Type.Missing);
                    double v_fFactor = 1;
                    //设置图片大小
                    if (v_Picture.Width * (double)v_Range.Height > v_Picture.Height * (double)v_Range.Width)
                    {
                        v_fFactor = (double)v_Range.Width / (double)v_Picture.Width;
                    }
                    else
                    {
                        v_fFactor = (double)v_Range.Height / (double)v_Picture.Height;
                    }
                    v_Picture.Left = (double)v_Range.Left + ((double)v_Range.Width - (v_Picture.Width * v_fFactor)) / 2 + 1;
                    v_Picture.Top = (double)v_Range.Top + ((double)v_Range.Height - (v_Picture.Height * v_fFactor)) / 2 + 1;
                    v_Picture.Width = v_Picture.Width - 0.5d; // *v_fFactor - 0.5d;
                    v_Picture.Height = v_Picture.Height - 0.5d;// *v_fFactor - 0.5d;
                }
                catch (Exception ex)
                {
                    throw new Exception("Excel添加图片出错！\n" + ex.Message);
                }
            }
        }
        
        /// <summary>
        /// 方法名称： SetPictureToRange
        /// 内容描述： 写入图片到Range
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="iRow"></param>
        /// <param name="iCol"></param>
        /// <param name="picFilePath"></param>
        public void SetPictureToRange(int iRow, int iCol, string picFilePath)
        {
            if (File.Exists(picFilePath))
            {
                try
                {
                    Excel.Pictures v_Pictures = (Excel.Pictures)ws.Pictures(Type.Missing);
                    Excel.Picture v_Picture = v_Pictures.Insert(picFilePath, Type.Missing);// + ".jpeg", Type.Missing);
                    // Excel的get_Range方法可以得到Excel的单元格，可以用来设置图片显示的位置
                    Excel.Range v_Range = ws.get_Range(ws.Cells[iRow, iCol], ws.Cells[iRow, iCol]);
                    double v_fFactor = 1;
                    //设置图片大小
                    if (v_Picture.Width * (double)v_Range.Height > v_Picture.Height * (double)v_Range.Width)
                    {
                        v_fFactor = (double)v_Range.Width / (double)v_Picture.Width;
                    }
                    else
                    {
                        v_fFactor = (double)v_Range.Height / (double)v_Picture.Height;
                    }
                    v_Picture.Left = (double)v_Range.Left + ((double)v_Range.Width - (v_Picture.Width * v_fFactor)) / 2 + 1;
                    v_Picture.Top = (double)v_Range.Top + ((double)v_Range.Height - (v_Picture.Height * v_fFactor)) / 2 + 1;
                    v_Picture.Width = v_Picture.Width - 0.5d; // *v_fFactor - 0.5d;
                    v_Picture.Height = v_Picture.Height - 0.5d;// *v_fFactor - 0.5d;
                }
                catch (Exception ex)
                {
                    throw new Exception("Excel添加图片出错！\n" + ex.Message);
                }
            }
        }

        /// <summary>
        /// 方法名称： SetPictureToRectangle
        /// 内容描述： 写入图片到Rectangle
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <param name="rect"></param>
        /// <param name="picFilePath"></param>
        public void SetPictureToRectangle(System.Drawing.Rectangle rect, string picFilePath)
        {
            if (File.Exists(picFilePath))
            {
                /*Excel.Pictures v_Pictures = (Excel.Pictures)ws.Pictures(Type.Missing);
                Excel.Picture v_Picture = v_Pictures.Insert(PICPATH, Type.Missing);// + ".jpeg", Type.Missing);
                // Excel的get_Range方法可以得到Excel的单元格，可以用来设置图片显示的位置
                Excel.Range v_Range = ws.get_Range((object)p_strRangeName, Type.Missing);
                double v_fFactor = 1;
                //设置图片大小
                if (v_Picture.Width * (double)v_Range.Height > v_Picture.Height * (double)v_Range.Width)
                {
                    v_fFactor = (double)v_Range.Width / (double)v_Picture.Width;
                }
                else
                {
                    v_fFactor = (double)v_Range.Height / (double)v_Picture.Height;
                }
                v_Picture.Left = (double)v_Range.Left + ((double)v_Range.Width - (v_Picture.Width * v_fFactor)) / 2 + 1;
                v_Picture.Top = (double)v_Range.Top + ((double)v_Range.Height - (v_Picture.Height * v_fFactor)) / 2 + 1;
                v_Picture.Width = v_Picture.Width * v_fFactor - 0.5d;
                v_Picture.Height = v_Picture.Height * v_fFactor - 0.5d;*/
                try
                {
                    ws.Shapes.AddPicture(picFilePath, Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, rect.Left, rect.Top, rect.Width, rect.Height);
                }
                catch (Exception ex)
                {
                    throw new Exception("Excel添加图片出错！\n" + ex.Message);
                }
            }
        }
        /// <summary>
        /// 另存为Excel文件
        /// 作    者： KELL
        /// 日    期： 2012-3-10 23:46:00
        /// </summary>
        /// <param name="savePath">保存路径</param>
        /// <param name="format">另存格式，默认为xlExcel7(即Office2003二进制格式)</param>
        public void SaveAs(string savePath, Excel.XlFileFormat format = Excel.XlFileFormat.xlExcel7)
        {
            wb.SaveAs(savePath, format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//,Type.Missing);
        }
        /// <summary>
        /// 存储
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public void Save()
        {
            foreach (Excel.Workbook wb in app.Workbooks)
            {
                wb.Save();
            }
        }

        /// <summary>
        /// 资源释放
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public bool Close()
        {
            try
            {
                Dispose();
                KillExcelProcess();
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 方法名称： Dispose
        /// 内容描述： 如果对Excel的操作没有引发异常的话，用这个方法可以正常结束Excel进程
        /// 否则要用KillExcelProcess()方法来结束Excel进程
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        internal void Dispose()
        {
            //注意：这里用到的所有Excel对象都要执行这个操作，否则结束不了Excel进程
            if (wb != null)
            {
                wb.Close(null, null, null);
                app.Workbooks.Close();
                app.Quit();
            }
            if (rng != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rng);
                rng = null;
            }
            if (tb != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(tb);
                tb = null;
            }
            if (ws != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                ws = null;
            }
            if (wb != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                wb = null;
            }
            if (app != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }

            GC.Collect();
        }

        /// <summary>
        /// 方法名称： Dispose
        /// 内容描述： 如果对Excel的操作没有引发异常的话，用这个方法可以正常结束Excel进程
        /// 否则要用KillExcelProcess()方法来结束Excel进程
        /// 作    者： KELL
        /// 日    期： 2011-7-21
        /// </summary>
        /// <param name="closeApp">是否把Excel.Application也关闭？</param>
        internal void Dispose(bool closeApp)
        {
            //注意：这里用到的所有Excel对象都要执行这个操作，否则结束不了Excel进程
            if (wb != null)
            {
                wb.Close(null, null, null);
                app.Workbooks.Close();
                app.Quit();
            }
            if (ws != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
                ws = null;
            }
            if (wb != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
                wb = null;
            }
            if (closeApp && app != null)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
                app = null;
            }

            GC.Collect();
        }

        /// <summary>
        /// 资源释放
        /// 作    者： KELL
        /// 日    期： 2011-7-21
        /// </summary>
        /// <param name="closeApp">是否把Excel.Application也关闭？</param>
        public bool Close(bool closeApp)
        {
            try
            {
                Dispose(closeApp);
                KillExcelProcess();
                return true;
            }
            catch
            {
                return false;
            }
        }
        /// <summary>
        /// 方法名称： KillExcelProcess
        /// 内容描述： 用Process方法结束Excel进程
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public void KillExcelProcess()
        {
            Process[] myProcesses;
            DateTime startTime;
            myProcesses = Process.GetProcessesByName("Excel");

            //得不到Excel进程ID，暂时只能判断进程启动时间
            foreach (Process myProcess in myProcesses)
            {
                startTime = myProcess.StartTime;

                if (startTime > beforeTime && startTime < afterTime)
                {
                    myProcess.Kill();
                }
            }
        }
        /// <summary>
        /// 关闭所有Excel进程
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        public static void KillAllExcelProcess()
        {
            Process[] myProcesses;
            myProcesses = Process.GetProcessesByName("Excel");
            foreach (Process myProcess in myProcesses)
            {
                if (!myProcess.HasExited)
                {
                    myProcess.Kill();
                }
            }
        }
        /// <summary>
        /// 根据sheet名称获取ExcelSheetIndex对象
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static ExcelSheetIndex GetSheetIndexByName(string filename, string sheetName)
        {
            ExcelSheetIndex esi = ExcelSheetIndex.CurrentSheet;
            MyExcel excel = new MyExcel();
            try
            {
                if (excel.OpenCreate(filename, ExcelSheetIndex.CurrentSheet, false, false))
                {
                    foreach (Excel.Worksheet sheet in excel.WorkBook.Worksheets)
                    {
                        if (sheet.Name.ToLower() == sheetName.ToLower())
                        {
                            esi = (ExcelSheetIndex)Enum.ToObject(typeof(ExcelSheetIndex), sheet.Index);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
            }
            finally
            {
                excel.Close();
            }
            return esi;
        }
    }

    /// <summary>
    /// 公用类库
    /// </summary>
    public class Common
    {
        /// <summary>
        /// 读取Excel中的内容
        /// </summary>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public static string ReadSheet(string excelFile)
        {
            string content = "";
            try
            {
                content = File.ReadAllText(excelFile);
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("导出sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            return content;
        }

        /// <summary>
        /// 将指定的字符串写入到指定的excelFile文件中
        /// </summary>
        /// <param name="dt">要写入的文本</param>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public static bool WriteSheet(string content, string excelFile)
        {
            bool flag = false;
            try
            {
                using (StreamWriter sw = File.CreateText(excelFile))
                {
                    sw.Write(content);
                }
                flag = true;
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("写入sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            return flag;
        }
        /// <summary>
        /// 读取Excel中的指定Sheet的内容，并以DataTable的形式输出(支持Excel11.0)
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static DataTable ReadSheet(string excelFile, string sheetName)
        {
            string strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + excelFile + ";Extended Properties=\"Excel 8.0;HDR=NO;IMEX=1\"";
            string strExcel = "select * from [" + sheetName + "$]";
            DataTable dt = new DataTable();
            OleDbConnection conns = new OleDbConnection(strConn);
            try
            {
                conns.Open();
                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, conns);
                adapter.Fill(dt);
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("导出sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            finally
            {
                conns.Close();
            }
            return dt;
        }

        /// <summary>
        /// 将指定的DataTable写入到指定的excelFile文件中名字为的sheetName的Sheet中
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="dt">只能写入DataTable中的文本</param>
        /// <param name="excelFile"></param>
        /// <param name="sheetIndex">从1开始的Sheet索引，默认为1</param>
        /// <param name="showColumnName">默认为false</param>
        /// <returns></returns>
        public static bool WriteSheet(DataTable dt, string excelFile, string sheetName, int sheetIndex = 1, bool showColumnName = false)
        {
            bool flag = false;
            MyExcel excel = new MyExcel();
            try
            {
                if (excel.OpenCreate(excelFile, MyExcel.GetSheetIndexByName(excelFile, sheetName), false, false))
                {
                    Excel.Worksheet ws = (Excel.Worksheet)excel.WorkBook.Worksheets[sheetIndex];
                    ws.Cells.Clear();
                    int colHeight = 0;
                    int col = dt.Columns.Count;
                    if (showColumnName)
                    {
                        colHeight = 1;
                        for (int i = 0; i < col; i++)
                        {
                            ws.Cells[1, 1 + i] = dt.Columns[i].ColumnName;
                            ws.get_Range(ws.Cells[1, 1 + i], ws.Cells[1, 1 + i]).Font.Bold = true;
                        }
                    }
                    if (dt.Rows.Count > 0)
                    {
                        int row = dt.Rows.Count;
                        for (int i = 0; i < row; i++)
                        {
                            for (int j = 0; j < col; j++)
                            {
                                string str = dt.Rows[i][j].ToString();
                                ws.Cells[i + 1 + colHeight, j + 1] = str;
                            }
                        }
                    }
                    ws.Name = dt.TableName;
                    excel.WorkBook.Save();
                    flag = true;
                }
            }
            catch (Exception err)
            {
                System.Windows.Forms.MessageBox.Show("写入sheet出错！错误原因：" + err.Message, "提示信息",
                    System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
            }
            finally
            {
                excel.Close();
            }
            return flag;
        }
        /// <summary>
        /// 根据指定的Excel文件和Sheet索引，获取Sheet名字
        /// 作    者： KELL
        /// 日    期： 2011-7-19
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="sheetIndex">从1开始的索引值</param>
        /// <returns></returns>
        public static string GetSheetName(string excelFile, int sheetIndex)
        {
            string sheetName = "Sheet1";
            MyExcel excel = new MyExcel();
            try
            {
                if (excel.OpenCreate(excelFile, ExcelSheetIndex.CurrentSheet, false, false))
                {
                    sheetName = ((Excel.Worksheet)excel.WorkBook.Worksheets[sheetIndex]).Name;
                }
            }
            catch { }
            finally
            {
                excel.Close();
            }
            return sheetName;
        }
    }
}
