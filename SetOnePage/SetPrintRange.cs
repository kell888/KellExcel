using System;
using System.IO;
using System.Diagnostics;
using System.Data;
using System.Data.OleDb;
using System.Text;

namespace KellSetOnePage
{
    public class SetPrintRange
    {
        public void Set(string excelPath, bool display)
        {
            try
            {
                if (OpenCreate(excelPath, display))
                {
                    SetPrintFitToOnePage();
                }
                else
                {
                    throw new Exception("设置为1页打印时出错！");
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                Close();
            }
        }
        private string filePath = "";
        public string FilePath
        {
            get { return filePath; }
        }
        DateTime beforeTime, afterTime;
        Excel.ApplicationClass app;
        Excel.Workbook wb;

        public Excel.Workbook Workbook
        {
            get { return wb; }
        }
        Excel.Worksheet ws;

        public Excel.Worksheet Worksheet
        {
            get { return ws; }
        }
        public const int MaxSheetCount = 32;
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
            catch (Exception e)
            {
                KillExcelProcess();
                throw new Exception("创建Excel应用对象出现错误，请检查你的机器！\n" + e.Message);
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
        /// 方法名称： Open
        /// 内容描述： 无
        /// 实现流程： 打开/连接一个excel数据文档，只能指定一个存在的文件
        /// 作    者： KELL
        /// 日    期： 2007-5-18 18:10:00
        /// </summary>
        /// <returns></returns>
        public bool Open(string filepath, bool display)
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
            try
            {
                if (IsOpen == false)
                {
                    if (wb == null)
                        wb = app.Workbooks.Open(filepath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws = (Excel.Worksheet)wb.Worksheets[1];
                    ws.Activate();
                }
                bolRetValue = true;
            }
            catch (Exception ex)
            {
                KillExcelProcess();
                throw new Exception("打开或连接Excel文档错误，请检查！\n" + ex.Message);
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
        public bool OpenCreate(string filepath, bool display)
        {
            bool bolRet = false;
            try
            {
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
                bolRet = Open(filepath, display);
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
        /// 另存为Excel二进制文件
        /// 作    者： KELL
        /// 日    期： 2012-3-10 23:41:00
        /// </summary>
        /// <param name="savePath">保存路径</param>
        /// <param name="format">另存格式，默认为xlExcel7(即Office2003二进制格式)</param>
        public void SaveAs(string savePath, Excel.XlFileFormat format = Excel.XlFileFormat.xlExcel7)
        {
            wb.SaveAs(savePath, format, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);//,Type.Missing);
        }
        /// <summary>
        /// 设置当前Sheet的横向打印区域缩放为pageCount页
        /// </summary>
        /// <param name="pageCount">页数</param>
        public void SetPrintFitToPagesWidth(int pageCount)
        {
            ws.PageSetup.Zoom = false;
            ws.PageSetup.FitToPagesWide = pageCount;
            wb.Save();
        }
        /// <summary>
        /// 设置当前Sheet的纵向打印区域缩放为pageCount页
        /// </summary>
        /// <param name="pageCount">页数</param>
        public void SetPrintFitToPagesHeight(int pageCount)
        {
            ws.PageSetup.Zoom = false;
            ws.PageSetup.FitToPagesTall = pageCount;
            wb.Save();
        }
        /// <summary>
        /// 设置当前Sheet的打印区域缩放为1页(包括横向和纵向)
        /// </summary>
        public void SetPrintFitToOnePage()
        {
            ws.PageSetup.Zoom = false;
            ws.PageSetup.FitToPagesWide = 1;
            ws.PageSetup.FitToPagesTall = 1;
            wb.Save();
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
                wb.Save();
            }
        }
        public void SetAllRowHeight(int rowBegin, int rowEnd, int colBegin, int colEnd, int height)
        {
            if (height > 0)
            {
                ws.get_Range(ws.Cells[rowBegin, colBegin], ws.Cells[rowEnd, colEnd]).RowHeight = height;
                wb.Save();
            }
        }
        public void GotoNextSheet()
        {
            int current = ws.Index;
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
        public void GotoPrevSheet()
        {
            int current = ws.Index;
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
        /// 在Web下面行不通
        /// </summary>
        public void Copy()
        {
            ws.Select();
            System.Windows.Forms.SendKeys.Send("^a");
            System.Windows.Forms.SendKeys.Send("^c");
        }
        /// <summary>
        /// 在Web下面行不通
        /// </summary>
        public void Paste()
        {
            ws.Select();
            System.Windows.Forms.SendKeys.Send("^v");
        }
        /// <summary>
        /// 在当前的Worksheet之后添加1个外部的Worksheet，有问题！
        /// </summary>
        /// <param name="externalExcelFile">外部的Excel文件</param>
        public void AddAnExternalSheet(string externalExcelFile)
        {
            SetPrintRange externalExcel = new SetPrintRange();
            try
            {
                //if (externalExcel.OpenCreate(externalExcelFile, false))
                //{
                //externalExcel.Worksheet.Copy(Type.Missing, ws);
                //DataTable dt = ReadSheet(externalExcelFile, GetSheetName(externalExcelFile, 1));
                //string content = ReadSheetEx(externalExcelFile);
                GotoNextSheet(); this.wb.MergeWorkbook(externalExcelFile);
                //WriteSheet(content);//会把原来的其他sheet内容重写掉！！！
                //WriteSheet(dt, ws.Index);
                //externalExcel.Worksheet.Copy();
                //this.ws.Activate();
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
        /// </summary>
        /// <param name="externalExcelFile">外部的Excel文件</param>
        public void InsertAnExternalSheet(string externalExcelFile)
        {
            SetPrintRange externalExcel = new SetPrintRange();
            try
            {
                //if (externalExcel.OpenCreate(externalExcelFile, false))
                //{
                //externalExcel.Worksheet.Copy(ws, Type.Missing);
                //DataTable dt = ReadSheet(externalExcelFile, GetSheetName(externalExcelFile, 1));
                //string content = ReadSheetEx(externalExcelFile);
                GotoPrevSheet(); this.wb.MergeWorkbook(externalExcelFile);
                //WriteSheet(content);//会把原来的其他sheet内容重写掉！！！
                //WriteSheet(dt, ws.Index);
                //externalExcel.Worksheet.Copy();
                //this.ws.Activate();
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
                //using (FileStream fs = new FileStream(this.filePath, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                //{
                //    byte[] data = Encoding.Default.GetBytes(content);
                //    fs.Write(data, 0, data.Length);
                //}
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
        /// </summary>
        /// <param name="dt">只能写入DataTable中的文本</param>
        /// <param name="sheetIndex">从1开始的Sheet索引，默认为1</param>
        /// <param name="showColumnName">默认为false</param>
        /// <returns></returns>
        public bool WriteSheet(DataTable dt, int sheetIndex = 1, bool showColumnName = false)
        {
            bool flag = false;
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
            return flag;
        }

        /// <summary>
        /// 读取Excel中的指定Sheet的内容，并以DataTable的形式输出(支持Excel11.0)
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
        /// </summary>
        /// <param name="dt">只能写入DataTable中的文本</param>
        /// <param name="excelFile"></param>
        /// <param name="sheetIndex">从1开始的Sheet索引，默认为1</param>
        /// <param name="showColumnName">默认为false</param>
        /// <returns></returns>
        public static bool WriteSheet(DataTable dt, string excelFile, int sheetIndex = 1, bool showColumnName = false)
        {
            bool flag = false;
            SetPrintRange excel = new SetPrintRange();
            try
            {
                if (excel.OpenCreate(excelFile, false))
                {
                    Excel.Worksheet ws = (Excel.Worksheet)excel.Workbook.Worksheets[sheetIndex];
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
                    excel.Workbook.Save();
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
        /// 读取Excel中的内容
        /// </summary>
        /// <param name="excelFile"></param>
        /// <returns></returns>
        public static string ReadSheetEx(string excelFile)
        {
            string content = "";
            try
            {
                content = File.ReadAllText(excelFile);
                //using (FileStream fs = new FileStream(excelFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                //{
                //    byte[] data = new byte[1024];
                //    int len = 0;
                //    StringBuilder sb = new StringBuilder();
                //    while ((len = fs.Read(data, 0, data.Length)) > 0)
                //    {
                //        sb.Append(Encoding.Default.GetString(data, 0, len));
                //    }
                //    content = sb.ToString();
                //}
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
        /// 根据指定的Excel文件和Sheet索引，获取Sheet名字
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="sheetIndex">从1开始的索引值</param>
        /// <returns></returns>
        public static string GetSheetName(string excelFile, int sheetIndex)
        {
            string sheetName = "Sheet1";
            SetPrintRange spr = new SetPrintRange();
            try
            {
                if (spr.OpenCreate(excelFile, false))
                {
                    sheetName = ((Excel.Worksheet)spr.Workbook.Worksheets[sheetIndex]).Name;
                }
            }
            catch { }
            finally
            {
                spr.Close();
            }
            return sheetName;
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
        /// 日    期： 2007-5-18 18:10:00
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
    }
}
