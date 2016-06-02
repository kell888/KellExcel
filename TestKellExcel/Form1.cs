using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace TestKellExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        KellExcel.MyExcel excel;
        private void button1_Click(object sender, EventArgs e)
        {
            excel = new KellExcel.MyExcel();
            openFileDialog1.Filter = "Excel文件|*.xls";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox9.Text = openFileDialog1.FileName;
            }
            openFileDialog1.Dispose();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (excel != null)
            {
                excel.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (excel != null)
            {
                saveFileDialog1.Filter = "Excel File | *.xls";
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    excel.Save();
                    excel.SaveAs(saveFileDialog1.FileName);
                }
                saveFileDialog1.Dispose();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (excel != null)
            {
                excel.IsLink = true;
                excel.LinkFile = @"C:\2.bmp";
                int iRow = int.Parse(textBox2.Text.Trim());
                int iCol = int.Parse(textBox3.Text.Trim());
                if (excel.WriteCell(iRow, iCol, textBox1.Text))
                    MessageBox.Show(textBox1.Text + "写入到" + KellExcel.MyExcel.GetCellNameByIndexs(iRow, iCol) + "成功！");
            }
        }
        KellExcel.CellIndexs ci;
        string cn = "";
        private void button4_Click(object sender, EventArgs e)
        {
            ci = KellExcel.MyExcel.GetCellIndexsByName(cn);
            MessageBox.Show(ci.ToString());
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //int iRow = int.Parse(textBox2.Text.Trim());
            //int iCol = int.Parse(textBox3.Text.Trim());
            //cn = KellExcel.MyExcel.GetCellNameByIndexs(iRow, iCol);
            cn = KellExcel.MyExcel.GetCellNameByIndexs(ci);
            MessageBox.Show(cn);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (excel != null)
            {
                if (radioButton1.Checked)
                {
                    ci.Col++;
                    textBox3.Text = ci.Col.ToString();
                }
                else
                {
                    ci.Row++;
                    textBox2.Text = ci.Row.ToString();
                }
                textBox4.Text = KellExcel.MyExcel.GetCellNameByIndexs(ci);

                excel.IsLink = true;
                excel.LinkFile = @"C:\2.bmp";
                if (excel.WriteCell(ci.Row, ci.Col, textBox1.Text))
                    MessageBox.Show(textBox1.Text + "写入到" + KellExcel.MyExcel.GetCellNameByIndexs(ci) + "成功！");
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            ci.Row = int.Parse(textBox2.Text.Trim());
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            ci.Col = int.Parse(textBox3.Text.Trim());
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            cn = textBox4.Text.Trim();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ci.Row = int.Parse(textBox2.Text.Trim());
            ci.Col = int.Parse(textBox3.Text.Trim());
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int current = excel.GetCurrentSheetIndex();
            //MessageBox.Show(current.ToString());
            if (current >= KellExcel.MyExcel.MaxSheetCount)
            {
                MessageBox.Show("Sheet数量已满！");
            }
            excel.GotoNextSheet();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            excel.AddSheet();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            bool has = excel.ExistsSheetIndex(int.Parse(textBox5.Text.Trim()));
            MessageBox.Show(has.ToString());
        }

        private void button10_Click(object sender, EventArgs e)
        {
            excel.SetPrintFitToPagesWidth(int.Parse(textBox6.Text.Trim()));
        }

        private void button11_Click(object sender, EventArgs e)
        {
            excel.SetPrintFitToPagesHeight(int.Parse(textBox7.Text.Trim()));
        }

        private void button12_Click(object sender, EventArgs e)
        {
            excel.SetPrintFitToOnePage();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            excel.PrintPreview();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            excel.SetPrintRangeZoom(int.Parse(textBox8.Text.Trim()));
        }

        private void button15_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Excel文件|*.xls";
            if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                ExcelWithoutCOM.ExcelWriter excel = new ExcelWithoutCOM.ExcelWriter(saveFileDialog1.FileName);
                excel.BeginWrite();
                if (checkBox3.Checked)
                    excel.WriteNumber(short.Parse(textBox2.Text.Trim()), short.Parse(textBox3.Text.Trim()), double.Parse(textBox1.Text.Trim()));
                else
                excel.WriteString(short.Parse(textBox2.Text.Trim()), short.Parse(textBox3.Text.Trim()), textBox1.Text);
                excel.EndWrite();
            }
            saveFileDialog1.Dispose();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (File.Exists(textBox9.Text.Trim()))
            {
                KellExcel.MyExcel ex = new KellExcel.MyExcel();
                try
                {
                    if (ex.OpenCreate(textBox9.Text.Trim(), KellExcel.ExcelSheetIndex.CurrentSheet, false, false))
                    {
                        ex.SetPrintFitToOnePage();
                    }
                }
                catch (Exception ee)
                {
                    MessageBox.Show(ee.Message);
                }
                finally
                {
                    ex.Close();
                }
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            excel.OpenCreate(openFileDialog1.FileName, KellExcel.ExcelSheetIndex.CurrentSheet, checkBox1.Checked, checkBox2.Checked);
        }
    }
}