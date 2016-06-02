using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace SetOnPage
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        KellSetOnePage.SetPrintRange spr = new KellSetOnePage.SetPrintRange();

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim() != "")
            {
                try
                {
                    if (spr.OpenCreate(textBox1.Text.Trim(), checkBox1.Checked))
                    {
                        spr.SetPrintFitToOnePage();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    spr.Close(!checkBox1.Checked);
                }
            }
            else
            {
                MessageBox.Show("先选择Excel文件");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel文件|*.xls";
            if (openFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
            }
            openFileDialog1.Dispose();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                label1.Text = "合并开始...";
                KellSetOnePage.SetPrintRange spr = new KellSetOnePage.SetPrintRange();
                string all = folderBrowserDialog1.SelectedPath + "\\All.xls";
                if (spr.OpenCreate(all, false))
                {
                    string[] files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.xls");
                    foreach (string f in files)
                    {
                        FileInfo fi = new FileInfo(f);
                        if (fi.Name != "All.xls")
                        {
                            label1.Text = "合并[" + fi.Name + "]开始...";
                            spr.AddAnExternalSheet(f);
                            label1.Text = "合并[" + fi.Name + "]完成.";
                        }
                    }
                }
                spr.Close();
                label1.Text = "合并结束.";
            }
            folderBrowserDialog1.Dispose();
        }

        private void label1_TextChanged(object sender, EventArgs e)
        {
            label1.Refresh();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (checkBox1.Checked)
                spr.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Trim() != "")
            {
                try
                {
                    KellSetOnePage.SetPrintRange spr = new KellSetOnePage.SetPrintRange();
                    if (spr.OpenCreate(openFileDialog1.FileName, checkBox1.Checked))
                    {
                        saveFileDialog1.Filter = "Excel文件|*.xls";
                        if (saveFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            spr.SaveAs(saveFileDialog1.FileName);
                            MessageBox.Show("保存完毕！");
                        }
                        saveFileDialog1.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
                finally
                {
                    spr.Close(!checkBox1.Checked);
                }
            }
            else
            {
                MessageBox.Show("先选择Excel文件");
            }
        }
    }
}
