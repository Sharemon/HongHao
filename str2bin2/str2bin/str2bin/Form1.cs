using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace str2bin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private const int size = 16777216;
        private byte[] buf = new byte[size];
        private string fileNameBin;
        private string fileNameTxt;
        private FileStream fs;
        private StreamWriter sw;
        private bool newFlag = false;
        Label[] series = new Label[100];
        Label[] address = new Label[100];
        TextBox[] content = new TextBox[100];
        CheckBox[] enter = new CheckBox[100];
        CheckBox[] newLine = new CheckBox[100];
        private void Form1_Load(object sender, EventArgs e)
        {
            int ini = Convert.ToInt32(baseAddr.Text,16);
            int add = Convert.ToInt32(upAddr.Text,16); 
            for (int i = 0; i < 100; i++)
            {
                //序号列
                series[i] = new Label();
                series[i].Parent = this;
                series[i].Width = 48;
                series[i].Left = 13;
                series[i].Top = 113 + i * 27;
                series[i].Text = i.ToString("00");
                series[i].Visible = true;

                //地址
                address[i] = new Label();
                address[i].Parent = this;
                address[i].Width = 48;
                address[i].Left = 69;
                address[i].Top = 113 + i * 27;
                address[i].Text = string.Format("{0:X6}", (ini + add * i));
                address[i].Visible = true; 

                //回车
                enter[i] = new CheckBox();
                enter[i].Parent = this;
                enter[i].Width = 48;
                enter[i].Left = 397;
                enter[i].Top = 113 + i * 27;
                enter[i].Text = "0D";
                enter[i].Visible = true;

                //换行
                newLine[i] = new CheckBox();
                newLine[i].Parent = this;
                newLine[i].Width = 48;
                newLine[i].Left = 451;
                newLine[i].Top = 113 + i * 27;
                newLine[i].Text = "0A";
                newLine[i].Visible = true;
            }
            for (int i = 0; i < 100; i++)
            {
                //地址
                content[i] = new TextBox();
                content[i].Parent = this;
                content[i].Width = 230;
                content[i].Left = 139;
                content[i].Top = 113 + i * 27;
                content[i].Text = "";
                content[i].Visible = true;
                this.content[i].TextChanged += new System.EventHandler(this.content_TextChanged);
            }
            this.WindowState = FormWindowState.Normal;
        }

        private void baseAddr_TextChanged(object sender, EventArgs e)
        {
            if ((baseAddr.Text != ""))
            {
                char getchar = baseAddr.Text[baseAddr.Text.Length - 1];
                if ((getchar <= 57 && getchar >= 48) || (getchar <= 70 && getchar >= 65) || (getchar <= 102 && getchar >= 97))
                {
                    if (baseAddr.Text.Length > 6)
                    {
                        MessageBox.Show("地址大小超过限制！", "提示");
                        string str = baseAddr.Text.Remove(baseAddr.Text.Length - 1);
                        baseAddr.Text = str;
                    }
                    if (upAddr.Text != "")
                    {
                        int ini = Convert.ToInt32(baseAddr.Text, 16);
                        int add = Convert.ToInt32(upAddr.Text, 16);
                        int addr;
                        for (int i = 0; i < 100; i++)
                        {
                            addr = (ini + add * i);
                            if (addr < size)
                            {
                                address[i].Text = string.Format("{0:X6}", addr);
                            }
                            else
                            {
                                address[i].Text = "";
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请输入一个正确的十六进制地址！", "提示");
                    string str = baseAddr.Text.Remove(baseAddr.Text.Length - 1);
                    baseAddr.Text = str;
                }
            }
        }

        private void upAddr_TextChanged(object sender, EventArgs e)
        {
            if ((upAddr.Text != ""))
            {
                char getchar = upAddr.Text[upAddr.Text.Length - 1];
                if ((getchar <= 57 && getchar >= 48) || (getchar <= 70 && getchar >= 65) || (getchar <= 102 && getchar >= 97))
                {
                    if (upAddr.Text.Length > 6)
                    {
                        MessageBox.Show("地址大小超过限制！", "提示");
                        string str = upAddr.Text.Remove(upAddr.Text.Length - 1);
                        upAddr.Text = str;
                    }
                    if (baseAddr.Text != "")
                    {
                        int ini = Convert.ToInt32(baseAddr.Text, 16);
                        int add = Convert.ToInt32(upAddr.Text, 16);
                        int addr;
                        for (int i = 0; i < 100; i++)
                        {
                            addr = (ini + add * i);
                            if (addr < size)
                            {
                                address[i].Text = string.Format("{0:X6}", addr);
                            }
                            else
                            {
                                address[i].Text = "";
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("请输入一个正确的十六进制地址！", "提示");
                    string str = upAddr.Text.Remove(upAddr.Text.Length - 1);
                    upAddr.Text = str;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            string dic = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\BinData";
            if (!Directory.Exists(dic))
            {
                Directory.CreateDirectory(dic);
            }
            sfd.InitialDirectory = dic;
            sfd.Filter = "bin文件(*.bin)|*.bin";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                fileNameBin = sfd.FileName;
                byte by = Convert.ToByte(fill.Text, 16);
                for (int i = 0; i < size; i++)
                {
                    buf[i] = by;
                }
                fs = new FileStream(fileNameBin, FileMode.Create);
                fs.Write(buf, 0, buf.Length);
                fs.Close();

                this.Text = "str2bin" + "   " + fileNameBin;

                fileNameTxt = fileNameBin.Remove(fileNameBin.Length - 4) + ".txt";
                sw = new StreamWriter(fileNameTxt, false);
                sw.Write("空白填充：\t" + fill.Text + "\r\n起始地址：\t" + baseAddr.Text + "\r\n自增地址：\t" + upAddr.Text + "\r\n");
                sw.Close();
                for (int i = 0; i < 100; i++)
                {
                    if (!string.IsNullOrEmpty(address[i].Text))
                    {
                        if (true/*!string.IsNullOrEmpty(content[i].Text)*/)
                        {
                            string enLine = "";
                            if (enter[i].Checked)
                            {
                                enLine = enLine + "+0D";
                            }
                            if (newLine[i].Checked)
                            {
                                enLine = enLine + "+0A";
                            }
                            sw = new StreamWriter(fileNameTxt, true);
                            string str = address[i].Text + "\t:" + content[i].Text + enLine;
                            sw.WriteLine(str);
                            sw.Close();
                        }
                    }
                }
                newFlag = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            string dic = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\BinData";
            if (!Directory.Exists(dic))
            {
                Directory.CreateDirectory(dic);
            }
            ofd.InitialDirectory = dic;
            ofd.Filter = "bin文件(*.bin)|*.bin";
            ofd.RestoreDirectory = true;
            ofd.FilterIndex = 1;
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                fileNameBin = ofd.FileName;
                //fs = new FileStream(fileNameBin, FileMode.Open);
                //fs.Read(buf, 0, buf.Length);
                //fs.Close();

                this.Text = "str2bin" + "   " + fileNameBin;

                fileNameTxt = fileNameBin.Remove(fileNameBin.Length - 4) + ".txt";
                if (System.IO.File.Exists(fileNameTxt))
                {
                    StreamReader sr = new StreamReader(fileNameTxt);
                    ReadTxt(sr);
                    sr.Close();
                    //sw = new StreamWriter(fileNameTxt, true);
                }
                else
                {
                    //sw = new StreamWriter(fileNameTxt, false);
                    MessageBox.Show("无相应txt文件！", "提示");
                }
                //sw.Close();
                newFlag = true;
            }
        }

        private void ReadTxt(StreamReader sr)
        {
            string str = sr.ReadLine();
            if (str == null)
            {
                MessageBox.Show("此文本为空！", "提示");
            }
            else
            {
                str = str.Substring(str.LastIndexOf('\t') + 1);
                int data = Convert.ToInt32(str, 16);
                fill.Text = string.Format("{0:X2}", data);
                str = sr.ReadLine();
                str = str.Substring(str.LastIndexOf('\t') + 1);
                data = Convert.ToInt32(str, 16);
                baseAddr.Text = data.ToString("X");
                str = sr.ReadLine();
                str = str.Substring(str.LastIndexOf('\t') + 1);
                data = Convert.ToInt32(str, 16);
                upAddr.Text = data.ToString("X");

                for (int i = 0; i < 100; i++)
                {
                    str = sr.ReadLine();
                    if (str == null)
                    {
                        break;
                    }
                    str = str.Substring(str.LastIndexOf(':') + 1);
                    if ((str.Length >= 3) && (str.Substring(str.Length - 3) == "+0A"))
                    {
                        newLine[i].Checked = true;
                        str = str.Substring(0, str.Length - 3);
                    }
                    else
                    {
                        newLine[i].Checked = false;
                    }
                    if ((str.Length >= 3) && str.Substring(str.Length - 3) == "+0D")
                    {
                        enter[i].Checked = true;
                        str = str.Substring(0, str.Length - 3);
                    }
                    else
                    {
                        enter[i].Checked = false;
                    }
                    content[i].Text = str;
                    content[i].SelectionStart = content[i].Text.Length;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (newFlag)
            {
                sw = new StreamWriter(fileNameTxt, false);
                sw.Write("空白填充：\t" + fill.Text + "\r\n起始地址：\t" + baseAddr.Text + "\r\n自增地址：\t" + upAddr.Text + "\r\n");
                sw.Close();
                byte by = Convert.ToByte(fill.Text, 16);
                for (int i = 0; i < size; i++)
                {
                    buf[i] = by;
                }
                for (int i = 0; i < 100; i++)
                {
                    if (!string.IsNullOrEmpty(address[i].Text))
                    {
                        if (true/*!string.IsNullOrEmpty(content[i].Text)*/)
                        {
                            int addr = Convert.ToInt32(address[i].Text, 16);
                            Byte[] info = new UTF8Encoding(true).GetBytes(content[i].Text);
                            int j;
                            for (j = 0; j < info.Length; j++)
                            {
                                buf[j + addr] = info[j];
                            }
                            //buf[addr + i] = 0;
                            string enLine = "";
                            if (enter[i].Checked)
                            {
                                buf[j + addr] = 13;
                                enLine = enLine + "+0D";
                                j++;
                            }
                            if (newLine[i].Checked)
                            {
                                buf[j + addr] = 10;
                                enLine = enLine + "+0A";
                                j++;
                            }
                            sw = new StreamWriter(fileNameTxt, true);
                            string str = address[i].Text + "\t:" + content[i].Text + enLine;
                            sw.WriteLine(str);
                            sw.Close();
                        }
                    }
                }
                fs = new FileStream(fileNameBin, FileMode.Create);
                fs.Write(buf, 0, buf.Length);
                fs.Close();
            }
            else
            {
                MessageBox.Show("请先新建/打开一个文件！", "提示");
            }
            MessageBox.Show("写入完成！", "提示");
        }

        private void allEnter_CheckedChanged(object sender, EventArgs e)
        {
            if (allEnter.Checked)
            {
                for (int i = 0; i < 100; i++)
                {
                    enter[i].Checked = true;
                }
            }
            else
            {
                for (int i = 0; i < 100; i++)
                {
                    enter[i].Checked = false;
                }
            }
        }

        private void allnewLine_CheckedChanged(object sender, EventArgs e)
        {
            if (allnewLine.Checked)
            {
                for (int i = 0; i < 100; i++)
                {
                    newLine[i].Checked = true;
                }
            }
            else
            {
                for (int i = 0; i < 100; i++)
                {
                    newLine[i].Checked = false;
                }
            }
        }

        private void content_TextChanged(object sender, EventArgs e)
        {
            /*
            string str = (string)sender;
            if (str.Length > Convert.ToInt32(upAddr.Text, 16))
            {
                MessageBox.Show("输入字符过多！", "提示",MessageBoxButtons.OKCancel);
            }
             * */
        }
    }
}
