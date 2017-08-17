using Aspose.Email.Mail;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MailCoverter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MailMessage message = MailMessage.Load(textBox1.Text, MessageFormat.Msg);
            message.Save(textBox2.Text, MailMessageSaveType.EmlFormat);  
        }


       

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.Title = "打开文件";
            oFD.ShowHelp = true;
            oFD.Filter = "所有文件|*.*";//过滤格式
            oFD.FilterIndex = 1;                                    //格式索引
            oFD.RestoreDirectory = false;
            oFD.InitialDirectory = "c:\\";                          //默认路径
            oFD.Multiselect = false;                                 //是否多选
            if (oFD.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = oFD.FileName;      
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
           // saveFileDialog1.InitialDirectory = Path.GetDirectoryName(strPartPath);
            //设置文件类型
            saveFileDialog1.Filter = "所有文件|*.*";
            //saveFileDialog1.FilterIndex = 1;//设置文件类型显示
            //saveFileDialog1.FileName = "自己取个";//设置默认文件名
            saveFileDialog1.RestoreDirectory = true;//保存对话框是否记忆上次打开的目录
            saveFileDialog1.CheckPathExists = true;//检查目录
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = saveFileDialog1.FileName;//文件路径
            }       
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MailMessage message = MailMessage.Load(textBox1.Text, MessageFormat.Eml);
            message.Save(textBox2.Text, MailMessageSaveType.OutlookMessageFormatUnicode);  
        }

       
    }
}
