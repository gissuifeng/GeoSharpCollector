using DevExpress.XtraEditors;
using GeoSharp2018.UtilClass;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GeoSharp2018.SystemForms
{
    public partial class RegForm : DevExpress.XtraEditors.XtraForm
    {
        private RegClass reg = new RegClass();

        public RegForm()
        {
            InitializeComponent();
        }

        private void RegForm_Load(object sender, EventArgs e)
        {
            reg = new RegClass();

            string mnum = reg.getMNum();

            textEdit_MachineNum.EditValue = mnum;

            regPic.Visible = false;

            try
            {
                string path = string.Format(@"{0}\{1}", Application.StartupPath, "sys.xml");
                StreamReader sr = new StreamReader(path, Encoding.Default);
                String line;
                while ((line = sr.ReadLine()) != null)
                {
                    textEdit_RegNum.EditValue = line;
                }

                if (textEdit_RegNum.EditValue.ToString() != "")
                {
                    if (textEdit_RegNum.EditValue.ToString() == reg.getRNum(reg.getMNum()))
                    {
                        regPic.Visible = true;
                    }
                    else
                    {
                        regPic.Visible = false;
                    }
                }

                sr.Close();

            }
            catch
            {

            }
        }

        private void btn_reg_Click(object sender, EventArgs e)
        {
            string s = textEdit_MachineNum.EditValue.ToString();

            if (s == "")
            {
                XtraMessageBox.Show("注册码不能为空！");
                return;
            }

            string path = string.Format(@"{0}\{1}", Application.StartupPath, "sys.xml");

            FileStream fs = new FileStream(path, FileMode.Create);
            StreamWriter sw = new StreamWriter(fs);

            sw.Write(textEdit_RegNum.EditValue.ToString());

            sw.Flush();

            sw.Close();
            fs.Close();

            if (textEdit_RegNum.EditValue.ToString() == reg.getRNum(reg.getMNum()))
            {
                regPic.Visible = true;
                XtraMessageBox.Show("注册成功");
            }
            else
            {
                regPic.Visible = false;
                XtraMessageBox.Show("注册码无效， 注册失败！");
            }
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
