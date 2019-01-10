using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace GeoSharp2018.ToolForms
{
    public partial class PopflowForm : DevExpress.XtraEditors.XtraForm
    {

        //定义delegate以便Invoke时使用  
        private delegate void SetProgressBarValue(int value);

        /// <summary>
        /// 系统表格控件
        /// </summary>
        private SpreadsheetControl spreadsheetControl;
       
        /// <summary>
        /// 主窗口对象
        /// </summary>
        private MainForm mainform;

        private string filePath;

        private Thread thread;
    

        public PopflowForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }


        private void PopflowForm_Load(object sender, EventArgs e)
        {
            comboBox_datatype.Properties.Items.Add("5");
            comboBox_datatype.Properties.Items.Add("10");
            comboBox_datatype.Properties.Items.Add("30");
            comboBox_datatype.Properties.Items.Add("60");
            comboBox_datatype.SelectedIndex = 0;

            dateEdit1.EditValue = DateTime.Now.AddDays(-1);
        }

        private void btn_savePath_Click(object sender, EventArgs e)
        {
            UtilClass.FolderBrowserDialog dlg = new UtilClass.FolderBrowserDialog();
            
            DialogResult dr = dlg.ShowDialog(this);

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                filePath = dlg.DirectoryPath;
                textEdit_file.EditValue = filePath;
            }
        }

        private void textEdit_file_EditValueChanged(object sender, EventArgs e)
        {
            filePath = textEdit_file.EditValue.ToString();
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {

            time_interval = Convert.ToInt32(comboBox_datatype.Properties.Items[comboBox_datatype.SelectedIndex].ToString());

            if (time_interval == 5)
            {
                progressBarControl1.Properties.Maximum = 288;
            }
            else if (time_interval == 10)
            {
                progressBarControl1.Properties.Maximum = 144;
            }
            else if (time_interval == 30)
            {
                progressBarControl1.Properties.Maximum = 48;
            }
            else if (time_interval == 60)
            {
                progressBarControl1.Properties.Maximum = 24;
            }

            thread = new Thread(new ThreadStart(fun1));
            thread.Start();
     
        }

        int sum = 0;
        int time_interval = 5;
        private void fun1()
        {

            DateTime dt = Convert.ToDateTime(dateEdit1.EditValue.ToString());

            string path = filePath;
            string date = Convert.ToDateTime(dateEdit1.EditValue.ToString()).ToString("yyyy-MM-dd");
            time_interval = Convert.ToInt32(comboBox_datatype.Properties.Items[comboBox_datatype.SelectedIndex].ToString());

            string h, m;

            for (int i = 0; i < 24; i++)
            {
                for (int j = 0; j < 60; j = j + time_interval)
                {
                    if (i < 10)
                    {
                        h = "0" + i;
                    }
                    else
                    {
                        h = i.ToString();
                    }
                    if (j < 10)
                    {
                        m = "0" + j;
                    }
                    else
                    {
                        m = j.ToString();
                    }


                    fun(path, date, h, m, "00");
                    sum++;

                    RunWithInoke(sum);
                }
            }

            taskExecuted = true;
            XtraMessageBox.Show(string.Format("已经成功采集完成{0}个文件！", time_interval));
        }

        private void RunWithInoke(int i)
        {

            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);

        }

        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;
        }

        /// <summary>
        /// 数据采集
        /// </summary>
        /// <param name="path">文件存储路径</param>
        /// <param name="date">采集时间</param>
        /// <param name="h">时</param>
        /// <param name="m">分</param>
        /// <param name="s">秒</param>
        public void fun(string path, string date, string h, string m, string s)
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            //庐山 1179
            //北京欢乐谷 5381
            //丽江古城 894
            //首都机场786
            //厦门大学6423
            string str = string.Format(@"https://heat.qq.com/api/getHeatDataByTime.php?region_id=5381&datetime={0}+{1}:{2}:{3}", date, h, m, s);
            Stream stream = client.OpenRead(str);

            string str_json = new StreamReader(stream).ReadToEnd();

            str_json = str_json.Replace("{", "").Replace("}", "");

            string[] strs = str_json.Split('"');

            CoordVal coordVal = null;

            Workbook workbook = new Workbook();
            workbook.CreateNewDocument();
            workbook.SaveDocument(string.Format(@"{0}\r{1}_{2}{3}.xls", path, date.Replace("-", ""), h, m));

            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "routeSheet";

            worksheet[0, 0].SetValue("xcoord");
            worksheet[0, 1].SetValue("ycoord");
            worksheet[0, 2].SetValue("val");

            int rowOrder = 0;
            for (int i = 1; i < strs.Length; i = i + 2)
            {
                rowOrder++;

                try
                {
                    coordVal = new CoordVal();

                    coordVal.coordX = Convert.ToInt32(strs[i].Split(',')[0]);
                    coordVal.coordY = Convert.ToInt32(strs[i].Split(',')[1]);
                    coordVal.val = Convert.ToInt32(strs[i + 1].Replace(":", "").Replace(",", ""));

                    //Console.WriteLine(string.Format("x:{0}, y:{1}, val:{2}", coordVal.coordX, coordVal.coordY, coordVal.val));

                    worksheet[rowOrder, 0].SetValue(coordVal.coordX);
                    worksheet[rowOrder, 1].SetValue(coordVal.coordY);
                    worksheet[rowOrder, 2].SetValue(coordVal.val);
                }
                catch (Exception ex)
                {
                    coordVal.coordX = 0;
                    coordVal.coordY = 0;
                    coordVal.val = 0;

                    worksheet[rowOrder, 0].SetValue(0);
                    worksheet[rowOrder, 1].SetValue(0);
                    worksheet[rowOrder, 2].SetValue(0);
                }

            }

            workbook.SaveDocument(string.Format(@"{0}\r{1}_{2}{3}.xls", path, date.Replace("-", ""), h, m));
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            DialogResult dr = XtraMessageBox.Show("是否退出当前任务", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                if (thread != null)
                {
                    thread.Abort();
                }

                this.Close();
            }
        }

        bool taskExecuted = false;
        private void btn_saveClose_Click(object sender, EventArgs e)
        {
            if (taskExecuted)
            {
                if (checkEdit_addToView.CheckState == CheckState.Checked)
                {
                    //spreadsheetControl.LoadDocument(docPath);

                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                }

            }
            else
            {
                XtraMessageBox.Show("未执行成功任何任务！");
            }
            this.Close();
        }

        /// <summary>
        /// 人流量数据采集命令行方法
        /// </summary>
        /// <param name="str_path">存储路径</param>
        /// <param name="str_num">区域编号</param>
        /// <param name="timeInterval">时间间隔</param>
        /// <param name="str_date">采集日期</param>
        public void Execute_by_command(string str_path, string str_num, int timeInterval, string str_date)
        {
            cmd_fun1(str_path, timeInterval,str_date, str_num);
        }

        private void cmd_fun1(string str_path, int timeInterval, string str_date, string str_num)
        {
            string path = str_path;
            string date = str_date;
            time_interval = timeInterval;

            string h, m;

            for (int i = 0; i < 24; i++)
            {
                for (int j = 0; j < 60; j = j + time_interval)
                {
                    if (i < 10)
                    {
                        h = "0" + i;
                    }
                    else
                    {
                        h = i.ToString();
                    }
                    if (j < 10)
                    {
                        m = "0" + j;
                    }
                    else
                    {
                        m = j.ToString();
                    }


                    cmd_fun(path, date, h, m, "00", str_num);
                    
                }
            }
        }

        /// <summary>
        /// 数据采集
        /// </summary>
        /// <param name="path">文件存储路径</param>
        /// <param name="date">采集时间</param>
        /// <param name="h">时</param>
        /// <param name="m">分</param>
        /// <param name="s">秒</param>
        /// <param name="n">区域编号</param>
        public void cmd_fun(string path, string date, string h, string m, string s, string n)
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            //庐山 1179
            //北京欢乐谷 5381
            //丽江古城 894
            //首都机场786
            //厦门大学6423
            string str = string.Format(@"https://heat.qq.com/api/getHeatDataByTime.php?region_id={4}&datetime={0}+{1}:{2}:{3}", date, h, m, s,n);
            Stream stream = client.OpenRead(str);

            string str_json = new StreamReader(stream).ReadToEnd();

            str_json = str_json.Replace("{", "").Replace("}", "");

            string[] strs = str_json.Split('"');

            CoordVal coordVal = null;

            Workbook workbook = new Workbook();
            workbook.CreateNewDocument();
            workbook.SaveDocument(string.Format(@"{0}\r{1}_{2}{3}.xls", path, date.Replace("-", ""), h, m));

            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "routeSheet";

            worksheet[0, 0].SetValue("xcoord");
            worksheet[0, 1].SetValue("ycoord");
            worksheet[0, 2].SetValue("val");

            int rowOrder = 0;
            for (int i = 1; i < strs.Length; i = i + 2)
            {
                rowOrder++;

                coordVal = new CoordVal();

                

                try
                {
                    coordVal.coordX = Convert.ToInt32(strs[i].Split(',')[0]);
                    coordVal.coordY = Convert.ToInt32(strs[i].Split(',')[1]);
                    coordVal.val = Convert.ToInt32(strs[i + 1].Replace(":", "").Replace(",", ""));

                    //Console.WriteLine(string.Format("x:{0}, y:{1}, val:{2}", coordVal.coordX, coordVal.coordY, coordVal.val));

                    worksheet[rowOrder, 0].SetValue(coordVal.coordX);
                    worksheet[rowOrder, 1].SetValue(coordVal.coordY);
                    worksheet[rowOrder, 2].SetValue(coordVal.val);
                }
                catch (Exception ex)
                {
                    coordVal.coordX = 0;
                    coordVal.coordY = 0;
                    coordVal.val = 0;

                    worksheet[rowOrder, 0].SetValue(0);
                    worksheet[rowOrder, 1].SetValue(0);
                    worksheet[rowOrder, 2].SetValue(0);
                }

            }

            workbook.SaveDocument(string.Format(@"{0}\r{1}_{2}{3}.xls", path, date.Replace("-", ""), h, m));
        }
    }
}
