using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraSpreadsheet;
using Newtonsoft.Json;
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
using System.Xml;

namespace GeoSharp2018.ToolForms
{
    public partial class ForecastColleForm : DevExpress.XtraEditors.XtraForm
    {

        //定义delegate以便Invoke时使用  
        private delegate void SetProgressBarValue(int value);
        /// <summary>
        /// 系统表格控件
        /// </summary>
        private SpreadsheetControl spreadsheetControl;
        /// <summary>
        /// 当前工作簿
        /// </summary>
        private Workbook workbook;
        /// <summary>
        /// 当前工作表
        /// </summary>
        private Worksheet worksheet;
        /// <summary>
        /// 解析字段列序号
        /// </summary>
        int col;
        /// <summary>
        /// 解析城市字段序号
        /// </summary>
        int col1;
        /// <summary>
        /// 是否是第一次加载
        /// </summary>
        private bool isFirstLoad = true;
        /// <summary>
        /// 文件路径
        /// </summary>
        string docPath;
        /// <summary>
        /// 是否有完成的执行任务
        /// </summary>
        bool taskExecuted = false;
        /// <summary>
        /// 主窗口对象
        /// </summary>
        private MainForm mainform;
       /// <summary>
       /// 采集数据主线程
       /// </summary>
        Thread thread = null;

        string strKey = "27caa753be4090132a65386ed3efff97";
        int sum = 0;
        string forecastCodeFile = null;

        public ForecastColleForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void ForecastColleForm_Load(object sender, EventArgs e)
        {


                comboBox_targetObj.Properties.Items.Add("地市类城市");
                comboBox_targetObj.Properties.Items.Add("区县类城市");

                comboBox_targetObj.SelectedIndex = 0;


        }

        private void btn_open_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;
            string fileName;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "excel2003文件|*.xls|excel文件|*.xlsx";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FilterIndex = 2;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                docPath = fileName = saveFileDialog.FileName;

                textEdit_file.EditValue = fileName;

                workbook = new Workbook();

                workbook.CreateNewDocument();

                workbook.SaveDocument(textEdit_file.EditValue.ToString());

                textEdit_fileName.EditValue = "main_sheet";

                isFirstLoad = false;
            }
        }


        string target_type;
        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                target_type = comboBox_targetObj.Properties.Items[comboBox_targetObj.SelectedIndex].ToString();

                if (target_type == "地市类城市")
                {
                    forecastCodeFile = Application.StartupPath + @"\data\city_code_forecast.xlsx";
                }
                else if (target_type == "区县类城市")
                {
                    forecastCodeFile = Application.StartupPath + @"\data\county_code_forecast.xlsx";
                }
                else
                {
                    XtraMessageBox.Show("所选采集对象类型不合法！请正确配置相关参数！");
                    return;
                }

                try
                {
                    worksheet = workbook.Worksheets[0];
                    worksheet.Name = textEdit_fileName.EditValue.ToString();

                    worksheet[0, 0].SetValue("id");
                    worksheet[0, 1].SetValue("province");
                    worksheet[0, 2].SetValue("city");
                    worksheet[0, 3].SetValue("adcode");

                    worksheet[0, 4].SetValue("weather");
                    worksheet[0, 5].SetValue("temperature");
                    worksheet[0, 6].SetValue("winddirection");
                    worksheet[0, 7].SetValue("windpower");
                    worksheet[0, 8].SetValue("humidity");
                    worksheet[0, 9].SetValue("reporttime");

                    Workbook tem_workbook = new Workbook();
                    tem_workbook.LoadDocument(forecastCodeFile);
                    Range range = tem_workbook.Worksheets[0].GetUsedRange();

                    progressBarControl1.Properties.Maximum = range.RowCount;

                    thread = new Thread(new ThreadStart(fun1));
                    thread.Start();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }

                btn_ok.Enabled = false;
            }
        }
      
        private string SetAddress(string strKey, string cityCode)
        {
            string s = "";
            try
            {
                s = string.Format(@"http://restapi.amap.com/v3/weather/weatherInfo?city={1}&key={0}&extensions=base", strKey, cityCode);

            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            return s;
        }

        private void fun1()
        {
            Workbook tem_workbook = new Workbook();
            tem_workbook.LoadDocument(forecastCodeFile);

            Range range = tem_workbook.Worksheets[0].GetUsedRange();

            WebClientto client = new WebClientto(4500);
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            string str_json = null;
            ForecastInfo forcastInfo = null;
            List<ForecastInfo> list = new List<ForecastInfo>();

            for (int i = 1; i < range.RowCount; i++)
            {
                try
                {
                    stream = client.OpenRead(SetAddress(strKey, tem_workbook.Worksheets[0][i, 1].Value.ToString()));
                }
                catch(Exception ex)
                {
                    MessageBox.Show("操作超市,当前工作将自动退出。请在稳定的网络环境下执行此任务！");
                    return;

                    if (thread.ThreadState == ThreadState.Running)
                    {
                        thread.Abort();
                    }
                    this.Close();
                }

                

                str_json = new StreamReader(stream).ReadToEnd();

                if (str_json != "")
                {
                    try
                    {
                        forcastInfo = JsonConvert.DeserializeObject<ForecastInfo>(str_json);

                        list.Add(forcastInfo);
                   }
                    catch(Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                sum++;
                RunWithInoke(sum);

                if (i == 1000)
                {
                    MessageBox.Show(progressBarControl1.EditValue.ToString());
                }
            }

            for (int i = 0; i < list.Count; i++)
            {
                worksheet[i + 1, 0].SetValue(i);
                worksheet[i + 1, 1].SetValue(list[i].lives[0].province);
                worksheet[i + 1, 2].SetValue(list[i].lives[0].city);
                worksheet[i + 1, 3].SetValue(list[i].lives[0].adcode);
                worksheet[i + 1, 4].SetValue(list[i].lives[0].weather);
                worksheet[i + 1, 5].SetValue(list[i].lives[0].temperature);
                worksheet[i + 1, 6].SetValue(list[i].lives[0].winddirection);
                worksheet[i + 1, 7].SetValue(list[i].lives[0].windpower);
                worksheet[i + 1, 8].SetValue(list[i].lives[0].humidity);
                worksheet[i + 1, 9].SetValue(list[i].lives[0].reporttime); 
            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);


            
            XtraMessageBox.Show("所有天气信息已经解析完成！");

        }

        private void RunWithInoke(int i)
        {

            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);

        }

        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;
        }

        private void textEdit_file_EditValueChanged(object sender, EventArgs e)
        {
            dxValidationProvider1.Validate();
        }

        private void textEdit_fileName_EditValueChanged(object sender, EventArgs e)
        {
            dxValidationProvider1.Validate();
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            DialogResult dr =  XtraMessageBox.Show("是否退出当前任务", "提示", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                if (thread != null)
                {
                    thread.Abort();
                }

                this.Close();
            }
        }

        private void btn_saveClose_Click(object sender, EventArgs e)
        {
            if (taskExecuted)
            {
                if (checkEdit_addToView.CheckState == CheckState.Checked)
                {
                    spreadsheetControl.LoadDocument(docPath);

                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                }

            }
            else
            {
                XtraMessageBox.Show("未执行成功任何任务！");
            }
            this.Close();
        }

        private void comboBox_targetType_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }   
    }

    public class ForecastInfo
    {
        public string status { get; set; }
        public string count { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }

        public List<LiveInfo> lives { get; set; }
    }

    public class LiveInfo
    {
        public string province{get;set;}
        public string city{get;set;}
        public string adcode{get;set;}
        public string weather{get;set;}
        public string temperature{get;set;}
        public string winddirection{get;set;}
        public string windpower{get;set;}
        public string humidity{get;set;}
        public string reporttime{get;set;}
    }

    public class WebClientto : WebClient
    {
        /// <summary>  
        /// 过期时间  
        /// </summary>  
        public int Timeout { get; set; }

        public WebClientto(int timeout)
        {
            Timeout = timeout;
        }

        /// <summary>  
        /// 重写GetWebRequest,添加WebRequest对象超时时间  
        /// </summary>  
        /// <param name="address"></param>  
        /// <returns></returns>  
        protected override WebRequest GetWebRequest(Uri address)
        {
            HttpWebRequest request = (HttpWebRequest)base.GetWebRequest(address);
            request.Timeout = Timeout;
            request.ReadWriteTimeout = Timeout;
            return request;
        }
    }  

   
}
