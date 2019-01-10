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

namespace GeoSharp2018.ToolForms
{
    public partial class ForecastAllColleForm : DevExpress.XtraEditors.XtraForm
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

        public ForecastAllColleForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void ForecastAllColleForm_Load(object sender, EventArgs e)
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
                    worksheet[0, 1].SetValue("city");
                    worksheet[0, 2].SetValue("adcode");
                    worksheet[0, 3].SetValue("province");
                    worksheet[0, 4].SetValue("reporttime");

                    worksheet[0, 5].SetValue("day_order");
                    worksheet[0, 6].SetValue("date");
                    worksheet[0, 7].SetValue("week");
                    worksheet[0, 8].SetValue("dayweather");
                    worksheet[0, 9].SetValue("nightweather");
                    worksheet[0, 10].SetValue("daytemp");
                    worksheet[0, 11].SetValue("nighttemp");
                    worksheet[0, 12].SetValue("daywind");
                    worksheet[0, 13].SetValue("nightwind");
                    worksheet[0, 14].SetValue("daypower");
                    worksheet[0, 15].SetValue("nightpower");

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

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            DialogResult dr = XtraMessageBox.Show("提示", "是否退出当前任务？", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning);

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

        private void btn_help_Click(object sender, EventArgs e)
        {

        }

        private string SetAddress(string strKey, string cityCode)
        {
            string s = "";
            try
            {
                s = string.Format(@"http://restapi.amap.com/v3/weather/weatherInfo?city={1}&key={0}&extensions=all", strKey, cityCode);

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

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            string str_json = null;
            ForecastAllInfo forcastAllInfo = null;
            List<ForecastAllInfo> list = new List<ForecastAllInfo>();

            for (int i = 1; i < range.RowCount; i++)
            {
                stream = client.OpenRead(SetAddress(strKey, tem_workbook.Worksheets[0][i, 1].Value.ToString()));
                str_json = new StreamReader(stream).ReadToEnd();

                if (str_json != "")
                {
                    try
                    {
                        forcastAllInfo = JsonConvert.DeserializeObject<ForecastAllInfo>(str_json);

                        list.Add(forcastAllInfo);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                sum++;
                RunWithInoke(sum);

                if (i == 1000)
                {
                    //MessageBox.Show(progressBarControl1.EditValue.ToString());
                }
            }

            int castsCount = -1;
            int rowOrder = -1;
            for (int i = 0; i < list.Count; i++)
            {
                castsCount = list[i].forecasts[0].casts.Count;
                for (int j = 0; j < castsCount; j++)
                {
                    rowOrder = i * castsCount + (j + 1);

                    worksheet[rowOrder, 0].SetValue(i * castsCount + (j + 1));
                    worksheet[rowOrder, 1].SetValue(list[i].forecasts[0].city);
                    worksheet[rowOrder, 2].SetValue(list[i].forecasts[0].adcode);
                    worksheet[rowOrder, 3].SetValue(list[i].forecasts[0].province);
                    worksheet[rowOrder, 4].SetValue(list[i].forecasts[0].reporttime);


                    worksheet[rowOrder, 5].SetValue(j + 1);
                    worksheet[rowOrder, 6].SetValue(list[i].forecasts[0].casts[j].date);
                    worksheet[rowOrder, 7].SetValue(list[i].forecasts[0].casts[j].week);
                    worksheet[rowOrder, 8].SetValue(list[i].forecasts[0].casts[j].dayweather);
                    worksheet[rowOrder, 9].SetValue(list[i].forecasts[0].casts[j].nightweather);
                    worksheet[rowOrder, 10].SetValue(list[i].forecasts[0].casts[j].daytemp);
                    worksheet[rowOrder, 11].SetValue(list[i].forecasts[0].casts[j].nighttemp);
                    worksheet[rowOrder, 12].SetValue(list[i].forecasts[0].casts[j].daywind);
                    worksheet[rowOrder, 13].SetValue(list[i].forecasts[0].casts[j].nightwind);
                    worksheet[rowOrder, 14].SetValue(list[i].forecasts[0].casts[j].daypower);
                    worksheet[rowOrder, 15].SetValue(list[i].forecasts[0].casts[j].nightpower);
                    
                }
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
    }

    public class ForecastAllInfo
    {
        public string status { get; set; }
        public string count { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }

        public List<Forecasts> forecasts { get; set; }
    }

    public class Forecasts
    {
        public string city { get; set; }
        public string adcode { get; set; }
        public string province { get; set; }
        public string reporttime { get; set; }

        public List<Casts> casts { get; set; }
        
    }

    public class Casts
    {
        public string date { get; set; }
        public string week { get; set; }
        public string dayweather { get; set; }
        public string nightweather { get; set; }
        public string daytemp { get; set; }
        public string nighttemp { get; set; }
        public string daywind { get; set; }
        public string nightwind { get; set; }
        public string daypower { get; set; }
        public string nightpower { get; set; }
    }
}
