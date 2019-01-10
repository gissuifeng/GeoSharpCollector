using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraSpreadsheet;
using Newtonsoft.Json.Linq;
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
    public partial class BaiduGeocodingForm : DevExpress.XtraEditors.XtraForm
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
        /// 主程序窗口
        /// </summary>
        private MainForm mainForm;
        /// <summary>
        /// 任务线程
        /// </summary>
        private Thread thread;
        /// <summary>
        /// 计算解析地址数量
        /// </summary>
        private int sum = 0;

        public BaiduGeocodingForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void BaiduGeocodingForm_Load(object sender, EventArgs e)
        {
            textEdit_lngFieldName.EditValue = "Gaode_Lng";
            textEdit_latFieldName.EditValue = "Gaode_Lat";
        }

        private void btn_open_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;
            comboBox_addressField.Properties.Items.Clear();
            comboBox_cityField.Properties.Items.Clear();
            comboBox_worksheet.Properties.Items.Clear();

            string fileName;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "excel2003文件|*.xls|excel文件|*.xlsx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 2;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                docPath = fileName = openFileDialog.FileName;

                textEdit_file.EditValue = fileName;

                workbook = new Workbook();

                workbook.LoadDocument(textEdit_file.EditValue.ToString());

                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    comboBox_worksheet.Properties.Items.Add(workbook.Worksheets[i].Name);
                }

                comboBox_worksheet.SelectedIndex = 0;

                worksheet = workbook.Worksheets[0];

                for (int j = 0; j < worksheet.GetDataRange().ColumnCount; j++)
                {
                    comboBox_addressField.Properties.Items.Add(worksheet[0, j].Value);
                    comboBox_cityField.Properties.Items.Add(worksheet[0, j].Value);
                }

                comboBox_addressField.SelectedIndex = 0;
                comboBox_cityField.SelectedIndex = 0;

                isFirstLoad = false;
            }
        }


        string[] strs;
        string[] strs1;
        int colCount;
        int rowCount;
        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                try
                {
                    col = comboBox_addressField.SelectedIndex;
                    col1 = comboBox_cityField.SelectedIndex;

                    colCount = worksheet.GetDataRange().ColumnCount;
                    rowCount = worksheet.GetDataRange().RowCount;

                    strs = new string[rowCount];
                    strs1 = new string[rowCount];

                    progressBarControl1.Properties.Maximum = rowCount;

                    for (int i = 1; i < rowCount; i++)
                    {
                        strs[i] = worksheet[i, col].Value.ToString();
                        strs1[i] = worksheet[i, col1].Value.ToString();
                    }

                    worksheet[0, colCount].SetValue(textEdit_lngFieldName.EditValue.ToString());
                    worksheet[0, colCount + 1].SetValue(textEdit_latFieldName.EditValue.ToString());

                    worksheet[0, colCount + 2].SetValue("f_status");
                    worksheet[0, colCount + 3].SetValue("f_precise");
                    worksheet[0, colCount + 4].SetValue("f_confidence");

                    worksheet[0, colCount + 5].SetValue("comprehension");
                    worksheet[0, colCount + 6].SetValue("level");

                    thread = new Thread(new ThreadStart(fun));
                    thread.Start();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }
            }
        }

        private void comboBox_worksheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox_addressField.Properties.Items.Clear();
            comboBox_cityField.Properties.Items.Clear();
            if (!isFirstLoad)
            {


                string targetSheetName = comboBox_worksheet.SelectedText;

                if (targetSheetName != "")
                {
                    worksheet = workbook.Worksheets[targetSheetName];

                    for (int j = 0; j < worksheet.GetDataRange().ColumnCount; j++)
                    {
                        comboBox_addressField.Properties.Items.Add(worksheet[0, j].Value);
                        comboBox_cityField.Properties.Items.Add(worksheet[0, j].Value);
                    }

                    comboBox_addressField.SelectedIndex = 0;
                    comboBox_cityField.SelectedIndex = 0;
                }
            }
        }

        private void fun()
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            string str_json = null;
            List<GeoCodingAllInfoBaidu> list = new List<GeoCodingAllInfoBaidu>();

            string addressName;
            string cityName;

            for (int i = 1; i < rowCount; i++)
            {
                addressName = worksheet[i, comboBox_addressField.SelectedIndex].Value.ToString();
                cityName = worksheet[i, comboBox_cityField.SelectedIndex].Value.ToString();

                try
                {
                    stream = client.OpenRead(SetAddress(addressName, cityName, gaodeKey));

                    str_json = new StreamReader(stream).ReadToEnd();
                }
                catch
                {

                }

                if (str_json != "")
                {
                    GeoCodingAllInfoBaidu geocodeAllInfo = new GeoCodingAllInfoBaidu();
                    ResultsInfoBaidu resultBaidu = new ResultsInfoBaidu();
                    LocationInfoBaidu locationInfoBaidu = new LocationInfoBaidu();
                    

                    JObject obj = JObject.Parse(str_json);

                    geocodeAllInfo.status = obj["status"].ToString();

                    locationInfoBaidu.lng = obj["result"]["location"]["lng"].ToString();
                    locationInfoBaidu.lat = obj["result"]["location"]["lat"].ToString();

                    resultBaidu.location = locationInfoBaidu;

                    resultBaidu.precise = obj["result"]["precise"].ToString();
                    resultBaidu.confidence = obj["result"]["confidence"].ToString();
                    resultBaidu.comprehension = obj["result"]["comprehension"].ToString();
                    resultBaidu.level = obj["result"]["level"].ToString();

                    geocodeAllInfo.result = resultBaidu;

                    list.Add(geocodeAllInfo);
                }
                sum++;
                RunWithInoke(sum);
            }

            for (int i = 0; i < list.Count; i++)
            {
                //MessageBox.Show(list[i].geocodes.location.ToString());
                worksheet[i + 1, colCount].SetValue(list[i].result.location.lng);
                worksheet[i + 1, colCount + 1].SetValue(list[i].result.location.lat);

                worksheet[i + 1, colCount + 2].SetValue(list[i].status);
                worksheet[i + 1, colCount + 3].SetValue(list[i].result.precise);
                worksheet[i + 1, colCount + 4].SetValue(list[i].result.confidence);
                worksheet[i + 1, colCount + 5].SetValue(list[i].result.comprehension);
                worksheet[i + 1, colCount + 6].SetValue(list[i].result.level);
            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);

            XtraMessageBox.Show("所有数据的地理编码任务已经完成！");
        }

        string gaodeKey = "NTkLYjCEkr7aunaFznXElWUBAA4SrBi6";
        private string SetAddress(string address, string city, string strKey)
        {
            return string.Format(@"http://api.map.baidu.com/geocoder/v2/?address={0}&output=json&ak={2}&city={1}", address, city, strKey);
        }

        private void RunWithInoke(int i)
        {
            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);
        }

        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            if (thread != null)
            {
                thread.Abort();
            }

            this.Close();
        }

        private void btn_cancel_Click_1(object sender, EventArgs e)
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
    }

    public class GeoCodingAllInfoBaidu
    {
        public string status { get; set; }

        public ResultsInfoBaidu result { get; set; }
       
    }

    public class ResultsInfoBaidu
    {
        public string precise { get; set; }
        public string confidence { get; set; }
        public string comprehension { get; set; }
        public string level { get; set; }

        public LocationInfoBaidu location { get; set; }
    }

    public class LocationInfoBaidu
    {
        public string lng { get; set; }

        public string lat { get; set; }
    }
}
