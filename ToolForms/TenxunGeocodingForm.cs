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
    public partial class TenxunGeocodingForm : DevExpress.XtraEditors.XtraForm
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

        public TenxunGeocodingForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void TenxunGeocodingForm_Load(object sender, EventArgs e)
        {
            textEdit_lngFieldName.EditValue = "Tenxun_Lng";
            textEdit_latFieldName.EditValue = "Tenxun_Lat";
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
                    worksheet[0, colCount + 3].SetValue("f_message");

                    worksheet[0, colCount + 4].SetValue("f_title");
                    worksheet[0, colCount + 5].SetValue("f_province");
                    worksheet[0, colCount + 6].SetValue("f_city");
                    worksheet[0, colCount + 7].SetValue("f_district");
                    worksheet[0, colCount + 8].SetValue("f_street");
                    worksheet[0, colCount + 9].SetValue("f_number");
                    worksheet[0, colCount + 10].SetValue("f_similarity");
                    worksheet[0, colCount + 11].SetValue("f_deviation");
                    worksheet[0, colCount + 12].SetValue("f_reliability");
                    worksheet[0, colCount + 13].SetValue("f_level");

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
            List<GeoCodingAllInfoTenxun> list = new List<GeoCodingAllInfoTenxun>();

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

                    GeoCodingAllInfoTenxun geocodeAllInfo = new GeoCodingAllInfoTenxun();
                    ResultInfo resultInfo = new ResultInfo();
                    LocationInfo locationInfo = new LocationInfo();
                    AddressComInfo addressComInfo = new AddressComInfo();

                    JObject obj = JObject.Parse(str_json);

                    geocodeAllInfo.status = obj["status"].ToString();
                    geocodeAllInfo.message = obj["message"].ToString();

                    resultInfo.title = obj["result"]["title"].ToString();

                    locationInfo.Lng = obj["result"]["location"]["lng"].ToString();
                    locationInfo.Lat = obj["result"]["location"]["lat"].ToString();

                    resultInfo.locationInfo = locationInfo;

                    addressComInfo.province = obj["result"]["address_components"]["province"].ToString();
                    addressComInfo.city = obj["result"]["address_components"]["city"].ToString();
                    addressComInfo.district = obj["result"]["address_components"]["district"].ToString();
                    addressComInfo.street = obj["result"]["address_components"]["street"].ToString();
                    addressComInfo.street_number = obj["result"]["address_components"]["street_number"].ToString();

                    resultInfo.addressComInfo = addressComInfo;

                    resultInfo.similarity = obj["result"]["similarity"].ToString();
                    resultInfo.similarity = obj["result"]["deviation"].ToString();
                    resultInfo.similarity = obj["result"]["reliability"].ToString();
                    resultInfo.similarity = obj["result"]["level"].ToString();

                    geocodeAllInfo.resultInfo = resultInfo;
                    

                    list.Add(geocodeAllInfo);
                }
                sum++;
                RunWithInoke(sum);

                Thread.Sleep(500);
            }

            for (int i = 0; i < list.Count; i++)
            {
                //MessageBox.Show(list[i].geocodes.location.ToString());
                worksheet[i + 1, colCount].SetValue(list[i].resultInfo.locationInfo.Lng);
                worksheet[i + 1, colCount + 1].SetValue(list[i].resultInfo.locationInfo.Lat);

                worksheet[i + 1, colCount + 2].SetValue(list[i].status);
                worksheet[i + 1, colCount + 3].SetValue(list[i].message);
                worksheet[i + 1, colCount + 4].SetValue(list[i].resultInfo.title);
                worksheet[i + 1, colCount + 5].SetValue(list[i].resultInfo.addressComInfo.province);
                worksheet[i + 1, colCount + 6].SetValue(list[i].resultInfo.addressComInfo.city);
                worksheet[i + 1, colCount + 7].SetValue(list[i].resultInfo.addressComInfo.district);
                worksheet[i + 1, colCount + 8].SetValue(list[i].resultInfo.addressComInfo.street);
                worksheet[i + 1, colCount + 9].SetValue(list[i].resultInfo.addressComInfo.street_number);
                worksheet[i + 1, colCount + 10].SetValue(list[i].resultInfo.similarity);
                worksheet[i + 1, colCount + 11].SetValue(list[i].resultInfo.deviation);
                worksheet[i + 1, colCount + 12].SetValue(list[i].resultInfo.reliability);
                worksheet[i + 1, colCount + 13].SetValue(list[i].resultInfo.level);

            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);

            XtraMessageBox.Show("所有数据的地理编码任务已经完成！");
        }

        string gaodeKey = "CUDBZ-MQ3W4-TTWUQ-DJDUP-ZNHHZ-W6BSC";
        private string SetAddress(string address, string city, string strKey)
        {
            string s = string.Format(@"https://apis.map.qq.com/ws/geocoder/v1/?address={0}&region={1}&key={2}", address, city, strKey);

            return s;
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

    public class GeoCodingAllInfoTenxun
    {
        public string status { get; set; }
        public string message { get; set; }
        public ResultInfo resultInfo { get; set; }
    }

    public class ResultInfo
    {
        public string title { get; set; }
        public LocationInfo locationInfo { get; set; }
        public AddressComInfo addressComInfo { get; set; }
        public string similarity { get; set; }
        public string deviation { get; set; }
        public string reliability { get; set; }
        public string level { get; set; }
    }

    public class LocationInfo
    {
        public string Lng { get; set; }
        public string Lat { get; set; }
    }

    public class AddressComInfo
    {
        public string province { get; set; }
        public string city { get; set; }
        public string district { get; set; }
        public string street { get; set; }
        public string street_number { get; set; }
    }
}
