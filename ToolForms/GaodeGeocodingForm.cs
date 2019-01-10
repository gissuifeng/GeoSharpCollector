using DevExpress.Spreadsheet;
using DevExpress.Utils;
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
    public partial class GaodeGeocodingForm : DevExpress.XtraEditors.XtraForm
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

        DateTime pro_dt;
        public GaodeGeocodingForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void GaodeGeocodingForm_Load(object sender, EventArgs e)
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
                    pro_dt = DateTime.Now;

                    for (int i = 1; i < rowCount; i++)
                    {
                        strs[i] = worksheet[i, col].Value.ToString();
                        strs1[i] = worksheet[i, col1].Value.ToString();
                    }

                    worksheet[0, colCount].SetValue(textEdit_lngFieldName.EditValue.ToString());
                    worksheet[0, colCount + 1].SetValue(textEdit_latFieldName.EditValue.ToString());

                    worksheet[0, colCount + 2].SetValue("f_status");
                    worksheet[0, colCount + 3].SetValue("f_info");
                    worksheet[0, colCount + 4].SetValue("f_count");

                    worksheet[0, colCount + 5].SetValue("formatted_address");
                    worksheet[0, colCount + 6].SetValue("province");
                    worksheet[0, colCount + 7].SetValue("city");
                    worksheet[0, colCount + 8].SetValue("citycode");
                    worksheet[0, colCount + 9].SetValue("district");
                    worksheet[0, colCount + 10].SetValue("township");
                    worksheet[0, colCount + 11].SetValue("street");
                    worksheet[0, colCount + 12].SetValue("number");
                    worksheet[0, colCount + 13].SetValue("adcode");
                    worksheet[0, colCount + 14].SetValue("level");
                   



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
            
        }

        private void fun()
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            string str_json = null;
            List<GeoCodingAllInfo> list = new List<GeoCodingAllInfo>();

            string addressName;
            string cityName;

            int i = 1;
            for (; i < rowCount; i++)
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
                    GeoCodingAllInfo geocodeAllInfo = new GeoCodingAllInfo();
                    Geocodes geocodes = new Geocodes();

                    JObject obj = JObject.Parse(str_json);

                    if (obj["status"].ToString() == "0")
                    {

                        geocodeAllInfo.status = "";
                        geocodeAllInfo.info = "";
                        geocodeAllInfo.infocode = "";
                        geocodeAllInfo.count = "";
                    }
                    else
                    {
                        geocodeAllInfo.status = obj["status"].ToString();
                        geocodeAllInfo.info = obj["info"].ToString();
                        geocodeAllInfo.infocode = obj["infocode"].ToString();
                        geocodeAllInfo.count = obj["count"].ToString();
                    }

                    if(geocodeAllInfo.count == "1")
                    {
                        JArray jlist = JArray.Parse(obj["geocodes"].ToString());
                        JObject obj1 = JObject.Parse(jlist[0].ToString());

                        geocodes.formatted_address = obj1["formatted_address"].ToString();
                        geocodes.province = obj1["province"].ToString();
                        geocodes.city = obj1["city"].ToString();
                        geocodes.citycode = obj1["city"].ToString();
                        geocodes.district = obj1["district"].ToString();
                        geocodes.township = obj1["township"].ToString();
                        geocodes.adcode = obj1["adcode"].ToString();
                        geocodes.street = obj1["street"].ToString();
                        geocodes.number = obj1["number"].ToString();
                        geocodes.level = obj1["level"].ToString();
                        geocodes.location = obj1["location"].ToString();
                    }
                    else
                    {
                        geocodes.formatted_address = "";
                        geocodes.province = "";
                        geocodes.city = "";
                        geocodes.citycode = "";
                        geocodes.district = "";
                        geocodes.township = "";
                        geocodes.adcode = "";
                        geocodes.street = "";
                        geocodes.number = "";
                        geocodes.level = "";
                        geocodes.location = "";
                    }

                    geocodeAllInfo.geocodes = geocodes;

                    list.Add(geocodeAllInfo);
                }
                sum++;
                RunWithInoke(sum);

                if((i % 1000) == 0 && (i / 1000) > 0)
                {
                    int cur_count = ((i / 1000)-1) * 1000 - 1;

                    for (int j = 0; j < list.Count; j++)
                    {
                        cur_count++;
                        //MessageBox.Show(list[i].geocodes.location.ToString());
                        if (list[j].geocodes.location != "")
                        {
                            worksheet[cur_count + 1, colCount].SetValue(list[j].geocodes.location.Split(',')[0]);
                            worksheet[cur_count + 1, colCount + 1].SetValue(list[j].geocodes.location.Split(',')[1]);
                        }
                        else
                        {
                            worksheet[cur_count + 1, colCount].SetValue("");
                            worksheet[cur_count + 1, colCount + 1].SetValue("");

                        }

                        worksheet[cur_count + 1, colCount + 2].SetValue(list[j].status);
                        worksheet[cur_count + 1, colCount + 3].SetValue(list[j].info);
                        worksheet[cur_count + 1, colCount + 4].SetValue(list[j].count);
                        worksheet[cur_count + 1, colCount + 5].SetValue(list[j].geocodes.formatted_address);
                        worksheet[cur_count + 1, colCount + 6].SetValue(list[j].geocodes.province);
                        worksheet[cur_count + 1, colCount + 7].SetValue(list[j].geocodes.city);
                        worksheet[cur_count + 1, colCount + 8].SetValue(list[j].geocodes.citycode);
                        worksheet[cur_count + 1, colCount + 9].SetValue(list[j].geocodes.district);
                        worksheet[cur_count + 1, colCount + 10].SetValue(list[j].geocodes.township);
                        worksheet[cur_count + 1, colCount + 11].SetValue(list[j].geocodes.street);
                        worksheet[cur_count + 1, colCount + 12].SetValue(list[j].geocodes.number);
                        worksheet[cur_count + 1, colCount + 13].SetValue(list[j].geocodes.adcode);

                        worksheet[cur_count + 1, colCount + 14].SetValue(list[j].geocodes.level);
                    }

                    workbook.SaveDocument(docPath);

                    list.Clear();
                }
            }

            int restRows = (rowCount) - (rowCount % 1000);

            for (int m = 0; m < list.Count; m++)
            {
                if (list[m].geocodes.location != "")
                {
                    worksheet[restRows + m + 1, colCount].SetValue(list[m].geocodes.location.Split(',')[0]);
                    worksheet[restRows + m + 1, colCount + 1].SetValue(list[m].geocodes.location.Split(',')[1]);
                }
                else
                {
                    worksheet[restRows + m + 1, colCount].SetValue("");
                    worksheet[restRows + m + 1, colCount + 1].SetValue("");
                }

                worksheet[restRows + m + 1, colCount + 2].SetValue(list[m].status);
                worksheet[restRows + m + 1, colCount + 3].SetValue(list[m].info);
                worksheet[restRows + m + 1, colCount + 4].SetValue(list[m].count);
                worksheet[restRows + m + 1, colCount + 5].SetValue(list[m].geocodes.formatted_address);
                worksheet[restRows + m + 1, colCount + 6].SetValue(list[m].geocodes.province);
                worksheet[restRows + m + 1, colCount + 7].SetValue(list[m].geocodes.city);
                worksheet[restRows + m + 1, colCount + 8].SetValue(list[m].geocodes.citycode);
                worksheet[restRows + m + 1, colCount + 9].SetValue(list[m].geocodes.district);
                worksheet[restRows + m + 1, colCount + 10].SetValue(list[m].geocodes.township);
                worksheet[restRows + m + 1, colCount + 11].SetValue(list[m].geocodes.street);
                worksheet[restRows + m + 1, colCount + 12].SetValue(list[m].geocodes.number);
                worksheet[restRows + m + 1, colCount + 13].SetValue(list[m].geocodes.adcode);

                worksheet[restRows + m + 1, colCount + 14].SetValue(list[m].geocodes.level);    
            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);

            XtraMessageBox.Show("所有数据的地理编码任务已经完成！");
        }

        string gaodeKey = "1ef994aee94edd79c70ea991d98c046c";//"e00536b393e9671af12bea182f75a36b";
        private string SetAddress(string address, string city, string strKey)
        {
            return string.Format(@"http://restapi.amap.com/v3/geocode/geo?key={2}&output=json&address={0}&city={1}", address, city, strKey);
        }

        private void RunWithInoke(int i)
        {

            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);

        }


        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;

            DateTime temp_dt = DateTime.Now;

            progressBarControl1.SuperTip.Items.AddTitle("任务当前信息");
            progressBarControl1.SuperTip.Items.AddSeparator();
            progressBarControl1.SuperTip.Items.Add("当前已经用时(时：分：秒)：" + ((temp_dt - pro_dt).ToString()));
            
            
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

        private void comboBox_worksheet_SelectedValueChanged(object sender, EventArgs e)
        {
            comboBox_addressField.Properties.Items.Clear();
            comboBox_cityField.Properties.Items.Clear();
            if (!isFirstLoad)
            {
                int v = comboBox_worksheet.SelectedIndex;

                string targetSheetName = comboBox_worksheet.SelectedItem.ToString() ;

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
    }

    public class GeoCodingAllInfo
    {
        public string status { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }
        public string count { get; set; }
        public Geocodes geocodes { get; set; } 
    }

    public class Geocodes
    {
        public string formatted_address { get; set; }
        public string province { get; set; }
        public string citycode { get; set; }
        public string city { get; set; }
        public string district { get; set; }
        public string township { get; set; }
        public string street { get; set; }
        public string number { get; set; }
        public string adcode { get; set; }
        public string location { get; set; }
        public string level { get; set; }
    }
}
