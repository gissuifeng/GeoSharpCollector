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
    public partial class GaodeDegeocodingForm : DevExpress.XtraEditors.XtraForm
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

        public GaodeDegeocodingForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void btn_open_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;
            comboBox_lngField.Properties.Items.Clear();
            comboBox_latField.Properties.Items.Clear();
            comboBox_idField.Properties.Items.Clear();
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
                    comboBox_lngField.Properties.Items.Add(worksheet[0, j].Value);
                    comboBox_latField.Properties.Items.Add(worksheet[0, j].Value);
                    comboBox_idField.Properties.Items.Add(worksheet[0, j].Value);
                }

                comboBox_lngField.SelectedIndex = 0;
                comboBox_latField.SelectedIndex = 0;
                comboBox_idField.SelectedIndex = 0;

                isFirstLoad = false;
            }
        }

        private void comboBox_worksheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox_lngField.Properties.Items.Clear();
            comboBox_latField.Properties.Items.Clear();
            comboBox_idField.Properties.Items.Clear();

            if (!isFirstLoad)
            {
                string targetSheetName = comboBox_worksheet.SelectedText;

                if (targetSheetName != "")
                {
                    worksheet = workbook.Worksheets[targetSheetName];

                    for (int j = 0; j < worksheet.GetDataRange().ColumnCount; j++)
                    {
                        comboBox_lngField.Properties.Items.Add(worksheet[0, j].Value);
                        comboBox_latField.Properties.Items.Add(worksheet[0, j].Value);
                        comboBox_idField.Properties.Items.Add(worksheet[0, j].Value);
                    }

                    comboBox_lngField.SelectedIndex = 0;
                    comboBox_latField.SelectedIndex = 0;
                    comboBox_idField.SelectedIndex = 0;
                }
            }
        }

        string[] strs_lng;
        string[] strs_lat;
        string[] strs_id;

        int colCount;
        int rowCount;

        int col_lng;
        int col_lat;
        int col_id;
        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                try
                {
                    col_lng = comboBox_lngField.SelectedIndex;
                    col_lat = comboBox_latField.SelectedIndex;
                    col_id = comboBox_idField.SelectedIndex;

                    colCount = worksheet.GetDataRange().ColumnCount;
                    rowCount = worksheet.GetDataRange().RowCount;

                    strs_lng = new string[rowCount];
                    strs_lat = new string[rowCount];
                    strs_id = new string[rowCount];

                    progressBarControl1.Properties.Maximum = rowCount;

                    for (int i = 1; i < rowCount; i++)
                    {
                        strs_lng[i] = worksheet[i, col_lng].Value.ToString();
                        strs_lat[i] = worksheet[i, col_lat].Value.ToString();
                        strs_id[i] = worksheet[i, col_id].Value.ToString();
                    }

                    worksheet[0, colCount].SetValue("status");
                    worksheet[0, colCount + 1].SetValue("info");
                    worksheet[0, colCount + 2].SetValue("infocode");

                    worksheet[0, colCount + 3].SetValue("formatted_address");

                    worksheet[0, colCount + 4].SetValue("country");
                    worksheet[0, colCount + 5].SetValue("province");
                    worksheet[0, colCount + 6].SetValue("city");
                    worksheet[0, colCount + 7].SetValue("citycode");
                    worksheet[0, colCount + 8].SetValue("district");
                    worksheet[0, colCount + 9].SetValue("adcode");
                    worksheet[0, colCount + 10].SetValue("township");
                    worksheet[0, colCount + 11].SetValue("towncode");

                    worksheet[0, colCount + 12].SetValue("neighborhood_name");
                    worksheet[0, colCount + 13].SetValue("neighborhood_type");

                    worksheet[0, colCount + 14].SetValue("building_name");
                    worksheet[0, colCount + 15].SetValue("building_type");

                    worksheet[0, colCount + 16].SetValue("street_name");
                    worksheet[0, colCount + 17].SetValue("street_number");
                    worksheet[0, colCount + 18].SetValue("street_lng");
                    worksheet[0, colCount + 19].SetValue("street_lat");
                    worksheet[0, colCount + 20].SetValue("street_dir");
                    worksheet[0, colCount + 21].SetValue("street_dis");

                    thread = new Thread(new ThreadStart(fun));
                    thread.Start();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
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
            List<AddreessAllInfoGaode> list = new List<AddreessAllInfoGaode>();

            string lngName;
            string latName;
            string idName;

            string str_lnglat;

            for (int i = 1; i < rowCount; i++)
            {
                lngName = worksheet[i, comboBox_lngField.SelectedIndex].Value.ToString();
                latName = worksheet[i, comboBox_latField.SelectedIndex].Value.ToString();
                idName = worksheet[i, comboBox_idField.SelectedIndex].Value.ToString();

                str_lnglat = string.Format($"{lngName},{latName}");

                try
                {
                    stream = client.OpenRead(SetAddress(str_lnglat, gaodeKey));

                    str_json = new StreamReader(stream).ReadToEnd();
                }
                catch
                {

                }

                if (str_json != "")
                {
                    AddreessAllInfoGaode addressAllInfo = new AddreessAllInfoGaode();
                    RegeocodeGaode regeocode = new RegeocodeGaode();
                    AddressComponentGaode addressComponent = new AddressComponentGaode();
                    NeighborhoodGaode neighborhoode = new NeighborhoodGaode();
                    BuildingGaode building = new BuildingGaode();
                    StreetNumberGaode streetNumber = new StreetNumberGaode();
                    BusinessAreasGaode businessAreas = new BusinessAreasGaode();
                    BusinessAreaGaode businessArea = new BusinessAreaGaode();

                    JObject obj = JObject.Parse(str_json);

                    addressAllInfo.status = obj["status"].ToString();
                    addressAllInfo.info = obj["info"].ToString();
                    addressAllInfo.infocode = obj["infocode"].ToString();

                    regeocode.formatted_address = obj["regeocode"]["formatted_address"].ToString();

                    addressComponent.country = obj["regeocode"]["addressComponent"]["country"].ToString();
                    addressComponent.province = obj["regeocode"]["addressComponent"]["province"].ToString();
                    addressComponent.city = obj["regeocode"]["addressComponent"]["city"].ToString();
                    addressComponent.citycode = obj["regeocode"]["addressComponent"]["citycode"].ToString();
                    addressComponent.district = obj["regeocode"]["addressComponent"]["district"].ToString();
                    addressComponent.adcode = obj["regeocode"]["addressComponent"]["adcode"].ToString();
                    addressComponent.towncode = obj["regeocode"]["addressComponent"]["towncode"].ToString();
                    addressComponent.township = obj["regeocode"]["addressComponent"]["township"].ToString();
                    addressComponent.towncode = obj["regeocode"]["addressComponent"]["towncode"].ToString();

                    neighborhoode.name = obj["regeocode"]["addressComponent"]["neighborhood"]["name"].ToString();
                    neighborhoode.type = obj["regeocode"]["addressComponent"]["neighborhood"]["type"].ToString();

                    building.name = obj["regeocode"]["addressComponent"]["building"]["name"].ToString();
                    building.type = obj["regeocode"]["addressComponent"]["building"]["type"].ToString();

                    streetNumber.street = obj["regeocode"]["addressComponent"]["streetNumber"]["street"].ToString();
                    streetNumber.number = obj["regeocode"]["addressComponent"]["streetNumber"]["number"].ToString();
                    streetNumber.location = obj["regeocode"]["addressComponent"]["streetNumber"]["location"].ToString();
                    streetNumber.direction = obj["regeocode"]["addressComponent"]["streetNumber"]["direction"].ToString();
                    streetNumber.distance = obj["regeocode"]["addressComponent"]["streetNumber"]["distance"].ToString();

                    streetNumber.distance = obj["regeocode"]["addressComponent"]["streetNumber"]["distance"].ToString();

                    JArray jlist = JArray.Parse(obj["regeocode"]["addressComponent"]["businessAreas"].ToString());

                    JObject obj1 = null;
                    for (int j = 0; j < jlist.Count; j++)
                    {
                        obj1 = JObject.Parse(jlist[j].ToString());

                        businessArea = new BusinessAreaGaode();
                        businessArea.location = obj1["location"].ToString();
                        businessArea.name = obj1["name"].ToString();
                        businessArea.id = obj1["id"].ToString();

                        businessAreas.businessAreas.Add(businessArea);
                    }

                    addressComponent.neighborhood = neighborhoode;
                    addressComponent.building = building;
                    addressComponent.streetNumber = streetNumber;
                    addressComponent.businessAreas = businessAreas;

                    regeocode.addressComponent = addressComponent;

                    addressAllInfo.regeocode = regeocode;

                    list.Add(addressAllInfo);
                }
                sum++;
                RunWithInoke(sum);
            }

            for (int i = 0; i < list.Count; i++)
            {
                //MessageBox.Show(list[i].geocodes.location.ToString());

                worksheet[i + 1, colCount].SetValue(list[i].status);
                worksheet[i + 1, colCount + 1].SetValue(list[i].info);
                worksheet[i + 1, colCount + 2].SetValue(list[i].infocode);

                worksheet[i + 1, colCount + 3].SetValue(list[i].regeocode.formatted_address);
                worksheet[i + 1, colCount + 4].SetValue(list[i].regeocode.addressComponent.country);
                worksheet[i + 1, colCount + 5].SetValue(list[i].regeocode.addressComponent.province);
                worksheet[i + 1, colCount + 6].SetValue(list[i].regeocode.addressComponent.city);
                worksheet[i + 1, colCount + 7].SetValue(list[i].regeocode.addressComponent.citycode);
                worksheet[i + 1, colCount + 8].SetValue(list[i].regeocode.addressComponent.district);
                worksheet[i + 1, colCount + 9].SetValue(list[i].regeocode.addressComponent.adcode);
                worksheet[i + 1, colCount + 10].SetValue(list[i].regeocode.addressComponent.township);
                worksheet[i + 1, colCount + 11].SetValue(list[i].regeocode.addressComponent.towncode);

                worksheet[i + 1, colCount + 12].SetValue(list[i].regeocode.addressComponent.neighborhood.name);
                worksheet[i + 1, colCount + 13].SetValue(list[i].regeocode.addressComponent.neighborhood.type);

                worksheet[i + 1, colCount + 14].SetValue(list[i].regeocode.addressComponent.building.name);
                worksheet[i + 1, colCount + 15].SetValue(list[i].regeocode.addressComponent.building.type);

                worksheet[i + 1, colCount + 16].SetValue(list[i].regeocode.addressComponent.streetNumber.street);
                worksheet[i + 1, colCount + 17].SetValue(list[i].regeocode.addressComponent.streetNumber.number);

                worksheet[i + 1, colCount + 18].SetValue(list[i].regeocode.addressComponent.streetNumber.location.Split(',')[0]);
                worksheet[i + 1, colCount + 19].SetValue(list[i].regeocode.addressComponent.streetNumber.location.Split(',')[1]);

                worksheet[i + 1, colCount + 20].SetValue(list[i].regeocode.addressComponent.streetNumber.direction);
                worksheet[i + 1, colCount + 21].SetValue(list[i].regeocode.addressComponent.streetNumber.distance);

            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);

            XtraMessageBox.Show("所有数据的地理编码任务已经完成！");
        }

        string gaodeKey = "e00536b393e9671af12bea182f75a36b";
        private string SetAddress(string strLnglat,string strKey)
        {
            return string.Format(@"http://restapi.amap.com/v3/geocode/regeo?key={1}&location={0}", strLnglat, strKey);
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

        private void GaodeDegeocodingForm_Load(object sender, EventArgs e)
        {

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

    public class AddreessAllInfoGaode
    {
        public string status { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }

        public RegeocodeGaode regeocode { get; set; }

    }

    public class RegeocodeGaode
    {
        public string formatted_address { get; set; }

        public AddressComponentGaode addressComponent { get; set; }
    }

    public class AddressComponentGaode
    {
        public string country { get; set; }
        public string province { get; set; }
        public string city { get; set; }
        public string citycode { get; set; }
        public string district { get; set; }
        public string adcode { get; set; }
        public string township { get; set; }
        public string towncode { get; set; }

        public NeighborhoodGaode neighborhood { get; set; }

        public BuildingGaode building { get; set; }

        public StreetNumberGaode streetNumber { get; set; }

        public BusinessAreasGaode businessAreas { get; set; }

    }

    public class NeighborhoodGaode
    {
        public string name { get; set; }
        public string type { get; set; }
    }

    public class BuildingGaode
    {
        public string name { get; set; }
        public string type { get; set; }
    }

    public class StreetNumberGaode
    {
        public string street { get; set; }
        public string number { get; set; }
        public string location { get; set; }
        public string direction { get; set; }
        public string distance { get; set; }
    }

    public class BusinessAreasGaode
    {
        public List<BusinessAreaGaode> businessAreas = new List<BusinessAreaGaode>();
    }

    public class BusinessAreaGaode
    {
        public string location { get; set; }
        public string name { get; set; }
        public string id { get; set; }
    }
}
