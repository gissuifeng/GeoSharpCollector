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
    public partial class GaodeDegeocodingAdForm : DevExpress.XtraEditors.XtraForm
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

        public GaodeDegeocodingAdForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void GaodeDegeocodingAdForm_Load(object sender, EventArgs e)
        {

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

        Worksheet worksheet_businessAreas;
        Worksheet worksheet_pois;
        Worksheet worksheet_roads;
        Worksheet worksheet_roadInters;
        Worksheet worksheet_aois;
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

                    worksheet_businessAreas = workbook.Worksheets.Add("businessAreas_sheet");
                    worksheet_businessAreas[0, 0].SetValue("uid");
                    worksheet_businessAreas[0, 1].SetValue("lng");
                    worksheet_businessAreas[0, 2].SetValue("lat");
                    worksheet_businessAreas[0, 3].SetValue("name");
                    worksheet_businessAreas[0, 4].SetValue("id");

                    worksheet_pois = workbook.Worksheets.Add("pois_sheet");
                    worksheet_pois[0, 0].SetValue("uid");
                    worksheet_pois[0, 1].SetValue("id");
                    worksheet_pois[0, 2].SetValue("name");
                    worksheet_pois[0, 3].SetValue("type");
                    worksheet_pois[0, 4].SetValue("tel");
                    worksheet_pois[0, 5].SetValue("direction");
                    worksheet_pois[0, 6].SetValue("distance");
                    worksheet_pois[0, 7].SetValue("lng");
                    worksheet_pois[0, 8].SetValue("lat");
                    worksheet_pois[0, 9].SetValue("address");
                    worksheet_pois[0, 10].SetValue("poiweight");
                    worksheet_pois[0, 11].SetValue("businessArea");

                    worksheet_roads = workbook.Worksheets.Add("roads_sheet");
                    worksheet_roads[0, 0].SetValue("uid");
                    worksheet_roads[0, 1].SetValue("id");
                    worksheet_roads[0, 2].SetValue("name");
                    worksheet_roads[0, 3].SetValue("type");
                    worksheet_roads[0, 4].SetValue("tel");
                    worksheet_roads[0, 5].SetValue("direction");
                    worksheet_roads[0, 6].SetValue("distance");
                    worksheet_roads[0, 7].SetValue("lng");
                    worksheet_roads[0, 8].SetValue("lat");
                    worksheet_roads[0, 9].SetValue("address");
                    worksheet_roads[0, 10].SetValue("poiweight");
                    worksheet_roads[0, 11].SetValue("businessArea");

                    worksheet_roadInters = workbook.Worksheets.Add("roadInters_sheet");
                    worksheet_roadInters[0, 0].SetValue("uid");
                    worksheet_roadInters[0, 1].SetValue("direction");
                    worksheet_roadInters[0, 2].SetValue("distance");
                    worksheet_roadInters[0, 3].SetValue("location");
                    worksheet_roadInters[0, 4].SetValue("firstId");
                    worksheet_roadInters[0, 5].SetValue("firstName");
                    worksheet_roadInters[0, 6].SetValue("secondeId");
                    worksheet_roadInters[0, 7].SetValue("secondeName");

                    worksheet_aois = workbook.Worksheets.Add("aois_sheet");
                    worksheet_aois[0, 0].SetValue("uid");
                    worksheet_aois[0, 1].SetValue("id");
                    worksheet_aois[0, 2].SetValue("name");
                    worksheet_aois[0, 3].SetValue("lng");
                    worksheet_aois[0, 4].SetValue("lat");
                    worksheet_aois[0, 5].SetValue("adcode");
                    worksheet_aois[0, 6].SetValue("area");
                    worksheet_aois[0, 7].SetValue("distance");
                    worksheet_aois[0, 8].SetValue("type");


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
            List<AddreessAllInfoGaodeAd> list = new List<AddreessAllInfoGaodeAd>();

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
                    AddreessAllInfoGaodeAd addressAllInfo = new AddreessAllInfoGaodeAd();
                    RegeocodeGaodeAd regeocode = new RegeocodeGaodeAd();
                    AddressComponentGaodeAd addressComponent = new AddressComponentGaodeAd();
                    NeighborhoodGaodeAd neighborhoode = new NeighborhoodGaodeAd();
                    BuildingGaodeAd building = new BuildingGaodeAd();
                    StreetNumberGaodeAd streetNumber = new StreetNumberGaodeAd();
                    BusinessAreasGaodeAd businessAreas = new BusinessAreasGaodeAd();
                    BusinessAreaGaodeAd businessArea = new BusinessAreaGaodeAd();

                    POIsInfoGaodeAd poisInfo = new POIsInfoGaodeAd();
                    RoadsInfoGaodeAd roadsInfo = new RoadsInfoGaodeAd();
                    RoadIntersGaodeAd roadintersInfo = new RoadIntersGaodeAd();
                    AoisGaodeAd aoisInfo = new AoisGaodeAd();

                    addressAllInfo.uid = idName;

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

                        businessArea = new BusinessAreaGaodeAd();
                        businessArea.location = obj1["location"].ToString();
                        businessArea.name = obj1["name"].ToString();
                        businessArea.id = obj1["id"].ToString();

                        businessAreas.businessAreas.Add(businessArea);
                    }

                    addressComponent.neighborhood = neighborhoode;
                    addressComponent.building = building;
                    addressComponent.streetNumber = streetNumber;
                    addressComponent.businessAreas = businessAreas;

                    JArray jlist_pois = JArray.Parse(obj["regeocode"]["pois"].ToString());
                    JObject obj_poi = null;
                    PoiInfoGaodeAd poiInfo = null;
                    for (int m = 0; m < jlist_pois.Count; m++)
                    {
                        obj_poi = JObject.Parse(jlist_pois[m].ToString());

                        poiInfo = new PoiInfoGaodeAd();
                        poiInfo.id = obj_poi["id"].ToString();
                        poiInfo.name = obj_poi["name"].ToString();
                        poiInfo.type = obj_poi["type"].ToString();
                        poiInfo.tel = obj_poi["tel"].ToString();
                        poiInfo.direction = obj_poi["direction"].ToString();
                        poiInfo.distance = obj_poi["distance"].ToString();
                        poiInfo.location = obj_poi["location"].ToString();
                        poiInfo.address = obj_poi["address"].ToString();
                        poiInfo.poiWeight = obj_poi["poiweight"].ToString();
                        poiInfo.businessArea = obj_poi["businessarea"].ToString();

                        poisInfo.pois.Add(poiInfo);
                    }

                    JArray jlist_roads = JArray.Parse(obj["regeocode"]["roads"].ToString());
                    JObject obj_road = null;
                    RoadInfoGaodeAd roadInfo = null;
                    for (int n = 0; n < jlist_roads.Count; n++)
                    {
                        obj_road = JObject.Parse(jlist_roads[n].ToString());

                        roadInfo = new RoadInfoGaodeAd();
                        roadInfo.id = obj_road["id"].ToString();
                        roadInfo.name = obj_road["name"].ToString();
                        roadInfo.direction = obj_road["direction"].ToString();
                        roadInfo.distance = obj_road["distance"].ToString();
                        roadInfo.location = obj_road["location"].ToString();
                       
                        roadsInfo.roads.Add(roadInfo);
                    }

                    JArray jlist_roadinters = JArray.Parse(obj["regeocode"]["roadinters"].ToString());
                    JObject obj_roadinter = null;
                    RoadInterGaodeAd roadinter = null;
                    for (int u = 0; u < jlist_roadinters.Count; u++)
                    {
                        obj_roadinter = JObject.Parse(jlist_roadinters[u].ToString());

                        roadinter = new RoadInterGaodeAd();
                        roadinter.direction = obj_roadinter["direction"].ToString();
                        roadinter.distance = obj_roadinter["distance"].ToString();
                        roadinter.location = obj_roadinter["location"].ToString();
                        roadinter.first_id = obj_roadinter["first_id"].ToString();
                        roadinter.first_name = obj_roadinter["first_name"].ToString();
                        roadinter.second_id = obj_roadinter["second_id"].ToString();
                        roadinter.second_name = obj_roadinter["second_name"].ToString();

                        roadintersInfo.roadInters.Add(roadinter);
                    }

                    JArray jlist_aois = JArray.Parse(obj["regeocode"]["aois"].ToString());
                    JObject obj_aoi = null;
                    AoiGaodeAd aoi = null;
                    for (int v = 0; v < jlist_aois.Count; v++)
                    {
                        obj_aoi = JObject.Parse(jlist_aois[v].ToString());

                        aoi = new AoiGaodeAd();
                        aoi.id = obj_aoi["id"].ToString();
                        aoi.name = obj_aoi["name"].ToString();
                        aoi.adcode = obj_aoi["adcode"].ToString();
                        aoi.location = obj_aoi["location"].ToString();
                        aoi.area = obj_aoi["area"].ToString();
                        aoi.distance = obj_aoi["distance"].ToString();
                        aoi.type = obj_aoi["type"].ToString();

                        aoisInfo.aois.Add(aoi);
                    }

                    regeocode.addressComponent = addressComponent;
                    addressComponent.businessAreas = businessAreas;

                    addressAllInfo.regeocode = regeocode;

                    regeocode.pois = poisInfo;
                    regeocode.roads = roadsInfo;
                    regeocode.roadInters = roadintersInfo;
                    regeocode.aois = aoisInfo;

                    addressAllInfo.regeocode = regeocode;

                    list.Add(addressAllInfo);
                }
                sum++;
                RunWithInoke(sum);
            }

            int business_order = 0;
            int poi_order = 0;
            int road_order = 0;
            int roadinter_order = 0;
            int aoi_order = 0;
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


                for (int business_i = 0; business_i < list[i].regeocode.addressComponent.businessAreas.businessAreas.Count; business_i++)
                {
                    worksheet_businessAreas[business_order + 1, 0].SetValue(list[i].uid);
                    worksheet_businessAreas[business_order + 1, 1].SetValue(list[i].regeocode.addressComponent.businessAreas.businessAreas[business_i].location.Split(',')[0]);
                    worksheet_businessAreas[business_order + 1, 2].SetValue(list[i].regeocode.addressComponent.businessAreas.businessAreas[business_i].location.Split(',')[1]);
                    worksheet_businessAreas[business_order + 1, 3].SetValue(list[i].regeocode.addressComponent.businessAreas.businessAreas[business_i].name);
                    worksheet_businessAreas[business_order + 1, 4].SetValue(list[i].regeocode.addressComponent.businessAreas.businessAreas[business_i].id);

                    business_order++;
                }

                for (int poi_i = 0; poi_i < list[i].regeocode.pois.pois.Count; poi_i++)
                {
                    worksheet_pois[poi_order + 1, 0].SetValue(list[i].uid);
                    worksheet_pois[poi_order + 1, 1].SetValue(list[i].regeocode.pois.pois[poi_i].id);
                    worksheet_pois[poi_order + 1, 2].SetValue(list[i].regeocode.pois.pois[poi_i].name);
                    worksheet_pois[poi_order + 1, 3].SetValue(list[i].regeocode.pois.pois[poi_i].type);
                    worksheet_pois[poi_order + 1, 4].SetValue(list[i].regeocode.pois.pois[poi_i].tel);
                    worksheet_pois[poi_order + 1, 5].SetValue(list[i].regeocode.pois.pois[poi_i].direction);
                    worksheet_pois[poi_order + 1, 6].SetValue(list[i].regeocode.pois.pois[poi_i].distance);
                    worksheet_pois[poi_order + 1, 7].SetValue(list[i].regeocode.pois.pois[poi_i].location.Split(',')[0]);
                    worksheet_pois[poi_order + 1, 8].SetValue(list[i].regeocode.pois.pois[poi_i].location.Split(',')[1]);
                    worksheet_pois[poi_order + 1, 9].SetValue(list[i].regeocode.pois.pois[poi_i].address);
                    worksheet_pois[poi_order + 1, 10].SetValue(list[i].regeocode.pois.pois[poi_i].poiWeight);
                    worksheet_pois[poi_order + 1, 11].SetValue(list[i].regeocode.pois.pois[poi_i].businessArea);

                    poi_order++;
                }
                for (int road_i = 0; road_i < list[i].regeocode.roads.roads.Count; road_i++)
                {
                    worksheet_roads[road_order + 1, 0].SetValue(list[i].uid);
                    worksheet_roads[road_order + 1, 1].SetValue(list[i].regeocode.roads.roads[road_i].id);
                    worksheet_roads[road_order + 1, 2].SetValue(list[i].regeocode.roads.roads[road_i].name);
                    worksheet_roads[road_order + 1, 3].SetValue(list[i].regeocode.roads.roads[road_i].direction);
                    worksheet_roads[road_order + 1, 4].SetValue(list[i].regeocode.roads.roads[road_i].distance);
                    worksheet_roads[road_order + 1, 5].SetValue(list[i].regeocode.roads.roads[road_i].location.Split(',')[0]);
                    worksheet_roads[road_order + 1, 6].SetValue(list[i].regeocode.roads.roads[road_i].location.Split(',')[1]);

                    road_order++;
                }
                for (int roadinter_i = 0; roadinter_i < list[i].regeocode.roadInters.roadInters.Count; roadinter_i++)
                {
                    worksheet_roadInters[roadinter_order + 1, 0].SetValue(list[i].uid);
                    worksheet_roadInters[roadinter_order + 1, 1].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].direction);
                    worksheet_roadInters[roadinter_order + 1, 2].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].distance);
                    worksheet_roadInters[roadinter_order + 1, 3].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].location.Split(',')[0]);
                    worksheet_roadInters[roadinter_order + 1, 4].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].location.Split(',')[1]);
                    worksheet_roadInters[roadinter_order + 1, 5].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].first_id);
                    worksheet_roadInters[roadinter_order + 1, 6].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].first_name);
                    worksheet_roadInters[roadinter_order + 1, 7].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].second_id);
                    worksheet_roadInters[roadinter_order + 1, 8].SetValue(list[i].regeocode.roadInters.roadInters[roadinter_i].second_name);

                    roadinter_order++;
                }
                for (int aoi_i = 0; aoi_i < list[i].regeocode.aois.aois.Count; aoi_i++)
                {
                    worksheet_aois[aoi_order + 1, 0].SetValue(list[i].uid);
                    worksheet_aois[aoi_order + 1, 1].SetValue(list[i].regeocode.aois.aois[aoi_i].id);
                    worksheet_aois[aoi_order + 1, 2].SetValue(list[i].regeocode.aois.aois[aoi_i].name);
                    worksheet_aois[aoi_order + 1, 3].SetValue(list[i].regeocode.aois.aois[aoi_i].adcode);
                    worksheet_aois[aoi_order + 1, 4].SetValue(list[i].regeocode.aois.aois[aoi_i].location.Split(',')[0]);
                    worksheet_aois[aoi_order + 1, 5].SetValue(list[i].regeocode.aois.aois[aoi_i].location.Split(',')[1]);
                    worksheet_aois[aoi_order + 1, 6].SetValue(list[i].regeocode.aois.aois[aoi_i].area);
                    worksheet_aois[aoi_order + 1, 7].SetValue(list[i].regeocode.aois.aois[aoi_i].distance);
                    worksheet_aois[aoi_order + 1, 8].SetValue(list[i].regeocode.aois.aois[aoi_i].type);

                    aoi_order++;

                }
            }

           

            taskExecuted = true;
            workbook.SaveDocument(docPath);

            XtraMessageBox.Show("所有数据的地理编码任务已经完成！");
        }

        string gaodeKey = "e00536b393e9671af12bea182f75a36b";
        private string SetAddress(string strLnglat, string strKey)
        {
            return string.Format(@"http://restapi.amap.com/v3/geocode/regeo?key={1}&location={0}&extensions=all", strLnglat, strKey);
        }

        private void RunWithInoke(int i)
        {

            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);

        }

        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;
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

    public class AddreessAllInfoGaodeAd
    {
        public string status { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }

        public RegeocodeGaodeAd regeocode { get; set; }

        public string uid { get; set; }

    }

    public class RegeocodeGaodeAd
    {
        public string formatted_address { get; set; }

        public AddressComponentGaodeAd addressComponent { get; set; }

        public POIsInfoGaodeAd pois { get; set; }

        public RoadsInfoGaodeAd roads { get; set; }

        public RoadIntersGaodeAd roadInters { get; set; }

        public AoisGaodeAd aois { get; set; }
    }

    public class AddressComponentGaodeAd
    {
        public string country { get; set; }
        public string province { get; set; }
        public string city { get; set; }
        public string citycode { get; set; }
        public string district { get; set; }
        public string adcode { get; set; }
        public string township { get; set; }
        public string towncode { get; set; }

        public NeighborhoodGaodeAd neighborhood { get; set; }

        public BuildingGaodeAd building { get; set; }

        public StreetNumberGaodeAd streetNumber { get; set; }

        public BusinessAreasGaodeAd businessAreas { get; set; }

    }

    public class NeighborhoodGaodeAd
    {
        public string name { get; set; }
        public string type { get; set; }
    }

    public class BuildingGaodeAd
    {
        public string name { get; set; }
        public string type { get; set; }
    }

    public class StreetNumberGaodeAd
    {
        public string street { get; set; }
        public string number { get; set; }
        public string location { get; set; }
        public string direction { get; set; }
        public string distance { get; set; }
    }

    public class BusinessAreasGaodeAd
    {
        public List<BusinessAreaGaodeAd> businessAreas = new List<BusinessAreaGaodeAd>();
    }

    public class BusinessAreaGaodeAd
    {
        public string location { get; set; }
        public string name { get; set; }
        public string id { get; set; }
    }

    public class POIsInfoGaodeAd
    {
        public List<PoiInfoGaodeAd> pois = new List<PoiInfoGaodeAd>();
    }

    public class PoiInfoGaodeAd
    {
        public string id { get; set; }
        public string name { get; set; }
        public string type { get; set; }
        public string tel { get; set; }
        public string direction { get; set; }
        public string distance { get; set; }
        public string location { get; set; }
        public string address { get; set; }
        public string poiWeight { get; set; }
        public string businessArea { get; set; }
    }

    public class RoadsInfoGaodeAd
    {
        public List<RoadInfoGaodeAd> roads = new List<RoadInfoGaodeAd>();
    }

    public class RoadInfoGaodeAd
    {
        public string id { get; set; }
        public string name { get; set; }
        public string direction { get; set; }
        public string distance { get; set; }
        public string location { get; set; }
    }

    public class RoadIntersGaodeAd
    {
        public List<RoadInterGaodeAd> roadInters = new List<RoadInterGaodeAd>();
    }

    public class RoadInterGaodeAd
    {
        public string direction { get; set; }
        public string distance { get; set; }
        public string location { get; set; }
        public string first_id { get; set; }
        public string first_name { get; set; }
        public string second_id { get; set; }
        public string second_name { get; set; }
    }

    public class AoisGaodeAd
    {
        public List<AoiGaodeAd> aois = new List<AoiGaodeAd>();
    }

    public class AoiGaodeAd
    {
        public string id { get; set; }
        public string name { get; set; }
        public string adcode { get; set; }
        public string location { get; set; }
        public string area { get; set; }
        public string distance { get; set; }
        public string type { get; set; }
    }
}
