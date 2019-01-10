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
    public partial class GaodePOIGetForm : DevExpress.XtraEditors.XtraForm
    {
        //定义delegate以便Invoke时使用  
        private delegate void SetProgressBarValue(int value);
        /// <summary>
        /// 主窗口对象
        /// </summary>
        private MainForm mainForm;
        /// <summary>
        /// 工作表控件
        /// </summary>
        private SpreadsheetControl spreadsheetControl;
        /// <summary>
        /// 表格文件路径
        /// </summary>
        private string docPath;
        /// <summary>
        /// 工作簿
        /// </summary>
        private Workbook workbook;
        /// <summary>
        /// 工作表
        /// </summary>
        private Worksheet worksheet;
        /// <summary>
        /// 数据采集线程
        /// </summary>
        private Thread thread;

        private bool taskExecuted = false;

        private string typeKeyword;

        private string codeType;

        public GaodePOIGetForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void GaodePOIGetForm_Load(object sender, EventArgs e)
        {
            textEdit_topleftX.EditValue = "118.74";
            textEdit_topleftY.EditValue = "32.08";

            textEdit_bottomRightX.EditValue = "118.86";
            textEdit_bottomRightY.EditValue = "32.00";
        }

        private void btn_open_Click(object sender, EventArgs e)
        {
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
                workbook.LoadDocument(textEdit_file.EditValue.ToString());

                worksheet = workbook.Worksheets[0];
            }
        }


        public string poiTypeCode = "";
        public string poiKeyword = "";
        private void btn_poiType_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            GaodePOITypeForm dlg = new GaodePOITypeForm(this);

            DialogResult rs = dlg.ShowDialog();

            if(rs == DialogResult.OK)
            {
                btn_poiType.Text = poiKeyword;
                tb_poiCode.EditValue = poiTypeCode;
            }
        }

        private void btn_addDefaultExt_Click(object sender, EventArgs e)
        {
            textEdit_topleftX.EditValue = mainForm.extentCoord.Lx;
            textEdit_topleftY.EditValue = mainForm.extentCoord.Ly;

            textEdit_bottomRightX.EditValue = mainForm.extentCoord.Rx;
            textEdit_bottomRightY.EditValue = mainForm.extentCoord.Ry;
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            if (tb_poiCode.Text == "POI类型编码未知")
            {
                XtraMessageBox.Show("请选择POI关键词，不要手动填写");
                return;
            }

            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                try
                {
                    btn_ok.Enabled = false;

                    typeKeyword = btn_poiType.EditValue.ToString();
                    codeType = tb_poiCode.Text;

                    worksheet[0, 0].SetValue("status");
                    worksheet[0, 1].SetValue("count");
                    worksheet[0, 2].SetValue("info");
                    worksheet[0, 3].SetValue("infocode");

                    worksheet[0, 4].SetValue("keywords");
                    worksheet[0, 5].SetValue("cities");

                    worksheet[0, 6].SetValue("id");
                    worksheet[0, 7].SetValue("name");
                    worksheet[0, 8].SetValue("type");
                    worksheet[0, 9].SetValue("typecode");
                    worksheet[0, 10].SetValue("biz_type");
                    worksheet[0, 11].SetValue("address");
                    worksheet[0, 12].SetValue("lng");
                    worksheet[0, 13].SetValue("lat");
                    worksheet[0, 14].SetValue("tel");
                    worksheet[0, 15].SetValue("distance");
                    worksheet[0, 16].SetValue("biz_ext");
                    worksheet[0, 17].SetValue("pname");
                    worksheet[0, 18].SetValue("cityname");
                    worksheet[0, 19].SetValue("adname");

                    double ox = Convert.ToDouble(textEdit_topleftX.EditValue);
                    double oy = Convert.ToDouble(textEdit_topleftY.EditValue);

                    double dx = Convert.ToDouble(textEdit_bottomRightX.EditValue);
                    double dy = Convert.ToDouble(textEdit_bottomRightY.EditValue);

                    double len = 0.005;

                    int rowCount = GetGridRowCount(oy - dy, len);
                    int colCount = GetGridColCount(dx - ox, len);

                    progressBarControl1.Properties.Maximum = rowCount * colCount;

                    thread = new Thread(new ThreadStart(fun1));
                    thread.Start();
                }
                catch
                {

                }
            }
        }

        public void fun1()
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            List<POIAllInfoGaode> list = new List<POIAllInfoGaode>();

            double ox = Convert.ToDouble(textEdit_topleftX.EditValue);
            double oy = Convert.ToDouble(textEdit_topleftY.EditValue);

            double dx = Convert.ToDouble(textEdit_bottomRightX.EditValue);
            double dy = Convert.ToDouble(textEdit_bottomRightY.EditValue);

            double len = 0.005;

            int rowCount = GetGridRowCount(oy - dy, len);
            int colCount = GetGridColCount(dx - ox, len);

            int rowCur = 1;
            int sum = 1;

            double oy_val;
            string str_json = "";
            POIAllInfoGaode poiAllInfo = null;
            for (int row = 1; row <= rowCount; row++)
            {
                oy_val = oy - (row - 1) * 0.005;

                for (int i = 1; i <= colCount; i++)
                {

                    try
                    {
                        stream = client.OpenRead(SetAddress(GetPolygonString(ox, oy_val, len, i), typeKeyword, 1, codeType));
                       
                        str_json = new StreamReader(stream).ReadToEnd();
                    }
                    catch (Exception ex)
                    {

                      
                    }

                    if (str_json != "")
                    {
                        poiAllInfo = new POIAllInfoGaode();
                        SuggestionGaode suggestion = new SuggestionGaode();
                        List<PoiGaode> pois = new List<PoiGaode>();
                        
                        JObject obj = JObject.Parse(str_json);

                        poiAllInfo.status = obj["status"].ToString();
                        poiAllInfo.count = obj["count"].ToString();
                        poiAllInfo.info = obj["info"].ToString();
                        poiAllInfo.infocode = obj["infocode"].ToString();

                        suggestion.keywords = obj["suggestion"]["keywords"].ToString();
                        suggestion.cities = obj["suggestion"]["cities"].ToString();

                        poiAllInfo.sugeestion = suggestion;

                        JArray jlist = JArray.Parse(obj["pois"].ToString());
                        PoiGaode poi = null;
                        JObject obj1 = null;
                        for (int j = 0; j < jlist.Count; j++)
                        {
                            obj1 = JObject.Parse(jlist[j].ToString());

                            poi = new PoiGaode();

                            poi.id = obj1["id"].ToString();
                            poi.name = obj1["name"].ToString();
                            poi.type = obj1["type"].ToString();
                            poi.typecode = obj1["typecode"].ToString();
                            poi.biz_type = obj1["biz_type"].ToString();
                            poi.address = obj1["address"].ToString();
                            poi.location = obj1["location"].ToString();
                            poi.tel = obj1["tel"].ToString();
                            poi.distance = obj1["distance"].ToString();
                            poi.biz_ext = obj1["biz_ext"].ToString();
                            poi.pname = obj1["pname"].ToString();
                            poi.cityname = obj1["cityname"].ToString();
                            poi.adname = obj1["adname"].ToString();

                            pois.Add(poi);
                        }

                        poiAllInfo.pois = pois;

                        Console.WriteLine("poi_count" + pois.Count);
                    }

                    list.Add(poiAllInfo);
                    sum++;
                    RunWithInoke(sum);
                }

            }

            int rowOrder = 1;
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < list[i].pois.Count; j++)
                {

                    worksheet[rowOrder, 0].SetValue(list[i].status);
                    worksheet[rowOrder, 1].SetValue(list[i].count);
                    worksheet[rowOrder, 2].SetValue(list[i].info);
                    worksheet[rowOrder, 3].SetValue(list[i].infocode);

                    worksheet[rowOrder, 4].SetValue(list[i].sugeestion.keywords);
                    worksheet[rowOrder, 5].SetValue(list[i].sugeestion.cities);

                    worksheet[rowOrder, 6].SetValue(list[i].pois[j].id);
                    worksheet[rowOrder, 7].SetValue(list[i].pois[j].name);
                    worksheet[rowOrder, 8].SetValue(list[i].pois[j].type);
                    worksheet[rowOrder, 9].SetValue(list[i].pois[j].typecode);
                    worksheet[rowOrder, 10].SetValue(list[i].pois[j].biz_type);
                    worksheet[rowOrder, 11].SetValue(list[i].pois[j].address);
                    worksheet[rowOrder, 12].SetValue(list[i].pois[j].location.Split(',')[0]);
                    worksheet[rowOrder, 13].SetValue(list[i].pois[j].location.Split(',')[1]);
                    worksheet[rowOrder, 14].SetValue(list[i].pois[j].tel);
                    worksheet[rowOrder, 15].SetValue(list[i].pois[j].distance);
                    worksheet[rowOrder, 16].SetValue(list[i].pois[j].biz_ext);
                    worksheet[rowOrder, 17].SetValue(list[i].pois[j].pname);
                    worksheet[rowOrder, 18].SetValue(list[i].pois[j].cityname);
                    worksheet[rowOrder, 19].SetValue(list[i].pois[j].adname);

                    rowOrder++;
                }

            }

            XtraMessageBox.Show("解析完成！");

            taskExecuted = true;

            workbook.SaveDocument(docPath);
        }
        private string GetPolygonString(double ox, double oy, double len, int i)
        {
            string polygonStr = "";

            double fx = from_pointX(ox, len, i);
            double fy = from_pointY(oy, len, i);

            double tx = to_pointX(ox, len, i);
            double ty = to_pointY(oy, len, i);

            polygonStr = string.Format("{0},{1};{2},{3}", fx, fy, tx, ty);

            return polygonStr;
        }

        private void RunWithInoke(int i)
        {

            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);

        }

        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;

            if (Convert.ToInt32(progressBarControl1.EditValue) == 5)
            {
                progressBarControl1.EditValue = 0;
            }
        }

        string strKey = "e00536b393e9671af12bea182f75a36b";

        private string SetAddress(string polygonStr, string keyWords, int page, string typecode)
        {
            string s = string.Format(@"http://restapi.amap.com/v3/place/polygon?polygon={0}&keywords={1}&output=json&key={4}&offset=50&page={2}&extensions=base&types={3}", polygonStr, keyWords, page, typecode, strKey);
            Console.WriteLine(s);
            return s;
        }

        private int GetGridColCount(double x, double len)
        {
            int colCount = (int)Math.Ceiling((double)x / len);

            return colCount;
        }

        private int GetGridRowCount(double y, double len)
        {
            int colCount = (int)Math.Ceiling((double)y / len);

            return colCount;
        }

        /// <summary>
        /// 矩形框的左上角x坐标
        /// </summary>
        /// <param name="ox">原点x坐标</param>
        /// <param name="len">矩形边长</param>
        /// <param name="i">矩形序号</param>
        /// <returns></returns>
        private double from_pointX(double ox, double len, int i)
        {
            double from_x = ox + (i - 1) * len;

            return from_x;
        }

        /// <summary>
        /// 矩形框的左上角y坐标
        /// </summary>
        /// <param name="oy">原点y坐标</param>
        /// <param name="len">矩形边长</param>
        /// <param name="i">矩形序号</param>
        /// <returns></returns>
        private double from_pointY(double oy, double len, int i)
        {
            double from_y = oy;

            return from_y;
        }

        /// <summary>
        /// 矩形框的右下角x坐标
        /// </summary>
        /// <param name="ox">原点x坐标</param>
        /// <param name="len">矩形边长</param>
        /// <param name="i">矩形序号</param>
        /// <returns></returns>
        private double to_pointX(double ox, double len, int i)
        {
            double to_x = ox + i * len;

            return to_x;
        }

        /// <summary>
        /// 矩形框的右下角y坐标
        /// </summary>
        /// <param name="oy">原点y坐标</param>
        /// <param name="len">矩形边长</param>
        /// <param name="i">矩形序号</param>
        /// <returns></returns>
        private double to_pointY(double oy, double len, int i)
        {
            double to_y = oy - len;
            //MessageBox.Show(to_y.ToString());
            return to_y;
        }

        private void btn_cancel_Click(object sender, EventArgs e)
        {
            if (thread != null)
            {
                thread.Abort();
            }
            this.Close();
        }

        private void btn_saveClose_Click(object sender, EventArgs e)
        {
            if (thread != null && thread.ThreadState != ThreadState.Stopped)
            {
                return;
            }
            if (taskExecuted)
            {
                if (checkEdit_addToView.CheckState == CheckState.Checked)
                {
                    spreadsheetControl.LoadDocument(docPath);
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

    public class POIAllInfoGaode
    {
        public string status { get; set; }
        public string count { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }

        public SuggestionGaode sugeestion { get; set; }

        public List<PoiGaode> pois = new List<PoiGaode>();
    }

    public class SuggestionGaode
    {
        public string keywords { get; set; }
        public string cities { get; set; }
    }

    public class PoiGaode
    {
        public string id { get; set; }
        public string name { get; set; }
        public string type { get; set; }
        public string typecode { get; set; }
        public string biz_type { get; set; }
        public string address { get; set; }
        public string location { get; set; }
        public string tel { get; set; }
        public string distance { get; set; }
        public string biz_ext { get; set; }
        public string pname { get; set; }
        public string cityname { get; set; }
        public string adname { get; set; }
    }
}
