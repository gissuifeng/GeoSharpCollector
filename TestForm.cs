using DevExpress.Spreadsheet;
using Newtonsoft.Json;
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
using System.Windows.Forms;

namespace GeoSharp2018
{
    public partial class TestForm : Form
    {
        public TestForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            string str = @"http://restapi.amap.com/v3/direction/walking?origin=116.434307,39.90909&destination=116.434446,39.90816&key=27caa753be4090132a65386ed3efff97";
            Stream stream = client.OpenRead(str);

            RouteAllInfo routeAllInfo = new RouteAllInfo();
            RouteInfo routeInfo = new RouteInfo();
            Paths paths = new Paths();
            List<Steps> stepList = new List<Steps>();

            string str_json = new StreamReader(stream).ReadToEnd();

            JObject obj = JObject.Parse(str_json);

            routeAllInfo.status = obj["status"].ToString();
            routeAllInfo.info = obj["info"].ToString();
            routeAllInfo.infocode = obj["infocode"].ToString();
            routeAllInfo.count = obj["count"].ToString();

            routeInfo.origin = obj["route"]["origin"].ToString();
            routeInfo.destination = obj["route"]["destination"].ToString();

            JArray jlist = JArray.Parse(obj["route"]["paths"].ToString());
            JObject obj1 = JObject.Parse(jlist[0].ToString());

            paths.distance = obj1["distance"].ToString();
            paths.duration = obj1["duration"].ToString();

            JArray jlist1 = JArray.Parse(obj1["steps"].ToString());

            Steps steps = null;
            JObject tempJO = null;
            for (int i = 0; i < jlist1.Count(); i++)
            {
                tempJO = JObject.Parse(jlist1[i].ToString());
                steps = new Steps();

                steps.instruction = tempJO["instruction"].ToString();
                steps.orientation = tempJO["orientation"].ToString();
                steps.road = tempJO["road"].ToString();
                steps.distance = tempJO["distance"].ToString();
                steps.duration = tempJO["duration"].ToString();
                steps.polyline = tempJO["polyline"].ToString();
                steps.action = tempJO["action"].ToString();
                steps.assistant_action = tempJO["assistant_action"].ToString();
                steps.walk_type = tempJO["walk_type"].ToString();

                stepList.Add(steps);
            }

            paths.steps = stepList;
            routeInfo.pathInfos = paths;
            routeAllInfo.routeInfos = routeInfo;

                

            int s;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            for (DateTime dt = new DateTime(2018, 5, 1); dt < new DateTime(2018, 5, 31); dt = dt.AddDays(1))
            {
                string date = dt.ToString("yyyy-MM-dd").Replace("/", "-");
                Console.WriteLine(date);
                string codate = date.Replace("-", "");
                string path = string.Format(@"C:\Users\Zhanghaiping\Desktop\test\V3\首都机场{0}", codate);
                Console.WriteLine(path);
                bool exi = Directory.Exists(path);
                if (!exi)
                {
                    Directory.CreateDirectory(path);
                }

                string h, m;

                for (int i = 0; i < 24; i++)
                {
                    for (int j = 0; j < 60; j = j + 5)
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
                    }
                }
                
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string path = @"G:\数据采集\景点人流\北京欢乐谷";
            string date = "2018-05-20";
            

            string h, m;

            for (int i = 0; i < 24; i++)
            {
                for (int j = 0; j < 60;j=j+5 )
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
                }
            }
            
        }

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
            workbook.SaveDocument(string.Format(@"{0}\r{1}_{2}{3}.xls",path, date.Replace("-",""), h, m));

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

                coordVal.coordX = Convert.ToInt32(strs[i].Split(',')[0]);
                coordVal.coordY = Convert.ToInt32(strs[i].Split(',')[1]);
                coordVal.val = Convert.ToInt32(strs[i + 1].Replace(":", "").Replace(",", ""));

                //Console.WriteLine(string.Format("x:{0}, y:{1}, val:{2}", coordVal.coordX, coordVal.coordY, coordVal.val));

                worksheet[rowOrder, 0].SetValue(coordVal.coordX);
                worksheet[rowOrder, 1].SetValue(coordVal.coordY);
                worksheet[rowOrder, 2].SetValue(coordVal.val);

            }

            workbook.SaveDocument(string.Format(@"{0}\r{1}_{2}{3}.xls", path, date.Replace("-",""), h, m));
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string path = @"G:\数据采集\景点人流\test7";
            string date = "2018-04-29";

            string h, m;

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            string str = @"https://heat.qq.com/api/getRegionHeatMapInfoById.php?id=5381";
            Stream stream = client.OpenRead(str);

            string str_json = new StreamReader(stream).ReadToEnd();

            JObject obj = JObject.Parse(str_json);

            string max = obj["max"].ToString();
            string name = obj["name"].ToString();
            string boundary = obj["boundary"].ToString();
            string center_gcj = obj["center_gcj"].ToString();
            string lower_left = obj["lower_left"].ToString();
            string upper_right = obj["upper_right"].ToString();

            Workbook workbook = new Workbook();
            workbook.CreateNewDocument();
            workbook.SaveDocument(string.Format(@"{0}\tt01.xlsx", path));

            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Name = "routeSheet";

            worksheet[0, 0].SetValue("oid");
            worksheet[0, 1].SetValue("b_name");
            worksheet[0, 2].SetValue("b_type");
            worksheet[0, 3].SetValue("lng");
            worksheet[0, 4].SetValue("lat");

            worksheet[1, 0].SetValue(1);
            worksheet[1, 1].SetValue(name);
            worksheet[1, 2].SetValue("center_pnt");
            worksheet[1, 3].SetValue(center_gcj.Split(',')[0]);
            worksheet[1, 4].SetValue(center_gcj.Split(',')[1]);

            worksheet[2, 0].SetValue(2);
            worksheet[2, 1].SetValue(name);
            worksheet[2, 2].SetValue("lower_left");
            worksheet[2, 3].SetValue(lower_left.Split(',')[0]);
            worksheet[2, 4].SetValue(lower_left.Split(',')[1]);

            worksheet[3, 0].SetValue(3);
            worksheet[3, 1].SetValue(name);
            worksheet[3, 2].SetValue("upper_right");
            worksheet[3, 3].SetValue(upper_right.Split(',')[0]);
            worksheet[3, 4].SetValue(upper_right.Split(',')[1]);

            string[] boundStrs = boundary.Split('|');
            for (int i = 0; i < boundStrs.Length-1; i++)
            {
                worksheet[i+4, 0].SetValue(i+4);
                worksheet[i + 4, 1].SetValue(name);
                worksheet[i + 4, 2].SetValue("bound_pnt");
                worksheet[i + 4, 3].SetValue(boundStrs[i].Split(',')[0]);
                worksheet[i + 4, 4].SetValue(boundStrs[i].Split(',')[1]);
            }

            workbook.SaveDocument(string.Format(@"{0}\tt01.xlsx", path));


                //MessageBox.Show(obj["max"].ToString());
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\Zhanghaiping\Desktop\test\V2\厦大0429";
            string date = "2018-04-29";

            string h, m;

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            string str = @"https://xingyun.map.qq.com/api/getXingyunPoints";
            Stream stream = client.OpenRead(str);

            string str_json = new StreamReader(stream).ReadToEnd();

            Console.WriteLine(str_json);
        }

        private void TestForm_Load(object sender, EventArgs e)
        {

        }

        
    }

    public class CoordVal
    {
        public int coordX { get; set; }
        public int coordY { get; set; }
        public int val { get; set; }
    }

    public class RouteAllInfo
    {
        public string status { get; set; }
        public string info { get; set; }
        public string infocode { get; set; }
        public string count { get; set; }

        public RouteInfo routeInfos { get; set; }
    }

    public class RouteInfo
    {
        public string origin { get; set; }
        public string destination { get; set; }

        public Paths pathInfos { get; set; }

    }

    public class Paths
    {
        public string distance { get; set; }
        public string duration { get; set; }

        public List<Steps> steps { get; set; }
    }

    public class Steps
    {
        public string instruction { get; set; }
        public string orientation { get; set; }
        public string road { get; set; }
        public string distance { get; set; }
        public string duration { get; set; }
        public string polyline { get; set; }
        public string action { get; set; }
        public string assistant_action { get; set; }
        public string walk_type { get; set; }
    }
    
}
