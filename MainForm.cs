using AutocompleteMenuNS;
using CSScriptLibrary;
using DevExpress.LookAndFeel;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.XtraRichEdit.Services;
using DevExpress.XtraTreeList.Nodes;
using DotSpatial.Controls;
using GeoSharp2018.SystemForms;
using GeoSharp2018.ToolForms;
using GeoSharp2018.ToolForms.CoordTrans;
using GeoSharp2018.UtilClass;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace GeoSharp2018
{
    public partial class MainForm : DevExpress.XtraBars.Ribbon.RibbonForm
    {

        public DefaultLookAndFeel defaultLookAndFeel;

        public string currentCmd;

        public Map mapControl;

        public Legend legendControl;

        public ExtentCoord extentCoord;

        public MainForm()
        {
            InitializeComponent();

            defaultLookAndFeel = defaultLookAndFeel1;

            extentCoord = new ExtentCoord();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            mapControl = new Map();
            mapControl.Dock = DockStyle.Fill;
            mapControl.MapFrame.LegendText = "MapFrame";
            
            dockPanel_localMapView.Controls.Add(mapControl);

            legendControl = new Legend();
            legendControl.Dock = DockStyle.Fill;

            dockPanel_layerMgr.Controls.Add(legendControl);

            mapControl.Legend = legendControl;


            InitialSkin();

            //string layout_path = string.Format(@"{0}\{1}", Application.StartupPath, "layout.xml");
            //dockManager1.RestoreLayoutFromXml(layout_path);

            InitialToolBoxTree();

            BuildAutocompleteMenu();

            try
            {
                webBrowser1.Navigate(Application.StartupPath + @"\map\gaode.html");
            }
            catch
            {

            }




        }

        private void InitialSkin()
        {
            string skinPath = string.Format(@"{0}\{1}", Application.StartupPath, "skins.xml");
            StreamReader sr = new StreamReader(skinPath, Encoding.Default);
            String line;
            string[] strParts = new string[3];
            while ((line = sr.ReadLine()) != null)
            {
                strParts = line.Split(',');

                if (strParts[2] == "1")
                {
                    defaultLookAndFeel.LookAndFeel.SkinName = strParts[1];
                    return;
                }

                defaultLookAndFeel1.LookAndFeel.SkinName = "Office 2013";
            }

            sr.Close();
        }

        private void InitialToolBoxTree()
        {
            TreeListNode rootNode = CreateTreeNode("Toolbox", 0, 0, null);

            TreeListNode toolboxNode;
            TreeListNode toolsetNode;
            TreeListNode toolNode;

            //地理编码工具箱
            toolboxNode = CreateTreeNode("Geocoding toolbox", 1, 1, rootNode);

            toolsetNode = CreateTreeNode("Baidu geocoding toolset", 2, 2, toolboxNode);
            toolNode = CreateTreeNode("Baidu map geocoding tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("Baidu map reverse geocoding tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("Baidu map reverse geocoding pro tool", 3, 4, toolsetNode);

            toolsetNode = CreateTreeNode("Gaode geocoding toolset", 2, 2, toolboxNode);
            toolNode = CreateTreeNode("AMap geocoding tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("AMap reverse geocoding tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("AMap reverse geocoding pro tool", 3, 4, toolsetNode);

            toolsetNode = CreateTreeNode("Tencent geocoding toolset", 2, 2, toolboxNode);
            toolNode = CreateTreeNode("Tencent geocoding tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("Tencent reverse geocoding tool", 3, 4, toolsetNode);

            //坐标转换工具箱
            toolboxNode = CreateTreeNode("Coordinate toolbox", 1, 1, rootNode);

            toolNode = CreateTreeNode("Baidu_to_WGS84", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("GCJ_to_WGS84", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("WGS84_to_Baidu", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("WGS84_to_GCJ", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("GCJ_to_Baidu", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("Baidu_to_GCJ", 3, 4, toolboxNode);

            //POI数据采集工具箱
            toolboxNode = CreateTreeNode("POI collection toolbox", 1, 1, rootNode);

            toolsetNode = CreateTreeNode("AMap POI collection toolset", 2, 2, toolboxNode);
            toolNode = CreateTreeNode("AMap POI tool", 3, 4, toolsetNode);

            toolsetNode = CreateTreeNode("Baidu POI collection toolset", 2, 2, toolboxNode);
            toolNode = CreateTreeNode("Baidu POI tool", 3, 4, toolsetNode);

            //AOI数据采集工具箱
            toolboxNode = CreateTreeNode("AOI collection toolbox", 1, 1, rootNode);

            toolNode = CreateTreeNode("AMap AOI collection tool", 3, 4, toolboxNode);

            //交通态势数据采集工具箱
            toolboxNode = CreateTreeNode("Traffic flow data toolbox", 1, 1, rootNode);

            toolNode = CreateTreeNode("Traffic data tool", 3, 4, toolboxNode);

            //天气数据采集工具箱
            toolboxNode = CreateTreeNode("Weather forecast data toolbox", 1, 1, rootNode);
            toolNode = CreateTreeNode("Real time weather forecast tool", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("Weather forecast tool", 3, 4, toolboxNode);

            //行政区划数据采集工具箱
            toolboxNode = CreateTreeNode("Administrative data toolbox", 1, 1, rootNode);

            toolNode = CreateTreeNode("Province data tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("City data tool", 3, 4, toolsetNode);
            toolNode = CreateTreeNode("County data tool", 3, 4, toolsetNode);

            toolboxNode = CreateTreeNode("Tencent big data toolbox", 1, 1, rootNode);

            toolNode = CreateTreeNode("Tencent big data tool", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("Tencent big data batch tool", 3, 4, toolboxNode);
            toolNode = CreateTreeNode("Population flow data tool", 3, 4, toolboxNode);

            toolboxNode = CreateTreeNode("Trajectory data toolbox", 1, 1, rootNode);

            toolNode = CreateTreeNode("Walk trajectory tool", 3, 4, toolboxNode);
            

            rootNode.ExpandAll();
        }

        private void BuildAutocompleteMenu()
        {
            var items = new List<AutocompleteItem>();

            items.Add(new SubstringAutocompleteItem("GeoPy.Version()", true) { ImageIndex = 0 });
            items.Add(new SubstringAutocompleteItem("GeoPy.Add()", true) { ImageIndex = 0 });
            items.Add(new SubstringAutocompleteItem("GeoPy.Sub()", true) { ImageIndex = 0 });
            items.Add(new SubstringAutocompleteItem("GeoPy.Pro()", true) { ImageIndex = 0 });
            items.Add(new SubstringAutocompleteItem("GeoPy.Div()", true) { ImageIndex = 0 });

            items.Add(new SubstringAutocompleteItem("GeoPy.GetPopFlowData()", true) { ImageIndex = 0 });

            autocompleteMenu1.SetAutocompleteItems(items);
        }

        private TreeListNode CreateTreeNode(string name, int imageIndex, int selectedIndex, TreeListNode parentNode)
        {
            TreeListNode node;
            node = treeList_toolbox.AppendNode(new object[] { name }, parentNode);

            node.ImageIndex = imageIndex;
            node.SelectImageIndex = selectedIndex;

            return node;
        }

        private void richEditControl1_Click(object sender, EventArgs e)
        {

        }

        private void treeList_toolbox_DoubleClick(object sender, EventArgs e)
        {
            string nodeName = treeList_toolbox.FocusedNode[0].ToString();

            if (nodeName == "AMap geocoding tool")
            {
                GaodeGeocodingForm dlg = new GaodeGeocodingForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "AMap reverse geocoding tool")
            {
                GaodeDegeocodingForm dlg = new GaodeDegeocodingForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "AMap reverse geocoding pro tool")
            {
                GaodeDegeocodingAdForm dlg = new GaodeDegeocodingAdForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Baidu map geocoding tool")
            {
                BaiduGeocodingForm dlg = new BaiduGeocodingForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Tencent geocoding tool")
            {
                TenxunGeocodingForm dlg = new TenxunGeocodingForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }

            if (nodeName == "Real time weather forecast tool")
            {
                ForecastColleForm dlg = new ForecastColleForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Weather forecast tool")
            {
                ForecastAllColleForm dlg = new ForecastAllColleForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            
            if (nodeName == "GCJ_to_WGS84")
            {
                Gaode2WGS84Form dlg = new Gaode2WGS84Form(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "WGS84_to_GCJ")
            {
                WGS842GaodeForm dlg = new WGS842GaodeForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
           
            if (nodeName == "WGS84_to_Baidu")
            {
                WGS842BaiduForm dlg = new WGS842BaiduForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Baidu_to_WGS84")
            {
                Baidu2WGS84Form dlg = new Baidu2WGS84Form(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Baidu_to_GCJ")
            {
                Baidu2GaodeForm dlg = new Baidu2GaodeForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "GCJ_to_Baidu")
            {
                Gaode2BaiduForm dlg = new Gaode2BaiduForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Tencent big data tool")
            {
                TenxunQianxiForm dlg = new TenxunQianxiForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }

            if (nodeName == "Tencent big data batch tool")
            {
                TenxunQianxiBatchForm dlg = new TenxunQianxiBatchForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            
            if (nodeName == "Walk trajectory tool")
            {
                WalkRouteForm dlg = new WalkRouteForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }

            if (nodeName == "Population flow data tool")
            {
                PopflowForm dlg = new PopflowForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "AMap POI tool")
            {
                GaodePOIGetForm dlg = new GaodePOIGetForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "Traffic data tool")
            {
                TrafficSituationGetByRecForm dlg = new TrafficSituationGetByRecForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            if (nodeName == "AMap AOI collection tool")
            {
                GaodeAOIGetForm dlg = new GaodeAOIGetForm(this, this.spreadsheetControl1);
                dlg.ShowDialog();
            }
            




        }

        private void btn_runCommand_Click(object sender, EventArgs e)
        {

            IGeoPy geopy = CSScript.Evaluator
                                 .LoadFile<IGeoPy>(Application.StartupPath + @"\HelloScript.cs");

            currentCmd = textEdit_cmdline.Text;
            string msg;

            if (currentCmd.Equals("GeoPy.Version()"))
            {
                msg = GepPySystem.Version();

                memoEdit_cmdMsg.Text += string.Format("{0}{1}", msg, "\r\n");
            }



            if (currentCmd.StartsWith("GeoPy.Add"))
            {

                string[] sArray1 = currentCmd.Split(new string[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
                string str1 = sArray1[1];
                string[] sArray2 = str1.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);


                int a = Convert.ToInt32(sArray2[0]);
                int b = Convert.ToInt32(sArray2[1]);

                memoEdit_cmdMsg.Text += string.Format("执行语句：{0}", "\r\n");
                memoEdit_cmdMsg.Text += string.Format("{0}{1}", currentCmd, "\r\n");
                memoEdit_cmdMsg.Text += string.Format("执行结果：{0}", "\r\n");
                memoEdit_cmdMsg.Text += string.Format("{0}{1}", (geopy.Add(a, b)), "\r\n");
                memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");

            }
            if (currentCmd.StartsWith("GeoPy.GetPopFlowData"))
            {
                if (currentCmd.Equals("GeoPy.GetPopFlowData()"))
                {
                    memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");
                    memoEdit_cmdMsg.Text += string.Format("{0}{1}", "调用人口流量数据采集工具完成！", "\r\n");
                    memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");

                    PopflowForm dlg = new PopflowForm(this, spreadsheetControl1);
                    dlg.Show();
                }
                else
                {
                    string[] sArray1 = currentCmd.Split(new string[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
                    string str1 = sArray1[1];
                    string[] sArray2 = str1.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);


                    string a_path = sArray2[0].ToString();
                    string b_areaNum = sArray2[1].ToString();
                    int c_timeInterval = Convert.ToInt32(sArray2[2]);
                   
                    string from_date = sArray2[3].ToString();
                    string to_date = sArray2[4].ToString();
                    
                    int f_yy = Convert.ToInt32(from_date.Split('-')[0].ToString());
                    int f_mm = Convert.ToInt32(from_date.Split('-')[1].ToString());
                    int f_dd = Convert.ToInt32(from_date.Split('-')[2].ToString());

                    int t_yy = Convert.ToInt32(to_date.Split('-')[0].ToString());
                    int t_mm = Convert.ToInt32(to_date.Split('-')[1].ToString());
                    int t_dd = Convert.ToInt32(to_date.Split('-')[2].ToString());


                    PopflowForm dlg = new PopflowForm(this, spreadsheetControl1);

                    for (DateTime dt = new DateTime(f_yy, f_mm, f_dd); dt < new DateTime(t_yy, t_mm, t_dd); dt = dt.AddDays(1))
                    {

                        string date = dt.ToString("yyyy-MM-dd").Replace("/", "-");
            
                        string codate = date.Replace("-", "");

                        string str_p = string.Format(@"{0}\f{1}",a_path,codate);

                        bool exi = Directory.Exists(str_p);
                        if (!exi)
                        {
                            Directory.CreateDirectory(str_p);

                            memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");
                            memoEdit_cmdMsg.Text += string.Format("文件夹{0}创建完成！{1}", str_p, "\r\n");
                            memoEdit_cmdMsg.Text += string.Format("任务执行正常{0}", "\r\n");
                            memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");

                            memoEdit_cmdMsg.SelectionStart = memoEdit_cmdMsg.Text.Length;
                            memoEdit_cmdMsg.ScrollToCaret();
                            Application.DoEvents();

                            dlg.Execute_by_command(str_p, b_areaNum, c_timeInterval, date);
                        }


                    }

                    if ("aa" == "bb")
                    {
                        // (G:\数据采集\景点人流\test7,5381,60,2018-5-1,2018-5-7)
                        //dlg.Execute_by_command(a_path, b_areaNum, c_timeInterval, d_date);

                        //memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");
                        //memoEdit_cmdMsg.Text += string.Format("正在执行任务：{0}{1}", i, "\r\n");
                        //memoEdit_cmdMsg.Text += string.Format("{0}{1}", "------------------------------------------------------------------------------------------------------------", "\r\n");

                        //memoEdit_cmdMsg.SelectionStart = memoEdit_cmdMsg.Text.Length;
                        //memoEdit_cmdMsg.ScrollToCaret();
                        //Application.DoEvents();
                    }

                }
                
                
            }
            
        }

        /// <summary>
        /// 新建表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barBtn_newGrid_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            spreadsheetControl1.CreateNewDocument();
        }
        /// <summary>
        /// 打开表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barBtn_openGrid_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            spreadsheetControl1.LoadDocument(this);
        }
        /// <summary>
        /// 保存表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barBtn_saveGrid_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            spreadsheetControl1.SaveDocument(this);
        }
        /// <summary>
        /// 另存表格
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void barBtn_saveAsGrid_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            spreadsheetControl1.SaveDocumentAs(this);
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            
        }

        private void fun_unknown()
        {
            string str = @"https://lbs.gtimg.com/maplbs/qianxi/00000000/37010006.js";

            WebClientto client = new WebClientto(4500);
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            string str_json = null;

            stream = client.OpenRead(str);

            str_json = new StreamReader(stream).ReadToEnd();

            Regex regexObj = new Regex(@"\[(?<result>)[^[\]]+\]");
            System.Text.RegularExpressions.Match matchResult = regexObj.Match(str_json);
            while (matchResult.Success)
            {
                MessageBox.Show(matchResult.Groups[0].Value);
                matchResult = matchResult.NextMatch();
            }
        }

        public void ttt()
        {
            string str = @"https://lbs.gtimg.com/maplbs/qianxi/20180411/00000000/37010006.js";

            HttpWebRequest request;
            request = (HttpWebRequest)WebRequest.Create(str);
            request.Method = "GET";
            HttpWebResponse response;
            response = (HttpWebResponse)request.GetResponse();
            Stream s;
            s = response.GetResponseStream();
            string StrDate = "";
            string strValue = "";

            StreamReader Reader = new StreamReader(s, Encoding.UTF8);
            while ((StrDate = Reader.ReadLine()) != null)
            {
                //strValue += StrDate + "\r\n";

                string s1 = StrDate.Split(new char[] { '(', ')' })[1].Replace("[", "").Replace("]", "");

                MessageBox.Show(s1);
            }

            //MessageBox.Show(strValue);
        }

        ///注册配置
        private void barBtn_authority_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            RegForm dlg = new RegForm();
            dlg.ShowDialog();
        }
        ///系统皮肤
        private void barBtn_sysStyleMgr_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            SkinsForm dlg = new SkinsForm(this.defaultLookAndFeel1);
            dlg.ShowDialog();
        }
        ///系统布局
        private void barBtn_layout_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string layout_path = string.Format(@"{0}\{1}", Application.StartupPath, "layout.xml");
            dockManager1.SaveLayoutToXml(layout_path);
        }
        ///帮助文档
        private void barBtn_helpDocument_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {

        }
        ///关于系统
        private void barBtn_systemAbout_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            AboutForm dlg = new AboutForm();
            dlg.ShowDialog();
        }

        private void barBtn_poiExtentMgr_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            ExtentLibForm dlg = new ExtentLibForm(this);

            dlg.ShowDialog();
        }

        //高德地图
        private void barBtn_gaodeMap_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                webBrowser1.Navigate(Application.StartupPath + @"\map\gaode.html");
            }
            catch
            {

            }
        }

        //百度地图
        private void barBtn_baiduMap_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                webBrowser1.Navigate(Application.StartupPath + @"\map\baidu.html");
            }
            catch
            {

            }
        }

        //腾讯地图
        private void barBtn_tengxunMap_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            try
            {
                webBrowser1.Navigate(Application.StartupPath + @"\map\tenxun.html");
            }
            catch
            {

            }
        }

        //工具箱
        private void barBtn_toolboxView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_Toolbox .Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_Toolbox.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_Toolbox.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }

        //地图视图
        private void barBtn_mapView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_webMapView.Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_webMapView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_webMapView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }

        //表格视图
        private void barBtn_gridView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_gridView.Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_gridView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_gridView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }

        //命令窗口
        private void barBtn_cmdView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_command.Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_command.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_command.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }

        //消息窗口
        private void barBtn_msgView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_msgView.Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_msgView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_msgView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }

        private void barBtn_addSpatialData_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.AddLayer();
        }

        private void barBtn_addRasterData_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.AddImageLayer();
        }

        private void barBtn_zoomIn_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.FunctionMode = FunctionMode.ZoomIn;
        }

        private void barBtn_zoomOut_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.FunctionMode = FunctionMode.ZoomOut;
        }

        private void barBtn_pan_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.FunctionMode = FunctionMode.Pan;
        }

        private void barBtn_fullExtent_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.ZoomToMaxExtent();
        }

        private void barBtn_select_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.FunctionMode = FunctionMode.Select;
        }

        private void barBtn_query_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            mapControl.FunctionMode = FunctionMode.Info;
        }

        private void barBtn_localMapView_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_localMapView.Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_localMapView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_localMapView.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }

        private void barBtn_layerMgr_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            if (dockPanel_layerMgr.Visibility == DevExpress.XtraBars.Docking.DockVisibility.Hidden)
            {
                dockPanel_layerMgr.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Visible;
            }
            else
            {
                dockPanel_layerMgr.Visibility = DevExpress.XtraBars.Docking.DockVisibility.Hidden;
            }
        }
    }

    public interface IGeoPy
    {
        int Add(int x, int y);
        int Sub(int x, int y);
        int Pro(int x, int y);
        int Div(int x, int y);
    }

    public class GepPySystem
    {
        public static string Version()
        {
            return "当前版本：GeoSharp2.0";
        }
    }
}
