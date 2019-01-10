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
    public partial class GaodeAOIGetForm : DevExpress.XtraEditors.XtraForm
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
       
        public GaodeAOIGetForm(MainForm mainForm, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainForm = mainForm;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void btn_open_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;
            comboBox_AOICodeField.Properties.Items.Clear();
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
                    comboBox_AOICodeField.Properties.Items.Add(worksheet[0, j].Value);
                }

                comboBox_AOICodeField.SelectedIndex = 0;

                isFirstLoad = false;
            }
        }

        string[] strs;
        int colCount;
        int rowCount;

        Worksheet worksheet_new;
        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                try
                {
                    col = comboBox_AOICodeField.SelectedIndex;

                    colCount = worksheet.GetDataRange().ColumnCount;
                    rowCount = worksheet.GetDataRange().RowCount;

                    strs = new string[rowCount];

                    progressBarControl1.Properties.Maximum = rowCount;
                    pro_dt = DateTime.Now;

                    for (int i = 1; i < rowCount; i++)
                    {
                        strs[i] = worksheet[i, col].Value.ToString();
                    }

                    worksheet[0, colCount].SetValue("status");
                    worksheet[0, colCount + 1].SetValue("code");
                    worksheet[0, colCount + 2].SetValue("city_name");
                    worksheet[0, colCount + 3].SetValue("new_type");
                    worksheet[0, colCount + 4].SetValue("title");

                    worksheet[0, colCount + 5].SetValue("classify");
                    worksheet[0, colCount + 6].SetValue("business");
                    worksheet[0, colCount + 7].SetValue("cityadcode");
                    worksheet[0, colCount + 8].SetValue("cre_flag");
                    worksheet[0, colCount + 9].SetValue("std_t_tag_0_v");
                    worksheet[0, colCount + 10].SetValue("navi_geometry");

                    worksheet[0, colCount + 12].SetValue("poiid");
                    worksheet[0, colCount + 13].SetValue("address");
            

                    worksheet_new = workbook.Worksheets.Add();

                    worksheet_new[0, 0].SetValue("uid");
                    worksheet_new[0, 1].SetValue("AOIID");
                    worksheet_new[0, 2].SetValue("lng");
                    worksheet_new[0, 3].SetValue("lat");

                    thread = new Thread(new ThreadStart(fun));
                    thread.Start();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }

            }
        }

        string coordStr = "";
        private void fun()
        {
            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");
            client.Headers.Add("Cookie", "cna=4LMgEHtMgygCAdOiGvZVKFqd; l=AtzcahYI/-GKqMsd/C5BlLYMrPCOVYB/; isg=BO_vu7kn6N-5kOD93d3GUDu1fQA5PEI-2L7ACQF9t95lUApSCWHZBgpO1oBLMxsu; CNZZDATA1255626299=1780542841-1522410376-https%253A%252F%252Fwww.baidu.com%252F%7C1538873610; _uab_collina=152241487386224945552091; guid=3518-40cb-d687-efee; _ga=GA1.2.1630710146.1536216740; UM_distinctid=166432e04be14b-05a28b1845a6718-4c312878-1fa400-166432e04bf3dd; passport_login=OTk2MzAxMSxnaXNzdWlmZW5nLDdmazNrMWc5bGN0NTljeThkeGpjMzFhYjZzaG9keDA5LDE1Mzg3MjU0MTEsTm1aaFpqRTFZMlJtTVRoaVl6YzBZVFppTnpGaE5URTFPVEEzTlRjMU1EZz0%3D; dev_help=Gcg7p7C6v42wQCkYJ5IEJjg5OWNlZGI3YjBjMmRmNjJlMTMyYWRhYzliNjVjY2I0OTJkM2ExZDU1Y2VkMTA2MTFmODc0OTU3Y2YxMWRjNDmI3c8KuhOrEJVuVJ35UM8dhxtnxeeCGjaEHNsWe9fRiWDACZbNXQ%2Fql2q9yKGjFfgOCFbYEspFLNVO1G3PH764kNVtbGJ9dl6%2FMIuI2QKRK4UqGX8kFVVSvmRkRdqsMGw%3D; _umdata=BA335E4DD2FD504F217402697004971AD7520D922636FA245583C1035B20126578D68D34A1D57411CD43AD3E795C914CFEA6A7DA26A1BEC3656F3B9F0DAD776B; key=bfe31f4e0fb231d29e1d3ce951e2c780; x5sec=7b22617365727665723b32223a2234613762306633653261636536343230366433343134643364373834663566364349727035643046455072736a634c486d70364654773d3d227d");

            Stream stream = null;
            string str_json = null;
            List<AOIInfoCls> list = new List<AOIInfoCls>();
            List<string> coorList = new List<string>();

            string addressName;

            int i = 1;
            for (; i < rowCount; i++)
            {
                addressName = worksheet[i, comboBox_AOICodeField.SelectedIndex].Value.ToString();

                try
                {
                    stream = client.OpenRead(SetAddress(addressName));

                    str_json = new StreamReader(stream).ReadToEnd();
                }
                catch
                {

                }

                if (str_json != "")
                {
                    AOIInfoCls aoiinfo = new AOIInfoCls();

                    JObject obj = JObject.Parse(str_json);

                    if (obj["status"].ToString() != "1")
                    {

                        aoiinfo.status = obj["status"].ToString();

                    }
                    else
                    {

                        aoiinfo.status = obj["status"].ToString();

                        aoiinfo.code = obj["data"]["base"]["code"].ToString();
                        aoiinfo.new_type = obj["data"]["base"]["new_type"].ToString();
                        aoiinfo.city_name = obj["data"]["base"]["city_name"].ToString();
                        aoiinfo.title = obj["data"]["base"]["title"].ToString();

                        aoiinfo.classify = obj["data"]["base"]["classify"].ToString();
                        aoiinfo.business = obj["data"]["base"]["business"].ToString();
                        aoiinfo.cityadcode = obj["data"]["base"]["city_adcode"].ToString();
                        aoiinfo.cre_flag = obj["data"]["base"]["cre_flag"].ToString();

                        aoiinfo.poiid = obj["data"]["base"]["poiid"].ToString();
                        aoiinfo.address = obj["data"]["base"]["address"].ToString();

                        coordStr = obj["data"]["spec"]["shape"].ToString();
                    }

                    list.Add(aoiinfo);
                    coorList.Add(coordStr);
                }
                sum++;
                RunWithInoke(sum);

                if ((i % 1000) == 0 && (i / 1000) > 0)
                {
                    int cur_count = ((i / 1000) - 1) * 1000 - 1;

                    for (int j = 0; j < list.Count; j++)
                    {
                        cur_count++;


                        worksheet[cur_count + 1, colCount].SetValue(list[j].status);
                        worksheet[cur_count + 1, colCount + 1].SetValue(list[j].code);

                        worksheet[cur_count + 1, colCount + 2].SetValue(list[j].new_type);
                        worksheet[cur_count + 1, colCount + 3].SetValue(list[j].city_name);
                        worksheet[cur_count + 1, colCount + 4].SetValue(list[j].title);

                        worksheet[cur_count + 1, colCount + 5].SetValue(list[j].classify);
                        worksheet[cur_count + 1, colCount + 6].SetValue(list[j].business);
                        worksheet[cur_count + 1, colCount + 7].SetValue(list[j].cityadcode);
                        worksheet[cur_count + 1, colCount + 8].SetValue(list[j].cre_flag);

                        worksheet[cur_count + 1, colCount + 9].SetValue(list[j].poiid);
                        worksheet[cur_count + 1, colCount + 10].SetValue(list[j].address);

                        //coordStr = obj["data"]["spec"]["shape"].ToString();

                    }

                    workbook.SaveDocument(docPath);

                    list.Clear();
                }
            }

            int restRows = (rowCount) - (rowCount % 1000);

            for (int m = 0; m < list.Count; m++)
            {
                
                worksheet[restRows + m + 1, colCount].SetValue(list[m].status);
                worksheet[restRows + m + 1, colCount + 1].SetValue(list[m].code);

                worksheet[restRows + m + 1, colCount + 2].SetValue(list[m].new_type);
                worksheet[restRows + m + 1, colCount + 3].SetValue(list[m].city_name);
                worksheet[restRows + m + 1, colCount + 4].SetValue(list[m].title);

                worksheet[restRows + m + 1, colCount + 5].SetValue(list[m].classify);
                worksheet[restRows + m + 1, colCount + 6].SetValue(list[m].business);
                worksheet[restRows + m + 1, colCount + 7].SetValue(list[m].cityadcode);
                worksheet[restRows + m + 1, colCount + 8].SetValue(list[m].cre_flag);
                
                worksheet[restRows + m + 1, colCount + 9].SetValue(list[m].poiid);
                worksheet[restRows + m + 1, colCount + 10].SetValue(list[m].address);

            }

            
            for(int n=0; n<coorList.Count; n++)
            {
                

                worksheet_new[n + 1, 0].SetValue("");
            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);

            XtraMessageBox.Show("所有数据的地理编码任务已经完成！");
        }

        string gaodeKey = "1ef994aee94edd79c70ea991d98c046c";//"e00536b393e9671af12bea182f75a36b";
        private string SetAddress(string address)
        {
            
            string s = string.Format(@"https://www.amap.com/detail/get/detail?id={0}", address);
            Console.WriteLine(s);
            return s;
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
    }

    public class AOIInfoCls
    {
        public string status;
        public string code;
        public string new_type;
        public string city_name;
        public string title;

        public string classify;
        public string business;
        public string cityadcode;
        public string cre_flag;

        public string poiid;
        public string address;
    }


}
