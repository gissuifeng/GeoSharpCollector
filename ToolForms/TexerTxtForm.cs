using Baidu.Aip.Nlp;
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
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace GeoSharp2018.ToolForms
{
    public partial class TexerTxtForm : DevExpress.XtraEditors.XtraForm
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
        private Workbook workbookOriginal;
        /// <summary>
        /// 当前工作簿
        /// </summary>
        private Workbook workbookNew;
        /// <summary>
        /// 当前工作表
        /// </summary>
        private Worksheet worksheetOriginal;
        /// <summary>
        /// 当前工作表
        /// </summary>
        private Worksheet worksheetNew;
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
        string docPathOriginal;
        /// <summary>
        /// 文件路径
        /// </summary>
        string docPathNew;
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

        // 调用getAccessToken()获取的 access_token建议根据expires_in 时间 设置缓存
        // 返回token示例
        public static String TOKEN = "24.aaf9630b7f203a1d970f112006945230.2592000.1525444092.282335-11050040";

        // 百度云中开通对应服务应用的 API Key 建议开通应用的时候多选服务
        private static String clientId = "p6ESoMXyl39ofm6AkMBjdf9Y";
        // 百度云中开通对应服务应用的 Secret Key
        private static String clientSecret = "B8WB0UF9MpvNNQGMmS6sh4Siam2MZgGV";

        public TexerTxtForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void TexerTxtForm_Load(object sender, EventArgs e)
        {

        }

        private string inputFilesPath = null;
        private void btn_inputOpen_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            dialog.Description = "请选择TXT所在文件夹";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                if (string.IsNullOrEmpty(dialog.SelectedPath))
                {
                    XtraMessageBox.Show(this, "文件夹路径不能为空", "提示");
                    return;
                }

                inputFilesPath = dialog.SelectedPath;
                textEdit_inputFile.EditValue = inputFilesPath;

            }

            inputFilesPath = @"F:\GIS_Study\paper\话题地理学\source_data\处理后";
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            try
            {
                worksheetNew = workbookNew.Worksheets[0];
                worksheetNew.Name = textEdit_outputName.EditValue.ToString();

                worksheetNew[0, 0].SetValue("uid");
                worksheetNew[0, 1].SetValue("log_id");
                worksheetNew[0, 2].SetValue("text");
                worksheetNew[0, 3].SetValue("loc_details");
                worksheetNew[0, 4].SetValue("byte_offset");

                worksheetNew[0, 5].SetValue("uri");
                worksheetNew[0, 6].SetValue("pos");
                worksheetNew[0, 7].SetValue("ne");
                worksheetNew[0, 8].SetValue("item");
                worksheetNew[0, 9].SetValue("basic_words");
                worksheetNew[0, 10].SetValue("byte_lenth");
                worksheetNew[0, 11].SetValue("formal");
                worksheetNew[0, 12].SetValue("order_id");

                DirectoryInfo dirInfo = new DirectoryInfo(inputFilesPath);
                FileInfo[] fis = dirInfo.GetFiles();

                progressBarControl1.Properties.Maximum = fis.Count();

                thread = new Thread(new ThreadStart(fun1));
                thread.Start();
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }

            btn_ok.Enabled = false;
        }

        private void fun1()
        {
            Account1 account = null;
            List<Account1> list = new List<Account1>();

            FlagCls flagcls = null;
            List<FlagCls> list1 = new List<FlagCls>();

           

            var nlp = new Nlp(clientId, clientSecret);

            string targetStr = "";


            StreamReader sr = null;
            StreamWriter file = null;
            string str;

            DirectoryInfo dirInfo = new DirectoryInfo(inputFilesPath);
            FileInfo[] fis = dirInfo.GetFiles();

            
            long sum_all= 0;
            long sum_item = 0;
            for (int i = 0; i < fis.Length; i++)
            {

                sum_item++;
                sr = new StreamReader(fis[i].FullName);
                
                while ((str = sr.ReadLine()) != null)
                {
                    str = RemovePunctuation(str);
                    
                    if (str != "" && !str.Contains("jpg"))
                    {
                        var result = nlp.Lexer(RemovePunctuation(str).Trim());
                        //MessageBox.Show(result.ToString());
                        if (result.ToString() != "")
                        {
                            try
                            {
                                account = JsonConvert.DeserializeObject<Account1>(result.ToString());

                                list.Add(account);
                                
                                sum_all++;

                                flagcls = new FlagCls();
                                flagcls.uid = sum_all.ToString();
                                flagcls.order = sum_item.ToString();
                                list1.Add(flagcls);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }
                    }
                }

                sum++;
                RunWithInoke(sum);
              
                sr.Close();                                 
            }



            int temp_count = 0;
            try
            {
                for (int i = 0; i < list.Count; i++)
                {
                    for (int j = 0; j < list[i].items.Count; j++)
                    {


                        worksheetNew[temp_count + 1, 0].SetValue(temp_count + 1);
                        worksheetNew[temp_count + 1, 1].SetValue(list[i].log_id);
                        worksheetNew[temp_count + 1, 2].SetValue("");

                        worksheetNew[temp_count + 1, 3].SetValue(list[i].items[j].item);
                        worksheetNew[temp_count + 1, 4].SetValue(list[i].items[j].bype_offset);
                        worksheetNew[temp_count + 1, 5].SetValue(list[i].items[j].uri);
                        worksheetNew[temp_count + 1, 6].SetValue(list[i].items[j].pos);
                        worksheetNew[temp_count + 1, 7].SetValue(list[i].items[j].ne);
                        worksheetNew[temp_count + 1, 8].SetValue(list[i].items[j].item);
                        worksheetNew[temp_count + 1, 9].SetValue(list[i].items[j].item);
                        worksheetNew[temp_count + 1, 10].SetValue(list[i].items[j].byte_length);
                        worksheetNew[temp_count + 1, 11].SetValue(list[i].items[j].formal);
                        worksheetNew[temp_count + 1, 12].SetValue(list1[i].order);

                        temp_count++;
                    }

                }
            }
            catch (Exception ex)
            {

            }

                taskExecuted = true;
                workbookNew.SaveDocument(docPathNew);



                XtraMessageBox.Show("所有信息已经解析完成！");

        }

        private void RunWithInoke(int i)
        {

            progressBarControl1.Invoke(new SetProgressBarValue(SetProgressValue), i);

        }

        private void SetProgressValue(int value)
        {
            progressBarControl1.EditValue = value + 1;
        }

        public static string RemovePunctuation(string str)
        {

            if (str != "")
            {
                str = str.Replace(",", "")
                              .Replace("，", "")
                              .Replace(".", "")
                              .Replace("。", "")
                              .Replace("!", "")
                              .Replace("！", "")
                              .Replace("?", "")
                              .Replace("？", "")
                              .Replace(":", "")
                              .Replace("：", "")
                              .Replace(";", "")
                              .Replace("；", "")
                              .Replace("～", "")
                              .Replace("-", "")
                              .Replace("_", "")
                              .Replace("——", "")
                              .Replace("—", "")
                              .Replace("--", "")
                              .Replace("【", "")
                              .Replace("】", "")
                              .Replace("\\", "")
                              .Replace("(", "")
                              .Replace(")", "")
                              .Replace("（", "")
                              .Replace("）", "")
                              .Replace("#", "")
                              .Replace("”", "")
                              .Replace("“", "")
                              .Replace("\"", "")
                              .Replace("△", "")
                              .Replace("$", "");
            }



            return str;

        }

        private void btn_outputOpen_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;
            string fileName;

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "excel2003文件|*.xls|excel文件|*.xlsx";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FilterIndex = 2;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                docPathNew = fileName = saveFileDialog.FileName;

                textEdit_outputFile.EditValue = fileName;

                workbookNew = new Workbook();

                workbookNew.CreateNewDocument();

                workbookNew.SaveDocument(textEdit_outputFile.EditValue.ToString());

                textEdit_outputName.EditValue = "main_sheet";

                isFirstLoad = false;
            }
        }


    }

    public class Account1
    {
        public string log_id { get; set; }
        public string text { get; set; }
        public IList<Item1> items { get; set; }
    }



    public class Item1
    {
        public IList<string> loc_details { get; set; }
        public int bype_offset { get; set; }
        public string uri { get; set; }
        public string pos { get; set; }
        public string ne { get; set; }
        public string item { get; set; }
        public IList<string> baasec_words { get; set; }
        public int byte_length { get; set; }
        public string formal { get; set; }

    }

    public class FlagCls
    {
        public string uid;
        public string order;
    }
}
