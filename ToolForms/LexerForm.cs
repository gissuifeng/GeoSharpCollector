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
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace GeoSharp2018.ToolForms
{
    public partial class LexerForm : DevExpress.XtraEditors.XtraForm
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
    

        public LexerForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void labelControl3_Click(object sender, EventArgs e)
        {

        }

        private void btn_inputOpen_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;

            comboBox_inputSheet.Properties.Items.Clear();
            comboBox_inputField.Properties.Items.Clear();

            string fileName;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "excel2003文件|*.xls|excel文件|*.xlsx";
            openFileDialog.RestoreDirectory = true;
            openFileDialog.FilterIndex = 2;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                docPathOriginal = fileName = openFileDialog.FileName;

                textEdit_inputFile.EditValue = fileName;

                workbookOriginal = new Workbook();

                workbookOriginal.LoadDocument(textEdit_inputFile.EditValue.ToString());

                for (int i = 0; i < workbookOriginal.Worksheets.Count; i++)
                {
                    comboBox_inputSheet.Properties.Items.Add(workbookOriginal.Worksheets[i].Name);
                }

                comboBox_inputSheet.SelectedIndex = 0;

                worksheetOriginal = workbookOriginal.Worksheets[0];

                for (int j = 0; j < worksheetOriginal.GetDataRange().ColumnCount; j++)
                {
                    comboBox_inputField.Properties.Items.Add(worksheetOriginal[0, j].Value);                   
                }

                comboBox_inputField.SelectedIndex = 0;

                isFirstLoad = false;
            }
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

        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
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

                    Workbook tem_workbook = new Workbook();
                    tem_workbook.LoadDocument(textEdit_inputFile.EditValue.ToString());
                    Range range = tem_workbook.Worksheets[0].GetUsedRange();

                    progressBarControl1.Properties.Maximum = range.RowCount;

                    thread = new Thread(new ThreadStart(fun1));
                    thread.Start();
                }
                catch (Exception ex)
                {
                    XtraMessageBox.Show(ex.Message);
                }

                btn_ok.Enabled = false;
            }

        }



        /// <summary>

        /// 删除标点符号

        /// </summary>

        /// <param name="str"></param>

        /// <returns></returns>

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


        private void fun1()
        {
            Account account = null;
            List<Account> list = new List<Account>();

            Workbook tem_workbook = new Workbook();
            tem_workbook.LoadDocument(textEdit_inputFile.EditValue.ToString());

            int col = comboBox_inputField.SelectedIndex;

            Range range = tem_workbook.Worksheets[0].GetUsedRange();

            var nlp = new Nlp(clientId, clientSecret);

            string targetStr = "";
            for (int i = 1; i < range.RowCount; i++)
            {
                targetStr = tem_workbook.Worksheets[0][i, col].Value.ToString();
                //MessageBox.Show(targetStr);
                var result = nlp.Lexer(RemovePunctuation(targetStr).Trim());
                //MessageBox.Show(result.ToString());
                if (result.ToString() != "")
                {
                    try
                    {
                        account = JsonConvert.DeserializeObject<Account>(result.ToString());

                        list.Add(account);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                sum++;
                RunWithInoke(sum);
            }

            int temp_count=0;
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = 0; j < list[i].items.Count;j++ )
                {
                    

                    worksheetNew[temp_count + 1, 0].SetValue(temp_count + 1);
                    worksheetNew[temp_count + 1, 1].SetValue(list[i].log_id);
                    worksheetNew[temp_count + 1, 2].SetValue(list[i].text);

                    worksheetNew[temp_count + 1, 3].SetValue(list[i].items[j].item);
                    worksheetNew[temp_count + 1, 4].SetValue(list[i].items[j].bype_offset);
                    worksheetNew[temp_count + 1, 5].SetValue(list[i].items[j].uri);
                    worksheetNew[temp_count + 1, 6].SetValue(list[i].items[j].pos);
                    worksheetNew[temp_count + 1, 7].SetValue(list[i].items[j].ne);
                    worksheetNew[temp_count + 1, 8].SetValue(list[i].items[j].item);
                    worksheetNew[temp_count + 1, 9].SetValue(list[i].items[j].item);
                    worksheetNew[temp_count + 1, 10].SetValue(list[i].items[j].byte_length);
                    worksheetNew[temp_count + 1, 11].SetValue(list[i].items[j].formal);

                    temp_count++;
                }
                
            }

            taskExecuted = true;
            workbookNew.SaveDocument(docPathNew);



            XtraMessageBox.Show("所有天气信息已经解析完成！");

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

        }

        private void btn_saveClose_Click(object sender, EventArgs e)
        {

        }

        private void btn_help_Click(object sender, EventArgs e)
        {

        }

        private void LexerForm_Load(object sender, EventArgs e)
        {

        }
    }

    public class Account
    {
        public string log_id { get; set; }
        public string text { get; set; }
        public IList<Item> items { get; set; }
    }



    public class Item
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
}
