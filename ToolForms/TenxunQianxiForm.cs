using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
using DevExpress.XtraSpreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

namespace GeoSharp2018.ToolForms
{
    public partial class TenxunQianxiForm : DevExpress.XtraEditors.XtraForm
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
        /// 主窗口对象
        /// </summary>
        private MainForm mainform;
        /// <summary>
        /// 采集数据主线程
        /// </summary>
        Thread thread = null;
        private string strKey = "27caa753be4090132a65386ed3efff97";
        int sum = 0;
        string forecastCodeFile = null;

        public TenxunQianxiForm()
        {
            InitializeComponent();
        }

        private void TenxunQianxiForm_Load(object sender, EventArgs e)
        {
            comboBox_targetObj.Properties.Items.Add("迁入数据");
            comboBox_targetObj.Properties.Items.Add("迁出数据");

            comboBox_targetObj.SelectedIndex = 0;

            dateEdit1.EditValue = DateTime.Now.AddDays(-1);

            comboBox_datatype.Properties.Items.Add("全部");
            comboBox_datatype.Properties.Items.Add("飞机");
            comboBox_datatype.Properties.Items.Add("火车");
            comboBox_datatype.Properties.Items.Add("汽车");
            comboBox_datatype.SelectedIndex = 0;
        }

        public TenxunQianxiForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();


            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void Btn_open_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;
            string fileName;

            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "excel2003文件|*.xls|excel文件|*.xlsx",
                RestoreDirectory = true,
                FilterIndex = 2
            };
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                docPath = fileName = saveFileDialog.FileName;

                textEdit_file.EditValue = fileName;

                workbook = new Workbook();

                workbook.CreateNewDocument();

                workbook.SaveDocument(textEdit_file.EditValue.ToString());

                textEdit_fileName.EditValue = "main_sheet";

                isFirstLoad = false;
            }
        }

        string target_type;
        string target_datatype;
        
        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                target_type = comboBox_targetObj.Properties.Items[comboBox_targetObj.SelectedIndex].ToString();
                forecastCodeFile = Application.StartupPath + @"\data\地级市.xlsx";

                if (target_type == "迁入数据")
                {
                    dataType = "0";
                }
                else if (target_type == "迁出数据")
                {
                    dataType = "1";
                }
                else
                {
                    XtraMessageBox.Show("所选采集对象类型不合法！请正确配置相关参数！");
                    return;
                }

                target_datatype = comboBox_datatype.Properties.Items[comboBox_datatype.SelectedIndex].ToString();
                //MessageBox.Show(target_datatype);
                try
                {
                    worksheet = workbook.Worksheets[0];
                    worksheet.Name = textEdit_fileName.EditValue.ToString();

                    if (target_datatype == "全部")
                    {
                        worksheet[0, 0].SetValue("data_id");
                        worksheet[0, 1].SetValue("data_type");
                        worksheet[0, 2].SetValue("city_name1");
                        worksheet[0, 3].SetValue("date");

                        worksheet[0, 4].SetValue("city_name2");
                        worksheet[0, 5].SetValue("hot_val");
                        worksheet[0, 6].SetValue("car");
                        worksheet[0, 7].SetValue("train");
                        worksheet[0, 8].SetValue("plane");
                        worksheet[0, 9].SetValue("get_date");

                        worksheet[0, 10].SetValue("lng");
                        worksheet[0, 11].SetValue("lat");
                        worksheet[0, 12].SetValue("adcode");
                    }
                    if (target_datatype == "飞机")
                    {
                        //MessageBox.Show("feiji");
                        worksheet[0, 0].SetValue("data_id");
                        worksheet[0, 1].SetValue("data_type");
                        worksheet[0, 2].SetValue("city_name1");
                        worksheet[0, 3].SetValue("date");

                        worksheet[0, 4].SetValue("city_name2");
                        worksheet[0, 5].SetValue("hot_val");
                        worksheet[0, 6].SetValue("plane");
     
                        worksheet[0, 7].SetValue("get_date");

                        worksheet[0, 8].SetValue("lng");
                        worksheet[0, 9].SetValue("lat");
                        worksheet[0, 10].SetValue("adcode");
                    }
                    if (target_datatype == "火车")
                    {
                        worksheet[0, 0].SetValue("data_id");
                        worksheet[0, 1].SetValue("data_type");
                        worksheet[0, 2].SetValue("city_name1");
                        worksheet[0, 3].SetValue("date");

                        worksheet[0, 4].SetValue("city_name2");
                        worksheet[0, 5].SetValue("hot_val");
                        worksheet[0, 6].SetValue("train");

                        worksheet[0, 7].SetValue("get_date");

                        worksheet[0, 8].SetValue("lng");
                        worksheet[0, 9].SetValue("lat");
                        worksheet[0, 10].SetValue("adcode");
                    }
                    if (target_datatype == "汽车")
                    {
                        worksheet[0, 0].SetValue("data_id");
                        worksheet[0, 1].SetValue("data_type");
                        worksheet[0, 2].SetValue("city_name1");
                        worksheet[0, 3].SetValue("date");

                        worksheet[0, 4].SetValue("city_name2");
                        worksheet[0, 5].SetValue("hot_val");
                        worksheet[0, 6].SetValue("car");

                        worksheet[0, 7].SetValue("get_date");

                        worksheet[0, 8].SetValue("lng");
                        worksheet[0, 9].SetValue("lat");
                        worksheet[0, 10].SetValue("adcode");
                    }

                    Workbook tem_workbook = new Workbook();
                    tem_workbook.LoadDocument(forecastCodeFile);
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

        private string SetAddress(string date, string cityCode, string type)
        {
            string s = "";
            try
            {
                if (target_datatype == "全部")
                {
                    s = string.Format(@"https://lbs.gtimg.com/maplbs/qianxi/{0}/{1}{2}6.js", date, cityCode, type);
                }
                if (target_datatype == "飞机")
                {
                    s = string.Format(@"https://lbs.gtimg.com/maplbs/qianxi/{0}/{1}{2}3.js", date, cityCode, type);
                }
                if (target_datatype == "火车")
                {
                    s = string.Format(@"https://lbs.gtimg.com/maplbs/qianxi/{0}/{1}{2}2.js", date, cityCode, type);
                }
                if (target_datatype == "汽车")
                {
                    s = string.Format(@"https://lbs.gtimg.com/maplbs/qianxi/{0}/{1}{2}1.js", date, cityCode, type);
                }
                
               
            }
            catch (Exception ex)
            {
                XtraMessageBox.Show(ex.Message);
            }
            return s;
        }

        string dataType = "";
        private void fun1()
        {
            DateTime dt = Convert.ToDateTime(dateEdit1.EditValue.ToString());
            string date = dt.ToString("yyyyMMdd");

            Workbook tem_workbook = new Workbook();
            tem_workbook.LoadDocument(forecastCodeFile);

            Range range = tem_workbook.Worksheets[0].GetUsedRange();

            WebClient client = new WebClient();
            client.Encoding = Encoding.UTF8;
            client.Headers.Add("Content-Type", "application/x-www-form-urlencoded");

            Stream stream = null;
            string str_json = null;
            List<string> list = new List<string>();
            string[] itemStrs = null;

            for (int i = 1; i < range.RowCount; i++)
            {
                try
                {
                    string vv = SetAddress(date, tem_workbook.Worksheets[0][i, 1].Value.ToString(), dataType);
                    stream = client.OpenRead(vv);
                }
                catch (Exception ex)
                {

                    MessageBox.Show(tem_workbook.Worksheets[0][i, 1].Value.ToString() + "," + tem_workbook.Worksheets[0][i, 2].Value.ToString());
                    MessageBox.Show("操作超时,当前工作将自动退出。请在稳定的网络环境下执行此任务！");

                    continue;

                    if (thread.ThreadState == System.Threading.ThreadState.Running)
                    {
                        thread.Abort();
                    }
                    this.Close();
                }



                str_json = new StreamReader(stream).ReadToEnd();

                if (str_json != "")
                {
                    try
                    {
                        //MessageBox.Show(str_json);
                        Regex regexObj = new Regex(@"\[(?<result>)[^[\]]+\]");
                        System.Text.RegularExpressions.Match matchResult = regexObj.Match(str_json);
                        int tempRow = 1;
                        while (matchResult.Success)
                        {
                            string s1 = matchResult.Groups[0].Value.Replace("[", "").Replace("]", "");
                            //MessageBox.Show(string.Format("i:{0}, {1}", i, s1));
                            itemStrs = s1.Split(',');

                            if (target_datatype == "全部")
                            {
                                worksheet[10 * (i - 1) + tempRow, 0].SetValue(10 * (i - 1));
                                worksheet[10 * (i - 1) + tempRow, 1].SetValue(dataType);
                                worksheet[10 * (i - 1) + tempRow, 2].SetValue(tem_workbook.Worksheets[0][i, 2].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 3].SetValue(date);

                                worksheet[10 * (i - 1) + tempRow, 4].SetValue(itemStrs[0].Replace("\"", ""));
                                worksheet[10 * (i - 1) + tempRow, 5].SetValue(itemStrs[1]);
                                worksheet[10 * (i - 1) + tempRow, 6].SetValue(itemStrs[2]);
                                worksheet[10 * (i - 1) + tempRow, 7].SetValue(itemStrs[3]);
                                worksheet[10 * (i - 1) + tempRow, 8].SetValue(itemStrs[4]);
                                worksheet[10 * (i - 1) + tempRow, 9].SetValue(DateTime.Now);

                                worksheet[10 * (i - 1) + tempRow, 10].SetValue(tem_workbook.Worksheets[0][i, 4].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 11].SetValue(tem_workbook.Worksheets[0][i, 5].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 12].SetValue(tem_workbook.Worksheets[0][i, 1].Value.ToString());
                            }
                            if (target_datatype == "飞机")
                            {
                                worksheet[10 * (i - 1) + tempRow, 0].SetValue(10 * (i - 1));
                                worksheet[10 * (i - 1) + tempRow, 1].SetValue(dataType);
                                worksheet[10 * (i - 1) + tempRow, 2].SetValue(tem_workbook.Worksheets[0][i, 2].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 3].SetValue(date);

                                worksheet[10 * (i - 1) + tempRow, 4].SetValue(itemStrs[0].Replace("\"", ""));
                                worksheet[10 * (i - 1) + tempRow, 5].SetValue(itemStrs[1]);
                                worksheet[10 * (i - 1) + tempRow, 6].SetValue(itemStrs[2]);
                            
                                worksheet[10 * (i - 1) + tempRow, 7].SetValue(DateTime.Now);

                                worksheet[10 * (i - 1) + tempRow, 8].SetValue(tem_workbook.Worksheets[0][i, 4].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 9].SetValue(tem_workbook.Worksheets[0][i, 5].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 10].SetValue(tem_workbook.Worksheets[0][i, 1].Value.ToString());
                            }
                            if (target_datatype == "火车")
                            {
                                worksheet[10 * (i - 1) + tempRow, 0].SetValue(10 * (i - 1));
                                worksheet[10 * (i - 1) + tempRow, 1].SetValue(dataType);
                                worksheet[10 * (i - 1) + tempRow, 2].SetValue(tem_workbook.Worksheets[0][i, 2].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 3].SetValue(date);

                                worksheet[10 * (i - 1) + tempRow, 4].SetValue(itemStrs[0].Replace("\"", ""));
                                worksheet[10 * (i - 1) + tempRow, 5].SetValue(itemStrs[1]);
                                worksheet[10 * (i - 1) + tempRow, 6].SetValue(itemStrs[2]);

                                worksheet[10 * (i - 1) + tempRow, 7].SetValue(DateTime.Now);

                                worksheet[10 * (i - 1) + tempRow, 8].SetValue(tem_workbook.Worksheets[0][i, 4].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 9].SetValue(tem_workbook.Worksheets[0][i, 5].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 10].SetValue(tem_workbook.Worksheets[0][i, 1].Value.ToString());
                            }
                            if (target_datatype == "汽车")
                            {
                                worksheet[10 * (i - 1) + tempRow, 0].SetValue(10 * (i - 1));
                                worksheet[10 * (i - 1) + tempRow, 1].SetValue(dataType);
                                worksheet[10 * (i - 1) + tempRow, 2].SetValue(tem_workbook.Worksheets[0][i, 2].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 3].SetValue(date);

                                worksheet[10 * (i - 1) + tempRow, 4].SetValue(itemStrs[0].Replace("\"", ""));
                                worksheet[10 * (i - 1) + tempRow, 5].SetValue(itemStrs[1]);
                                worksheet[10 * (i - 1) + tempRow, 6].SetValue(itemStrs[2]);

                                worksheet[10 * (i - 1) + tempRow, 7].SetValue(DateTime.Now);

                                worksheet[10 * (i - 1) + tempRow, 8].SetValue(tem_workbook.Worksheets[0][i, 4].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 9].SetValue(tem_workbook.Worksheets[0][i, 5].Value.ToString());
                                worksheet[10 * (i - 1) + tempRow, 10].SetValue(tem_workbook.Worksheets[0][i, 1].Value.ToString());
                            }

                            matchResult = matchResult.NextMatch();

                            tempRow++;
                        }

                        

                        

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }

                sum++;
                RunWithInoke(sum);

                
            }

            taskExecuted = true;
            workbook.SaveDocument(docPath);



            XtraMessageBox.Show("所有数据已经采集完成！");

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

        private void textEdit_file_EditValueChanged(object sender, EventArgs e)
        {
            dxValidationProvider1.Validate();
        }

        private void textEdit_fileName_EditValueChanged(object sender, EventArgs e)
        {
            dxValidationProvider1.Validate();
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
    }
}
