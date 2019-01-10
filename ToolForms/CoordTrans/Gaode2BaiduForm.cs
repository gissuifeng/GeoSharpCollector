using DevExpress.Spreadsheet;
using DevExpress.XtraSpreadsheet;
using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using GeoSharp2018.UtilClass;

namespace GeoSharp2018.ToolForms.CoordTrans
{
    public partial class Gaode2BaiduForm : DevExpress.XtraEditors.XtraForm
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
        /// 当前工作表
        /// </summary>
        private Worksheet worksheetOriginal;
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


        public Gaode2BaiduForm(MainForm mainform, SpreadsheetControl spreadsheetControl)
        {
            InitializeComponent();

            this.mainform = mainform;
            this.spreadsheetControl = spreadsheetControl;
        }

        private void btn_inputOpen_Click(object sender, EventArgs e)
        {
            isFirstLoad = true;

            comboBox_inputSheet.Properties.Items.Clear();
            comboBox_inputLng.Properties.Items.Clear();
            comboBox_inputLat.Properties.Items.Clear();

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
                    comboBox_inputLng.Properties.Items.Add(worksheetOriginal[0, j].Value);
                    comboBox_inputLat.Properties.Items.Add(worksheetOriginal[0, j].Value);
                }

                comboBox_inputLng.SelectedIndex = 0;
                comboBox_inputLat.SelectedIndex = 0;

                isFirstLoad = false;
            }
        }
        int colCount;
        private void btn_ok_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                try
                {
                    col = comboBox_inputLng.SelectedIndex;
                    col1 = comboBox_inputLat.SelectedIndex;

                    colCount = worksheetOriginal.GetDataRange().ColumnCount;

                    progressBarControl1.Properties.Maximum = worksheetOriginal.GetDataRange().RowCount;

                    worksheetOriginal[0, colCount].SetValue(comboBox_outputLng.EditValue.ToString());
                    worksheetOriginal[0, colCount + 1].SetValue(comboBox_outputLat.EditValue.ToString());



                    Thread thread = new Thread(fun);
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
            double ox = 0.0, oy = 0.0;
            double dx = 0.0, dy = 0.0;
            Gps gps;
            for (int i = 1; i < worksheetOriginal.GetDataRange().RowCount; i++)
            {
                //MessageBox.Show(worksheet[i, col].Value.ToString());
                string str_ox = worksheetOriginal[i, col].Value.ToString();
                string str_oy = worksheetOriginal[i, col1].Value.ToString();
                if (string.IsNullOrEmpty(str_ox) || string.IsNullOrEmpty(str_oy))
                {
                    str_ox = "0";
                    str_oy = "0";
                }
                ox = Convert.ToDouble(str_ox);
                oy = Convert.ToDouble(str_oy);

                gps = CoordUtil.gcj02_To_Bd09(oy, ox);

                dx = gps.getWgLon();
                dy = gps.getWgLat();

                worksheetOriginal[i, colCount].SetValue(dx);
                worksheetOriginal[i, colCount + 1].SetValue(dy);

                RunWithInoke(i + 1);
            }

            workbookOriginal.SaveDocument(docPathOriginal);
            taskExecuted = true;

            XtraMessageBox.Show("转换完成！");
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
                    spreadsheetControl.LoadDocument(docPathOriginal);
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

        private void Gaode2BaiduForm_Load(object sender, EventArgs e)
        {
            comboBox_outputLng.EditValue = "baidu_lng";
            comboBox_outputLat.EditValue = "baidu_lat";
        }

        private void textEdit_inputFile_EditValueChanged(object sender, EventArgs e)
        {

            comboBox_inputSheet.Properties.Items.Clear();

            if (!isFirstLoad)
            {
                dxValidationProvider1.Validate();

                for (int i = 0; i < workbookOriginal.Worksheets.Count; i++)
                {
                    comboBox_inputSheet.Properties.Items.Add(workbookOriginal.Worksheets[i].Name);
                }

                comboBox_inputSheet.SelectedIndex = 0;
            }
        }

        private void comboBox_inputSheet_EditValueChanged(object sender, EventArgs e)
        {
            comboBox_inputLng.Properties.Items.Clear();
            comboBox_inputLat.Properties.Items.Clear();

            if (!isFirstLoad)
            {

                dxValidationProvider1.Validate();

                string targetSheetName = comboBox_inputSheet.SelectedText;

                if (targetSheetName != "")
                {
                    worksheetOriginal = workbookOriginal.Worksheets[targetSheetName];

                    for (int j = 0; j < worksheetOriginal.GetDataRange().ColumnCount; j++)
                    {
                        comboBox_inputLng.Properties.Items.Add(worksheetOriginal[0, j].Value);
                        comboBox_inputLat.Properties.Items.Add(worksheetOriginal[0, j].Value);
                    }

                    comboBox_inputLng.SelectedIndex = 0;
                    comboBox_inputLat.SelectedIndex = 0;
                }
            }

        }

        private void comboBox_inputLng_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isFirstLoad)
            {
                dxValidationProvider1.Validate();
            }

        }

        private void comboBox_inputLat_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!isFirstLoad)
            {
                dxValidationProvider1.Validate();
            }
        }


        private void comboBox_outputLat_EditValueChanged_1(object sender, EventArgs e)
        {
            if (!isFirstLoad)
            {
                dxValidationProvider1.Validate();
            }
        }

        private void comboBox_outputLng_EditValueChanged_1(object sender, EventArgs e)
        {
            if (!isFirstLoad)
            {
                dxValidationProvider1.Validate();
            }
        }
    }
}
