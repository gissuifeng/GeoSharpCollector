using DevExpress.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GeoSharp2018.ToolForms
{
    public partial class ExtentLibForm : DevExpress.XtraEditors.XtraForm
    {

        private Workbook workbook;

        private Worksheet worksheet;

        private int selRowIndex = -1;

        private bool isEditing = false;

        private MainForm mainform;

        public ExtentLibForm(MainForm mainform)
        {
            InitializeComponent();

            this.mainform = mainform;
        }

        private void ExtentLibForm_Load(object sender, EventArgs e)
        {
            workbook = new Workbook();

            workbook.LoadDocument(Application.StartupPath + @"\data\poi_extent.xlsx");

            worksheet = workbook.Worksheets[0];

            initialGrid();

            txtEnableFun(false);

            btn_editStop.Enabled = false;
            btn_newRow.Enabled = false;
            
            btn_editSave.Enabled = false;
            btn_deleteRow.Enabled = false;

            btn_editStart.Enabled = true;
            btn_setDefault.Enabled = true;
        }

        private void initialGrid()
        {
            DataTable dt = new DataTable();

            Range range = worksheet.GetUsedRange();

            int rowCount = range.RowCount;
            int colCount = range.ColumnCount;

            DataRow dr = null;

            for (int row = 0; row < rowCount; row++)
            {
                dr = dt.NewRow();

                for (int col = 0; col < colCount; col++)
                {
                    if (row == 0)
                    {

                        dt.Columns.Add(worksheet[row, col].Value.ToString());
                    }
                    else
                    {
                        //MessageBox.Show(worksheet[row, col].Value.ToString());


                        dr[col] = worksheet[row, col].Value.ToString();
                    }
                }

                if (row != 0)
                {
                    dt.Rows.Add(dr);
                }
            }

            gridControl1.DataSource = dt;
        }

        public void txtEnableFun(bool sta)
        {
            txt_code.Enabled = false;
            txt_name.Enabled = sta;
            txt_firstXY.Enabled = sta;
            txt_secondXY.Enabled = sta;
        }

        public void txtClearFun()
        {
            txt_code.EditValue = "";
            txt_name.EditValue = "";
            txt_firstXY.EditValue = ""; ;
            txt_secondXY.EditValue = ""; ;
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_editStop_Click(object sender, EventArgs e)
        {
            initialGrid();
            txtEnableFun(false);
            txtClearFun();

            isEditing = false;

            btn_editStop.Enabled = false;
            btn_newRow.Enabled = false;
            btn_editSave.Enabled = false;
            btn_deleteRow.Enabled = false;

            btn_editStart.Enabled = true;
            btn_setDefault.Enabled = true;
        }

        private void btn_editSave_Click(object sender, EventArgs e)
        {
            bool bb = true; // dxValidationProvider1.Validate();

            if (bb)
            {
                Range range = worksheet.GetUsedRange();

                int rowCount = range.RowCount;

                string lng01 = txt_firstXY.EditValue.ToString().Split(',')[0];
                string lat01 = txt_firstXY.EditValue.ToString().Split(',')[1];

                string lng02 = txt_secondXY.EditValue.ToString().Split(',')[0];
                string lat02 = txt_secondXY.EditValue.ToString().Split(',')[1];

                worksheet[selRowIndex + 1, 1].SetValue(txt_name.EditValue);
                worksheet[selRowIndex + 1, 2].SetValue(lng01);
                worksheet[selRowIndex + 1, 3].SetValue(lat01);
                worksheet[selRowIndex + 1, 4].SetValue(lng02);
                worksheet[selRowIndex + 1, 5].SetValue(lat02);
                worksheet[selRowIndex + 1, 6].SetValue(memoEdit_note.EditValue);


                workbook.SaveDocument(Application.StartupPath + @"\data\poi_extent.xlsx");

                initialGrid();
            }
        }

        private void btn_newRow_Click(object sender, EventArgs e)
        {
            bool bb = true; // dxValidationProvider1.Validate();

            if (bb)
            {
                Range range = worksheet.GetUsedRange();

                int rowCount = range.RowCount;

                string code = string.Format("{0}{1}{2}{3}{4}{5}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, DateTime.Now.Hour, DateTime.Now.Minute, DateTime.Now.Second);
                txt_code.EditValue = code;

                string lng01 = txt_firstXY.EditValue.ToString().Split(',')[0];
                string lat01 = txt_firstXY.EditValue.ToString().Split(',')[1];

                string lng02 = txt_secondXY.EditValue.ToString().Split(',')[0];
                string lat02 = txt_secondXY.EditValue.ToString().Split(',')[1];

                worksheet[rowCount, 0].SetValue(txt_code.EditValue);
                worksheet[rowCount, 1].SetValue(txt_name.EditValue);
                worksheet[rowCount, 2].SetValue(lng01);
                worksheet[rowCount, 3].SetValue(lat01);
                worksheet[rowCount, 4].SetValue(lng02);
                worksheet[rowCount, 5].SetValue(lat02);
                worksheet[rowCount, 6].SetValue(memoEdit_note.EditValue);


                workbook.SaveDocument(Application.StartupPath + @"\data\poi_extent.xlsx");

                initialGrid();
            }
        }

        private void btn_deleteRow_Click(object sender, EventArgs e)
        {
            worksheet.Rows.Remove(selRowIndex + 1);

            initialGrid();
        }

        private void btn_setDefault_Click(object sender, EventArgs e)
        {
            //currentCode.Text = worksheet[selRowIndex + 1, 0].Value.ToString();

            double lx = Convert.ToDouble(worksheet[selRowIndex + 1, 2].Value.ToString());
            double ly = Convert.ToDouble(worksheet[selRowIndex + 1, 3].Value.ToString());
            double rx = Convert.ToDouble(worksheet[selRowIndex + 1, 4].Value.ToString());
            double ry = Convert.ToDouble(worksheet[selRowIndex + 1, 5].Value.ToString());

            mainform.extentCoord.SetCoor(lx, ly, rx, ry);
        }

        private void btn_editStart_Click(object sender, EventArgs e)
        {
            initialGrid();
            txtEnableFun(true);
            txtClearFun();

            isEditing = true;

            btn_editStop.Enabled = true;
            btn_newRow.Enabled = true;
            btn_editSave.Enabled = true;
            btn_deleteRow.Enabled = true;

            btn_editStart.Enabled = false;
            btn_setDefault.Enabled = false;
        }

        private void gridView1_FocusedRowChanged(object sender, DevExpress.XtraGrid.Views.Base.FocusedRowChangedEventArgs e)
        {
            if (isEditing)
            {
                DataRow dr = gridView1.GetFocusedDataRow();

                txt_code.EditValue = dr[0];
                txt_name.EditValue = dr[1];
                txt_firstXY.EditValue = string.Format("{0},{1}", dr[2], dr[3]);
                txt_secondXY.EditValue = string.Format("{0},{1}", dr[4], dr[5]);

                memoEdit_note.EditValue = dr[6];


            }
            selRowIndex = e.FocusedRowHandle;
        }
    }
}
