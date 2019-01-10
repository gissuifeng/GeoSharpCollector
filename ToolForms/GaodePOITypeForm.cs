using DevExpress.Spreadsheet;
using DevExpress.XtraEditors;
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
    public partial class GaodePOITypeForm : DevExpress.XtraEditors.XtraForm
    {
        public GaodePOIGetForm gaodePoiForm;

        public GaodePOITypeForm(GaodePOIGetForm gaodePoiForm)
        {
            InitializeComponent();

            this.gaodePoiForm = gaodePoiForm;
        }

        private void GaodePOITypeForm_Load(object sender, EventArgs e)
        {
            Workbook workbook = new Workbook();
            workbook.LoadDocument(Application.StartupPath + @"\data\gaode_poi_code.xlsx");

            Worksheet worksheet = workbook.Worksheets[0];

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

        private void btn_select_Click(object sender, EventArgs e)
        {
            bool bb = dxValidationProvider1.Validate();

            if (bb)
            {
                this.DialogResult = System.Windows.Forms.DialogResult.OK;

                this.Close();
            }
        }

        private void gridControl1_DoubleClick(object sender, EventArgs e)
        {
            string s = gridView1.FocusedColumn.FieldName;

            if (s == "分类码")
            {
                XtraMessageBox.Show("不能选择<编号>列");
            }
            else
            {
                string cellVal = gridView1.GetFocusedValue().ToString();

                DataRow dr = gridView1.GetFocusedDataRow();

                string poiTypeCode = dr[0].ToString();

                int fcol = -1;
                for (int i = 0; i < 4; i++)
                {
                    if (dr[i].ToString() == cellVal)
                    {
                        fcol = i;
                    }
                }

                if (fcol == 0 || fcol == 1)
                {
                    poiTypeCode = poiTypeCode.Substring(0, 2);
                }
                if (fcol == 2)
                {
                    poiTypeCode = poiTypeCode.Substring(0, 4);
                }
                if (fcol == 3)
                {
                    poiTypeCode = poiTypeCode.Substring(0, 6);
                }

                gaodePoiForm.poiTypeCode = poiTypeCode;
                gaodePoiForm.poiKeyword = cellVal;

                tb_result.EditValue = cellVal;
            }
        }
    }
}
