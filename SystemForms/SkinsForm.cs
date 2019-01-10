using DevExpress.LookAndFeel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GeoSharp2018.SystemForms
{
    public partial class SkinsForm : DevExpress.XtraEditors.XtraForm
    {
        private DefaultLookAndFeel defaultLookAndFeel;

        List<SkinsLine> skinLineList;

        bool isFirst = false;

        public SkinsForm(DefaultLookAndFeel defaultLookAndFeel)
        {
            InitializeComponent();

            this.defaultLookAndFeel = defaultLookAndFeel;
        }

        private void btn_ok_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SkinsForm_Load(object sender, EventArgs e)
        {
            skinLineList = new List<SkinsLine>();

            comboBox_skins.Properties.Items.Clear();

            string path = string.Format(@"{0}\{1}", Application.StartupPath, "skins.xml");
            StreamReader sr = new StreamReader(path, Encoding.Default);
            String line;
            string[] strParts = new string[3];
            while((line = sr.ReadLine()) != null)
            {
                strParts = line.Split(',');
                comboBox_skins.Properties.Items.Add(strParts[1]);

                if (strParts[2] == "1")
                {
                    comboBox_skins.SelectedIndex = Convert.ToInt32(strParts[0]);
                }

                skinLineList.Add(new SkinsLine(line));
            }


            sr.Close();



            //defaultLookAndFeel.LookAndFeel.SkinName = "Valentine";

            isFirst = true;
        }

        private void comboBox_skins_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(isFirst)
            {
                int selIdx = comboBox_skins.SelectedIndex;

                foreach (var item in skinLineList)
                {
                    item.sflag = 0;
                }

                skinLineList[selIdx].sflag = 1;

                string path = string.Format(@"{0}\{1}", Application.StartupPath, "skins.xml");

                FileStream fs = new FileStream(path, FileMode.Create);
                StreamWriter sw = new StreamWriter(fs);
                foreach (var item1 in skinLineList)
                {
                    sw.WriteLine(string.Format($"{item1.sid},{item1.stype},{item1.sflag}"));
                }

                sw.Flush();

                sw.Close();
                fs.Close();

                defaultLookAndFeel.LookAndFeel.SkinName = skinLineList[selIdx].stype;
            }
        }

        public class SkinsLine
        {
            public int sid;
            public string stype;
            public int sflag;
            public SkinsLine(string line)
            {
                string[] elem = line.Split(',');

                sid = Convert.ToInt32(elem[0]);
                stype = elem[1];
                sflag = Convert.ToInt32(elem[2]);

            }
        }

        
    }
}
