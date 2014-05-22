using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Data;
using System.Data;

namespace ChangeSCScoreHS
{
    public partial class ChangeDataForm : FISCA.Presentation.Controls.BaseForm
    {
        Dictionary<string, string> _ExamDict;

        public ChangeDataForm()
        {
            InitializeComponent();
            _ExamDict = new Dictionary<string, string>();

        }

        private void ChangeDataForm_Load(object sender, EventArgs e)
        {
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            // 取得考試
            GetExam();
            foreach (string name in _ExamDict.Keys)
            {
                cbxTotal.Items.Add(name);
                cbxSScore.Items.Add(name);
                cbxAScore.Items.Add(name);                
            }
        }

        private void GetExam()
        {
            _ExamDict.Clear();
            string query1 = "select id,exam_name from exam;";
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query1);
            foreach (DataRow dr in dt.Rows)
                _ExamDict.Add(dr["exam_name"].ToString(), dr["id"].ToString());
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {

        }  

    }
}
