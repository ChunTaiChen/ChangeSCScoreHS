using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Data;
using System.Xml.Linq;

namespace ChangeSCScoreHS
{
    public partial class ChangeDataForm : FISCA.Presentation.Controls.BaseForm
    {
        Dictionary<string, string> _ExamDict;

        List<string> _SCAttendIDList;

        public ChangeDataForm()
        {
            InitializeComponent();
            _ExamDict = new Dictionary<string, string>();
            _SCAttendIDList = new List<string>();

        }

        private void ChangeDataForm_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;

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
            if (string.IsNullOrEmpty(cbxTotal.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("沒有合併評量名稱");
                return;
            }

            if (string.IsNullOrEmpty(cbxSScore .Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("沒有定期評量名稱");
                return;
            }

            if (string.IsNullOrEmpty(cbxAScore.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("沒有平時評量名稱");
                return;
            }

            btnRun.Enabled = false;

            List<string> StudenIDList = K12.Presentation.NLDPanels.Student.SelectedSource;
            int SchoolYear=iptSchoolYear.Value;
            int Semester=iptSemester.Value;

            // 處理小考名稱與小考成績
            string oldASExamID = "";
            string oldSSExamID = "";
            string newTSExamID = "";
            
            if (_ExamDict.ContainsKey(cbxAScore.Text))
                oldASExamID = _ExamDict[cbxAScore.Text];

            if (_ExamDict.ContainsKey(cbxSScore.Text))
                oldSSExamID = _ExamDict[cbxSScore.Text];

            if (_ExamDict.ContainsKey(cbxTotal.Text))
                newTSExamID = _ExamDict[cbxTotal.Text];

            // 修改課程上小考項目與學生修課上的小考成績
            EditCourseSCAttendExtesions(oldASExamID, newTSExamID, StudenIDList, SchoolYear, Semester);

            
            List<string> delSCTakeIDss = new List<string>();
            List<string> delSCTakeIDas = new List<string>();
            // Score
            Dictionary<string, string> SSDict = new Dictionary<string, string>();
            // A Score
            Dictionary<string, string> SADict = new Dictionary<string, string>();

            string queryKey = string.Join(",", _SCAttendIDList.ToArray());

            // 讀取舊定期成績
            string querySS = @"select id,ref_sc_attend_id,extension from sce_take where ref_sc_attend_id in("+queryKey+") and ref_exam_id="+oldSSExamID;
            QueryHelper qhss = new QueryHelper();
            DataTable dtss = qhss.Select(querySS);
            foreach (DataRow drss in dtss.Rows)
            {
                delSCTakeIDss.Add(drss["id"].ToString());
                string scid = drss["ref_sc_attend_id"].ToString();
                string ext = drss["extension"].ToString();
                string score="";
                if (!string.IsNullOrEmpty(ext))
                {
                    XElement elm = null;
                    try
                    {
                        elm = XElement.Parse(ext);
                    }
                    catch (Exception ex)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("讀取定期成績發生錯誤" + ex.Message);
                    }

                    if (elm != null)
                    {
                        if (elm.Element("Score") != null)
                            score = elm.Element("Score").Value;
                    }
                }
                SSDict.Add(scid, score);
            }
            
            // 讀取舊平時成績
            string queryAS = @"select id,ref_sc_attend_id,extension from sce_take where ref_sc_attend_id in(" + queryKey + ") and ref_exam_id=" + oldASExamID;
            QueryHelper qhas = new QueryHelper();
            DataTable dtas = qhas.Select(queryAS);
            foreach (DataRow dras in dtas.Rows)
            {
                delSCTakeIDas.Add(dras["id"].ToString());
                string scid = dras["ref_sc_attend_id"].ToString();
                string ext = dras["extension"].ToString();
                string ascore = "";
                if (!string.IsNullOrEmpty(ext))
                {
                    XElement elm = null;
                    try
                    {
                        elm = XElement.Parse(ext);
                    }
                    catch (Exception ex)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("讀取平時成績發生錯誤:" + ex.Message);
                    }

                    if (elm != null)
                    {
                        if (elm.Element("AssignmentScore") != null)
                            ascore = elm.Element("AssignmentScore").Value;
                    }
                }
                SADict.Add(scid, ascore);
            }

            // 寫入定期/平時成績
            StringBuilder sbInsert = new StringBuilder();
            foreach (string scID in _SCAttendIDList)
            {
                string ss = "", sa = "";
                // 定期
                if (SSDict.ContainsKey(scID))
                    ss = SSDict[scID];

                // 平時
                if (SADict.ContainsKey(scID))
                    sa = SADict[scID];

                XElement elm = new XElement("Extension");
                elm.SetElementValue("Score", ss);
                elm.SetElementValue("AssignmentScore", sa);
                elm.SetElementValue("Text", "");

                string str = @"insert into sce_take(ref_sc_attend_id,ref_exam_id,score,extension) values(" + scID + "," + newTSExamID + ",0,'" + elm.ToString() + "');";
                sbInsert.AppendLine(str);
            }
                try
                {
                    if (sbInsert.Length > 0)
                    {
                        UpdateHelper uhIns = new UpdateHelper();
                        uhIns.Execute(sbInsert.ToString());
                    }
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("資料新增錯誤：" + ex.Message);
                }
            
            FISCA.Presentation.Controls.MsgBox.Show("資料轉換新增成功.");
            btnRun.Enabled = true;
        }

        /// <summary>
        /// 修改課程與修課小考 Extesions
        /// </summary>
        /// <param name="newExamID"></param>
        private void EditCourseSCAttendExtesions(string oldExamID,string newExamID,List<string> StudIDList,int SchoolYear,int Semester)
        {
            StringBuilder sb1 = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();

            string idList=string.Join(",",StudIDList.ToArray());
            // 修改課程
            string query1 = @"select distinct course.id as course_id,course.extensions as course_extensions from course inner join sc_attend on course.id=sc_attend.ref_course_id where course.school_year="+SchoolYear+" and course.semester="+Semester+" and sc_attend.ref_student_id in("+idList+")";
            QueryHelper qh1 = new QueryHelper();
            DataTable dt1 = qh1.Select(query1);
            foreach (DataRow dr1 in dt1.Rows)
            {
                string id = dr1["course_id"].ToString();
                string ext = dr1["course_extensions"].ToString();
                if (!string.IsNullOrEmpty(ext))
                {
                    XElement elm = null;
                    try
                    {
                        elm = XElement.Parse(ext);
                    }
                    catch (Exception ex)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("處理課程小考發生錯誤："+ex.Message);
                    }

                    if (elm != null)
                    {
                        foreach (XElement elms1 in elm.Elements("Extension"))
                        {
                            if (elms1.Attribute("Name") != null && elms1.Attribute("Name").Value == "GradeItem")
                            {
                                foreach (XElement elms2 in elms1.Elements("GradeItem"))
                                {
                                    foreach (XElement elms3 in elms2.Elements("Item"))
                                    {
                                        // 更換 ExamID
                                        if (elms3.Attribute("ExamID").Value == oldExamID)
                                            elms3.Attribute("ExamID").SetValue(newExamID);
                                    }                            
                                }                            
                            }
                        }

                        string updateStr = @"update course set extensions='"+ext.ToString()+ "' where id="+id+";";
                        sb1.AppendLine(updateStr);
                    }
                }
            }

            if (sb1.Length > 0)
            {
                UpdateHelper uh1 = new UpdateHelper();
                uh1.Execute(sb1.ToString());
            }

            _SCAttendIDList.Clear();

            // 修改修課
            string query2 = @"select sc_attend.id as sc_attend_id, sc_attend.extensions as sc_attend_extensions from course inner join sc_attend on course.id=sc_attend.ref_course_id where course.school_year="+SchoolYear+" and course.semester="+Semester+" and sc_attend.ref_student_id in("+idList+")";
            QueryHelper qh2 = new QueryHelper();
            DataTable dt2 = qh1.Select(query2);
            foreach (DataRow dr2 in dt2.Rows)
            {
                string id = dr2["sc_attend_id"].ToString();
                _SCAttendIDList.Add(id);

                string ext = dr2["sc_attend_extensions"].ToString();
                if (!string.IsNullOrEmpty(ext))
                {
                    XElement elm = null;
                    try
                    {
                        elm = XElement.Parse(ext);
                    }
                    catch (Exception ex)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("處理修課小考發生錯誤：" + ex.Message);
                    }

                    if (elm != null)
                    {
                        foreach (XElement elms1 in elm.Elements("Extension"))
                        {
                            if (elms1.Attribute("Name") != null && elms1.Attribute("Name").Value == "GradeBook")
                            {
                                foreach (XElement elms2 in elms1.Elements("Exam"))
                                {
                                        // 更換 ExamID
                                    if (elms2.Attribute("ExamID").Value == oldExamID)
                                        elms2.Attribute("ExamID").SetValue(newExamID);
                                }
                            }
                        }

                        string updateStr = @"update sc_attend set extensions='" + ext.ToString() + "' where id=" + id + ";";
                        sb2.AppendLine(updateStr);
                    }
                }
            }

            if (sb2.Length > 0)
            {
                UpdateHelper uh2 = new UpdateHelper();
                uh2.Execute(sb2.ToString());
            }


        }
    }
}
