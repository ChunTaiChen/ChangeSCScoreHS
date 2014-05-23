using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ChangeSCScoreHS
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void Main()
        {        
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["月考成績與小考轉換(新竹)"].Enable = true;
                K12.Presentation.NLDPanels.Student.ListPaneContexMenu["月考成績與小考轉換(新竹)"].Click += delegate
                {

                    if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                    {
                        ChangeDataForm cdf = new ChangeDataForm();
                        cdf.ShowDialog();
                    }
                    else
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("沒有選學生");                            
                    }
                };
        }
    }
}
