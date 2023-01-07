using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace RefusalReasonList
{
    public partial class rrRibbon
    {
        private void rrRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            using (RefusalReasonList fileList = new RefusalReasonList())
            {
                if (fileList.出願番号列の判定("出願番号") == true)
                {
                    fileList.書込み列の取得(@"@条文");
                    fileList.行数の取得();
                    fileList.m_wordConvert = false;
                    fileList.DoGetRefusalReason();
                    MessageBox.Show("出願番号列あり");
                }
                else
                {
                    MessageBox.Show("出願番号列なし");
                }
            }
        }

        private void buttonRR2Word_Click(object sender, RibbonControlEventArgs e)
        {
            using (RefusalReasonList rrList = new RefusalReasonList())
            {
                if (rrList.出願番号列の判定("出願番号") == true)
                {
                    rrList.書込み列の取得(@"@条文");
                    rrList.行数の取得();
                    rrList.m_wordConvert = true;
                    rrList.DoGetRefusalReason();
                    MessageBox.Show("出願番号列あり");
                }
                else
                {
                    MessageBox.Show("出願番号列なし");
                }
            }
        }

        private void buttonAccount_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.TaskPane.Visible = !Globals.ThisAddIn.TaskPane.Visible;
        }
    }
}
