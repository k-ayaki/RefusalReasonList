using JpoApi;
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

        private void buttonRR2Word_Click(object sender, RibbonControlEventArgs e)
        {
            using (Account ac = new Account())
            {
                using (AccessToken at = new AccessToken(ac.m_id, ac.m_password, ac.m_path))
                {
                    if (at.m_access_token.access_token.Length == 0)
                    {
                        MessageBox.Show("アカウントが正しく設定されていません。");
                        return;
                    }
                    at.Dispose();
                }
                ac.Dispose();
            }

            using (RefusalReasonList rrList = new RefusalReasonList())
            {
                if (rrList.m_workbook.Path.Length > 0)
                {
                    if (rrList.出願番号列の判定("出願番号") == true)
                    {
                        rrList.書込み列の取得(@"@審査記録");
                        rrList.行数の取得();
                        rrList.DoGetRefusalReason(true);
                        MessageBox.Show("出願番号列あり");
                    }
                    else
                    {
                        MessageBox.Show("出願番号列なし");
                    }
                } 
                else
                {
                    MessageBox.Show("ワークシートを保存してください。保存先に拒絶理由通知のWordファイルが生成されます。");
                }
            }
        }

        private void buttonAccount_Click(object sender, RibbonControlEventArgs e)
        {
            if (Globals.ThisAddIn.TaskPane != null)
            {
                Globals.ThisAddIn.TaskPane.Visible = !Globals.ThisAddIn.TaskPane.Visible;
            }
        }

        private void buttonVersion_Click(object sender, RibbonControlEventArgs e)
        {
            VersionForm f = new VersionForm();
            f.StartPosition = FormStartPosition.CenterScreen;
            f.Show();
        }
    }
}
