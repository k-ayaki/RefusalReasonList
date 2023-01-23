using JpoApi;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using JpoApi;

namespace RefusalReasonList
{
    public partial class UserControlAccount : UserControl
    {
        public UserControlAccount()
        {
            InitializeComponent();
            using (Account ac = new Account())
            {
                this.textBoxID.Text = ac.m_id;
                this.textBoxPassword.Text = ac.m_password;
                this.textBoxPath.Text = ac.m_path;
                this.textBoxCacheEffective.Text = ac.m_cacheEffective.ToString();
            }
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            using (Account ac = new Account())
            {
                ac.m_id = this.textBoxID.Text;
                ac.m_password = this.textBoxPassword.Text;
                ac.m_path = this.textBoxPath.Text;
                ac.m_cacheEffective = Int32.Parse(this.textBoxCacheEffective.Text);
            }
        }

        private void buttonGetToken_Click(object sender, EventArgs e)
        {
            AccessToken ac = new AccessToken(this.textBoxID.Text, this.textBoxPassword.Text, this.textBoxPath.Text);
            if (ac.m_access_token.access_token.Length > 0)
            {
                MessageBox.Show(ac.m_access_token.access_token, "トークン取得");
            } else
            {
                MessageBox.Show("トークンが取得できませんでした。", "トークン取得");
            }
        }
    }
}
