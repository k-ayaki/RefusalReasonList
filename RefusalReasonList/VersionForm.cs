using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RefusalReasonList
{
    public partial class VersionForm : Form
    {
        public VersionForm()
        {
            InitializeComponent();
            //自分自身のバージョン情報を取得する
            this.labelName.Text = Properties.Resources.Name;
            this.labelVersion.Text = Properties.Resources.Version;
            this.labelAuthor.Text = Properties.Resources.Author;
            this.Text = Properties.Resources.Text;
        }

        private void buttonOK_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
