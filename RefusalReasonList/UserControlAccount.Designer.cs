namespace RefusalReasonList
{
    partial class UserControlAccount
    {
        /// <summary> 
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary> 
        /// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
        /// コード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.textBoxID = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxPassword = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.buttonOK = new System.Windows.Forms.Button();
            this.textBoxPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxCacheEffective = new System.Windows.Forms.TextBox();
            this.labelCache = new System.Windows.Forms.Label();
            this.buttonGetToken = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // textBoxID
            // 
            this.textBoxID.Location = new System.Drawing.Point(107, 34);
            this.textBoxID.Name = "textBoxID";
            this.textBoxID.Size = new System.Drawing.Size(226, 25);
            this.textBoxID.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(45, 39);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 18);
            this.label1.TabIndex = 2;
            this.label1.Text = "ID  ";
            // 
            // textBoxPassword
            // 
            this.textBoxPassword.Location = new System.Drawing.Point(107, 84);
            this.textBoxPassword.Name = "textBoxPassword";
            this.textBoxPassword.Size = new System.Drawing.Size(226, 25);
            this.textBoxPassword.TabIndex = 6;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 94);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 18);
            this.label2.TabIndex = 7;
            this.label2.Text = "パスワード";
            // 
            // buttonOK
            // 
            this.buttonOK.Location = new System.Drawing.Point(213, 273);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(119, 30);
            this.buttonOK.TabIndex = 8;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.buttonOK_Click);
            // 
            // textBoxPath
            // 
            this.textBoxPath.Location = new System.Drawing.Point(107, 141);
            this.textBoxPath.Name = "textBoxPath";
            this.textBoxPath.Size = new System.Drawing.Size(226, 25);
            this.textBoxPath.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 152);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(90, 18);
            this.label3.TabIndex = 10;
            this.label3.Text = "トークンパス";
            // 
            // textBoxCacheEffective
            // 
            this.textBoxCacheEffective.Location = new System.Drawing.Point(189, 200);
            this.textBoxCacheEffective.Name = "textBoxCacheEffective";
            this.textBoxCacheEffective.Size = new System.Drawing.Size(144, 25);
            this.textBoxCacheEffective.TabIndex = 11;
            this.textBoxCacheEffective.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // labelCache
            // 
            this.labelCache.AutoSize = true;
            this.labelCache.Location = new System.Drawing.Point(3, 207);
            this.labelCache.Name = "labelCache";
            this.labelCache.Size = new System.Drawing.Size(126, 18);
            this.labelCache.TabIndex = 12;
            this.labelCache.Text = "キャッシュ有効日";
            // 
            // buttonGetToken
            // 
            this.buttonGetToken.Location = new System.Drawing.Point(16, 273);
            this.buttonGetToken.Name = "buttonGetToken";
            this.buttonGetToken.Size = new System.Drawing.Size(130, 30);
            this.buttonGetToken.TabIndex = 13;
            this.buttonGetToken.Text = "トークン取得";
            this.buttonGetToken.UseVisualStyleBackColor = true;
            this.buttonGetToken.Click += new System.EventHandler(this.buttonGetToken_Click);
            // 
            // UserControlAccount
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonGetToken);
            this.Controls.Add(this.labelCache);
            this.Controls.Add(this.textBoxCacheEffective);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.textBoxPath);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.textBoxPassword);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.textBoxID);
            this.Margin = new System.Windows.Forms.Padding(5, 4, 5, 4);
            this.Name = "UserControlAccount";
            this.Size = new System.Drawing.Size(352, 1010);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox textBoxID;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox textBoxPassword;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.TextBox textBoxPath;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxCacheEffective;
        private System.Windows.Forms.Label labelCache;
        private System.Windows.Forms.Button buttonGetToken;
    }
}
