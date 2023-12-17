using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using RefusalReasonList;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Xml.Linq;
using JpoApi;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic;

namespace RefusalReasonList
{
    class RefusalReasonList : IDisposable
    {
        private bool disposedValue;

        public const int ErrDiv0 = -2146826281; // #DIV0!
        public const int ErrNA = -2146826246;   // #N/A
        public const int ErrName = -2146826259; // #NAME
        public const int ErrNull = -2146826288; // #NULL!
        public const int ErrNum = -2146826252;  // #NUM!
        public const int ErrRef = -2146826265;  // #REF!
        public const int ErrValue = -2146826273;    // #VALUE!

        private Excel.Worksheet m_activeSheet { get; set; }
        public  Excel.Workbook m_workbook { get; set; }
        private string m_outPath { get; set; }
        private string m_relativePath { get; set; }
        private int m_in_column { get; set; }
        private int m_max_row { get; set; }
        private int m_oaCount { get; set; }
        private int m_out_column { get; set; }
        private string m_appendColumn { get; set; }

        private List<Xml2Word> m_xml2WordList = new List<Xml2Word>();
        private bool m_isProvisions { get; set; }
        public RefusalReasonList()
        {
            this.m_activeSheet = Globals.ThisAddIn.Application.ActiveSheet
              as Excel.Worksheet;
            this.m_workbook = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
            this.m_outPath = this.m_workbook.Path + @"\" + Path.GetFileNameWithoutExtension(this.m_workbook.Name);
            this.m_relativePath = @".\" + Path.GetFileNameWithoutExtension(this.m_workbook.Name);
            this.m_in_column = 0;
            this.m_out_column = 0;
            this.m_max_row = 0;
            this.m_oaCount = 0;
            this.m_appendColumn = @"@審査記録";
            this.m_isProvisions = false;
        }
        public bool 出願番号列の判定(string fileNumberRow = "出願番号")
        {
            this.m_in_column = 0;

            for(int column = 1; column < 65535; column++)
            {
                object obj = this.m_activeSheet.Cells[1, column].value;
                if (obj == null)
                {
                    break;
                }
                else if(obj.ToString() == "")
                {
                    break;
                }
                if (obj.ToString().IndexOf(fileNumberRow)==0)
                {
                    this.m_in_column = column;
                    return true;
                }
            }
            this.m_in_column = 0;
            return false;
        }
        public bool 行数の取得()
        {
            for (int row = 2; row < 65535; row++)
            {
                object obj = this.m_activeSheet.Cells[row, this.m_in_column].value;
                if (obj == null)
                {
                    ;
                }
                else if (obj.ToString() == "")
                {
                    ;
                }
                else if (obj.ToString().Length > 0)
                {
                    this.m_max_row = row;
                }
            }
            if (this.m_max_row == 0) return false;
            return true;
        }
        public void 書込み列の取得(string appendColumn = @"@審査記録")
        {
            this.m_appendColumn = appendColumn;
            for(int col=1; col<65535; col++)
            {
                object obj = this.m_activeSheet.Cells[1, col].value;
                if (obj == null || obj.ToString() == "")
                {
                    break;
                }
                else if(obj.ToString().IndexOf(appendColumn)==0)
                {
                    Range range = this.m_activeSheet.Columns[col];
                    range.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    col--;
                }
            }
            for (int column = 1; column < 65535; column++)
            {
                object obj = this.m_activeSheet.Cells[1, column].value;
                if (obj == null || obj.ToString() == "")
                {
                    this.m_out_column = column;
                    break;
                }
            }
            return;
        }
        // 明細書のパラグラフをリストに格納
        public bool DoGetRefusalReason(bool isProvisions)
        {
            bool fRet = true;
            this.m_isProvisions = isProvisions;

            using (ProgressForm pd = new ProgressForm("審査記録",
                    new DoWorkEventHandler(ProgressDialog_Support_DoGetRefusalReason),
                    0))
            {
                //進行状況ダイアログを表示する
                DialogResult result = pd.ShowDialog();
                //結果を取得する
                if (result == DialogResult.Cancel)
                {
                    MessageBox.Show("キャンセルされました");
                    fRet = false;
                }
                else if (result == DialogResult.Abort)
                {
                    //エラー情報を取得する
                    Exception ex = pd.Error;
                    MessageBox.Show("エラー: " + ex.Message);
                    fRet = false;
                }
                else if (result == DialogResult.OK)
                {
                    //結果を取得する
                    int stopTime = (int)pd.Result;
                    fRet = true;
                }
                //後始末
                pd.Dispose();
            }
            return fRet;
        }
        // DoMAイベントハンドラ
        // 形態素解析
        private void ProgressDialog_Support_DoGetRefusalReason(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker bw = (BackgroundWorker)sender;
            DateTime currentDate;
            //パラメータを取得する
            int stopTime = (int)e.Argument;

            int i = 0;
            currentDate = DateTime.Now;
            long lastTick = currentDate.Ticks-1200;
            long currTick;
            bool isAppDocContRefusalReasonDecision = true;

            using (Account account = new Account())
            {
                AccessToken token = new AccessToken(account.m_id, account.m_password, account.m_path);
                for (int row = 2; row <= this.m_max_row; row++)
                {
                    string fileNumber = 出願番号取得(this.m_activeSheet.Cells[row, this.m_in_column].Value);
                    Regex rx0 = new Regex(@"^[0-9]{10,10}$", RegexOptions.None);
                    Match w_match0 = rx0.Match(fileNumber);
                    if (w_match0.Success == false)
                    {
                        this.m_activeSheet.Cells[row, this.m_out_column].value = @"";
                        this.m_activeSheet.Cells[row, this.m_out_column].Formula = @"#N/A";
                        continue;
                    }
                    System.Threading.Thread.Sleep(16);
                    i++;
                    //キャンセルされたか調べる
                    if (bw.CancellationPending)
                    {
                        //キャンセルされたとき
                        e.Cancel = true;
                        break;
                    }
                    currentDate = DateTime.Now;
                    currTick = currentDate.Ticks;
                    if (currTick - lastTick > 600 * 10000)
                    {
                        //指定された時間待機する
                        System.Threading.Thread.Sleep(16);

                        int percent = i * 100 / m_max_row;
                        //bw.ReportProgress(percent, i.ToString());
                        bw.ReportProgress(percent, fileNumber);
                        lastTick = currTick;
                    }
                    this.m_xml2WordList = new List<Xml2Word>();

                    token.refresh();
                    if (token.m_access_token.access_token.Length == 0)
                    {
                        MessageBox.Show("アカウントが正しく入力されていません。");
                        break;
                    }
                    AppDocContRefusalReasonDecision tj1 = new AppDocContRefusalReasonDecision(fileNumber, token.m_access_token.access_token);
                    if (tj1.m_error == tj1.e_NONE && tj1.m_files != null)
                    {
                        foreach (string f in tj1.m_files)
                        {
                            Xml2Word xml2word = new Xml2Word(f, fileNumber, this.m_outPath);
                            if (xml2word != null)
                            {
                                m_xml2WordList.Add(xml2word);
                            }
                        }
                    }
                    else
                    if (tj1.m_error == tj1.e_SERVER || tj1.m_error == tj1.e_NETWORK)
                    {
                        MessageBox.Show(tj1.m_result.errorMessage);
                        break;
                    }
                    //キャンセルされたか調べる
                    if (bw.CancellationPending)
                    {
                        //キャンセルされたとき
                        e.Cancel = true;
                        break;
                    }
                    currentDate = DateTime.Now;
                    currTick = currentDate.Ticks;
                    if (currTick - lastTick > 600 * 10000)
                    {
                        //指定された時間待機する
                        System.Threading.Thread.Sleep(16);

                        int percent = i * 100 / m_max_row;
                        //bw.ReportProgress(percent, i.ToString());
                        bw.ReportProgress(percent, fileNumber);
                        lastTick = currTick;
                    }
                    token.refresh();
                    if (token.m_access_token.access_token.Length == 0)
                    {
                        MessageBox.Show("アカウントが正しく入力されていません。");
                        break;
                    }
                    AppDocContOpinionAmendment tj2 = new AppDocContOpinionAmendment(fileNumber, token.m_access_token.access_token);
                    if (tj2.m_error == tj2.e_NONE && tj2.m_files != null)
                    {
                        foreach (string f in tj2.m_files)
                        {
                            Xml2Word xml2word = new Xml2Word(f, fileNumber, this.m_outPath);
                            if (xml2word != null)
                            {
                                this.m_xml2WordList.Add(xml2word);
                            }
                        }
                    }
                    else
                    if (tj2.m_error == tj2.e_SERVER || tj2.m_error == tj2.e_NETWORK)
                    {
                        MessageBox.Show(tj2.m_result.errorMessage);
                        break;
                    }
                    this.m_xml2WordList.Sort((a, b) => string.Compare(a.m_Date, b.m_Date));
                    int j = 0;
                    foreach (Xml2Word xml2word in this.m_xml2WordList)
                    {
                        if (xml2word.m_DocumentName.Length > 0)
                        {
                            string szName = xml2word.m_DocumentName;
                            if (this.m_isProvisions)
                            {
                                if (xml2word.m_provisions.Length > 0)
                                {
                                    szName = xml2word.m_provisions;
                                }
                            }
                            if (xml2word.m_wordFilePath.Length > 0)
                            {
                                this.m_activeSheet.Cells[row, this.m_out_column + j].Formula = "";
                                this.m_activeSheet.Hyperlinks.Add(this.m_activeSheet.Cells[row, this.m_out_column + j], this.m_relativePath + @"\" + Path.GetFileName(xml2word.m_outFileName), Type.Missing, "", szName);
                                j++;
                            }
                            else
                            {
                                this.m_activeSheet.Cells[row, this.m_out_column + j].value = szName;
                                this.m_activeSheet.Cells[row, this.m_out_column + j].Formula = "";
                                j++;
                            }
                        }
                        else
                        {
                            this.m_activeSheet.Cells[row, this.m_out_column + j].value = "";
                            this.m_activeSheet.Cells[row, this.m_out_column + j].Formula = @"#NULL!";
                            j++;
                        }
                        if (this.m_oaCount < j)
                        {
                            this.m_oaCount = j;
                            this.m_activeSheet.Cells[1, this.m_out_column + j - 1].value = this.m_appendColumn + j.ToString();
                        }
                    }
                }
                for (i = 0; i < this.m_oaCount; i++)
                {
                    this.m_activeSheet.Cells[1, this.m_out_column + i].value = this.m_appendColumn + (i + 1).ToString();
                }
                //結果を設定する
                e.Result = 0;
                token.Dispose();
                account.Dispose();
            }
        }

        private string 出願番号取得(object objFileNumber)
        {
            string fileNumber = "";
            if (objFileNumber != null)
            {
                if (objFileNumber.GetType() == typeof(int))
                {
                    fileNumber = ((int)objFileNumber).ToString();
                }
                else if (objFileNumber.GetType() == typeof(double))
                {
                    fileNumber = ((double)objFileNumber).ToString();
                }
                else if (objFileNumber.GetType() == typeof(float))
                {
                    fileNumber = ((float)objFileNumber).ToString();
                }
                else if (objFileNumber.GetType() == typeof(string))
                {
                    fileNumber = ((string)objFileNumber);
                } else
                {
                    fileNumber = ((int)objFileNumber).ToString();
                }
                fileNumber = Strings.StrConv(fileNumber, VbStrConv.Narrow, 0x411);
                Regex rx1 = new Regex(@"^特願(?<year>[0-9]{4,4})-(?<no>[0-9]+)$", RegexOptions.None);
                Match w_match1 = rx1.Match(fileNumber);
                if (w_match1.Success)
                {
                    fileNumber = w_match1.Groups["year"].Value + (w_match1.Groups["no"].Value).PadLeft(6, '0');
                    return fileNumber;
                }
                Regex rx3 = new Regex(@"^特願平(?<year>[0-9]+)-(?<no>[0-9]+)$", RegexOptions.None);
                Match w_match3 = rx3.Match(fileNumber);
                if (w_match3.Success)
                {
                    int gengo = int.Parse(w_match3.Groups["year"].Value) + 1988;
                    fileNumber = gengo.ToString() + (w_match3.Groups["no"].Value).PadLeft(6, '0');
                    return fileNumber;
                }
                Regex rx2 = new Regex(@"^(?<year>[0-9]{4,4})-(?<no>[0-9]+)$", RegexOptions.None);
                Match w_match2 = rx2.Match(fileNumber);
                if (w_match2.Success)
                {
                    fileNumber = w_match2.Groups["year"].Value + (w_match2.Groups["no"].Value).PadLeft(6, '0');
                    return fileNumber;
                }
                Regex rx4 = new Regex(@"^[0-9]{10,10}$", RegexOptions.None);
                Match w_match4 = rx4.Match(fileNumber);
                {
                    return fileNumber;
                }
            }
            return fileNumber;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    // TODO: マネージド状態を破棄します (マネージド オブジェクト)
                }

                // TODO: アンマネージド リソース (アンマネージド オブジェクト) を解放し、ファイナライザーをオーバーライドします
                // TODO: 大きなフィールドを null に設定します
                disposedValue = true;
            }
        }

        // // TODO: 'Dispose(bool disposing)' にアンマネージド リソースを解放するコードが含まれる場合にのみ、ファイナライザーをオーバーライドします
        // ~FileList()
        // {
        //     // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
        //     Dispose(disposing: false);
        // }

        public void Dispose()
        {
            // このコードを変更しないでください。クリーンアップ コードを 'Dispose(bool disposing)' メソッドに記述します
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }
}
