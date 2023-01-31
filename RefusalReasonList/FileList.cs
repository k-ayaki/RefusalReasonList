using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using RefusalReasonList;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using AppLint01;
using Microsoft.Office.Interop.Excel;
using System.ComponentModel;
using System.Xml.Linq;
using JpoApi;
using System.IO;
using System.Text.RegularExpressions;

namespace RefusalReasonList
{
    class FileList : IDisposable
    {
        public const int ErrDiv0 = -2146826281; // #DIV0!
        public const int ErrNA = -2146826246;   // #N/A
        public const int ErrName = -2146826259; // #NAME
        public const int ErrNull = -2146826288; // #NULL!
        public const int ErrNum = -2146826252;  // #NUM!
        public const int ErrRef = -2146826265;  // #REF!
        public const int ErrValue = -2146826273;    // #VALUE!

        private bool disposedValue;

        private Excel.Worksheet m_activeSheet;
        private Excel.Workbook m_workbook;
        private string m_outPath;

        private int m_in_column;
        private int m_max_row;
        private int m_oaCount;

        private int m_out_column;
        private string m_appendColumn { get; set; }
        public FileList()
        {
            m_activeSheet = Globals.ThisAddIn.Application.ActiveSheet
              as Excel.Worksheet;
            m_workbook = Globals.ThisAddIn.Application.ActiveWorkbook as Excel.Workbook;
            m_outPath = m_workbook.Path + @"\" + Path.GetFileNameWithoutExtension(m_workbook.Name);
            m_in_column = 0;
            m_out_column = 0;
            m_max_row = 0;
            m_oaCount = 0;
            m_appendColumn = "条文";
        }
        public bool 出願番号列の判定(string fileNumberRow = "出願番号")
        {
            m_in_column = 0;

            for(int column = 1; column < 65535; column++)
            {
                object obj = m_activeSheet.Cells[1, column].value;
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
                    m_in_column = column;
                    return true;
                }
            }
            m_in_column = 0;

            return false;
        }
        public bool 行数の取得()
        {
            for (int row = 2; row < 65535; row++)
            {
                object obj = m_activeSheet.Cells[row, m_in_column].value;
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
                    m_max_row = row;
                }
            }
            if (m_max_row == 0) return false;
            return true;
        }
        public void 書込み列の取得(string appendColumn = @"@条文")
        {
            m_appendColumn = appendColumn;
            for(int col=1; col<65535; col++)
            {
                object obj = m_activeSheet.Cells[1, col].value;
                if (obj == null
                || obj.ToString() == "")
                {
                    break;
                }
                else if(obj.ToString().IndexOf(appendColumn)==0)
                {
                    Range range = m_activeSheet.Columns[col];
                    range.Delete(XlDeleteShiftDirection.xlShiftToLeft);
                    col--;
                }
            }
            for (int column = 1; column < 65535; column++)
            {
                object obj = m_activeSheet.Cells[1, column].value;
                if (obj == null
                || obj.ToString() == "")
                {
                    m_out_column = column;
                    break;
                }
            }
            return;
        }
        // 明細書のパラグラフをリストに格納
        public bool DoGetRefusalReason()
        {
            bool fRet = true;

            using (ProgressForm pd = new ProgressForm("拒絶理由通知",
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

            Account ac = new Account();
            AccessToken at = new AccessToken(ac.m_id, ac.m_password, ac.m_path);
            int errCode = 0;

            for (int row=2; row <= m_max_row; row++)
            {
                if (m_activeSheet.Cells[row, m_in_column].Value == null)
                {
                    break;
                }
                object objFileNumber = m_activeSheet.Cells[row, m_in_column].Value;
                string fileNumber = "";
                if (objFileNumber.GetType() == typeof(int))
                {
                    fileNumber = ((int)m_activeSheet.Cells[row, m_in_column].Value).ToString();
                }
                else if (objFileNumber.GetType() == typeof(string))
                {
                    fileNumber = ((string)m_activeSheet.Cells[row, m_in_column].Value);
                }
                if (fileNumber == null
                || fileNumber.Length == 0)
                {
                    break;
                }
                Regex rx1 = new Regex(@"^特願(?<year>[0-9]{4,4})-(?<no>[0-9]+)$", RegexOptions.None);
                Match w_match1 = rx1.Match(fileNumber);
                if (w_match1.Success)
                {
                    fileNumber = w_match1.Groups["year"].Value + (w_match1.Groups["no"].Value).PadLeft(6, '0');
                }
                System.Threading.Thread.Sleep(16);
                i++;
                //キャンセルされたか調べる
                if (bw.CancellationPending)
                {
                    //キャンセルされたとき
                    e.Cancel = true;
                    for (i = 0; i < m_oaCount; i++)
                    {
                        m_activeSheet.Cells[1, m_out_column + i].value = m_appendColumn + i.ToString();
                    }
                    return;
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
                Regex rx0 = new Regex(@"^[0-9]{10,10}$", RegexOptions.None);
                Match w_match0 = rx0.Match(fileNumber);
                if (w_match0.Success)
                {
                    AppDocContRefusalReason tj5 = new AppDocContRefusalReason(fileNumber, at.m_access_token.access_token);
                    if (tj5.m_error == tj5.e_CONTENT)
                    {
                        if(tj5.m_result.statusCode == "108")
                        {
                            m_activeSheet.Cells[row, m_out_column].value = @"";
                            m_activeSheet.Cells[row, m_out_column].Formula = @"#N/A";
                        }
                        else
                        {
                            m_activeSheet.Cells[row, m_out_column].value = @"";
                            m_activeSheet.Cells[row, m_out_column].Formula = @"#REF!";
                        }
                    }
                    else
                    if (tj5.m_error == tj5.e_NONE && tj5.m_files != null)
                    {
                        int j = 0;
                        foreach (string f in tj5.m_files)
                        {
                            if (m_appendColumn == "w条文")
                            {
                                if (System.IO.Directory.Exists(m_outPath) == false)
                                {
                                    Directory.CreateDirectory(m_outPath);
                                }
                                Xml2Word xml2Word = new Xml2Word(f, fileNumber, m_outPath);
                            }
                            notice_pat_exam npe = new notice_pat_exam(f);
                            if (npe != null)
                            {
                                m_activeSheet.Cells[row, m_out_column + j].Formula = "";
                                m_activeSheet.Cells[row, m_out_column + j].value = npe.provisions();
                                j++;
                                m_activeSheet.Cells[row, m_out_column + j].Formula = "";
                                m_activeSheet.Cells[row, m_out_column + j].value = npe.refusal_sentences();
                                j++;
                                m_activeSheet.Cells[row, m_out_column + j].Formula = "";
                                m_activeSheet.Cells[row, m_out_column + j].value = npe.xString("//jp:drafting-body");
                                j++;
                            }
                            else
                            {
                                m_activeSheet.Cells[row, m_out_column + j].Formula = "";
                                m_activeSheet.Cells[row, m_out_column + j].value = "";
                                j++;
                            }
                            if(m_oaCount < j)
                            {
                                m_oaCount = j;
                                m_activeSheet.Cells[1, m_out_column + j].value = m_appendColumn + j.ToString();
                            }
                        }
                    }
                    else
                    {
                        m_activeSheet.Cells[row, m_out_column].value = @"";
                        m_activeSheet.Cells[row, m_out_column].Formula = @"#REF!";
                    }
                }
                else
                {
                    m_activeSheet.Cells[row, m_out_column].value = @"";
                    m_activeSheet.Cells[row, m_out_column].Formula = @"#REF!";
                }
            }
            for(i=0; i<m_oaCount; i++)
            {
                m_activeSheet.Cells[1, m_out_column + i].value = m_appendColumn + i.ToString();
            }
            //結果を設定する
            e.Result = 0;
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
