using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Drawing;
using Microsoft.VisualBasic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;
using System.Web;

namespace RefusalReasonList
{
    public class notice_pat_exam : IDisposable
    {
        private bool disposedValue;
        public int m_error { get; set; }
        public readonly int e_NONE = 0x00000000;
        public readonly int e_NETWORK = 0x00000001;
        public readonly int e_SERVER = 0x00000002;
        public readonly int e_TIMEOVER = 0x00000004;
        public readonly int e_CONTENT = 0x00000008;
        public readonly int e_ZIPFILE = 0x00000010;
        public readonly int e_CACHE = 0x00000020;

        public string m_xmlPath { get; set; }
        public string m_dirName { get; set; }

        public XmlDocument m_xDoc { get; set; }

        public XmlNamespaceManager m_xmlNsManager { get; set; }

        public XmlNode node_notice_pat_exam;
        public notice_pat_exam(string xmlPath)
        {
            m_xDoc = new XmlDocument();
            node_notice_pat_exam = null;
            try
            {
                m_xmlNsManager = new XmlNamespaceManager(m_xDoc.NameTable);
                m_xmlNsManager.AddNamespace("jp", "http://www.jpo.go.jp");

                m_dirName = System.IO.Path.GetDirectoryName(xmlPath);

                m_xDoc.XmlResolver = null;
                m_xmlPath = xmlPath;
                if (File.Exists(xmlPath))
                {
                    m_xDoc.Load(xmlPath);
                    m_error = e_NONE;
                    return;
                }
                else
                {
                    m_error = e_CACHE;
                    return;
                }
            }
            catch (System.IO.FileNotFoundException ex)
            {
                m_error = e_CACHE;
                return;
            }
            catch (System.UnauthorizedAccessException ex)
            {
                m_error = e_CACHE;
                return;
            }
        }

        public string provisions()
        {
            string provisionsDetail = "";

            string drafting_body = xString("//jp:drafting-body");

            // xmlのdrafting-body を改行ごとに区分
            string[] del = { "\r\n" };
            string[] sentences = drafting_body.Split(del, StringSplitOptions.None);

            // 条文列挙部分の接続
            for (int i = 0; i + 1 < sentences.Length; i++)
            {
                if (sentences[i].IndexOf(@"<br />") == 36
                && sentences[i].Substring(35, 1) != "。")
                {
                    sentences[i] = sentences[i].Substring(0, 36);
                    sentences[i] += sentences[i + 1];
                    sentences[i] = sentences[i].Replace("　", "");
                    sentences[i + 1] = "";
                }
                if (sentences[i].IndexOf("　記") >= 0)
                {
                    break;
                }
            }

            foreach (string sentence in sentences)
            {
                string sentence2 = sentence.Replace(@"<br />", "");
                sentence2 = sentence2.Replace("　", "");
                sentence2 = Strings.StrConv(sentence2, VbStrConv.Wide, 0x411);

                // 条文列挙部分の抽出
                if (Regex.IsMatch(sentence2, "この出願(の|は[、，]?)(下記の請求項|請求項[０-９，、]+|特許請求の範囲|発明の詳細な説明|特許請求の範囲又は発明の詳細な説明|明細書|下記)")
                || Regex.IsMatch(sentence2, "その出願の日前の(日本語)?特許出願であって、")
                || Regex.IsMatch(sentence2, "[０-９]+年[０-９]+月[０-９]+日付けでした手続補正は[、，]")
                || sentence2.IndexOf("特許を受けることができない") >= 0
                || (sentence2.IndexOf("要件を") >= 0 && sentence2.IndexOf("満たしていない") >= 0))
                {
                    // 括弧部分の除去
                    string sentence3 = "";
                    foreach (Match match0 in Regex.Matches(sentence2, "(?<lv0>[^（]*)?(?<lv1>（[^）]*）)?"))
                    {
                        sentence3 += match0.Groups["lv0"].Value;
                    }

                    string prov1 = "";
                    string lv1 = "";
                    string lv2 = "";
                    string lv3 = "";
                    foreach (Match match in Regex.Matches(sentence3, "(?<prov1>(?<lv1>特許法第?(?<lv11>[０-９]+条(の[０-９]+)?))?(?<lv2>第[０-９]+項(柱書)?)?(?<lv3>(?<lv31>第[０-９]+)(、|，)?(?<lv32>[０-９]+)?号)?)(および|及び|または|又は|亦は|叉は|ならびに?|並びに?|、|，|[のにで](規定|該当))"))
                    {
                        prov1 = match.Groups["prov1"].Value;
                        if (prov1.Length > 0)
                        {
                            if (match.Groups["lv1"].Value.Length > 0)
                            {
                                lv1 = "特許法第" + match.Groups["lv11"].Value;
                                lv2 = match.Groups["lv2"].Value;
                            }
                            else
                            if (lv1.Length > 0)
                            {
                                if (match.Groups["lv2"].Value.Length > 0)
                                {
                                    lv2 = match.Groups["lv2"].Value;
                                }
                            }
                            else
                            {
                                continue;
                            }
                            if (match.Groups["lv32"].Value.Length > 0)
                            {
                                lv3 = match.Groups["lv31"].Value + "号";
                            }
                            else
                            {
                                lv3 = match.Groups["lv3"].Value;
                            }
                            prov1 = lv1 + lv2 + lv3;
                            if (lv1 != "特許法第４１条")
                            {
                                if (provisionsDetail.IndexOf(prov1) == -1)
                                {
                                    if (provisionsDetail.Length > 0)
                                    {
                                        provisionsDetail += ",";
                                    }
                                    provisionsDetail += prov1;
                                }
                            }
                            if (match.Groups["lv32"].Value.Length > 0
                            && lv1 != "特許法第４１条")
                            {
                                lv3 = "第" + match.Groups["lv32"].Value + "号";
                                prov1 = lv1 + lv2 + lv3;
                                if (provisionsDetail.IndexOf(prov1) == -1)
                                {
                                    if (provisionsDetail.Length > 0)
                                    {
                                        provisionsDetail += ",";
                                    }
                                    provisionsDetail += prov1;
                                }
                            }
                        }
                    }
                }
            }
            return provisionsDetail;
        }
        public string refusal_sentences()
        {
            string refusals = "";

            string drafting_body = xString("//jp:drafting-body");
            // xmlのdrafting-body を改行ごとに区分
            string[] del = { "\r\n" };
            string[] sentences = drafting_body.Split(del, StringSplitOptions.None);

            // 条文列挙部分の接続
            for (int i = 0; i + 1 < sentences.Length; i++)
            {
                if (sentences[i].IndexOf(@"<br />") == 36
                && sentences[i].Substring(35,1) != "。")
                {
                    sentences[i] = sentences[i].Substring(0, 36);
                    sentences[i] += sentences[i + 1];
                    sentences[i] = sentences[i].Replace("　", "");
                    sentences[i + 1] = "";
                }
                if (sentences[i].IndexOf("　記") >= 0 )
                {
                    break;
                }
            }
            foreach (string sentence in sentences)
            {
                string sentence2 = sentence.Replace(@"<br />", "");
                sentence2 = sentence2.Replace("　", "");
                sentence2 = Strings.StrConv(sentence2, VbStrConv.Wide, 0x411);

                // 条文列挙部分の抽出
                if (Regex.IsMatch(sentence2, "この出願(の|は[、，]?)(下記の請求項|請求項[０-９，、]+|特許請求の範囲|発明の詳細な説明|特許請求の範囲又は発明の詳細な説明|明細書|下記)")
                || Regex.IsMatch(sentence2, "その出願の日前の(日本語)?特許出願であって、")
                || Regex.IsMatch(sentence2, "[０-９]+年[０-９]+月[０-９]+日付けでした手続補正は[、，]")
                || sentence2.IndexOf("特許を受けることができない") >= 0
                || (sentence2.IndexOf("要件を") >= 0 && sentence2.IndexOf("満たしていない") >= 0))
                {
                    // 括弧部分の除去
                    string sentence3 = "";
                    foreach (Match match0 in Regex.Matches(sentence2, "(?<lv0>[^（]*)?(?<lv1>（[^）]*）)?"))
                    {
                        sentence3 += match0.Groups["lv0"].Value;
                    }
                    if (refusals.IndexOf(sentence3) == -1)
                    {
                        if (refusals.Length > 0)
                        {
                            refusals += "\r\n";
                        }
                        refusals += sentence3;
                    }
                }
            }
            return refusals;
        }

        public string xString(string xpath)
        {
            XmlNode node = m_xDoc.SelectSingleNode(xpath, m_xmlNsManager);
            if (node != null)
            {
                return p2html(node);
            }
            return "";
        }

        private string p2html(XmlNode nodeP)
        {
            string wHtmlbody = "<p>";
            XmlNodeList nodeList = nodeP.ChildNodes;
            foreach (XmlNode node in nodeList)
            {
                if (node.LocalName == "img")
                {
                    wHtmlbody += "\r\n" + node_img(node);
                }
                else
                if (node.LocalName == "chemistry")
                {
                    wHtmlbody += "\r\n【化" + Strings.StrConv(node.Attributes["num"].Value, VbStrConv.Wide, 0x411) + "】\r\n";
                    wHtmlbody += p2html(node);
                }
                else
                if (node.LocalName == "tables")
                {
                    wHtmlbody += "\r\n【表" + Strings.StrConv(node.Attributes["num"].Value, VbStrConv.Wide, 0x411) + "】\r\n";
                    wHtmlbody += p2html(node);
                }
                else
                if (node.LocalName == "maths")
                {
                    wHtmlbody += "\r\n【数" + Strings.StrConv(node.Attributes["num"].Value, VbStrConv.Wide, 0x411) + "】\r\n";
                    wHtmlbody += p2html(node);
                }
                else
                if (node.LocalName == "#text")
                {
                    wHtmlbody += HttpUtility.HtmlEncode(node.OuterXml);
                }
                else
                {
                    wHtmlbody += node.OuterXml;
                }
            }
            wHtmlbody += "</p>\r\n";
            return wHtmlbody;
        }
        private string node_img(XmlNode node)
        {
            string wHtmlbody = "<p>";
            int height = (int)(3.777 * double.Parse(node.Attributes["he"].Value));
            int width = (int)(3.777 * double.Parse(node.Attributes["wi"].Value));
            string w_src_png = Path.GetFileNameWithoutExtension(node.Attributes["file"].Value) + ".png";
            string w_src1 = m_dirName + @"\" + w_src_png;

            string w_src0 = m_dirName + @"\" + node.Attributes["file"].Value;
            System.Drawing.Image img = System.Drawing.Bitmap.FromFile(w_src0);
            img.Save(w_src1, System.Drawing.Imaging.ImageFormat.Png);
            wHtmlbody += "<img height=" + height.ToString() + " width=" + width.ToString() + " src=\"" + w_src_png + "\"></p>\r\n";
            return wHtmlbody;
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
        // ~notice_pat_exam()
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