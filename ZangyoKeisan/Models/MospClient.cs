using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Livet;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net;
using System.Text.RegularExpressions;

namespace ZangyoKeisan.Models
{
    public class MospClient : NotificationObject
    {
        /// <summary>
        /// MOSPログインURL
        /// </summary>
        const string MOSP_URL = "https://www.nskint.co.jp/kintai/srv/";

        /// <summary>
        /// Excel勤怠簿ダウンロード先フォルダパス
        /// </summary>
        readonly string DOWNLOAD_FOLDER_PATH = Path.GetTempPath();

        /// <summary>
        /// Mospで、ログインするためにPOSTする値
        /// </summary>
        const string MOSP_CMD_LOGIN = "PF0020";


        #region DownloadStatus変更通知プロパティ
        private string _DownloadStatus;

        public string DownloadStatus
        {
            get
            { return _DownloadStatus; }
            set
            { 
                if (_DownloadStatus == value)
                    return;
                _DownloadStatus = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        /*
         * NotificationObjectはプロパティ変更通知の仕組みを実装したオブジェクトです。
         */
        public async Task<string> downloadExcel(string id ,string password, string year, string month)
        {
            
            var cookieContainer = new CookieContainer();

            var handler = new HttpClientHandler()
            {
                CookieContainer = cookieContainer
            };
            handler.UseCookies = true;

            var client = new HttpClient(handler);
            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36");
            client.DefaultRequestHeaders.Add("Accept-Language", "ja-JP");

            HttpResponseMessage response;

            var idPassword = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"txtUserId", id},           // ID
                {"txtPassWord", password},   // パスワード
                // 以下2行はログイン処理に必要な記述（実際にPOSTされた値を見ただけなので、内容は不明）
                {"cmd", MOSP_CMD_LOGIN},
                {"procSeq", "2" }
            });

            response = await client.PostAsync(MOSP_URL, idPassword);

            DownloadStatus = "ログインしました";

            string res = await response.Content.ReadAsStringAsync();
            string procSeq = getProcSeq(res);

            var kintaiList = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"cmd", "TM1100" },
                {"procSeq", procSeq },
                {"transferredMenuKey", "AttendanceList"}
            });
            var cookie = response.Headers.GetValues("Set-Cookie").First().Replace("JSESSIONID=", "").Replace("; Path=/kintai; Secure", "");

            Cookie sessioncookie = new Cookie("JSESSIONID", cookie.ToString());
            sessioncookie.Domain = "www.nskint.co.jp";
            sessioncookie.Path = "/kintai";
            sessioncookie.Secure = true;

            //cookieContainer.Add(new Uri(mospURL), sessioncookie);

            var kintaiListRes = await client.PostAsync(MOSP_URL, kintaiList);

            res = await kintaiListRes.Content.ReadAsStringAsync();
            procSeq = getProcSeq(res);

            var pageMoveParam = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"pltSelectYear", year},
                {"pltSelectMonth", month},
                { "cmd", "TM1102" },
                {"procSeq", procSeq }
            });

            kintaiListRes = await client.PostAsync(MOSP_URL, pageMoveParam);

            res = await kintaiListRes.Content.ReadAsStringAsync();
            procSeq = getProcSeq(res);

            var excelDownload = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"pltSelectYear", year},
                {"pltSelectMonth", month},
                {"cmd",  "TM1151"},
                {"procSeq", procSeq }
            });

            DownloadStatus = "ダウンロード中...";

            var fileDownload = await client.PostAsync(MOSP_URL, excelDownload);
            var fileStream = File.Create(DOWNLOAD_FOLDER_PATH + "text.xls");
            var httpStream = await fileDownload.Content.ReadAsStreamAsync();
            await httpStream.CopyToAsync(fileStream);
            fileStream.Flush();

            DownloadStatus = "ダウンロードが完了しました";

            string result = await response.Content.ReadAsStringAsync();
            return result;
        }

        /// <summary>
        /// "procSeq"（MospにパラメータをPOSTする際に必要となる）をHTMLから探して返す
        /// </summary>
        /// <param name="html">Mospから取得したHTML</param>
        /// <returns>procSeqの値</returns>
        private string getProcSeq(string html)
        {
            string procSeq = "";

            Regex regex = new Regex("var procSeq = \"(?<procSeqNum>.*)\";");
            Match match = regex.Match(html);

            if (match.Success == true)
            {
                procSeq = match.Groups["procSeqNum"].Value;
            }

            return procSeq;
        }
    }
}
