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

        /// <summary>
        /// Mospで、「勤怠一覧」画面に遷移するためにPOSTする値
        /// </summary>
        const string MOSP_CMD_KINTAI_ITIRAN = "TM1100";

        /// <summary>
        /// Mospで、「勤怠一覧」画面において指定した年月の勤怠記録画面を開くためにPOSTする値
        /// </summary>
        const string MOSP_CMD_PAGE_MOVE = "TM1102";

        /// <summary>
        /// Mospで、勤怠簿をダウンロードするためにPOSTする値
        /// </summary>
        const string MOSP_CMD_EXCEL_DOWNLOAD = "TM1151";

        /*
         * NotificationObjectはプロパティ変更通知の仕組みを実装したオブジェクトです。
         */
        /// <summary>
        /// Mospから勤怠簿をダウンロードする
        /// </summary>
        /// <param name="progress">ダウンロード状況を示す</param>
        /// <param name="id">MospのログインID</param>
        /// <param name="password">Mospのログインパスワード</param>
        /// <param name="year">ダウンロード対象の年</param>
        /// <param name="month">ダウンロード対象の月</param>
        /// <returns></returns>
        public async Task<string> downloadExcel(IProgress<ProgressInfo> progress, string id ,string password, DateTime targetDate)
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

            // ログイン処理
            HttpResponseMessage loginResult = await loginMosp(client, id, password);

            string loginResultHTML = await loginResult.Content.ReadAsStringAsync();
            string procSeq = getProcSeq(loginResultHTML);

            // ログイン判定
            if (loginResult == null || procSeq == "")
            {
                progress.Report(new ProgressInfo("ログインに失敗しました。", true));

                return "";
            }
            else
            {
                progress.Report(new ProgressInfo("ログインしました"));
            }

            // セッションID取得
            var cookie = loginResult.Headers.GetValues("Set-Cookie").First().Replace("JSESSIONID=", "").Replace("; Path=/kintai; Secure", "");

            Cookie sessioncookie = new Cookie("JSESSIONID", cookie.ToString());
            sessioncookie.Domain = "www.nskint.co.jp";
            sessioncookie.Path = "/kintai";
            sessioncookie.Secure = true;        

            // 勤怠一覧画面に遷移
            procSeq = await moveKintaiItiran(client, procSeq);

            // 遷移に失敗した場合
            if (procSeq == "")
            {
                return "";
            }

            // ダウンロード対象年月の勤怠一覧画面に遷移
            // （遷移しないと、次のダウンロード処理のパラメータに年月を渡しても反映されない）
            procSeq = await moveToKintaiItiranTargetDate(client, targetDate, procSeq);

            // Excel勤怠簿ダウンロード
            progress.Report(new ProgressInfo("ダウンロード中..."));

            string kintaiboFilePath = await downloadKintaibo(client, targetDate, procSeq);

            // 勤怠簿ダウンロード判定
            if (kintaiboFilePath == "")
            {
                progress.Report(new ProgressInfo("勤怠簿のダウンロードに失敗しました", true));
                return "";
            }
            
            progress.Report(new ProgressInfo("ダウンロードが完了しました"));

            return kintaiboFilePath;
        }

        /// <summary>
        /// "procSeq"（MospにパラメータをPOSTする際に必要となる）をHTMLから探して返す
        /// </summary>
        /// <param name="html">Mospから取得したHTML</param>
        /// <returns>procSeqの値（取得できなかった場合は""（空文字列）を返す</returns>
        private string getProcSeq(string html)
        {
            string procSeq = "";

            // procSeqはJavaScriptの変数にセットされている
            Regex regex = new Regex("var procSeq = \"(?<procSeqNum>.*)\";");
            Match match = regex.Match(html);

            if (match.Success == true)
            {
                procSeq = match.Groups["procSeqNum"].Value;
            }

            return procSeq;
        }

        /// <summary>
        /// Mospにログインする
        /// </summary>
        /// <param name="client">HttpClientオブジェクト</param>
        /// <param name="id">MospのログインID</param>
        /// <param name="password">Mospのログインパスワード</param>
        /// <returns>ログイン後ページのHttpResponseMessage（ログインに失敗した場合はnull））</returns>
        private async Task<HttpResponseMessage> loginMosp(HttpClient client, string id, string password)
        {
            var postParam = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"txtUserId", id},           // ID
                {"txtPassWord", password},   // パスワード
                // 以下2行はログイン処理に必要な記述（実際にPOSTされた値を見ただけなので、内容は不明）
                {"cmd", MOSP_CMD_LOGIN},
                {"procSeq", "0" }
            });

            HttpResponseMessage mospResponse = await client.PostAsync(MOSP_URL, postParam);

            string res = await mospResponse.Content.ReadAsStringAsync();

            // ログイン判定

            // ログインに失敗した場合、エラーコードを含むHTMLが返される
            string loginErrorCode = "PFW9111";
            if (res.ToString().Contains(loginErrorCode))
            {
                return null;
            }

            return mospResponse;
        }

        /// <summary>
        /// Mospの勤怠一覧画面に遷移する
        /// </summary>
        /// <param name="client">HttpClientオブジェクト</param>
        /// <param name="procSeq">POSTする際に使用するprocSeq</param>
        /// <returns>遷移後のprocSeq（遷移に失敗した場合は""（空文字列））</returns>
        private async Task<string> moveKintaiItiran(HttpClient client, string procSeq)
        {
            var postParam = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"cmd", MOSP_CMD_KINTAI_ITIRAN },
                {"procSeq", procSeq },
                {"transferredMenuKey", "AttendanceList"}
            });

            HttpResponseMessage kintaiListRes = await client.PostAsync(MOSP_URL, postParam);

            var res = await kintaiListRes.Content.ReadAsStringAsync();

            procSeq = getProcSeq(res);

            return procSeq;
        }

        /// <summary>
        /// Mospの勤怠簿ダウンロード対象年月の勤怠一覧画面に遷移する
        /// </summary>
        /// <param name="client">HttpClientオブジェクト</param>
        /// <param name="targetDate">遷移対象の年月を示すDateTimeオブジェクト（日時分秒は無視される）</param>
        /// <param name="procSeq">POSTする際に使用するprocSeq</param>
        /// <returns>遷移後のprocSeq（遷移に失敗した場合は""（空文字列））</returns>
        private async Task<string> moveToKintaiItiranTargetDate(HttpClient client, DateTime targetDate, string procSeq)
        {
            var postParam = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"pltSelectYear", targetDate.Year.ToString()},
                {"pltSelectMonth", targetDate.Month.ToString()},
                {"cmd", MOSP_CMD_PAGE_MOVE },
                {"procSeq", procSeq }
            });

            HttpResponseMessage mospResponse = await client.PostAsync(MOSP_URL, postParam);

            var res = await mospResponse.Content.ReadAsStringAsync();
            procSeq = getProcSeq(res);

            return procSeq;
        }

        /// <summary>
        /// 勤怠簿をダウンロードする
        /// </summary>
        /// <param name="client">HttpClientオブジェクト</param>
        /// <param name="targetDate">ダウンロード対象年月を示すDateTimeオブジェクト（日時分秒は無視される）</param>
        /// <param name="procSeq">POSTする際に使用するprocSeq</param>
        /// <returns>ダウンロードした勤怠簿のパス（ダウンロードに失敗した場合は""（空文字列））</returns>
        private async Task<string> downloadKintaibo(HttpClient client, DateTime targetDate, string procSeq)
        {
            var postParam = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"pltSelectYear", targetDate.Year.ToString()},
                {"pltSelectMonth", targetDate.Year.ToString()},
                {"cmd",  MOSP_CMD_EXCEL_DOWNLOAD},
                {"procSeq", procSeq }
            });

            var fileDownload = await client.PostAsync(MOSP_URL, postParam);
            string kintaiboFilePath = DOWNLOAD_FOLDER_PATH + "mosp_kintaibo_" + targetDate.Year.ToString() + "_" + targetDate.Month.ToString() + ".xls";
            var fileStream = File.Create(kintaiboFilePath);
            var httpStream = await fileDownload.Content.ReadAsStreamAsync();
            await httpStream.CopyToAsync(fileStream);
            fileStream.Flush();

            // ダウンロード結果判定

            // ファイルサイズチェック
            // （何らかの理由でダウンロードできなかった場合）
            if (fileStream.Length == 0)
            {
                
                return "";
            }

            // 次の判定でファイルを開くため、ここで閉じる
            // （上の処理でファイルサイズを取得しているため、これより前ではCloseできない）
            fileStream.Close();

            // ファイル内容チェック
            using (StreamReader reader = new StreamReader(kintaiboFilePath, Encoding.GetEncoding("Shift_JIS")))
            {
                string text = reader.ReadToEnd();

                // Mosp側でエラーが発生した場合、ログインページに戻される。
                // この場合、エラーページのHTMLをダウンロードしてしまうため、チェックする
                if (text.Contains("<html>"))
                {
                    return "";
                }
            }

            return kintaiboFilePath;
        }
    }

    /// <summary>
    /// ダウンロード状況
    /// </summary>
    public class ProgressInfo
    {
        public ProgressInfo(string message, bool isError = false)
        {
            Message = message;
            IsError = isError;
        }

        /// <summary>
        /// 現在の状態を示すメッセージ
        /// </summary>
        public string Message { get; private set; }

        /// <summary>
        /// エラーが発生したか（trueならエラー発生）
        /// </summary>
        public bool IsError { get; private set; }
    }
}
