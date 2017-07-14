using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Livet;
using System.Net.Http;
using System.Threading.Tasks;
using System.Net;

namespace ZangyoKeisan.Models
{
    public class MospClient : NotificationObject
    {
        /*
         * NotificationObjectはプロパティ変更通知の仕組みを実装したオブジェクトです。
         */
         public async Task<string> downloadExcel()
        {
            
            var cookieContainer = new CookieContainer();

            string mospURL = "https://www.nskint.co.jp/kintai/srv/";

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
                {"txtUserId", ""},           // ID
                {"txtPassWord", ""},        // パスワード
                // 以下2行はログイン処理に必要な記述（実際にPOSTされた値を見ただけなので、内容は不明）
                {"cmd", "PF0020"},
                {"procSeq", "2" }
            });

            response = await client.PostAsync(mospURL, idPassword);
            
            var kintaiList = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"cmd", "TM1100" },
                {"procSeq", "0" },
                {"transferredMenuKey", "AttendanceList"}
            });
            var cookie = response.Headers.GetValues("Set-Cookie").First().Replace("JSESSIONID=", "").Replace("; Path=/kintai; Secure", "");

            Cookie sessioncookie = new Cookie("JSESSIONID", cookie.ToString());
            sessioncookie.Domain = "www.nskint.co.jp";
            sessioncookie.Path = "/kintai";
            sessioncookie.Secure = true;

            //cookieContainer.Add(new Uri(mospURL), sessioncookie);

            var kintaiListRes = await client.PostAsync(mospURL, kintaiList);

            var excelDownload = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                {"pltSelectYear", "2017"},
                {"pltSelectMonth", "5"},
                {"cmd",  "TM1151"},
                {"procSeq", "1" }
            });


            var fileDownload = await client.PostAsync(mospURL, excelDownload);
            var fileStream = File.Create(Path.GetTempPath() + "text.xls");
            var httpStream = await fileDownload.Content.ReadAsStreamAsync();
            await httpStream.CopyToAsync(fileStream);
            fileStream.Flush();

            string result = await response.Content.ReadAsStringAsync();
            return result;
        }
    }
}
