using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;
using System.IO;

using Livet;

namespace ZangyoKeisan
{
    /// <summary>
    /// App.xaml の相互作用ロジック
    /// </summary>
    public partial class App : Application
    {
        private void Application_Startup(object sender, StartupEventArgs e)
        {
            DispatcherHelper.UIDispatcher = Dispatcher;
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
        }

        //集約エラーハンドラ
        private void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            string logFileName = "error.log";

            using (FileStream fileStream = new FileStream(logFileName, FileMode.Append))
            using (StreamWriter streamWriter = new StreamWriter(fileStream))
            {
                streamWriter.Write(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.ff") + Environment.NewLine + e.ExceptionObject.ToString() + Environment.NewLine + Environment.NewLine);
            }
        
            MessageBox.Show(
                "エラーが発生したため、プログラムを終了します。申し訳ございません。\n\nエラーが発生した状況とログを報告して頂けると幸いです。\n" + System.AppDomain.CurrentDomain.BaseDirectory + "\\" + logFileName,
                "エラー",
               MessageBoxButton.OK,
                MessageBoxImage.Error);
        
            Environment.Exit(1);
        }
    }
}
