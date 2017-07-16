using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;

using Livet;
using Livet.Commands;
using Livet.Messaging;
using Livet.Messaging.IO;
using Livet.EventListeners;
using Livet.Messaging.Windows;

using ZangyoKeisan.Models;

using System.Windows.Interactivity;
using System.Windows.Controls;
using System.Windows;
using System.Windows.Input;
using Microsoft.Expression.Interactivity;

namespace ZangyoKeisan.ViewModels
{
    public class DownloadWindowViewModel : ViewModel
    {
        /* コマンド、プロパティの定義にはそれぞれ 
         * 
         *  lvcom   : ViewModelCommand
         *  lvcomn  : ViewModelCommand(CanExecute無)
         *  llcom   : ListenerCommand(パラメータ有のコマンド)
         *  llcomn  : ListenerCommand(パラメータ有のコマンド・CanExecute無)
         *  lprop   : 変更通知プロパティ(.NET4.5ではlpropn)
         *  
         * を使用してください。
         * 
         * Modelが十分にリッチであるならコマンドにこだわる必要はありません。
         * View側のコードビハインドを使用しないMVVMパターンの実装を行う場合でも、ViewModelにメソッドを定義し、
         * LivetCallMethodActionなどから直接メソッドを呼び出してください。
         * 
         * ViewModelのコマンドを呼び出せるLivetのすべてのビヘイビア・トリガー・アクションは
         * 同様に直接ViewModelのメソッドを呼び出し可能です。
         */

        /* ViewModelからViewを操作したい場合は、View側のコードビハインド無で処理を行いたい場合は
         * Messengerプロパティからメッセージ(各種InteractionMessage)を発信する事を検討してください。
         */

        /* Modelからの変更通知などの各種イベントを受け取る場合は、PropertyChangedEventListenerや
         * CollectionChangedEventListenerを使うと便利です。各種ListenerはViewModelに定義されている
         * CompositeDisposableプロパティ(LivetCompositeDisposable型)に格納しておく事でイベント解放を容易に行えます。
         * 
         * ReactiveExtensionsなどを併用する場合は、ReactiveExtensionsのCompositeDisposableを
         * ViewModelのCompositeDisposableプロパティに格納しておくのを推奨します。
         * 
         * LivetのWindowテンプレートではViewのウィンドウが閉じる際にDataContextDisposeActionが動作するようになっており、
         * ViewModelのDisposeが呼ばれCompositeDisposableプロパティに格納されたすべてのIDisposable型のインスタンスが解放されます。
         * 
         * ViewModelを使いまわしたい時などは、ViewからDataContextDisposeActionを取り除くか、発動のタイミングをずらす事で対応可能です。
         */

        /* UIDispatcherを操作する場合は、DispatcherHelperのメソッドを操作してください。
         * UIDispatcher自体はApp.xaml.csでインスタンスを確保してあります。
         * 
         * LivetのViewModelではプロパティ変更通知(RaisePropertyChanged)やDispatcherCollectionを使ったコレクション変更通知は
         * 自動的にUIDispatcher上での通知に変換されます。変更通知に際してUIDispatcherを操作する必要はありません。
         */

        Model model;
        PropertyChangedEventListener statusListener;

        public void Initialize()
        {
            model = Model.GetInstance();
            statusListener = new PropertyChangedEventListener(model);
            statusListener.RegisterHandler(() => model.DownloadStatus, (s, e) =>
            {
                DownloadStatus = model.DownloadStatus;
            });

            TargetYears = new ObservableSynchronizedCollection<string>()
            {
                "2015",
                "2016",
                "2017"
            };

            TargetMonthes = new ObservableSynchronizedCollection<string>()
            {
                "1",
                "2",
                "3",
                "4",
                "5",
                "6",
                "7",
                "8",
                "9",
                "10",
                "12"
            };
        }



        #region DownloadCommand
        private ViewModelCommand _DownloadCommand;

        public ViewModelCommand DownloadCommand
        {
            get
            {
                if (_DownloadCommand == null)
                {
                    _DownloadCommand = new ViewModelCommand(Download, CanDownload);
                }
                return _DownloadCommand;
            }
        }

        public bool CanDownload()
        {
            // 必要な項目に入力されている場合のみ押下できる
            if (Id != null && Password != null && SelectedYear != null && SelectedMonth != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public async void Download()
        {
            string kintaiboFilePath = await model.downloadExcel(Id, Password, SelectedYear, SelectedMonth);

            // ダウンロードに成功した場合
            if (kintaiboFilePath != "")
            {
                // ダウンロード画面を閉じる
                Messenger.Raise(new WindowActionMessage(WindowAction.Close, "CloseWindow"));
            }
        }
        #endregion

        #region Id変更通知プロパティ
        private string _Id;

        public string Id
        {
            get
            { return _Id; }
            set
            { 
                if (_Id == value)
                    return;
                _Id = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        #region Password変更通知プロパティ
        private string _Password;

        public string Password
        {
            get
            { return _Password; }
            set
            { 
                if (_Password == value)
                    return;
                _Password = value;
                RaisePropertyChanged();

                DownloadCommand.RaiseCanExecuteChanged();
            }
        }
        #endregion


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

                DownloadCommand.RaiseCanExecuteChanged();
            }
        }
        #endregion



        #region TargetYearsList変更通知プロパティ
        private ObservableSynchronizedCollection<string> _TargetYearsList;

        public ObservableSynchronizedCollection<string> TargetYears
        {
            get
            { return _TargetYearsList; }
            set
            { 
                if (_TargetYearsList == value)
                    return;
                _TargetYearsList = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        #region SelectedYear変更通知プロパティ
        private string _SelectedYear;

        public string SelectedYear
        {
            get
            { return _SelectedYear; }
            set
            { 
                if (_SelectedYear == value)
                    return;
                _SelectedYear = value;
                RaisePropertyChanged();

                DownloadCommand.RaiseCanExecuteChanged();
            }
        }
        #endregion


        #region TargetMonthes変更通知プロパティ
        private ObservableSynchronizedCollection<string> _TargetMonthes;

        public ObservableSynchronizedCollection<string> TargetMonthes
        {
            get
            { return _TargetMonthes; }
            set
            { 
                if (_TargetMonthes == value)
                    return;
                _TargetMonthes = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        #region SelectedMonth変更通知プロパティ
        private string _SelectedMonth;

        public string SelectedMonth
        {
            get
            { return _SelectedMonth; }
            set
            { 
                if (_SelectedMonth == value)
                    return;
                _SelectedMonth = value;
                RaisePropertyChanged();

                DownloadCommand.RaiseCanExecuteChanged();
            }
        }
        #endregion
    }
}
