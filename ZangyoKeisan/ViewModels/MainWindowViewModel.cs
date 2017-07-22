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
using System.Threading.Tasks;
using Microsoft.Win32;

namespace ZangyoKeisan.ViewModels
{
    public class MainWindowViewModel : ViewModel
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
        PropertyChangedEventListener kintaiDataListener;

        public void Initialize()
        {
            model = Model.GetInstance();

            // 初期表示時は、メイン画面を表示するためのデータがそろっていない
            IsDisplayReady = false;

            // 勤怠記録を取得できたら表示する
            kintaiDataListener = new PropertyChangedEventListener(model);
            kintaiDataListener.RegisterHandler(
                () => model.KintaiData, (_, __) => {
                    KintaiData = model.KintaiData;
                    IsDisplayReady = true;
            });
        }

        #region KintaiList変更通知プロパティ
        private ObservableSynchronizedCollection<Kintai> _KintaiList;

        public ObservableSynchronizedCollection<Kintai> KintaiList
        {
            get
            { return _KintaiList; }
            set
            {
                if (_KintaiList == value)
                    return;
                _KintaiList = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        #region LoadExcelCommand
        private ViewModelCommand _LoadExcelCommand;

        public ViewModelCommand LoadExcelCommand
        {
            get
            {
                if (_LoadExcelCommand == null)
                {
                    _LoadExcelCommand = new ViewModelCommand(LoadExcel);
                }
                return _LoadExcelCommand;
            }
        }

        /// <summary>
        /// 「ファイルを開く」ダイアログを表示して、エクセルファイルを読み込む
        /// </summary>
        public void LoadExcel()
        {
                var dialog = new OpenFileDialog();
                dialog.Title = "Excelファイルを開く";
                dialog.Filter = "Excelファイル(*.xls, *.xlsx)|*.xls;*.xlsx";
                if (dialog.ShowDialog() == true)
                {
                    model.loadKintaiFromExcel(dialog.FileName);
                }
        }
        #endregion

        public void LoadExcel(string kintaiboPath)
        {
            model.loadKintaiFromExcel(kintaiboPath);
        }

        #region IsDisplayReady変更通知プロパティ
        private bool _IsDisplayReady;

        /// <summary>
        /// メイン画面を表示する準備が整っているか（trueなら整っている）
        /// </summary>
        public bool IsDisplayReady
        {
            get
            { return _IsDisplayReady; }
            set
            {
                if (_IsDisplayReady == value)
                    return;
                _IsDisplayReady = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        #region KintaiData変更通知プロパティ
        private Kintai _KintaiData;

        /// <summary>
        /// 勤怠記録（個人）
        /// </summary>
        public Kintai KintaiData
        {
            get
            { return _KintaiData; }
            set
            { 
                if (_KintaiData == value)
                    return;
                _KintaiData = value;
                RaisePropertyChanged();
            }
        }
        #endregion

    }
}
