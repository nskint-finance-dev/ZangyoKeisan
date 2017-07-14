using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Livet;
using NPOI.SS.UserModel;
using System.IO;
using System.Text.RegularExpressions;

namespace ZangyoKeisan.Models
{
    public class Model : NotificationObject
    {
        /*
         * NotificationObjectはプロパティ変更通知の仕組みを実装したオブジェクトです。
         */

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

        public Model()
        {
            KintaiList = new ObservableSynchronizedCollection<Kintai>();
        }

        /// <summary>
        /// Mospから出力したエクセルファイルを読み込む
        /// </summary>
        /// <param name="filePath">エクセルファイルパス</param>
        public void loadKintaiFromExcel(string filePath)
        {
            try
            {
                IWorkbook workBook;
                using (FileStream infile = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    // Excelファイルを読み込む
                    workBook = WorkbookFactory.Create(infile, ImportOption.All);
                }

                // エクセル勤怠簿のデータをオブジェクトに変換
                ObservableSynchronizedCollection<Kintai> kintaiList = parseExceltoDataObject(workBook);

                KintaiList.Clear();

                foreach (var item in kintaiList)
                {
                    // Add で追加しないと変更が反映されない
                    KintaiList.Add(item);
                }

                foreach (var kintai in KintaiList)
                {
                    kintai.zangyo = calclateZangyo(kintai);
                }
            }
            catch (System.IO.FileNotFoundException ex)
            {
                Console.WriteLine("ファイルが見つかりません!");
            }
        }

        /// <summary>
        /// 「氏名」が記載されたセルのアドレス
        /// </summary>
        const string CELL_ADDRESS_NAME = "D4";
        /// <summary>
        /// 「所属」が記載されたセルのアドレス
        /// </summary>
        const string CELL_ADDRESS_DEPARTMENT = "D3";
        /// <summary>
        /// 「社員番号」が記載されたセルのアドレス
        /// </summary>
        const string CELL_ADDRESS_EMPLOYEE_ID = "D5";
        /// <summary>
        /// 日ごとの勤怠が記載始められている行
        /// </summary>
        const int ROW_START_KINTAI = 8;
        /// <summary>
        /// 対象年月が書かれたセルのアドレス
        /// </summary>
        const string CELL_TARGET_YYYYMM = "D1";
        /// <summary>
        /// 時刻が記入されていないことを示す文字
        /// </summary>
        const string CELL_NULL_CHAR = "-";

        /// <summary>
        /// Excelの勤怠情報をC#のオブジェクトに変換する
        /// </summary>
        /// <param name="workBook">Excelワークブック</param>
        private ObservableSynchronizedCollection<Kintai> parseExceltoDataObject(IWorkbook workBook)
        {
            ObservableSynchronizedCollection<Kintai> kintaiList = new ObservableSynchronizedCollection<Kintai>();

            // シート（個人）ごとに処理する
            for (int sheetNum = 0; sheetNum < workBook.NumberOfSheets; sheetNum++)
            {
                // 個人の勤怠
                Kintai kintai = new Kintai();

                // シートごとにセルの値を取得する
                ISheet workSheet = workBook.GetSheetAt(sheetNum);
                // データの最終行
                int lastRow = workSheet.LastRowNum;

                // １回取得すればよい項目を取得する

                // チェック対象の年月
                ICell tmpCell = getCellByAddress(CELL_TARGET_YYYYMM, workSheet);
                int year = tmpCell.DateCellValue.Year;
                int month = tmpCell.DateCellValue.Month;

                // 名前
                tmpCell = getCellByAddress(CELL_ADDRESS_NAME, workSheet);
                kintai.name = tmpCell?.StringCellValue;

                // 所属
                tmpCell = getCellByAddress(CELL_ADDRESS_DEPARTMENT, workSheet);
                kintai.department = tmpCell?.StringCellValue;

                // 社員番号
                tmpCell = getCellByAddress(CELL_ADDRESS_EMPLOYEE_ID, workSheet);
                kintai.employeeID = tmpCell?.StringCellValue.ToString();

                // その月の日数
                int daysInMonth = DateTime.DaysInMonth(year, month);

                // 日ごとの勤怠を取得する
                for (int rowNum = ROW_START_KINTAI - 1; rowNum < ROW_START_KINTAI + daysInMonth - 1; rowNum++)
                {
                    IRow row = workSheet.GetRow(rowNum);
                    // セルに何も入力されていない場合、row.LastCellNumはnullになる
                    short? lastColumn = row?.LastCellNum;

                    // １日ごとの勤怠
                    Kintai1day kintai1day = new Kintai1day();

                    // 項目ごとに値を取得する
                    for (int columnNum = 0; columnNum < lastColumn; columnNum++)
                    {
                        ICell cell = row?.GetCell(columnNum);
                        // セルの値を取得する
                        // （cell(5, 4)など、アドレスを指定して値を取得することはできないため、若干面倒）

                        // 操作対象の日付（作業用）
                        int tmpDay;
                        // 作業用の時
                        string tmpHour;
                        // 作業用の分
                        string tmpMinute;

                        switch (parseNumbertoAlphabet(cell.ColumnIndex))
                        {
                            // 形態
                            case "B":
                                kintai1day.keitai = cell.StringCellValue;
                                break;
                            // 始業
                            case "C":
                                if (cell.StringCellValue != CELL_NULL_CHAR)
                                {
                                    tmpDay = row.GetCell(parseAlphabettoNumber('A')).DateCellValue.Day;
                                    tmpHour = cell.StringCellValue.Split(':')[0];
                                    tmpMinute = cell.StringCellValue.Split(':')[1];
                                    kintai1day.shigyo = new DateTime(year, month, tmpDay, int.Parse(tmpHour), int.Parse(tmpMinute), 0);
                                }
                                else
                                {
                                    kintai1day.shigyo = null;
                                }

                                break;
                            // 終業
                            case "D":
                                if (cell.StringCellValue != CELL_NULL_CHAR)
                                {
                                    tmpDay = row.GetCell(parseAlphabettoNumber('A')).DateCellValue.Day;
                                    tmpHour = cell.StringCellValue.Split(':')[0];
                                    tmpMinute = cell.StringCellValue.Split(':')[1];
                                    try
                                    {
                                        // 24時以降に対応するための処理
                                        DateTime tmpSyugyoDate = new DateTime(year, month, tmpDay);
                                        TimeSpan tmpSyugyoTime = new TimeSpan(int.Parse(tmpHour), int.Parse(tmpMinute), 0);
                                        kintai1day.syugyo = tmpSyugyoDate + tmpSyugyoTime;
                                    }catch(ArgumentOutOfRangeException ex)
                                    {
                                        Console.WriteLine("終業時刻を変換できませんでした (" + parseNumbertoAlphabet(cell.ColumnIndex) + cell.RowIndex + ")");
                                    }
                                }
                                else
                                {
                                    kintai1day.syugyo = null;
                                }
                                break;
                            // 勤務時間
                            case "E":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.kinmuJikan = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.kinmuJikan = new TimeSpan();
                                }
                                break;
                            // 休憩時間
                            case "F":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.kyukei = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.kyukei = new TimeSpan();
                                }
                                break;
                            // 外出時間
                            case "G":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.gaisyutsu = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.gaisyutsu = new TimeSpan();
                                }
                                break;
                            // 遅刻・早退
                            case "H":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.tikokuSotai = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.tikokuSotai = new TimeSpan();
                                }
                                break;
                            // 自己啓発
                            case "I":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.jikoKeihatsu = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.jikoKeihatsu = new TimeSpan();
                                }
                                break;
                            // 勤務外残業
                            case "J":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.gaizan = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.gaizan = new TimeSpan();
                                }
                                break;
                            // 休日出勤
                            case "K":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.kyujitsuSyukkin = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.kyujitsuSyukkin = new TimeSpan();
                                }
                                break;
                            // 深夜勤務
                            case "L":
                                if (cell.CellType == CellType.Numeric)
                                {
                                    kintai1day.sinyaKinmu = parseDecimaltoTimeSpan(cell.NumericCellValue);
                                }
                                else
                                {
                                    kintai1day.sinyaKinmu = new TimeSpan();
                                }
                                break;
                            // 備考
                            case "M":
                                kintai1day.biko = cell?.StringCellValue;
                                break;

                            case "A":
                                // 日付
                                tmpDay = row.GetCell(parseAlphabettoNumber('A')).DateCellValue.Day;
                                kintai1day.date = new DateTime(year, month, tmpDay);

                                // 曜日
                                kintai1day.dayOfWeek = kintai1day.date.DayOfWeek;
                                break;

                            default:
                                break;
                        }


                    }
                    kintai.kinmuInfo.Add(kintai1day);
                }

                kintaiList.Add(kintai);
            }

            return kintaiList;
        }

        /// <summary>
        /// ExcelのセルアドレスをR1C1参照方式（列、行とも数字）からA1参照方式（アルファベット+数字）で表す形式に変換する
        /// </summary>
        /// <param name="columnNum">列番号</param>
        /// <param name="rowNum">行番号</param>
        /// <returns>A1参照方式で表したセルアドレス</returns>
        private string parseToAddress(int columnNum, int rowNum)
        {
            // アルファベットに変換するためのテーブル
            const string columnNameTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (columnNum >= 26)
            {
                throw new ArgumentOutOfRangeException("引数には、25 以下の数字を指定してください: " + columnNum);
            }

            return columnNameTable.Substring(columnNum, 1) + (rowNum + 1);
        }

        /// <summary>
        /// Excelの列を表す数字をアルファベットに変換する
        /// </summary>
        /// <param name="num">列を示す数字</param>
        /// <returns>列を表すアルファベット</returns>
        private string parseNumbertoAlphabet(int num)
        {
            // 数字に変換するためのテーブル
            const string alphabetTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (num >= 26)
            {
                throw new ArgumentOutOfRangeException("引数には、25 以下の数字を指定してください: " + num);
            }

            return alphabetTable.Substring(num, 1);
        }

        /// <summary>
        /// Excelの列を表すアルファベットを数字（0始まり）に変換する
        /// </summary>
        /// <param name="colName">列を示すアルファベット</param>
        /// <returns>列を示す数字</returns>
        private int parseAlphabettoNumber(char colName)
        {
            // 数字に変換するためのテーブル
            const string alphabetTable = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (alphabetTable.IndexOf(colName) == -1)
            {
                throw new ArgumentException("変換できない文字です: " + colName);
            }

            return alphabetTable.IndexOf(colName);
        }

        /// <summary>
        /// A1参照方式で指定したセルオブジェクトを取得する
        /// </summary>
        /// <param name="address">取得したいセルのアドレス（A1参照方式）</param>
        /// <param name="workSheet">セルが存在するシート</param>
        /// <returns>セル</returns>
        private ICell getCellByAddress(string address, ISheet workSheet)
        {
            // 引数（セルのアドレス）チェック用正規表現
            string cellAddressPattern = @"^[A-Z][1-9][0-9]*";
            // アドレスの形式をチェック
            if (!Regex.IsMatch(address, cellAddressPattern))
            {
                throw new ArgumentException("セルのアドレス指定方法が対応外です: " + address);
            }

            // 列を示す数字
            string columnName = address.Substring(0, 1);
            int columnNum = parseAlphabettoNumber(char.Parse(columnName));
            // 行番号
            int rowNum = int.Parse(address.Substring(1)) - 1;

            // セルを取得
            IRow row = workSheet.GetRow(rowNum);
            ICell cell = row.GetCell(columnNum);

            return cell;
        }

        /// <summary>
        /// 残業時間を計算する
        /// </summary>
        /// <param name="">１ヶ月分の勤怠記録</param>
        /// <returns>残業時間</returns>
        private TimeSpan calclateZangyo(Kintai kintai)
        {
            // １ヶ月の残業時間
            TimeSpan monthlyZangyoTime = new TimeSpan();

            // １日の法定労働時間
            TimeSpan dailyHoteiRodoTime = new TimeSpan(8, 0, 0);
            // １週間の法定労働時間
            TimeSpan weeklyHoteiRodoTime = new TimeSpan(40, 0, 0);
            // １日あたりの残業時間（８時間超過）の合計
            TimeSpan sumDailyZangyo = new TimeSpan();
            // １週間ごとの残業時間（４０時間超過）
            TimeSpan weeklyZangyo = new TimeSpan();
            // １週間ごとの合計勤務時間
            TimeSpan weeklyRodoTime = new TimeSpan();

            foreach (var dailyKintai in kintai.kinmuInfo)
            {
                // １日ごとの残業時間を計算
                if (dailyKintai.dayOfWeek != DayOfWeek.Sunday && dailyKintai.keitai != "有休" && dailyKintai.keitai != "所休日" && dailyKintai.keitai != "法代休")
                {
                    // 残業時間の計算対象日の場合
                    if (dailyKintai.kinmuJikan.TotalHours > dailyHoteiRodoTime.TotalHours)
                    {
                        // 法定労働時間（８時間）を超えて勤務した時間を、１日あたりの残業時間として算出
                        sumDailyZangyo += dailyKintai.kinmuJikan - dailyHoteiRodoTime;
                    }

                    // １週間ごとの残業時間計算のため、勤務時間を合算
                    weeklyRodoTime += dailyKintai.kinmuJikan;
                }

                // １週間ごとの残業時間を計算
                if (dailyKintai.dayOfWeek == DayOfWeek.Sunday || dailyKintai == kintai.kinmuInfo.Last())
                {
                    if (weeklyRodoTime - sumDailyZangyo > weeklyHoteiRodoTime)
                    {
                        weeklyZangyo = weeklyRodoTime - weeklyHoteiRodoTime - sumDailyZangyo;
                    }

                    // １ヶ月の合計残業時間に合算
                    monthlyZangyoTime += sumDailyZangyo + weeklyZangyo;

                    // 初期化
                    sumDailyZangyo = new TimeSpan(0, 0, 0);
                    weeklyZangyo = new TimeSpan(0, 0, 0);
                    weeklyRodoTime = new TimeSpan(0, 0, 0);
                }
            }

            return monthlyZangyoTime;
        }

        /// <summary>
        /// 小数で表した時間を"hh:mm"形式に変換する
        /// </summary>
        /// <param name="num">小数で表した時間</param>
        /// <returns>"hh:mm"形式で表した時間</returns>
        private TimeSpan parseDecimaltoTimeSpan(double num)
        {
            return TimeSpan.Parse(Math.Truncate(num) + ":" + Math.Round(num % 1 * 60, MidpointRounding.AwayFromZero));
        }
    }

    /// <summary>
    /// １ヶ月ごとの勤務情報
    /// </summary>
    public class Kintai : NotificationObject
    {
        #region name変更通知プロパティ
        private string _name;
        /// <summary>
        /// 氏名
        /// </summary>
        public string name
        {
            get
            { return _name; }
            set
            {
                if (_name == value)
                    return;
                _name = value;
                RaisePropertyChanged();
            }
        }
        #endregion

        #region employeeID変更通知プロパティ
        private string _employeeID;
        /// <summary>
        /// 社員番号
        /// </summary>
        public string employeeID
        {
            get
            { return _employeeID; }
            set
            {
                if (_employeeID == value)
                    return;
                _employeeID = value;
                RaisePropertyChanged();
            }
        }
        #endregion

        #region department変更通知プロパティ
        private string _department;
        /// <summary>
        /// 所属
        /// </summary>
        public string department
        {
            get
            { return _department; }
            set
            {
                if (_department == value)
                    return;
                _department = value;
                RaisePropertyChanged();
            }
        }
        #endregion


        #region kinmuInfo変更通知プロパティ
        private List<Kintai1day> _kinmuInfo;
        /// <summary>
        /// １日ごとの勤務情報
        /// </summary>
        public List<Kintai1day> kinmuInfo
        {
            get
            { return _kinmuInfo; }
            set
            {
                if (_kinmuInfo == value)
                    return;
                _kinmuInfo = value;
                RaisePropertyChanged();
            }
        }
        #endregion

        #region zangyo変更通知プロパティ
        private TimeSpan? _zangyo;
        /// <summary>
        /// 基準に基づいて計算した残業時間
        /// </summary>
        public TimeSpan? zangyo
        {
            get
            { return _zangyo; }
            set
            {
                if (_zangyo == value)
                    return;
                _zangyo = value;
                RaisePropertyChanged();
            }
        }
        #endregion

        public Kintai()
        {
            kinmuInfo = new List<Kintai1day>();
        }
    }

    /// <summary>
    /// １日ごとの勤務情報
    /// </summary>
    public class Kintai1day
    {
        /// <summary>
        /// 日付
        /// </summary>
        public DateTime date;
        /// <summary>
        /// 曜日
        /// </summary>
        public DayOfWeek dayOfWeek;
        /// <summary>
        /// 形態（勤務先名、所休日、法休日など）
        /// </summary>
        public string keitai;
        /// <summary>
        /// 始業時刻
        /// </summary>
        public DateTime? shigyo;
        /// <summary>
        /// 終業時刻
        /// </summary>
        public DateTime? syugyo;
        /// <summary>
        /// 勤務時間
        /// </summary>
        public TimeSpan kinmuJikan;
        /// <summary>
        /// 休憩時間
        /// </summary>
        public TimeSpan kyukei;
        /// <summary>
        /// 外出時間
        /// </summary>
        public TimeSpan gaisyutsu;
        /// <summary>
        /// 遅刻・早退時間
        /// </summary>
        public TimeSpan tikokuSotai;
        /// <summary>
        /// 自己啓発時間
        /// </summary>
        public TimeSpan jikoKeihatsu;
        /// <summary>
        /// 時間外労働時間
        /// </summary>
        public TimeSpan gaizan;
        /// <summary>
        /// 休日勤務時間
        /// </summary>
        public TimeSpan kyujitsuSyukkin;
        /// <summary>
        /// 深夜勤務時間
        /// </summary>
        public TimeSpan sinyaKinmu;
        /// <summary>
        /// 備考
        /// </summary>
        public String biko;
    }
}
