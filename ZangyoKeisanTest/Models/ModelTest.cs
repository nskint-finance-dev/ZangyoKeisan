using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ZangyoKeisan.Models;
using NPOI.SS.UserModel;
using System.IO;
using Livet;
using DeepEqual.Syntax;
using NUnit.Framework.Constraints;

namespace ZangyoKeisanTest
{
    class Excelツール
    {
        Model model;

        [SetUp]
        public void setUp()
        {
            model = Model.GetInstance();
        }

        #region parseNumbertoAlphabet テスト
        // parseNumbertoAlphabet テスト用データ
        static object[] numToAlphabetTestValues =
        {
            new object[]{0,"A"},
            new object[]{1,"B"},
            new object[]{2,"C"},
            new object[]{3,"D"},
            new object[]{4,"E"},
            new object[]{5,"F"},
            new object[]{6,"G"},
            new object[]{7,"H"},
            new object[]{8,"I"},
            new object[]{9,"J"},
            new object[]{10,"K"},
            new object[]{11,"L"},
            new object[]{12,"M"},
            new object[]{13,"N"},
            new object[]{14,"O"},
            new object[]{15,"P"},
            new object[]{16,"Q"},
            new object[]{17,"R"},
            new object[]{18,"S"},
            new object[]{19,"T"},
            new object[]{20,"U"},
            new object[]{21,"V"},
            new object[]{22,"W"},
            new object[]{23,"X"},
            new object[]{24,"Y"},
            new object[]{25,"Z"}
        };

        [TestCaseSource("numToAlphabetTestValues")]
        public void 正常_parseNumbertoAlphabet(int num, string res)
        {
            Type type = model.GetType();
            MethodInfo oMethod = type.GetMethod("parseNumbertoAlphabet", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            Assert.That(oMethod.Invoke(model, new object[] { num }), Is.EqualTo(res));
        }
        #endregion parseNumbertoAlphabet テスト
    }

    class 残業計算 {
        Model model;

        [SetUp]
        public void setUp()
        {
            model = Model.GetInstance();
        }

        #region calclateZangyo
        /// <summary>
        /// テストデータリストから各データを抽出し、テストを実行する
        /// </summary>
        static IEnumerable<TestCaseData> calclateZangyoTestDataProvider
        {
            get {
                
                List<KintaiTestData> calclateZangyoTestDataList = calclateZangyoTestData;

                foreach(var calclateZangyoTestDataSet in calclateZangyoTestDataList)
                {
                    Kintai kintai = new Kintai();

                    foreach(var kintai1DayTestData in calclateZangyoTestDataSet.testData)
                    {
                        Kintai1day kintai1day = new Kintai1day();
                        kintai1day.date = (DateTime)kintai1DayTestData.date;
                        kintai1day.dayOfWeek = kintai1day.date.DayOfWeek;
                        kintai1day.keitai = (string)kintai1DayTestData.keitai;
                        kintai1day.kinmuJikan = (TimeSpan)kintai1DayTestData.kinmuJikan;

                        kintai.kinmuInfo.Add(kintai1day);
                    }

                    yield return new TestCaseData(kintai).Returns(calclateZangyoTestDataSet.result).SetName(calclateZangyoTestDataSet.testName);
                }
            }
        }

        /// <summary>
        /// テストデータリスト
        /// </summary>
        static List<KintaiTestData> calclateZangyoTestData = new List<KintaiTestData>() {
            new KintaiTestData( 
                "１日分（平日・残業あり）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 4, 1), "勤務先", new TimeSpan(9, 0, 0) )
                },
                new TimeSpan(1, 0, 0)
            ),
            new KintaiTestData(
                "１週間分（土曜なし・残業なし）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 5, 7), "勤務先（日）", new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 8), "勤務先（月）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 9), "勤務先（火）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 10), "勤務先（水）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 11), "勤務先（木）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 12), "勤務先（金）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 13), "勤務先（土）", new TimeSpan(0, 0, 0) ),
                },
                new TimeSpan(0, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜なし・残業あり）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 5, 7), "勤務先（日）", new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 8), "勤務先（月）", new TimeSpan(9, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 9), "勤務先（火）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 10), "勤務先（水）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 11), "勤務先（木）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 12), "勤務先（金）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 13), "勤務先（土）", new TimeSpan(0, 0, 0) ),
                },
                new TimeSpan(1, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜あり・残業あり）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 5, 7), "勤務先（日）", new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 8), "勤務先（月）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 9), "勤務先（火）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 10), "勤務先（水）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 11), "勤務先（木）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 12), "勤務先（金）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 13), "勤務先（土）", new TimeSpan(3, 0, 0) ),
                },
                new TimeSpan(3, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜なし・残業なし・日曜カウント外の確認）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 5, 7), "勤務先（日）", new TimeSpan(1, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 8), "勤務先（月）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 9), "勤務先（火）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 10), "勤務先（水）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 11), "勤務先（木）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 12), "勤務先（金）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 13), "勤務先（土）", new TimeSpan(0, 0, 0) ),
                },
                new TimeSpan(0, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜なし・残業なし・日曜カウント外の確認）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 5, 7), "勤務先（日）", new TimeSpan(1, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 8), "勤務先（月）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 9), "勤務先（火）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 10), "勤務先（水）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 11), "勤務先（木）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 12), "勤務先（金）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 13), "勤務先（土）", new TimeSpan(0, 0, 0) ),
                },
                new TimeSpan(0, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜あり・残業あり・日曜カウント外の確認）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 5, 7), "勤務先（日）", new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 8), "勤務先（月）", new TimeSpan(10, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 9), "勤務先（火）", new TimeSpan(10, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 10), "勤務先（水）", new TimeSpan(10, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 11), "勤務先（木）", new TimeSpan(10, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 12), "勤務先（金）", new TimeSpan(10, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 5, 13), "勤務先（土）", new TimeSpan(10, 0, 0) ),
                },
                new TimeSpan(20, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜あり・残業なし・祝日あり）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 3, 19), "勤務先（日）",        new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 20), "勤務先（月・祝）",    new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 21), "勤務先（火）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 22), "勤務先（水）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 23), "勤務先（木）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 24), "勤務先（金）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 25), "勤務先（土）",        new TimeSpan(8, 0, 0) ),
                },
                new TimeSpan(0, 0, 0)
            ),
            new KintaiTestData(
                "１週間（土曜あり・日ごとの残業あり・週ごとの残業なし）",
                new List<Kintai1dayTestData>() {
                    new Kintai1dayTestData( new DateTime(2017, 3, 19), "勤務先（日）",        new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 20), "勤務先（月・祝）",    new TimeSpan(0, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 21), "勤務先（火）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 22), "勤務先（水）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 23), "勤務先（木）",        new TimeSpan(8, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 24), "勤務先（金）",        new TimeSpan(9, 0, 0) ),
                    new Kintai1dayTestData( new DateTime(2017, 3, 25), "勤務先（土）",        new TimeSpan(8, 0, 0) ),
                },
                new TimeSpan(1, 0, 0)
            )
        };

        /// <summary>
        /// テスト用の勤怠情報クラス
        /// （実際のクラスにはコンストラクターがないため使用）
        /// </summary>
        class Kintai1dayTestData : Kintai1day
        {
            public Kintai1dayTestData(DateTime date, string keitai, TimeSpan kinmuJikan)
            {
                this.date = date;
                this.keitai = keitai;
                this.kinmuJikan = kinmuJikan;
            }

            public Kintai1dayTestData(DateTime date, DayOfWeek dayOfWeek, string keitai, DateTime? shigyo, DateTime? syugyo, TimeSpan kinmuJikan, TimeSpan kyukei, TimeSpan gaisyutsu, TimeSpan tikokuSotai, TimeSpan jikoKeihatsu, TimeSpan gaizan, TimeSpan kyujitsuSyukkin, TimeSpan sinyaKinmu, String biko)
            {
                this.date = date;
                this.dayOfWeek = dayOfWeek;
                this.keitai = keitai;
                this.shigyo = shigyo;
                this.syugyo = syugyo;
                this.kinmuJikan = kinmuJikan;
                this.kyukei = kyukei;
                this.gaisyutsu = gaisyutsu;
                this.tikokuSotai = tikokuSotai;
                this.jikoKeihatsu = jikoKeihatsu;
                this.gaizan = gaizan;
                this.kyujitsuSyukkin = kyujitsuSyukkin;
                this.sinyaKinmu = sinyaKinmu;
                this.biko = biko;
            }
    }
        /// <summary>
        /// テストデータセット
        /// </summary>
        class KintaiTestData
        {
            /// <summary>
            /// 勤怠情報（テスト対象メソッドに渡す引数）
            /// </summary>
            public List<Kintai1dayTestData> testData;
            /// <summary>
            /// 残業時間（テスト対象メソッドの返り値）
            /// </summary>
            public TimeSpan result;
            /// <summary>
            /// テストに付ける名前
            /// </summary>
            public string testName;

            public KintaiTestData(string testName,List<Kintai1dayTestData> testData, TimeSpan result)
            {
                this.testName = testName;
                this.testData = testData;
                this.result = result;
            }
        }

        [TestCaseSource("calclateZangyoTestDataProvider")]
        public object 残業計算_calclateZangyo(Kintai kintai)
        {
            Type type = model.GetType();
            MethodInfo oMethod = type.GetMethod("calclateZangyo", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            return oMethod.Invoke(model, new object[] { kintai });            
        }

        #endregion calclateZangyo

        [Test]
        public void Excel読み込み()
        {
            Kintai kintai = new Kintai();
            
            // Excelをオブジェクトに変換した結果（期待値）
            // （各日のデータを書く際はExcelファイルにあるコピペ用コードを使うとよい）
            List<Kintai1dayTestData> resultExcelData = new List<Kintai1dayTestData>{
            new Kintai1dayTestData(new DateTime(2017, 3, 1),new DateTime(2017, 3, 1).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 1, 9, 0, 0),new DateTime(2017, 3, 1, 20, 30, 0),new TimeSpan(10, 0, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2, 0, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 2),new DateTime(2017, 3, 2).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 2, 9, 0, 0),new DateTime(2017, 3, 2, 21, 0, 0),new TimeSpan(9,30, 0),new TimeSpan(2,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 3),new DateTime(2017, 3, 3).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 3, 9, 0, 0),new DateTime(2017, 3, 3, 20, 15, 0),new TimeSpan(9,45, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(1,45, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 4),new DateTime(2017, 3, 4).DayOfWeek,"所休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 5),new DateTime(2017, 3, 5).DayOfWeek,"法休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 6),new DateTime(2017, 3, 6).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 6, 9, 0, 0),new DateTime(2017, 3, 6, 21, 0, 0),new TimeSpan(10,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 7),new DateTime(2017, 3, 7).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 7, 9, 0, 0),new DateTime(2017, 3, 7, 21, 0, 0),new TimeSpan(10,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 8),new DateTime(2017, 3, 8).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 8, 9, 0, 0),new DateTime(2017, 3, 8, 21, 0, 0),new TimeSpan(10,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 9),new DateTime(2017, 3, 9).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 9, 9, 0, 0),new DateTime(2017, 3, 9, 20, 30, 0),new TimeSpan(10, 0, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2, 0, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 10),new DateTime(2017, 3, 10).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 10, 9, 0, 0),new DateTime(2017, 3, 10, 20, 0, 0),new TimeSpan(9,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 11),new DateTime(2017, 3, 11).DayOfWeek,"所休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 12),new DateTime(2017, 3, 12).DayOfWeek,"法休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 13),new DateTime(2017, 3, 13).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 13, 9, 0, 0),new DateTime(2017, 3, 13, 20, 0, 0),new TimeSpan(9,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 14),new DateTime(2017, 3, 14).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 14, 9, 0, 0),new DateTime(2017, 3, 14, 22, 10, 0),new TimeSpan(10,25, 0),new TimeSpan(2,45, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2,25, 0),new TimeSpan(),new TimeSpan(0,10, 0),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 15),new DateTime(2017, 3, 15).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 15, 9, 0, 0),new DateTime(2017, 3, 15, 21, 10, 0),new TimeSpan(9,40, 0),new TimeSpan(2,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(1,40, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 16),new DateTime(2017, 3, 16).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 16, 9, 0, 0),new DateTime(2017, 3, 16, 18, 0, 0),new TimeSpan(7,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 17),new DateTime(2017, 3, 17).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 17, 9, 0, 0),new DateTime(2017, 3, 17, 18, 30, 0),new TimeSpan(8, 0, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 18),new DateTime(2017, 3, 18).DayOfWeek,"所休出",new DateTime(2017, 3, 18, 13, 0, 0),new DateTime(2017, 3, 18, 17, 0, 0),new TimeSpan(4, 0, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),"休出済"),
            new Kintai1dayTestData(new DateTime(2017, 3, 19),new DateTime(2017, 3, 19).DayOfWeek,"法休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 20),new DateTime(2017, 3, 20).DayOfWeek,"所休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),"春分の日"),
            new Kintai1dayTestData(new DateTime(2017, 3, 21),new DateTime(2017, 3, 21).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 21, 9, 0, 0),new DateTime(2017, 3, 21, 18, 0, 0),new TimeSpan(7,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 22),new DateTime(2017, 3, 22).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 22, 9, 0, 0),new DateTime(2017, 3, 22, 18, 0, 0),new TimeSpan(7,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 23),new DateTime(2017, 3, 23).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 23, 9, 0, 0),new DateTime(2017, 3, 23, 18, 30, 0),new TimeSpan(8, 0, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 24),new DateTime(2017, 3, 24).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 24, 9, 0, 0),new DateTime(2017, 3, 24, 18, 30, 0),new TimeSpan(8, 0, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 25),new DateTime(2017, 3, 25).DayOfWeek,"所休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 26),new DateTime(2017, 3, 26).DayOfWeek,"法休日",null,null,new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 27),new DateTime(2017, 3, 27).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 27, 9, 0, 0),new DateTime(2017, 3, 27, 21, 0, 0),new TimeSpan(10,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(2,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 28),new DateTime(2017, 3, 28).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 28, 9, 0, 0),new DateTime(2017, 3, 28, 19, 0, 0),new TimeSpan(8,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(0,30, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 29),new DateTime(2017, 3, 29).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 29, 9, 0, 0),new DateTime(2017, 3, 29, 19, 30, 0),new TimeSpan(9, 0, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(1, 0, 0),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 30),new DateTime(2017, 3, 30).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 30, 9, 0, 0),new DateTime(2017, 3, 30, 17, 30, 0),new TimeSpan(7,30, 0),new TimeSpan(1, 0, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),
            new Kintai1dayTestData(new DateTime(2017, 3, 31),new DateTime(2017, 3, 31).DayOfWeek,"テスト勤務先",new DateTime(2017, 3, 31, 9, 0, 0),new DateTime(2017, 3, 31, 18, 0, 0),new TimeSpan(7,30, 0),new TimeSpan(1,30, 0),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),new TimeSpan(),""),

            };

            foreach (var item in resultExcelData)
            {
                Kintai1day dailyKintai = new Kintai1day();
                dailyKintai.date = item.date;
                dailyKintai.dayOfWeek = item.dayOfWeek;
                dailyKintai.keitai = item.keitai;
                dailyKintai.shigyo = item.shigyo;
                dailyKintai.syugyo = item.syugyo;
                dailyKintai.kinmuJikan = item.kinmuJikan;
                dailyKintai.kyukei = item.kyukei;
                dailyKintai.gaisyutsu = item.gaisyutsu;
                dailyKintai.tikokuSotai = item.tikokuSotai;
                dailyKintai.jikoKeihatsu = item.jikoKeihatsu;
                dailyKintai.gaizan = item.gaizan;
                dailyKintai.kyujitsuSyukkin = item.kyujitsuSyukkin;
                dailyKintai.sinyaKinmu = item.sinyaKinmu;
                dailyKintai.biko = item.biko;

                kintai.kinmuInfo.Add(dailyKintai);
            }

            ObservableSynchronizedCollection<Kintai> result = new ObservableSynchronizedCollection<Kintai>();

            kintai.department = "AAA事業部";
            kintai.name = "山田 太郎";
            kintai.employeeID = "99999";
            result.Add(kintai);

            // Excelファイルを読み込む処理
            string testFolderPath = System.AppDomain.CurrentDomain.BaseDirectory + @"..\..\Test\";
            string filePath = testFolderPath + @"単体テスト\test.xls";
            IWorkbook workBook;
            using (FileStream infile = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                // Excelファイルを読み込む
                workBook = WorkbookFactory.Create(infile, ImportOption.All);
            }

            model.loadKintaiFromExcel(filePath);

            // private メソッドをテストするための設定
            Type type = model.GetType();
            MethodInfo oMethod = type.GetMethod("parseExceltoDataObject", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            ObservableSynchronizedCollection<Kintai> expected = (ObservableSynchronizedCollection<Kintai>)oMethod.Invoke(model, new object[] { workBook });

            Assert.That(expected, Is.KintaiEqualTo(result));
        }
    }

    public class Is : NUnit.Framework.Is
    {
        public static KintaiConstraint KintaiEqualTo(ObservableSynchronizedCollection<Kintai> expected)
        {
            return new KintaiConstraint(expected);
        }
    }

    public class KintaiConstraint : Constraint
    {
        private ObservableSynchronizedCollection<Kintai> expected;

        public KintaiConstraint(ObservableSynchronizedCollection<Kintai> expected)
        {
            this.expected = expected;
        }

        public override ConstraintResult ApplyTo<TActual>(TActual actual)
        {
            ObservableSynchronizedCollection<Kintai> actualList = actual as ObservableSynchronizedCollection<Kintai>;
            if (actualList == null)
            {
                return new ConstraintResult(this, 1, ConstraintStatus.Failure);
            }
            else
            {
                for (var i = 0; i < actualList.Count; i++)
                {
                    if (actualList[i].name != expected[i].name)
                    {
                        Description = "Index [" + i + "].name = " + expected[i].name;
                        return new ConstraintResult(this, actualList[i].name, ConstraintStatus.Failure);
                    }

                    if (actualList[i].department != expected[i].department)
                    {
                        Description = "Index [" + i + "].department = " + expected[i].department;
                        return new ConstraintResult(this, actualList[i].department, ConstraintStatus.Failure);
                    }

                    if (actualList[i].employeeID != expected[i].employeeID)
                    {
                        Description = "Index [" + i + "].department = " + expected[i].employeeID;
                        return new ConstraintResult(this, actualList[i].employeeID, ConstraintStatus.Failure);
                    }

                    for (int j = 0; j < actualList[i].kinmuInfo.Count; j++)
                    {
                        if (actualList[i].kinmuInfo[j].date != expected[i].kinmuInfo[j].date)
                        {
                            Description = $"Index [{i}] -> kinmuInfo[{j}].date = {expected[i].kinmuInfo[j].date}";
                            return new ConstraintResult(this, actualList[i].kinmuInfo[j].date, ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].dayOfWeek != expected[i].kinmuInfo[j].dayOfWeek)
                        {
                            Description = $"Index [{i}] -> kinmuInfo[{j}].dayOfWeek = {expected[i].kinmuInfo[j].dayOfWeek}";
                            return new ConstraintResult(this, actualList[i].kinmuInfo[j].dayOfWeek, ConstraintStatus.Failure);
                        }

                        DateTime targetDate = actualList[i].kinmuInfo[j].date;

                        if (actualList[i].kinmuInfo[j].keitai != expected[i].kinmuInfo[j].keitai)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 形態 = {expected[i].kinmuInfo[j].keitai}";
                            return new ConstraintResult(this, actualList[i].kinmuInfo[j].keitai, ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].shigyo != expected[i].kinmuInfo[j].shigyo)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 始業 = {expected[i].kinmuInfo[j].shigyo?.ToString("HH:mm")} ({expected[i].kinmuInfo[j].shigyo})";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].shigyo?.ToString("HH:mm")} ({actualList[i].kinmuInfo[j].shigyo})", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].syugyo != expected[i].kinmuInfo[j].syugyo)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 終業 = {expected[i].kinmuInfo[j].syugyo?.ToString("HH:mm")} ({expected[i].kinmuInfo[j].syugyo})";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].syugyo?.ToString("HH:mm")} ({actualList[i].kinmuInfo[j].syugyo})", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].kinmuJikan != expected[i].kinmuInfo[j].kinmuJikan)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 勤務時間 = {expected[i].kinmuInfo[j].kinmuJikan.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].kinmuJikan.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].kyukei != expected[i].kinmuInfo[j].kyukei)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 休憩時間 = {expected[i].kinmuInfo[j].kyukei.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].kyukei.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].gaisyutsu != expected[i].kinmuInfo[j].gaisyutsu)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 外出時間 = {expected[i].kinmuInfo[j].gaisyutsu.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].gaisyutsu.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].tikokuSotai != expected[i].kinmuInfo[j].tikokuSotai)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 遅刻・早退時間 = {expected[i].kinmuInfo[j].tikokuSotai.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].tikokuSotai.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].jikoKeihatsu != expected[i].kinmuInfo[j].jikoKeihatsu)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 自己啓発時間 = {expected[i].kinmuInfo[j].jikoKeihatsu.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].jikoKeihatsu.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].gaizan != expected[i].kinmuInfo[j].gaizan)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 時間外労働時間 = {expected[i].kinmuInfo[j].gaizan.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].gaizan.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].kyujitsuSyukkin != expected[i].kinmuInfo[j].kyujitsuSyukkin)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 休日勤務時間 = {expected[i].kinmuInfo[j].kyujitsuSyukkin.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].kyujitsuSyukkin.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].sinyaKinmu != expected[i].kinmuInfo[j].sinyaKinmu)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 深夜勤務時間 = {expected[i].kinmuInfo[j].sinyaKinmu.ToString("c")}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].sinyaKinmu.ToString("c")}", ConstraintStatus.Failure);
                        }

                        if (actualList[i].kinmuInfo[j].biko != expected[i].kinmuInfo[j].biko)
                        {
                            Description = $"{targetDate.ToString("M/d(ddd)")} (Index:{i}) : 備考 = {expected[i].kinmuInfo[j].biko}";
                            return new ConstraintResult(this, $"{actualList[i].kinmuInfo[j].biko}", ConstraintStatus.Failure);
                        }
                    }
                }
            }
            return new ConstraintResult(this, actual, ConstraintStatus.Success);
        }
    }

    
}
