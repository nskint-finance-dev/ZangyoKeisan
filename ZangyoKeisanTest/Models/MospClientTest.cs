using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ZangyoKeisan.Models;


namespace ZangyoKeisanTest.Models
{
    class MospClientTest
    {
        [Test]
        public async Task 勤怠簿ダウンロード()
        {
            var model = new MospClient();
            string res = await model.downloadExcel("", "", "", "");

            Console.WriteLine(res);

            Assert.That(res, Is.EqualTo("test"));
        }

        [Test]
        public void procSeq取得()
        {
            MospClient mospClient = new MospClient();
            Type type = mospClient.GetType();
            MethodInfo oMethod = type.GetMethod("getProcSeq", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            string result =  (string)oMethod.Invoke(mospClient, new object[] { "<script>var procSeq = \"9\";</script>" });

            Assert.That("9", Is.EqualTo(result));
        }

        [Test]
        public void procSeq取得_失敗()
        {
            MospClient mospClient = new MospClient();
            Type type = mospClient.GetType();
            MethodInfo oMethod = type.GetMethod("getProcSeq", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            string result = (string)oMethod.Invoke(mospClient, new object[] { "<script>test</script>" });

            Assert.That("", Is.EqualTo(result));
        }
    }
}
