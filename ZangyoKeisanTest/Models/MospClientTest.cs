using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ZangyoKeisan.Models;
using RichardSzalay.MockHttp;
using System.Net.Http;


namespace ZangyoKeisanTest.Models
{
    class MospClientTest
    {
        [Test]
        public void procSeq取得()
        {
            // Mospから取得したHTML
            string html = "<script>var procSeq = \"9\";</script>";

            // private メソッドをテストする
            MospClient mospClient = new MospClient();
            Type type = mospClient.GetType();
            MethodInfo oMethod = type.GetMethod("getProcSeq", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            string result =  (string)oMethod.Invoke(mospClient, new object[] { html });

            Assert.That("9", Is.EqualTo(result));
        }

        [Test]
        public void procSeq取得_失敗()
        {
            // Mospから取得したHTML
            string html = "<script>test</script>";

            MospClient mospClient = new MospClient();
            Type type = mospClient.GetType();
            MethodInfo oMethod = type.GetMethod("getProcSeq", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            string result = (string)oMethod.Invoke(mospClient, new object[] { html });

            Assert.That("", Is.EqualTo(result));
        }
    }
}
