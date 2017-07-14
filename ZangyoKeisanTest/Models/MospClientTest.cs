using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
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
            string res = await model.downloadExcel();

            Console.WriteLine(res);

            Assert.That(res, Is.EqualTo("test"));
        }
    }
}
