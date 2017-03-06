using Microsoft.VisualStudio.TestTools.UnitTesting;
using WindowsFormsApplication2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApplication2.Tests
{
    [TestClass()]
    public class Form1Tests
    {
        [TestMethod()]
        public void getStringTest()
        {
            Assert.IsTrue(Form1.something.getString("12345") == "12");
            Assert.Fail();
        }
    }
}