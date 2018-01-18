using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace MSCRMDataSyncTests
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            MSCRMDataSync.Program.Main(new string[] { "Test1.xml" });
        }

        [TestMethod]
        public void TestMethod2()
        {
            MSCRMDataSync.Program.Main(new string[] { "Test2.xml" });
        }

        [TestMethod]
        public void TestMethod3()
        {
            MSCRMDataSync.Program.Main(new string[] { "Test3.xml" });
        }
    }
}
