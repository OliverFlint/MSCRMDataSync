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
    }
}
