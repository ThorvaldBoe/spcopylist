using Microsoft.VisualStudio.TestTools.UnitTesting;
using SPCopyListLib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPCopyListLib.Tests
{
    [TestClass()]
    public class SPCopyListLibTests
    {
        [TestMethod()]
        public void SPCopyListTest()
        {
            SPCopyListLib lib = new SPCopyListLib();
            lib.SPCopyList("", "", "");

        }
    }
}