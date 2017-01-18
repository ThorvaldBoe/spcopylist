using Microsoft.VisualStudio.TestTools.UnitTesting;
using SPCopyListLib;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPCopyListLib.Tests
{
    [TestClass()]
    public class SPCopyListLibTests
    {
        [ClassInitialize]
        public static void Initialize(TestContext context)
        {
            
        }

        [TestMethod()]
        public void SPCopyListAdhocTest()
        {
            var lib = SetupSPCopyListInstance();
            lib.SPCopyList(lib.SourceSite, lib.TargetSite, "TestList1");

        }

        private SPCopyListLib SetupSPCopyListInstance()
        {
            SPCopyListLib lib = new SPCopyListLib();
            string sourceSiteBase = ConfigurationManager.AppSettings["sourceSiteBase"];
            string sourceSite = sourceSiteBase + "/source1";
            string targetSite = sourceSiteBase + "/target1";
            
            string userNameSource = ConfigurationManager.AppSettings["sourceSiteUser"];
            string userNameTarget = ConfigurationManager.AppSettings["targetSiteUser"];
            string passwordSource = ConfigurationManager.AppSettings["sourceSitePassword"];
            string passwordTarget = ConfigurationManager.AppSettings["targetSitePassword"];

            lib.ConnectSource(sourceSite, userNameSource, passwordSource);
            lib.ConnectTarget(targetSite, userNameTarget, passwordTarget);

            return lib;
        }

        [TestMethod()]
        public void SPCopyListTest()
        {
            SPCopyListLib lib = new SPCopyListLib();
            string sourceSiteBase = ConfigurationManager.AppSettings["sourceSiteBase"];
            string sourceSite = sourceSiteBase + "/source1";
            string targetSite = sourceSiteBase + "/target1";
            string listName = "TestList1";

            try
            {
                lib.SPCopyList(sourceSite, targetSite, listName);
            } catch (Exception ex)
            {
                Assert.AreEqual("Source or Target context is not connected", ex.Message);
            }

            string userNameSource = ConfigurationManager.AppSettings["sourceSiteUser"];
            string userNameTarget = ConfigurationManager.AppSettings["targetSiteUser"];
            string passwordSource = ConfigurationManager.AppSettings["sourceSitePassword"];
            string passwordTarget = ConfigurationManager.AppSettings["targetSitePassword"];
            lib.ConnectSource(sourceSite, userNameSource, passwordSource);

            try
            {
                lib.SPCopyList(sourceSite, targetSite, listName);
            }
            catch (Exception ex)
            {
                Assert.AreEqual("Source or Target context is not connected", ex.Message);
            }

            lib.ConnectTarget(targetSite, userNameTarget, passwordTarget);
            var copyResult = lib.SPCopyList(sourceSite, targetSite, listName);
            if (copyResult.Conflict)
                Assert.AreEqual(SPCopyListCopyResult.SPCopyListConflictAction.Merged, copyResult.ConflictAction);
            else
                Assert.AreEqual(SPCopyListCopyResult.SPCopyListConflictAction.None, copyResult.ConflictAction);


            //the list now exists in both sites. Let's test conflict handling:

            //Abort
            copyResult = lib.SPCopyList(sourceSite, targetSite, listName, SPCopyListLib.SPCopyListCopyOptions.AbortOnDuplicate);
            Assert.AreEqual(SPCopyListCopyResult.SPCopyListConflictAction.Aborted, copyResult.ConflictAction);
            Assert.AreEqual(true, copyResult.Conflict);
            Assert.AreEqual(false, copyResult.ListCopied);

            //Merge
            copyResult = lib.SPCopyList(sourceSite, targetSite, listName, SPCopyListLib.SPCopyListCopyOptions.Merge);
            Assert.AreEqual(SPCopyListCopyResult.SPCopyListConflictAction.Merged, copyResult.ConflictAction);
            Assert.AreEqual(true, copyResult.Conflict);
            Assert.AreEqual(true, copyResult.ListCopied);

            //Overwrite
            copyResult = lib.SPCopyList(sourceSite, targetSite, listName, SPCopyListLib.SPCopyListCopyOptions.Overwrite);
            Assert.AreEqual(SPCopyListCopyResult.SPCopyListConflictAction.Overwritten, copyResult.ConflictAction);
            Assert.AreEqual(true, copyResult.Conflict);
            Assert.AreEqual(true, copyResult.ListCopied);

            //Rename
            copyResult = lib.SPCopyList(sourceSite, targetSite, listName, SPCopyListLib.SPCopyListCopyOptions.Rename);
            Assert.AreEqual(SPCopyListCopyResult.SPCopyListConflictAction.Renamed, copyResult.ConflictAction);
            Assert.AreEqual(true, copyResult.Conflict);
            Assert.AreEqual(true, copyResult.ListCopied);

        }

        [TestMethod()]
        public void CopyLookupListTest()
        {
            SPCopyListLib lib = new SPCopyListLib();
            string sourceSiteBase = ConfigurationManager.AppSettings["sourceSiteBase"];
            string sourceSite = sourceSiteBase + "/source1";
            string targetSite = sourceSiteBase + "/target1";
            string listName = "MyCustomLookup"; //a list with a lookup field

            string userNameSource = ConfigurationManager.AppSettings["sourceSiteUser"];
            string userNameTarget = ConfigurationManager.AppSettings["targetSiteUser"];
            string passwordSource = ConfigurationManager.AppSettings["sourceSitePassword"];
            string passwordTarget = ConfigurationManager.AppSettings["targetSitePassword"];
            lib.ConnectSource(sourceSite, userNameSource, passwordSource);
            lib.ConnectTarget(targetSite, userNameTarget, passwordTarget);

            //TODO: Add handling for lookup fields
            //var copyResult = lib.SPCopyList(sourceSite, targetSite, listName);
        }

        [TestMethod()]
        public void ConnectSourceTest()
        {
            SPCopyListLib lib = new SPCopyListLib();
            Assert.IsNull(lib.SourceContext);
            Assert.IsNull(lib.TargetContext);

            string sourceSiteBase = ConfigurationManager.AppSettings["sourceSiteBase"];
            string user = ConfigurationManager.AppSettings["sourceSiteUser"];
            string password = ConfigurationManager.AppSettings["sourceSitePassword"];
            lib.ConnectSource(sourceSiteBase, user, password);

            Assert.IsNotNull(lib.SourceContext);
            Assert.AreEqual(sourceSiteBase, lib.SourceSite);
        }

        [TestMethod()]
        public void ConnectTargetTest()
        {
            SPCopyListLib lib = new SPCopyListLib();
            Assert.IsNull(lib.TargetContext);
            Assert.IsNull(lib.TargetContext);

            string targetSiteBase = ConfigurationManager.AppSettings["targetSiteBase"];
            string user = ConfigurationManager.AppSettings["targetSiteUser"];
            string password = ConfigurationManager.AppSettings["targetSitePassword"];
            lib.ConnectTarget(targetSiteBase, user, password);

            Assert.IsNotNull(lib.TargetContext);
            Assert.AreEqual(targetSiteBase, lib.TargetSite);


        }
    }
}