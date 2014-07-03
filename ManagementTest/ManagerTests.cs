using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using VSMacros;
using VSMacros.Engines;

/* BEHAVIORS TO BE CHECKED
 *      - 
*/
namespace ManagementTest
{
    [TestClass]
    public class ManagerTests
    {
        //[TestMethod]
        public void Instantiation_LoadsShortcut()
        {
            Assert.IsNotNull(Manager.Instance.Shortcuts);
        }
    }
}
