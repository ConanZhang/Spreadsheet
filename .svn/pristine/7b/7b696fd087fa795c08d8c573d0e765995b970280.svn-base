// Assignment testing implementation written by Conan Zhang, u0409453
// for CS3500 Assignment #6. November, 2014.
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;


namespace GUITestProject1
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class GUITests
    {
        public GUITests()
        {
        }

        /// <summary>
        /// Tests 1 + 2 = 3 as formula.
        /// </summary>
        [TestMethod]
        public void SimpleAddition1()
        {
            this.UIMap.AdditionSetup1();
            this.UIMap.AdditionAssert1();
            this.UIMap.AdditionClosing1();
        }

        /// <summary>
        /// Tests 1 + 2 = 3 as formula.
        /// </summary>
        [TestMethod]
        public void NewOpenSaveClose1()
        {
            this.UIMap.NewOpenSaveClose1();
        }

        /// <summary>
        /// Tests circular exception.
        /// </summary>
        [TestMethod]
        public void CircularException1()
        {
            this.UIMap.CircularSetup1();
            this.UIMap.SaveCircular1();
        }

        /// <summary>
        /// Tests formula exception.
        /// </summary>
        [TestMethod]
        public void FormulaException1()
        {
            this.UIMap.FormulaExceptionSetUp1();
            this.UIMap.FormulaExceptionMove1();
            this.UIMap.FormulaExceptionAssertion1();
            this.UIMap.FormulaExceptionSetUp2();
            this.UIMap.FormulaExceptionAssertion2();
            this.UIMap.FormulaExceptionClosing1();
        }

        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;

        public UIMap UIMap
        {
            get
            {
                if ((this.map == null))
                {
                    this.map = new UIMap();
                }

                return this.map;
            }
        }

        private UIMap map;
    }
}
