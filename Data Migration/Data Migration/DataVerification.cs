    using Data_Migration;
    using NUnit.Framework;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Support.UI;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.OleDb;
    using System.Linq;
    using System.Text;
    using Actions = OpenQA.Selenium.Interactions.Actions;

    namespace Data_Migration.Data_Migration

{

        [TestFixture]

        public class DataVerification : TestBase
    {

            //Policy-Servicing

            private string sheet;



            [SetUp]

            public void startBrowser()

            {

                _driver = TestBase.SiteConnection();
                sheet = "Policy-Verification";

            }



            [Test, Order(1)]

            public void PolicyServicingTestSuite()

            {

                using (OleDbConnection conn = new OleDbConnection(_connString))
                {
                    try
                    {

                        // Open connection
                        conn.Open();
                        string cmdQuery = "SELECT * FROM [" + sheet + "$]";

                        OleDbCommand cmd = new OleDbCommand(cmdQuery, conn);

                        // Create new OleDbDataAdapter
                        OleDbDataAdapter oleda = new OleDbDataAdapter();

                        oleda.SelectCommand = cmd;

                        // Create a DataSet which will hold the data extracted from the worksheet.
                        DataSet ds = new DataSet();

                        // Fill the DataSet from the data extracted from the worksheet.
                        oleda.Fill(ds, "Policies");



                        foreach (var row in ds.Tables[0].DefaultView)
                        {
                            var contractRef = ((System.Data.DataRowView)row).Row.ItemArray[0].ToString(); ;
                            var func = ((System.Data.DataRowView)row).Row.ItemArray[1].ToString();

                            if (contractRef != "" && func != "")
                            {
                                Delay(8);
                                clickOnMainMenu();
                                try
                                {
                                    switch (func)
                                    {
                                        case "Dataverification":
                                        Dataverification(contractRef);
                                            break;

                                        default:
                                            break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    // TakeScreenshot(_driver, $@"{_screenShotFolder}\Failed_Scenarios\", func);
                                    var errMsg = ex.Message;
                                    StringBuilder error = new StringBuilder();
                                    var charsToRemove = new char[] { '@', ',', '.', ';', '"', '{', '}' };
                                    var stringsToRemove = new String[] { "'" };


                                    foreach (char c in errMsg)
                                    {
                                        var isInvalidChar = charsToRemove.Contains(c);
                                        var isInvalidString = stringsToRemove.Contains(c.ToString());
                                        if (!isInvalidChar && !isInvalidString)
                                        {
                                            error.Append(c);
                                        }
                                    }

                                    cmd = conn.CreateCommand();

                                    var testDate = DateTime.Now.ToString();
                                    var errorMsg = error.ToString();
                                    if (errorMsg.Length > 250)
                                    {
                                        errorMsg = errorMsg.Substring(0, 250);
                                    }

                                    //Test_Date
                                    cmd.CommandText = $"UPDATE [{sheet}$] SET Test_Date = '{testDate}' WHERE Function = '{func}';";
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandText = $"UPDATE [{sheet}$] SET Test_Results  = 'Failed' WHERE Function = '{func}';";
                                    cmd.ExecuteNonQuery();
                                    cmd.CommandText = $"UPDATE [{sheet}$] SET Comments  = '{errorMsg}' WHERE Function = '{func}';";
                                    cmd.ExecuteNonQuery();
                                }

                            }

                        }


                    }
                    finally
                    {
                        conn.Close();
                        conn.Dispose();
                    }
                }




            }
            private void clickOnMainMenu()
            {
                try
                {
                    //find the contract search option
                    var search = _driver.FindElement(By.XPath("//*[@id='t0_764']/a"));
                }
                catch
                {
                    //clickOnMainMenu
                    _driver.FindElement(By.Name("CBWeb")).Click();
                }
            }

            public void Dataverification(string contractRef)



        {



            String test_url_1 = "http://ilr-int.safrican.co.za/web/wspd_cgi.sh/WService=wsb_ilrint/run.w?";
            String test_url_1_title = "MIP - Sanlam ARL - Warpspeed Lookup Window";
            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;




            string results = "";



            string date = DateTime.Today.ToString("g");




            policySearch(contractRef);



            //Contract Status validation
            Delay(2);
            SetproductName("ReInstate");
            var Cancelled = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;





            IWebElement policyOptionElement3 = _driver.FindElement(By.XPath("//*[@id='m0i0o1']"));




            //Creating object of an Actions class
            Actions action2 = new Actions(_driver);





            //Performing the mouse hover action on the target element.
            action2.MoveToElement(policyOptionElement3).Perform();
            Delay(2);



            //Click on Reinstate
            _driver.FindElement(By.XPath("//*[@id='m0t0']/tbody/tr[8]/td/div/div[3]/a/img")).Click();
            Delay(2);

            SelectElement oSelect2 = new SelectElement(_driver.FindElement(By.Name("frmReason")));
            oSelect2.SelectByValue("ReinstateReason2");
            Delay(2);


            //Click submit
            _driver.FindElement(By.Name("btnctcrereinstatecsu5")).Click();
            Delay(4);




            //Click submit
            _driver.FindElement(By.Name("btnctcrereinstatecsu2")).Click();
            Delay(5);





            //Contract Status validation



            var StatusInForce = _driver.FindElement(By.XPath("//*[@id='CntContentsDiv8']/table/tbody/tr[2]/td[2]/u/font")).Text;





            Assert.IsTrue(StatusInForce.Equals("In-Force", StringComparison.CurrentCultureIgnoreCase));



            Delay(3);



            if (StatusInForce == "In-Force")
            {
                results = "Passed";
            }
            else
            {
                results = "Failed";
            }



            base.writeResultsToExcell(results, sheet, "ReInstate");
        }

        [Category("Policy Search")]
        public void policySearch(string contractRef)
        {

            IJavaScriptExecutor js = (IJavaScriptExecutor)_driver;


            //Click on contract search 
            _driver.FindElement(By.Name("alf-ICF8_00000222")).Click();
            Delay(2);




            //Type in contract ref 

            _driver.FindElement(By.Name("frmContractReference")).SendKeys(contractRef);



            Delay(4);

            //Click on Search Icon 
            _driver.FindElement(By.Name("btncbcts0")).Click();

            Delay(5);
            _driver.FindElement(By.Name("hl_960735872.488")).Click();

            Delay(5);

        }


        [TearDown]

        public void closeBrowser()

        {

            base.DisconnectBrowser();

        }



    }

}

