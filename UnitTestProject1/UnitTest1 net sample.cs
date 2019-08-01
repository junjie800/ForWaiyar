using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Reflection;
using System.Text.RegularExpressions;
using OfficeOpenXml;

namespace UnitTestProject1
{
    [TestClass]
    public class UnitTest1
    {
        static void Main(string[] args)
        {
            TestMethod1();
        }

        [TestMethod]
        public static void TestMethod1()
        {
            string questions;
            string columnheader;
            string newcolumnheaders;
            string answercells;
            string responsecells;
            int count = 0;
            int yescount = 0;


            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://www.nyp.edu.sg/student-life/events/2018/webchat.html");
            driver.Manage().Window.Maximize();
            IWebElement imageclick = driver.FindElement(By.XPath("//img[@src='https://asknypadmin.azurewebsites.net/BotFolder/NYPChatBotRight.png']"));
            imageclick.Click();
            IWebElement frame = driver.FindElement(By.XPath(".//iframe[@id='nypBot']"));
            driver.SwitchTo().Frame(frame);
            driver.FindElement(By.XPath("/html/body/div[1]/div/div/div[3]/div/input")).Click();

            //create a list to hold all the values
            List<string> excelData = new List<string>();
            //read the Excel file as byte array
            byte[] bin = File.ReadAllBytes(@"C:\Users\manyp\Desktop\JJ\Real JJ Project\Overall_QnA4.xlsx");
            //create a new Excel package in a memorystream
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {

                //loop all worksheets
                for (int w = 4; w <= 7; w++)   //Sheets 1 2 3 are Subject Areas, Normalized Values, EAE which we dont use.
                //foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets.ToList())
                {
                    //Console.WriteLine("TEST 1 Pass");
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[w];
                    //Delete normalized values sheet
                    Console.WriteLine("TEST 2 Pass");

                    //loop all rows
                    for (int i = worksheet.Dimension.Start.Row; i <= worksheet.Dimension.End.Row; i++)
                    {
                        //loop all columns in a row
                        for (int j = worksheet.Dimension.Start.Column; j <=
                       worksheet.Dimension.End.Column; j++)
                        {
                            //add the cell data to the List
                            if (worksheet.Cells[i, j].Value != null)
                            {
                                excelData.Add(worksheet.Cells[i, j].Value.ToString());
                            }
                            //worksheet.Cells[1, worksheet.Dimension.End.Column + 1].Value = "Response";  //may need to set as .Value.ToString()
                        }   
                    }
                    int colCount = worksheet.Dimension.End.Column;
                    int rowCount = worksheet.Dimension.End.Row;
                    worksheet.Cells[1, colCount + 1].Value = "Responses";
                    worksheet.Cells[1, colCount + 2].Value = "Timing of response retrieval";
                    worksheet.Cells[1, colCount + 3].Value = "Does the answer and response match?";
                    Console.WriteLine("Rows Count: " + (rowCount - 1));

                    var numOfRes = 1;
                    var newNumOfRes = 0;
                    var testString = "";
                    var testString2 = "";

                    for (int i = 2; i <= rowCount; i++)
                    {
                        Console.WriteLine("Worksheet name" + worksheet);
                        questions = worksheet.Cells[i,1].Text;
                        Console.WriteLine("Questions are:" + questions);
                        driver.FindElement(By.XPath("//html/body/div[1]/div/div/div[3]/div/input")).SendKeys(questions); //"Send questions"
                        driver.FindElement(By.XPath("//html/body/div[1]/div/div/div[3]/button[1]")).Click(); //click button to send the question
                        Thread.Sleep(1000); //Code passed till here so far (Checkpoint 1) (tick)

                        var textboxmsg = driver.FindElements(By.ClassName("format-markdown")); //Element whereby they give a specified answer
                        var wrongqnmsg = driver.FindElements(By.XPath("//div[@class='wc-list']")); //Element whereby they ask "do you mean this?" because no specified answer

                        count += 1;

                        newNumOfRes = textboxmsg.Count() + wrongqnmsg.Count();
                        if (newNumOfRes >= 2)
                        {
                                testString = textboxmsg.Last().GetAttribute("outerHTML");
                                Console.WriteLine("Whole chunk of response : " + testString);
                                testString2 = wrongqnmsg.Last().GetAttribute("outerHTML");
                                Console.WriteLine("Invalid scenario : " + testString2);

                        }
                        while (newNumOfRes == numOfRes)
                        {
                            Thread.Sleep(1000);
                            textboxmsg = driver.FindElements(By.ClassName("format-markdown"));
                            wrongqnmsg = driver.FindElements(By.XPath("//div[@class='wc-list']"));
                            newNumOfRes = textboxmsg.Count() + wrongqnmsg.Count();
                            try
                            {
                                testString = textboxmsg.Last().GetAttribute("outerHTML");
                                testString2 = wrongqnmsg.Last().GetAttribute("outerHTML");
                            }
                            catch
                            { }
                        }  //Code passed till here (Checkpoint 2)

                        //TRY THIS CODE OUT

                        numOfRes = newNumOfRes;
                        // foreach (var textmsg in textboxmsg)
                        // {
                        for (int c = 1; c <= colCount + 2; c++)
                        {

                            columnheader = worksheet.Cells[1, c].Text;
                            newcolumnheaders = worksheet.Cells[1, c].Text;
                            answercells = worksheet.Cells[i, 2].Text;
                            responsecells = worksheet.Cells[i, 3].Text;

                            //retrieve response with all tags then remove all the tags below
                            //try
                            //{
                            var outerhtml = testString; //figure this out later (system no element exception)
                                                        //}
                                                        //catch
                                                        //{

                            //}

                            outerhtml = outerhtml.Replace("<br />", Environment.NewLine);
                            outerhtml = Regex.Replace(outerhtml, @"<(?!a|/a|ol|ul[\x20/>])[^<>]+>", string.Empty);
                            outerhtml = outerhtml.TrimEnd('\r', '\n');  //remove carriage return
                                                                        //all to replace some to match
                            outerhtml = outerhtml.Replace("“", "\"");
                            outerhtml = outerhtml.Replace("”", "\"");
                            outerhtml = outerhtml.Replace("<ul>", "-");
                            outerhtml = outerhtml.Replace("‘", "'");
                            outerhtml = outerhtml.Replace("’", "'");

                            var outerhtml2 = wrongqnmsg.Last().GetAttribute("outerHTML");
                            outerhtml2 = outerhtml2.Replace("<br />", Environment.NewLine);
                            outerhtml2 = Regex.Replace(outerhtml2, @"<(?!a|/a|ol|ul[\x20/>])[^<>]+>", string.Empty);
                            outerhtml2 = outerhtml2.TrimEnd('\r', '\n');  //remove carriage return
                                                                          //all to replace some to match
                            outerhtml2 = outerhtml2.Replace("“", "\"");
                            outerhtml2 = outerhtml2.Replace("”", "\"");
                            outerhtml2 = outerhtml2.Replace("<ul>", "-");
                            outerhtml2 = outerhtml2.Replace("‘", "'");
                            outerhtml2 = outerhtml2.Replace("’", "'");


                            //to replace ol with numerics
                            int result = 0;
                            StringBuilder sb = new StringBuilder(outerhtml);
                            result = outerhtml.IndexOf("<ol");
                            while (result > -1)
                            {
                                if (result == outerhtml.IndexOf("<ol>"))
                                {
                                    sb.Remove(result, 4);
                                    sb.Insert(result, "1)");
                                }
                                else
                                {
                                    char number = outerhtml[result + 11];
                                    sb.Remove(result, 14);
                                    sb.Insert(result, number + ")");

                                }
                                outerhtml = sb.ToString();
                                result = outerhtml.IndexOf("<ol");
                            } //Code passed till here (Checkpoint 3)
                            //below is to remove linebreaks and whitespace for both answer and response cells to do matching
                            var compareresponsecells = Regex.Replace(outerhtml, @"\r\n?|\n", ""); //to remove line breaks for comparison
                            compareresponsecells = Regex.Replace(compareresponsecells, @"\s+", ""); //to remove whitespace for comparison
                            var compareanswercells = Regex.Replace(answercells, @"\r\n?|\n", "");
                            compareanswercells = Regex.Replace(compareanswercells, @"\s+", "");

                            //Console.WriteLine(newcolumnheaders);
                            if (columnheader == "Question")
                            {

                            }
                            else if (columnheader == "Answer")
                            {

                            }
                            else if (columnheader == "Answers")
                            {

                            } //Code passed till here (Checkpoint 4)
                            else if (newcolumnheaders == "Responses")
                            {
                                //Console.WriteLine("YES " + count);
                                try
                                {
                                    worksheet.Cells[i, c].Value = outerhtml;
                                    Console.WriteLine("SICK FEELING :" + outerhtml);
                                    var wcmessagecontented = driver.FindElements(By.XPath("//div[@class='wc-message-content']"));
                                    var lastwcmsgcontent = wcmessagecontented.Last();
                                    var child = lastwcmsgcontent.FindElement(By.XPath("./div/div")); // ./ means go down from this element
                                                                                                     //Console.WriteLine("LAST CHILD :" + child.GetAttribute("outerHTML"));
                                    if (child.GetAttribute("class").Contains("wc-list"))
                                    {
                                        worksheet.Cells[i, c].Value = outerhtml2;
                                        //Console.WriteLine("Panini");
                                    }
                                    //NewWorkSheet.Cells[i, c] = outerhtml2;
                                }
                                catch
                                {
                                } // Code passed till here (Checkpoint 5)
                                //outerhtml.Contains((char)13);
                                //Console.WriteLine(outerhtml.Contains((char)13));
                                //Console.WriteLine("WELP:" + outerhtml);

                            }
                            else
                            { }
                        }


                    } 
                    
                    excelPackage.SaveAs(new FileInfo(@"D:\New.xlsx"));
                }
            }

        }

    }
}
