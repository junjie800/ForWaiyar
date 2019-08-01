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
            string fakeresponses;
            int count = 0;
            int yescount = 0;


            IWebDriver driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://asknypadmindev.azurewebsites.net/botmain");
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

                    for (int i = 2; i <= rowCount; i++)
                    {
                        Console.WriteLine("Worksheet name" + worksheet);
                        questions = worksheet.Cells[i, 1].Text;
                        Console.WriteLine("Questions are:" + questions);
                        driver.FindElement(By.XPath("//html/body/div[1]/div/div/div[3]/div/input")).SendKeys(questions); //"Send questions"
                        driver.FindElement(By.XPath("//html/body/div[1]/div/div/div[3]/button[1]")).Click(); //click button to send the question
                        Thread.Sleep(1000); //Code passed till here so far (Checkpoint 1) (tick)
                        fakeresponses = worksheet.Cells[i, 2].Text;
                        worksheet.Cells[i, colCount + 1].Value = fakeresponses;

                        count += 1;

                        //TRY THIS CODE OUT
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




                            //Console.WriteLine(newcolumnheaders);
                            if (columnheader == "Question")
                            {

                            }
                            else if (columnheader == "Answer")
                            {

                            }
                            else if (columnheader == "Answers")
                            {

                            }
                            else if (newcolumnheaders == "Responses")
                            {
                                //Console.WriteLine("YES " + count);
                                try
                                {

                                    //NewWorkSheet.Cells[i, c] = outerhtml2;
                                }
                                catch
                                {
                                }
                                //outerhtml.Contains((char)13);
                                //Console.WriteLine(outerhtml.Contains((char)13));
                                //Console.WriteLine("WELP:" + outerhtml);

                            }
                            else if (newcolumnheaders == "Timing of response retrieval")
                            {
                                try
                                {
                                    worksheet.Cells[i, c].Value = DateTime.Now.ToString("MM/dd/yyyy HH:mm:ss");
                                }
                                catch
                                {
                                }
                            }
                            else if (newcolumnheaders == "Does the answer and response match?")
                            {
                                if (answercells.Equals(responsecells))
                                {
                                    try
                                    {
                                        worksheet.Cells[i, c].Value = "Yes";
                                    }
                                    catch
                                    {

                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        worksheet.Cells[i, c].Value = "No";
                                    }
                                    catch
                                    {
                                    }
                                }
                            }
                            else if (columnheader == null)
                            {

                            }
                            else
                            {

                                worksheet.DeleteColumn(c);
                                c--;  
                                Console.WriteLine("HEY MAN STOP IT");

                            }
                        }

                    }

                        excelPackage.SaveAs(new FileInfo(@"D:\New.xlsx"));
                    }
                }

            }

        }
    }
