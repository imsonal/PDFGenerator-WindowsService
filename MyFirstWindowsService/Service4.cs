using iTextSharp.text;
using iTextSharp.text.html;
using iTextSharp.text.pdf;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Configuration;
using MyFirstWindowsService.Model;
using MyFirstWindowsService.Service;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.Web.UI.DataVisualization.Charting;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using Paragraph = iTextSharp.text.Paragraph;
using Path = System.IO.Path;
using Rectangle = iTextSharp.text.Rectangle;
using DocumentITextPDF = iTextSharp.text.Document;
using FontITP = iTextSharp.text.Font;
using ImageITP = iTextSharp.text.Image;
using System.Drawing.Imaging;
using ceTe.DynamicPDF.Rasterizer;
using ImageFormat = ceTe.DynamicPDF.Rasterizer.ImageFormat;
using System.Web.UI.DataVisualization.Charting;
using Font = System.Drawing.Font;
using Color = System.Drawing.Color;
using Image = iTextSharp.text.Image;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Globalization;
using System.Runtime.ConstrainedExecution;
using MyFirstWindowsService.Model;
using MyFirstWindowsService.Services;
using System;
using WorkerServiceDemoProject.Model;
using System.Collections.Generic;
using System.IO;
using Microsoft.AspNetCore.Http;
using System.Data.SqlClient;

namespace MyFirstWindowsService
{
    public partial class Service4 : ServiceBase
    {
        private readonly IConfiguration _configuration;
        private readonly IHostingEnvironment _hostingEnvironment;
        private readonly HttpClient _httpClient;
        private readonly IHttpContextAccessor _httpContextAccessor;

        public Service4(IHttpContextAccessor httpContextAccessor, IConfiguration configuration, HttpClient httpClient, IHostingEnvironment hostingEnvironment)
        {
            _httpClient = httpClient;
            _httpContextAccessor = httpContextAccessor;
            _configuration = configuration;
            _hostingEnvironment = hostingEnvironment;
            // InitializeComponent();
        }


        Timer timer = new Timer(); // name space(using System.Timers;)

        private static string con = "Data Source=103.86.177.2;Initial Catalog=WaltCapitalDB_DEV;User Id=WCMDEV2022;Password=WCM@2022;";
        //ConfigurationManager.ConnectionStrings["ABCD"].ToString();

        private SqlConnection sn = new SqlConnection(con);

        private SqlCommand sm;
        public Service4()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            timer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            timer.Interval = 5000; //number in milisecinds
            timer.Enabled = true;
        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Before Service Call at  " + DateTime.Now);
            GenerateSummaryReportBySummaryDetailsAsync();
            //InsertRecord();
            //TraceService("Success");
            //WriteToFile("Service is recall at " + DateTime.Now);
            WriteToFile("After Service Call at " + DateTime.Now);
        }

        public void WriteToFile(string Message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = AppDomain.CurrentDomain.BaseDirectory + "\\Logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";

            if (!File.Exists(filepath))
            {
                // Create a file to write to. 
                using (StreamWriter sw = File.CreateText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filepath))
                {
                    sw.WriteLine(Message);
                }
            }
        }


        private void InsertRecord()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection("Server=192.168.1.199,1433;Database=WaltCapitalDB;Trusted_Connection=false;User Id=sa;Password=sa@2022;"))

                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        command.Connection = connection;
                        command.CommandText = "Insert into .....')";
                        connection.Open();
                        command.ExecuteNonQuery();
                        connection.Close();
                    }
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void TraceService(string content)
        {


            sn.Open();
            // sm = new SqlCommand("Update_Table", sn);
            //  sm.CommandType = CommandType.StoredProcedure;
            try
            {
                //    sm.Parameters.AddWithValue("@value", "0");
                //    sm.ExecuteNonQuery();
            }
            catch
            {
                throw;
            }
            finally
            {
                sm.Dispose();
                sn.Close();
                sn.Dispose();
            }


            //set up a filestream
            FileStream fs = new FileStream(@"d:\MeghaService.txt", FileMode.OpenOrCreate, FileAccess.Write);

            //set up a streamwriter for adding text
            StreamWriter sw = new StreamWriter(fs);

            //find the end of the underlying filestream
            sw.BaseStream.Seek(0, SeekOrigin.End);

            //add the text
            sw.WriteLine(content);
            //add the text to the underlying filestream

            sw.Flush();
            //close the writer
            sw.Close();
        }
        public string GetRootPath()
        {
            string path = "";
            string ServerDomain = _httpContextAccessor.HttpContext.Request.Host.Value;
            ServerDomain = "";
            var FileBaseURL = Convert.ToString(_configuration["FileBaseURL"]);
            if (!string.IsNullOrWhiteSpace(ServerDomain))
            {
                path = ServerDomain;
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(FileBaseURL))
                {
                    path = FileBaseURL;
                }
                else
                {
                    path = _hostingEnvironment.ContentRootPath;
                }
            }
            return path;
        }

        public Byte[] GenerateGraphImage()
        {
            var chart = new System.Web.UI.DataVisualization.Charting.Chart
            {
                Width = 700,
                Height = 450,
                RenderType = RenderType.ImageTag,
                AntiAliasing = AntiAliasingStyles.All,
                TextAntiAliasingQuality = TextAntiAliasingQuality.High
            };

            chart.Titles.Add("Portfolio Performance");
            chart.Titles[0].Font = new Font("Arial", 20f, FontStyle.Bold);
            chart.Titles[0].Alignment = ContentAlignment.MiddleLeft;

            chart.ChartAreas.Add("");
            chart.ChartAreas[0].AxisX.Title = "";
            chart.ChartAreas[0].AxisY.Title = "";
            chart.ChartAreas[0].AxisX.LabelStyle.Font = new Font("Arial", 10f);
            //chart.ChartAreas[0].AxisX.LabelStyle.Angle = -90;
            chart.ChartAreas[0].BackColor = Color.White;
            chart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
            chart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;

            chart.Series.Add("");
            chart.Series[0].MarkerBorderWidth = 0;
            chart.Series[0].ChartType = SeriesChartType.Spline;
            chart.Series[0].Color = Color.SteelBlue;

            chart.Series[0].BorderWidth = 3;

            chart.Series.Add("");
            chart.Series[1].MarkerBorderWidth = 0;
            chart.Series[1].ChartType = SeriesChartType.Spline;
            chart.Series[1].Color = Color.Black;
            chart.Series[1].BorderWidth = 3;

            List<UserDataModel> userDataModel = new List<UserDataModel>();

            for (int i = 0; i < 12; i++)
            {
                UserDataModel UserData = new UserDataModel();
                UserData.Id = i + 1;
                UserData.Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i + 1);
                UserData.NoOfOrders = Convert.ToDouble(GenerateDyanicVal("double", 130));
                UserData.NoOfOrdersNet = Convert.ToDouble(GenerateDyanicVal("double", 130));
                UserData.NoOfOrdersTotal = Convert.ToDouble(GenerateDyanicVal("double", 130));
                userDataModel.Add(UserData);
            }

            foreach (var q in userDataModel)
            {
                chart.Series[0].Points.AddXY(q.Name, Convert.ToDouble(q.NoOfOrders));
                chart.Series[1].Points.AddXY(q.Name, Convert.ToDouble(q.NoOfOrdersTotal));
            }

            using (var chartimage = new MemoryStream())
            {
                chart.SaveImage(chartimage, ChartImageFormat.Png);
                var data = chartimage.GetBuffer();
                return data;
            }

        }

        public string GenerateDyanicVal(string generateType, int range)
        {

            string val = "";
            Random r = new Random();
            if (generateType == "int")
            {
                int rInt = r.Next(80, range);
                val = rInt.ToString();
            }

            //for doubles
            if (generateType == "double")
            {
                double rDouble = r.NextDouble() * range;
                val = rDouble.ToString();
            }
            return val;
        }

        public string GetPhysicalRootPath()
        {
            WriteToFile("Started Finding Physical Path - " + DateTime.Now);
            string path = "";
            //string ServerDomain = _httpContextAccessor.HttpContext.Request.Host.Value;
            string ServerDomain = "";
            var FileBaseURL = Convert.ToString(_configuration["FileBaseURL"]);
            FileBaseURL = "";
            if (!string.IsNullOrWhiteSpace(ServerDomain))
            {
                path = ServerDomain;
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(FileBaseURL))
                {
                    // path = FileBaseURL;
                    path = @"D:\Files";
                }
                else
                {
                    // path = _hostingEnvironment.ContentRootPath;
                    path = @"D:\Files";
                }
            }
            WriteToFile("Finding Physical Path Successfully" + DateTime.Now);
            return path;
        }


        public string GenerateSummaryReportBySummaryDetailsAsync()
        {
            try
            {
                WriteToFile("Started Creating document PDF  " + DateTime.Now);
                //sn.Open();
                //sm = new SqlCommand("FactSheetMst", sn);
                //sm.CommandType = CommandType.TableDirect;

                //sm.Parameters.InitializeLifetimeService();

                //sm.ExecuteNonQuery();

                // sm = new SqlCommand("Update_Table", sn);
                //  sm.CommandType = CommandType.StoredProcedure;

                //    sm.Parameters.AddWithValue("@value", "0");
                //    sm.ExecuteNonQuery();
                using (SqlConnection connection = new SqlConnection("Server=103.86.177.2;Database=WaltCapitalDB_DEV;Trusted_Connection=false;User Id=WCMDEV2022;Password=WCM@2022;"))

                {
                    using (SqlCommand command = new SqlCommand())
                    {
                        WriteToFile("Database Connection Successfully  " + DateTime.Now);
                        command.Connection = connection;
                        DataSet ds = new DataSet();
                        SqlDataAdapter sda = new SqlDataAdapter();
                        //  command.CommandText = "Insert into .....')";
                        //connection.Open();

                       // command.CommandText = "select PortfolioManager from FactSheetMst where FundId = 15 ";
                        string filePath = string.Empty;
                        string pdfFileName = string.Empty;
                        int width = 4500;
                        int tempHeight = 10200;
                        iTextSharp.text.Rectangle pageSize = new iTextSharp.text.Rectangle(width, tempHeight);
                        DocumentITextPDF document = new DocumentITextPDF();
                        document.SetMargins(15, 15, 12, 12);
                        //document.SetPageSize(iTextSharp.text.PageSize.A4.Rotate());

                        var SummaryR = FontFactory.GetFont("Poppins", 18, FontITP.NORMAL, WebColors.GetRGBColor("#bfbfbf"));
                        var FR = FontFactory.GetFont("Poppins", 12, FontITP.NORMAL, BaseColor.BLACK);
                        var SR = FontFactory.GetFont("Poppins", 10, FontITP.NORMAL, BaseColor.GRAY);
                        var THF = FontFactory.GetFont("Arial", 8, FontITP.BOLD, BaseColor.BLACK);
                        var TDF = FontFactory.GetFont("Arial", 8, FontITP.NORMAL, BaseColor.BLACK);
                        var TblHeaderFont = FontFactory.GetFont("Arial", 8, FontITP.BOLD, BaseColor.WHITE);


                        using (System.IO.MemoryStream memoryStream = new System.IO.MemoryStream())
                        {
                            WriteToFile("Started Creating document PDF at-2   " + DateTime.Now);
                            PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                            document.Open();
                            BaseColor backgroundColor = WebColors.GetRGBColor("#ffffff");
                            BaseColor tableMainHeaderColor = WebColors.GetRGBColor("#0b6aab");

                            BaseColor tableHeaderColor = WebColors.GetRGBColor("#eaeff7");
                            BaseColor tableHeaderColor2 = WebColors.GetRGBColor("#d2ddef");

                            BaseColor evenCellColor = WebColors.GetRGBColor("#d2ddef");
                            BaseColor performaceTblColor = WebColors.GetRGBColor("#5b9bd5");
                            PdfContentByte content = writer.DirectContentUnder;

                            //main table create
                            PdfPTable MainTable = new PdfPTable(12);
                            MainTable.WidthPercentage = 100;

                            FontITP fontSubHeaderBlue = FontFactory.GetFont("Arial", 10, FontITP.BOLD, WebColors.GetRGBColor("#0b6aab"));
                            FontITP fontSubHeaderBlueNormal = FontFactory.GetFont("Arial", 9, FontITP.NORMAL, WebColors.GetRGBColor("#0b6aab"));

                            FontITP fontBody = FontFactory.GetFont("Arial", 8, FontITP.NORMAL, BaseColor.BLACK);
                            FontITP fontBodyRed = FontFactory.GetFont("Arial", 8, FontITP.NORMAL, BaseColor.RED);
                            FontITP fontBodyBold = FontFactory.GetFont("Arial", 8, FontITP.BOLD, BaseColor.BLACK);


                            FontITP fontSummary = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 24, FontITP.NORMAL, BaseColor.BLACK);
                            FontITP fontCHWTitle = FontFactory.GetFont("Arial", 9, FontITP.NORMAL, BaseColor.BLACK);
                            FontITP fontSubHeader = FontFactory.GetFont("Arial", 10, FontITP.NORMAL, BaseColor.BLACK);
                            var fontTableHeader = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 12, FontITP.NORMAL, BaseColor.BLACK);
                            var fontTableRow = FontFactory.GetFont(BaseFont.TIMES_ROMAN, 10, FontITP.NORMAL, BaseColor.GRAY);

                            PdfPTable table = new PdfPTable(12);
                            table.WidthPercentage = 100;

                            table.HorizontalAlignment = Rectangle.ALIGN_LEFT;
                            WriteToFile("Finding Path  " + DateTime.Now);

                            ////string? logoFilename1 = logoFilename;
                            string logoFilename = "logo walt.png";
                            //  string rootPath = GetPhysicalRootPath();
                            WriteToFile("Finding Path -1  " + DateTime.Now);

                            // var logoPath = Path.Combine(rootPath, "wwwroot", "Files", "AppLogo", logoFilename);
                            var logoPath = Path.Combine(@"D:\Files\AppLogo", logoFilename);

                            WriteToFile("Finding Applogo Successfully   " + DateTime.Now);

                            //var URL = new Uri(logoPath);
                            ImageITP Img = ImageITP.GetInstance(logoPath);

                            PdfPCell cell2 = new PdfPCell(Img);
                            cell2.Colspan = 3;
                            cell2.FixedHeight = 30;
                            cell2.Border = Rectangle.NO_BORDER;
                            cell2.PaddingTop = 7;
                            cell2.PaddingBottom = 10;
                            cell2.PaddingRight = 0;
                            cell2.HorizontalAlignment = Rectangle.ALIGN_LEFT;
                            table.AddCell(cell2);


                            WriteToFile("Global Portfolio (Rand)  " + DateTime.Now);
                            PdfPCell cellHeader = new PdfPCell(new Phrase("Walt Capital Management Global Portfolio (Rand)\r\n", fontCHWTitle));
                            cellHeader.Colspan = 5;
                            cellHeader.Border = Rectangle.NO_BORDER;
                            cellHeader.HorizontalAlignment = Rectangle.ALIGN_CENTER;
                            cellHeader.PaddingTop = 5;
                            cellHeader.PaddingBottom = 10;
                            table.AddCell(cellHeader);

                            PdfPCell headerLine = new PdfPCell(new Phrase("", fontCHWTitle));
                            headerLine.Colspan = 4;
                            headerLine.Border = Rectangle.NO_BORDER;
                            headerLine.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                            headerLine.PaddingTop = 5;
                            headerLine.PaddingBottom = 10;

                            table.AddCell(headerLine);

                            Paragraph p = new Paragraph(new Chunk(new iTextSharp.text.pdf.draw.LineSeparator(0.0F, 100.0F, BaseColor.BLACK, Element.ALIGN_LEFT, 1)));

                            document.Add(table);

                            WriteToFile("Monthly Portfolio Fact Sheet -2" + DateTime.Now);

                            PdfPTable tableLeftContent = new PdfPTable(6);
                            tableLeftContent.WidthPercentage = 50;
                            tableLeftContent.HorizontalAlignment = Rectangle.ALIGN_LEFT;

                            PdfPCell cellLeftDesc = new PdfPCell(new Phrase("Monthly Portfolio Fact Sheet\r\n", fontSubHeader));
                            cellLeftDesc.Colspan = 6;
                            cellLeftDesc.Border = Rectangle.NO_BORDER;
                            cellLeftDesc.HorizontalAlignment = Rectangle.ALIGN_LEFT;
                            cellLeftDesc.PaddingTop = 20;
                            tableLeftContent.AddCell(cellLeftDesc);

                            PdfPCell cellLeftDesc1 = new PdfPCell(new Phrase("As at 01 November 2022\r\n", fontSubHeaderBlueNormal));
                            cellLeftDesc1.Colspan = 6;
                            cellLeftDesc1.Border = Rectangle.NO_BORDER;
                            cellLeftDesc1.HorizontalAlignment = Rectangle.ALIGN_LEFT;
                            tableLeftContent.AddCell(cellLeftDesc1);

                            PdfPCell investmentObj = new PdfPCell(new Phrase("Investment objective\r\n", fontSubHeaderBlue));
                            investmentObj.PaddingTop = 10;
                            investmentObj.Colspan = 6;
                            investmentObj.Border = Rectangle.NO_BORDER;

                            tableLeftContent.AddCell(investmentObj);

                            PdfPCell investmentObjInfo = new PdfPCell(new Phrase("No Objective.\r\n", fontBody));
                            investmentObjInfo.Colspan = 6;
                            investmentObjInfo.Border = Rectangle.NO_BORDER;
                            tableLeftContent.AddCell(investmentObjInfo);

                            PdfPCell keyFacts = new PdfPCell(new Phrase("Key Facts\r\n", fontSubHeaderBlue));
                            keyFacts.Border = Rectangle.NO_BORDER;
                            keyFacts.Colspan = 6;
                            keyFacts.PaddingTop = 10;
                            tableLeftContent.AddCell(keyFacts);

                            PdfPCell keyFactsInfo = new PdfPCell(new Phrase("Portfolio manager : \r\n", fontBodyBold));
                            keyFactsInfo.Border = Rectangle.NO_BORDER;
                            keyFactsInfo.Colspan = 2;
                            tableLeftContent.AddCell(keyFactsInfo);

                            WriteToFile("Before Portfolio Manager Value Generated Successfully" + DateTime.Now);

                            command.CommandText = "select PortfolioManager from FactSheetMst where FundId = 15 ";
                            var tableredirect = CommandType.TableDirect.ToString("select PortfolioManager from FactSheetMst where FundId = 15 ");
                           // CommandType.TableDirect = "select PortfolioManager from FactSheetMst where FundId = 15 ";

                            string PortfolioManagerValue = command.CommandText;
                          //PdfPCell keyFactsInfoData = new PdfPCell(new Phrase("PortfolioManagerValue \r\n", fontBody));
                          PdfPCell keyFactsInfoData = new PdfPCell(new Phrase("select PortfolioManager from FactSheetMst where FundId = 15 ", fontBody));
                            connection.Open();
                            command.ExecuteNonQuery();

                            //if (cmdText == CommandType.TableDirect) //Type: Table Direct
                            //{
                            //    cmd.CommandType = CommandType.Text;
                            //    cmd.CommandText = Query;
                            //}


                            WriteToFile("Portfolio Manager Value Generated Successfully" + DateTime.Now);

                            keyFactsInfoData.Border = Rectangle.NO_BORDER;
                            keyFactsInfoData.HorizontalAlignment = Rectangle.ALIGN_LEFT;
                            keyFactsInfoData.PaddingLeft = 0;
                            keyFactsInfoData.Colspan = 4;
                            tableLeftContent.AddCell(keyFactsInfoData);

                            PdfPCell KFIEmail = new PdfPCell(new Phrase("Email : ", fontBodyBold));
                            KFIEmail.Border = Rectangle.NO_BORDER;
                            KFIEmail.Colspan = 1;
                            KFIEmail.PaddingRight = 0;
                            tableLeftContent.AddCell(KFIEmail);

                            PdfPCell KFIEmailData = new PdfPCell(new Phrase("tony@gmail.com", fontBody));
                            KFIEmailData.PaddingLeft = 0;
                            KFIEmailData.Border = Rectangle.NO_BORDER;
                            KFIEmailData.Colspan = 5;
                            tableLeftContent.AddCell(KFIEmailData);

                            PdfPCell KFIFSPNo = new PdfPCell(new Phrase("Walt Capital Managment FSP Number : \r\n", fontBodyBold));
                            KFIFSPNo.Border = Rectangle.NO_BORDER;
                            KFIFSPNo.Colspan = 4;
                            tableLeftContent.AddCell(KFIFSPNo);

                            PdfPCell KFIFSPNoDetail = new PdfPCell(new Phrase("123123123.5454\r\n", fontBody));
                            KFIFSPNoDetail.Border = Rectangle.NO_BORDER;
                            KFIFSPNoDetail.Colspan = 2;
                            tableLeftContent.AddCell(KFIFSPNoDetail);

                            PdfPCell KFITel = new PdfPCell(new Phrase("Tel : \r\n", fontBodyBold));
                            KFITel.Border = Rectangle.NO_BORDER;
                            KFITel.Colspan = 1;
                            tableLeftContent.AddCell(KFITel);

                            PdfPCell KFITelNo = new PdfPCell(new Phrase("+27 43524878787\r\n", fontBody));
                            KFITelNo.Border = Rectangle.NO_BORDER;
                            KFITelNo.Colspan = 5;
                            tableLeftContent.AddCell(KFITelNo);

                            PdfPCell KFInceptionDate = new PdfPCell(new Phrase("Inception Date : \r\n", fontBodyBold));
                            KFInceptionDate.Border = Rectangle.NO_BORDER;
                            KFInceptionDate.Colspan = 2;
                            tableLeftContent.AddCell(KFInceptionDate);

                            PdfPCell KFInceptionDateDetail = new PdfPCell(new Phrase("27 October 2022\r\n", fontBody));
                            KFInceptionDateDetail.Border = Rectangle.NO_BORDER;
                            KFInceptionDateDetail.Colspan = 4;
                            tableLeftContent.AddCell(KFInceptionDateDetail);

                            PdfPCell KFISector = new PdfPCell(new Phrase("Sector : \r\n", fontBodyBold));
                            KFISector.Border = Rectangle.NO_BORDER;
                            KFISector.Colspan = 1;
                            tableLeftContent.AddCell(KFISector);

                            PdfPCell KFISectorDetail = new PdfPCell(new Phrase("Sector 51\r\n", fontBody));
                            KFISectorDetail.Border = Rectangle.NO_BORDER;
                            KFISectorDetail.Colspan = 5;
                            tableLeftContent.AddCell(KFISectorDetail);

                            PdfPCell KFITargetReturns = new PdfPCell(new Phrase("Target Returns : \r\n", fontBodyBold));
                            KFITargetReturns.Border = Rectangle.NO_BORDER;
                            KFITargetReturns.Colspan = 2;
                            tableLeftContent.AddCell(KFITargetReturns);

                            PdfPCell KFITargetReturnDetail = new PdfPCell(new Phrase("12.343\r\n", fontBody));
                            KFITargetReturnDetail.Border = Rectangle.NO_BORDER;
                            KFITargetReturnDetail.Colspan = 4;
                            tableLeftContent.AddCell(KFITargetReturnDetail);

                            PdfPCell KFIPps = new PdfPCell(new Phrase("Possible participatory structures: \r\n", fontBodyBold));
                            KFIPps.Border = Rectangle.NO_BORDER;
                            KFIPps.Colspan = 3;
                            tableLeftContent.AddCell(KFIPps);

                            PdfPCell KFIPpsDetail = new PdfPCell(new Phrase(" \r\n", fontBody));
                            KFIPpsDetail.Border = Rectangle.NO_BORDER;
                            KFIPpsDetail.Colspan = 3;
                            tableLeftContent.AddCell(KFIPpsDetail);

                            PdfPCell KFIInvesterInvest = new PdfPCell(new Phrase("Investor may invest in No\r\n", fontBody));
                            KFIInvesterInvest.Border = Rectangle.NO_BORDER;
                            KFIInvesterInvest.Colspan = 6;
                            tableLeftContent.AddCell(KFIInvesterInvest);

                            PdfPCell KFIMinInvest = new PdfPCell(new Phrase("Minimum Investment : Min : \r\n", fontBodyBold));
                            KFIMinInvest.Border = Rectangle.NO_BORDER;
                            KFIMinInvest.Colspan = 3;
                            tableLeftContent.AddCell(KFIMinInvest);

                            PdfPCell KFIMinInvestDetail = new PdfPCell(new Phrase("Rand 10,002.8687875\r\n", fontBody));
                            KFIMinInvestDetail.Border = Rectangle.NO_BORDER;
                            KFIMinInvestDetail.Colspan = 3;
                            tableLeftContent.AddCell(KFIMinInvestDetail);

                            PdfPCell KFIMinInvestRecommend = new PdfPCell(new Phrase("Recommend : \r\n", fontBodyBold));
                            KFIMinInvestRecommend.Border = Rectangle.NO_BORDER;
                            KFIMinInvestRecommend.Colspan = 2;
                            tableLeftContent.AddCell(KFIMinInvestRecommend);

                            PdfPCell KFIMinInvestRecommendDetail = new PdfPCell(new Phrase("Rand 2002454\r\n", fontBody));
                            KFIMinInvestRecommendDetail.Border = Rectangle.NO_BORDER;
                            KFIMinInvestRecommendDetail.Colspan = 4;
                            tableLeftContent.AddCell(KFIMinInvestRecommendDetail);

                            PdfPCell feeCalculations = new PdfPCell(new Phrase("Fee and Calculations \r\n", fontSubHeaderBlue));
                            feeCalculations.Border = Rectangle.NO_BORDER;
                            feeCalculations.Colspan = 6;
                            feeCalculations.PaddingTop = 10;
                            tableLeftContent.AddCell(feeCalculations);

                            PdfPCell baseFee = new PdfPCell(new Phrase("Base Fee : 0\r\n", fontBodyBold));
                            baseFee.Border = Rectangle.NO_BORDER;
                            baseFee.Colspan = 6;
                            tableLeftContent.AddCell(baseFee);

                            PdfPCell feeHurdle = new PdfPCell(new Phrase("Fee Hurdle : 12.09092\r\n", fontBodyBold));
                            feeHurdle.Border = Rectangle.NO_BORDER;
                            feeHurdle.Colspan = 6;
                            tableLeftContent.AddCell(feeHurdle);

                            PdfPCell sharingRatio = new PdfPCell(new Phrase("Sharing Ration : 0\r\n", fontBodyBold));
                            sharingRatio.Border = Rectangle.NO_BORDER;
                            sharingRatio.Colspan = 6;
                            tableLeftContent.AddCell(sharingRatio);

                            PdfPCell feeExample = new PdfPCell(new Phrase("Fee Example : 12.89898\r\n", fontBodyBold));
                            feeExample.Border = Rectangle.NO_BORDER;
                            feeExample.Colspan = 6;
                            tableLeftContent.AddCell(feeExample);

                            PdfPCell methodOfCalculating = new PdfPCell(new Phrase("Method of calculating : 12.89898\r\n", fontBodyBold));
                            methodOfCalculating.Border = Rectangle.NO_BORDER;
                            methodOfCalculating.Colspan = 6;
                            tableLeftContent.AddCell(methodOfCalculating);

                            document.Add(tableLeftContent);

                            //Add Graph

                            var image = Image.GetInstance(GenerateGraphImage());
                            image.ScalePercent(40f);
                            image.SetAbsolutePosition(300, 590);
                            document.Add(image);

                            PdfPTable portfolioTable = new PdfPTable(6);
                            portfolioTable.WidthPercentage = 50;

                            //Table Heading

                            WriteToFile("Model Portfolio Comparison - 3" + DateTime.Now);
                            PdfPCell mainHeaderCell = new PdfPCell(new Phrase("Model Portfolio Comparison", TblHeaderFont)) { BackgroundColor = tableMainHeaderColor };
                            mainHeaderCell.Colspan = 6;
                            mainHeaderCell.PaddingBottom = 12;
                            mainHeaderCell.PaddingTop = 10;
                            mainHeaderCell.Border = Rectangle.NO_BORDER;
                            mainHeaderCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable.AddCell(mainHeaderCell);

                            //Table Header

                            PdfPCell cellBlankRow = new PdfPCell(new Phrase("")) { BackgroundColor = tableHeaderColor };
                            cellBlankRow.Colspan = 3;
                            cellBlankRow.PaddingBottom = 12;
                            cellBlankRow.PaddingTop = 10;
                            cellBlankRow.Border = Rectangle.NO_BORDER;
                            cellBlankRow.HorizontalAlignment = Element.ALIGN_LEFT;
                            portfolioTable.AddCell(cellBlankRow);

                            PdfPCell portfolioCell = new PdfPCell(new Phrase("Portfolio", THF)) { BackgroundColor = tableHeaderColor };
                            portfolioCell.Colspan = 1;
                            portfolioCell.PaddingBottom = 12;
                            portfolioCell.PaddingTop = 10;
                            portfolioCell.Border = Rectangle.NO_BORDER;
                            portfolioCell.HorizontalAlignment = Element.ALIGN_LEFT;
                            portfolioTable.AddCell(portfolioCell);

                            PdfPCell sP500Cell = new PdfPCell(new Phrase("S&P500 (TR)", THF)) { BackgroundColor = tableHeaderColor };
                            sP500Cell.Colspan = 2;
                            sP500Cell.PaddingBottom = 12;
                            sP500Cell.PaddingTop = 10;
                            sP500Cell.Border = Rectangle.NO_BORDER;
                            sP500Cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable.AddCell(sP500Cell);

                            //Generate Dynamic Data for Table

                            List<PortfolioComparisonModel> ModelPortfolioComparisonData = new List<PortfolioComparisonModel>();

                            for (int i = 0; i < 6; i++)
                            {
                                PortfolioComparisonModel PortfolioComparisonData = new PortfolioComparisonModel();
                                PortfolioComparisonData.Id = i + 1;
                                PortfolioComparisonData.Name = CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i + 1);
                                PortfolioComparisonData.PortfolioRate = Convert.ToDouble(GenerateDyanicVal("double", 100));
                                PortfolioComparisonData.SPRate = Convert.ToDouble(GenerateDyanicVal("double", 100));
                                ModelPortfolioComparisonData.Add(PortfolioComparisonData);
                            }


                            portfolioTable.PaddingTop = 0;
                            portfolioTable.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                            document.Add(portfolioTable);


                            //Table Dynamic Data

                            for (var i = 0; i < ModelPortfolioComparisonData.Count; i++)
                            {
                                PdfPTable PortfolioComparisonTable = new PdfPTable(6);
                                PortfolioComparisonTable.WidthPercentage = 50;
                                PortfolioComparisonTable.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                                PortfolioComparisonTable.PaddingTop = 0;

                                BaseColor cellBackgroundColor = tableHeaderColor;
                                if (i % 2 == 0)
                                {
                                    cellBackgroundColor = tableHeaderColor2;
                                }

                                PdfPCell cellBlankRow11 = new PdfPCell(new Phrase(ModelPortfolioComparisonData[i].Name, TDF)) { BackgroundColor = cellBackgroundColor };
                                cellBlankRow11.Colspan = 3;
                                cellBlankRow11.PaddingBottom = 12;
                                cellBlankRow11.PaddingTop = 10;
                                cellBlankRow11.Border = Rectangle.NO_BORDER;
                                cellBlankRow11.HorizontalAlignment = Element.ALIGN_LEFT;
                                PortfolioComparisonTable.AddCell(cellBlankRow11);

                                var portfolioRate = String.Format("{0:0.00}", ModelPortfolioComparisonData[i].PortfolioRate) + "%";
                                PdfPCell portfolioCell11 = new PdfPCell(new Phrase(portfolioRate, TDF)) { BackgroundColor = cellBackgroundColor };
                                portfolioCell11.Colspan = 1;
                                portfolioCell11.PaddingBottom = 12;
                                portfolioCell11.PaddingTop = 10;
                                portfolioCell11.Border = Rectangle.NO_BORDER;
                                portfolioCell11.HorizontalAlignment = Element.ALIGN_LEFT;
                                PortfolioComparisonTable.AddCell(portfolioCell11);

                                var SPRate = String.Format("{0:0.00}", ModelPortfolioComparisonData[i].SPRate) + "%";
                                PdfPCell sP500Cell11 = new PdfPCell(new Phrase(SPRate, TDF)) { BackgroundColor = cellBackgroundColor };
                                sP500Cell11.Colspan = 2;
                                sP500Cell11.PaddingBottom = 12;
                                sP500Cell11.PaddingTop = 10;
                                sP500Cell11.Border = Rectangle.NO_BORDER;
                                sP500Cell11.HorizontalAlignment = Element.ALIGN_CENTER;
                                PortfolioComparisonTable.AddCell(sP500Cell11);

                                document.Add(PortfolioComparisonTable);
                            }

                            //dynamic table end

                            document.AddTitle("Title");

                            PdfPTable monthlyPerformanceTable = new PdfPTable(12);
                            monthlyPerformanceTable.WidthPercentage = 80;
                            monthlyPerformanceTable.PaddingTop = 30;


                            //Table Heading

                            //PdfPCell mpMainHeaderCell = new PdfPCell(new Phrase("Monthly Performance", TblHeaderFont)) { BackgroundColor = tableMainHeaderColor };
                            //mpMainHeaderCell.Colspan = 12;
                            //mpMainHeaderCell.PaddingTop = 10;
                            //mpMainHeaderCell.PaddingBottom = 12;
                            //mpMainHeaderCell.Border = Rectangle.NO_BORDER;
                            //mpMainHeaderCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            //monthlyPerformanceTable.AddCell(mpMainHeaderCell);

                            //PdfPCell mpMainHeaderCell1 = new PdfPCell(new Phrase("Monthly Performance", TblHeaderFont)) { BackgroundColor = tableMainHeaderColor };
                            ////mpMainHeaderCell1.Colspan = 2;
                            //mpMainHeaderCell1.PaddingTop = 10;
                            //mpMainHeaderCell1.PaddingBottom = 12;
                            ////mpMainHeaderCell1.Border = Rectangle.NO_BORDER;
                            //mpMainHeaderCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            //monthlyPerformanceTable.AddCell(mpMainHeaderCell1);

                            //document.Add(monthlyPerformanceTable);

                            //End
                            WriteToFile("Monthly Performance -4" + DateTime.Now);
                            PdfPTable portfolioTable1 = new PdfPTable(15);
                            portfolioTable1.WidthPercentage = 100;
                            portfolioTable1.PaddingTop = 20;

                            //Table Heading

                            PdfPCell mainHeaderCell1 = new PdfPCell(new Phrase("Monthly Performance", TblHeaderFont)) { BackgroundColor = tableMainHeaderColor };
                            mainHeaderCell1.Colspan = 15;
                            mainHeaderCell1.PaddingBottom = 12;
                            mainHeaderCell1.PaddingTop = 10;
                            mainHeaderCell1.Border = Rectangle.NO_BORDER;
                            mainHeaderCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(mainHeaderCell1);

                            //Table Header

                            PdfPCell yearsRow1 = new PdfPCell(new Phrase("Year", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            yearsRow1.Colspan = 1;
                            yearsRow1.PaddingBottom = 12;
                            yearsRow1.PaddingTop = 10;
                            //yearsRow1.Border = Rectangle.NO_BORDER;
                            yearsRow1.BorderWidthRight = 1;
                            yearsRow1.BorderWidthBottom = 1;
                            yearsRow1.BorderWidthTop = 0;
                            yearsRow1.BorderWidthLeft = 0;
                            yearsRow1.BorderColorRight = tableHeaderColor;
                            yearsRow1.BorderColorBottom = tableHeaderColor;
                            yearsRow1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(yearsRow1);

                            PdfPCell portfolioCell1 = new PdfPCell(new Phrase("", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            portfolioCell1.Colspan = 1;
                            portfolioCell1.PaddingBottom = 12;
                            portfolioCell1.PaddingTop = 10;
                            portfolioCell1.Border = Rectangle.NO_BORDER;
                            portfolioCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(portfolioCell1);

                            PdfPCell monthJan = new PdfPCell(new Phrase("Jan", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthJan.Colspan = 1;
                            monthJan.PaddingBottom = 12;
                            monthJan.PaddingTop = 10;
                            monthJan.Border = Rectangle.NO_BORDER;
                            monthJan.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan);

                            PdfPCell monthFeb = new PdfPCell(new Phrase("Feb", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthFeb.Colspan = 1;
                            monthFeb.PaddingBottom = 12;
                            monthFeb.PaddingTop = 10;
                            monthFeb.Border = Rectangle.NO_BORDER;
                            monthFeb.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb);

                            PdfPCell monthMar = new PdfPCell(new Phrase("Mar", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthMar.Colspan = 1;
                            monthMar.PaddingBottom = 12;
                            monthMar.PaddingTop = 10;
                            monthMar.Border = Rectangle.NO_BORDER;
                            monthMar.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMar);

                            PdfPCell monthApril = new PdfPCell(new Phrase("April", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthApril.Colspan = 1;
                            monthApril.PaddingBottom = 12;
                            monthApril.PaddingTop = 10;
                            monthApril.Border = Rectangle.NO_BORDER;
                            monthApril.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril);

                            PdfPCell monthMay = new PdfPCell(new Phrase("May", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthMay.Colspan = 1;
                            monthMay.PaddingBottom = 12;
                            monthMay.PaddingTop = 10;
                            monthMay.Border = Rectangle.NO_BORDER;
                            monthMay.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay);

                            PdfPCell monthJune = new PdfPCell(new Phrase("June", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthJune.Colspan = 1;
                            monthJune.PaddingBottom = 12;
                            monthJune.PaddingTop = 10;
                            monthJune.Border = Rectangle.NO_BORDER;
                            monthJune.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune);

                            PdfPCell monthJuly = new PdfPCell(new Phrase("July", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthJuly.Colspan = 1;
                            monthJuly.PaddingBottom = 12;
                            monthJuly.PaddingTop = 10;
                            monthJuly.Border = Rectangle.NO_BORDER;
                            monthJuly.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly);

                            PdfPCell monthAug = new PdfPCell(new Phrase("Aug", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthAug.Colspan = 1;
                            monthAug.PaddingBottom = 12;
                            monthAug.PaddingTop = 10;
                            monthAug.Border = Rectangle.NO_BORDER;
                            monthAug.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug);

                            PdfPCell monthSep = new PdfPCell(new Phrase("Sep", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthSep.Colspan = 1;
                            monthSep.PaddingBottom = 12;
                            monthSep.PaddingTop = 10;
                            monthSep.Border = Rectangle.NO_BORDER;
                            monthSep.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep);

                            PdfPCell monthOct = new PdfPCell(new Phrase("Oct", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthOct.Colspan = 1;
                            monthOct.PaddingBottom = 12;
                            monthOct.PaddingTop = 10;
                            monthOct.Border = Rectangle.NO_BORDER;
                            monthOct.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct);

                            PdfPCell monthNov = new PdfPCell(new Phrase("Nov", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthNov.Colspan = 1;
                            monthNov.PaddingBottom = 12;
                            monthNov.PaddingTop = 10;
                            monthNov.Border = Rectangle.NO_BORDER;
                            monthNov.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov);

                            PdfPCell monthDec = new PdfPCell(new Phrase("Dec", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            monthDec.Colspan = 1;
                            monthDec.PaddingBottom = 12;
                            monthDec.PaddingTop = 10;
                            monthDec.Border = Rectangle.NO_BORDER;
                            monthDec.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec);

                            PdfPCell ytdRow = new PdfPCell(new Phrase("YTD", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            ytdRow.Colspan = 1;
                            ytdRow.PaddingBottom = 12;
                            ytdRow.PaddingTop = 10;
                            ytdRow.Border = Rectangle.NO_BORDER;
                            ytdRow.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdRow);

                            PdfPCell year1Cell = new PdfPCell(new Phrase("2020", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            year1Cell.Colspan = 1;
                            year1Cell.Rowspan = 2;
                            year1Cell.PaddingBottom = 15;
                            year1Cell.BorderWidthRight = 0;
                            year1Cell.BorderWidthBottom = 1;
                            year1Cell.BorderWidthTop = 0;
                            year1Cell.BorderWidthLeft = 0;
                            year1Cell.BorderColorBottom = tableHeaderColor;

                            year1Cell.PaddingTop = 15;
                            //year1Cell.Border = Rectangle.NO_BORDER;
                            year1Cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(year1Cell);

                            PdfPCell performanceTypeCell = new PdfPCell(new Phrase("Portfolio", THF)) { BackgroundColor = tableHeaderColor2 };
                            performanceTypeCell.Colspan = 1;
                            performanceTypeCell.PaddingBottom = 7;
                            performanceTypeCell.PaddingTop = 6;
                            performanceTypeCell.Border = Rectangle.NO_BORDER;
                            performanceTypeCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(performanceTypeCell);

                            PdfPCell monthJan1 = new PdfPCell(new Phrase("", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthJan1.Colspan = 1;
                            monthJan1.PaddingBottom = 7;
                            monthJan1.PaddingTop = 6;
                            monthJan1.Border = Rectangle.NO_BORDER;
                            monthJan1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan1);

                            PdfPCell monthFeb1 = new PdfPCell(new Phrase("0.0%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthFeb1.Colspan = 1;
                            monthFeb1.PaddingBottom = 7;
                            monthFeb1.PaddingTop = 6;
                            monthFeb1.Border = Rectangle.NO_BORDER;
                            monthFeb1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb1);

                            PdfPCell monthMarch1 = new PdfPCell(new Phrase("0.7%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthMarch1.Colspan = 1;
                            monthMarch1.PaddingBottom = 7;
                            monthMarch1.PaddingTop = 6;
                            monthMarch1.Border = Rectangle.NO_BORDER;
                            monthMarch1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMarch1);

                            PdfPCell monthApril1 = new PdfPCell(new Phrase("0.7%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthApril1.Colspan = 1;
                            monthApril1.PaddingBottom = 7;
                            monthApril1.PaddingTop = 6;
                            monthApril1.Border = Rectangle.NO_BORDER;
                            monthApril1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril1);

                            PdfPCell monthMay1 = new PdfPCell(new Phrase("4.8%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthMay1.Colspan = 1;
                            monthMay1.PaddingBottom = 7;
                            monthMay1.PaddingTop = 6;
                            monthMay1.Border = Rectangle.NO_BORDER;
                            monthMay1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay1);

                            PdfPCell monthJune1 = new PdfPCell(new Phrase("7.8%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthJune1.Colspan = 1;
                            monthJune1.PaddingBottom = 7;
                            monthJune1.PaddingTop = 6;
                            monthJune1.Border = Rectangle.NO_BORDER;
                            monthJune1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune1);

                            PdfPCell monthJuly1 = new PdfPCell(new Phrase("8.9%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthJuly1.Colspan = 1;
                            monthJuly1.PaddingBottom = 7;
                            monthJuly1.PaddingTop = 6;
                            monthJuly1.Border = Rectangle.NO_BORDER;
                            monthJuly1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly1);

                            PdfPCell monthAug1 = new PdfPCell(new Phrase("1.0%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthAug1.Colspan = 1;
                            monthAug1.PaddingBottom = 7;
                            monthAug1.PaddingTop = 6;
                            monthAug1.Border = Rectangle.NO_BORDER;
                            monthAug1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug1);

                            PdfPCell monthSep1 = new PdfPCell(new Phrase("-1.4%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthSep1.Colspan = 1;
                            monthSep1.PaddingBottom = 7;
                            monthSep1.PaddingTop = 6;
                            monthSep1.Border = Rectangle.NO_BORDER;
                            monthSep1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep1);

                            PdfPCell monthOct1 = new PdfPCell(new Phrase("3.7%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthOct1.Colspan = 1;
                            monthOct1.PaddingBottom = 7;
                            monthOct1.PaddingTop = 6;
                            monthOct1.Border = Rectangle.NO_BORDER;
                            monthOct1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct1);

                            PdfPCell monthNov1 = new PdfPCell(new Phrase("6.1%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthNov1.Colspan = 1;
                            monthNov1.PaddingBottom = 7;
                            monthNov1.PaddingTop = 6;
                            monthNov1.Border = Rectangle.NO_BORDER;
                            monthNov1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov1);

                            PdfPCell monthDec1 = new PdfPCell(new Phrase("8.3%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthDec1.Colspan = 1;
                            monthDec1.PaddingBottom = 7;
                            monthDec1.PaddingTop = 6;
                            monthDec1.Border = Rectangle.NO_BORDER;
                            monthDec1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec1);

                            PdfPCell ytdCell1 = new PdfPCell(new Phrase("52.8%", THF)) { BackgroundColor = tableHeaderColor2 };
                            ytdCell1.Colspan = 1;
                            ytdCell1.PaddingBottom = 7;
                            ytdCell1.PaddingTop = 6;
                            ytdCell1.Border = Rectangle.NO_BORDER;
                            ytdCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdCell1);

                            PdfPCell performanceTypeCell1 = new PdfPCell(new Phrase("S&P 500", THF)) { BackgroundColor = tableHeaderColor };
                            performanceTypeCell1.Colspan = 1;
                            performanceTypeCell1.Border = Rectangle.NO_BORDER;
                            performanceTypeCell1.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(performanceTypeCell1);

                            PdfPCell monthJan2 = new PdfPCell(new Phrase("", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJan2.Colspan = 1;
                            monthJan2.Border = Rectangle.NO_BORDER;
                            monthJan2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan2);

                            PdfPCell monthFeb2 = new PdfPCell(new Phrase("-8.2%", fontBodyRed)) { BackgroundColor = tableHeaderColor };
                            monthFeb2.Colspan = 1;
                            monthFeb2.Border = Rectangle.NO_BORDER;
                            monthFeb2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb2);

                            PdfPCell monthMarch2 = new PdfPCell(new Phrase("-12.4%", fontBodyRed)) { BackgroundColor = tableHeaderColor };
                            monthMarch2.Colspan = 1;
                            monthMarch2.Border = Rectangle.NO_BORDER;
                            monthMarch2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMarch2);

                            PdfPCell monthApril2 = new PdfPCell(new Phrase("12.8%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthApril2.Colspan = 1;
                            monthApril2.Border = Rectangle.NO_BORDER;
                            monthApril2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril2);

                            PdfPCell monthMay2 = new PdfPCell(new Phrase("4.8%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthMay2.Colspan = 1;
                            monthMay2.Border = Rectangle.NO_BORDER;
                            monthMay2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay2);

                            PdfPCell monthJune2 = new PdfPCell(new Phrase("2.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJune2.Colspan = 1;
                            monthJune2.Border = Rectangle.NO_BORDER;
                            monthJune2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune2);

                            PdfPCell monthJuly2 = new PdfPCell(new Phrase("5.9%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJuly2.Colspan = 1;
                            monthJuly2.Border = Rectangle.NO_BORDER;
                            monthJuly2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly2);

                            PdfPCell monthAug2 = new PdfPCell(new Phrase("7.2%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthAug2.Colspan = 1;
                            monthAug2.Border = Rectangle.NO_BORDER;
                            monthAug2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug2);

                            PdfPCell monthSep2 = new PdfPCell(new Phrase("-3.4%", fontBodyRed)) { BackgroundColor = tableHeaderColor };
                            monthSep2.Colspan = 1;
                            monthSep2.Border = Rectangle.NO_BORDER;
                            monthSep2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep2);

                            PdfPCell monthOct2 = new PdfPCell(new Phrase("-2.7%", fontBodyRed)) { BackgroundColor = tableHeaderColor };
                            monthOct2.Colspan = 1;
                            monthOct2.Border = Rectangle.NO_BORDER;
                            monthOct2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct2);

                            PdfPCell monthNov2 = new PdfPCell(new Phrase("10.9%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthNov2.Colspan = 1;
                            monthNov2.Border = Rectangle.NO_BORDER;
                            monthNov2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov2);

                            PdfPCell monthDec2 = new PdfPCell(new Phrase("8.3%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthDec2.Colspan = 1;
                            monthDec2.Border = Rectangle.NO_BORDER;
                            monthDec2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec2);

                            PdfPCell ytdCell2 = new PdfPCell(new Phrase("18.45%", THF)) { BackgroundColor = tableHeaderColor };
                            ytdCell2.Colspan = 1;
                            ytdCell2.Border = Rectangle.NO_BORDER;
                            ytdCell2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdCell2);

                            // 2021 Data

                            PdfPCell year2Cell = new PdfPCell(new Phrase("2021", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            year2Cell.Colspan = 1;
                            year2Cell.Rowspan = 2;
                            year2Cell.PaddingBottom = 15;
                            year2Cell.PaddingTop = 15;
                            year2Cell.BorderWidthRight = 0;
                            year2Cell.BorderWidthBottom = 1;
                            year2Cell.BorderWidthTop = 0;
                            year2Cell.BorderWidthLeft = 0;
                            year2Cell.BorderColorBottom = tableHeaderColor;

                            //year2Cell.Border = Rectangle.NO_BORDER;
                            year2Cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(year2Cell);

                            PdfPCell performanceTypeCell2 = new PdfPCell(new Phrase("Portfolio", THF)) { BackgroundColor = tableHeaderColor2 };
                            performanceTypeCell2.Colspan = 1;
                            performanceTypeCell2.PaddingBottom = 7;
                            performanceTypeCell2.PaddingTop = 6;
                            performanceTypeCell2.Border = Rectangle.NO_BORDER;
                            performanceTypeCell2.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(performanceTypeCell2);

                            PdfPCell monthJan3 = new PdfPCell(new Phrase("-0.3%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthJan3.Colspan = 1;
                            monthJan3.PaddingBottom = 7;
                            monthJan3.PaddingTop = 6;
                            monthJan3.Border = Rectangle.NO_BORDER;
                            monthJan3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan3);

                            PdfPCell monthFeb3 = new PdfPCell(new Phrase("-2.8%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthFeb3.Colspan = 1;
                            monthFeb3.PaddingBottom = 7;
                            monthFeb3.PaddingTop = 6;
                            monthFeb3.Border = Rectangle.NO_BORDER;
                            monthFeb3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb3);

                            PdfPCell monthMarch3 = new PdfPCell(new Phrase("-5.7%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthMarch3.Colspan = 1;
                            monthMarch3.PaddingBottom = 7;
                            monthMarch3.PaddingTop = 6;
                            monthMarch3.Border = Rectangle.NO_BORDER;
                            monthMarch3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMarch3);

                            PdfPCell monthApril3 = new PdfPCell(new Phrase("5.9%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthApril3.Colspan = 1;
                            monthApril3.PaddingBottom = 7;
                            monthApril3.PaddingTop = 6;
                            monthApril3.Border = Rectangle.NO_BORDER;
                            monthApril3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril3);

                            PdfPCell monthMay3 = new PdfPCell(new Phrase("4.1%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthMay3.Colspan = 1;
                            monthMay3.PaddingBottom = 7;
                            monthMay3.PaddingTop = 6;
                            monthMay3.Border = Rectangle.NO_BORDER;
                            monthMay3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay3);

                            PdfPCell monthJune3 = new PdfPCell(new Phrase("-0.8%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthJune3.Colspan = 1;
                            monthJune3.PaddingBottom = 7;
                            monthJune3.PaddingTop = 6;
                            monthJune3.Border = Rectangle.NO_BORDER;
                            monthJune3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune3);

                            PdfPCell monthJuly3 = new PdfPCell(new Phrase("-8.9%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthJuly3.Colspan = 1;
                            monthJuly3.PaddingBottom = 7;
                            monthJuly3.PaddingTop = 6;
                            monthJuly3.Border = Rectangle.NO_BORDER;
                            monthJuly3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly3);

                            PdfPCell monthAug3 = new PdfPCell(new Phrase("-4.3%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthAug3.Colspan = 1;
                            monthAug3.PaddingBottom = 7;
                            monthAug3.PaddingTop = 6;
                            monthAug3.Border = Rectangle.NO_BORDER;
                            monthAug3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug3);

                            PdfPCell monthSep3 = new PdfPCell(new Phrase("2.4%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthSep3.Colspan = 1;
                            monthSep3.PaddingBottom = 7;
                            monthSep3.PaddingTop = 6;
                            monthSep3.Border = Rectangle.NO_BORDER;
                            monthSep3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep3);

                            PdfPCell monthOct3 = new PdfPCell(new Phrase("3.7%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthOct3.Colspan = 1;
                            monthOct3.PaddingBottom = 7;
                            monthOct3.PaddingTop = 6;
                            monthOct3.Border = Rectangle.NO_BORDER;
                            monthOct3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct3);

                            PdfPCell monthNov3 = new PdfPCell(new Phrase("6.1%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthNov3.Colspan = 1;
                            monthNov3.PaddingBottom = 7;
                            monthNov3.PaddingTop = 6;
                            monthNov3.Border = Rectangle.NO_BORDER;
                            monthNov3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov3);

                            PdfPCell monthDec3 = new PdfPCell(new Phrase("8.3%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthDec3.Colspan = 1;
                            monthDec3.PaddingBottom = 7;
                            monthDec3.PaddingTop = 6;
                            monthDec3.Border = Rectangle.NO_BORDER;
                            monthDec3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec3);

                            PdfPCell ytdCell3 = new PdfPCell(new Phrase("-6.45%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            ytdCell3.Colspan = 1;
                            ytdCell3.PaddingBottom = 7;
                            ytdCell3.PaddingTop = 6;
                            ytdCell3.Border = Rectangle.NO_BORDER;
                            ytdCell3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdCell3);

                            PdfPCell performanceTypeCell4 = new PdfPCell(new Phrase("S&P 500", THF)) { BackgroundColor = tableHeaderColor };
                            performanceTypeCell4.Colspan = 1;
                            performanceTypeCell4.Border = Rectangle.NO_BORDER;
                            performanceTypeCell4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(performanceTypeCell4);

                            PdfPCell monthJan4 = new PdfPCell(new Phrase("-1.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJan4.Colspan = 1;
                            monthJan4.Border = Rectangle.NO_BORDER;
                            monthJan4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan4);

                            PdfPCell monthFeb4 = new PdfPCell(new Phrase("0.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthFeb4.Colspan = 1;
                            monthFeb4.Border = Rectangle.NO_BORDER;
                            monthFeb4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb4);

                            PdfPCell monthMarch4 = new PdfPCell(new Phrase("0.7%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthMarch4.Colspan = 1;
                            monthMarch4.Border = Rectangle.NO_BORDER;
                            monthMarch4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMarch4);

                            PdfPCell monthApril4 = new PdfPCell(new Phrase("0.7%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthApril4.Colspan = 1;
                            monthApril4.Border = Rectangle.NO_BORDER;
                            monthApril4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril4);

                            PdfPCell monthMay4 = new PdfPCell(new Phrase("4.8%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthMay4.Colspan = 1;
                            monthMay4.Border = Rectangle.NO_BORDER;
                            monthMay4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay4);

                            PdfPCell monthJune4 = new PdfPCell(new Phrase("7.8%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJune4.Colspan = 1;
                            monthJune4.Border = Rectangle.NO_BORDER;
                            monthJune4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune4);

                            PdfPCell monthJuly4 = new PdfPCell(new Phrase("8.9%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJuly4.Colspan = 1;
                            monthJuly4.Border = Rectangle.NO_BORDER;
                            monthJuly4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly4);

                            PdfPCell monthAug4 = new PdfPCell(new Phrase("1.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthAug4.Colspan = 1;
                            monthAug4.Border = Rectangle.NO_BORDER;
                            monthAug4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug4);

                            PdfPCell monthSep4 = new PdfPCell(new Phrase("-1.4%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthSep4.Colspan = 1;
                            monthSep4.Border = Rectangle.NO_BORDER;
                            monthSep4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep4);

                            PdfPCell monthOct4 = new PdfPCell(new Phrase("3.7%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthOct4.Colspan = 1;
                            monthOct4.Border = Rectangle.NO_BORDER;
                            monthOct4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct4);

                            PdfPCell monthNov4 = new PdfPCell(new Phrase("6.1%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthNov4.Colspan = 1;
                            monthNov4.Border = Rectangle.NO_BORDER;
                            monthNov4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov4);

                            PdfPCell monthDec4 = new PdfPCell(new Phrase("8.3%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthDec4.Colspan = 1;
                            monthDec4.Border = Rectangle.NO_BORDER;
                            monthDec4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec4);

                            PdfPCell ytdCell4 = new PdfPCell(new Phrase("28.71%", THF)) { BackgroundColor = tableHeaderColor };
                            ytdCell4.Colspan = 1;
                            ytdCell4.Border = Rectangle.NO_BORDER;
                            ytdCell4.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdCell4);


                            // 2022 Data

                            PdfPCell year3Cell = new PdfPCell(new Phrase("2021", TblHeaderFont)) { BackgroundColor = performaceTblColor };
                            year3Cell.Colspan = 1;
                            year3Cell.Rowspan = 2;
                            year3Cell.PaddingBottom = 15;
                            year3Cell.PaddingTop = 15;

                            year3Cell.BorderWidthRight = 0;
                            year3Cell.BorderWidthBottom = 1;
                            year3Cell.BorderWidthTop = 0;
                            year3Cell.BorderWidthLeft = 0;
                            year3Cell.BorderColorBottom = tableHeaderColor;

                            //year2Cell.Border = Rectangle.NO_BORDER;
                            year3Cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(year3Cell);

                            PdfPCell performanceTypeCell3 = new PdfPCell(new Phrase("Portfolio", THF)) { BackgroundColor = tableHeaderColor2 };
                            performanceTypeCell3.Colspan = 1;
                            performanceTypeCell3.PaddingBottom = 7;
                            performanceTypeCell3.PaddingTop = 6;
                            performanceTypeCell3.Border = Rectangle.NO_BORDER;
                            performanceTypeCell3.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(performanceTypeCell3);

                            PdfPCell monthJan5 = new PdfPCell(new Phrase("-0.3%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthJan5.Colspan = 1;
                            monthJan5.PaddingBottom = 7;
                            monthJan5.PaddingTop = 6;
                            monthJan5.Border = Rectangle.NO_BORDER;
                            monthJan5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan5);

                            PdfPCell monthFeb5 = new PdfPCell(new Phrase("-2.8%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthFeb5.Colspan = 1;
                            monthFeb5.PaddingBottom = 7;
                            monthFeb5.PaddingTop = 6;
                            monthFeb5.Border = Rectangle.NO_BORDER;
                            monthFeb5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb5);

                            PdfPCell monthMarch5 = new PdfPCell(new Phrase("-5.7%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthMarch5.Colspan = 1;
                            monthMarch5.PaddingBottom = 7;
                            monthMarch5.PaddingTop = 6;
                            monthMarch5.Border = Rectangle.NO_BORDER;
                            monthMarch5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMarch5);

                            PdfPCell monthApril5 = new PdfPCell(new Phrase("5.9%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthApril5.Colspan = 1;
                            monthApril5.PaddingBottom = 7;
                            monthApril5.PaddingTop = 6;
                            monthApril5.Border = Rectangle.NO_BORDER;
                            monthApril5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril5);

                            PdfPCell monthMay5 = new PdfPCell(new Phrase("4.1%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthMay5.Colspan = 1;
                            monthMay5.PaddingBottom = 7;
                            monthMay5.PaddingTop = 6;
                            monthMay5.Border = Rectangle.NO_BORDER;
                            monthMay5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay5);

                            PdfPCell monthJune5 = new PdfPCell(new Phrase("-0.8%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthJune5.Colspan = 1;
                            monthJune5.PaddingBottom = 7;
                            monthJune5.PaddingTop = 6;
                            monthJune5.Border = Rectangle.NO_BORDER;
                            monthJune5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune5);

                            PdfPCell monthJuly5 = new PdfPCell(new Phrase("-8.9%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthJuly5.Colspan = 1;
                            monthJuly5.PaddingBottom = 7;
                            monthJuly5.PaddingTop = 6;
                            monthJuly5.Border = Rectangle.NO_BORDER;
                            monthJuly5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly5);

                            PdfPCell monthAug5 = new PdfPCell(new Phrase("-4.3%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            monthAug5.Colspan = 1;
                            monthAug5.PaddingBottom = 7;
                            monthAug5.PaddingTop = 6;
                            monthAug5.Border = Rectangle.NO_BORDER;
                            monthAug5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug5);

                            PdfPCell monthSep5 = new PdfPCell(new Phrase("2.4%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthSep5.Colspan = 1;
                            monthSep5.PaddingBottom = 7;
                            monthSep5.PaddingTop = 6;
                            monthSep5.Border = Rectangle.NO_BORDER;
                            monthSep5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep5);

                            PdfPCell monthOct5 = new PdfPCell(new Phrase("3.7%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthOct5.Colspan = 1;
                            monthOct5.PaddingBottom = 7;
                            monthOct5.PaddingTop = 6;
                            monthOct5.Border = Rectangle.NO_BORDER;
                            monthOct5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct5);

                            PdfPCell monthNov5 = new PdfPCell(new Phrase("6.1%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthNov5.Colspan = 1;
                            monthNov5.PaddingBottom = 7;
                            monthNov5.PaddingTop = 6;
                            monthNov5.Border = Rectangle.NO_BORDER;
                            monthNov5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov5);

                            PdfPCell monthDec5 = new PdfPCell(new Phrase("8.3%", fontBody)) { BackgroundColor = tableHeaderColor2 };
                            monthDec5.Colspan = 1;
                            monthDec5.PaddingBottom = 7;
                            monthDec5.PaddingTop = 6;
                            monthDec5.Border = Rectangle.NO_BORDER;
                            monthDec5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec5);

                            PdfPCell ytdCell5 = new PdfPCell(new Phrase("-6.45%", fontBodyRed)) { BackgroundColor = tableHeaderColor2 };
                            ytdCell5.Colspan = 1;
                            ytdCell5.PaddingBottom = 7;
                            ytdCell5.PaddingTop = 6;
                            ytdCell5.Border = Rectangle.NO_BORDER;
                            ytdCell5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdCell5);

                            PdfPCell performanceTypeCell5 = new PdfPCell(new Phrase("S&P 500", THF)) { BackgroundColor = tableHeaderColor };
                            performanceTypeCell5.Colspan = 1;
                            performanceTypeCell5.Border = Rectangle.NO_BORDER;
                            performanceTypeCell5.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable.AddCell(performanceTypeCell5);

                            PdfPCell monthJan6 = new PdfPCell(new Phrase("-1.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJan6.Colspan = 1;
                            monthJan6.Border = Rectangle.NO_BORDER;
                            monthJan6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJan6);

                            PdfPCell monthFeb6 = new PdfPCell(new Phrase("0.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthFeb6.Colspan = 1;
                            monthFeb6.Border = Rectangle.NO_BORDER;
                            monthFeb6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthFeb6);

                            PdfPCell monthMarch6 = new PdfPCell(new Phrase("0.7%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthMarch6.Colspan = 1;
                            monthMarch6.Border = Rectangle.NO_BORDER;
                            monthMarch6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMarch6);

                            PdfPCell monthApril6 = new PdfPCell(new Phrase("0.7%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthApril6.Colspan = 1;
                            monthApril6.Border = Rectangle.NO_BORDER;
                            monthApril6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthApril6);

                            PdfPCell monthMay6 = new PdfPCell(new Phrase("4.8%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthMay6.Colspan = 1;
                            monthMay6.Border = Rectangle.NO_BORDER;
                            monthMay6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthMay6);

                            PdfPCell monthJune6 = new PdfPCell(new Phrase("7.8%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJune6.Colspan = 1;
                            monthJune6.Border = Rectangle.NO_BORDER;
                            monthJune6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJune6);

                            PdfPCell monthJuly6 = new PdfPCell(new Phrase("8.9%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthJuly6.Colspan = 1;
                            monthJuly6.Border = Rectangle.NO_BORDER;
                            monthJuly6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthJuly6);

                            PdfPCell monthAug6 = new PdfPCell(new Phrase("1.0%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthAug6.Colspan = 1;
                            monthAug6.Border = Rectangle.NO_BORDER;
                            monthAug6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthAug6);

                            PdfPCell monthSep6 = new PdfPCell(new Phrase("-1.4%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthSep6.Colspan = 1;
                            monthSep6.Border = Rectangle.NO_BORDER;
                            monthSep6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthSep6);

                            PdfPCell monthOct6 = new PdfPCell(new Phrase("3.7%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthOct6.Colspan = 1;
                            monthOct6.Border = Rectangle.NO_BORDER;
                            monthOct6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthOct6);

                            PdfPCell monthNov6 = new PdfPCell(new Phrase("6.1%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthNov6.Colspan = 1;
                            monthNov6.Border = Rectangle.NO_BORDER;
                            monthNov6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthNov6);

                            PdfPCell monthDec6 = new PdfPCell(new Phrase("8.3%", fontBody)) { BackgroundColor = tableHeaderColor };
                            monthDec6.Colspan = 1;
                            monthDec6.Border = Rectangle.NO_BORDER;
                            monthDec6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(monthDec6);

                            PdfPCell ytdCell6 = new PdfPCell(new Phrase("28.71%", THF)) { BackgroundColor = tableHeaderColor };
                            ytdCell6.Colspan = 1;
                            ytdCell6.Border = Rectangle.NO_BORDER;
                            ytdCell6.HorizontalAlignment = Element.ALIGN_CENTER;
                            portfolioTable1.AddCell(ytdCell6);

                            //Generate Dynamic Data for Table

                            portfolioTable.PaddingTop = 0;
                            //portfolioTable.HorizontalAlignment = Rectangle.ALIGN_RIGHT;
                            document.Add(portfolioTable1);

                            pageSize.BackgroundColor = new BaseColor(234, 244, 251);
                            document.Close();
                            byte[] bytes = memoryStream.ToArray();
                            memoryStream.Close();
                            var datetime = DateTime.Now.ToString();
                            datetime = datetime.Replace(" ", "_");
                            datetime = datetime.Replace(":", "_");
                            var filename = "WCM_" + datetime + ".pdf";
                            WriteToFile("GenerateFilepath" + DateTime.Now);
                            //filename = "WCM_" + datetime.Replace(" ", "_") + ".pdf";

                            //  var target = Path.Combine(GetPhysicalRootPath(), "wwwroot", "Files", "GeneratedReports");
                            var target = Path.Combine(@"D:\Files\GeneratedReports");

                            //var strPath = Path.Combine(new Uri(target).ToString(), filename);
                            var strPath = Path.Combine(target, filename);
                            WriteToFile("Successfully Created PDF " + DateTime.Now);
                            //var filePath1 = Path.Combine(GetPhysicalRootPath(), "wwwroot", "Files", "GeneratedReports");
                            var filePath1 = Path.Combine(@"D:\Files\GeneratedReports");
                            var exists = Directory.Exists(filePath1);
                            if (!exists)
                            {
                                Directory.CreateDirectory(filePath1);
                            }
                            filePath1 = Path.Combine(filePath1.ToString(), filename);
                            File.WriteAllBytes(filePath1, bytes);
                            return strPath;
                            command.ExecuteNonQuery();
                            connection.Close();

                        }
                    }
                }
               
            }
            catch (Exception ex)
            {
                return ex.Message;
                WriteToFile("Exception call at " + DateTime.Now + ex);
            }
        }
    }


}

