using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using iText.IO.Font;
using iText.IO.Font.Constants;
using iText.Kernel.Events;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System.IO;
using elinknext.Data.Services;
using elinknext.Models;

namespace elinknext.Data.Common
{
    public class GrantSummaryReport
    {
        private List<VGrantActivityTabBudgetList> AllGrantAppActivities { get; set; }
        public List<VGrantActivityTabBudgetList> AllGrantAppActivitiesThisOne;
        //private List<VGrantActivityIndicator> AllGrantProposedIndicatorsApplication { get; set; }
        //private List<VGrantActivityIndicator> AllGrantProposedIndicators { get; set; }
        public List<VAllAttachments> WorkplanAttachments { get; set; }
        private List<VGrantActivityTabList> AllActivitiesTab;
        private List<ActivityResults> AllActivityResults { get; set; }
        private List<VGrantActivityTabBudgetList> AllBudgetResults { get; set; }
        private List<VGrantActivityDetailGet> AllActivityDetailResults { get; set; }
        private List<VGrantActivityIndicatorSum> AllSumIndicators { get; set; }
        private List<VGrantActivityIndicatorSumApp> AllSumIndicatorsApp { get; set; }
        private List<VGrantActivityIndicatorReport> AllActivityAction { get; set; }

        IGrantService grantService;
        IActivityService activityService;
        IGrantAgreementService grantAgreeService;

        public GrantSummaryReport(IGrantService _grantService, IActivityService _activityService, IGrantAgreementService _grantAgreeService)
        {
            grantService = _grantService;
            activityService = _activityService;
            grantAgreeService = _grantAgreeService;
        }

        public async Task<string> GenerateAllDetailsWorkplanReport(long GrantId, bool blnWorkPlan,decimal TotalMatchRequired,decimal TotalBudgetAmt, decimal TotalMatchAmt, decimal TotalGrantSpent,decimal TotalMatchSpent, decimal TotalOtherAmt, decimal TotalOtherSpent)
        {
            //Get Data
            var grantData = await grantService.GetGrantData(GrantId);

            string AwardedAmtTxt = "";
            if (grantData.AwardedAmt != null)
            {
                AwardedAmtTxt = ((decimal)grantData.AwardedAmt).ToString("C");
            }

            string AmendmentAwardedAmtTxt = "";
            if (grantData.AmendedAmt != null)
            {
                AmendmentAwardedAmtTxt = ((decimal)grantData.AmendedAmt).ToString("C");
            }

            string AmendmentAwardedDate = "";
            if (grantData.AmendEndDate.HasValue)
            {
                AmendmentAwardedDate = Convert.ToDateTime(grantData.AmendEndDate).ToString("MM/dd/yyyy");
            }

            string GAExecutedDate = "";
            if (grantData.GAExecutedDate.HasValue)
            {
                GAExecutedDate = Convert.ToDateTime(grantData.GAExecutedDate).ToString("MM/dd/yyyy");
            }

            string AgreementEndDate = "";
            if (grantData.AgreementEndDate.HasValue)
            {
                AgreementEndDate = Convert.ToDateTime(grantData.CurrentEndDate).ToString("MM/dd/yyyy");
            }

            string FiscalAgentVal = "";
            if (grantData.GrantFiscalAgentOrgName != null)
            {
                FiscalAgentVal = grantData.GrantFiscalAgentOrgName;
            }

            string TotalBudgetAmtString = TotalBudgetAmt.ToString("C");
            string TotalGrantSpentString = TotalGrantSpent.ToString("C");
            string TotalMatchAmtString = TotalMatchAmt.ToString("C");
            string TotalMatchSpentString = TotalMatchSpent.ToString("C");
            string TotalOtherAmtString = TotalOtherAmt.ToString("C");
            string TotalOtherSpentString = TotalOtherSpent.ToString("C");

            string GrantFundsBalanceRemainingString = "";
            if (grantData.AwardedAmt != null)
            {
                GrantFundsBalanceRemainingString = ((decimal)grantData.AwardedAmt - TotalGrantSpent).ToString("C");
            } else
            {
                GrantFundsBalanceRemainingString = (TotalGrantSpent).ToString("C");
            }

            string MatchFundsBalanceRemainingString = (TotalMatchAmt - TotalMatchSpent).ToString("C");
            string OtherFundsBalanceRemainingString = (TotalOtherAmt - TotalOtherSpent).ToString("C");

            string TotalBudgeted = (TotalBudgetAmt + TotalMatchAmt + TotalOtherAmt).ToString("C");
            string TotalSpent = (TotalGrantSpent + TotalMatchSpent + TotalOtherSpent).ToString("C");
            string TotalBalance = "";
            if (grantData.AwardedAmt != null)
            {
                TotalBalance = (((decimal)grantData.AwardedAmt - TotalGrantSpent) + (TotalMatchAmt - TotalMatchSpent) + (TotalOtherAmt - TotalOtherSpent)).ToString("C");
            }

            //var percentdivide = (double)grantData.MatchFundPercent * 0.01;
            string MatchFundPercentStr = "";
            if (grantData.MatchFundPercent != null)
            {
                MatchFundPercentStr = grantData.MatchFundPercent.ToString();
            }

            //DaytoDayContact
            var GAdaytoday = await grantAgreeService.GetDaytoDayContact(grantData.ApplicantOrgId);
            string DayDayContact;
            if (GAdaytoday.Count > 0)
            {
                DayDayContact = GAdaytoday[0].FullName;
            } else
            {
                DayDayContact = "No Day-To-Day Contact Found!";
            }
            //DayDayEmail = GAdaytoday[0].Email;

            ///////////////////////
            AllGrantAppActivities = await activityService.GetAllBudgetSummaryResults(GrantId);
            AllSumIndicatorsApp = await grantService.GetAllProposedIndicatorsApplication(GrantId);
            WorkplanAttachments = await activityService.GetWorkplanAttachments(GrantId);
            AllActivitiesTab = await activityService.GetActivityTabList(GrantId);
            //AllGrantProposedIndicators = await activityService.GetAllProposedIndicatorsGrant(GrantId, AllActivitiesTab);
            AllSumIndicators = await activityService.GetAllProposedIndicatorsGrant(GrantId, AllActivitiesTab);

            bool RatesHrs = false;
            foreach (var activityDesc in AllActivitiesTab)
            { 
                if (activityDesc.StaffTimeBilled == "Y")
                {
                    RatesHrs = true;
                    break;
                }  
            }

            /////Begin PDF Creation///////////////////////////////////////////////////////////////////////////////////////////////////////////
            String Logo = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\Data\Common\eLINK_logo_RGB.png";
            byte[] pdfBytes;
            string results;

            using (var stream = new MemoryStream())
            {
                //Initialize PDF writer
                PdfWriter writer = new PdfWriter(stream, new WriterProperties().AddUAXmpMetadata().SetPdfVersion(PdfVersion.PDF_2_0));
                //Initialize PDF document
                PdfDocument pdf = new PdfDocument(writer);
                
                Document document = new Document(pdf, PageSize.A4.Rotate());
                //Document document = new Document(pdf, PageSize.A4.Rotate(), false);

                // Creating a PdfCanvas object
                Color ColorBlack = new DeviceRgb(0, 0, 0);

                PdfPage page = pdf.AddNewPage();
                PdfCanvas canvas = new PdfCanvas(page);
                canvas.SetStrokeColor(ColorBlack);
                canvas.MoveTo(25, 382);
                canvas.LineTo(815, 382);
                canvas.SetLineWidth(2);
                canvas.ClosePathStroke();

                pdf.AddEventHandler(PdfDocumentEvent.END_PAGE, new TextFooterEventHandler(document));

                //Accessibility Stuff
                pdf.SetTagged();
                pdf.GetCatalog().SetLang(new PdfString("en-US"));
                pdf.GetCatalog().SetViewerPreferences(new PdfViewerPreferences().SetDisplayDocTitle(true));
                PdfDocumentInfo info = pdf.GetDocumentInfo();
                info.SetTitle(grantData.GrantTitle);
                PdfFont fontCalibri = PdfFontFactory.CreateFont("c:/windows/fonts/calibri.ttf", PdfEncodings.WINANSI, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                PdfFont fontCalibriBold = PdfFontFactory.CreateFont("c:/windows/fonts/calibrib.ttf", PdfEncodings.WINANSI, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                PdfFont fontCalibriItalic = PdfFontFactory.CreateFont("c:/windows/fonts/calibrii.ttf", PdfEncodings.WINANSI, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
                PdfFont fontCalibriBoldItalic = PdfFontFactory.CreateFont("c:/windows/fonts/calibriz.ttf", PdfEncodings.WINANSI, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);

                Color ColorBlue = new DeviceRgb(0, 56, 101);
                Color ColorGray = new DeviceRgb(83, 86, 90);
                Color ColorModerateLightGray = new DeviceRgb(217, 217, 217);
                Color ColorLightGray = new DeviceRgb(242, 242, 242);
                Color ColorSuperLightGray = new DeviceRgb(191, 191, 191);
                Color ColorWhite = new DeviceRgb(255, 255, 255);

                document.SetMargins(30, 35, 30, 35);
                Rectangle pageSize = pdf.GetDefaultPageSize();
                float width = pageSize.GetWidth() - document.GetLeftMargin() - document.GetRightMargin();

                // Compose Paragraph
                Image logoImage = new Image(ImageDataFactory.Create(Logo), 35, 505, 73);
                logoImage.GetAccessibilityProperties().SetAlternateDescription("Logo");
                document.Add(logoImage);

                //Center
                List<TabStop> tabStops = new List<TabStop>();
                tabStops.Add(new TabStop(width / 2, TabAlignment.CENTER));

                ///////////Title///////////////////////////////////////////////////////////////////////////
                Paragraph MainTitle = new Paragraph().SetFontSize(20).SetFont(fontCalibriBold).SetFontColor(ColorBlue).SetPaddingTop(0).SetMarginTop(0).SetTextAlignment(TextAlignment.CENTER);
                
                if (blnWorkPlan)
                {
                    MainTitle.Add(new Tab()).Add("Grant Work Plan").Add(new Tab());
                } else
                {
                    MainTitle.Add(new Tab()).Add("Grant Progress Report").Add(new Tab());
                }
                document.Add(MainTitle);

                Paragraph SubTitle = new Paragraph().SetFontSize(15).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetPaddingTop(0).SetMarginTop(0).SetPaddingBottom(5).SetTextAlignment(TextAlignment.CENTER);
                SubTitle.Add(new Tab()).Add(grantData.AllocationDescription).Add(new Tab());
                document.Add(SubTitle);

                /////////////General Table Header///////////////////////////////////////////////////////////////////////////
                document.Add(CreateTableHeader(fontCalibriBold, fontCalibri, grantData.GrantTitle, grantData.GrantCode, grantData.ApplicantOrgName, FiscalAgentVal, DayDayContact, AwardedAmtTxt, AmendmentAwardedAmtTxt, TotalMatchRequired.ToString("C"), MatchFundPercentStr, GAExecutedDate, AgreementEndDate, AmendmentAwardedDate));

                /////////////Budget Summary////////////////////////////////////////////////////////////////////////////////
                document.Add(CreateTableSummary(fontCalibriBold, fontCalibri, ColorGray, ColorLightGray, ColorWhite, TotalBudgetAmtString, TotalGrantSpentString, TotalMatchAmtString, TotalMatchSpentString, GrantFundsBalanceRemainingString, MatchFundsBalanceRemainingString, TotalOtherAmtString, TotalOtherSpentString, OtherFundsBalanceRemainingString, TotalBudgeted, TotalSpent, TotalBalance));

                Paragraph BalanceRemaining = new Paragraph().SetFontSize(11).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetPaddingTop(5).SetMarginTop(0).SetPaddingBottom(5);
                BalanceRemaining.Add("*Grant balance remaining is the difference between the Awarded Amount and the Spent Amount. Other values compare budgeted and spent amounts.");
                document.Add(BalanceRemaining);

                /////////Budget Details///////////////////////////////////////////////////////////////////////////
                if (AllGrantAppActivities.Count > 0)
                {
                    Paragraph TitleBudgetSummary = new Paragraph("Budget Details").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(TitleBudgetSummary);
                    document.Add(CreateTableBudgetSummary(fontCalibriBoldItalic, fontCalibri, AllGrantAppActivities, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray));
                }

                ///////////Final Indicators Summary///////////////////////////////////////////////////////////////////////////
                //ProgressReport only?

                if (AllSumIndicatorsApp.Count() > 0 || AllSumIndicators.Count() > 0)
                {
                    Paragraph FinalIndicatorsSummary = new Paragraph("Indicator Summary").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(FinalIndicatorsSummary);
                    Table tableIndicators = new Table(UnitValue.CreatePercentArray(2)).SetBorder(null).UseAllAvailableWidth();

                    tableIndicators.AddCell(new Cell().SetBorder(null).Add(CreateIndicatorsApplicationSummary(fontCalibriBoldItalic, fontCalibri, AllSumIndicatorsApp, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray)));
                 
                    //tableIndicators.AddCell(new Cell().SetBorder(null).Add(CreateIndicatorsFinalSummary(fontCalibriBoldItalic, fontCalibri, AllGrantProposedIndicators, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray)));

                    tableIndicators.AddCell(new Cell().SetBorder(null).Add(CreateIndicatorsFinalSummary(fontCalibriBoldItalic, fontCalibri, AllSumIndicators, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray)));
                    document.Add(tableIndicators);
                }

                ///////////Grant Activity///////////////////////////////////////////////////////////////////////////
                if (AllActivitiesTab.Count > 0)
                {
                    Paragraph GrantActivitySummary = new Paragraph("Grant Activities").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(GrantActivitySummary);
                    //document.Add(CreateActivitySummary(fontCalibriBold, fontCalibri, activityService, AllGrantAppActivitiesThisOne, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray, GrantId, activityDesc));

                    int iCntActivity = 0;
                    foreach (var activityDesc in AllActivitiesTab)
                    {
                        iCntActivity += 1;

                        Table table;
                        if (iCntActivity == 1)
                        {
                            table = new Table(UnitValue.CreatePointArray(new float[] { 550f, 200f })).UseAllAvailableWidth().SetMarginBottom(0).SetKeepTogether(true);
                        } else
                        {
                            table = new Table(UnitValue.CreatePointArray(new float[] { 550f, 200f })).UseAllAvailableWidth().SetMarginBottom(0).SetMarginTop(30).SetKeepTogether(true);
                        }

                        table.AddCell(new Cell(1,2).SetFont(fontCalibriBold).SetFontSize(15).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Name: " + activityDesc.ActivityName)));

                        Text first = new Text("Activity Category: ").SetFontSize(11).SetFont(fontCalibriBold).SetFontColor(ColorBlack);
                        Text second = new Text(activityDesc.ActivityCategory).SetFontSize(11).SetFont(fontCalibri).SetFontColor(ColorBlack);
                        Paragraph ActivityCategoryString = new Paragraph().Add(first).Add(second);
                        table.AddCell(new Cell().SetFont(fontCalibriBold).SetFontSize(11).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorModerateLightGray).Add(ActivityCategoryString));

                        Text firstResults = new Text("Staff time?: ").SetFontSize(11).SetFont(fontCalibriBold).SetFontColor(ColorBlack);
                        Text secondResults;

                        if (activityDesc.StaffTimeBilled == "Y")
                        {
                            secondResults = new Text("Yes").SetFontSize(11).SetFont(fontCalibri).SetFontColor(ColorBlack);
                        }
                        else
                        {
                            secondResults = new Text("No").SetFontSize(11).SetFont(fontCalibri).SetFontColor(ColorBlack);
                        }

                        Paragraph ActivityCategoryStringResults = new Paragraph().Add(firstResults).Add(secondResults);
                        table.AddCell(new Cell().SetFont(fontCalibriBold).SetFontSize(11).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorModerateLightGray).Add(ActivityCategoryStringResults));

                        Text ActivityDescFirst = new Text("Description: ").SetFontSize(11).SetFont(fontCalibriBold).SetFontColor(ColorBlack);

                        string ActivityDescriptionStr = "";
                        if (activityDesc.ActivityDescription != null)
                        {
                            ActivityDescriptionStr = activityDesc.ActivityDescription;
                        }

                        Text ActivityDescSecond = new Text(ActivityDescriptionStr).SetFontSize(11).SetFont(fontCalibri).SetFontColor(ColorBlack);
                        Paragraph ActivityDescString = new Paragraph().Add(ActivityDescFirst).Add(ActivityDescSecond);
                        table.AddCell(new Cell(1,2).SetFont(fontCalibriBold).SetFontSize(11).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorWhite).Add(ActivityDescString));

                        AllBudgetResults = await activityService.GetAllBudgetResults(activityDesc.Id, GrantId);
                        if (AllBudgetResults.Count > 0)
                        {
                            table.AddCell(new Cell(1,2).SetFont(fontCalibriBold).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorModerateLightGray).Add(new Paragraph("Budget Details")));
                            document.Add(table);
                            document.Add(CreateTableBudgetActivitySummary(fontCalibriBold, fontCalibri, AllBudgetResults, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray));
                        } else
                        {
                            document.Add(table);
                        }

                        AllActivityResults = await activityService.GetAllActivityResults(activityDesc.Id);
                        if (AllActivityResults.Count > 0)
                        {
                            Table tableResult = new Table(UnitValue.CreatePercentArray(1)).UseAllAvailableWidth().SetMarginBottom(0).SetKeepWithNext(true);
                            tableResult.AddCell(new Cell().SetFont(fontCalibriBold).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorModerateLightGray).Add(new Paragraph("Actual Results")));
                            document.Add(tableResult);
                            document.Add(CreateTableResultsSummary(fontCalibriBoldItalic, fontCalibri, AllActivityResults, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray));
                        }

                        AllSumIndicators = await activityService.GetProposedIndicatorSum(activityDesc.Id);
                        int iCntSum = 0;
                        if (AllSumIndicators.Count > 0)
                        {
                            iCntSum += 1;
                            Table tableFinalIndicator = new Table(UnitValue.CreatePercentArray(1)).UseAllAvailableWidth().SetMarginBottom(0).SetKeepWithNext(true);
                            tableFinalIndicator.AddCell(new Cell().SetFont(fontCalibriBold).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorModerateLightGray).Add(new Paragraph("Final Indicators")));
                            document.Add(tableFinalIndicator);
                            document.Add(CreateFinalIndicator(fontCalibriBold, fontCalibri, AllSumIndicators, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray, iCntSum, AllSumIndicators.Count()));
                        }

                        ////Activity Action - (Detail)
                        AllActivityDetailResults = await activityService.GetAllActivityDetailResults(activityDesc.Id);
                        if (AllActivityDetailResults == null)
                        {
                            AllActivityDetailResults = new List<VGrantActivityDetailGet>();
                        }

                        int iCnt = 0;
                        foreach (var vgrantDetail in AllActivityDetailResults)
                        {
                            iCnt += 1;
                            document.Add(CreateActivityActionSummary(fontCalibriBold, fontCalibri, vgrantDetail, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray));

                            AllActivityAction = await activityService.GetProposedIndicatorReport(vgrantDetail.Id);
                            if (AllActivityAction.Count > 0)
                            {
                                document.Add(CreateActivityActionIndicator(fontCalibriBold, fontCalibri, AllActivityAction, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray));
                            }            
                        }
                    }
                }

                /////////////////
                document.Close();

                pdfBytes = stream.ToArray();
                results = Convert.ToBase64String(pdfBytes);
            }

            return results;
        }

        /////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private static Table CreateTableHeader(PdfFont titleFont,PdfFont defaultFont,string GrantTitle,string GrantCode,string Grantee,string FiscalAgent,string DaytoDayContact, string GrantAward, string AmendedGrantAward, string RequiredMatchMoney, string RequiredMatchPercent,  string GrantExecDate, string GrantEndDate, string AmendedGrantEndDate)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(3)).SetBorder(null).UseAllAvailableWidth().SetPaddingTop(28);
            //colWidths = new float[] { 140f, 60f, 60f };
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 130f, 65f, 65f })).SetBorder(null).UseAllAvailableWidth().SetPaddingTop(28);

            table.SetFixedLayout();
            int FontSizeVal = 12;

            Text first = new Text("Grant Title: ").SetFont(titleFont);
            Text second = new Text(GrantTitle).SetFont(defaultFont);
            Paragraph TitleText = new Paragraph().Add(first).Add(second);

            Text first2 = new Text("Grant ID: ").SetFont(titleFont);
            Text second2 = new Text(GrantCode).SetFont(defaultFont);
            Paragraph TitleText2 = new Paragraph().Add(first2).Add(second2);

            Text first3 = new Text("Grantee: ").SetFont(titleFont);
            Text second3 = new Text(Grantee).SetFont(defaultFont);
            Paragraph TitleText3 = new Paragraph().Add(first3).Add(second3);

            Text first4 = new Text("Fiscal Agent: ").SetFont(titleFont);
            Text second4 = new Text(FiscalAgent).SetFont(defaultFont);
            Paragraph TitleText4 = new Paragraph().Add(first4).Add(second4);

            Text first5 = new Text("Grant Day-to-Day Contact: ").SetFont(titleFont);
            Text second5 = new Text(DaytoDayContact).SetFont(defaultFont);
            Paragraph TitleText5 = new Paragraph().Add(first5).Add(second5);

            Text first6 = new Text("Grant Award ($): ").SetFont(titleFont);
            Text second6 = new Text(GrantAward).SetFont(defaultFont);
            Paragraph TitleText6 = new Paragraph().Add(first6).Add(second6);

            Paragraph TitleText7 = new Paragraph();
            Paragraph TitleText8 = new Paragraph();
            Paragraph TitleText9 = new Paragraph();
            if (AmendedGrantAward == "")
            {
                Text first7 = new Text("Required Match (%): ").SetFont(titleFont);
                Text second7 = new Text(RequiredMatchPercent).SetFont(defaultFont);
                TitleText7 = new Paragraph().Add(first7).Add(second7);

                Text first8 = new Text("Required Match ($): ").SetFont(titleFont);
                Text second8 = new Text(RequiredMatchMoney).SetFont(defaultFont);
               TitleText8 = new Paragraph().Add(first8).Add(second8);
            } else
            {
                Text first7 = new Text("Amended Grant Award ($): ").SetFont(titleFont);
                Text second7 = new Text(AmendedGrantAward).SetFont(defaultFont);
                TitleText7 = new Paragraph().Add(first7).Add(second7);

                Text first8 = new Text("Required Match (%): ").SetFont(titleFont);
                Text second8 = new Text(RequiredMatchPercent).SetFont(defaultFont);
                TitleText8 = new Paragraph().Add(first8).Add(second8);

                Text first9 = new Text("Required Match ($): ").SetFont(titleFont);
                Text second9 = new Text(RequiredMatchMoney).SetFont(defaultFont);
                TitleText9 = new Paragraph().Add(first9).Add(second9);
            }

            Text first10 = new Text("Grant Execution Date: ").SetFont(titleFont);
            Text second10 = new Text(GrantExecDate).SetFont(defaultFont);
            Paragraph TitleText10 = new Paragraph().Add(first10).Add(second10);

            //Text first11 = new Text("Grant End Date: ").SetFont(titleFont);
            //Text second11 = new Text(GrantEndDate).SetFont(defaultFont);
            //Paragraph TitleText11 = new Paragraph().Add(first11).Add(second11);

            Paragraph TitleText11 = new Paragraph();
            Paragraph TitleText12 = new Paragraph();
            if (AmendedGrantAward != "")
            {
                Text first11 = new Text("Grant End Date: ").SetFont(titleFont);
                Text second11 = new Text(GrantEndDate).SetFont(defaultFont);
                TitleText11 = new Paragraph().Add(first11).Add(second11);

                Text first12 = new Text("Amended Grant End Date: ").SetFont(titleFont);
                Text second12 = new Text(AmendedGrantEndDate).SetFont(defaultFont);
                TitleText12 = new Paragraph().Add(first12).Add(second12);
            } else
            {
                Text first11 = new Text("Grant End Date: ").SetFont(titleFont);
                Text second11 = new Text(GrantEndDate).SetFont(defaultFont);
                TitleText11 = new Paragraph().Add(first11).Add(second11);
            }

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText6));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText10));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText2));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText7));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText11));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText3));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText8));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText12));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText4));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText9));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(new Paragraph("")));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(TitleText5));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(new Paragraph("")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetPadding(0).Add(new Paragraph("")));

            return table;
        }

        private static Table CreateTableSummary(PdfFont titleFont, PdfFont defaultFont, Color ColorGray, Color ColorLightGray, Color ColorWhite,string TotalBudgetAmt, string TotalGrantSpent, string TotalMatchAmt, string TotalMatchSpent,string GrantFundsBalanceRemaining, string MatchFundsBalanceRemaining, string TotalOtherAmt, string TotalOtherSpent, string OtherFundsBalanceRemainingString, string TotalBudgeted, string TotalSpent, string TotalBalance)
        {
            Table table = new Table(UnitValue.CreatePercentArray(4)).SetBorder(null).UseAllAvailableWidth().SetMarginTop(32);
            table.SetFixedLayout();
            int FontSizeVal = 11;

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorWhite).Add(new Paragraph("")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Total Budgeted")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Total Spent")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Balance Remaining*")));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Grant Funds")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(TotalBudgetAmt)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(TotalGrantSpent)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(GrantFundsBalanceRemaining)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Match Funds")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(TotalMatchAmt)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(TotalMatchSpent)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(MatchFundsBalanceRemaining)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Other Funds")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(TotalOtherAmt)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(TotalOtherSpent)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(OtherFundsBalanceRemainingString)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Total")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(TotalBudgeted)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(TotalSpent)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(TotalBalance)));

            return table;
        }

        private static Table CreateTableBudgetSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityTabBudgetList> AllGrantAppActivities, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(8)).SetBorder(null).UseAllAvailableWidth();
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 110f, 100f, 120f, 185f, 55f, 55f, 60f, 35f })).SetBorder(null).UseAllAvailableWidth();

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Name")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Category")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Source Type")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Source Description")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Budgeted")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Spent")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Balance Remaining")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Match Fund?")));

            bool blnColorChange = false;
            foreach (var activityrow in AllGrantAppActivities)
            {
                string ActivityName = activityrow.ActivityName;
                string ActivityCategory = activityrow.ActivityCategory;
                string SourceType;
                string SourceDesc = activityrow.SourceTypeDesc;
                string BudgetAmt = ((decimal)activityrow.BudgetAmt).ToString("C");

                if (activityrow.SelectedGrantId == activityrow.GrantId)
                {
                    SourceType = activityrow.Source;
                } else
                {
                    if (activityrow.LkSourceTypeId == 3)
                    {
                        SourceType = "Other Funds";
                    } else
                    {
                        SourceType = activityrow.Source;
                    }
                }

                string ExpenseAmt = "";
                if (activityrow.ExpenseAmt != null)
                {
                    ExpenseAmt = ((decimal)activityrow.ExpenseAmt).ToString("C");
                }

                decimal BudgetAmtVal = 0;
                if (activityrow.BudgetAmt != null)
                {
                    BudgetAmtVal = (decimal)activityrow.BudgetAmt;
                }

                decimal ExpenseAmtVal = 0;
                if (activityrow.ExpenseAmt != null)
                {
                    ExpenseAmtVal = (decimal)activityrow.ExpenseAmt;
                }

                decimal BalanceRemainingCalc = BudgetAmtVal - ExpenseAmtVal;
                string BalanceRemaining = BalanceRemainingCalc.ToString("C");
                string MatchingFund = activityrow.IsMatch;

                if (ActivityName == null)
                {
                    ActivityName = "";
                }
                if (ActivityCategory == null)
                {
                    ActivityCategory = "";
                }
                if (SourceType == null)
                {
                    SourceType = "";
                }
                if (SourceDesc == null)
                {
                    SourceDesc = "";
                }
                if (BudgetAmt == null)
                {
                    BudgetAmt = "";
                }
                if (ExpenseAmt == null)
                {
                    ExpenseAmt = "";
                }
                if (BalanceRemaining == null)
                {
                    BalanceRemaining = "";
                }
                if (MatchingFund == null)
                {
                    MatchingFund = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityName)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(SourceType)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(SourceDesc)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(BudgetAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ExpenseAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(BalanceRemaining)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(MatchingFund)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityName)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(SourceType)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(SourceDesc)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(BudgetAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ExpenseAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(BalanceRemaining)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(MatchingFund)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        private static Table CreateTableResultsSummary(PdfFont titleFont, PdfFont defaultFont, List<ActivityResults> AllActivityResults, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            Table table = new Table(UnitValue.CreatePercentArray(1)).SetBorder(null).UseAllAvailableWidth().SetKeepTogether(true);

            bool blnColorChange = false;
            foreach (var activitydetailrow in AllActivityResults)
            {
                string DetailString = activitydetailrow.Results;
                if (DetailString == null)
                {
                    DetailString = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(DetailString)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(DetailString)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        private static Table CreateTableBudgetActivitySummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityTabBudgetList> AllGrantAppActivities, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(7)).SetBorder(null).UseAllAvailableWidth().SetKeepTogether(true);
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 120f, 205f, 70f, 70f, 95f, 100f, 70f })).SetBorder(null).UseAllAvailableWidth().SetKeepTogether(true);

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph("Source Type")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).Add(new Paragraph("Source Description")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).Add(new Paragraph("Budgeted")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).Add(new Paragraph("Spent")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).Add(new Paragraph("Balance Remaining")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).Add(new Paragraph("Last Transaction Date")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph("Match Fund?")));

            bool blnColorChange = false;
            foreach (var activityrow in AllGrantAppActivities)
            {
                string SourceType = activityrow.Source;
                string SourceDesc = activityrow.SourceTypeDesc;
                string BudgetAmt = ((decimal)activityrow.BudgetAmt).ToString("C");

                string ExpenseAmt = "";
                if (activityrow.ExpenseAmt != null)
                {
                    ExpenseAmt = ((decimal)activityrow.ExpenseAmt).ToString("C");
                }

                decimal BudgetAmtVal = 0;
                if (activityrow.BudgetAmt != null)
                {
                    BudgetAmtVal = (decimal)activityrow.BudgetAmt;
                }

                decimal ExpenseAmtVal = 0;
                if (activityrow.ExpenseAmt != null)
                {
                    ExpenseAmtVal = (decimal)activityrow.ExpenseAmt;
                }

                decimal BalanceRemainingCalc = BudgetAmtVal - ExpenseAmtVal;
                string BalanceRemaining = BalanceRemainingCalc.ToString("C");

                string LastTransactionDate = "";
                if (activityrow.ExpenseMaxDate != null)
                {
                    LastTransactionDate = @Convert.ToDateTime(activityrow.ExpenseMaxDate).ToString("MM/dd/yyyy");
                }

                string MatchingFund = activityrow.IsMatch;

                if (SourceType == null)
                {
                    SourceType = "";
                }
                if (SourceDesc == null)
                {
                    SourceDesc = "";
                }
                if (BudgetAmt == null)
                {
                    BudgetAmt = "";
                }
                if (ExpenseAmt == null)
                {
                    ExpenseAmt = "";
                }
                if (LastTransactionDate == null)
                {
                    LastTransactionDate = "";
                }
                if (MatchingFund == null)
                {
                    MatchingFund = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(SourceType)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).Add(new Paragraph(SourceDesc)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).Add(new Paragraph(BudgetAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).Add(new Paragraph(ExpenseAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).Add(new Paragraph(BalanceRemaining)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).Add(new Paragraph(LastTransactionDate)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(MatchingFund)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(SourceType)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).Add(new Paragraph(SourceDesc)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).Add(new Paragraph(BudgetAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ExpenseAmt)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).Add(new Paragraph(BalanceRemaining)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).Add(new Paragraph(LastTransactionDate)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(MatchingFund)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        private static Table CreateIndicatorsApplicationSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityIndicatorSumApp> AllSumIndicators, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(4)).SetBorder(null).SetWidth(375).SetKeepTogether(true);
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 100f, 145f, 60f, 70f })).SetBorder(null).SetWidth(375).SetKeepTogether(true);

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Indicator Category")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Proposed Indicator")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Total Value")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Unit")));

            bool blnColorChange = false;
            foreach (var indicatorrow in AllSumIndicators)
            {
                //string ActivityCategory = indicatorrow.Category;
                string IndicatorCategory = indicatorrow.Category;
                string IndicatorValue = indicatorrow.IndicatorValue.ToString();
                string IndicatorNameUnits = indicatorrow.NameUnits;
                string IndicatorUnitDesc = indicatorrow.UnitDesc;

                //if (ActivityCategory == null)
                //{
                //    ActivityCategory = "";
                //}

                if (IndicatorCategory == null)
                {
                    IndicatorCategory = "";
                }
                if (IndicatorValue == null)
                {
                    IndicatorValue = "";
                }
                if (IndicatorNameUnits == null)
                {
                    IndicatorNameUnits = "";
                }
                if (IndicatorUnitDesc == null)
                {
                    IndicatorUnitDesc = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorNameUnits)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorValue)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorUnitDesc)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorNameUnits)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorValue)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorUnitDesc)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        //List<VGrantActivityIndicatorSum> AllSumIndicators
        //private static Table CreateIndicatorsFinalSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityIndicator> AllGrantProposedIndicators, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        private static Table CreateIndicatorsFinalSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityIndicatorSum> AllSumIndicators, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(4)).SetBorder(null).SetWidth(375).SetKeepTogether(true);
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 100f, 145f, 60f, 70f })).SetBorder(null).SetWidth(375).SetKeepTogether(true);

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Indicator Category")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Final Indicator")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Total Value")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Unit")));

            bool blnColorChange = false;
            foreach (var indicatorrow in AllSumIndicators)
            {
                //string ActivityCategory = indicatorrow.Category;
                string IndicatorCategory = indicatorrow.Category;
                string IndicatorValue = indicatorrow.IndicatorValue.ToString();
                string IndicatorNameUnits = indicatorrow.NameUnits;
                string IndicatorUnitDesc = indicatorrow.UnitDesc;

                //if (ActivityCategory == null)
                //{
                //    ActivityCategory = "";
                //}
                if (IndicatorCategory == null)
                {
                    IndicatorCategory = "";
                }
                if (IndicatorValue == null)
                {
                    IndicatorValue = "";
                }
                if (IndicatorNameUnits == null)
                {
                    IndicatorNameUnits = "";
                }
                if (IndicatorUnitDesc == null)
                {
                    IndicatorUnitDesc = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorNameUnits)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorValue)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorUnitDesc)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorNameUnits)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorValue)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorUnitDesc)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        private static Table CreateActivityActionSummary(PdfFont titleFont, PdfFont defaultFont, VGrantActivityDetailGet ActDetail, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            Table tableAction = new Table(UnitValue.CreatePointArray(new float[] { 23f, 102f, 30f })).SetBorder(null).UseAllAvailableWidth().SetMarginBottom(0).SetMarginTop(15).SetKeepTogether(true);
            tableAction.SetFixedLayout();

            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Activity Action Name:")));
            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph(ActDetail.DetailName)));

            if (ActDetail.Count == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Activity Count: 0")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderTop(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Activity Count: " + ActDetail.Count.ToString())));
            }

            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Practice Type:")));
            if (ActDetail.Practice == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).Add(new Paragraph("")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).Add(new Paragraph(ActDetail.Practice)));
            }

            if (ActDetail.ActualSize == null || ActDetail.Units == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Size/Units: ")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Size/Units: " + ActDetail.ActualSize + " - " + ActDetail.Units)));
            }

            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("TA Provider/JAA:")));
            if (ActDetail.TAProvider == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).Add(new Paragraph("")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).Add(new Paragraph(ActDetail.TAProvider)));
            }

            if (ActDetail.Lifespan == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Lifespan: ")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Lifespan: " + ActDetail.Lifespan)));
            }

            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Practice Description:")));
            if (ActDetail.PracticeDescription == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).Add(new Paragraph("")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).Add(new Paragraph(ActDetail.PracticeDescription)));
            }

            if (ActDetail.InstallationDate.HasValue)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Install Date: " + @Convert.ToDateTime(ActDetail.InstallationDate).ToString("MM/dd/yyyy"))));
            } else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Install Date: ")));
            }

            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("")));
            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("")));

            if (ActDetail.Mapped == null)
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Mapped: ")));
            }
            else
            {
                tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorBlack, 0.25f)).Add(new Paragraph("Mapped: " + ActDetail.Mapped)));
            }

            return tableAction;
        }

        private static Table CreateFinalIndicator(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityIndicatorSum> AllDetail, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray, int IndicatorCount, int IndicatorCountAll)
        {
            bool finalindicator = false;
            if (IndicatorCountAll == IndicatorCount)
            {
                finalindicator = false;
            }

            Table tableAction = new Table(UnitValue.CreatePointArray(new float[] { 55f, 30f, 150f })).SetBorder(null).UseAllAvailableWidth().SetMarginTop(0).SetKeepTogether(true);
            tableAction.SetFixedLayout();

            tableAction.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph("Indicator")));
            tableAction.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).Add(new Paragraph("Total Value")));
            tableAction.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetUnderline(1.5f, -3).SetBorder(null).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph("Unit")));

            bool blnColorChange = false;
            foreach (var indicatorrow in AllDetail)
            {
                string strIndicatorUnit = "";
                if (indicatorrow.NameUnits != null)
                {
                    strIndicatorUnit = indicatorrow.NameUnits;
                }

                string strIndicatorValue = "";
                if (indicatorrow.IndicatorValue != null)
                {
                    strIndicatorValue = indicatorrow.IndicatorValue.ToString();
                }

                string strUnitDesc = "";
                if (indicatorrow.UnitDesc != null)
                {
                    strUnitDesc = indicatorrow.UnitDesc;
                }

                if (blnColorChange)
                {
                    if (finalindicator)
                    {
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strIndicatorUnit)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strIndicatorValue)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strUnitDesc)));
                    } else
                    {
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strIndicatorUnit)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).Add(new Paragraph(strIndicatorValue)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strUnitDesc)));
                    }
                    blnColorChange = false;
                }
                else
                {
                    if (finalindicator)
                    {
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strIndicatorUnit)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strIndicatorValue)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderBottom(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(strUnitDesc)));
                    }
                    else
                    {
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderLeft(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(strIndicatorUnit)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBackgroundColor(ColorLightGray).Add(new Paragraph(strIndicatorValue)));
                        tableAction.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(null).SetBorderRight(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(strUnitDesc)));
                    }

                    blnColorChange = true;
                }
            }

            return tableAction;
        }

        private static Table CreateActivityActionIndicator(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityIndicatorReport> AllDetail, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //int MarginBottomVal = 0;
            //if (ActionCountAll == ActionCount)
            //{
            //    MarginBottomVal = 30;
            //} 

            Table tableAction = new Table(UnitValue.CreatePointArray(new float[] { 35f, 20f, 30f, 75f, 75f })).SetBorder(null).UseAllAvailableWidth().SetMarginTop(0).SetKeepTogether(true);
            //Table tableAction = new Table(UnitValue.CreatePointArray(new float[] { 55f, 30f, 75f, 75f })).SetBorder(null).UseAllAvailableWidth().SetMarginTop(0).SetMarginBottom(MarginBottomVal).SetKeepTogether(true);
            tableAction.SetFixedLayout();

            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBackgroundColor(ColorLightGray).Add(new Paragraph("Indicator Name")));
            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBackgroundColor(ColorLightGray).Add(new Paragraph("Units")));
            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBackgroundColor(ColorLightGray).Add(new Paragraph("Value")));
            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBackgroundColor(ColorLightGray).Add(new Paragraph("Calculation Tool")));
            tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).SetBackgroundColor(ColorLightGray).Add(new Paragraph("Waterbody")));

            bool blnColorChange = false;
            foreach (var indicatorrow in AllDetail)
            {
                string strIndicatorUnit = "";
                if (indicatorrow.IndicatorUnit != null)
                {
                    strIndicatorUnit = indicatorrow.NameUnits;
                    //strIndicatorUnit = indicatorrow.NameUnits + " / " + indicatorrow.IndicatorUnit;
                }

                string strIndicatorUnitDesc = "";
                if (indicatorrow.IndicatorUnit != null)
                {
                    strIndicatorUnitDesc = indicatorrow.IndicatorUnit.ToString();
                }

                string strIndicatorValue = "";
                if (indicatorrow.IndicatorValue != null)
                {
                    strIndicatorValue = indicatorrow.IndicatorValue.ToString();
                }

                string strIndicatorWaterBody = "";
                if (indicatorrow.IndicatorWaterBody != null)
                {
                    strIndicatorWaterBody = indicatorrow.IndicatorWaterBody;
                }

                string strCalcTool = "";
                if (indicatorrow.CalcTool != null)
                {
                    strCalcTool = indicatorrow.CalcTool;
                }

                if (blnColorChange)
                {
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorUnit)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorUnitDesc)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorValue)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strCalcTool)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorWaterBody)));
                    blnColorChange = false;
                }
                else
                {
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorUnit)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorUnitDesc)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorValue)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strCalcTool)));
                    tableAction.AddCell(new Cell().SetFont(defaultFont).SetFontSize(11).Add(new Paragraph(strIndicatorWaterBody)));
                    blnColorChange = true;
                }
            }

            return tableAction;
        }

        private class TextFooterEventHandler : IEventHandler
        {
            protected Document doc;

            public TextFooterEventHandler(Document doc)
            {
                this.doc = doc;
            }

            public void HandleEvent(Event currentEvent)
            {
                PdfDocumentEvent docEvent = (PdfDocumentEvent)currentEvent;

                PdfDocument pdf = docEvent.GetDocument();
                PdfPage page = docEvent.GetPage();
                int pageNumber = pdf.GetPageNumber(page);

                Rectangle pageSize = page.GetPageSize();
                PdfFont font = null;
                try
                {
                    font = PdfFontFactory.CreateFont(StandardFonts.HELVETICA_OBLIQUE);
                }
                catch (IOException e)
                {
                    Console.Error.WriteLine(e.Message);
                }

                //int numberOfPages2 = pdf.GetNumberOfPages();

                string DateToday = DateTime.Now.ToString("MM/dd/yyyy");

                //float coordX = ((pageSize.GetLeft() + doc.GetLeftMargin()) + (pageSize.GetRight() - doc.GetRightMargin())) / 2;
                float coordXdate = 158;
                float coordX = 805;
                float coordXcredit = 515;
                //float headerY = pageSize.GetTop() - doc.GetTopMargin() + 10;
                //float footerY = doc.GetBottomMargin();
                float footerY = 13;
                Canvas canvas = new Canvas(docEvent.GetPage(), pageSize);
                canvas
                    .SetFont(font)
                    .SetFontSize(9)
                    .ShowTextAligned("Report created on: " + DateToday.ToString(), coordXdate, footerY, TextAlignment.RIGHT)
                    .ShowTextAligned("Generated by iTEXT (https://itextpdf.com/).", coordXcredit, footerY, TextAlignment.RIGHT)
                    .ShowTextAligned(pageNumber.ToString(), coordX, footerY, TextAlignment.RIGHT)
                    .Close();
            }
        }

    }
}
