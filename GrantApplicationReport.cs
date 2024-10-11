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
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System.IO;
using elinknext.Data.Services;
using elinknext.Models;

namespace elinknext.Data.Common
{
    public class GrantApplicationReport
    {
        //private VContactSearchTeamMembers SelectedPrimaryContact { get; set; }
        private VGrantTeam SelectedPrimaryContact { get; set; }
        private IQueryable<VGrantTeam> AllPrimaryContacts;
        //private IQueryable<VContactSearchTeamMembers> AllPrimaryContacts;
        private List<VGrantAppActivity> AllGrantAppActivities { get; set; }
        private List<VCountyList> AllGrantCounty { get; set; }
        private List<GrantHUCList> AllGrantHUCListParams { get; set; }
        private List<GrantApplicationDetails> AllGrantApplicationQuestions { get; set; }
        private List<VGrantActivityIndicator> AllGrantProposedIndicators { get; set; }
        private List<VGrantAppActivityDetailsGet> AllGrantAppActivityQuestionsDescription { get; set; }

        private GrantAttachmentFiles GrantAttachmentsFilesImageFileParams { get; set; }
        private GrantAttachmentFiles GrantAttachmentsFilesMapImageFileParams { get; set; }
        private List<VGrantAttachmentFiles> AllGrantAttachmentFiles { get; set; }
        public object FontConstants { get; private set; }

        IGrantService grantService;

        public GrantApplicationReport(IGrantService _grantService)
        {
            grantService = _grantService;
        }

        public async Task<string> GeneratePreviewReport(long GrantId)
        {
            //Get Data
            var grantData = await grantService.GetGrantData(GrantId);

            string OtherAmountTxt = "";
            if (grantData.OtherAmount != null)
            {
                OtherAmountTxt = "$" + grantData.OtherAmount.ToString();
            }

            string SubmitDate = "";
            if (grantData.SubmittedDate.HasValue)
            {
                SubmitDate = Convert.ToDateTime(grantData.SubmittedDate).ToString("MM/dd/yyyy");
            }

            AllPrimaryContacts = await grantService.IGetGrantTeam(GrantId);
            //AllPrimaryContacts = await grantService.IGetContactsByOrg(grantData.ApplicantOrgId);
            //AllPrimaryContacts = await grantService.GetContactsByOrg(grantData.ApplicantOrgId);
            AllGrantAppActivities = await grantService.GetAllAppActivityData(GrantId);
            AllGrantProposedIndicators = await grantService.GetAllProposedIndicatorsReport(GrantId);
            AllGrantAppActivityQuestionsDescription = await grantService.GetExistingActivityQuestionsReport(GrantId);

            if (grantData.ContactPersonId != null)
                SelectedPrimaryContact = AllPrimaryContacts.Where(x => x.GrantTeamUserId == grantData.ContactPersonId).SingleOrDefault();
            var RequestAmounTemp = (decimal)AllGrantAppActivities.Where(x => x.GrantId == GrantId).Select(x => x.Amount).Sum();
            string RequestAmountVal = RequestAmounTemp.ToString("C");
            decimal MatchAmountValTmp = 0;
            if (grantData.MatchAmount != null)
            {
                MatchAmountValTmp = (decimal)grantData.MatchAmount;
            }

            string MatchAmountVal = MatchAmountValTmp.ToString("C");
            //string CalcMatchPerc = (MatchAmountValTmp / grantData.MatchAmount).ToString();

            string Countylist = "";
            AllGrantCounty = await grantService.GetGrantCountyNames(GrantId);
            if (AllGrantCounty.Count != 0)
            {
                foreach (var countyrow in AllGrantCounty)
                {
                    if (Countylist == "")
                    {
                        Countylist = countyrow.CountyName;
                    }
                    else
                    {
                        Countylist = Countylist + "," + countyrow.CountyName;
                    }
                }
            }

            string HUClist = "";
            AllGrantHUCListParams = await grantService.GetHUCs(GrantId);
            if (AllGrantHUCListParams.Count != 0)
            {
                foreach(var hucrow in AllGrantHUCListParams)
                {
                    if (HUClist == "")
                    {
                        HUClist = hucrow.HucNumber;
                    }
                    else
                    {
                        HUClist = HUClist + "," + hucrow.HucNumber;
                    }
                }
            }

            string PrimaryFullName = "";
            if (SelectedPrimaryContact != null)
            {
                if (SelectedPrimaryContact.FullName != null)
                {
                    PrimaryFullName = SelectedPrimaryContact.FullName;
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

                Color ColorBlue = new DeviceRgb(0,56,101);
                Color ColorGray = new DeviceRgb(83,86,90);
                Color ColorBlack = new DeviceRgb(0,0,0);
                Color ColorLightGray = new DeviceRgb(242,242,242);
                Color ColorSuperLightGray = new DeviceRgb(191,191,191);
                Color ColorWhite = new DeviceRgb(255,255,255);

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

                Paragraph MainTitle = new Paragraph().SetFontSize(20).AddTabStops(tabStops).SetFont(fontCalibriBold).SetFontColor(ColorBlue).SetPaddingTop(12);
                MainTitle.Add(new Tab()).Add("Grant Application").Add(new Tab());
                document.Add(MainTitle);

                ///////////Sub Title///////////////////////////////////////////////////////////////////////////
                Paragraph GrantTitle = new Paragraph();
                GrantTitle.Add("Grant Name ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0).SetPaddingTop(20);
                GrantTitle.Add(new Text("- " + grantData.GrantTitle).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(GrantTitle);

                Paragraph GrantID = new Paragraph();
                GrantID.Add("Grant ID ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                GrantID.Add(new Text("- " + grantData.GrantCode).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(GrantID);

                Paragraph Organization = new Paragraph();
                Organization.Add("Organization ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetPaddingBottom(8);
                Organization.Add(new Text("- " + grantData.ApplicantOrgName).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(Organization);

                ///////////General Summary///////////////////////////////////////////////////////////////////////////
                document.Add(CreateTableSummary(fontCalibriBold, fontCalibri, grantData.AllocationDescription, grantData.GrantFiscalAgentOrgName, MatchAmountVal, OtherAmountTxt, RequestAmountVal, grantData.MatchFundPercent.ToString(), PrimaryFullName, grantData.ProjectAbstract, grantData.MeasureableOutcome, HUClist, Countylist, SubmitDate, ColorGray, ColorBlack, ColorLightGray, ColorWhite));

                ///////////Required Checkbox///////////////////////////////////////////////////////////////////////////
                //Will always occur
                Paragraph pChecky = new Paragraph();
                PdfFont zapfdingbats = PdfFontFactory.CreateFont(StandardFonts.ZAPFDINGBATS);
                Text chunk = new Text("4").SetFont(zapfdingbats).SetFontSize(14);
                pChecky.Add(chunk);
                pChecky.Add(" **Required** MN Statute 16B.981 Subd. 2 (6) requires that no current principals of a grantee have been convicted of a felony financial crime in the last 10 years. A principal is defined as a public official, a board member, or staff (paid or volunteer) with the authority to access funds provided by this grant opportunity. By checking this box, I attest that no current principal of my organization with authority to access funds has been convicted of a felony financial crime in the last 10 years.");
                document.Add(pChecky);

                ///////////Narrative Summary///////////////////////////////////////////////////////////////////////////
                AllGrantApplicationQuestions = await grantService.GetGrantApplicationQuestions(GrantId);
                
                if (AllGrantApplicationQuestions.Count > 0)
                {
                    Paragraph NarrativeSummary = new Paragraph("Narrative").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(NarrativeSummary);

                    Table tableNarr = new Table(UnitValue.CreatePercentArray(1)).UseAllAvailableWidth();
                    tableNarr.SetFixedLayout();

                    foreach (var vgrantQues in AllGrantApplicationQuestions)
                    {
                        tableNarr.AddCell(new Cell().SetFont(fontCalibriBold).SetFontSize(11).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph(vgrantQues.ColumnDescription)));

                        if (vgrantQues.RequiredValue == null)
                        {
                            tableNarr.AddCell(new Cell().SetFont(fontCalibri).SetFontSize(11).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph("")));
                        } else
                        {
                            tableNarr.AddCell(new Cell().SetFont(fontCalibri).SetFontSize(11).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(vgrantQues.RequiredValue)));
                        }
                    }

                    document.Add(tableNarr);
                }

                ///////////Budget Summary///////////////////////////////////////////////////////////////////////////
                if (AllGrantAppActivities.Count > 0)
                {
                    Paragraph TitleBudgetSummary = new Paragraph("Application Budget").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(TitleBudgetSummary);
                    document.Add(CreateTableBudgetSummary(fontCalibriBoldItalic, fontCalibri, AllGrantAppActivities, ColorGray, ColorBlack, ColorLightGray, ColorWhite, ColorSuperLightGray));
                }

                ///////////Proposed Activity Indicators///////////////////////////////////////////////////////////////////////////
                if (AllGrantProposedIndicators.Count > 0)
                {
                    Paragraph TitleProposedActivityIndicators = new Paragraph("Proposed Activity Indicators").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(TitleProposedActivityIndicators);
                    document.Add(CreateTableProposedActivityIndicatorsSummary(fontCalibriBoldItalic, fontCalibri, AllGrantProposedIndicators, ColorGray, ColorBlack, ColorLightGray, ColorWhite,ColorSuperLightGray));
                }

                ///////////Activity Details Summary///////////////////////////////////////////////////////////////////////////
                if (AllGrantAppActivityQuestionsDescription.Count > 0)
                {
                    Paragraph TitleActivityDetailsSummary = new Paragraph("Activity Details").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(TitleActivityDetailsSummary);

                    document.Add(CreateTableActivityDetailsSummary(fontCalibriBoldItalic, fontCalibri, AllGrantAppActivityQuestionsDescription, ColorGray, ColorBlack, ColorLightGray, ColorWhite,ColorSuperLightGray));
                }

                ///////////Grant Application Attachments///////////////////////////////////////////////////////////////////////////
                AllGrantAttachmentFiles = await grantService.GetGrantAllocationAttachments(GrantId);
                if (AllGrantAttachmentFiles.Count > 0)
                {
                    Paragraph GrantAppImages = new Paragraph("Grant Application Attachments").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(GrantAppImages);

                    Table table = new Table(UnitValue.CreatePointArray(new float[] { 168f, 194f, 413f })).SetBorder(null);

                    table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(fontCalibriBoldItalic).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Document Name")));
                    table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(fontCalibriBoldItalic).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Document Type")));
                    table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(fontCalibriBoldItalic).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Description")));

                    bool blnColorChange = false;

                    foreach (var vAttach in AllGrantAttachmentFiles)
                    {
                        string DocLabel = vAttach.AttachmentLabel;
                        string DocType = vAttach.FileType;
                        string DocFileName = System.IO.Path.GetFileName(vAttach.FileLocation);
                        if (DocLabel == null)
                        {
                            DocLabel = "";
                        }
                        if (DocType == null)
                        {
                            DocType = "";
                        }
                        if (DocFileName != null)
                        {
                            if (DocFileName != "")
                            {
                                if (blnColorChange)
                                {
                                    table.AddCell(new Cell().SetFontSize(11).SetFont(fontCalibri).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(DocLabel)));
                                    table.AddCell(new Cell().SetFontSize(11).SetFont(fontCalibri).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(DocType)));
                                    table.AddCell(new Cell().SetFontSize(11).SetFont(fontCalibri).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(DocFileName)));
                                    blnColorChange = false;
                                }
                                else
                                {
                                    table.AddCell(new Cell().SetFontSize(11).SetFont(fontCalibri).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(DocLabel)));
                                    table.AddCell(new Cell().SetFontSize(11).SetFont(fontCalibri).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(DocType)));
                                    table.AddCell(new Cell().SetFontSize(11).SetFont(fontCalibri).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(DocFileName)));
                                    blnColorChange = true;
                                }
                            }
                        }

                    }
                    document.Add(table);
                }

                ///////////Application Image///////////////////////////////////////////////////////////////////////////
                GrantAttachmentsFilesImageFileParams = await grantService.GetImageAttachmentFile(GrantId, "Image");
                if (File.Exists(GrantAttachmentsFilesImageFileParams.FileLocation))
                {
                    document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

                    Paragraph AppImage = new Paragraph("Application Image").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(AppImage);

                    Image ApplicationImage = new Image(ImageDataFactory.Create(GrantAttachmentsFilesImageFileParams.FileLocation));
                    ApplicationImage.GetAccessibilityProperties().SetAlternateDescription("Application Image");

                    bool ExtraSize = false;
                    if (ApplicationImage.GetImageWidth() > 765)
                    {
                        ExtraSize = true;
                    }
                    if (ApplicationImage.GetImageHeight() > 415)
                    {
                        ExtraSize = true;
                    }

                    if (ExtraSize)
                    {
                        ApplicationImage.SetAutoScale(true);
                    }

                    document.Add(ApplicationImage);
                }

                ///////////Map Image///////////////////////////////////////////////////////////////////////////
                GrantAttachmentsFilesMapImageFileParams = await grantService.GetImageAttachmentFile(GrantId, "MapImage");
                if (File.Exists(GrantAttachmentsFilesMapImageFileParams.FileLocation))
                {
                    document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));

                    Paragraph MapImageLbl = new Paragraph("Map Image").SetKeepWithNext(true).SetFontSize(12).SetFont(fontCalibriBold).SetPaddingTop(8).SetFontColor(ColorGray);
                    document.Add(MapImageLbl);

                    Image MapImage = new Image(ImageDataFactory.Create(GrantAttachmentsFilesMapImageFileParams.FileLocation));
                    MapImage.GetAccessibilityProperties().SetAlternateDescription("Map Image");

                    bool ExtraSizeApp = false;
                    if (MapImage.GetImageWidth() > 765)
                    {
                        ExtraSizeApp = true;
                    }
                    if (MapImage.GetImageHeight() > 415)
                    {
                        ExtraSizeApp = true;
                    }

                    if (ExtraSizeApp)
                    {
                        MapImage.SetAutoScale(true);
                    }

                    document.Add(MapImage);
                }

                document.Close();

                pdfBytes = stream.ToArray();
                results = Convert.ToBase64String(pdfBytes);
            }

            return results;
        }

        private static Table CreateTableSummary(PdfFont titleFont, PdfFont defaultFont, string Allocation, string FiscalOrg, string MatchAmount, string OtherAmount,string RequestAmountVal, string MatchFundPercent, string PrimaryContact,string ProjectAbstract, string MeasureableOutcome,string HUClist, string Countylist, string SubmitDate, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite)
        {
            Table table = new Table(UnitValue.CreatePercentArray(4)).SetBorder(null).UseAllAvailableWidth();
            table.SetFixedLayout();
            int FontSizeVal = 11;

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Allocation")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(Allocation)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Grant Contact")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(PrimaryContact)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Total Grant Amount Requested")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(RequestAmountVal)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("County(s)")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(Countylist)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Grant Match Amount")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(MatchAmount)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("12 Digit HUC(s)")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(HUClist)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Required Match %")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(MatchFundPercent + "%")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Fiscal Agent")));

            if (FiscalOrg == null)
            {
                table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph("")));
            } else
            {
                table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(FiscalOrg)));
            }

            //table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Calculated Match %")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Other Amount")));
            //table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(CalcMatchPerc + "%")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(OtherAmount)));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Application Submitted Date")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(SubmitDate)));

            //table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Other Amount")));
            //table.AddCell(new Cell(1, 3).SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(OtherAmount)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Project Abstract")));

            string ProjectAbstractStr = "";
            if (ProjectAbstract != null)
            {
                ProjectAbstractStr = ProjectAbstract;
            }
            table.AddCell(new Cell(1, 3).SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ProjectAbstractStr)));

            string MeasureableOutcomeStr = "";
            if (MeasureableOutcome != null)
            {
                MeasureableOutcomeStr = MeasureableOutcome;
            }
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Proposed Measurable Outcomes")));
            table.AddCell(new Cell(1, 3).SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(MeasureableOutcomeStr)));

            return table;
        }

        private static Table CreateTableBudgetSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantAppActivity> AllGrantAppActivities, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(5)).SetBorder(null).UseAllAvailableWidth();
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 110f, 260f, 260f, 75f, 70f })).SetBorder(null);

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Name")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Description")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Category")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("State Grant $ Requested")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Lifespan (yrs)")));

            bool blnColorChange = false;
            foreach(var activityrow in AllGrantAppActivities)
            {
                string ActivityDescription = activityrow.Description;
                string ActivityActivityCategory = activityrow.ActivityCategory;
                string ActivityAmount = activityrow.Amount.Value.ToString("C");
                string ActivityLifespan = activityrow.ActivityLifespan.ToString();
                if (ActivityDescription == null)
                {
                    ActivityDescription = "";
                }
                if (ActivityActivityCategory == null)
                {
                    ActivityActivityCategory = "";
                }
                if (ActivityAmount == null)
                {
                    ActivityAmount = "";
                }
                if (ActivityLifespan == null)
                {
                    ActivityLifespan = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(activityrow.ActivityName)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityDescription)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityActivityCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityAmount)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityLifespan)));
                    blnColorChange = false;
                } else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(activityrow.ActivityName)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityDescription)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityActivityCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityAmount)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityLifespan)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        private static Table CreateTableProposedActivityIndicatorsSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantActivityIndicator> AllGrantProposedIndicators, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            //Table table = new Table(UnitValue.CreatePercentArray(6)).SetBorder(null).UseAllAvailableWidth();
            Table table = new Table(UnitValue.CreatePointArray(new float[] { 160f, 125f, 80f, 120f, 120f, 160f })).SetBorder(null);

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Name")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Indicator Name")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Value & Units")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Waterbody")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Calculation Tool")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Comments")));

            bool blnColorChange = false;
            foreach (var indicatorrow in AllGrantProposedIndicators)
            {
                string IndicatorCategory = indicatorrow.Category;
                string IndicatorNameUnits = indicatorrow.NameUnits;
                string IndicatorValue = indicatorrow.ValueUnits;
                string IndicatorWaterbody = indicatorrow.Waterbody;
                string IndicatorCalcTool = indicatorrow.CalcTool;
                string IndicatorComments = indicatorrow.Comments;

                if (IndicatorCategory == null)
                {
                    IndicatorCategory = "";
                }
                if (IndicatorNameUnits == null)
                {
                    IndicatorNameUnits = "";
                }
                if (IndicatorValue == null)
                {
                    IndicatorValue = "";
                }
                if (IndicatorWaterbody == null)
                {
                    IndicatorWaterbody = "";
                }
                if (IndicatorCalcTool == null)
                {
                    IndicatorCalcTool = "";
                }
                if (IndicatorComments == null)
                {
                    IndicatorComments = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorNameUnits)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorValue)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorWaterbody)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorCalcTool)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(IndicatorComments)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorCategory)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorNameUnits)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorValue)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorWaterbody)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorCalcTool)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(IndicatorComments)));
                    blnColorChange = true;
                }
            }

            return table;
        }

        private static Table CreateTableActivityDetailsSummary(PdfFont titleFont, PdfFont defaultFont, List<VGrantAppActivityDetailsGet> AllGrantAppActivityQuestionsDescription, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, Color ColorSuperLightGray)
        {
            Table table = new Table(UnitValue.CreatePercentArray(3)).SetBorder(null).UseAllAvailableWidth();

            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Activity Name")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Question")));
            table.AddHeaderCell(new Cell().SetFontSize(11).SetFont(titleFont).SetBorder(null).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Answer")));

            bool blnColorChange = false;
            foreach (var detailsrow in AllGrantAppActivityQuestionsDescription)
            {
                string ActivityName = detailsrow.ActivityName;
                string ColumnDescription = detailsrow.ColumnDescription;
                string RequiredValue = detailsrow.RequiredValue;

                if (ActivityName == null)
                {
                    ActivityName = "";
                }
                if (ColumnDescription == null)
                {
                    ColumnDescription = "";
                }
                if (RequiredValue == null)
                {
                    RequiredValue = "";
                }

                if (blnColorChange)
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ActivityName)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(ColumnDescription)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).Add(new Paragraph(RequiredValue)));
                    blnColorChange = false;
                }
                else
                {
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ActivityName)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ColumnDescription)));
                    table.AddCell(new Cell().SetFontSize(11).SetFont(defaultFont).SetBorder(new iText.Layout.Borders.SolidBorder(ColorSuperLightGray, 0.25f)).SetBackgroundColor(ColorLightGray).Add(new Paragraph(RequiredValue)));
                    blnColorChange = true;
                }
            }

            return table;
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
                float coordXcredit = 515;
                float coordX = 805;
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
