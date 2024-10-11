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
    public class ReturnFundReport
    {
        private VContactSearchTeamMembers SelectedPrimaryContact { get; set; }
        private IQueryable<VContactSearchTeamMembers> AllPrimaryContacts;

        IGrantService grantService;

        public ReturnFundReport(IGrantService _grantService)
        {
            grantService = _grantService;
        }

        public async Task<string> GenerateReturnFundsReport(long GrantId)
        {
            //Get Data
            var grantData = await grantService.GetGrantData(GrantId);
            AllPrimaryContacts = await grantService.IGetContactsByOrg(grantData.ApplicantOrgId);
            if (grantData.ContactPersonId != null)
                SelectedPrimaryContact = AllPrimaryContacts.Where(x => x.UserId == grantData.ContactPersonId).SingleOrDefault();

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

                Document document = new Document(pdf);

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
                Color ColorBlack = new DeviceRgb(0, 0, 0);
                Color ColorLightGray = new DeviceRgb(242, 242, 242);
                Color ColorSuperLightGray = new DeviceRgb(191, 191, 191);
                Color ColorWhite = new DeviceRgb(255, 255, 255);

                document.SetMargins(30, 35, 30, 35);
                Rectangle pageSize = pdf.GetDefaultPageSize();
                float width = pageSize.GetWidth() - document.GetLeftMargin() - document.GetRightMargin();

                // Compose Paragraph
                Image logoImage = new Image(ImageDataFactory.Create(Logo), 35, 740, 73);
                logoImage.GetAccessibilityProperties().SetAlternateDescription("Logo");
                document.Add(logoImage);

                //Center
                List<TabStop> tabStops = new List<TabStop>();
                tabStops.Add(new TabStop(width / 2, TabAlignment.CENTER));

                ///////////Title///////////////////////////////////////////////////////////////////////////
                Paragraph MainTitle = new Paragraph().SetFontSize(20).SetFont(fontCalibriBold).SetFontColor(ColorBlue).SetPaddingTop(8).SetTextAlignment(TextAlignment.CENTER);
                MainTitle.Add(new Tab()).Add("Minnesota Board of Water and Soil Resources").Add(new Tab());
                document.Add(MainTitle);

                Paragraph SubTitle = new Paragraph().SetFontSize(15).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetTextAlignment(TextAlignment.CENTER);
                SubTitle.Add(new Tab()).Add("Return of State Grant Funds").Add(new Tab());
                document.Add(SubTitle);

                ///////////Description///////////////////////////////////////////////////////////////////////////
                Text first = new Text("This form is to be used when returning unspent or unencumbered State of MN grant funds. As stated within the ").SetFontSize(11).SetFont(fontCalibri).SetFontColor(ColorBlack);
                Text second = new Text("Closing out a BWSR Grant").SetFontSize(11).SetFont(fontCalibriBold).SetFontColor(ColorBlack);
                Text third = new Text(" section of the BWSR Grant Administration Manual, any funds remaining unspent or becoming unobligated or unencumbered at the conclusion of the grant contract period, must be returned within 30 calendar days of the end of the grant agreement period.").SetFontSize(11).SetFont(fontCalibri).SetFontColor(ColorBlack);
                Paragraph Terms = new Paragraph().Add(first).Add(second).Add(third).SetMarginBottom(0).SetPaddingTop(20);
                document.Add(Terms);

                ///////////Sub Title///////////////////////////////////////////////////////////////////////////
                Paragraph GrantTitle = new Paragraph();
                GrantTitle.Add("Grant Title").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0).SetPaddingTop(15);
                GrantTitle.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                GrantTitle.Add(new Text(grantData.GrantTitle).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(GrantTitle);

                Paragraph GrantID = new Paragraph();
                GrantID.Add("Grant Code").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                GrantID.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray);
                GrantID.Add(new Text( grantData.GrantCode).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(GrantID);

                Paragraph GrantAllocation = new Paragraph();
                GrantAllocation.Add("Grant Allocation").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                GrantAllocation.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray);
                GrantAllocation.Add(new Text(grantData.AllocationDescription).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(GrantAllocation);

                Paragraph Organization = new Paragraph();
                Organization.Add("Grantee").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                Organization.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray);
                Organization.Add(new Text(grantData.ApplicantOrgName).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(Organization);

                Paragraph FiscalAgent = new Paragraph();
                FiscalAgent.Add("Fiscal Agent").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                FiscalAgent.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray);

                if (grantData.GrantFiscalAgentOrgName == null)
                {
                    FiscalAgent.Add(new Text("").SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                } else
                {
                    FiscalAgent.Add(new Text(grantData.GrantFiscalAgentOrgName).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                }
                document.Add(FiscalAgent);

                Paragraph FiscalYear = new Paragraph();
                FiscalYear.Add("Grant Fiscal Year").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetMarginBottom(0);
                FiscalYear.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray);
                FiscalYear.Add(new Text(grantData.AllocationYear.ToString()).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                document.Add(FiscalYear);

                Paragraph AgreementPO = new Paragraph();
                AgreementPO.Add("Agreement PO#").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetPaddingBottom(10);
                AgreementPO.Add(": ").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray);
                if (grantData.PONumber == null)
                {
                    AgreementPO.Add(new Text("").SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                } else
                {
                    AgreementPO.Add(new Text(grantData.PONumber.ToString()).SetFontSize(12).SetFont(fontCalibri).SetFontColor(ColorGray));
                }

                document.Add(AgreementPO);

                Paragraph EnsureProcess = new Paragraph();
                EnsureProcess.Add("To ensure accurate processing, please review the following required details and update as needed:").SetFontSize(12).SetFont(fontCalibriBold).SetFontColor(ColorGray).SetPaddingBottom(15);
                document.Add(EnsureProcess);

                ///////////General Summary///////////////////////////////////////////////////////////////////////////
                string RefundCurrency = "";
                if (grantData.ReturnedAmt != null)
                {
                    RefundCurrency = grantData.ReturnedAmt.Value.ToString("C");
                }

                string PrimaryFullName = "";
                string PhoneNumber = "";
                if (SelectedPrimaryContact != null)
                {
                    if (SelectedPrimaryContact.PhoneNumber != null)
                    {
                        PrimaryFullName = SelectedPrimaryContact.FullName;
                        PhoneNumber = SelectedPrimaryContact.PhoneNumber;
                    }
                }

                string CheckNumber = "";
                if (grantData.CheckNumber != null)
                {
                    CheckNumber = grantData.CheckNumber.ToString();
                }

                document.Add(CreateTableSummary(fontCalibriBold, fontCalibri, ColorGray, ColorBlack, ColorLightGray, ColorWhite,CheckNumber, RefundCurrency, PrimaryFullName, PhoneNumber));

                ////Mail Stuff/////////////////
                Paragraph MakeChecks = new Paragraph("This completed form and the check for unspent grant funds\nshould be mailed to BWSR. Retain a copy for your file.\nMake check payable to and mail to:").SetFontSize(11).SetFont(fontCalibri).SetPaddingTop(25).SetFontColor(ColorBlack).SetTextAlignment(TextAlignment.CENTER);
                document.Add(MakeChecks);
                Paragraph MakeChecksAdd = new Paragraph("MN Board of Water and Soil Resources\n520 Lafayette Road N\nSt. Paul, MN 55155\n651-296-3767").SetFontSize(11).SetFont(fontCalibri).SetPaddingTop(10).SetTextAlignment(TextAlignment.CENTER);
                document.Add(MakeChecksAdd);

                document.Close();

                pdfBytes = stream.ToArray();
                results = Convert.ToBase64String(pdfBytes);
            }

            return results;
        }

        private static Table CreateTableSummary(PdfFont titleFont, PdfFont defaultFont, Color ColorGray, Color ColorBlack, Color ColorLightGray, Color ColorWhite, string CheckNumber, string RefundAmt,string ContactName, string ContactPhone)
        {
            Table table = new Table(UnitValue.CreatePercentArray(2)).SetBorder(null).UseAllAvailableWidth();
            table.SetFixedLayout();
            int FontSizeVal = 11;

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Check Number")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(CheckNumber)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Refund Amount")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(RefundAmt)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Contact Name")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorLightGray).Add(new Paragraph(ContactName)));

            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(titleFont).SetBackgroundColor(ColorGray).SetFontColor(ColorWhite).Add(new Paragraph("Contact Phone Number")));
            table.AddCell(new Cell().SetFontSize(FontSizeVal).SetBorder(null).SetFont(defaultFont).SetBackgroundColor(ColorWhite).Add(new Paragraph(ContactPhone)));

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

                float coordXdate = 158;
                float coordX = 570;
                float coordXcredit = 390;
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