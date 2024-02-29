using IEMS_WEB.Areas.DepoTransfers.Models.Response;
using IEMS_WEB.Areas.OnlinePermitGenerate.Models.Request;
using IEMS_WEB.Areas.TransferInOut.Models.Response;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.qrcode;
using System.IO;

namespace IEMS_WEB.Comman
{
    public static class Reports
    {
        public static byte[] SpriteFL5(FL5ReportResponseModel model, string logoPath)
        {
            byte[] fileContents;
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Document document = new Document(new Rectangle(595, 842), 0f, 0f, 0f, 0f))
                    {
                        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                        document.Open();
                        #region Added image Logo
                        var imagePath = logoPath;
                        Image myImage = Image.GetInstance(imagePath);
                        float fixedWidth = logoPath.ToLower().Contains("excise") ? 150 : 380;
                        float fixedHeight = 150;
                        float widthScalingFactor = fixedWidth / myImage.Width;
                        float heightScalingFactor = fixedHeight / myImage.Height;
                        float scalingFactor = Math.Min(widthScalingFactor, heightScalingFactor);
                        myImage.ScaleAbsolute(myImage.Width * scalingFactor, myImage.Height * scalingFactor);
                        float xCentered = (document.PageSize.Width - myImage.ScaledWidth) / 2;
                        myImage.SetAbsolutePosition(xCentered, document.PageSize.Height - myImage.ScaledHeight - document.TopMargin);
                        document.Add(myImage);
                        #endregion
                        Font blueFont = FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLUE);
                        Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK);
                        PdfPTable dataTable = new PdfPTable(12);
                        PdfPCell BelowCell = new PdfPCell(new Phrase("FL5 Serial No : " + model.FL5NO, FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell.Border = PdfPCell.NO_BORDER;
                        BelowCell.Colspan = 6;
                        PdfPCell rightBelowCell = new PdfPCell(new Phrase("Date of Issue :" + model.FL5Date, FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLUE)));
                        rightBelowCell.Border = PdfPCell.NO_BORDER;
                        rightBelowCell.Colspan = 6;
                        rightBelowCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        PdfPCell leftCell1 = new PdfPCell(new Phrase("1 .Consigner Name", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell1.Border = PdfPCell.NO_BORDER;
                        leftCell1.Colspan = 4;
                        PdfPCell centerCell = new PdfPCell(new Phrase(":   " + model.consignerName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell.Border = PdfPCell.NO_BORDER;
                        centerCell.Colspan = 4;
                        PdfPCell centerCell1 = new PdfPCell(new Phrase("   ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell1.Border = PdfPCell.NO_BORDER;
                        centerCell1.Colspan = 4;
                        PdfPCell leftCell2 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell2.Border = PdfPCell.NO_BORDER;
                        leftCell2.Colspan = 2;
                        PdfPCell leftCell21 = new PdfPCell(new Phrase("Address", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell21.Border = PdfPCell.NO_BORDER;
                        leftCell21.Colspan = 2;
                        PdfPCell centerCell2 = new PdfPCell(new Phrase(":  " + model.consignerAddress, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell2.Border = PdfPCell.NO_BORDER;
                        centerCell2.Colspan = 4;
                        PdfPCell centerCell3 = new PdfPCell(new Phrase("  ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell3.Border = PdfPCell.NO_BORDER;
                        centerCell3.Colspan = 4;
                        PdfPCell leftCell3 = new PdfPCell(new Phrase("2 .Consignee Name", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell3.Border = PdfPCell.NO_BORDER;
                        leftCell3.Colspan = 4;
                        PdfPCell centerCell4 = new PdfPCell(new Phrase(":   " + model.consigneeName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell4.Border = PdfPCell.NO_BORDER;
                        centerCell4.Colspan = 4;
                        PdfPCell centerCell5 = new PdfPCell(new Phrase("   ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell5.Border = PdfPCell.NO_BORDER;
                        centerCell5.Colspan = 4;
                        PdfPCell leftCell4 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell4.Border = PdfPCell.NO_BORDER;
                        leftCell4.Colspan = 2;
                        PdfPCell leftCell5 = new PdfPCell(new Phrase("Address", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell5.Border = PdfPCell.NO_BORDER;
                        leftCell5.Colspan = 2;
                        PdfPCell centerCell6 = new PdfPCell(new Phrase(":  " + model.consigneeAddress, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell6.Border = PdfPCell.NO_BORDER;
                        centerCell6.Colspan = 4;
                        PdfPCell centerCell7 = new PdfPCell(new Phrase("  ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell7.Border = PdfPCell.NO_BORDER;
                        centerCell7.Colspan = 4;
                        PdfPCell leftCell6 = new PdfPCell(new Phrase("3.Product Detail", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell6.Border = PdfPCell.NO_BORDER;
                        leftCell6.Colspan = 12;
                        PdfPCell dataCell1 = new PdfPCell(new Phrase("SNo", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell2 = new PdfPCell(new Phrase("Product", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell3 = new PdfPCell(new Phrase("Brand", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dataCell3.Colspan = 2;
                        PdfPCell dataCell4 = new PdfPCell(new Phrase("Qty(BL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell5 = new PdfPCell(new Phrase("Qty(LPL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell6 = new PdfPCell(new Phrase("Duty Amount", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dataCell6.Colspan = 6;
                        dataCell6.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        dataTable.AddCell(BelowCell);
                        dataTable.AddCell(rightBelowCell);
                        dataTable.AddCell(leftCell1);
                        dataTable.AddCell(centerCell);
                        dataTable.AddCell(centerCell1);
                        dataTable.AddCell(leftCell2);
                        dataTable.AddCell(leftCell21);
                        dataTable.AddCell(centerCell2);
                        dataTable.AddCell(centerCell3);
                        dataTable.AddCell(leftCell3);
                        dataTable.AddCell(centerCell4);
                        dataTable.AddCell(centerCell5);
                        dataTable.AddCell(leftCell4);
                        dataTable.AddCell(leftCell5);
                        dataTable.AddCell(centerCell6);
                        dataTable.AddCell(centerCell7);
                        dataTable.AddCell(leftCell6);
                        dataTable.AddCell(dataCell1);
                        dataTable.AddCell(dataCell2);
                        dataTable.AddCell(dataCell3);
                        dataTable.AddCell(dataCell4);
                        dataTable.AddCell(dataCell5);
                        dataTable.AddCell(dataCell6);
                        int i = 0;
                        foreach (var itm in model.lstProduct)
                        {
                            i = i + 1;
                            PdfPCell dataCell12 = new PdfPCell(new Phrase(Convert.ToString(i), FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell22 = new PdfPCell(new Phrase(itm.productName, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell32 = new PdfPCell(new Phrase(itm.brandName, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataCell32.Colspan = 2;
                            PdfPCell dataCell42 = new PdfPCell(new Phrase(itm.BLQty, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell52 = new PdfPCell(new Phrase(itm.LPLQty, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell62 = new PdfPCell(new Phrase(itm.dutyType, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataCell62.Colspan = 3;
                            PdfPCell dataCell72 = new PdfPCell(new Phrase(itm.dutyAmount, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataCell72.Colspan = 3;
                            dataTable.AddCell(dataCell12);
                            dataTable.AddCell(dataCell22);
                            dataTable.AddCell(dataCell32);
                            dataTable.AddCell(dataCell42);
                            dataTable.AddCell(dataCell52);
                            dataTable.AddCell(dataCell62);
                            dataTable.AddCell(dataCell72);
                        }
                        PdfPCell leftCell61 = new PdfPCell(new Phrase("4.Route", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell61.Border = PdfPCell.NO_BORDER;
                        leftCell61.Colspan = 4;
                        PdfPCell centerCell51 = new PdfPCell(new Phrase(": " + model.routeName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell51.Border = PdfPCell.NO_BORDER;
                        centerCell51.Colspan = 4;
                        PdfPCell centerCell52 = new PdfPCell(new Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell52.Border = PdfPCell.NO_BORDER;
                        centerCell52.Colspan = 4;
                        PdfPCell leftCell7 = new PdfPCell(new Phrase("5.Challan Adjustment Details  :", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell7.Border = PdfPCell.NO_BORDER;
                        leftCell7.Colspan = 12;
                        PdfPCell dataCell11 = new PdfPCell(new Phrase("Challan no/Date", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dataCell11.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        dataCell11.Colspan = 3;
                        PdfPCell dataCell21 = new PdfPCell(new Phrase("Bank", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell31 = new PdfPCell(new Phrase("Fee Type", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dataCell31.Colspan = 3;
                        dataCell31.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell dataCell41 = new PdfPCell(new Phrase("Ch. Amt", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell51 = new PdfPCell(new Phrase("Avl.Amt", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell61 = new PdfPCell(new Phrase("Total Fee", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell71 = new PdfPCell(new Phrase("Adj.Fee", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dataCell81 = new PdfPCell(new Phrase("Avl.Balance", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dataTable.AddCell(leftCell61);
                        dataTable.AddCell(centerCell51);
                        dataTable.AddCell(centerCell52);
                        dataTable.AddCell(leftCell7);
                        dataTable.AddCell(dataCell11);
                        dataTable.AddCell(dataCell21);
                        dataTable.AddCell(dataCell31);
                        dataTable.AddCell(dataCell41);
                        dataTable.AddCell(dataCell51);
                        dataTable.AddCell(dataCell61);
                        dataTable.AddCell(dataCell71);
                        dataTable.AddCell(dataCell81);
                        foreach (var itm in model.lstProduct)
                        {
                            PdfPCell dataCell12 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.challanNo ?? "" + "" + model.lstChallan?.FirstOrDefault()?.challanDate ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataCell12.Colspan = 3;
                            PdfPCell dataCell22 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.bankName ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell32 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.feeType ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataCell32.Colspan = 3;
                            PdfPCell dataCell42 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.challanAmount ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell52 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.availAmount ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell62 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.totalFee ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell72 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.adjustmentAmount ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell dataCell82 = new PdfPCell(new Phrase(model.lstChallan?.FirstOrDefault()?.availBalance ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataTable.AddCell(dataCell12);
                            dataTable.AddCell(dataCell22);
                            dataTable.AddCell(dataCell32);
                            dataTable.AddCell(dataCell42);
                            dataTable.AddCell(dataCell52);
                            dataTable.AddCell(dataCell62);
                            dataTable.AddCell(dataCell72);
                            dataTable.AddCell(dataCell82);
                        }
                        PdfPCell belowCell = new PdfPCell(new Phrase("6.Valid Upto", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell.Border = PdfPCell.NO_BORDER;
                        belowCell.Colspan = 4;
                        PdfPCell belowCell1 = new PdfPCell(new Phrase(":" + model.Fl5ValidityDate, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell1.Border = PdfPCell.NO_BORDER;
                        belowCell1.Colspan = 4;
                        PdfPCell belowCell2 = new PdfPCell(new Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell2.Border = PdfPCell.NO_BORDER;
                        belowCell2.Colspan = 4;
                        PdfPCell belowCell3 = new PdfPCell(new Phrase("7.Seal No.", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell3.Border = PdfPCell.NO_BORDER;
                        belowCell3.Colspan = 4;
                        PdfPCell belowCell4 = new PdfPCell(new Phrase(":" + model.sealNo, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell4.Border = PdfPCell.NO_BORDER;
                        belowCell4.Colspan = 4;
                        PdfPCell belowCell5 = new PdfPCell(new Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell5.Border = PdfPCell.NO_BORDER;
                        belowCell5.Colspan = 4;
                        PdfPCell belowCell6 = new PdfPCell(new Phrase("8.Name of Escort", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell6.Border = PdfPCell.NO_BORDER;
                        belowCell6.Colspan = 4;
                        PdfPCell belowCell7 = new PdfPCell(new Phrase(":" + model.nameOfEscort, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell7.Border = PdfPCell.NO_BORDER;
                        belowCell7.Colspan = 4;
                        PdfPCell belowCell8 = new PdfPCell(new Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell8.Border = PdfPCell.NO_BORDER;
                        belowCell8.Colspan = 4;
                        PdfPCell belowCell9 = new PdfPCell(new Phrase("Authorised Digital Signature by:- ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell9.Border = PdfPCell.NO_BORDER;
                        belowCell9.Colspan = 12;
                        belowCell9.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        PdfPCell belowCell10 = new PdfPCell(new Phrase("DEO: ALWAR", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell10.Border = PdfPCell.NO_BORDER;
                        belowCell10.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        belowCell10.Colspan = 12;
                        PdfPCell belowCell11 = new PdfPCell(new Phrase("SHUBHAM KUMAR", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell11.Border = PdfPCell.NO_BORDER;
                        belowCell11.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        belowCell11.Colspan = 12;
                        PdfPCell belowCell12 = new PdfPCell(new Phrase("Date: 1-JAN-2024", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell12.Border = PdfPCell.NO_BORDER;
                        belowCell12.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        belowCell12.Colspan = 12;
                        PdfPCell belowCell31 = new PdfPCell(new Phrase("1.CC to UNITED SPIRITS LTD. UNIT ALWAR \r\nFor their import No 0 dated 08/11/2023", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell31.Border = PdfPCell.NO_BORDER;
                        belowCell31.Colspan = 12;
                        PdfPCell belowCell41 = new PdfPCell(new Phrase("CC for Information and Necessary Action", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        belowCell41.Border = PdfPCell.NO_BORDER;
                        belowCell41.Colspan = 12;
                        PdfPCell belowCell51 = new PdfPCell(new Phrase("2.CC to DEO (Consigner) other", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell51.Border = PdfPCell.NO_BORDER;
                        belowCell51.Colspan = 12;
                        PdfPCell belowCell61 = new PdfPCell(new Phrase("3.CC to DEO (Consigner) Alwar/AEO/CL-", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell61.Border = PdfPCell.NO_BORDER;
                        belowCell61.Colspan = 12;
                        PdfPCell belowCell71 = new PdfPCell(new Phrase("4. Office Copy", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        belowCell71.Border = PdfPCell.NO_BORDER;
                        belowCell71.Colspan = 12;
                        dataTable.AddCell(belowCell);
                        dataTable.AddCell(belowCell1);
                        dataTable.AddCell(belowCell2);
                        dataTable.AddCell(belowCell3);
                        dataTable.AddCell(belowCell4);
                        dataTable.AddCell(belowCell5);
                        dataTable.AddCell(belowCell6);
                        dataTable.AddCell(belowCell7);
                        dataTable.AddCell(belowCell8);
                        dataTable.AddCell(belowCell9);
                        dataTable.AddCell(belowCell10);
                        dataTable.AddCell(belowCell11);
                        dataTable.AddCell(belowCell12);
                        dataTable.AddCell(belowCell31);
                        dataTable.AddCell(belowCell41);
                        dataTable.AddCell(belowCell51);
                        dataTable.AddCell(belowCell61);
                        dataTable.AddCell(belowCell71);
                        PdfContentByte canvas = writer.DirectContent;
                        PdfTemplate template = canvas.CreateTemplate(595, 842);
                        ColumnText dataTableColumnText = new ColumnText(template);
                        dataTableColumnText.SetSimpleColumn(new Rectangle(0, 0, 600, 780));
                        dataTableColumnText.AddElement(dataTable);
                        dataTableColumnText.Go();
                        canvas.AddTemplate(template, 0, 0);
                        document.Close();
                        fileContents = memoryStream.ToArray();
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return fileContents;
        }
        public static byte[] MisFinal(FinalMISResponseModel model, string logoPath, string reqType)
        {

            byte[] fileContents;
            try
            {
                string suffix = string.Empty;
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Document document = new Document(new Rectangle(595, 842), 0f, 0f, 0f, 0f))
                    {
                        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                        document.Open();
                        #region Added image Logo
                        var imagePath = logoPath;
                        Image myImage = Image.GetInstance(imagePath);
                        float fixedWidth = logoPath.ToLower().Contains("excise") ? 150 : 380;
                        float fixedHeight = 150;
                        float widthScalingFactor = fixedWidth / myImage.Width;
                        float heightScalingFactor = fixedHeight / myImage.Height;
                        float scalingFactor = Math.Min(widthScalingFactor, heightScalingFactor);
                        myImage.ScaleAbsolute(myImage.Width * scalingFactor, myImage.Height * scalingFactor);
                        float xCentered = (document.PageSize.Width - myImage.ScaledWidth) / 2;
                        myImage.SetAbsolutePosition(xCentered, document.PageSize.Height - myImage.ScaledHeight - document.TopMargin);
                        document.Add(myImage);
                        #endregion

                        Font blueFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLUE);
                        Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
                        PdfPTable leftTable = new PdfPTable(1);
                        PdfPCell leftCell = new PdfPCell(new Phrase("RAJASTHAN STATE BEVERAGES CORPORATION LIMITED", blueFont));
                        leftCell.PaddingTop = 50;
                        leftCell.PaddingLeft = 90;
                        leftCell.Border = PdfPCell.NO_BORDER;
                        PdfPCell leftCell1 = new PdfPCell(new Phrase("(A Govt. of Rajasthan Undertaking)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        leftCell1.Border = PdfPCell.NO_BORDER;
                        leftCell1.PaddingLeft = 155;
                        PdfPCell leftCell2 = new PdfPCell(new Phrase("Material Inward Slip(MIS)", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLUE)));
                        leftCell2.Border = PdfPCell.NO_BORDER;
                        leftCell2.PaddingLeft = 155;
                        leftCell2.PaddingTop = 5;
                        leftCell1.PaddingLeft = 155;

                        if (reqType == "OFS")
                        {
                            suffix = "MIS ";
                        }
                        else
                        {
                            suffix = "TIS ";
                        }
                        PdfPCell leftCell3 = new PdfPCell(new Phrase(suffix + "NO: " + model.MISDetails.MIS_NO, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        leftCell3.Border = PdfPCell.NO_BORDER;
                        leftCell3.PaddingTop = 3;



                        PdfPTable leftTable1 = new PdfPTable(1);
                        PdfPCell leftTableCell = new PdfPCell(new Phrase(suffix + "Date :" + model.MISDetails.MIS_Date, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        leftTableCell.Border = PdfPCell.NO_BORDER;
                        leftTable1.AddCell(leftTableCell);

                        PdfPTable leftTable2 = new PdfPTable(1);
                        PdfPCell leftTableCell2 = new PdfPCell(new Phrase("Supplier Name :" + model.MISDetails.SuppierName, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        leftTableCell2.Border = PdfPCell.NO_BORDER;
                        leftTable2.AddCell(leftTableCell2);

                        PdfPTable leftTable3 = new PdfPTable(1);
                        PdfPCell leftTableCell3 = new PdfPCell(new Phrase(reqType + " Number : " + model.MISDetails.OFSNO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        leftTableCell3.Border = PdfPCell.NO_BORDER;
                        leftTable3.AddCell(leftTableCell3);

                        PdfPTable leftTable4 = new PdfPTable(1);
                        PdfPCell leftTableCell4 = new PdfPCell(new Phrase(reqType + " DATE: " + model.MISDetails.OFSDate, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        leftTableCell4.Border = PdfPCell.NO_BORDER;
                        leftTable4.AddCell(leftTableCell4);

                        PdfPTable leftTable5 = new PdfPTable(1);
                        PdfPCell leftTableCell5 = new PdfPCell(new Phrase(reqType + " VALIDITY : " + model.MISDetails.OFS_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        leftTableCell5.Border = PdfPCell.NO_BORDER;
                        leftTable5.AddCell(leftTableCell5);




                        PdfContentByte canvas = writer.DirectContent;
                        PdfTemplate template = canvas.CreateTemplate(595, 842);
                        ColumnText leftColumnText = new ColumnText(template);
                        leftColumnText.SetSimpleColumn(new Rectangle(0, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 110));
                        leftColumnText.AddElement(leftTable1);
                        leftColumnText.AddElement(leftTable2);
                        leftColumnText.AddElement(leftTable3);
                        leftColumnText.AddElement(leftTable4);
                        leftColumnText.AddElement(leftTable5);
                        leftColumnText.Go();


                        PdfPTable centerTable1 = new PdfPTable(1);
                        PdfPCell centerTableCell = new PdfPCell(new Phrase("FL-5 NO : " + model.MISDetails.FL5_NO, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell.Border = PdfPCell.NO_BORDER;
                        centerTable1.AddCell(centerTableCell);

                        PdfPTable centerTable2 = new PdfPTable(1);
                        PdfPCell centerTableCell2 = new PdfPCell(new Phrase("FL-5 DATE : " + model.MISDetails.FL5_Date, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell2.Border = PdfPCell.NO_BORDER;
                        centerTable2.AddCell(centerTableCell2);

                        PdfPTable centerTable3 = new PdfPTable(1);
                        PdfPCell centerTableCell3 = new PdfPCell(new Phrase("FL-5 VALIDITY :" + model.MISDetails.FL5_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell3.Border = PdfPCell.NO_BORDER;
                        centerTable3.AddCell(centerTableCell3);

                        PdfPTable centerTable4 = new PdfPTable(1);
                        PdfPCell centerTableCell4 = new PdfPCell(new Phrase("TP NO : " + model.MISDetails.TP_NO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell4.Border = PdfPCell.NO_BORDER;
                        centerTableCell4.PaddingTop = 5;
                        centerTable4.AddCell(centerTableCell4);

                        PdfPTable centerTable5 = new PdfPTable(1);
                        PdfPCell centerTableCell5 = new PdfPCell(new Phrase("TP Date :" + model.MISDetails.TP_Date, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell5.Border = PdfPCell.NO_BORDER;
                        centerTable5.AddCell(centerTableCell5);

                        PdfPTable centerTable6 = new PdfPTable(1);
                        PdfPCell centerTableCell6 = new PdfPCell(new Phrase("TP Validity : " + model.MISDetails.TP_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell6.Border = PdfPCell.NO_BORDER;
                        centerTable6.AddCell(centerTableCell6);
                        PdfPTable centerTable7 = new PdfPTable(1);
                        PdfPCell centerTableCell7 = new PdfPCell(new Phrase("INVOICE NO : " + model.MISDetails.InvoiceNO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        centerTableCell7.Border = PdfPCell.NO_BORDER;
                        centerTable7.AddCell(centerTableCell7);

                        ColumnText centerColumnText = new ColumnText(template);
                        centerColumnText.SetSimpleColumn(new Rectangle(PageSize.A4.Width / 3, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 110));
                        centerColumnText.AddElement(centerTable1);
                        centerColumnText.AddElement(centerTable2);
                        centerColumnText.AddElement(centerTable3);
                        centerColumnText.AddElement(centerTable4);
                        centerColumnText.AddElement(centerTable5);
                        centerColumnText.AddElement(centerTable6);
                        centerColumnText.AddElement(centerTable7);
                        centerColumnText.Go();

                        PdfPTable rightTable1 = new PdfPTable(1);
                        PdfPCell rightTableCell = new PdfPCell(new Phrase("Gate Entry Date/Time " + model.MISDetails.GateEntryOn, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        rightTableCell.Border = PdfPCell.NO_BORDER;
                        rightTable1.AddCell(rightTableCell);

                        Font smallFont = FontFactory.GetFont(FontFactory.HELVETICA, 6, Font.BOLD, BaseColor.BLACK);
                        Font bFont = FontFactory.GetFont(FontFactory.HELVETICA, 6, Font.NORMAL, BaseColor.BLACK);



                        PdfPTable dataTable = new PdfPTable(15);
                        PdfPCell sno = new PdfPCell(new Phrase("SNO", smallFont));
                        sno.Rowspan = 2;
                        PdfPCell product = new PdfPCell(new Phrase("Product", smallFont));
                        product.Rowspan = 2;
                        PdfPCell ofs = new PdfPCell(new Phrase(reqType, smallFont));
                        PdfPCell invoice = new PdfPCell(new Phrase("Invoice", smallFont));
                        PdfPCell good = new PdfPCell(new Phrase("Good", smallFont));
                        good.Colspan = 2;
                        good.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell damageQty = new PdfPCell(new Phrase("Damage Qty", smallFont));
                        damageQty.Colspan = 2;
                        damageQty.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell shortageQty = new PdfPCell(new Phrase("Shortage Qty", smallFont));
                        shortageQty.Colspan = 2;
                        shortageQty.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell qtyBl = new PdfPCell(new Phrase("Qty(BL)", smallFont));
                        qtyBl.Rowspan = 2;
                        PdfPCell batchNo = new PdfPCell(new Phrase("BatchNo", smallFont));
                        batchNo.Rowspan = 2;
                        PdfPCell mfgDate = new PdfPCell(new Phrase("Mfg Date", smallFont));
                        mfgDate.Rowspan = 2;
                        PdfPCell amount = new PdfPCell(new Phrase("Amount", smallFont));
                        amount.Rowspan = 2;
                        PdfPCell unreconciledQty = new PdfPCell(new Phrase("Unreconclied Qty", smallFont));

                        unreconciledQty.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell ofsCase = new PdfPCell(new Phrase("Case", smallFont));
                        PdfPCell invoiceCase = new PdfPCell(new Phrase("Case", smallFont));
                        PdfPCell goodCase = new PdfPCell(new Phrase("Case", smallFont));
                        PdfPCell goodbottle = new PdfPCell(new Phrase("Btls", smallFont));
                        PdfPCell damageCase = new PdfPCell(new Phrase("Case", smallFont));
                        PdfPCell damagebottle = new PdfPCell(new Phrase("Btls", smallFont));
                        PdfPCell shortageCase = new PdfPCell(new Phrase("Case", smallFont));
                        PdfPCell shortagebottle = new PdfPCell(new Phrase("Btls", smallFont));
                        PdfPCell offlineQtyCase = new PdfPCell(new Phrase("case", smallFont));



                        dataTable.AddCell(sno);
                        dataTable.AddCell(product);
                        dataTable.AddCell(ofs);
                        dataTable.AddCell(invoice);
                        dataTable.AddCell(good);
                        dataTable.AddCell(damageQty);
                        dataTable.AddCell(shortageQty);
                        dataTable.AddCell(qtyBl);
                        dataTable.AddCell(batchNo);
                        dataTable.AddCell(mfgDate);
                        dataTable.AddCell(amount);
                        dataTable.AddCell(unreconciledQty);
                        dataTable.AddCell(ofsCase);
                        dataTable.AddCell(invoiceCase);
                        dataTable.AddCell(goodCase);
                        dataTable.AddCell(goodbottle);
                        dataTable.AddCell(damageCase);
                        dataTable.AddCell(damagebottle);
                        dataTable.AddCell(shortageCase);
                        dataTable.AddCell(shortagebottle);
                        dataTable.AddCell(offlineQtyCase);


                        int i = 0;
                        PdfPTable dataTable1 = null;
                        foreach (var itm in model.MISStockList)
                        {
                            i = i + 1;
                            dataTable1 = new PdfPTable(15);
                            PdfPCell sncell = new PdfPCell(new Phrase(Convert.ToString(i), bFont));
                            sncell.Rowspan = 2;
                            PdfPCell productcell = new PdfPCell(new Phrase(itm.ProductName, bFont));
                            productcell.Rowspan = 2;

                            PdfPCell ofscell = new PdfPCell(new Phrase(Convert.ToString(itm.OFSQty), bFont));

                            PdfPCell invoicecell = new PdfPCell(new Phrase(Convert.ToString(itm.InvoiceQty), bFont));

                            PdfPCell goodCasecell = new PdfPCell(new Phrase(Convert.ToString(itm.GoodCasesQty), bFont));
                            PdfPCell goodBottleCell = new PdfPCell(new Phrase(Convert.ToString(itm.GoodBottleQty), bFont));
                            PdfPCell DamageCasecell = new PdfPCell(new Phrase(Convert.ToString(itm.DamageCasesQty), bFont));
                            PdfPCell DamageBottleCell = new PdfPCell(new Phrase(Convert.ToString(itm.DamageBottleQty), bFont));
                            PdfPCell shortageCasecell = new PdfPCell(new Phrase(Convert.ToString(itm.ShortageCasesQty), bFont));
                            PdfPCell shortageBottleCell = new PdfPCell(new Phrase(Convert.ToString(itm.ShortageBottleQty), bFont));
                            PdfPCell qtyBlcell = new PdfPCell(new Phrase(Convert.ToString(itm.GoodBLQty), bFont));
                            PdfPCell batchNoCell = new PdfPCell(new Phrase(Convert.ToString(itm.BatchNo), bFont));
                            PdfPCell mfgDatecell = new PdfPCell(new Phrase(Convert.ToString(itm.MfgDate), bFont));
                            PdfPCell amountcell = new PdfPCell(new Phrase(Convert.ToString(itm.PurchaseValue), bFont));
                            PdfPCell offlineQtyCasecell = new PdfPCell(new Phrase(Convert.ToString(itm.UnReconcilled), bFont));

                            dataTable1.AddCell(sncell);
                            dataTable1.AddCell(productcell);
                            dataTable1.AddCell(ofscell);
                            dataTable1.AddCell(invoicecell);
                            dataTable1.AddCell(goodCasecell);
                            dataTable1.AddCell(goodBottleCell);
                            dataTable1.AddCell(DamageCasecell);
                            dataTable1.AddCell(DamageBottleCell);
                            dataTable1.AddCell(shortageCasecell);
                            dataTable1.AddCell(shortageBottleCell);
                            dataTable1.AddCell(qtyBlcell);
                            dataTable1.AddCell(batchNoCell);
                            dataTable1.AddCell(mfgDatecell);
                            dataTable1.AddCell(amountcell);
                            dataTable1.AddCell(offlineQtyCasecell);
                        }

                        ColumnText dtColumnText = new ColumnText(template);
                        dtColumnText.SetSimpleColumn(new Rectangle(0, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 220));
                        dtColumnText.AddElement(dataTable);
                        dtColumnText.AddElement(dataTable1);

                        dtColumnText.Go();

                        PdfPTable BeowTable = new PdfPTable(1);
                        PdfPCell BelowCell = new PdfPCell(new Phrase("For Rajasthan State Beverages Corporation Ltd.", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell.Border = PdfPCell.NO_BORDER;
                        BeowTable.AddCell(BelowCell);
                        PdfPTable BeowTable1 = new PdfPTable(1);
                        PdfPCell BelowCell1 = new PdfPCell(new Phrase(" Depot Manager", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell1.Border = PdfPCell.NO_BORDER;
                        BelowCell1.PaddingTop = 15;
                        BeowTable1.AddCell(BelowCell1);
                        PdfPTable BeowTable2 = new PdfPTable(1);
                        PdfPCell BelowCell2 = new PdfPCell(new Phrase(" DM Signature", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell2.Border = PdfPCell.NO_BORDER;
                        BelowCell2.PaddingTop = 35;
                        BeowTable2.AddCell(BelowCell2);
                        ColumnText BelowColumnText = new ColumnText(template);
                        BelowColumnText.SetSimpleColumn(new Rectangle(PageSize.A4.Width / 2 + 90, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 520));
                        BelowColumnText.AddElement(BeowTable);
                        BelowColumnText.AddElement(BeowTable1);
                        BelowColumnText.AddElement(BeowTable2);
                        BelowColumnText.Go();

                        PdfPTable footer = new PdfPTable(1);
                        PdfPCell footercell = new PdfPCell(new Phrase(" Godown Assistant", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        footercell.Border = PdfPCell.NO_BORDER;
                        footercell.PaddingTop = 25;

                        footer.AddCell(footercell);

                        PdfPTable footer1 = new PdfPTable(1);
                        PdfPCell footercell1 = new PdfPCell(new Phrase(" Driver Name", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        footercell1.Border = PdfPCell.NO_BORDER;
                        footercell1.PaddingTop = 35;

                        footer1.AddCell(footercell1);
                        ColumnText footerColumnText = new ColumnText(template);
                        footerColumnText.SetSimpleColumn(new Rectangle(0, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 520));
                        footerColumnText.AddElement(footer);
                        footerColumnText.AddElement(footer1);
                        footerColumnText.Go();

                        PdfPTable footer2 = new PdfPTable(1);
                        PdfPCell footercell12 = new PdfPCell(new Phrase(" Driver Signature", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        footercell12.Border = PdfPCell.NO_BORDER;
                        footercell12.PaddingTop = 80;
                        footer2.AddCell(footercell12);
                        ColumnText rightColumnText1 = new ColumnText(template);
                        rightColumnText1.SetSimpleColumn(new Rectangle(new Rectangle(200, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 520)));
                        rightColumnText1.AddElement(footer2);
                        rightColumnText1.Go();



                        ColumnText rightColumnText = new ColumnText(template);
                        rightColumnText.SetSimpleColumn(new Rectangle(PageSize.A4.Width / 2 + 90, 0, PageSize.A4.Width + 16, PageSize.A4.Height - 110));
                        rightColumnText.AddElement(rightTable1);

                        rightColumnText.Go();


                        leftTable.AddCell(leftCell);
                        leftTable.AddCell(leftCell1);
                        leftTable.AddCell(leftCell2);
                        leftTable.AddCell(leftCell3);


                        document.Add(leftTable);
                        canvas.AddTemplate(template, 0, 0);
                        document.Close();
                        fileContents = memoryStream.ToArray();

                    }

                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return fileContents;

        }
        public static byte[] DraftMs(MISViewResponseModel model, string ReqType, string logoPath)
        {
            byte[] fileContents;
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Document document = new Document(new Rectangle(595, 842), 0f, 0f, 0f, 0f))
                    {
                        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                        document.Open();
                        #region Added image Logo
                        var imagePath = logoPath;
                        Image myImage = Image.GetInstance(imagePath);
                        float fixedWidth = logoPath.ToLower().Contains("excise") ? 150 : 380;
                        float fixedHeight = 150;
                        float widthScalingFactor = fixedWidth / myImage.Width;
                        float heightScalingFactor = fixedHeight / myImage.Height;
                        float scalingFactor = Math.Min(widthScalingFactor, heightScalingFactor);
                        myImage.ScaleAbsolute(myImage.Width * scalingFactor, myImage.Height * scalingFactor);
                        float xCentered = (document.PageSize.Width - myImage.ScaledWidth) / 2;
                        myImage.SetAbsolutePosition(xCentered, document.PageSize.Height - myImage.ScaledHeight - document.TopMargin);
                        document.Add(myImage);
                        #endregion

                        Font blueFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLUE);
                        Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
                        PdfPTable leftTable = new PdfPTable(18);
                        PdfPCell leftCell = new PdfPCell(new Phrase("", blueFont));
                        leftCell.PaddingLeft = 90;
                        leftCell.Border = PdfPCell.NO_BORDER;
                        leftCell.Colspan = 18;

                        PdfPCell leftCell1 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        leftCell1.Border = PdfPCell.NO_BORDER;
                        leftCell1.PaddingLeft = 155;
                        leftCell1.PaddingTop = 20;
                        leftCell1.Colspan = 18;


                        PdfPCell leftCell2 = new PdfPCell(new Phrase(" Material Inward Check List", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLUE)));
                        leftCell2.Border = PdfPCell.NO_BORDER;
                        leftCell2.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        leftCell2.Colspan = 18;


                        PdfPCell leftCell3;

                        if (ReqType == "ofs")
                        {
                            Phrase misNo = new Phrase("MIS NO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK));
                            Phrase misValue = new Phrase(model.MISDetails.MIS_NO,
                                FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE));
                            leftCell3 = new PdfPCell(new Phrase("MIS NO " + model.MISDetails.MIS_NO,
                                FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        }
                        else
                        {
                            leftCell3 = new PdfPCell(new Phrase("TIS NO " + model.MISDetails.MIS_NO, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        }
                        leftCell3.Border = PdfPCell.NO_BORDER;
                        leftCell3.Left = PdfPCell.ALIGN_LEFT;
                        leftCell3.Colspan = 18;
                        PdfPCell leftTableCell = new PdfPCell(new Phrase("To\r\n" + model.MISDetails.SuppierName + "\r\n,\r\n" + model.MISDetails.SupplierAddress, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        leftTableCell.Border = PdfPCell.NO_BORDER;
                        leftTableCell.Colspan = 18;
                        PdfPCell leftTableCell3;
                        if (ReqType == "ofs")
                        {
                            leftTableCell3 = new PdfPCell(new Phrase("MIS Date : " + model.MISDetails.MIS_Date, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }
                        else
                        {
                            leftTableCell3 = new PdfPCell(new Phrase("TIS Date : " + model.MISDetails.MIS_Date, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }

                        leftTableCell3.Border = PdfPCell.NO_BORDER;
                        leftTableCell3.Colspan = 6;
                        PdfPCell leftTableCell4 = new PdfPCell(new Phrase("Suppier Name : " + model.MISDetails.SuppierName, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        leftTableCell4.Border = PdfPCell.NO_BORDER;
                        leftTableCell4.Colspan = 6;


                        PdfPCell leftTableCell5;
                        if (ReqType == "ofs")
                        {
                            leftTableCell5 = new PdfPCell(new Phrase("OFS NO:" + model.MISDetails.OFSNO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }
                        else
                        {
                            leftTableCell5 = new PdfPCell(new Phrase("TOO NO:" + model.MISDetails.OFSNO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }
                        leftTableCell5.Border = PdfPCell.NO_BORDER;
                        leftTableCell5.Colspan = 6;

                        PdfPCell centerTableCell;
                        if (ReqType == "ofs")
                        {
                            centerTableCell = new PdfPCell(new Phrase("OFS Date: " + model.MISDetails.OFSDate, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }
                        else
                        {
                            centerTableCell = new PdfPCell(new Phrase("TOO Date: " + model.MISDetails.OFSDate, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }

                        centerTableCell.Border = PdfPCell.NO_BORDER;
                        centerTableCell.Colspan = 6;

                        PdfPCell centerTableCell2;
                        if (ReqType == "ofs")
                        {
                            centerTableCell2 = new PdfPCell(new Phrase("OFS Validity: " + model.MISDetails.OFS_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }
                        else
                        {
                            centerTableCell2 = new PdfPCell(new Phrase("TOO Validity: " + model.MISDetails.OFS_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        }
                        centerTableCell2.Border = PdfPCell.NO_BORDER;
                        centerTableCell2.Colspan = 6;
                        PdfPCell centerTableCell3 = new PdfPCell(new Phrase("FL-5 No:" + model.MISDetails.FL5_NO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        centerTableCell3.Border = PdfPCell.NO_BORDER;
                        centerTableCell3.Colspan = 6;
                        PdfPCell centerTableCell4 = new PdfPCell(new Phrase("FL-5 Date :  " + model.MISDetails.FL5_Date, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        centerTableCell4.Border = PdfPCell.NO_BORDER;
                        centerTableCell4.Colspan = 6;
                        PdfPCell centerTableCell5 = new PdfPCell(new Phrase("FL-5 Validity " + model.MISDetails.FL5_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        centerTableCell5.Border = PdfPCell.NO_BORDER;
                        centerTableCell5.Colspan = 6;
                        PdfPCell centerTableCell6 = new PdfPCell(new Phrase("TP NO: " + model.MISDetails.TP_NO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        centerTableCell6.Border = PdfPCell.NO_BORDER;
                        centerTableCell6.Colspan = 6;

                        PdfPCell centerTableCell7 = new PdfPCell(new Phrase("TP Date : " + model.MISDetails.TP_Date, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.BOLD, BaseColor.BLACK)));
                        centerTableCell7.Border = PdfPCell.NO_BORDER;
                        centerTableCell7.Colspan = 6;
                        PdfPCell rightTableCell = new PdfPCell(new Phrase("TP Validity : " + model.MISDetails.TP_Validity, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        rightTableCell.Border = PdfPCell.NO_BORDER;
                        rightTableCell.Colspan = 6;
                        PdfPCell rightTableCell2;
                        if (ReqType == "ofs")
                        {
                            rightTableCell2 = new PdfPCell(new Phrase("Invoice No : " + model.MISDetails.InvoiceNO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        }
                        else
                        {
                            rightTableCell2 = new PdfPCell(new Phrase("TOS No : " + model.MISDetails.InvoiceNO, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        }
                        rightTableCell2.Border = PdfPCell.NO_BORDER;
                        rightTableCell2.Colspan = 6;

                        PdfPCell rightTableCell3 = new PdfPCell(new Phrase("Gate Entry Date/Time : " + model.MISDetails.GateEntryOn, FontFactory.GetFont(FontFactory.HELVETICA, 7, Font.NORMAL, BaseColor.BLACK)));
                        rightTableCell3.Border = PdfPCell.NO_BORDER;
                        rightTableCell3.PaddingTop = 25;
                        rightTableCell3.Colspan = 18;

                        leftTable.AddCell(leftCell);
                        leftTable.AddCell(leftCell1);
                        leftTable.AddCell(leftCell2);
                        leftTable.AddCell(leftCell3);
                        leftTable.AddCell(leftTableCell);
                        leftTable.AddCell(leftTableCell3);
                        leftTable.AddCell(leftTableCell4);
                        leftTable.AddCell(leftTableCell5);
                        leftTable.AddCell(centerTableCell);
                        leftTable.AddCell(centerTableCell2);
                        leftTable.AddCell(centerTableCell3);
                        leftTable.AddCell(centerTableCell4);
                        leftTable.AddCell(centerTableCell5);
                        leftTable.AddCell(centerTableCell6);
                        leftTable.AddCell(centerTableCell7);
                        leftTable.AddCell(rightTableCell);
                        leftTable.AddCell(rightTableCell2);
                        leftTable.AddCell(rightTableCell3);

                        Font smallFont = FontFactory.GetFont(FontFactory.HELVETICA, 6, Font.BOLD, BaseColor.BLACK);
                        Font bFont = FontFactory.GetFont(FontFactory.HELVETICA, 6, Font.NORMAL, BaseColor.BLACK);

                        PdfPCell dtcell1 = new PdfPCell(new Phrase("SNO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        dtcell1.Rowspan = 2;

                        PdfPCell dtcell2 = new PdfPCell(new Phrase("Product", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        //dtcell2.Colspan = 2;
                        dtcell2.Rowspan = 2;
                        dtcell2.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell dtcell3 = new PdfPCell(new Phrase("Packing", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        dtcell3.Rowspan = 2;
                        PdfPCell dtcell4 = new PdfPCell(new Phrase("BatchNo", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        dtcell4.Rowspan = 2;
                        PdfPCell dtcell5 = new PdfPCell(new Phrase("Mfg Date", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        dtcell5.Rowspan = 2;
                        PdfPCell ofsCaseQty = new PdfPCell(new Phrase("Case Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        ofsCaseQty.Colspan = 4;
                        ofsCaseQty.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell shortagebottleQty = new PdfPCell(new Phrase("Bottle Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        shortagebottleQty.Colspan = 4;
                        shortagebottleQty.HorizontalAlignment = PdfPCell.ALIGN_CENTER;

                        PdfPCell blQty = new PdfPCell(new Phrase("BL Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        blQty.Rowspan = 2;
                        PdfPCell rackNo = new PdfPCell(new Phrase("Rack No", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        rackNo.Rowspan = 2;
                        PdfPCell binNo = new PdfPCell(new Phrase("Bin NO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        binNo.Rowspan = 2;
                        PdfPCell sampleStatus = new PdfPCell(new Phrase("Sample Status", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        sampleStatus.Rowspan = 2;
                        PdfPCell amount = new PdfPCell(new Phrase("Amount", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        amount.Rowspan = 2;

                        PdfPCell dtcell12 = new PdfPCell(new Phrase(ReqType.ToUpper(), FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell13 = new PdfPCell(new Phrase("Invoice", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell14 = new PdfPCell(new Phrase("Damage", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell15 = new PdfPCell(new Phrase("Shortage", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell16 = new PdfPCell(new Phrase("OFS", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell17 = new PdfPCell(new Phrase("Invoice", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell18 = new PdfPCell(new Phrase("Damage", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell dtcell19 = new PdfPCell(new Phrase("Shortage", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));

                        leftTable.AddCell(dtcell1);
                        leftTable.AddCell(dtcell2);
                        leftTable.AddCell(dtcell3);
                        leftTable.AddCell(dtcell4);
                        leftTable.AddCell(dtcell5);
                        leftTable.AddCell(ofsCaseQty);
                        leftTable.AddCell(shortagebottleQty);
                        leftTable.AddCell(blQty);
                        leftTable.AddCell(rackNo);
                        leftTable.AddCell(binNo);
                        leftTable.AddCell(sampleStatus);
                        leftTable.AddCell(amount);
                        leftTable.AddCell(dtcell12);
                        leftTable.AddCell(dtcell13);
                        leftTable.AddCell(dtcell14);
                        leftTable.AddCell(dtcell15);
                        leftTable.AddCell(dtcell16);
                        leftTable.AddCell(dtcell17);
                        leftTable.AddCell(dtcell18);
                        leftTable.AddCell(dtcell19);
                        PdfPCell shortageCaseQty = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell invoiceCaseQty = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell damageCaseQty = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell ofsBottleQty = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell invoiceBottleQty = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        PdfPCell damageBottleQty = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));

                        int i = 0;
                        foreach (var item in model.InvoiceList)
                        {
                            i = i + 1;
                            dtcell1 = new PdfPCell(new Phrase(i.ToString(), smallFont));

                            dtcell2 = new PdfPCell(new Phrase(item.ProductName, smallFont));
                            //dtcell2.Colspan = 2;
                            dtcell3 = new PdfPCell(new Phrase(item.Packing, smallFont));

                            dtcell4 = new PdfPCell(new Phrase(item.BatchNo, smallFont));

                            dtcell5 = new PdfPCell(new Phrase(item.MfgDate, smallFont));
                            damageCaseQty = new PdfPCell(new Phrase("0", smallFont));
                            shortageCaseQty = new PdfPCell(new Phrase("0", smallFont));
                            invoiceCaseQty = new PdfPCell(new Phrase(Convert.ToString(item.CasesQTY), smallFont));



                            ofsCaseQty = new PdfPCell(new Phrase(Convert.ToString(item.CasesQTY), smallFont));
                            ofsBottleQty = new PdfPCell(new Phrase("0", smallFont));
                            invoiceBottleQty = new PdfPCell(new Phrase("0", smallFont));
                            damageBottleQty = new PdfPCell(new Phrase("0", smallFont));


                            shortagebottleQty = new PdfPCell(new Phrase("0", smallFont));

                            blQty = new PdfPCell(new Phrase(item.BLQty, smallFont));

                            rackNo = new PdfPCell(new Phrase(item.RackNo, smallFont));

                            binNo = new PdfPCell(new Phrase(item.BinNo, smallFont));

                            if (item.IsSamplingReq == 0)
                            {
                                sampleStatus = new PdfPCell(new Phrase("NO", smallFont));
                            }
                            else
                            {
                                sampleStatus = new PdfPCell(new Phrase("YES", smallFont));
                            }

                            amount = new PdfPCell(new Phrase(item.Amount, smallFont));

                            leftTable.AddCell(dtcell1);
                            leftTable.AddCell(dtcell2);
                            leftTable.AddCell(dtcell3);
                            leftTable.AddCell(dtcell4);
                            leftTable.AddCell(dtcell5);

                            leftTable.AddCell(ofsCaseQty);
                            leftTable.AddCell(invoiceCaseQty);
                            leftTable.AddCell(damageCaseQty);
                            leftTable.AddCell(shortageCaseQty);

                            leftTable.AddCell(ofsBottleQty);
                            leftTable.AddCell(invoiceBottleQty);
                            leftTable.AddCell(damageBottleQty);
                            leftTable.AddCell(shortagebottleQty);

                            leftTable.AddCell(blQty);
                            leftTable.AddCell(rackNo);
                            leftTable.AddCell(binNo);
                            leftTable.AddCell(sampleStatus);
                            leftTable.AddCell(amount);
                        }

                        PdfPCell BelowCell = new PdfPCell(new Phrase("For Rajasthan State Beverages Corporation Ltd.", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell.Border = PdfPCell.NO_BORDER;
                        BelowCell.PaddingTop = 10;

                        BelowCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        BelowCell.Colspan = 18;

                        PdfPCell footercell = new PdfPCell(new Phrase(" Signature Depot Manager", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        footercell.Border = PdfPCell.NO_BORDER;
                        footercell.PaddingTop = 95;
                        footercell.Colspan = 9;
                        PdfPCell BelowCell1 = new PdfPCell(new Phrase(" Signature Supervisor", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell1.Border = PdfPCell.NO_BORDER;
                        BelowCell1.Colspan = 9;
                        BelowCell1.PaddingTop = 95;
                        BelowCell1.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        PdfPCell BelowCell2 = new PdfPCell(new Phrase(" (Name & Signature)\r\nTruck Driver", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell2.Border = PdfPCell.NO_BORDER;
                        BelowCell2.PaddingTop = 35;
                        BelowCell2.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        BelowCell2.Colspan = 18;
                        leftTable.AddCell(BelowCell);
                        leftTable.AddCell(footercell);
                        leftTable.AddCell(BelowCell1);
                        leftTable.AddCell(BelowCell2);




                        PdfContentByte canvas = writer.DirectContent;
                        PdfTemplate template = canvas.CreateTemplate(595, 842);
                        ColumnText CenterColumnText = new ColumnText(template);
                        CenterColumnText.SetSimpleColumn(new Rectangle(0, 0, 600, 780));
                        CenterColumnText.AddElement(leftTable);
                        CenterColumnText.Go();
                        canvas.AddTemplate(template, 0, 0);
                        document.Close();
                        fileContents = memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return fileContents;
        }
        public static byte[] SpriteFL6(FL6ReportResponseModel model, string logoPath)
        {
            byte[] fileContents;
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Document document = new Document(new Rectangle(595, 842), 0f, 0f, 0f, 0f))
                    {
                        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                        document.Open();
                        #region Added image Logo
                        var imagePath = logoPath;
                        Image myImage = Image.GetInstance(imagePath);
                        float fixedWidth = logoPath.ToLower().Contains("excise") ? 150 : 380;
                        float fixedHeight = 150;
                        float widthScalingFactor = fixedWidth / myImage.Width;
                        float heightScalingFactor = fixedHeight / myImage.Height;
                        float scalingFactor = Math.Min(widthScalingFactor, heightScalingFactor);
                        myImage.ScaleAbsolute(myImage.Width * scalingFactor, myImage.Height * scalingFactor);
                        float xCentered = (document.PageSize.Width - myImage.ScaledWidth) / 2;
                        myImage.SetAbsolutePosition(xCentered, document.PageSize.Height - myImage.ScaledHeight - document.TopMargin);
                        document.Add(myImage);
                        #endregion
                        PdfContentByte canvas = writer.DirectContent;
                        PdfTemplate template = canvas.CreateTemplate(595, 842);
                        Font blueFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.NORMAL, BaseColor.BLUE);
                        Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
                        PdfPTable dataTable = new PdfPTable(12);
                        PdfPCell BelowCell = new PdfPCell(new Phrase("TP No : " + model.FL6No, FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell.Border = PdfPCell.NO_BORDER;
                        BelowCell.Colspan = 6;
                        PdfPCell rightBelowCell = new PdfPCell(new Phrase("Date of TP : " + model.FL6Date, FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLUE)));
                        rightBelowCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        rightBelowCell.Border = PdfPCell.NO_BORDER;
                        rightBelowCell.Colspan = 6;
                        PdfPCell leftCell1 = new PdfPCell(new Phrase("Distillery/Godown Name", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell1.Border = PdfPCell.NO_BORDER;
                        leftCell1.Colspan = 6;
                        leftCell1.PaddingTop = 40f;
                        PdfPCell centerCell1 = new PdfPCell(new Phrase(": " + model.consignerName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell1.Border = PdfPCell.NO_BORDER;
                        centerCell1.Colspan = 6;
                        centerCell1.PaddingTop = 40f;
                        PdfPCell leftCell2 = new PdfPCell(new Phrase("Address", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell2.Border = PdfPCell.NO_BORDER;
                        leftCell2.Colspan = 6;
                        leftCell2.PaddingTop = 20f;
                        PdfPCell centerCell2 = new PdfPCell(new Phrase(": " + model.consignerAddress, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell2.Border = PdfPCell.NO_BORDER;
                        centerCell2.Colspan = 6;
                        centerCell2.PaddingTop = 20f;
                        PdfPCell leftCell3 = new PdfPCell(new Phrase("Consignee Name", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell3.Border = PdfPCell.NO_BORDER;
                        leftCell3.Colspan = 6;
                        leftCell3.PaddingTop = 20f;
                        PdfPCell centerCell3 = new PdfPCell(new Phrase(": " + model.consigneeName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell3.Border = PdfPCell.NO_BORDER;
                        centerCell3.Colspan = 6;
                        centerCell3.PaddingTop = 20f;
                        PdfPCell leftCell4 = new PdfPCell(new Phrase("Address", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell4.Border = PdfPCell.NO_BORDER;
                        leftCell4.Colspan = 6;
                        leftCell4.PaddingTop = 20f;
                        PdfPCell centerCell4 = new PdfPCell(new Phrase(": " + model.consigneeAddress, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell4.Border = PdfPCell.NO_BORDER;
                        centerCell4.Colspan = 6;
                        centerCell4.PaddingTop = 20f;
                        PdfPCell leftCell5 = new PdfPCell(new Phrase("FL5 No", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell5.Border = PdfPCell.NO_BORDER;
                        leftCell5.Colspan = 6;
                        leftCell5.PaddingTop = 20f;
                        PdfPCell centerCell5 = new PdfPCell(new Phrase(": " + model.FL5NO, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell5.Border = PdfPCell.NO_BORDER;
                        centerCell5.Colspan = 6;
                        centerCell5.PaddingTop = 20f;
                        PdfPCell leftCell6 = new PdfPCell(new Phrase("Validity Date", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell6.Border = PdfPCell.NO_BORDER;
                        leftCell6.Colspan = 3;
                        leftCell6.PaddingTop = 20f;
                        PdfPCell centerCell6 = new PdfPCell(new Phrase(": " + model.Fl6ValidityDate, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell6.Border = PdfPCell.NO_BORDER;
                        centerCell6.Colspan = 3;
                        centerCell6.PaddingTop = 20f;
                        PdfPCell centerCell61 = new PdfPCell(new Phrase("Days:", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell61.Border = PdfPCell.NO_BORDER;
                        centerCell61.Colspan = 3;
                        centerCell61.PaddingTop = 20f;
                        PdfPCell centerCell62 = new PdfPCell(new Phrase(model.days, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell62.Border = PdfPCell.NO_BORDER;
                        centerCell62.Colspan = 3;
                        centerCell62.PaddingTop = 20f;
                        PdfPCell leftCell7 = new PdfPCell(new Phrase("Vehicle No", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell7.Border = PdfPCell.NO_BORDER;
                        leftCell7.Colspan = 6;
                        leftCell7.PaddingTop = 20f;
                        PdfPCell centerCell7 = new PdfPCell(new Phrase(": " + model.vehicleNo, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell7.Border = PdfPCell.NO_BORDER;
                        centerCell7.Colspan = 6;
                        centerCell7.PaddingTop = 20f;
                        PdfPCell leftCell8 = new PdfPCell(new Phrase("Person Carrying Consigment ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell8.Border = PdfPCell.NO_BORDER;
                        leftCell8.Colspan = 6;
                        leftCell8.PaddingTop = 20f;
                        PdfPCell centerCell8 = new PdfPCell(new Phrase(": " + model.personCarrying, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell8.Border = PdfPCell.NO_BORDER;
                        centerCell8.Colspan = 6;
                        centerCell8.PaddingTop = 20f;
                        PdfPCell leftCell9 = new PdfPCell(new Phrase("Route ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell9.Border = PdfPCell.NO_BORDER;
                        leftCell9.Colspan = 6;
                        leftCell9.PaddingTop = 20f;
                        PdfPCell centerCell9 = new PdfPCell(new Phrase(": " + model.routeName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell9.Border = PdfPCell.NO_BORDER;
                        centerCell9.Colspan = 6;
                        centerCell9.PaddingTop = 20f;
                        PdfPCell leftCell10 = new PdfPCell(new Phrase("Product Detail ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell10.Border = PdfPCell.NO_BORDER;
                        leftCell10.Colspan = 6;
                        leftCell10.PaddingTop = 20f;
                        PdfPCell centerCell10 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell10.Border = PdfPCell.NO_BORDER;
                        centerCell10.Colspan = 6;
                        centerCell10.PaddingTop = 20f;
                        PdfPCell datacell1 = new PdfPCell(new Phrase("SI NO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell1.Rowspan = 2;
                        PdfPCell datacell2 = new PdfPCell(new Phrase("Product ", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell2.Rowspan = 2;
                        PdfPCell datacell3 = new PdfPCell(new Phrase("Brand", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell3.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        datacell3.Colspan = 3;
                        datacell3.Rowspan = 2;
                        PdfPCell datacell4 = new PdfPCell(new Phrase("Packing", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell4.Rowspan = 2;
                        PdfPCell datacell5 = new PdfPCell(new Phrase("Batch No", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell5.Rowspan = 2;
                        PdfPCell datacell51 = new PdfPCell(new Phrase("TP", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell51.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        datacell51.Colspan = 3;
                        PdfPCell datacell61 = new PdfPCell(new Phrase("Westage", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell61.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        datacell61.Colspan = 2;
                        PdfPCell datacell6 = new PdfPCell(new Phrase("Qty(BL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell7 = new PdfPCell(new Phrase("Strength", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell8 = new PdfPCell(new Phrase("Qty(LPL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell9 = new PdfPCell(new Phrase("Qty(BL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell10 = new PdfPCell(new Phrase("Qty(LPL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell12 = new PdfPCell(new Phrase("1", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell22 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.productName ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell32 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.brandName ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        datacell32.Colspan = 3;
                        PdfPCell datacell42 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.packing ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell52 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.batchNo ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell62 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.BLQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell72 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.strength ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell82 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.LPLQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell92 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.wastageBLQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell datacell93 = new PdfPCell(new Phrase(model.lstProduct?.FirstOrDefault()?.wastageLPLQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA,
                            8, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell leftCell11 = new PdfPCell(new Phrase("CC to : ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell11.Border = PdfPCell.NO_BORDER;
                        leftCell11.Colspan = 6;
                        leftCell11.PaddingTop = 20f;
                        PdfPCell centerCell11 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell11.Border = PdfPCell.NO_BORDER;
                        centerCell11.Colspan = 6;
                        PdfPCell leftCell12 = new PdfPCell(new Phrase("DEO(Consignee) ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell12.Border = PdfPCell.NO_BORDER;
                        leftCell12.Colspan = 6;
                        leftCell12.PaddingTop = 20f;
                        PdfPCell centerCell12 = new PdfPCell(new Phrase(": " + model.deoConsignee ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell12.Border = PdfPCell.NO_BORDER;
                        centerCell12.Colspan = 6;
                        centerCell12.PaddingTop = 20f;
                        PdfPCell leftCell13 = new PdfPCell(new Phrase("DEO(Consigner)", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell13.Border = PdfPCell.NO_BORDER;
                        leftCell13.Colspan = 6;
                        leftCell13.PaddingTop = 20f;
                        PdfPCell centerCell13 = new PdfPCell(new Phrase(": " + model.deoConsigner ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        centerCell13.Border = PdfPCell.NO_BORDER;
                        centerCell13.Colspan = 6;
                        centerCell13.PaddingTop = 20f;
                        PdfPCell leftCell14 = new PdfPCell(new Phrase("It is verified that following liquor receipt is accepted for this consignment, and it is entered on",
                            FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        leftCell14.Border = PdfPCell.NO_BORDER;
                        leftCell14.Colspan = 12;
                        leftCell14.PaddingTop = 20f;
                        PdfPCell datacell11 = new PdfPCell(new Phrase("SI NO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell21 = new PdfPCell(new Phrase("Product ", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell31 = new PdfPCell(new Phrase("Brand", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        datacell31.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        datacell31.Colspan = 2;
                        PdfPCell datacell41 = new PdfPCell(new Phrase("Packing", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell5111 = new PdfPCell(new Phrase("Batch No", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell511 = new PdfPCell(new Phrase("Qty(Case)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell611 = new PdfPCell(new Phrase("Qty(BL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell6111 = new PdfPCell(new Phrase("Strength", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell71 = new PdfPCell(new Phrase("Qty(LPL)", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell81 = new PdfPCell(new Phrase("Duty Rate", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell datacell91 = new PdfPCell(new Phrase("Duty Paid", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dataTable.AddCell(BelowCell);
                        dataTable.AddCell(rightBelowCell);
                        dataTable.AddCell(leftCell1);
                        dataTable.AddCell(centerCell1);
                        dataTable.AddCell(leftCell2);
                        dataTable.AddCell(centerCell2);
                        dataTable.AddCell(leftCell3);
                        dataTable.AddCell(centerCell3);
                        dataTable.AddCell(leftCell4);
                        dataTable.AddCell(centerCell4);
                        dataTable.AddCell(leftCell5);
                        dataTable.AddCell(centerCell5);
                        dataTable.AddCell(leftCell6);
                        dataTable.AddCell(centerCell6);
                        dataTable.AddCell(centerCell61);
                        dataTable.AddCell(centerCell62);
                        dataTable.AddCell(leftCell7);
                        dataTable.AddCell(centerCell7);
                        dataTable.AddCell(leftCell8);
                        dataTable.AddCell(centerCell8);
                        dataTable.AddCell(leftCell9);
                        dataTable.AddCell(centerCell9);
                        dataTable.AddCell(leftCell10);
                        dataTable.AddCell(centerCell10);
                        dataTable.AddCell(datacell11);
                        dataTable.AddCell(datacell21);
                        dataTable.AddCell(datacell31);
                        dataTable.AddCell(datacell41);
                        dataTable.AddCell(datacell5111);
                        dataTable.AddCell(datacell511);
                        dataTable.AddCell(datacell611);
                        dataTable.AddCell(datacell6111);
                        dataTable.AddCell(datacell71);
                        dataTable.AddCell(datacell81);
                        dataTable.AddCell(datacell91);
                        int i = 0;
                        foreach (var item in model.lstProduct)
                        {
                            i++;
                            PdfPCell datacell121 = new PdfPCell(new Phrase(Convert.ToString(i), FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell221 = new PdfPCell(new Phrase(item?.productName ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell321 = new PdfPCell(new Phrase(item?.brandName ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            datacell321.Colspan = 2;
                            PdfPCell datacell421 = new PdfPCell(new Phrase(item?.packing ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell521 = new PdfPCell(new Phrase(item?.batchNo ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell621 = new PdfPCell(new Phrase(item?.caseQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell721 = new PdfPCell(new Phrase(item?.BLQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell821 = new PdfPCell(new Phrase(item?.strength ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell921 = new PdfPCell(new Phrase(item?.LPLQty ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell931 = new PdfPCell(new Phrase(item?.dutyAmount ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            PdfPCell datacell941 = new PdfPCell(new Phrase(item?.dutyRate ?? "", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLACK)));
                            dataTable.AddCell(datacell121);
                            dataTable.AddCell(datacell221);
                            dataTable.AddCell(datacell321);
                            dataTable.AddCell(datacell421);
                            dataTable.AddCell(datacell521);
                            dataTable.AddCell(datacell621);
                            dataTable.AddCell(datacell721);
                            dataTable.AddCell(datacell821);
                            dataTable.AddCell(datacell921);
                            dataTable.AddCell(datacell931);
                            dataTable.AddCell(datacell941);
                        }
                        PdfPCell totalCell = new PdfPCell(new Phrase("Total", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        totalCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        totalCell.Colspan = 6;
                        PdfPCell totalCell1 = new PdfPCell(new Phrase(Convert.ToString(model.lstProduct.Sum(s => Convert.ToInt64(s.caseQty))),
                            FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell totalCell2 = new PdfPCell(new Phrase(Convert.ToString(model.lstProduct.Sum(s => Convert.ToInt64(s.BLQty))),
                            FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell totalCell3 = new PdfPCell(new Phrase(" ", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell totalCell4 = new PdfPCell(new Phrase(Convert.ToString(model.lstProduct.Sum(s => Convert.ToInt64(s.LPLQty))),
                            FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell totalCell5 = null;
                        if (model.lstProduct.FirstOrDefault().dutyRate.Contains("null"))
                        {
                            totalCell5 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        }
                        else
                        {
                            totalCell5 = new PdfPCell(new Phrase(Convert.ToString(model.lstProduct.Sum(s => Convert.ToInt64(s.dutyRate))),
                                FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        }
                        PdfPCell totalCell6 = null;
                        if (model.lstProduct.FirstOrDefault().dutyAmount.Contains("null"))
                        {
                            totalCell6 = new PdfPCell(new Phrase("", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        }
                        else
                        {
                            totalCell6 = new PdfPCell(new Phrase(Convert.ToString(model.lstProduct.Sum(s => Convert.ToInt64(s.dutyAmount))),
                                FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        }
                        dataTable.AddCell(totalCell);
                        dataTable.AddCell(totalCell1);
                        dataTable.AddCell(totalCell2);
                        dataTable.AddCell(totalCell3);
                        dataTable.AddCell(totalCell4);
                        dataTable.AddCell(totalCell5);
                        dataTable.AddCell(totalCell6);
                        PdfPCell BelowCell1 = new PdfPCell(new Phrase("Authorised Signature"));
                        BelowCell1.Border = PdfPCell.NO_BORDER;
                        BelowCell1.Colspan = 6;
                        BelowCell1.PaddingTop = 40;
                        PdfPCell BelowCell2 = new PdfPCell(new Phrase("Authorised Signature"));
                        BelowCell2.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        BelowCell2.Border = PdfPCell.NO_BORDER;
                        BelowCell2.Colspan = 6;
                        BelowCell2.PaddingTop = 40;
                        dataTable.AddCell(leftCell11);
                        dataTable.AddCell(centerCell11);
                        dataTable.AddCell(leftCell12);
                        dataTable.AddCell(centerCell12);
                        dataTable.AddCell(leftCell13);
                        dataTable.AddCell(centerCell13);
                        dataTable.AddCell(leftCell14);
                        dataTable.AddCell(datacell1);
                        dataTable.AddCell(datacell2);
                        dataTable.AddCell(datacell3);
                        dataTable.AddCell(datacell4);
                        dataTable.AddCell(datacell5);
                        dataTable.AddCell(datacell51);
                        dataTable.AddCell(datacell61);
                        dataTable.AddCell(datacell6);
                        dataTable.AddCell(datacell7);
                        dataTable.AddCell(datacell8);
                        dataTable.AddCell(datacell9);
                        dataTable.AddCell(datacell10);
                        dataTable.AddCell(datacell12);
                        dataTable.AddCell(datacell22);
                        dataTable.AddCell(datacell32);
                        dataTable.AddCell(datacell42);
                        dataTable.AddCell(datacell52);
                        dataTable.AddCell(datacell62);
                        dataTable.AddCell(datacell72);
                        dataTable.AddCell(datacell82);
                        dataTable.AddCell(datacell92);
                        dataTable.AddCell(datacell93);
                        dataTable.AddCell(BelowCell1);
                        dataTable.AddCell(BelowCell2);
                        ColumnText dataTableColumnText = new ColumnText(template);
                        dataTableColumnText.SetSimpleColumn(new Rectangle(0, 0, 600, 780));
                        dataTableColumnText.AddElement(dataTable);
                        dataTableColumnText.Go();
                        canvas.AddTemplate(template, 0, 0);
                        document.Close();
                        fileContents = memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return fileContents;
        }
        public static Image GenerateDyanamicQRCode(string Input1, string Input2 = "", string Input3 = "")
        {
            BarcodeQRCode qrCode = new BarcodeQRCode(Input1 + Input2 + Input3, 80, 80, null);
            Image qrCodeImage = qrCode.GetImage();
            qrCodeImage.SetAbsolutePosition(495, 770); // Adjust position as needed
            return qrCodeImage;
        }
        public static byte[] TOS(TOOReportResponseModel model, string logoPath)
        {
            byte[] fileContents;
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Document document = new Document(new Rectangle(595, 842), 0f, 0f, 0f, 0f))
                    {
                        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                        document.Open();
                        #region Added image Logo
                        var imagePath = logoPath;
                        Image myImage = Image.GetInstance(imagePath);
                        float fixedWidth = logoPath.ToLower().Contains("excise") ? 150 : 380;
                        float fixedHeight = 150;
                        float widthScalingFactor = fixedWidth / myImage.Width;
                        float heightScalingFactor = fixedHeight / myImage.Height;
                        float scalingFactor = Math.Min(widthScalingFactor, heightScalingFactor);
                        myImage.ScaleAbsolute(myImage.Width * scalingFactor, myImage.Height * scalingFactor);
                        float xCentered = (document.PageSize.Width - myImage.ScaledWidth) / 2;
                        myImage.SetAbsolutePosition(xCentered, document.PageSize.Height - myImage.ScaledHeight - document.TopMargin);
                        document.Add(myImage);
                        #endregion
                        Image qrCodeImage = Reports.GenerateDyanamicQRCode(model.listReport.FirstOrDefault().TOONo,
                            model.listReport.FirstOrDefault().consigneeId, "TOS");
                        document.Add(qrCodeImage);
                        Font blueFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLUE);
                        Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
                        Font smallFont = FontFactory.GetFont(FontFactory.HELVETICA, 6, Font.BOLD, BaseColor.BLACK);
                        Font smallFont1 = FontFactory.GetFont(FontFactory.HELVETICA, 6, Font.NORMAL, BaseColor.BLACK);
                        PdfPTable leftTable = new PdfPTable(15);
                        PdfPCell leftCell = new PdfPCell(new Phrase("TELE:" + model.listReport.FirstOrDefault().teleNo, blueFont));
                        leftCell.Border = PdfPCell.NO_BORDER;
                        leftCell.Colspan = 15;
                        PdfPCell leftCell1 = new PdfPCell(new Phrase("FAX: " + model.listReport.FirstOrDefault().faxNo, blueFont));
                        leftCell1.Border = PdfPCell.NO_BORDER;
                        leftCell1.Colspan = 15;
                        PdfPCell HeaderCell = new PdfPCell(new Phrase("TRANSFER OUT SLIP", FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.BOLD, BaseColor.BLUE)));
                        HeaderCell.Border = PdfPCell.NO_BORDER;
                        HeaderCell.Colspan = 15;
                        HeaderCell.HorizontalAlignment = PdfPCell.ALIGN_CENTER;
                        PdfPCell HeaderRightCell = new PdfPCell(new Phrase(model.listReport.FirstOrDefault().OrgName + " TIN :" + model.listReport.FirstOrDefault().TinNo,
                            FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        HeaderRightCell.Border = PdfPCell.NO_BORDER;
                        HeaderRightCell.Colspan = 15;
                        HeaderRightCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        PdfPCell HeaderRightCell1 = new PdfPCell(new Phrase("CIN :" + model.listReport.FirstOrDefault().CIN, FontFactory.GetFont(FontFactory.HELVETICA,
                            9, Font.BOLD, BaseColor.BLACK)));
                        HeaderRightCell1.Border = PdfPCell.NO_BORDER;
                        HeaderRightCell1.Colspan = 15;
                        HeaderRightCell1.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        Phrase p = new Phrase("Transfer In Depot Name:\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK));
                        Phrase p1 = new Phrase(model.listReport.FirstOrDefault().consigneeAddress, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK));

                        PdfPCell LeftCell = new PdfPCell();
                        LeftCell.AddElement(p);
                        LeftCell.AddElement(p1);
                        LeftCell.Border = PdfPCell.NO_BORDER;
                        LeftCell.Rowspan = 5;
                        LeftCell.Colspan = 5;

                        PdfPCell RightCell = new PdfPCell(new Phrase("Tos NO", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        RightCell.Colspan = 7;
                        RightCell.Border = PdfPCell.NO_BORDER;
                        PdfPCell RightCell1 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TOSNo, FontFactory.GetFont(FontFactory.HELVETICA,
                            10, Font.NORMAL, BaseColor.BLACK)));
                        RightCell1.Colspan = 3;
                        RightCell1.Border = PdfPCell.NO_BORDER;

                        PdfPCell RightCell2 = new PdfPCell(new Phrase("Tos Date", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        RightCell2.Colspan = 7;
                        RightCell2.Border = PdfPCell.NO_BORDER;
                        PdfPCell RightCell3 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TOSDate, FontFactory.GetFont(FontFactory.HELVETICA,
                            10, Font.NORMAL, BaseColor.BLACK)));
                        RightCell3.Colspan = 3;
                        RightCell3.Border = PdfPCell.NO_BORDER;
                        PdfPCell RightCell4 = new PdfPCell(new Phrase("TOO Number", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        RightCell4.Colspan = 7;

                        RightCell4.Border = PdfPCell.NO_BORDER;
                        PdfPCell RightCell5 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TOONo, FontFactory.GetFont(FontFactory.HELVETICA,
                            10, Font.NORMAL, BaseColor.BLACK)));
                        RightCell5.Colspan = 3;
                        RightCell5.Border = PdfPCell.NO_BORDER;
                        PdfPCell RightCell51 = new PdfPCell(new Phrase("Dispatch Through", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        RightCell51.Colspan = 7;
                        RightCell51.Border = PdfPCell.NO_BORDER;
                        //PdfPCell RightCell52 = new PdfPCell(new Phrase(": Hard Coded"/* + model.listReport.FirstOrDefault().transferType*/,
                        //    FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell RightCell52 = new PdfPCell(new Phrase(": Hard Coded", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        RightCell52.Colspan = 3;
                        RightCell52.Border = PdfPCell.NO_BORDER;
                        PdfPCell RightCell53 = new PdfPCell(new Phrase("Terms of Delivery As Per\r\n Agrement Validity",
                            FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK)));
                        RightCell53.Colspan = 7;
                        RightCell53.Border = PdfPCell.NO_BORDER;
                        //PdfPCell RightCell54 = new PdfPCell(new Phrase(": Hard Coded" /*+ model.listReport.FirstOrDefault().TPDate*/, 
                        //    FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell RightCell54 = new PdfPCell(new Phrase(": Hard Coded", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        RightCell54.Colspan = 3;
                        RightCell54.Border = PdfPCell.NO_BORDER;
                        p = new Phrase("Transfer Out Depot Name :\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK));
                        p1 = new Phrase(model.listReport.FirstOrDefault().consignerAddress, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK));
                        PdfPCell LeftCell1 = new PdfPCell();
                        LeftCell1.AddElement(p);
                        LeftCell1.AddElement(p1);
                        LeftCell1.Border = PdfPCell.NO_BORDER;
                        LeftCell1.Rowspan = 5;
                        LeftCell1.Colspan = 5;
                        p = new Phrase("Supplier Name :\r\n", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK));
                        p1 = new Phrase("Hard Coded", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK));

                        PdfPCell LeftCell2 = new PdfPCell();
                        LeftCell2.AddElement(p);
                        LeftCell2.AddElement(p1);

                        LeftCell2.Border = PdfPCell.NO_BORDER;
                        LeftCell2.Rowspan = 5;
                        LeftCell2.Colspan = 7;
                        //PdfPCell LeftCell3 = new PdfPCell(new Phrase(" : Hard Coded" /*+ model.listReport.FirstOrDefault().consigneeAddress*/,
                        //    FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        PdfPCell LeftCell3 = new PdfPCell(new Phrase(" : Hard Coded",
                            FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK)));
                        LeftCell3.Border = PdfPCell.NO_BORDER;
                        LeftCell3.Rowspan = 5;
                        LeftCell3.Colspan = 3;
                        p = new Phrase("Vehicle Number :  ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK));
                        p1 = new Phrase(model.listReport.FirstOrDefault().TPVehicleNo, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK));

                        PdfPCell LeftCell4 = new PdfPCell();
                        LeftCell4.AddElement(p);
                        LeftCell4.AddElement(p1);
                        LeftCell4.Border = PdfPCell.NO_BORDER;
                        LeftCell4.Colspan = 5;

                        p = new Phrase("Driver Name :  ", FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLACK));
                        p1 = new Phrase(model.listReport.FirstOrDefault().TPDriverName, FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.NORMAL, BaseColor.BLACK));

                        PdfPCell LeftCell5 = new PdfPCell();
                        LeftCell5.AddElement(p);
                        LeftCell5.AddElement(p1);
                        LeftCell5.Border = PdfPCell.NO_BORDER;
                        LeftCell5.Colspan = 10;


                        leftTable.AddCell(leftCell);
                        leftTable.AddCell(leftCell1);
                        leftTable.AddCell(HeaderCell);
                        leftTable.AddCell(HeaderRightCell);
                        leftTable.AddCell(HeaderRightCell1);
                        leftTable.AddCell(LeftCell);
                        leftTable.AddCell(RightCell);
                        leftTable.AddCell(RightCell1);

                        leftTable.AddCell(RightCell2);
                        leftTable.AddCell(RightCell3);
                        leftTable.AddCell(RightCell4);
                        leftTable.AddCell(RightCell5);
                        leftTable.AddCell(RightCell51);
                        leftTable.AddCell(RightCell52);
                        leftTable.AddCell(RightCell53);
                        leftTable.AddCell(RightCell54);
                        leftTable.AddCell(LeftCell1);
                        leftTable.AddCell(LeftCell2);
                        leftTable.AddCell(LeftCell3);
                        leftTable.AddCell(LeftCell4);
                        leftTable.AddCell(LeftCell5);

                        PdfPCell spaceCell = new PdfPCell();
                        spaceCell.FixedHeight = 20f;
                        spaceCell.Colspan = 15;
                        spaceCell.Border = PdfPCell.NO_BORDER;
                        leftTable.AddCell(spaceCell);

                        PdfPCell dtcell1 = new PdfPCell(new Phrase("SNO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));

                        PdfPCell dtcell2 = new PdfPCell(new Phrase("Product Category", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dtcell2.Colspan = 3;
                        PdfPCell dtcell3 = new PdfPCell(new Phrase("Product", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dtcell3.Colspan = 3;
                        PdfPCell dtcell4 = new PdfPCell(new Phrase("Packing ", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dtcell4.Colspan = 3;
                        PdfPCell dtcell5 = new PdfPCell(new Phrase("Case Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dtcell6 = new PdfPCell(new Phrase("Bottle Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dtcell7 = new PdfPCell(new Phrase("BL  Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dtcell8 = new PdfPCell(new Phrase("LPL  Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));

                        PdfPCell dtcell9 = new PdfPCell(new Phrase("Purchase Value", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        leftTable.AddCell(dtcell1);
                        leftTable.AddCell(dtcell2);
                        leftTable.AddCell(dtcell3);
                        leftTable.AddCell(dtcell4);
                        leftTable.AddCell(dtcell5);
                        leftTable.AddCell(dtcell6);
                        leftTable.AddCell(dtcell7);
                        leftTable.AddCell(dtcell8);
                        leftTable.AddCell(dtcell9);

                        int i = 0;
                        foreach (var item in model.listProduct)
                        {
                            i = i + 1;


                            dtcell1 = new PdfPCell(new Phrase(i.ToString(), smallFont1));

                            dtcell2 = new PdfPCell(new Phrase(item.productName, smallFont1));
                            dtcell2.Colspan = 3;
                            dtcell3 = new PdfPCell(new Phrase(item.brandName, smallFont1));
                            dtcell3.Colspan = 3;
                            dtcell4 = new PdfPCell(new Phrase(item.packingName, smallFont1));
                            dtcell4.Colspan = 3;
                            dtcell5 = new PdfPCell(new Phrase(item.caseQty, smallFont1));

                            dtcell6 = new PdfPCell(new Phrase(Convert.ToString(item.bottleQty), smallFont1));

                            dtcell7 = new PdfPCell(new Phrase(Convert.ToString(item.BLQty), smallFont1));
                            dtcell8 = new PdfPCell(new Phrase(Convert.ToString(item.LPLQty), smallFont1));

                            dtcell9 = new PdfPCell(new Phrase(item.amount, smallFont1));


                            leftTable.AddCell(dtcell1);
                            leftTable.AddCell(dtcell2);
                            leftTable.AddCell(dtcell3);
                            leftTable.AddCell(dtcell4);
                            leftTable.AddCell(dtcell5);
                            leftTable.AddCell(dtcell6);
                            leftTable.AddCell(dtcell7);
                            leftTable.AddCell(dtcell8);
                            leftTable.AddCell(dtcell9);
                        }

                        PdfPCell BelowCell = new PdfPCell(new Phrase("For " + model.listReport.FirstOrDefault().OrgName,
                            FontFactory.GetFont(FontFactory.HELVETICA, 10, Font.BOLD, BaseColor.BLUE)));
                        BelowCell.Border = PdfPCell.NO_BORDER;
                        BelowCell.Colspan = 15;
                        BelowCell.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;

                        PdfPCell BelowCell3 = new PdfPCell(new Phrase("Subject to the sourcing and sales policy of the Corporation Issured from time to time",
                            FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell3.Border = PdfPCell.NO_BORDER;
                        BelowCell3.Colspan = 15;

                        PdfPCell BelowCell4 = new PdfPCell(new Phrase("ALL SUBJECT TO JAIPUR JURISDICTION ONLY.\r\n E. & O.E.",
                            FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell4.Border = PdfPCell.NO_BORDER;
                        BelowCell4.Colspan = 15;
                        PdfPCell BelowCell2 = new PdfPCell(new Phrase(model.listReport.FirstOrDefault().DepoIncharge, FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.NORMAL, BaseColor.BLUE)));
                        BelowCell2.Border = PdfPCell.NO_BORDER;
                        BelowCell2.Colspan = 15;
                        BelowCell2.HorizontalAlignment = PdfPCell.ALIGN_RIGHT;
                        leftTable.AddCell(BelowCell);
                        leftTable.AddCell(BelowCell2);
                        leftTable.AddCell(BelowCell3);
                        leftTable.AddCell(BelowCell4);
                        PdfContentByte canvas = writer.DirectContent;
                        PdfTemplate template = canvas.CreateTemplate(595, 842);
                        ColumnText leftColumnText = new ColumnText(template);
                        leftColumnText.SetSimpleColumn(new Rectangle(0, 0, 600, 780));
                        leftColumnText.AddElement(leftTable);
                        leftColumnText.Go();
                        canvas.AddTemplate(template, 0, 0);
                        document.Close();
                        fileContents = memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return fileContents;
        }
        public static byte[] TOOTP(TOOReportResponseModel model, string logoPath)
        {
            byte[] fileContents;
            try
            {
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    using (Document document = new Document(new Rectangle(595, 842), 0f, 0f, 0f, 0f))
                    {
                        PdfWriter writer = PdfWriter.GetInstance(document, memoryStream);
                        document.Open();
                        #region Added image Logo
                        var imagePath = logoPath;
                        Image myImage = Image.GetInstance(imagePath);
                        float fixedWidth = logoPath.ToLower().Contains("excise") ? 150 : 380;
                        float fixedHeight = 150;
                        float widthScalingFactor = fixedWidth / myImage.Width;
                        float heightScalingFactor = fixedHeight / myImage.Height;
                        float scalingFactor = Math.Min(widthScalingFactor, heightScalingFactor);
                        myImage.ScaleAbsolute(myImage.Width * scalingFactor, myImage.Height * scalingFactor);
                        float xCentered = (document.PageSize.Width - myImage.ScaledWidth) / 2;
                        myImage.SetAbsolutePosition(xCentered, document.PageSize.Height - myImage.ScaledHeight - document.TopMargin);
                        document.Add(myImage);
                        #endregion
                        Font blueFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLUE);
                        Font boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 12, Font.BOLD, BaseColor.BLACK);
                        Font smallFont = FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK);

                        PdfPTable leftTable = new PdfPTable(1);
                        PdfPCell leftCell = new PdfPCell(new Phrase("TP No     :" + model.listReport.FirstOrDefault().TPNo, blueFont));
                        leftCell.Border = PdfPCell.NO_BORDER;
                        leftTable.AddCell(leftCell);

                        PdfContentByte canvas = writer.DirectContent;
                        PdfTemplate template = canvas.CreateTemplate(595, 842);
                        ColumnText leftColumnText = new ColumnText(template);
                        leftColumnText.SetSimpleColumn(new Rectangle(30, 30, 400, 800));
                        leftColumnText.AddElement(leftTable);
                        leftColumnText.Go();

                        PdfPTable pdfPtable1 = new PdfPTable(12);
                        pdfPtable1.SpacingBefore = 5f;
                        pdfPtable1.SpacingAfter = 5f;
                        PdfPCell pdfPCell1 = new PdfPCell(new Phrase("Consigner Name", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;


                        PdfPCell pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().consignerUnitName, FontFactory.GetFont(FontFactory.HELVETICA,
                            9, Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable1.AddCell(pdfPCell1);
                        pdfPtable1.AddCell(pdfPCell2);

                        PdfPTable pdfPtable2 = new PdfPTable(12);
                        pdfPtable2.SpacingBefore = 5f;
                        pdfPtable2.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Addresss From ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().consignerAddress, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable2.AddCell(pdfPCell1);
                        pdfPtable2.AddCell(pdfPCell2);

                        PdfPTable pdfPtable3 = new PdfPTable(12);
                        pdfPtable3.SpacingBefore = 5f;
                        pdfPtable3.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Consignee Name ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().consigneeUnitName, FontFactory.GetFont(FontFactory.HELVETICA,
                            9, Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable3.AddCell(pdfPCell1);
                        pdfPtable3.AddCell(pdfPCell2);

                        PdfPTable pdfPtable4 = new PdfPTable(12);
                        pdfPtable4.SpacingBefore = 5f;
                        pdfPtable4.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Addresss TO ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().consigneeAddress, FontFactory.GetFont(FontFactory.HELVETICA,
                            9, Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable4.AddCell(pdfPCell1);
                        pdfPtable4.AddCell(pdfPCell2);

                        PdfPTable pdfPtable6 = new PdfPTable(12);
                        pdfPtable6.SpacingBefore = 5f;
                        pdfPtable6.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("TOO NO ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TOONo, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable6.AddCell(pdfPCell1);
                        pdfPtable6.AddCell(pdfPCell2);

                        PdfPTable pdfPtable7 = new PdfPTable(12);
                        pdfPtable7.SpacingBefore = 5f;
                        pdfPtable7.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Depo Type ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().depoType, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable7.AddCell(pdfPCell1);
                        pdfPtable7.AddCell(pdfPCell2);

                        PdfPTable pdfPtable8 = new PdfPTable(12);
                        pdfPtable8.SpacingBefore = 5f;
                        pdfPtable8.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("TransFer Type ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().transferType, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable8.AddCell(pdfPCell1);
                        pdfPtable8.AddCell(pdfPCell2);

                        PdfPTable pdfPtable9 = new PdfPTable(12);
                        pdfPtable9.SpacingBefore = 5f;
                        pdfPtable9.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("TP No ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TPNo, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable9.AddCell(pdfPCell1);
                        pdfPtable9.AddCell(pdfPCell2);

                        PdfPTable pdfPtable10 = new PdfPTable(12);
                        pdfPtable10.SpacingBefore = 5f;
                        pdfPtable10.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("TP Date ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TPDate, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable10.AddCell(pdfPCell1);
                        pdfPtable10.AddCell(pdfPCell2);

                        PdfPTable pdfPtable11 = new PdfPTable(12);
                        pdfPtable11.SpacingBefore = 5f;
                        pdfPtable11.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("TP Validity ", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TPValidityDate, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable11.AddCell(pdfPCell1);
                        pdfPtable11.AddCell(pdfPCell2);


                        PdfPTable pdfPtable12 = new PdfPTable(12);
                        pdfPtable12.SpacingBefore = 5f;
                        pdfPtable12.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Transporter Name", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TPTransporter, FontFactory.GetFont(FontFactory.HELVETICA, 9,
                            Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable12.AddCell(pdfPCell1);
                        pdfPtable12.AddCell(pdfPCell2);

                        PdfPTable pdfPtable13 = new PdfPTable(12);
                        pdfPtable13.SpacingBefore = 5f;
                        pdfPtable13.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Vehicle Number", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TPVehicleNo, FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable13.AddCell(pdfPCell1);
                        pdfPtable13.AddCell(pdfPCell2);


                        PdfPTable pdfPtable14 = new PdfPTable(12);
                        pdfPtable14.SpacingBefore = 5f;
                        pdfPtable14.SpacingAfter = 5f;
                        pdfPCell1 = new PdfPCell(new Phrase("Driver Name", FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.BOLD, BaseColor.BLACK)));
                        pdfPCell1.Colspan = 6;
                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPCell2 = new PdfPCell(new Phrase(":" + model.listReport.FirstOrDefault().TPDriverName, FontFactory.GetFont(FontFactory.HELVETICA, 9, Font.NORMAL, BaseColor.BLACK)));
                        pdfPCell2.Colspan = 6;
                        pdfPCell2.Border = PdfPCell.NO_BORDER;
                        pdfPtable14.AddCell(pdfPCell1);
                        pdfPtable14.AddCell(pdfPCell2);

                        PdfPTable pdfPtable5 = new PdfPTable(12);

                        pdfPCell1 = new PdfPCell(new Phrase("Product Detail"));
                        pdfPCell1.Colspan = 12;

                        ColumnText leftColumnText1 = new ColumnText(template);
                        leftColumnText1.SetSimpleColumn(new Rectangle(30, 30, 600, 780));
                        leftColumnText1.AddElement(pdfPtable1);
                        leftColumnText1.AddElement(pdfPtable2);
                        leftColumnText1.AddElement(pdfPtable3);
                        leftColumnText1.AddElement(pdfPtable4);

                        leftColumnText1.AddElement(pdfPtable6);
                        leftColumnText1.AddElement(pdfPtable7);
                        leftColumnText1.AddElement(pdfPtable8);
                        leftColumnText1.AddElement(pdfPtable9);
                        leftColumnText1.AddElement(pdfPtable10);
                        leftColumnText1.AddElement(pdfPtable11);
                        leftColumnText1.AddElement(pdfPtable12);
                        leftColumnText1.AddElement(pdfPtable13);
                        leftColumnText1.AddElement(pdfPtable14);
                        leftColumnText1.AddElement(pdfPtable5);

                        pdfPCell1.Border = PdfPCell.NO_BORDER;
                        pdfPtable5.AddCell(pdfPCell1);

                        PdfPTable dataTable = new PdfPTable(13);
                        dataTable.SpacingBefore = 10f;


                        PdfPCell dtcell1 = new PdfPCell(new Phrase("SNO", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));

                        PdfPCell dtcell2 = new PdfPCell(new Phrase("Product", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dtcell2.Colspan = 3;
                        PdfPCell dtcell3 = new PdfPCell(new Phrase("Brand", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dtcell3.Colspan = 2;
                        PdfPCell dtcell4 = new PdfPCell(new Phrase("Packing ", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        dtcell4.Colspan = 2;
                        PdfPCell dtcell5 = new PdfPCell(new Phrase("Case Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dtcell6 = new PdfPCell(new Phrase("Bottle Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));

                        PdfPCell dtcell7 = new PdfPCell(new Phrase("BL  Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));
                        PdfPCell dtcell8 = new PdfPCell(new Phrase("LPL  Qty", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));

                        PdfPCell dtcell9 = new PdfPCell(new Phrase("Purchase Value", FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD, BaseColor.BLACK)));

                        dataTable.AddCell(dtcell1);
                        dataTable.AddCell(dtcell2);
                        dataTable.AddCell(dtcell3);
                        dataTable.AddCell(dtcell4);
                        dataTable.AddCell(dtcell5);
                        dataTable.AddCell(dtcell6);
                        dataTable.AddCell(dtcell7);
                        dataTable.AddCell(dtcell8);
                        dataTable.AddCell(dtcell9);
                        leftColumnText1.AddElement(dataTable);

                        int i = 0;
                        foreach (var item in model.listProduct)
                        {
                            i = i + 1;

                            dataTable = new PdfPTable(13);
                            dtcell1 = new PdfPCell(new Phrase(i.ToString(), smallFont));

                            dtcell2 = new PdfPCell(new Phrase(item.productName, smallFont));
                            dtcell2.Colspan = 3;
                            dtcell3 = new PdfPCell(new Phrase(item.brandName, smallFont));
                            dtcell3.Colspan = 2;
                            dtcell4 = new PdfPCell(new Phrase(item.packingName, smallFont));
                            dtcell4.Colspan = 2;
                            dtcell5 = new PdfPCell(new Phrase(item.caseQty, smallFont));

                            dtcell6 = new PdfPCell(new Phrase(Convert.ToString(item.bottleQty), smallFont));

                            dtcell7 = new PdfPCell(new Phrase(Convert.ToString(item.BLQty), smallFont));
                            dtcell8 = new PdfPCell(new Phrase(Convert.ToString(item.LPLQty), smallFont));

                            dtcell9 = new PdfPCell(new Phrase(item.amount, smallFont));


                            dataTable.AddCell(dtcell1);
                            dataTable.AddCell(dtcell2);
                            dataTable.AddCell(dtcell3);
                            dataTable.AddCell(dtcell4);
                            dataTable.AddCell(dtcell5);
                            dataTable.AddCell(dtcell6);
                            dataTable.AddCell(dtcell7);
                            dataTable.AddCell(dtcell8);
                            dataTable.AddCell(dtcell9);

                            leftColumnText1.AddElement(dataTable);

                        }

                        leftColumnText1.Go();

                        PdfPTable rightTable = new PdfPTable(1);
                        PdfPCell rightCell = new PdfPCell(new Phrase("Date of Issue    :" + model.listReport.FirstOrDefault().TPDate, blueFont));
                        rightCell.Border = PdfPCell.NO_BORDER;
                        rightTable.AddCell(rightCell);
                        ColumnText rightColumnText = new ColumnText(template);
                        rightColumnText.SetSimpleColumn(new Rectangle(700, 30, 300, 800));
                        rightColumnText.AddElement(rightTable);
                        rightColumnText.Go();

                        canvas.AddTemplate(template, 0, 0);
                        document.Close();

                        fileContents = memoryStream.ToArray();
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
            return fileContents;
        }
    }
}
