using Syncfusion.Presentation;
using Syncfusion.Drawing;
using System.IO;
using System.Linq;
using System;
using SkiaSharp;
using System.Xml.Serialization;
using RdlDOM = BoldReports.RDL.DOM;
using static System.Net.Mime.MediaTypeNames;
using BoldReports.RDL.DOM;
using Syncfusion.PresentationRenderer;

namespace PresentationToRDL
{
    internal class PresentationToRdlConverter
    {
        RDLHelper rdlHelper = null;

        /// <summary>
        /// Convert PowerPoint Presentation to RDL DOM.
        /// </summary>
        /// <param name="presentation">Presentation instance to convert RDL DOM</param>
        internal MemoryStream Convert(Syncfusion.Presentation.Presentation presentation, string pptFileName, string dataProvider, string connectionString, string csvFilePath, bool isDBPresent)
        {
            MemoryStream memoryStream = new MemoryStream();
          
            try
            {
                rdlHelper = new RDLHelper(pptFileName, connectionString, dataProvider, csvFilePath, isDBPresent);
                IteratePresentation(presentation);

                string nameSpace = "http://schemas.microsoft.com/sqlserver/reporting/2016/01/reportdefinition";
                rdlHelper.reportDef.RDLType = RdlDOM.RDLType.RDL2016;

                System.Xml.Serialization.XmlSerializerNamespaces serialize = new System.Xml.Serialization.XmlSerializerNamespaces();
                serialize.Add("rd", "http://schemas.microsoft.com/SQLServer/reporting/reportdesigner");
                XmlSerializer xmlSerializer = new XmlSerializer(typeof(RdlDOM.ReportDefinition), nameSpace);
                xmlSerializer.Serialize(memoryStream, rdlHelper.reportDef, serialize);
                memoryStream.Position = 0;
                return memoryStream;
            }
            catch (Exception ex)
            {
                //LoggerManager.LogError(ex.Message);
            }
            finally
            {
                Dispose();
            }


            return null;
        }

        /// <summary>
        /// Iterate through slides in the presentation.
        /// </summary>
        private void IteratePresentation(Presentation presentation)
        {
            try
            {
                double slideHeight = 0;
                double slideNo = 0;
                var connectionString = "Data Source=dataplatformdemodata.syncfusion.com;Initial Catalog=AdventureWorks;User ID='demoreadonly@data-platform-demo';Password='N@c)=Y8s*1&dh'";
                var newDataSource = rdlHelper.CreateDataSource("DataSource1", "SQL", connectionString);
                rdlHelper.reportDef.DataSources = new DataSources();
                rdlHelper.reportDef.DataSources.Add(newDataSource);
                var newDataSet = rdlHelper.CreateDataSet("DataSource1", "DataSet1");
                rdlHelper.reportDef.DataSets = new DataSets();
                rdlHelper.reportDef.DataSets.Add(newDataSet);

                ReportSection newSection = new ReportSection
                {
                    Width = presentation.SlideSize.Width + "pt",
                    Body = new Body()
                    {
                        Height = presentation.SlideSize.Height + "pt",
                        ReportItems = new ReportItems(),
                        Style = rdlHelper.AddStyle()
                    }
                };

                newSection.Page = new BoldReports.RDL.DOM.Page();
                newSection.Page.Style = rdlHelper.AddStyle();
                newSection.Page.PageHeight = presentation.SlideSize.Height + "pt";
                newSection.Page.PageWidth = presentation.SlideSize.Width + "pt";

                foreach (ISlide slide in presentation.Slides)
                {
                    slideNo += 1;

                    if(slideNo > 1)
                    {
                        slideHeight += presentation.SlideSize.Height;
                    }

                    IterateSlideContent(slide, presentation, newSection, slideHeight, slideNo);
                }

                if(rdlHelper.reportDef.ReportSections == null)
                {
                    rdlHelper.reportDef.ReportSections = new ReportSections();
                }
                rdlHelper.reportDef.ReportSections.Add(newSection);

            }
            catch (Exception ex)
            {
                //LoggerManager.LogError(ex.Message);
            }
        }

        /// <summary>
        /// Iterate slide content.
        /// </summary>
        private void IterateSlideContent(ISlide slide, Presentation presentation, ReportSection reportSection, double slideHeight,double slideNo)
        {
            foreach (IShape shape in slide.Shapes)
            {
                switch (shape.SlideItemType)
                {
                    case SlideItemType.AutoShape:
                        {
                            var rectangleF = new RectangleF((float)shape.Left, (float)shape.Top, (float)shape.Width, (float)shape.Height);
                            IterateTextBody(shape.ShapeName , shape.TextBody, rectangleF, reportSection, slideHeight, slideNo);
                            break;
                        }
                    case SlideItemType.Placeholder:
                        {
                            var rectangleF = new RectangleF((float)shape.Left, (float)shape.Top, (float)shape.Width, (float)shape.Height);
                            IterateTextBody(shape.ShapeName, shape.TextBody, rectangleF, reportSection, slideHeight, slideNo);
                            break;
                        }
                    case SlideItemType.Picture:
                        {
                            ProcessPicture((IPicture)shape, reportSection, slideHeight, slideNo);
                            break;
                        }
                    case SlideItemType.Table:
                        {
                            IterateTable((ITable)shape, reportSection, slideHeight, slideNo);
                            break;
                        }
                    //Print the GroupShape text in the console
                    case SlideItemType.GroupShape:
                        {

                            break;
                        }
                    //Print the chart data in the console
                    case SlideItemType.Chart:
                        {
                            IPresentationChart chart = shape as IPresentationChart;
                            presentation.PresentationRenderer = new PresentationRenderer();
                            //Converts the chart to image.
                            MemoryStream memoryStream = new MemoryStream();
                            presentation.PresentationRenderer.ConvertToImage(chart, memoryStream);
                            byte[] imageBytes = memoryStream.ToArray();
                            // IterateChart((IPresentationChart)shape, shape, reportSection);
                            var bounds = new RectangleF((float)shape.Left, (float)shape.Top, (float)shape.Width, (float)shape.Height);
                            rdlHelper.AddImageToRdl(shape.ShapeName, imageBytes, bounds, reportSection, slideHeight, slideNo);
                            break;
                        }
                    //Print the SmartArt text in the console
                    case SlideItemType.SmartArt:
                        {

                            break;
                        }
                }
            }
        }

        private void IterateChart(IPresentationChart chart, IShape shape, ReportSection reportSection)
        {
            rdlHelper.AddChartToRdl(chart, shape, reportSection);
        }

        /// <summary>
        /// Process text bodies (text elements inside slides)
        /// </summary>
        private void IterateTextBody(string itemName, ITextBody textBody, RectangleF bounds, ReportSection reportSection, double slideHeight, double slideNo)
        {
            foreach (IParagraph paragraph in textBody.Paragraphs)
            {
                string text = paragraph.Text;

                if(!string.IsNullOrEmpty(text))
                {
                    rdlHelper.AddTextToRdl(textBody, itemName, text, bounds, reportSection, slideHeight, slideNo);
                }
            }
        }

        /// <summary>
        /// Iterate PowerPoint tables.
        /// </summary>
        private void IterateTable(ITable table,ReportSection reportSection, double slideHeight, double slideNo)
        {
            var bounds = new RectangleF((float)table.Left, (float)table.Top, (float)table.Width, (float)table.Height);
            rdlHelper.AddTablixToRdl(table,table.ShapeName, table.Rows.Count(), table.Columns.Count(), bounds, reportSection, slideHeight, slideNo);
            //for (int row = 0; row < table.Rows; row++)
            //{
            //    for (int col = 0; col < table.Columns; col++)
            //    {
            //        ICell cell = table[row, col];
            //        if (cell.TextBody != null)
            //        {
            //            IterateTextBody(cell.TextBody);
            //        }
            //    }
            //}
        }

        /// <summary>
        /// Process PowerPoint images.
        /// </summary>
        private void ProcessPicture(IPicture picture, ReportSection reportSection, double slideHeight, double slideNo)
        {
            var bounds = new RectangleF((float)picture.Left, (float)picture.Top, (float)picture.Width, (float)picture.Height);
            rdlHelper.AddImageToRdl(picture.ShapeName, picture.ImageData, bounds, reportSection, slideHeight, slideNo);
        }

        internal void Dispose()
        {
            rdlHelper = null;
        }
    }
}
