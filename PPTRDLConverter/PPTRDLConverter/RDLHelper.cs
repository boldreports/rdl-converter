using Syncfusion.Presentation;
using Syncfusion.Drawing;
using System.IO;
using BoldReports.RDL.DOM;
using System;
using System.Linq;
using Syncfusion.XPS;
using System.Text;
using BitMiracle.LibTiff.Classic;
using System.Collections.Generic;
using Syncfusion.XlsIO.Implementation.Charts;
using BoldReports.Data;
using Field = BoldReports.RDL.DOM.Field;
using System.Reflection.Emit;
using Syncfusion.XlsIO;
using Rectangle = BoldReports.RDL.DOM.Rectangle;

namespace PresentationToRDL
{
    internal class RDLHelper
    {
        internal ReportDefinition reportDef = null;
        string fileName = string.Empty;

        public RDLHelper() { }

        public RDLHelper(string pptFileName, string connectionString, string dataProvider, string csvFilePath, bool isDBPresent)
        {
            try
            {
                reportDef = CreateReportDefinition();
                this.fileName = pptFileName;
            }
            catch (Exception ex)
            {
                //LoggerManager.LogError(ex.Message);
            }
        }

        public MemoryStream ConvertToRDL(string pptFilePath)
        {
            MemoryStream stream = new MemoryStream();
            try
            {
                Presentation presentation = (Presentation)Syncfusion.Presentation.Presentation.Open(pptFilePath);
                PresentationToRdlConverter converter = new PresentationToRdlConverter();
                stream = converter.Convert(presentation, System.IO.Path.GetFileNameWithoutExtension(pptFilePath), "", "", "", false);
                presentation.Close();
            }
            catch (Exception ex)
            {
                stream = null;
                //LoggerManager.LogError("PowerPoint path not found: " + ex.Message);
            }
            finally
            {
                Dispose();
            }
            return stream;
        }

        internal ReportDefinition CreateReportDefinition()
        {
            return new ReportDefinition
            {
                Author = "Bold Reports by Syncfusion",
                Description = "Converted PowerPoint to RDL",
                ReportUnitType = "Inch",
                RDLType = RDLType.RDL2010,

            };
        }

        internal void AddTextToRdl(ITextBody textBody, string itemName, string text, RectangleF bounds, ReportSection reportSection, double slideHeight, double slideNo)
        {
            itemName = itemName.Replace(" ", "");
            itemName = itemName + slideNo;

            if (reportSection.Body.ReportItems != null && reportSection.Body.ReportItems.Count > 0)
            {
                foreach (var reportItem in reportSection.Body.ReportItems)
                {
                    if (reportItem.Name == itemName && reportItem is TextBox)
                    {
                        return;
                    }
                }
            }

            TextBox textBox = new TextBox
            {
                Name = itemName,
                Width = bounds.Width + "pt",
                Height = bounds.Height + "pt",
                Left = bounds.X + "pt",
                Top = slideHeight + bounds.Y + "pt",
            };

            textBox.Paragraphs = new Paragraphs();

            foreach (Syncfusion.Presentation.IParagraph paragraph in textBody.Paragraphs)
            {
                textBox.Paragraphs.Add(UpdateParagraph(textBody, paragraph, null, paragraph.Text));
            }

            textBox.Style = AddStyle();

            reportSection.Body.ReportItems.Add(textBox);
        }
        

        internal void AddImageToRdl(string itemName, byte[] imageBytes, RectangleF bounds, ReportSection reportSection, double slideHeight, double slideNo)
        {
            itemName = itemName.Replace(" ", "");
            itemName = itemName + slideNo;

            if (reportSection.Body.ReportItems != null && reportSection.Body.ReportItems.Count > 0)
            {
                foreach (var reportItem in reportSection.Body.ReportItems)
                {
                    if (reportItem.Name == itemName && reportItem is BoldReports.RDL.DOM.Image)
                    {
                        return;
                    }
                }
            }
            
                BoldReports.RDL.DOM.Image image = new BoldReports.RDL.DOM.Image
            {
                Name = itemName,
                Width = bounds.Width + "pt",
                Height = bounds.Height + "pt",
                Left = bounds.X + "pt",
                Top = slideHeight + bounds.Y + "pt",
                Source = Source.Embedded,
                Sizing = Sizing.FitProportional,
                Value = "EmbeddedImage" + itemName
                };

            image.Style = AddStyle();

            if (reportDef.EmbeddedImages == null)
                reportDef.EmbeddedImages = new EmbeddedImages();

            reportDef.EmbeddedImages.Add(new EmbeddedImage
            {
                Name = "EmbeddedImage" + itemName,
                MIMEType = "image/png",
                ImageData = Convert.ToBase64String(imageBytes)
            });

            reportSection.Body.ReportItems.Add(image);
        }

        internal void AddTablixToRdl(ITable table, string itemName, int numRows, int numColumns, RectangleF bounds, ReportSection reportSection, double slideHeight, double slideNo)
        {
            itemName = itemName.Replace(" ", "");
            itemName = itemName + slideNo;

            string csvFilePath = "C:\\Users\\MahendranShanmugam\\source\\repos\\ConsoleApp3\\ConsoleApp3\\Data\\Report\\" + itemName + ".csv";
            //string excelFilePath = "C:\\Users\\MahendranShanmugam\\source\\repos\\ConsoleApp3\\ConsoleApp3\\Data\\Report\\" + itemName + ".xls";
            // Set your path
            //CreateCsvFile(table, csvFilePath);
            //AddCsvDataSource(csvFilePath, itemName);
            //AddCsvDataSet(numColumns, itemName);

            //CreateExcelFile(table, excelFilePath);
            //AddExcelDataSource(excelFilePath, itemName);
            //AddExcelDataSet(numColumns, itemName);

            //if (reportDef.ReportSections[0].Body.ReportItems != null && reportDef.ReportSections[0].Body.ReportItems.Count > 0)
            //{
            //    foreach (var reportItem in reportDef.ReportSections[0].Body.ReportItems)
            //    {
            //        if (reportItem.Name == itemName && reportItem is Tablix)
            //        {
            //            return;
            //        }
            //    }
            //}


            Tablix tablix = new Tablix
            {
                Name = itemName,
                DataSetName = "DataSet1", // Bind to CSV DataSet
                Width = bounds.Width + "pt",
                Height = bounds.Height + "pt",
                Left = bounds.X + "pt",
                Top = slideHeight + bounds.Y + "pt",
                TablixBody = new TablixBody
                {
                    TablixColumns = new TablixColumns(),
                    TablixRows = new TablixRows()
                },
                TablixColumnHierarchy = new TablixColumnHierarchy { TablixMembers = new TablixMembers() },
                TablixRowHierarchy = new TablixRowHierarchy { TablixMembers = new TablixMembers() }
            };

            // Extract font properties from the first cell in the table
            Style tablixStyle = new Style
            {
                Border = new Border
                {
                    Style = "Solid",
                    Width = "1pt",
                    Color = "Black"
                },
                //PaddingLeft = "2pt",
                //PaddingRight = "2pt",
                //PaddingTop = "2pt",
                //PaddingBottom = "2pt",
                LineHeight = "12pt"
            };

            if (numRows > 0 && numColumns > 0)
            {
                ICell firstCell = table[0, 0];
                if (firstCell.TextBody.Paragraphs.Count > 0)
                {
                    IParagraph paragraph = firstCell.TextBody.Paragraphs[0];

                    tablixStyle.FontFamily = paragraph.Font.FontName;
                    tablixStyle.FontSize = paragraph.Font.FontSize.ToString() + "pt";
                    tablixStyle.FontWeight = paragraph.Font.Bold ? "Bold" : "Normal";
                    tablixStyle.FontStyle = paragraph.Font.Italic ? "Italic" : "Normal";
                    tablixStyle.TextAlign = paragraph.HorizontalAlignment.ToString();
                }
            }

            tablix.Style = tablixStyle;

            // Add columns
            for (int col = 0; col < numColumns; col++)
            {
                tablix.TablixBody.TablixColumns.Add(new TablixColumn { Width = (bounds.Width / numColumns) + "pt" });
                tablix.TablixColumnHierarchy.TablixMembers.Add(new TablixMember());
            }

            // Add rows and cells
            for (int row = 0; row < numRows; row++)
            {
                // Create a header row
                TablixRow headerRow = new TablixRow { Height = "20pt" };
                headerRow.TablixCells = new TablixCells();

                //    for (int col = 0; col < numColumns; col++)
                //    {
                //        ICell cell = table[0, col];
                //        string headerText = cell.TextBody.Paragraphs.Count > 0 && !string.IsNullOrEmpty(cell.TextBody.Paragraphs[0].Text) ? cell.TextBody.Paragraphs[0].Text: "Column" + col;

                //        TextBox headerTextBox = new TextBox
                //        {
                //            Style = new BoldReports.RDL.DOM.Style
                //            {
                //                FontWeight = "Bold",
                //                BackgroundColor = "#D3D3D3",
                //                //BorderStyle = BorderStyle.Solid,
                //                Border = new Border() { Style = "Solid"}
                //            },
                //            Name = $"{itemName}_header_{col}",
                //            CanGrow = true,
                //            Paragraphs = new Paragraphs
                //{
                //    new BoldReports.RDL.DOM.Paragraph
                //    {
                //        Style = new BoldReports.RDL.DOM.Style { FontWeight = "Bold" },
                //        TextRuns = new TextRuns
                //        {
                //            new TextRun { Value = headerText, Style = new BoldReports.RDL.DOM.Style { FontWeight = "Bold" } }
                //        }
                //    }
                //}
                //        };

                //        headerRow.TablixCells.Add(new TablixCell
                //        {
                //            CellContents = new CellContents { ReportItem = headerTextBox, ColSpan = 1, RowSpan = 1 }
                //        });
                //    }

                //    // Add header row to Tablix
                //    tablix.TablixBody.TablixRows.Insert(0, headerRow); // Insert at the top
                //    tablix.TablixRowHierarchy.TablixMembers.Insert(0, new TablixMember()); // Add hierarchy member for header

                TablixRow tablixRow = new TablixRow { Height = (bounds.Height / numRows) + "pt" };
                tablixRow.TablixCells = new TablixCells();

                for (int col = 0; col < numColumns; col++)
                {
                    ICell cell = table[row, col];
                    string cellText = "Column" + (col + 1);

                    TextBox cellTextBox = new TextBox
                    {
                        Style = tablixStyle,
                        Name = $"{itemName}_textbox_{col}",
                        CanGrow = true,
                        Paragraphs = new Paragraphs
                {
                    new BoldReports.RDL.DOM.Paragraph
                    {
                        //Style = tablixStyle,
                        TextRuns = new TextRuns
                        {
                            new TextRun { Style = new Style{FontWeight = "Normal"}, Value = cell.TextBody.Paragraphs.Count > 0 ? cell.TextBody.Paragraphs[0].Text : "" }
                        }
                    }
                }
                    };

                    tablixRow.TablixCells.Add(new TablixCell
                    {
                        CellContents = new CellContents { ReportItem = cellTextBox, ColSpan = 1, RowSpan = 1 }
                    });
                }

                tablix.TablixBody.TablixRows.Add(tablixRow);
                tablix.TablixRowHierarchy.TablixMembers.Add(new TablixMember() { /*Group = AddGroup(table.ShapeName)*/ });
            }

            reportSection.Body.ReportItems.Add(tablix);
        }

        internal Group AddGroup(string itemName)
        {
            itemName = itemName.Replace(" ", "");

            return new Group
            {
                Name = itemName + "Details"
            };
        }

        internal Style AddStyle()
        {
            return new Style
            {
                Border = new Border
                {
                    Style = "None",
                    Width = "1pt",
                    Color = "Black"
                },
                //PaddingLeft = "2pt",
                //PaddingRight = "2pt",
                //PaddingTop = "2pt",
                //PaddingBottom = "2pt",
                //LineHeight = "12pt"
            };
        }

        internal BoldReports.RDL.DOM.Paragraph UpdateParagraph(ITextBody item, IParagraph paragraph, BoldReports.RDL.DOM.ReportItem textBox, string text)
        {
            try
            {
                var para = new BoldReports.RDL.DOM.Paragraph
                {
                    Style = new Style
                    {
                        Border = new Border
                        {
                            // Example: Set border properties (modify as needed)
                            Style = "Solid",
                            Width = "1pt",
                            Color = "Black"
                        },
                        FontFamily = paragraph.Font.FontName,
                        FontSize = paragraph.Font.FontSize.ToString() + "pt",
                        FontWeight = paragraph.Font.Bold ? "Bold" : "Normal",
                        FontStyle = paragraph.Font.Italic ? "Italic" : "Normal",
                        //TextDecoration = paragraph.Font.Underline ? "Underline" : "None",
                        //Color = paragraph.Font.Color.RGB,
                        //PaddingLeft = "2pt",
                        //PaddingRight = "2pt",
                        //PaddingTop = "2pt",
                        //PaddingBottom = "2pt",
                        TextAlign = paragraph.HorizontalAlignment.ToString(),
                        LineHeight = "12pt"
                    }
                };

                // Check if the paragraph has bullet formatting and apply list style
                if (paragraph.ListFormat.Type != ListType.None)
                {
                    para.ListStyle = paragraph.ListFormat.Type == ListType.Bulleted
                        ? BoldReports.RDL.DOM.ListStyle.Bulleted
                        : BoldReports.RDL.DOM.ListStyle.Numbered;
                    para.ListLevel = paragraph.ListFormat.StartValue + 1;
                }

                para.TextRuns = new TextRuns();
                var textRun = new TextRun();
                textRun.Style = new Style();
                textRun.Style.FontFamily = paragraph.Font.FontName;
                textRun.Style.FontSize = paragraph.Font.FontSize.ToString() + "pt";
                textRun.Style.FontWeight = paragraph.Font.Bold ? "Bold" : "Normal";
                textRun.Style.FontStyle = paragraph.Font.Italic ? "Italic" : "Normal";
                textRun.Value = text;
                para.TextRuns.Add(textRun);
                return para;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }

            return null;
        }

        internal void CreateCsvFile(ITable table, string filePath)
        {
            StringBuilder csvContent = new StringBuilder();

            // Add header row
            for (int col = 0; col < table.Columns.Count; col++)
            {
                ICell cell = table[0, col];
                string cellText = $"Column{col+1}";
                csvContent.Append(cellText + ",");
            }
            csvContent.Length--; // Remove last comma
            csvContent.AppendLine();

            // Add data rows
            for (int row = 1; row < table.Rows.Count(); row++)
            {
                for (int col = 0; col < table.Columns.Count(); col++)
                {
                    ICell cell = table[row, col];
                    string cellText = cell.TextBody.Paragraphs.Count > 0 ? cell.TextBody.Paragraphs[0].Text : "";
                    csvContent.Append($"\"{cellText}\",");
                }
                csvContent.Length--; // Remove last comma
                csvContent.AppendLine();
            }

            // Write CSV file
            File.WriteAllText(filePath, csvContent.ToString(), Encoding.UTF8);
        }

        internal void AddCsvDataSource(string csvFilePath, string itemName)
        {
            itemName = itemName + "_DataSource";

            if (reportDef.DataSources != null && reportDef.DataSources.Count > 0)
            {
                foreach(var dataSource in reportDef.DataSources)
                {
                    if (dataSource.Name == itemName) {
                        return;
                    }
                }
            }

            DataSource csvDataSource = new DataSource
            {
                Name = itemName,
                ConnectionProperties = new BoldReports.RDL.DOM.ConnectionProperties
                {
                    DataProvider = "CSV",
                    ConnectString = "{\"Data\":\"\",\"DataMode\":\"file\",\"URL\":\"\",\"IsCSVFirstRowHeader\":true,\"Separator\":\"comma\",\"Delimiter\":\"\"}",
                    EmbeddedData = new EmbeddedData
                    {
                        Name = Guid.NewGuid().ToString(), // Unique ID
                        FileType = "csv",
                        FileName = System.IO.Path.GetFileNameWithoutExtension(csvFilePath),
                        Data = GetEmbeddedDataFromCsv(csvFilePath) // Base64-encoded CSV data
                    }
                }
            };

            csvDataSource.ImpersonateUser = false;

            if (reportDef.DataSources == null)
                reportDef.DataSources = new DataSources();

            reportDef.DataSources.Add(csvDataSource);

        }


        private string GetEmbeddedDataFromCsv(string csvFilePath)
        {
            if (!File.Exists(csvFilePath))
                throw new FileNotFoundException("CSV file not found", csvFilePath);

            // Read CSV as UTF-8 and remove BOM if present
            string csvContent = File.ReadAllText(csvFilePath, new UTF8Encoding(true));
            if (csvContent.Length > 0 && csvContent[0] == '\ufeff')
            {
                csvContent = csvContent.Substring(1); // Remove BOM
            }

            byte[] csvBytes = Encoding.UTF8.GetBytes(csvContent);
            return Convert.ToBase64String(csvBytes); // Convert to Base64 for embedding
        }

        internal void AddCsvDataSet(int numColumns, string itemName)
        {
            if (reportDef.DataSets != null && reportDef.DataSets.Count > 0)
            {
                foreach (var dataSet in reportDef.DataSets)
                {
                    if (dataSet.Name == itemName + "_DataSet")
                    {
                        return;
                    }
                }
            }

            DataSet csvDataSet = new DataSet
            {
                Name = itemName + "_DataSet",
                Query = new Query
                {
                    DataSourceName = itemName + "_DataSource",
                    CommandType = BoldReports.RDL.DOM.CommandType.Text,
                    CommandText = $"{{\"Name\":\"DataSource\",\"Columns\":[]}}"
                }
            };

            // Add QueryDesignerState as an XML structure
            csvDataSet.Query.QueryDesignerState = new QueryDesignerState
            {
                Tables = new List<BoldReports.RDL.DOM.Table>
                {
                     new BoldReports.RDL.DOM.Table
                     {
                         Name = "Datasource",
                         Schema = "",
                         Columns = new List<BoldReports.RDL.DOM.Column>() // Initialize the list first
                     }
                }
            };

            // Populate columns dynamically using a loop
            for (int col = 0; col < numColumns; col++)
            {
                csvDataSet.Query.QueryDesignerState.Tables[0].Columns.Add(new BoldReports.RDL.DOM.Column
                {
                    Name = $"Column{col + 1}",
                    IsDuplicate = "False",
                    IsSelected = "True"
                });

                if (csvDataSet.Fields == null)
                    csvDataSet.Fields = new Fields();

                // Add field definitions inside DataSet fields
                csvDataSet.Fields.Add(new Field
                {
                    Name = $"Column{col + 1}",
                    DataField = $"Column{col + 1}",
                    TypeName = "System.String"
                });
            }
            if (reportDef.DataSets == null)
                reportDef.DataSets = new DataSets();

            reportDef.DataSets.Add(csvDataSet);
        }


        internal void CreateExcelFile(ITable table, string filePath)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;

                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];

                // Add headers
                for (int col = 0; col < table.Columns.Count; col++)
                {
                    ICell cell = table[0, col];
                    string cellText = cell.TextBody?.Paragraphs.Count > 0 ? cell.TextBody.Paragraphs[0].Text.Trim() : $"Column{col + 1}";
                    sheet.Range[1, col + 1].Text = cellText;
                }

                // Add data rows
                for (int row = 1; row < table.Rows.Count; row++)
                {
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        ICell cell = table[row, col];
                        string cellText = cell.TextBody?.Paragraphs.Count > 0 ? cell.TextBody.Paragraphs[0].Text.Trim() : "";
                        sheet.Range[row + 1, col + 1].Text = cellText;
                    }
                }

                // Save to file
                workbook.SaveAs(filePath);
            }
        }

        internal void AddExcelDataSource(string excelFilePath, string itemName)
        {
            itemName = itemName + "_DataSource";

            if (reportDef.DataSources?.Any(ds => ds.Name == itemName) == true)
            {
                return; // DataSource already exists
            }

            DataSource excelDataSource = new DataSource
            {
                Name = itemName,
                ConnectionProperties = new BoldReports.RDL.DOM.ConnectionProperties
                {
                    DataProvider = "Excel",
                    ConnectString = $"{{\"Data\":\"\",\"DataMode\":\"file\",\"URL\":\"\",\"IsCSVFirstRowHeader\":true,\"Separator\":\"\",\"Delimiter\":\"\",\"ExcelMode\":\"xlsx\"}}",
                    EmbeddedData = new EmbeddedData
                    {
                        Name = Guid.NewGuid().ToString(),
                        FileType = "excel",
                        FileName = System.IO.Path.GetFileNameWithoutExtension(excelFilePath),
                        Data = GetEmbeddedDataFromExcel(excelFilePath)
                    }
                }
            };

            excelDataSource.ImpersonateUser = false;

             if(reportDef.DataSources == null)
            {
                reportDef.DataSources = new DataSources();
            }
            reportDef.DataSources.Add(excelDataSource);
        }

        private string GetEmbeddedDataFromExcel(string excelFilePath)
        {
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException("Excel file not found", excelFilePath);

            byte[] excelBytes = File.ReadAllBytes(excelFilePath);
            return Convert.ToBase64String(excelBytes);
        }

        internal void AddExcelDataSet(int numColumns, string itemName)
        {
            string dataSetName = itemName + "_DataSet";
            if (reportDef.DataSets?.Any(ds => ds.Name == dataSetName) == true)
            {
                return; // DataSet already exists
            }

            DataSet excelDataSet = new DataSet
            {
                Name = dataSetName,
                Query = new Query
                {
                    DataSourceName = itemName + "_DataSource",
                    CommandType = BoldReports.RDL.DOM.CommandType.Text,
                    CommandText = $"{{\"Name\":\"DataSource\",\"Columns\":[{string.Join(",", Enumerable.Range(1, numColumns).Select(i => $"{{\"Name\":\"Column{i}\"}}"))}]}}"
                }
            };

            // Add QueryDesignerState
            excelDataSet.Query.QueryDesignerState = new QueryDesignerState
            {
                Tables = new List<BoldReports.RDL.DOM.Table>
        {
            new BoldReports.RDL.DOM.Table
            {
                Name = "Datasource",
                Schema = "",
                Columns = Enumerable.Range(1, numColumns).Select(i => new BoldReports.RDL.DOM.Column
                {
                    Name = $"Column{i}",
                    IsDuplicate = "False",
                    IsSelected = "True"
                }).ToList()
            }
        }
            };

            excelDataSet.Fields = new Fields();
            for (int col = 0; col < numColumns; col++)
            {
                excelDataSet.Fields.Add(new Field
                {
                    Name = $"Column{col + 1}",
                    DataField = $"Column{col + 1}",
                    TypeName = "System.String"
                });
            }

            if(reportDef.DataSets == null)
            {
                reportDef.DataSets = new DataSets();
            }
            reportDef.DataSets.Add(excelDataSet);
        }

        internal void AddChartToRdl(IPresentationChart pptChart, Syncfusion.Presentation.IShape shape, ReportSection reportSection)
        {
            // Create RDL Chart
            var chartItem = new Chart
            {
                Name = shape.ShapeName ?? "Chart",
                Height = new BoldReports.RDL.DOM.Size(shape.Height),
                Width = new BoldReports.RDL.DOM.Size(shape.Width),
                Left = new BoldReports.RDL.DOM.Size(shape.Left),
                Top = new BoldReports.RDL.DOM.Size(shape.Top),
                DataSetName = "DataSet1",
                Style = AddStyle(), // Assign background color
                //Palette = pptChart. ?? "Pacific", // Assign Palette from pptChart

                ChartCategoryHierarchy = new ChartCategoryHierarchy { ChartMembers = new ChartMembers() },
                ChartSeriesHierarchy = new ChartSeriesHierarchy { ChartMembers = new ChartMembers() },
                ChartData = new ChartData
                {
                    ChartDerivedSeriesCollection = new ChartDerivedSeriesCollection(),
                    ChartSeriesCollection = CreateChartSeriesCollection(pptChart)
                },
                ChartLegends = new ChartLegends { CreateChartLegend(pptChart) }, // Assign Chart Legend
                ChartAreas = new ChartAreas { CreateChartArea(pptChart) }, // Assign Chart Area
                ChartTitles = CreateChartTitles(pptChart), // Assign Chart Titles
                ChartBorderSkin = new ChartBorderSkin {  }, // Assign Border Style
                ChartNoDataMessage = new ChartNoDataMessage
                {
                    Caption = "No Data Available",
                    Name = "NoDataMessage"
                }
            };

            // Add Category Axis
            //if (pptChart.PrimaryCategoryAxis != null)
            //{
            //    var categoryMember = ChartCategoryMember(pptChart);
            //    categoryMember.Label = $"=Fields!{pptChart.PrimaryCategoryAxis.Title}.Value"; // Assign category axis title
            //    chartItem.ChartCategoryHierarchy.ChartMembers.Add(categoryMember);
            //}

            // Add Series Axis
            var seriesMember = ChartSeriesMember(pptChart);
            chartItem.ChartSeriesHierarchy.ChartMembers.Add(seriesMember);

            // Add chart to report
            // reportSection.Body.ReportItems.Add(chartItem);
        }

        // Assign Chart Title(s)
        BoldReports.RDL.DOM.ChartTitles CreateChartTitles(IPresentationChart pptChart)
        {
            var chartTitles = new ChartTitles();
            var chartTitle = new ChartTitle
            {
                Caption  = pptChart.ChartTitle ?? "Chart Title",
                Style = AddStyle()
            };

            chartTitles.Add(chartTitle);
            return chartTitles;
        }

        // Assign Chart Legend
        BoldReports.RDL.DOM.ChartLegend CreateChartLegend(IPresentationChart pptChart)
        {
            return new ChartLegend
            {

            };
        }

        // Assign Category Member
        BoldReports.RDL.DOM.ChartMember ChartCategoryMember(IPresentationChart pptChart)
        {
            return new ChartMember
            {
                ChartMembers = new ChartMembers(),
                CustomProperties = new CustomProperties(),
                //DataElementName = pptChart.PrimaryCategoryAxis?.Title,
                DataElementOutput = DataElementOutputs.Auto,
                Group = CreateChartGroup(pptChart),
               // Label = $"=Fields!{pptChart.PrimaryCategoryAxis?.Title}.Value",
                SortExpressions = new SortExpressions()
            };
        }

        // Assign Series Member
        BoldReports.RDL.DOM.ChartMember ChartSeriesMember(IPresentationChart pptChart)
        {
            return new ChartMember
            {
                ChartMembers = new ChartMembers(),
                CustomProperties = new CustomProperties(),
                //DataElementName = pptChart.Series.FirstOrDefault()?.Name,
                DataElementOutput = DataElementOutputs.Auto,
                //Label = $"=Fields!{pptChart.Series.FirstOrDefault()?.Name}.Value",
                SortExpressions = new SortExpressions()
            };
        }

        // Assign Chart Group
        BoldReports.RDL.DOM.Group CreateChartGroup(IPresentationChart pptChart)
        {
            return new Group
            {
                //DataElementName = pptChart.PrimaryCategoryAxis?.Title,
                DataElementOutput = DataElementOutputs.Auto,
                GroupExpressions = CreateGroupExpressions(pptChart),
                Name = "Chart2_CategoryGroup"
            };
        }

        // Assign Group Expressions
        BoldReports.RDL.DOM.GroupExpressions CreateGroupExpressions(IPresentationChart pptChart)
        {
            var groupExpressions = new GroupExpressions();
            groupExpressions.Add(new GroupExpression
            {
                //Value = $"=Fields!{pptChart.PrimaryCategoryAxis?.Title}.Value"
            });

            return groupExpressions;
        }

        // Assign Chart Area
        BoldReports.RDL.DOM.ChartArea CreateChartArea(IPresentationChart pptChart)
        {
            return new ChartArea
            {
                Style = AddStyle(),
                Name = "Default",
                ChartCategoryAxes = new ChartCategoryAxes
        {
            createChartAxis("Primary", pptChart, true),
            createChartAxis("Secondary", pptChart, true)
        },
                ChartValueAxes = new ChartValueAxes
        {
            createChartAxis("Primary", pptChart, false),
            createChartAxis("Secondary", pptChart, false)
        }
            };
        }

        // Assign Chart Axis
        BoldReports.RDL.DOM.ChartAxis createChartAxis(string Name, IPresentationChart pptChart, bool IsCategoryAxis)
        {
            ChartAxis axis = new ChartAxis();
            axis.ChartAxisTitle = new ChartAxisTitle();

            axis.Name = Name;
            axis.ChartAxisTitle.Style = new Style { FontWeight = "Bold" };

            
            //if (IsCategoryAxis && Name == "Primary")
            //{
            //    if (pptChart.PrimaryCategoryAxis != null)
            //    {
            //        axis.ChartAxisTitle.Caption = pptChart.PrimaryCategoryAxis?.Title ?? "Axis";
            //        axis.Maximum = pptChart.PrimaryCategoryAxis.MaximumValue.ToString();
            //        axis.Minimum = pptChart.PrimaryCategoryAxis.MinimumValue.ToString();
            //        axis.ChartMajorGridLines = (ChartMajorGridLines)pptChart.PrimaryCategoryAxis.MajorGridLines;
            //        //axis.ChartMajorTickMarks = (ChartMajorTickMarks)pptChart.PrimaryCategoryAxis.MajorTickMark;
            //        axis.ChartMinorGridLines = (ChartMinorGridLines)pptChart.PrimaryCategoryAxis.MinorGridLines;
            //        //axis.ChartMinorTickMarks = (ChartMinorTickMarks)pptChart.PrimaryCategoryAxis.MinorTickMark;
            //    }
            //}
            //else if (IsCategoryAxis && Name == "Secondary")
            //{
            //    if(pptChart.SecondaryCategoryAxis != null)
            //    {
            //        axis.ChartAxisTitle.Caption = pptChart.SecondaryCategoryAxis?.Title ?? "Axis";
            //        axis.Maximum = pptChart.SecondaryCategoryAxis.MaximumValue.ToString();
            //        axis.Minimum = pptChart.SecondaryCategoryAxis.MinimumValue.ToString();
            //        axis.ChartMajorGridLines = (ChartMajorGridLines)pptChart.SecondaryCategoryAxis.MajorGridLines;
            //        //axis.ChartMajorTickMarks = (ChartMajorTickMarks)pptChart.SecondaryCategoryAxis.MajorTickMark;
            //        axis.ChartMinorGridLines = (ChartMinorGridLines)pptChart.SecondaryCategoryAxis.MinorGridLines;
            //        //axis.ChartMinorTickMarks = (ChartMinorTickMarks)pptChart.SecondaryCategoryAxis.MinorTickMark;
            //    }
            //}
            //else if (!IsCategoryAxis && Name == "Primary")
            //{
            //    if(pptChart.PrimaryValueAxis != null)
            //    {
            //        axis.ChartAxisTitle.Caption = pptChart.PrimaryValueAxis?.Title ?? "Axis";
            //        axis.Maximum = pptChart.PrimaryValueAxis.MaximumValue.ToString();
            //        axis.Minimum = pptChart.PrimaryValueAxis.MinimumValue.ToString();
            //        axis.ChartMajorGridLines = (ChartMajorGridLines)pptChart.PrimaryValueAxis.MajorGridLines;
            //        //axis.ChartMajorTickMarks = (ChartMajorTickMarks)pptChart.PrimaryValueAxis.MajorTickMark;
            //        axis.ChartMinorGridLines = (ChartMinorGridLines)pptChart.PrimaryValueAxis.MinorGridLines;
            //        //axis.ChartMinorTickMarks = (ChartMinorTickMarks)pptChart.PrimaryValueAxis.MinorTickMark;
            //    }
            //}
            //else if (!IsCategoryAxis && Name == "Secondary")
            //{
            //    if(pptChart.SecondaryValueAxis != null)
            //    {
            //        axis.ChartAxisTitle.Caption = pptChart.SecondaryValueAxis?.Title ?? "Axis";
            //        axis.Maximum = pptChart.SecondaryValueAxis.MaximumValue.ToString();
            //        axis.Minimum = pptChart.SecondaryValueAxis.MinimumValue.ToString();
            //        axis.ChartMajorGridLines = (ChartMajorGridLines)pptChart.SecondaryValueAxis.MajorGridLines;
            //        //axis.ChartMajorTickMarks = (ChartMajorTickMarks)pptChart.SecondaryValueAxis.MajorTickMark;
            //        axis.ChartMinorGridLines = (ChartMinorGridLines)pptChart.SecondaryValueAxis.MinorGridLines;
            //        //axis.ChartMinorTickMarks = (ChartMinorTickMarks)pptChart.SecondaryValueAxis.MinorTickMark;
            //    }
            //}

            return axis;

        }

        // Assign Chart Series Collection
        BoldReports.RDL.DOM.ChartSeriesCollection CreateChartSeriesCollection(IPresentationChart pptChart)
        {
            var chartSeriesCollection = new BoldReports.RDL.DOM.ChartSeriesCollection();

            //foreach (var series in pptChart.Series)
            //{
            //    var chartSeries = new ChartSeries
            //    {
            //        Name = series.Name ?? "Series",
            //        Type = ConvertChartType(pptChart.ChartType),
            //        Subtype = VisualizationSubType.Plain,
            //        ChartDataLabel = new ChartDataLabel(),
            //        ChartDataPoints = CreateChartDataPoints(pptChart)
            //    };

            //    foreach (var point in series.DataPoints)
            //    {
            //        chartSeries.ChartDataPoints.Add(new ChartDataPoint
            //        {
            //            ChartDataPointValues = new ChartDataPointValues { Y = point.ToString() }
            //        });
            //    }

               // chartSeriesCollection.Add(chartSeries);
            //}

            return chartSeriesCollection;
        }

        private ChartDataPoints CreateChartDataPoints(IPresentationChart pptChart)
        {
            ChartDataPoints datapoints = new ChartDataPoints();

            //foreach (var series in pptChart.Series)
            //{
            //    // Assign data points
            //    foreach (var point in series.DataPoints)
            //    {
            //        datapoints.Add(new ChartDataPoint
            //        {
            //            ChartDataPointValues = new ChartDataPointValues { Y = point.ToString() }
            //        });
            //    }
            //}

            return datapoints;
        }

        //// Helper method to convert PowerPoint Chart Type to BoldReports Chart Type
        //internal VisualizationType ConvertChartType(OfficeChartType pptChartType)
        //{
        //    switch (pptChartType)
        //    {
        //        case OfficeChartType.Column_Clustered:
        //        case OfficeChartType.Column_Stacked:
        //        case OfficeChartType.Column_Stacked_100:
        //            return VisualizationType.Column;
        //        case OfficeChartType.Line:
        //        case OfficeChartType.Line_Stacked:
        //        case OfficeChartType.Line_Stacked_100:
        //            return VisualizationType.Line;
        //        //case OfficeChartType.Pie:
        //        //case OfficeChartType.Pie_Exploded:
        //        //    return VisualizationType.pie;
        //        case OfficeChartType.Bar_Clustered:
        //        case OfficeChartType.Bar_Stacked:
        //        case OfficeChartType.Bar_Stacked_100:
        //            return VisualizationType.Bar;
        //        case OfficeChartType.Area:
        //        case OfficeChartType.Area_Stacked:
        //            return VisualizationType.Area;
        //        default:
        //            return VisualizationType.Column; // Default type
        //}

       // }

        //If you are using RDLC report then you can pass any string value as connection string
        internal BoldReports.RDL.DOM.DataSource CreateDataSource(string name, string dataProvider, string connectionString)
        {
            var dataSource = new BoldReports.RDL.DOM.DataSource();
            dataSource.Name = name;
            dataSource.SecurityType = BoldReports.RDL.DOM.SecurityType.None;
            dataSource.ConnectionProperties = new BoldReports.RDL.DOM.ConnectionProperties();
            dataSource.ConnectionProperties.DataProvider = dataProvider;
            dataSource.ConnectionProperties.ConnectString = connectionString;
            //dataSource.ConnectionProperties.IntegratedSecurity = true;

            return dataSource;
        }

        internal BoldReports.RDL.DOM.DataSet CreateDataSet(string dataSourceName, string dataSetName)
        {
            var dataSet = new BoldReports.RDL.DOM.DataSet();
            dataSet.Name = dataSetName;
            dataSet.Query = new Query();
            dataSet.Query.DataSourceName = dataSourceName;
            dataSet.Query.CommandText = "SELECT HumanResources.Department.DepartmentID,HumanResources.Department.Name,HumanResources.Department.GroupName,HumanResources.Department.ModifiedDate FROM HumanResources.Department";
            dataSet.Query.QueryDesignerState = new QueryDesignerState();
            var table = CreateTable();
            dataSet.Query.QueryDesignerState.Tables = new List<BoldReports.RDL.DOM.Table>();
            dataSet.Query.QueryDesignerState.Tables.Add(table);
            dataSet.Fields = new Fields();
            dataSet.Fields.Add(CreateField("DepartmentID", "System.Int16"));
            dataSet.Fields.Add(CreateField("Name", "System.String"));
            dataSet.Fields.Add(CreateField("GroupName", "System.String"));
            dataSet.Fields.Add(CreateField("ModifiedDate", "System.DateTime"));

            return dataSet;
        }

        BoldReports.RDL.DOM.Table CreateTable()
        {
            var table = new BoldReports.RDL.DOM.Table();
            table.Schema = "HumanResources";
            table.Name = "Department";
            table.Columns = new List<BoldReports.RDL.DOM.Column>();
            table.Columns.Add(CreateColumn("DepartmentID"));
            table.Columns.Add(CreateColumn("Name"));
            table.Columns.Add(CreateColumn("GroupName"));
            table.Columns.Add(CreateColumn("ModifiedDate"));
            return table;
        }

        BoldReports.RDL.DOM.Column CreateColumn(string columnName)
        {
            var column = new BoldReports.RDL.DOM.Column();
            column.Name = columnName;
            return column;
        }

        BoldReports.RDL.DOM.Field CreateField(string fieldName, string type)
        {
            var field = new Field();
            field.Name = fieldName;
            field.TypeName = type;
            field.DataField = fieldName;
            return field;
        }

        internal void Dispose()
        {
            reportDef = null;
            fileName = string.Empty;
        }


    }
}
