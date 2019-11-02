using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using System;
using System.Collections.Generic;
using System.IO;
using System.Drawing;
using Color = System.Drawing.Color;

namespace WobbleSoft.Img2Sheet
{
    class Program
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "Img2Sheet";
        static string SpreadsheetId = "SHEETID";
        static SheetsService service;

        static int PixelSize = 8; 
        static int RowsPerUpdate = 100;

        static void Main(string[] args)
        {
            SetupSheetsService();
            DrawSheet("SheetName", new Bitmap("file.png"));
        }

        static void DrawSheet(string sheetName, Bitmap image)
        {
            string range = sheetName + $"!A1:{ColumnToLetter(image.Width)}{image.Height}";
            BatchUpdateSpreadsheetRequest body;
            int sheetId = -1;

            try
            {
                //try get sheet
                SpreadsheetsResource.GetRequest getSheet = service.Spreadsheets.Get(SpreadsheetId);
                getSheet.Ranges = range;
                getSheet.IncludeGridData = false;
                Spreadsheet sheetResponse = getSheet.Execute();
                sheetId = sheetResponse.Sheets[0].Properties.SheetId.Value;
            }
            catch (Google.GoogleApiException)
            {
                //Create new sheet
                body = new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>() {
                        new Request()
                        {
                            AddSheet = new AddSheetRequest()
                            {
                                Properties = new SheetProperties()
                                {
                                    Title = sheetName,
                                    GridProperties = new GridProperties()
                                    {
                                        ColumnCount = image.Width,
                                        RowCount = image.Height,
                                    },
                                }
                            }
                        },
                    }
                };
                var createSheetUpdate = service.Spreadsheets.BatchUpdate(body, SpreadsheetId);
                createSheetUpdate.Execute();

                //Fetch new sheet
                SpreadsheetsResource.GetRequest request = service.Spreadsheets.Get(SpreadsheetId);
                request.Ranges = range;
                request.IncludeGridData = false;
                Spreadsheet response = request.Execute();
                sheetId = response.Sheets[0].Properties.SheetId.Value;
            }

            //Turn cells into squares
            body = new BatchUpdateSpreadsheetRequest()
            {
                Requests = new List<Request>() {
                    new Request()
                    {
                        UpdateDimensionProperties = new UpdateDimensionPropertiesRequest()
                        {
                            Fields = "*",
                            Properties = new DimensionProperties()
                            {
                                PixelSize = PixelSize,
                            },
                            Range = new DimensionRange()
                            {
                                SheetId = sheetId,
                                Dimension = "ROWS"
                            }

                        }
                    },
                    new Request()
                    {
                        UpdateDimensionProperties = new UpdateDimensionPropertiesRequest()
                        {
                            Fields = "*",
                            Properties = new DimensionProperties()
                            {
                                PixelSize = PixelSize,
                            },
                            Range = new DimensionRange()
                            {
                                SheetId = sheetId,
                                Dimension = "COLUMNS"
                            }
                        }
                    }
                },
            };
            var batchUpdate = service.Spreadsheets.BatchUpdate(body, SpreadsheetId);
            batchUpdate.Execute();

            //Turn image pixels into cells, per row
            List<Request> colorRequests = new List<Request>();
            for (int y = 0; y < image.Height; y++)
            {
                //Fill row data
                var rowData = new RowData();
                rowData.Values = new List<CellData>();
                for (int x = 0; x < image.Width; x++)
                {
                    Color color = image.GetPixel(x, y);
                    if (color.A == 0f)
                        color = Color.White;
                    rowData.Values.Add(new CellData()
                    {
                        UserEnteredFormat = new CellFormat()
                        {
                            BackgroundColor = new Google.Apis.Sheets.v4.Data.Color()
                            {
                                Red = color.R / 255f,
                                Green = color.G / 255f,
                                Blue = color.B / 255f,
                                Alpha = 1
                            }
                        }
                    });
                }
                //Save as request
                colorRequests.Add(new Request()
                {
                    UpdateCells = new UpdateCellsRequest()
                    {
                        Fields = "*",
                        Range = new GridRange()
                        {
                            SheetId = sheetId,
                            StartRowIndex = y,
                            EndRowIndex = y + 1,
                        },
                        Rows = new List<RowData>() { rowData },
                    }
                });
            }

            //Push color updates
            for (int req = 0; req < colorRequests.Count; req+=RowsPerUpdate)
            {
                body = new BatchUpdateSpreadsheetRequest()
                {
                    Requests = new List<Request>()
                };
                for (int i = 0; i < RowsPerUpdate; i++)
                {
                    if (req+i >= colorRequests.Count) break;
                    body.Requests.Add(colorRequests[req + i]);
                }
                
                batchUpdate = service.Spreadsheets.BatchUpdate(body, SpreadsheetId);
                batchUpdate.Execute();
            }
        }

        static string ColumnToLetter(int column)
        {
            int temp;
            string letter = "";
            while (column > 0)
            {
                temp = (column - 1) % 26;
                letter = Convert.ToChar(temp + 65) + letter;
                column = (column - temp - 1) / 26;
            }
            return letter;
        }

        static void SetupSheetsService()
        {
            //See .NET Quickstart: https://developers.google.com/sheets/api/quickstart/dotnet
            GoogleCredential credential;
            using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(Scopes);
            }

            service = new SheetsService(new Google.Apis.Services.BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
            Console.WriteLine("Initialized Google Sheets Service");
        }
    }
}