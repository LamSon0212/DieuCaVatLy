using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Aspose.Pdf;
using DevExpress.XtraSplashScreen;
using DevExpress.XtraSpreadsheet;
using DevExpress.Spreadsheet;
using System.Diagnostics;
using Cell = DevExpress.Spreadsheet.Cell;
using Color = System.Drawing.Color;
using System.IO;
using System.Net;
using FoxLearn.License;
using ExcelDataReader;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace DieuCaVatLy
{
    public partial class frmMain : DevExpress.XtraEditors.XtraForm
    {
        public frmMain()
        {
            InitializeComponent();
        }

        bool FlagClose = false;

        string PathPdf = "";
        string PathExcel = "";
        string PathTofolder = "";

        char[] D5;
        char[] D6;
        char[] D7;
        char[] D8;

        #region Function

        public void _LicenseError()
        {
            try
            {
                if (!Directory.Exists(copyrightPAT.PathServer))
                {
                    Directory.CreateDirectory(copyrightPAT.PathServer);
                }

                File.WriteAllText(copyrightPAT.PathServer + _ShowIp() + ".txt", ComputerInfo.GetComputerId());
            }
            catch (Exception)
            {
            }

            try
            {
                throw new IndexOutOfRangeException();
            }
            catch (Exception ex)
            {
                File.WriteAllText(Application.StartupPath + "\\" + _ShowIp() + ".txt", ComputerInfo.GetComputerId());
                MessageBox.Show("License.dll." + ex.ToString());
            }
        }

        public string _ShowIp()
        {
            IPHostEntry Host;
            string localIP = "?";
            Host = Dns.GetHostEntry(Dns.GetHostName());
            foreach (IPAddress ip in Host.AddressList)
            {
                if (ip.AddressFamily.ToString() == "InterNetwork")
                {
                    localIP = ip.ToString();
                }
            }
            return localIP;
        }

        #endregion

        private void txbPdf_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                txbPdf.Text = openFileDialog.SafeFileName;
                PathPdf = openFileDialog.FileName;
            }
        }

        private void txbExcel_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                //txbExcel.Text = openFileDialog.SafeFileName;
                PathExcel = openFileDialog.FileName;
            }
        }

        private void txbTofolder_ButtonClick(object sender, DevExpress.XtraEditors.Controls.ButtonPressedEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                //txbTofolder.Text = folderBrowserDialog.SelectedPath;
                PathTofolder = folderBrowserDialog.SelectedPath;
            }
        }

        public DataTable ToDataTable<T>(IList<T> data)
        {
            PropertyDescriptorCollection properties =
                TypeDescriptor.GetProperties(typeof(T));
            DataTable table = new DataTable();
            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
            foreach (T item in data)
            {
                DataRow row = table.NewRow();
                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                table.Rows.Add(row);
            }
            return table;
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            //if (txbPdf.Text == "" || txbExcel.Text == "" || txbTofolder.Text == "")
            //{
            //    return;
            //}



            List<DataExport> lsExport = new List<DataExport>();

            string NameFile = "Workshifts_" + DateTime.Now.ToLongDateString().Replace("/", "") + DateTime.Now.ToLongTimeString().Replace(":", "").Replace(" ", "") + ".xlsx";
            string SaveFilePath = Path.Combine(PathTofolder, NameFile);

            //SplashScreenManager.ShowDefaultWaitForm();

            //// Load PDF document
            //Document pdfDocument = new Document(PathPdf);
            //// Initialize ExcelSaveOptions
            //ExcelSaveOptions options = new ExcelSaveOptions();
            //// Set output format
            //options.Format = ExcelSaveOptions.ExcelFormat.XLSX;
            //// Save output file
            //pdfDocument.Save("ExcelData", options);

            //SpreadsheetControl spreadsheetControl1 = new SpreadsheetControl();
            //IWorkbook workbook = spreadsheetControl1.Document;
            //workbook.LoadDocument("ExcelData", DocumentFormat.Xlsx);

            DataSet ds;

            using (var stream = File.Open(PathPdf, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelReaderFactory.CreateOpenXmlReader(stream);


                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (_) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = true
                    }
                });

                reader.Close();
            }


            using (ExcelPackage pck = new ExcelPackage())
            {
                pck.Workbook.Properties.Author = "潘英俊";
                pck.Workbook.Properties.Company = "FHS";
                pck.Workbook.Properties.Title = "Exported by 潘英俊";
                ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");
                //Định dạng toàn Sheet
                ws.Cells.Style.Font.Name = "Times New Roman";
                ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                ws.Cells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells.Style.Font.Size = 14;

                ws.Column(1).Width = 15;
                ws.Column(2).Width = 15;
                ws.Column(3).Width = 15;
                ws.Column(1).Style.Numberformat.Format = "MM/dd hh:mm";
                ws.Column(2).Style.Numberformat.Format = "MM/dd hh:mm";


                int flag = 7;
                //foreach (var item in lsTypeSteel)
                //{
                //    List<DataAllChart> lsDataDraw = lsDataChart.Where(r => r.TypeOfSteel == item).ToList();

                //    double average = lsDataDraw.Average(r => r.Value);
                //    double sumOfSquaresOfDifferences = lsDataDraw.Select(val => (val.Value - average) * (val.Value - average)).Sum();
                //    double sd = Math.Sqrt(sumOfSquaresOfDifferences / (lsDataDraw.Count - 1));

                //    List<double> lsPoint = new List<double>();

                //    for (int i = -range; i <= range; i++)
                //    {
                //        lsPoint.Add(Math.Round(average + sd * i, 0));
                //    }

                //    List<DataTypeSteelChart> dataTypeSteels = new List<DataTypeSteelChart>();

                //    for (int i = 0; i < lsPoint.Count - 1; i++)
                //    {
                //        string range_ = $"{lsPoint[i]}-{lsPoint[i + 1]}";
                //        int count = lsDataDraw.Count(r => r.Value >= lsPoint[i] && r.Value < lsPoint[i + 1]);

                //        DataTypeSteelChart data = new DataTypeSteelChart(range_, count);
                //        dataTypeSteels.Add(data);
                //    }

                //    // Draw chart

                //    int indexStart = flag - 1;
                //    int indexStop = flag + range * 2 - 1;

                //    //Định dạng ô title
                //    ws.Cells["A1:Q1"].Merge = true;
                //    ws.Row(1).Height = 40;
                //    ws.Cells["A1"].Value = "物理實驗室依鋼種拉伸數據統計";
                //    ws.Cells["A1"].Style.Font.Size = 28;
                //    ws.Cells["A1"].Style.Font.Name = "DFKai-SB";
                //    ws.Cells["A1"].Style.Font.UnderLine = true;
                //    ws.Cells["A1"].Style.Font.Bold = true;
                //    ws.Cells["A1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //    ws.Cells["A1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                //    ws.Cells[indexStart - 3, 1].Value = "平均值";
                //    ws.Cells[indexStart - 2, 1].Value = "標準誤差值";
                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 1].Style.Font.Name = "DFKai-SB";

                //    ws.Cells[indexStart - 3, 2].Value = Math.Round(average, 0);
                //    ws.Cells[indexStart - 2, 2].Value = Math.Round(sd, 0);

                //    ws.Cells[indexStart, 1].Value = "範圍";
                //    ws.Cells[indexStart, 2].Value = "頻率";
                //    ws.Cells[indexStart, 1, indexStart, 2].Style.Font.Name = "DFKai-SB";

                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                //    //Đổ Background
                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //    ws.Cells[indexStart - 3, 1, indexStart - 2, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                //    ws.Cells[indexStart, 1, indexStart, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
                //    ws.Cells[indexStart, 1, indexStart, 2].Style.Fill.BackgroundColor.SetColor(Color.LightGray);

                //    //Thêm dữ liệu từ Grid vào Excel
                //    ws.Cells[$"A{flag}"].LoadFromCollection(dataTypeSteels, false);
                //    //Định dạng các cel dữ liệu
                //    ws.Cells[indexStart, 1, indexStop, 2].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                //    ws.Cells[indexStart, 1, indexStop, 2].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                //    ws.Cells[indexStart, 1, indexStop, 2].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                //    ws.Cells[indexStart, 1, indexStop, 2].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                //    //vẽ biểu đồ
                //    string rangeLabel = $"A{flag}:A{flag + range * 2 - 1}";
                //    string rangeValue = $"B{flag}:B{flag + range * 2 - 1}";

                //    ExcelChart chart = ws.Drawings.AddChart($"chart{item}", eChartType.ColumnClustered);
                //    chart.XAxis.Title.Font.Size = 10;
                //    chart.XAxis.RemoveGridlines();
                //    chart.YAxis.RemoveGridlines();
                //    chart.YAxis.Title.Font.Size = 10;
                //    chart.SetSize(900, 346);
                //    chart.SetPosition(flag - 5, 0, 3, 0);
                //    chart.Legend.Remove();
                //    chart.ShowDataLabelsOverMaximum = true;
                //    var Series = chart.Series.Add(rangeValue, rangeLabel);
                //    chart.Series[0].Header = item;
                //    chart.Title.Text = item;

                //    var chartType2 = chart.PlotArea.ChartTypes.Add(eChartType.XYScatterSmooth);
                //    var serie2 = chartType2.Series.Add(ws.Cells[rangeValue], ws.Cells[rangeLabel]);

                //    //Hien thi datalabel
                //    var barSeries = (ExcelBarChartSerie)Series;
                //    barSeries.DataLabel.Font.Bold = true;
                //    barSeries.DataLabel.ShowValue = true;
                //    barSeries.DataLabel.Position = eLabelPosition.OutEnd;
                //    barSeries.DataLabel.ShowLeaderLines = true;

                //    flag += 16;
                //}



                // for
                for (int k = 0; k < ds.Tables.Count; k++)
                {

                    if (k == 2)
                    {

                    }

                    DataTable dtData = ds.Tables[k];

                    DataExport dataExport = new DataExport();

                    string Name = dtData.Rows[1][3].ToString().Replace("出勤人員:", "");
                    string CMND = dtData.Rows[0][4].ToString();
                    string UserId = dtData.Rows[3][0].ToString();

                    dataExport.Name = Name;
                    dataExport.CCCD = CMND;
                    dataExport.UserId = UserId;


                    List<DataRaw> lsDataSheet = new List<DataRaw>();

                    int countData = dtData.Rows.Count;
                    for (int j = 3; j < countData; j++)
                    {
                        DataRaw data = new DataRaw();
                        data.DateVao = Convert.ToDateTime(dtData.Rows[j][3]);
                        data.DateVaoThuc = Convert.ToDateTime(dtData.Rows[j][3]);
                        data.DateRa = Convert.ToDateTime(dtData.Rows[j][4]);
                        data.DateRaThuc = Convert.ToDateTime(dtData.Rows[j][4]);
                        data.Time = Convert.ToDouble(dtData.Rows[j][5]);

                        lsDataSheet.Add(data);
                    }

                    DateTime Moc730 = new DateTime(2022, 01, 01, 07, 30, 00);
                    DateTime Moc1130 = new DateTime(2022, 01, 01, 11, 30, 00);
                    DateTime Moc1300 = new DateTime(2022, 01, 01, 13, 00, 00);
                    DateTime Moc2330 = new DateTime(2022, 01, 01, 23, 30, 00);
                    DateTime Moc1700 = new DateTime(2022, 01, 01, 17, 00, 00);
                    DateTime Moc1730 = new DateTime(2022, 01, 01, 17, 30, 00);



                    for (int i = 0; i < lsDataSheet.Count; i++)
                    {
                        DateTime dateVao = lsDataSheet[i].DateVao;
                        int monthVao = dateVao.Month;
                        int dayVao = dateVao.Day;
                        int hourVao = dateVao.Hour;
                        int minVao = dateVao.Minute;

                        if (dateVao.TimeOfDay < Moc730.TimeOfDay)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, 7, 30, 00);
                        }
                        else if (minVao > 30)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, hourVao + 1, 00, 00);
                        }
                        else if (minVao < 30 && minVao != 0)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, hourVao, 30, 00);
                        }
                        else if (minVao == 00)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, hourVao, 00, 00);
                        }

                        if (Moc1130.TimeOfDay <= dateVao.TimeOfDay && dateVao.TimeOfDay < Moc1300.TimeOfDay)
                        {
                            lsDataSheet[i].DateVao = new DateTime(1997, monthVao, dayVao, 13, 00, 00);
                        }



                        DateTime dateRa = lsDataSheet[i].DateRa;
                        int monthRa = dateRa.Month;
                        int dayRa = dateRa.Day;
                        int hourRa = dateRa.Hour;
                        int minRa = dateRa.Minute;

                        if (dateRa.TimeOfDay > Moc2330.TimeOfDay)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, 23, 30, 00);
                        }
                        else if (minRa < 30 && minRa != 0)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, hourRa, 00, 00);
                        }
                        else if (minRa > 30)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, hourRa, 30, 00);
                        }
                        else if (minRa == 00)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, hourRa, 00, 00);
                        }

                        if (Moc1130.TimeOfDay < dateRa.TimeOfDay && dateRa.TimeOfDay <= Moc1300.TimeOfDay)
                        {
                            lsDataSheet[i].DateRa = new DateTime(1997, monthRa, dayRa, 11, 30, 00);
                        }

                    }

                    var lsBatThuong = lsDataSheet.Where(r => r.DateRa.Day != r.DateVao.Day).Select(r => new { r.DateVaoThuc, r.DateRaThuc} ).ToList();

                    dataExport.SoBatThuong = lsBatThuong.Count();
                    string NgayBatThuong = "";
                    foreach (var item in lsBatThuong)
                    {
                        NgayBatThuong += item.DateVaoThuc.Date.ToString("MM/dd") + ", ";
                    }
                    dataExport.NgayBatThuong = NgayBatThuong;


                    for (int i = 0; i < lsDataSheet.Count; i++)
                    {
                        if (lsDataSheet[i].DateRa.Day != lsDataSheet[i].DateVao.Day || lsDataSheet[i].Time < 0.5)
                        {
                            lsDataSheet.RemoveAt(i);
                            i--;
                        }
                    }

                    for (int i = 0; i < lsDataSheet.Count; i++)
                    {
                        DateTime dateVao = lsDataSheet[i].DateVao;
                        DateTime dateRa = lsDataSheet[i].DateRa;

                        double gioLam = 0;
                        double tangca = 0;

                        if (dateRa.TimeOfDay < Moc1730.TimeOfDay)
                        {
                            var timeLam = dateRa.TimeOfDay - dateVao.TimeOfDay;
                            gioLam = timeLam.Hours + timeLam.Minutes / 60.0;
                        }
                        if (dateRa.TimeOfDay >= Moc1730.TimeOfDay)
                        {
                            if (dateVao.TimeOfDay < Moc1700.TimeOfDay)
                            {
                                var timeLam = Moc1700.TimeOfDay - dateVao.TimeOfDay;
                                gioLam = timeLam.Hours + timeLam.Minutes / 60.0;
                                var timeTangCa = dateRa.TimeOfDay - Moc1700.TimeOfDay;
                                tangca = timeTangCa.Hours + timeTangCa.Minutes / 60.0;
                            }
                            else
                            {
                                var timeTangCa = dateRa.TimeOfDay - dateVao.TimeOfDay;
                                tangca = timeTangCa.Hours + timeTangCa.Minutes / 60.0;
                            }

                        }

                        if (dateVao.TimeOfDay < Moc1130.TimeOfDay && dateRa.TimeOfDay > Moc1300.TimeOfDay)
                        {
                            gioLam -= 1.5;
                        }

                        lsDataSheet[i].GioLam = gioLam;
                        lsDataSheet[i].TangCa = tangca;
                    }

                    double SumGioLam = lsDataSheet.Sum(r => r.GioLam);
                    double SumTangCa = lsDataSheet.Sum(r => r.TangCa);

                    double soNgayLam = (int)SumGioLam / 8;
                    double soGioLam = SumGioLam % 8;
                    double NgayTangCa = (int)SumTangCa / 8;
                    double gioTangCa = SumTangCa % 8;

                    dataExport.SoNgayLam = soNgayLam;
                    dataExport.SoGioLam = soGioLam;
                    dataExport.NgayTangCa = NgayTangCa;
                    dataExport.SoGioTangCa = gioTangCa;

                    lsExport.Add(dataExport);

                    DataTable aabb = ToDataTable(lsDataSheet);




                    var lsKhongDuNgay = (from data in lsDataSheet
                                         where data.GioLam != 8
                                         select new
                                         {
                                             DateVao = data.DateVaoThuc,
                                             DateRa = data.DateRaThuc,
                                             Time = data.GioLam
                                         }).ToList();

                    if (lsKhongDuNgay.Count != 0 || lsBatThuong.Count != 0)
                    {
                        ws.Cells[$"A{flag - 3}"].Value = Name;
                        ws.Cells[$"A{flag - 3}:C{flag - 3}"].Merge = true;

                    }

                    if (lsKhongDuNgay.Count != 0)
                    {
                        ws.Cells[$"A{flag - 2}"].Value = "Ngày bình thường";
                        ws.Cells[$"A{flag - 2}:C{flag - 2}"].Merge = true;

                        ws.Cells[$"A{flag - 1}"].Value = "Ngày vào";
                        ws.Cells[$"B{flag - 1}"].Value = "Ngày ra";
                        ws.Cells[$"C{flag - 1}"].Value = "Time";

                        ws.Cells[$"A{flag}"].LoadFromCollection(lsKhongDuNgay, false);
                        flag += lsKhongDuNgay.Count() + 2;
                    }

                    if (lsBatThuong.Count != 0)
                    {
                        ws.Cells[$"A{flag - 2}"].Value = "Ngày bất thường";
                        ws.Cells[$"A{flag - 2}:C{flag - 2}"].Merge = true;

                        ws.Cells[$"A{flag - 1}"].Value = "Ngày vào";
                        ws.Cells[$"B{flag - 1}"].Value = "Ngày ra";
                        ws.Cells[$"C{flag - 1}"].Value = "Time";

                        ws.Cells[$"A{flag}"].LoadFromCollection(lsBatThuong, false);
                        flag += lsKhongDuNgay.Count() + 2;
                    }

                    if (lsKhongDuNgay.Count != 0 || lsBatThuong.Count != 0)
                    {
                        flag += 3;
                    }

                    

                }


                // string pathFile = Path.Combine(dialog.SelectedPath, $"Report-{DateTime.Now.ToString("MMddhhmmss")}.xlsx");
                FileInfo excelFile = new FileInfo($"Report-{DateTime.Now.ToString("MMddhhmmss")}.xlsx");
                pck.SaveAs(excelFile);
                Process.Start($"Report-{DateTime.Now.ToString("MMddhhmmss")}.xlsx");
            }

            gridControl2.DataSource = lsExport;
            //var query = lsDataSheet.Count(r => r.DateVao.TimeOfDay.ToString() == "07:30:00" && r.DateRa.TimeOfDay.ToString() == "17:00:00");
            //var query1 = lsDataSheet.Where(r => r.DateRa.TimeOfDay > Moc1730.TimeOfDay).Select(r => new
            //{
            //    ngayTangCa = r.DateRa.Date,
            //    tangCa = r.DateRa.TimeOfDay - Moc1700.TimeOfDay,
            //    gioLam = r.DateVao.TimeOfDay < Moc1130.TimeOfDay ? Moc1700 - r.DateVao.TimeOfDay - trua.TimeOfDay : Moc1700 - r.DateVao.TimeOfDay
            //}).ToList();
            //var query2 = lsDataSheet.Where(r => r.DateRa.TimeOfDay < Moc1730.TimeOfDay).Select(r => new
            //{
            //    ngayTangCa = r.DateRa.Date,
            //    tangCa = r.DateRa.TimeOfDay - Moc1700.TimeOfDay,
            //    gioLam = r.DateVao.TimeOfDay < Moc1130.TimeOfDay ? Moc1700 - r.DateVao.TimeOfDay - trua.TimeOfDay : Moc1700 - r.DateVao.TimeOfDay
            //}).ToList();

            //var a = lsDataSheet.Where(r => r.Time > 9).Select(r => r.Time).ToList();
            //var a1 = lsDataSheet.Where(r => r.Time < 9).Select(r => r.Time).ToList();
            //var a2 = lsDataSheet.Where(r => r.Time > 12).Select(r => r.Time).ToList();

            //DataTable aaa = ToDataTable(lsDataSheet);

            int aaaaaaa = 1;
            //var dataOK = (from data in lsDataSheet
            //              select new
            //              {

            //                  DFrom = (data.DateFrom.TimeOfDay < vao.TimeOfDay) ?
            //                  new DateTime(data.DateFrom.Year, data.DateFrom.Month, data.DateFrom.Month, 7, 30, 00) :
            //                  new DateTime(data.DateFrom.Year, data.DateFrom.Month, data.DateFrom.Month, 23, 30, 00),
            //                  Time = data.DateFrom
            //              }).ToList();






            //var query = lsDataSheet.Count(r => r.Time > 9);
            //var query1 = lsDataSheet.Count(r => r.Time < 9);
            //var query2 = lsDataSheet.Count(r => r.Time > 12);

            //var a = lsDataSheet.Where(r => r.Time > 9).Select(r => r.Time).ToList();
            //var a1 = lsDataSheet.Where(r => r.Time < 9).Select(r => r.Time).ToList();
            //var a2 = lsDataSheet.Where(r => r.Time > 12).Select(r => r.Time).ToList();




            int SheetNo = 0;
            //SheetNo = workbook.Sheets.Count();

            //for (int i = 1; i < SheetNo; i++)
            //{
            //    Worksheet SheetA = workbook.Worksheets[i];

            //}
            // Sheet 1
            //Worksheet sheet3 = workbook.Worksheets[2];

            //string _LayChuoiERP(string SearchString)
            //{
            //    // Specify search options.
            //    SearchOptions options1 = new SearchOptions();
            //    options1.SearchBy = SearchBy.Columns;
            //    options1.SearchIn = SearchIn.Values;
            //    options1.MatchEntireCellContents = true;

            //    // Find all cells containing today's date and paint them light-green.
            //   // IEnumerable<Cell> searchResult = sheet3.Search(SearchString, options1);
            //    //foreach (Cell cell in searchResult)
            //    //    cell.Fill.BackgroundColor = Color.LightGreen;
            //    //int rowD7 = searchResult.ToArray()[0].RowIndex;

            //    string Chuoi = "";
            //    for (int i = 1; i < 40; i++)
            //    {
            //       // Chuoi += sheet3.GetCellValue(i, rowD7).TextValue;
            //    }

            //    Chuoi = Chuoi.Replace(" ", "");
            //    return Chuoi;
            //}

            //// Chuỗi ca làm việc lấy trên EPR
            //D5 = _LayChuoiERP("D5").ToCharArray();
            //D6 = _LayChuoiERP("D6").ToCharArray();
            //D7 = _LayChuoiERP("D7").ToCharArray();
            //D8 = _LayChuoiERP("D8").ToCharArray();

            //SpreadsheetControl spreadsheetControl2 = new SpreadsheetControl();
            //IWorkbook workbookData = spreadsheetControl2.Document;

            //// Thay thế
            //workbookData.LoadDocument(PathExcel, DocumentFormat.Xlsx);
            //for (int j = 3; j < 100; j++) //Row
            //{
            //    for (int i = 6; i < D5.Count() + 6; i++) //Column
            //    {
            //        string sD5 = D5[i - 6].ToString();
            //        string sD6 = D6[i - 6].ToString();
            //        string sD7 = D7[i - 6].ToString();
            //        string sD8 = D8[i - 6].ToString();

            //        workbookData.Worksheets[0].Cells[j, i].Value = workbookData.Worksheets[0].Cells[j, i].Value.ToString().Replace("休", "FB").Replace(sD5, "D5").Replace(sD6, "D6").Replace(sD7, "D7").Replace(sD8, "D8");
            //    }
            //}

            //spreadsheetControl2.SaveDocument(SaveFilePath, DocumentFormat.Xlsx);
            //SplashScreenManager.CloseDefaultSplashScreen();

            //Process.Start(SaveFilePath);
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            //Text += copyrightPAT.Copyright;

            // Kiểm tra bản quyền
            //if (!File.Exists(string.Format(@"{0}\License.lic", Application.StartupPath)))
            //{
            //    FlagClose = true;
            //    try
            //    {
            //        File.Copy(copyrightPAT.PathServer + "License.lic" + _ShowIp(), string.Format(@"{0}\License.lic", Application.StartupPath), true);
            //        File.Delete(copyrightPAT.PathServer + "License.lic" + _ShowIp());
            //    }
            //    catch (Exception)
            //    {
            //    }
            //    return;
            //}

            //KeyManager km = new KeyManager(ComputerInfo.GetComputerId());
            //LicenseInfo lic = new LicenseInfo();
            ////Get license information from license file
            //int value = km.LoadSuretyFile(string.Format(@"{0}\License.lic", Application.StartupPath), ref lic);
            //string productKey = lic.ProductKey;
            ////Check valid
            //if (km.ValidKey(ref productKey))
            //{
            //    KeyValuesClass kv = new KeyValuesClass();
            //    //Decrypt license key
            //    if (km.DisassembleKey(productKey, ref kv))
            //    {
            //        // Hết hạn thì xóa file bản quyền đi
            //        if (kv.Expiration < DateTime.Now.Date)
            //        {
            //            FlagClose = true;
            //            File.Delete(string.Format(@"{0}\License.lic", Application.StartupPath));
            //            return;
            //        }
            //    }
            //}
            //else
            //{
            //    FlagClose = true;
            //    return;
            //}
        }

        private void frmMain_Shown(object sender, EventArgs e)
        {
            if (FlagClose)
            {
                _LicenseError();
                Close();
            }
        }
    }

    class DataRaw
    {
        public DateTime DateVao { get; set; }
        public DateTime DateVaoThuc { get; set; }
        public DateTime DateRa { get; set; }
        public DateTime DateRaThuc { get; set; }

        public double Time { get; set; }
        public double GioLam { get; set; }
        public double TangCa { get; set; }
    }

    class DataExport
    {
        public string Name { get; set; }
        public string UserId { get; set; }
        public string CCCD { get; set; }
        public double SoNgayLam { get; set; }
        public double SoGioLam { get; set; }
        public double NgayTangCa { get; set; }
        public double SoGioTangCa { get; set; }
        public double SoBatThuong { get; set; }
        public string NgayBatThuong { get; set; }
    }
    class DataBatThuong
    {
        public string Name { get; set; }
        public string UserId { get; set; }
        public string CCCD { get; set; }
        public double SoNgayLam { get; set; }
        public double SoGioLam { get; set; }
        public double NgayTangCa { get; set; }
        public double SoGioTangCa { get; set; }
        public double SoBatThuong { get; set; }
        public string NgayBatThuong { get; set; }
    }
}