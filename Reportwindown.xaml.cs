using System;
using System.Data;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Data.SqlClient; // 🟢 เชื่อมต่อฐานข้อมูลระบบ MSSQL ยุคใหม่
using ClosedXML.Excel;          // 🟢 เรียกใช้งาน ClosedXML สรรค์สร้าง Excel แท้ (.xlsx)

namespace BWG_WasteScraper
{
    public partial class Reportwindown : Window
    {
        // ⚠️ รบกวนพี่ยังปรับ Connection String ให้ตรงกับพิกัดเซิร์ฟเวอร์จริงของพี่นะครับ
        string? _connString = AppSettings.ConnectionString;

        // ถังพักข้อมูลดาต้าเบสชั่วคราว เพื่อส่งต่อให้ปุ่ม Export ทำงานได้ทันที
        private DataTable? _dtCurrentReport = null;

        public Reportwindown()
        {
            InitializeComponent();
        }

        // ====================================================================
        // 🔍 ปุ่มกดค้นหาข้อมูลตามตัวกรองหลัก (BtnSearch_Click)
        // ====================================================================
        private async void BtnSearch_Click(object sender, RoutedEventArgs e)
        {
            BtnSearch.IsEnabled = false;
            TxtReportStatus.Text = "⏳ กำลังดึงข้อมูลจาก Database และกรองรายการ...";
            DgPreview.ItemsSource = null; // เคลียร์ตารางเดิมบนหน้าจอ

            // 🎯 ดักจับค่าตัวแปรจากหน้ากาก UI ตัวกรองหน้าแรก
            string selectedCompanyId = (CboCompany.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "";
            string requestNumberInput = TxtReqNumber.Text.Trim().Replace("'", "''");
            string selectedStatus = (CboStatus.SelectedItem as ComboBoxItem)?.Tag?.ToString() ?? "";
            // 1. ดึงค่าจากหน้าจอมาพักไว้ตามปกติของพี่
            DateTime? dateFrom = DpFrom.SelectedDate;
            DateTime? dateTo = DpTo.SelectedDate;

            // ตัวแปรสำหรับเอาไปสวมใน SQL Query (แปลงเป็น String รูปแบบ ค.ศ. เสมอ)
            string? sqlDateFrom = null;
            string? sqlDateTo = null;

            // 🎯 2. ระบบตรวจสอบและดัดหลังปี พ.ศ. ให้กลับเป็น ค.ศ. สำหรับกล่อง "ตั้งแต่วันที่" (dateFrom)
            if (dateFrom.HasValue)
            {
                int year = dateFrom.Value.Year;

                // ถ้าปีมากกว่า 2500 แสดงว่าเป็น พ.ศ. แน่นอน -> จับหักลบ 543 ให้คืนชีพเป็น ค.ศ.
                if (year >= 2500)
                {
                    year = year - 543;
                }

                // ประกอบร่างวันที่ใหม่โดยยึดปี ค.ศ. (yyyy) ที่ผ่านการกรองแล้ว ปลอดภัย 100%
                sqlDateFrom = $"{year}-{dateFrom.Value:MM-dd} 00:00:00";
            }

            // 🎯 3. ระบบตรวจสอบและดัดหลังปี พ.ศ. ให้กลับเป็น ค.ศ. สำหรับกล่อง "ถึงวันที่" (dateTo)
            if (dateTo.HasValue)
            {
                int year = dateTo.Value.Year;

                // ดักเช็กเช่นกัน ถ้าเป็น พ.ศ. ให้แปลงกลับเป็น ค.ศ.
                if (year >= 2500)
                {
                    year = year - 543;
                }

                sqlDateTo = $"{year}-{dateTo.Value:MM-dd} 23:59:59";
            }

            await Task.Run(async () =>
            {
                try
                {
                    // 1. ต่อคำสั่ง SQL ยึดตามคอลัมน์ HD และ DT ที่เราเคลียร์ปัญหากล่องแฝดเรียบร้อย
                    StringBuilder sqlBuilder = new StringBuilder();
                    sqlBuilder.Append(@"
                        SELECT
                        d.FactoryRegistrationNumber AS เลขทะเบียนโรงงานผู้ก่อกำเนิด, 
                        d.FactoryName AS ชื่อโรงงาน,
                        d.BusinessOperation AS ประกอบกิจการ,
                        d.[Address] AS ที่ตั้งโรงงาน,
                        d.LicenseeName AS ชื่อผู้รับใบอนุญาต,
                        h.SubmissionDate AS วันที่ยื่นขอ,
                        d.[Year] AS ปี,
                        h.RequestType AS ประเภทคำขอ,
                        h.RequestNumber AS เลขที่คำขอ,
                        h.[Status] AS สถานะคำขอหลัก,
                        d.SequenceNumber AS ลำดับย่อย,
                        d.WasteTypeCode AS รหัสประเภทหรือชนิดของเสีย,
                        d.WasteName AS ชื่อของเสีย,
                        d.QuantityMetricTons AS [ปริมาณ(ตัน)],    
                        d.ManagementCode AS รหัสการจัดการ,
	                    d.OperatorCode AS เลขทะเบียนโรงงานผู้รับบริการ
                        FROM tbWasteScraperHD h
                        LEFT JOIN tbWasteScraperDT d ON h.RequestNumber = d.RequestNumber
                        WHERE h.IsCheck is not null"); // เลือกดึงเฉพาะตัวที่เรากวาดประวัติสมบูรณ์แล้ว

                    // 2. แตกแขนงเงื่อนไข Dynamic Filter (ยึดตามข้อมูลหน้าแรก)
                    if (!string.IsNullOrEmpty(selectedCompanyId))
                    {
                        sqlBuilder.Append($" AND  h.CompanyCode LIKE '{selectedCompanyId}%' ");
                    }
                    if (!string.IsNullOrEmpty(requestNumberInput))
                    {
                        sqlBuilder.Append($" AND  h.RequestNumber LIKE '%{requestNumberInput}%' ");
                    }
                    if (!string.IsNullOrEmpty(selectedStatus))
                    {
                        sqlBuilder.Append($" AND  h.[RequestType] = N'{selectedStatus}' ");
                    }
                    if (dateFrom.HasValue)
                    {
                        sqlBuilder.Append($" AND  h.SubmissionDate >= '{sqlDateFrom}' ");
                    }
                    if (dateTo.HasValue)
                    {
                        sqlBuilder.Append($" AND  h.SubmissionDate <= '{sqlDateTo}' ");
                    }

                    // จัดเรียงลำดับให้สวยงามตามคิวหน้าระบบเว็บ
                    sqlBuilder.Append(" ORDER BY h.SubmissionDate ASC, d.SequenceNumber ASC ");

                    // 3. ยิงข้อมูลฝังรากลงใน DataTable ตัวแปรส่วนกลาง
                    _dtCurrentReport = new DataTable();
                    using (SqlConnection conn = new SqlConnection(_connString))
                    {
                        await conn.OpenAsync();
                        using (SqlCommand cmd = new SqlCommand(sqlBuilder.ToString(), conn))
                        {
                            using (SqlDataAdapter da = new SqlDataAdapter(cmd))
                            {
                                da.Fill(_dtCurrentReport);
                            }
                        }
                    }

                    // 4. คืนชีพหน้าต่างกลับสู่ UI Thread เพื่ออัปเดตตาราง Preview ให้พนักงานตรวจงาน
                    Dispatcher.Invoke(() =>
                    {
                        DgPreview.ItemsSource = _dtCurrentReport.DefaultView;
                        TxtReportStatus.Text = $"✅ ดึงสำเร็จ! สแกนพบข้อมูลเข้าเงื่อนไขทั้งหมด {_dtCurrentReport.Rows.Count} รายการย่อย";
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => TxtReportStatus.Text = $"❌ คิวรี่ขัดข้อง: {ex.Message}");
                }
                finally
                {
                    Dispatcher.Invoke(() => BtnSearch.IsEnabled = true);
                }
            });
        }

        // ====================================================================
        // 📊 ปุ่มสั่งแปลงและส่งออกรายงานเป็นไฟล์ Excel (.xlsx) แท้
        // ====================================================================
        private async void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            // เช็กระบบความปลอดภัย ป้องกัน User มือกดปุ่มรายงานทั้ง ๆ ที่ยังไม่กดค้นหาข้อมูล
            if (_dtCurrentReport == null || _dtCurrentReport.Rows.Count == 0)
            {
                MessageBox.Show("กรุณากดปุ่ม 🔍 ค้นหา ข้อมูลตามตัวกรองก่อนสั่งออกไฟล์ Excel ครับพี่ยัง", "แจ้งเตือนระบบ", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            BtnExportExcel.IsEnabled = false;
            TxtReportStatus.Text = "⏳ กำลังสลักตัวอักษรลงชีตและจัดรูปเล่มไฟล์ Excel...";

            await Task.Run(async () =>
            {
                try
                {
                    // ตั้งค่าระบุตำแหน่งเซฟไฟล์รายงานไปโผล่ที่หน้าจอ Desktop อัตโนมัติ
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string fileName = $"BWG_WasteReport_{timestamp}.xlsx";
                    string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    string filePath = Path.Combine(desktopPath, fileName);

                    // 🎯 ม้วนเสื่อเข้าสู่สเต็ปแปลงข้อมูลด้วย ClosedXML (ฟรี 100% ชัวร์ขาดลอย)
                    using (var workbook = new XLWorkbook())
                    {
                        // สับค่าคว่ำดาต้าเบสลงหน้าชีตรายงาน พร้อมแถมพิมพ์หัวข้อแถวให้ฟรี
                        var worksheet = workbook.Worksheets.Add(_dtCurrentReport, "รายงานระบบกรองพิเศษ");

                        // 🎨 ตกแต่งหัวตาราง (แถวที่ 1) ให้สวยหรู สไตล์รายงานผู้บริหาร
                        var headerRange = worksheet.Row(1).CellsUsed();
                        headerRange.Style.Font.Bold = true; // อักษรหนา
                        headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#007ACC"); // แถบพื้นหลังน้ำเงินสากล
                        headerRange.Style.Font.FontColor = XLColor.White; // ตัวอักษรขาวตัดเส้นชัดเจน

                        // 📐 รันระบบปรับความกว้างคอลัมน์อัตโนมัติ (AutoFit) ให้พอดีตัวหนังสือภาษาไทย
                        worksheet.Columns().AdjustToContents();

                        // สั่ง Save ยัดข้อมูลจริงลงหน้า Desktop
                        workbook.SaveAs(filePath);
                    }

                    Dispatcher.Invoke(() =>
                    {
                        MessageBox.Show($"📊 ส่งออกรายงาน Excel สำเร็จเรียบร้อยครับพี่ยัง!\nไฟล์ตั้งตระหง่านอยู่บน Desktop ชื่อไฟล์:\n{fileName}", "ระบบทำงานสำเร็จ", MessageBoxButton.OK, MessageBoxImage.Information);
                        TxtReportStatus.Text = $"✅ ส่งออกรายงานสำเร็จ! ชื่อไฟล์: {fileName}";
                    });
                }
                catch (Exception ex)
                {
                    Dispatcher.Invoke(() => TxtReportStatus.Text = $"❌ ไม่สามารถผลิตไฟล์ Excel ได้: {ex.Message}");
                }
                finally
                {
                    Dispatcher.Invoke(() => BtnExportExcel.IsEnabled = true);
                }
            });
        }


    }
}