using ClosedXML.Excel;
using Microsoft.Data.SqlClient; 
using Microsoft.Playwright;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;

namespace BWG_WasteScraper
{
    public partial class MainScraperWindow : Window
    {
        private List<string> exportData = new List<string>();
        private readonly string _targetUrl = "https://e-waste.diw.go.th/waste/authen/login.html";
        string? _connString = AppSettings.ConnectionString;
        public MainScraperWindow(string appUsername)
        {
            InitializeComponent();
            TxtWelcome.Text = $"ผู้ใช้งานแอป: {appUsername}";
        }

        private void UpdateLog(string message)
        {
            Dispatcher.Invoke(() =>
            {
                TxtLogs.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}\n");
                LogScrollViewer.ScrollToBottom();
            });
        }

        private void UpdateStatus(string statusText, string hexColor = "#00FF99")
        {
            Dispatcher.Invoke(() =>
            {
                TxtStatus.Text = statusText;
                TxtStatus.Foreground = (System.Windows.Media.Brush?)new System.Windows.Media.BrushConverter().ConvertFromString(hexColor)
                                       ?? System.Windows.Media.Brushes.Black;
            });
        }

        private async void BtnStart_Click(object sender, RoutedEventArgs e)
        {
            string webUser = TxtWebUsername.Text.Trim();
            string webPass = TxtWebPassword.Password.Trim();
            string selectedCompanyId = "";
            string selectedCompanyName = "";

            Dispatcher.Invoke(() =>
            {
                if (CboCompany.SelectedItem is ComboBoxItem selectedItem)
                {
                    selectedCompanyId = selectedItem.Tag?.ToString() ?? "";
                    selectedCompanyName = selectedItem.Content?.ToString() ?? "";
                }
            });

            if (string.IsNullOrWhiteSpace(webUser) || string.IsNullOrWhiteSpace(webPass) || string.IsNullOrWhiteSpace(selectedCompanyId))
            {
                MessageBox.Show("กรุณากรอกข้อมูลและเลือกบริษัทให้ครบถ้วนก่อนครับ", "แจ้งเตือน", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            BtnStart.IsEnabled = false;
            TxtLogs.Clear();
            exportData.Clear();
            UpdateStatus("⏳ บอทกำลังเริ่มทำงาน...", "#FFCC00");

            await Task.Run(async () =>
            {
                try
                {
                    // 🎯 สะสมชุดคำสั่ง SQL ทั้งหมด (ทั้งตาราง HD และ DT)
                    StringBuilder sqlBuilder = new StringBuilder();

                    UpdateLog("📦 กำลังตรวจสอบความพร้อมของ Chromium Browser ...");
                    Microsoft.Playwright.Program.Main(new[] { "install" });

                    UpdateLog("🚀 เริ่มต้นระบบ Playwright แบบซ่อนหน้าต่าง...");
                    using var playwright = await Playwright.CreateAsync();
                    await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = true });
                    var page = await browser.NewPageAsync();

                    UpdateLog($"🌐 นำทางไปยังหน้าเว็บไซต์: {_targetUrl}");
                    await page.GotoAsync(_targetUrl);

                    await page.Locator("#username").FillAsync(webUser);
                    await page.Locator("#password").FillAsync(webPass);
                    await page.Locator("#bttLogin").ClickAsync();

                    UpdateLog("⏳ รอระบบตรวจสอบรหัสผ่านหน้าเว็บ...");
                    await page.Locator("button:has-text('Log Out')").WaitForAsync(new() { Timeout = 60000 });
                    UpdateLog("✅ ล็อกอินสำเร็จ!");

                    UpdateLog($"🏢 ค้นหาแถวบริษัท: {selectedCompanyId} {selectedCompanyName} ...");
                    var targetRow = page.Locator("tr").Filter(new() { HasText = selectedCompanyId });
                    await targetRow.Locator("button:has-text('ดำเนินการ')").ClickAsync();

                    await page.Locator("a:has-text('1. ยืนยันความยินยอมรับดำเนินการ สิ่งปฏิกูลหรือวัสดุที่ไม่ใช้แล้ว')").ClickAsync();
                    await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                    var lengthSelect = page.Locator("select[name='waste_pro_table1_length']");
                    await lengthSelect.SelectOptionAsync(new[] { "100" });

                    var firstRowOfMyTable = page.Locator("#waste_pro_table1 tbody tr").First;
                    await firstRowOfMyTable.WaitForAsync(new()  { Timeout = 60000 , State = WaitForSelectorState.Visible });

                    // ====================================================================
                    // 📋 ดึงชื่อหัวข้อคอลัมน์จากหน้าแรกสุด
                    // ====================================================================
                    var mainThElements = await page.Locator("#waste_pro_table1 thead th").AllAsync();
                    var mainHeadersList = new List<string>();
                    foreach (var th in mainThElements)
                    {
                        string thText = await th.InnerTextAsync();
                        if (!string.IsNullOrWhiteSpace(thText)) mainHeadersList.Add($"\"{thText.Trim()}\"");
                    }
                    string mainHeaderStr = string.Join(",", mainHeadersList);

                    var allMainRows = page.Locator("#waste_pro_table1 > tbody > tr");
                    int rowCount = await allMainRows.CountAsync();
                    UpdateLog($"📡 เรดาร์พบข้อมูลในหน้าปัจจุบันทั้งหมด {rowCount} แถวหลัก เริ่มสกัดข้อมูล...");

                    // 🚀 ลูปหลัก (ตาราง HD)
                    for (int mainRowIdx = 0; mainRowIdx < rowCount; mainRowIdx++)
                    {
                        UpdateStatus($"⏳ กำลังดึงข้อมูลแถวที่ {mainRowIdx + 1} / {rowCount}...", "#FFCC00");
                        UpdateLog($"=======================================================================");
                        UpdateLog($"▶️ [เริ่มแถวหลักที่ {mainRowIdx + 1}]");

                        var targetMainRow = allMainRows.Nth(mainRowIdx);
                        var mainTds = targetMainRow.Locator("xpath=./td");
                        int mainTdCount = await mainTds.CountAsync();

                        if (mainTdCount <= 1) continue;

                        var mainValuesListtest = new List<string>();
                        for (int c = 0; c < mainHeadersList.Count && c < mainTdCount; c++)
                        {
                            string text = await mainTds.Nth(c).InnerTextAsync();
                            mainValuesListtest.Add($"\"{text.Replace("\n", " ").Replace("\r", "").Trim()}\"");
                        }
                        string firstPageDataStr = string.Join(",", mainValuesListtest);

                        // 🎯 ดึงข้อมูลหน้าแรกจัดเก็บลงตาราง tbWasteScraperHD (คงเดิมไว้ทั้งหมด)
                        int startPageValue = mainRowIdx + 1;
                        string requestNumber = (await mainTds.Nth(1).InnerTextAsync()).Trim();
                        string requestType = (await mainTds.Nth(2).InnerTextAsync()).Trim();
                        string rawSubmissionDate = (await mainTds.Nth(3).InnerTextAsync()).Trim();
                        string statusValue = (await mainTds.Nth(4).InnerTextAsync()).Trim();

                        string safeReqNum = requestNumber.Replace("'", "''");
                        requestType = requestType.Replace("'", "''");
                        statusValue = statusValue.Replace("'", "''");

                        string formattedDateForSql = "GETDATE()";
                        if (DateTime.TryParse(rawSubmissionDate, out DateTime parsedDate))
                        {
                            formattedDateForSql = $"'{parsedDate:yyyy-MM-dd HH:mm:ss}'";
                        }

                        // 🎯 1. ใช้ลอจิก IF NOT EXISTS ดักเช็กในระดับคำสั่ง SQL เพื่อไม่ให้มีการเก็บข้อมูลซ้ำถ้าเจอ RequestNumber เดิม
                        StringBuilder rowSqlBuilder = new StringBuilder();
                        rowSqlBuilder.AppendLine($"IF NOT EXISTS (SELECT 1 FROM tbWasteScraperHD WHERE RequestNumber = '{safeReqNum}' AND CompanyCode = '{selectedCompanyId}')");
                        rowSqlBuilder.AppendLine("BEGIN");
                        // ยิงบันทึก HD ลงคิว SQL
                        rowSqlBuilder.AppendLine(" INSERT INTO tbWasteScraperHD (StartPage, RequestNumber, RequestType, SubmissionDate, [Status], IsCheck, InsertDate,CompanyCode)");
                        rowSqlBuilder.AppendLine($" VALUES ({startPageValue}, '{safeReqNum}', N'{requestType}', {formattedDateForSql}, '{statusValue}', 'N', GETDATE() ,{selectedCompanyId});");
                        rowSqlBuilder.AppendLine("END");
                        rowSqlBuilder.AppendLine("ELSE");
                        rowSqlBuilder.AppendLine("BEGIN");
                        // 🔄 กรณีที่ 2: มีเลขคำขอนี้อยู่แล้ว -> ทำการอัปเดตสเตตัสพิกัดเดิมให้เป็น 'Y' ทันที
                        rowSqlBuilder.AppendLine($" UPDATE tbWasteScraperHD SET IsCheck = 'Y', [Status] = '{statusValue}' WHERE RequestNumber = '{safeReqNum}' AND CompanyCode = '{selectedCompanyId}' ;");
                        rowSqlBuilder.AppendLine("END");

                        // ยัดสคริปต์ฝั่ง HD ลงกล่องสะสมหลัก
                        sqlBuilder.Append(rowSqlBuilder.ToString());

                        // ====================================================================
                        // 🖱️ คลิกเปิดหน้าต่างเอกสารย่อย (Modal ตารางชั้นกลาง)
                        // ====================================================================
                        await RandomDelay(1500, 3000);
                        await mainTds.First.ClickAsync();

                        var targetModal = page.Locator(".modal-content:visible").Last;
                        try { await targetModal.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 12000 }); }
                        catch (TimeoutException) { UpdateLog($"❌ ข้ามรายการคำขอ {requestNumber} เนี่องจากโหลดหน้าต่างไม่ทัน"); continue; }

                        var innerTable = targetModal.Locator("table[id^='items_data_']").Last;

                        if (await innerTable.CountAsync() > 0)
                        {
                            var allRows = await page.Locator("#CleanerDataTable > table > tbody > tr").AllAsync();
                            string currentMainDataStr = "";
                            string innerHeaderStr = "";
                            bool isFullHeaderSaved = false;
                            int currentButtonIndex = 1;

                            // 🎯 ตัวนับลำดับ SequenceNumber สำหรับตาราง DT (เริ่มนับ 1 ใหม่ทุกคำขอหลัก)
                            int sequenceCounter = 1;

                            for (int i = 0; i < allRows.Count; i++)
                            {
                                var row = allRows[i];
                                var tds = row.Locator("xpath=./td");
                                int tdCount = await tds.CountAsync();

                                if (tdCount > 1)
                                {
                                    var mainValuesListtests = new List<string>();
                                    for (int c = 0; c < mainHeadersList.Count && c < tdCount; c++)
                                    {
                                        string text = await tds.Nth(c).InnerTextAsync();
                                        mainValuesListtests.Add($"\"{text.Replace("\n", " ").Replace("\r", "").Trim()}\"");
                                    }
                                    currentMainDataStr = string.Join(",", mainValuesListtests);
                                }
                                else if (tdCount == 1)
                                {
                                    var innerTables = row.Locator("table");
                                    if (await innerTables.CountAsync() > 0)
                                    {
                                        // 📋 [ของเดิมเป๊ะ] ระบบจัดการหัวข้อและจัดเก็บไฟล์ CSV (ห้ามแตะต้องเพื่อให้หัวข้ออยู่ครบ)
                                        if (string.IsNullOrEmpty(innerHeaderStr))
                                        {
                                            var innerThElementsM = await innerTables.Locator("thead th").AllAsync();
                                            var innerHeadersListM = new List<string>();
                                            foreach (var th in innerThElementsM) { innerHeadersListM.Add($"\"{th.InnerTextAsync().Result.Trim()}\""); }
                                            innerHeaderStr = string.Join(",", innerHeadersListM);
                                        }

                                        if (!isFullHeaderSaved && !string.IsNullOrEmpty(innerHeaderStr))
                                        {
                                            string combinedHeader = "หัวหน้าแรก_" + mainHeaderStr + "," + mainHeaderStr + "," + innerHeaderStr;
                                            if (exportData.Count == 0) { exportData.Insert(0, combinedHeader); }
                                            isFullHeaderSaved = true;
                                        }

                                        var innerDataRows = await innerTables.Locator("tbody tr").AllAsync();

                                        for (int j = 0; j < innerDataRows.Count; j++)
                                        {
                                            var innerTds = innerDataRows[j].Locator("xpath=./td");
                                            int innerTdCount = await innerTds.CountAsync();
                                            if (innerTdCount <= 1) continue;

                                            var innerValuesList = new List<string>();
                                            for (int c = 0; c < innerTdCount; c++)
                                            {
                                                string text = await innerTds.Nth(c).InnerTextAsync();
                                                innerValuesList.Add($"\"{text.Replace("\n", " ").Replace("\r", "").Trim()}\"");
                                            }
                                            string innerDataStr = string.Join(",", innerValuesList);

                                            // 📥 ดึงข้อมูลตารางชั้นกลาง มารองรับฟิลด์ DT (ผู้รับดำเนินการ, ประเภท, ปริมาณ, รหัสจัดการ, สิ้นสุดตอบรับ, สถานะย่อย)
                                            string operatorCode = (await innerTds.Nth(1).InnerTextAsync()).Trim();
                                            string operatorName = (await innerTds.Nth(2).InnerTextAsync()).Trim();
                                            string typeValue = (await innerTds.Nth(3).InnerTextAsync()).Trim();
                                            string rawQty = (await innerTds.Nth(4).InnerTextAsync()).Trim();
                                            string managementCode = (await innerTds.Nth(5).InnerTextAsync()).Trim();
                                            string rawDeadline = (await innerTds.Nth(6).InnerTextAsync()).Trim();
                                            string statusDt = (await innerTds.Nth(7).InnerTextAsync()).Trim();

                                            int buttonIndex = currentButtonIndex;
                                            var inspectBtn = targetModal.Locator($"button#btt_rd{buttonIndex}").First;

                                            if (await inspectBtn.IsVisibleAsync())
                                            {
                                                await inspectBtn.ScrollIntoViewIfNeededAsync();
                                                await inspectBtn.ClickAsync(new() { Force = true });
                                                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                                                // เจาะทะลวงเข้าสู่หน้าต่าง Modal ชั้นลึกสุด
                                                var detailModal = page.Locator(".modal-content:visible").Last;
                                                var firstInputInModal = detailModal.Locator("input").First;
                                                await firstInputInModal.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 60000 });

                                                // ====================================================================
                                                // ✨ หน้าต่างลึกสุดกางเต็มจอสำเร็จ! เริ่มสกัดข้อมูลรายกล่องคู่คอลัมน์...
                                                // ====================================================================
                                                var detailHeadersList = new List<string>();
                                                var detailValuesList = new List<string>();

                                                // 🎯 ประกาศตัวแปรรับค่าฟิลด์ต่าง ๆ เตรียมไว้ (คงเดิม)
                                                string yearVal = "0"; string facRegNo = ""; string facName = ""; string bizOp = "";
                                                string addr = ""; string licensee = ""; string process = ""; string tax = ""; string phone = ""; string fax = "";
                                                string itemNo = "0"; string wasteCode = ""; string haz = ""; string reason = "";
                                                // 🎯 1. เพิ่มตัวแปรพระเอกคู่ใหม่ สำหรับดักจับกล่องฝาแฝด
                                                string prop = "";     // สำหรับเก็บค่ากล่อง "ชื่อสิ่งปฏิกูลฯ" ตัวแรก (Properties)
                                                string wasteName = ""; // สำหรับเก็บค่ากล่อง "ชื่อสิ่งปฏิกูลฯ" ตัวที่สอง (WasteName)
                                                int twinNameCounter = 0; // ตัวนับว่าเจอคำว่า "ชื่อสิ่งปฏิกูล" กี่ครั้งแล้วในหน้านี้
                                                var labels = await detailModal.Locator("label").AllAsync();
                                                foreach (var lbl in labels)
                                                {
                                                    string headerText = await lbl.InnerTextAsync();
                                                    headerText = headerText.Replace("\n", " ").Replace("\r", "").Trim();
                                                    if (string.IsNullOrWhiteSpace(headerText) || headerText.Contains("เอกสารประกอบ")) continue;

                                                    var input = lbl.Locator("xpath=..//input | ..//textarea | ..//select").First;
                                                    if (await input.CountAsync() > 0)
                                                    {
                                                        var isInTable = input.Locator("xpath=ancestor::table[@id='ProduceTable']");
                                                        if (await isInTable.CountAsync() > 0) continue;

                                                        string val = await input.InputValueAsync();
                                                        val = val.Replace("\n", " ").Replace("\r", "").Trim();

                                                        // 📋 บันทึกเข้าไฟล์ CSV ของพี่ตามปกติเป๊ะ หัวข้อและข้อมูล CSV อยู่ครบ 100% ไม่พังแน่นอน
                                                        detailHeadersList.Add($"\"{headerText}\"");
                                                        detailValuesList.Add($"\"{val}\"");

                                                        // ====================================================================
                                                        // 🎯 2. สกัดค่าลงฐานข้อมูลรายตัว โดยใช้ระบบตรวจสอบคีย์เวิร์ด + นับจำนวนกล่องแฝด
                                                        // ====================================================================

                                                        if (headerText.Contains("ชื่อสิ่งปฏิกูล"))
                                                        {
                                                            twinNameCounter++; // นับคิวทันที (กล่องแรกได้ 1, กล่องสองได้ 2)
                                                            if (twinNameCounter == 1)
                                                            {
                                                                prop = val; // คิวแรก (ตัวบน) คือ คุณสมบัติ (Properties)
                                                            }
                                                            else if (twinNameCounter == 2)
                                                            {
                                                                wasteName = val; // คิวสอง (ตัวล่าง) คือ WasteName
                                                            }
                                                        }
                                                        else // กล่องอื่น ๆ ที่ชื่อไม่ซ้ำ ใช้สแกนหาคำคำใกล้เคียงตามปกติได้เลยครับพี่
                                                        {
                                                            if (headerText.Contains("ชื่อผู้รับใบอนูญาต")) licensee = val;
                                                            if (headerText.Contains("รายละเอียดของกิจกรรมที่ก่อให้เกิดของเสีย")) process = val;
                                                            if (headerText.Contains("ปี")) yearVal = val;
                                                            if (headerText.Contains("ทะเบียนโรงงาน")) facRegNo = val;
                                                            if (headerText.Contains("ชื่อโรงงาน")) facName = val;
                                                            if (headerText.Contains("ประกอบกิจการ")) bizOp = val;
                                                            if (headerText.Contains("ที่ตั้ง")) addr = val;
                                                            if (headerText.Contains("ผู้เสียภาษี")) tax = val;
                                                            if (headerText.Contains("โทรศัพท์")) phone = val;
                                                            if (headerText.Contains("โทรสาร")) fax = val;
                                                            if (headerText.Contains("รายการที่")) itemNo = val;
                                                            if (headerText.Contains("รหัสประเภท")) wasteCode = val;
                                                            if (headerText.Contains("HAZ")) haz = val;
                                                            if (headerText.Contains("เหตุผล")) reason = val;
                                                        }
                                                    }
                                                }

                                                // 📋 [ของเดิมเป๊ะ] จัดการรวมความกว้างบันทึกสายสตริงลงไฟล์สำรอง CSV 
                                                string detailHeaderStr = string.Join(",", detailHeadersList);
                                                string detailDataStr = string.Join(",", detailValuesList);
                                                string completeHeader = "หัวหน้าแรก_" + mainHeaderStr + "," + mainHeaderStr + "," + innerHeaderStr + "," + detailHeaderStr;
                                                if (exportData.Count == 0) { exportData.Add(completeHeader); }
                                                exportData.Add(firstPageDataStr + "," + currentMainDataStr + "," + innerDataStr + "," + detailDataStr);

                                                // ====================================================================
                                                // 🎯 [เพิ่มใหม่ยิงตรงคอลัมน์] ดึงจากตารางกลาง + Dictionary ลงตาราง tbWasteScraperDT
                                                // ====================================================================
                                                try
                                                {
                                                
                                                    // ปรับแต่งฟอร์แมตตัวเลขและวันที่ให้ปลอดภัยต่อโครงสร้าง MSSQL
                                                   
                                                    int.TryParse(yearVal, out int parsedYear);
                                                    int.TryParse(itemNo, out int parsedItemNo);
                                                
                                                    int LindedYear = 0; int.TryParse(yearVal, out LindedYear);
                                                    int LinkedItemNo = 0; int.TryParse(itemNo, out LinkedItemNo);
                                                    decimal.TryParse(rawQty.Replace(",", ""), out decimal qtyDecimal);

                                                    string formattedDeadlineSql = "NULL";
                                                    if (DateTime.TryParse(rawDeadline, out DateTime deadDate)) { formattedDeadlineSql = $"'{deadDate:yyyy-MM-dd HH:mm:ss}'"; }
                                                 
                                                    // 🎯 2. ใช้ลอจิก IF NOT EXISTS คลุมฝั่งตาราง DT ด้วยเช่นกัน เพื่อไม่ให้บันทึกข้อมูลย่อยซ้ำซ้อนซ่อนเงื่อน
                                                    StringBuilder dtSqlBuilder = new StringBuilder();
                                                    dtSqlBuilder.AppendLine($"IF NOT EXISTS (SELECT 1 FROM tbWasteScraperDT WHERE RequestNumber = '{safeReqNum}' AND SequenceNumber = {sequenceCounter})");
                                                    dtSqlBuilder.AppendLine("BEGIN");
                                                    dtSqlBuilder.AppendLine("    INSERT INTO tbWasteScraperDT (RequestNumber, SequenceNumber, OperatorCode, OperatorName, [Type], QuantityMetricTons, ManagementCode, AcknowledgementDeadline, [Status], [Year], FactoryRegistrationNumber, FactoryName, BusinessOperation, [Address], LicenseeName, TaxID, Phone, Fax, ItemNumber, WasteTypeCode, HazStatus, Properties, WasteName, WasteGenerationProcess, EvaluationReason)");
                                                    dtSqlBuilder.AppendLine($"    VALUES ('{safeReqNum}', {sequenceCounter}, '{operatorCode.Replace("'", "''")}', N'{operatorName.Replace("'", "''")}', N'{typeValue.Replace("'", "''")}', {qtyDecimal}, '{managementCode.Replace("'", "''")}', {formattedDeadlineSql}, '{statusDt.Replace("'", "''")}', {LindedYear}, '{facRegNo.Replace("'", "''")}', N'{facName.Replace("'", "''")}', N'{bizOp.Replace("'", "''")}', N'{addr.Replace("'", "''")}', N'{licensee.Replace("'", "''")}', '{tax.Replace("'", "''")}', '{phone.Replace("'", "''")}', '{fax.Replace("'", "''")}', {LinkedItemNo}, '{wasteCode.Replace("'", "''")}', '{haz.Replace("'", "''")}', N'{prop.Replace("'", "''")}', N'{wasteName.Replace("'", "''")}', N'{process.Replace("'", "''")}', N'{reason.Replace("'", "''")}');");
                                                    dtSqlBuilder.AppendLine("END");
                                                    sqlBuilder.Append(dtSqlBuilder.ToString());

                                                    UpdateLog($"⚡ สะสมคำสั่ง SQL บันทึกข้อมูล DT แถวที่ {sequenceCounter} สำเร็จ");
                                                    sequenceCounter++;
                                                }
                                                catch (Exception ex) { UpdateLog($"⚠️ ข้อผิดพลาดสกัดจัดฟิลด์ SQL DT: {ex.Message}"); }

                                                await RandomDelay(1000, 2000);

                                                // ลอจิกปิดหน้าต่างย่อยเพื่อคืนคิวกลับมา
                                                var closeBtn = detailModal.Locator("button[data-bs-dismiss='modal'].btn-secondary, button:has-text('ปิด')").First;
                                                if (await closeBtn.IsVisibleAsync()) { await closeBtn.ClickAsync(new() { Force = true }); }
                                                else { await detailModal.Locator("button.btn-close").First.ClickAsync(new() { Force = true }); }

                                                var backdrop = page.Locator(".modal-backdrop");
                                                int backdropCount = await backdrop.CountAsync();
                                                for (int b = 0; b < backdropCount; b++) { try { await backdrop.Nth(b).WaitForAsync(new() { State = WaitForSelectorState.Hidden, Timeout = 500 }); } catch { } }

                                                int nextJ = j + 1;
                                                if (nextJ < innerDataRows.Count)
                                                {
                                                    int nextButtonIndex = nextJ + 1;
                                                    var nextInspectBtn = targetModal.Locator($"button#btt_rd{nextButtonIndex}").First;
                                                    try { await nextInspectBtn.WaitForAsync(new() { State = WaitForSelectorState.Visible, Timeout = 6000 }); } catch { }
                                                }
                                                currentButtonIndex++;
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        var outerCloseBtn = targetModal.Locator("button[data-bs-dismiss='modal'], button:has-text('ปิด')").Last;
                        if (await outerCloseBtn.IsVisibleAsync())
                        {
                            await outerCloseBtn.ClickAsync(new() { Force = true });
                            await page.WaitForTimeoutAsync(1500);
                        }

                    } // 🔁 จบลูปแถวข้อมูลหน้าแรกหลักทั้งหมด

                    // ====================================================================
                    // 🗄️ [จบลูปใหญ่รวดเดียว] เชื่อมต่อ Database สั่งยิงตูมเดียวจบหลังบ้าน (Batch Execute)
                    // ====================================================================
                    string finalSqlCommands = sqlBuilder.ToString();
                    if (!string.IsNullOrWhiteSpace(finalSqlCommands))
                    {
                        UpdateLog("\n🗄️ เริ่มกระบวนการบันทึกข้อมูลเข้าสู่ฐานข้อมูลระบบ MSSQL...");
                        UpdateStatus("⏳ กำลังบันทึกข้อมูลเข้าตาราง HD และ DT...", "#FFCC00");

                        using (SqlConnection conn = new SqlConnection(_connString))
                        {
                            await conn.OpenAsync();
                            using (SqlCommand cmd = new SqlCommand(finalSqlCommands, conn))
                            {
                                cmd.CommandTimeout = 180; // เผื่อประมวลผลกรณีข้อมูลเยอะมาก ๆ ครับพี่
                                int totalRowsAffected = await cmd.ExecuteNonQueryAsync();
                                UpdateLog($"🏆 [DATABASE SUCCESS] บันทึกข้อมูลลงตาราง tbWasteScraperHD และ tbWasteScraperDT ครบถ้วนรวม {totalRowsAffected} แถวย่อย!");
                            }
                        }
                    }

                    Dispatcher.Invoke(() =>
                    {
                        // โยนปุ่มตัวเอง และสร้าง RoutedEventArgs ว่าง ๆ ส่งเข้าไป
                        //BtnExportExcel_Click(BtnExportExcel, new RoutedEventArgs());
                        ExportExcel();
                    });

                    UpdateStatus("✅ สกัดและจัดเก็บข้อมูลเข้าฐานข้อมูล HD/DT เรียบร้อย 100%!", "#00FF99");
                }
                catch (Exception ex)
                {
                    UpdateLog($"❌ ขัดข้องรุนแรงระหว่างสกัดส่งออก: {ex.Message}");
                    UpdateStatus("❌ โปรแกรมทำงานขัดข้อง", "#FF3B30");
                }
                finally
                {
                    Dispatcher.Invoke(() => BtnStart.IsEnabled = true);
                }
            });
        }

        public async Task RandomDelay(int minDelay, int maxDelay)
        {
            Random random = new Random();
            int delay = random.Next(minDelay, maxDelay);
            await Task.Delay(delay);
        }

        private  void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            Reportwindown report = new Reportwindown();
            report.Show();     
        }

        private async void ExportExcel()
        {
            
               // ดักจับรหัสบริษัทที่ User เลือกบนหน้าจอ เพื่อเอาไปคิวรี่เฉพาะของบริษัทนั้น ๆ
               string selectedCompanyId = "";
               Dispatcher.Invoke(() =>
               {
                   if (CboCompany.SelectedItem is ComboBoxItem selectedItem)
                   {
                       selectedCompanyId = selectedItem.Tag?.ToString() ?? "";
                   }
               });

               if (string.IsNullOrWhiteSpace(selectedCompanyId))
               {
                   MessageBox.Show("กรุณาเลือกบริษัทเป้าหมายก่อนส่งออกรายงานครับ", "แจ้งเตือน", MessageBoxButton.OK, MessageBoxImage.Warning);
                   return;
               }

          BtnExportExcel.IsEnabled = false;
           UpdateStatus("⏳ กำลังดึงข้อมูลและสร้างไฟล์ Excel ด้วย ClosedXML...", "#007ACC");

           await Task.Run(async () =>
           {
               try
               {
                   string query = @$"
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
                   WHERE h.ischeck <> 'Y' AND d.OperatorCode LIKE '{selectedCompanyId}%'
                   ORDER BY h.SubmissionDate ASC, d.SequenceNumber ASC;";

                   System.Data.DataTable dtReport = new System.Data.DataTable();
                   using (SqlConnection conn = new SqlConnection(_connString))
                   {
                       await conn.OpenAsync();
                       using (SqlCommand cmd = new SqlCommand(query, conn))
                       {
                           using (SqlDataAdapter da = new SqlDataAdapter(cmd)) { da.Fill(dtReport); }
                       }
                   }

                   if (dtReport.Rows.Count == 0)
                   {
                       Dispatcher.Invoke(() => MessageBox.Show("ไม่พบข้อมูลที่ดึงสำเร็จในระบบ", "รายงาน"));
                       return;
                   }

                   string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                   string fileName = $"WasteReport_{selectedCompanyId}_{timestamp}.xlsx";
                   string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                   string filePath = Path.Combine(desktopPath, fileName);

                   // 🎯 2. สั่งเขียนไฟล์ด้วย ClosedXML (ฟรี 100% ไม่มีตัวกรองสิทธิ์กวนใจ)
                   using (var workbook = new XLWorkbook())
                   {
                       // โหลดข้อมูลจาก DataTable ลงชีต และตั้งชื่อชีต
                       var worksheet = workbook.Worksheets.Add(dtReport, "ข้อมูลสิ่งปฏิกูล");

                       // ตกแต่งหัวตารางเป็นสีน้ำเงินตัวอักษรขาวเท่ ๆ
                       var headerRange = worksheet.Row(1).CellsUsed();
                       headerRange.Style.Font.Bold = true;
                       headerRange.Style.Fill.BackgroundColor = XLColor.FromHtml("#007ACC");
                       headerRange.Style.Font.FontColor = XLColor.White;

                       // สั่งขยายความกว้างคอลัมน์อัตโนมัติ (AutoFit)
                       worksheet.Columns().AdjustToContents();

                       // บันทึกไฟล์ลง Desktop
                       workbook.SaveAs(filePath);
                   }

                   UpdateLog($"🏆 SUCCESSFUL! ส่งออกสำเร็จด้วย ClosedXML ชื่อไฟล์: '{fileName}'");
               }
               catch (Exception ex)
               {
                   UpdateLog($"❌ ClosedXML ขัดข้อง: {ex.Message}");
               }
               finally
               {
                   Dispatcher.Invoke(() => {
                       BtnExportExcel.IsEnabled = true;
                       UpdateStatus("✅ ส่งออกรายงาน Excel สำเร็จ!", "#00FF99");
                   });
               }

           });

        }

    }
}