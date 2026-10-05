using ClosedXML.Excel;
using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.Data.SqlClient;
using Microsoft.Playwright;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace BWG_WasteScraper
{
    public partial class WastePermit : Window
    {
        private List<string> exportData = new List<string>();
        private readonly string _targetUrl = "https://waste-permit.diw.go.th/login";
        string? _connString = AppSettings.ConnectionString;
        private readonly AcceptanceDb _db = new AcceptanceDb();

        public WastePermit(string appUsername)
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
            var dateFrom = DpFrom.SelectedDate;
            var dateTo = DpTo.SelectedDate;
            if (dateFrom.HasValue && dateTo.HasValue && dateFrom.Value > dateTo.Value)
            {
                MessageBox.Show("วันที่เริ่มต้นต้องไม่มากกว่าวันที่สิ้นสุด", "ตรวจสอบวันที่",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string webUser = TxtWebUsername.Text.Trim();
            string webPass = TxtWebPassword.Password.Trim();

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
                    //Login(webUser, webPass);
                    const int FactorySwitchDelayMs = 10000;
                    using var playwright = await Playwright.CreateAsync();
                    await using var browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions { Headless = false });
                    var page = await browser.NewPageAsync();
                    // ---------- 1) Login ----------
                    UpdateLog($"🌐 นำทางไปยังหน้าเว็บไซต์: {_targetUrl}");
                    await page.GotoAsync(_targetUrl);
                    await page.GetByTestId("login-form").WaitForAsync(new() { Timeout = 15000 });
                    await page.GetByTestId("login-username-input").FillAsync(webUser);
                    await page.GetByTestId("login-password-input").FillAsync(webPass);

                    await page.GetByTestId("login-submit-btn").ClickAsync();
                    var factoryCard = page.GetByTestId("select-factory-card");
                    await factoryCard.WaitForAsync(new() { Timeout = 15000 });
                    UpdateLog("⏳ รอระบบตรวจสอบรหัสผ่านหน้าเว็บ...");
                    UpdateLog("✅ ล็อกอินสำเร็จ!");

                    // ---------- 2) เปิด dropdown (accordion) รายการแรก ----------
                    var summary = page.Locator("[data-testid^='accordion-summary-']").First;
                    await summary.WaitForAsync();

                    if (await summary.GetAttributeAsync("aria-expanded") != "true")
                    {
                        await summary.ClickAsync();
                    }
                    var firstRow = page.Locator("[data-testid^='factory-row-']").First;
                    await firstRow.WaitForAsync();

                    var regNo = (await firstRow.Locator("[data-testid^='factory-reg-no-']").InnerTextAsync()).Trim();
                    UpdateLog($"รายการแรก: เลขทะเบียนโรงงาน {regNo}");

                    // ---------- 3) กด "ดำเนินการ" ที่รายการแรก ----------
                    await firstRow.Locator("[data-testid^='btn-proceed-']").ClickAsync();

                    await factoryCard.WaitForAsync(new() { State = WaitForSelectorState.Detached, Timeout = 15000 });
                    UpdateLog($"เข้าสู่หน้าถัดไป: {page.Url}");

                    // ---------- 4) กด "ยืนยันดำเนินการ" ----------
                    var confirmBtn = page.GetByTestId("btn-confirm");
                    await confirmBtn.WaitForAsync(new() { Timeout = 15000 });
                    await confirmBtn.ClickAsync();
                    UpdateLog("กดยืนยันดำเนินการแล้ว");

                    // ---------- ฟังก์ชันจัดการ dropdown โรงงานที่หัวเว็บ ----------
                    var factoryInput = page.GetByTestId("content-header-factory-select-input");
                    var factoryRoot = page.Locator(".MuiAutocomplete-inputRoot").Filter(new() { Has = factoryInput });
                    var factoryOptions = page.GetByRole(AriaRole.Option);

                    async Task OpenFactoryListAsync()
                    {
                        await factoryRoot.Locator("button[aria-label='Open']").ClickAsync();
                        await factoryOptions.First.WaitForAsync(new() { Timeout = 10000 });
                    }

                    async Task<(List<string> Texts, int SelectedIndex)> ReadFactoryOptionsAsync()
                    {
                        await OpenFactoryListAsync();
                        var texts = new List<string>();
                        var selected = -1;
                        var n = await factoryOptions.CountAsync();
                        for (var i = 0; i < n; i++)
                        {
                            var opt = factoryOptions.Nth(i);
                            texts.Add((await opt.InnerTextAsync()).Trim());
                            if (await opt.GetAttributeAsync("aria-selected") == "true") selected = i;
                        }
                        await page.Keyboard.PressAsync("Escape");
                        return (texts, selected);
                    }

                    async Task SelectFactoryAsync(int index)
                    {
                        await OpenFactoryListAsync();
                        await factoryOptions.Nth(index).ClickAsync();

                        // รอให้เว็บเปลี่ยนหน้า/โหลดข้อมูลของโรงงานใหม่ ก่อนทำขั้นตอนต่อไป
                        await page.WaitForTimeoutAsync(FactorySwitchDelayMs);
                        await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                        // เมนูข้างต้องแสดงก่อน แปลว่าหน้าหลักโหลดเสร็จแล้ว
                        await page.GetByTestId("sidebar-nav-ko2").WaitForAsync(new() { Timeout = 30000 });
                    }

                    // ---------- วนทุกโรงงานใน dropdown ----------
                    var (optionTexts, currentIdx) = await ReadFactoryOptionsAsync();
                    UpdateLog($"พบโรงงานใน dropdown {optionTexts.Count} รายการ");
                    for (var i = 0; i < optionTexts.Count; i++)
                    UpdateLog($"[{i + 1}] {optionTexts[i]}");
                    var failed = new List<string>();
                    for (var i = 0; i < optionTexts.Count; i++)
                    {
                        UpdateLog($"\n===== โรงงาน {i + 1}/{optionTexts.Count}: {optionTexts[i]} =====");
                        try
                        {
                            await RandomDelay(3000, 6000);
                            await SelectFactoryAsync(i);

                            var name = (await factoryInput.InputValueAsync()).Trim();

                            // เลขทะเบียน: ลองดึงเลข 14 หลักจากข้อความตัวเลือก
                            // ถ้าไม่มี และเป็นโรงงานที่เลือกไว้ตั้งแต่ขั้นตอน 2 ใช้ regNo เดิม มิฉะนั้นปล่อยว่าง
                            var m = Regex.Match(optionTexts[i], @"\d{14}");
                            var reg = m.Success ? m.Value : (i == currentIdx ? regNo : "");

                            await ProcessFactoryAsync(name, reg, (i + 1).ToString("00"));
                        }
                        catch (Exception ex)
                        {
                            UpdateLog($"โรงงานที่ {i + 1} ผิดพลาด: {ex.Message}");
                            failed.Add($"[{i + 1}] {optionTexts[i]}");
                            await page.ScreenshotAsync(new() { Path = $"output/error-factory-{i + 1:00}.png" });
                            // ไปโรงงานถัดไปต่อ ไม่หยุดทั้งโปรแกรม
                        }
                    }
                    UpdateLog($"\nเสร็จสิ้น สำเร็จ {optionTexts.Count - failed.Count}/{optionTexts.Count} โรงงาน");
                    foreach (var f in failed) UpdateLog($"  ล้มเหลว: {f}");

                    async Task ProcessFactoryAsync(string operatorName, string operatorRegNo, string fileTag)
                    {
                        // ---------- 5) เมนูข้าง: ผู้รับดำเนินการ > รายการขออนุญาตฯ ----------
                        var menuParent = page.GetByTestId("sidebar-nav-ko2");
                        var acceptanceLink = page.GetByTestId("sidebar-nav-acceptance-list");

                        await menuParent.WaitForAsync(new() { Timeout = 15000 });

                        // คลิกเมนูหลักเฉพาะตอนที่เมนูย่อยยังไม่แสดง (กันกดแล้วเมนูพับปิด)
                        if (!await acceptanceLink.IsVisibleAsync())
                        {
                            await menuParent.ClickAsync();
                            await acceptanceLink.WaitForAsync();
                        }

                        await acceptanceLink.ClickAsync();
                        await page.WaitForURLAsync("**/ko1/acceptance", new() { Timeout = 15000 });
                        // รอให้หน้ารายการขออนุญาตฯ โหลดครบก่อนเลือกวันที่
                        await page.GetByTestId("acceptance-list-page").WaitForAsync(new() { Timeout = 20000 });
                        await page.GetByTestId("acceptance-search-date-from").WaitForAsync(new() { Timeout = 20000 });
                        await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                        await RandomDelay(500, 1000);
                        await page.WaitForTimeoutAsync(1000);   // กันเว็บ render ซ้ำหลังโหลดเสร็จ
                        UpdateLog($"เข้าหน้ารายการขออนุญาตฯ: {page.Url}");

                        // ---------- 6) เลือกวันที่ ----------
                        async Task PickDateAsync(string inputTestId, string name, DateTime target)
                        {
                            var input = page.GetByTestId(inputTestId);
                            await input.ClickAsync();

                            var popup = page.Locator(
                                "[role='dialog'], .MuiPopover-root, .MuiPickersPopper-root, .MuiPopper-root, .react-datepicker, .rdp, [class*='calendar' i]").Last;

                            try
                            {
                                await popup.WaitForAsync(new() { Timeout = 5000 });
                            }
                            catch (TimeoutException)
                            {
                                await input.Locator("xpath=following-sibling::div//*[name()='svg']").ClickAsync();
                                try { await popup.WaitForAsync(new() { Timeout = 5000 }); }
                                catch (TimeoutException)
                                {
                                    //await page.ScreenshotAsync(new() { Path = $"output/datepicker-{name}.png" });
                                    await File.WriteAllTextAsync($"output/datepicker-{name}.html", await page.ContentAsync());
                                    throw new Exception($"ไม่พบ popup ปฏิทิน ดู output/datepicker-{name}.png และ .html");
                                }
                            }

                            // ปฏิทินเปิดที่เดือนปัจจุบัน ให้เดินไปเดือนเป้าหมาย (ถอยหลังหรือไปข้างหน้าก็ได้)
                            var now = DateTime.Now;
                            var monthDiff = (target.Year - now.Year) * 12 + (target.Month - now.Month);

                            var nextSel = "button[aria-label*='next month' i], button[aria-label*='Next' i], button[aria-label*='ถัดไป'], button[title*='next' i]";
                            var prevSel = "button[aria-label*='previous month' i], button[aria-label*='Previous' i], button[aria-label*='ก่อนหน้า'], button[title*='previous' i]";

                            if (Math.Abs(monthDiff) > 36)
                                throw new Exception($"{name}: วันที่ {target:yyyy-MM-dd} ห่างจากเดือนปัจจุบันเกิน 36 เดือน");

                            for (var m = 0; m < Math.Abs(monthDiff); m++)
                            {
                                var navBtn = monthDiff > 0 ? popup.Locator(nextSel).Last : popup.Locator(prevSel).First;
                                if (await navBtn.CountAsync() == 0)
                                {
                                    await File.WriteAllTextAsync($"output/datepicker-{name}.html", await popup.EvaluateAsync<string>("e => e.outerHTML"));
                                    throw new Exception($"หาปุ่มเปลี่ยนเดือนไม่เจอ ดู output/datepicker-{name}.html");
                                }
                                await navBtn.ClickAsync();
                                await page.WaitForTimeoutAsync(200);   // รอปฏิทินวาดเดือนใหม่
                            }

                            await File.WriteAllTextAsync($"output/datepicker-{name}.html", await popup.EvaluateAsync<string>("e => e.outerHTML"));
                            //await page.ScreenshotAsync(new() { Path = $"output/datepicker-{name}.png" });

                            // วันที่ 1-15 ใช้ตัวแรก กันชนกับวันท้ายเดือนก่อนที่ปฏิทินแสดงโชว์, วันที่ 16+ ใช้ตัวสุดท้าย
                            var day = target.Day;
                            var days = popup.Locator("button:not([disabled]), [role='gridcell']:not([aria-disabled='true']), td:not(.disabled)")
                                .Filter(new() { HasTextRegex = new Regex($"^\\s*{day}\\s*$") });
                            var dayBtn = day <= 15 ? days.First : days.Last;
                            await dayBtn.ClickAsync();

                            var ok = popup.GetByRole(AriaRole.Button, new() { NameRegex = new Regex("^(ตกลง|ยืนยัน|OK)$", RegexOptions.IgnoreCase) });
                            if (await ok.CountAsync() > 0) await ok.First.ClickAsync();

                            await Assertions.Expect(input).ToHaveValueAsync(new Regex(".+"));
                            UpdateLog($"เลือก{name}แล้ว: {await input.InputValueAsync()}");
                            await RandomDelay(500, 1000);
                        }

                        //var startDate = DateTime.Now;
                        //var endDate = startDate.AddDays(10);

                        //await PickDateAsync("acceptance-search-date-from", "วันที่เริ่มต้น", startDate);
                        //await PickDateAsync("acceptance-search-date-to", "วันที่สิ้นสุด", endDate);

                        await SelectSearchDatesAsync(dateFrom, dateTo);

                        // ---------- กดปุ่มค้นหา ----------
                        var searchBtn = page.GetByTestId("acceptance-search-btn");
                        await searchBtn.WaitForAsync(new() { Timeout = 10000 });
                        await Assertions.Expect(searchBtn).ToBeEnabledAsync();

                        await RandomDelay(500, 1000);   // รอปฏิทินปิดสนิท ไม่ให้บังปุ่ม
                        await searchBtn.ClickAsync();
                        UpdateLog("กดปุ่มค้นหาแล้ว");

                        // รอผลการค้นหา
                        await RandomDelay(1500, 2500);                          // ให้เว็บเริ่มส่งคำขอค้นหา
                        await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                        await page.GetByTestId("acceptance-tabs").WaitForAsync(new() { Timeout = 15000 });
                        await RandomDelay(500, 1000);                           // ให้ตารางวาดเสร็จ

                        async Task SelectSearchDatesAsync(DateTime? dateFrom, DateTime? dateTo)
                        {
                            if (dateFrom.HasValue)
                                await PickDateAsync("acceptance-search-date-from", "วันที่เริ่มต้น", dateFrom.Value.Date);
                            if (dateTo.HasValue)
                                await PickDateAsync("acceptance-search-date-to", "วันที่สิ้นสุด", dateTo.Value.Date);
                        }

                        // แท็บที่ต้องวน (เรียงตามลำดับ)
                        var tabs = new[]
                        {
                            ("acceptance-tab-pending",        "รอตอบรับ"),
                            ("acceptance-tab-pending-review", "อยู่ระหว่างพิจารณา"),
                            ("acceptance-tab-completed",      "ดำเนินการเสร็จสิ้น"),
                        };

                        async Task SelectTabAsync(string tabTestId)
                        {
                            var tab = page.GetByTestId(tabTestId);
                            await tab.WaitForAsync();
                            if (await tab.GetAttributeAsync("aria-selected") != "true")
                            {
                                await tab.ClickAsync();
                                await Assertions.Expect(tab).ToHaveAttributeAsync("aria-selected", "true");
                                await RandomDelay(800, 1500);   // รอตารางเปลี่ยนเป็นของแท็บใหม่
                            }
                        }

                        async Task<int> ReadTabCountAsync(string tabTestId)
                        {
                            var m = Regex.Match(await page.GetByTestId(tabTestId).InnerTextAsync(), @"\((\d+)\)");
                            return m.Success ? int.Parse(m.Groups[1].Value) : 0;
                        }

                        async Task<AcceptanceRow> ReadRowAsync(ILocator row)
                        {
                            var id = await row.GetAttributeAsync("data-id") ?? "";
                            async Task<string> CellAsync(string field) =>
                                (await row.Locator($"[data-field='{field}']").InnerTextAsync()).Trim();

                            var factoryCell = row.Locator("[data-field='generatorFactory']");
                            var dot = row.GetByTestId($"acceptance-deadline-dot-{id}");   // แท็บที่เสร็จแล้วอาจไม่มีจุดกำหนดตอบรับ

                            return new AcceptanceRow(
                                Id: id,
                                Sequence: await CellAsync("sequence"),
                                RequestNumber: await CellAsync("requestNumber"),
                                FactoryName: (await factoryCell.Locator("p").InnerTextAsync()).Trim(),
                                FactoryRegNo: (await factoryCell.Locator("span").InnerTextAsync()).Trim(),
                                SubmittedAt: await CellAsync("submittedAt"),
                                Status: await CellAsync("acceptanceStatus"),
                                RequestType: await CellAsync("requestType"),
                                Deadline: await dot.CountAsync() > 0 ? (await dot.GetAttributeAsync("aria-label") ?? "") : "");
                        }

                        async Task ScrapeTabAsync(string tabTestId, string tabName, int total)
                        {
                            var grid = page.GetByTestId("acceptance-list-table");
                            var rows = grid.Locator("div[role='row'][data-id]");
                            var done = 0;
                            var pageNo = 1;
                            var results = new List<AcceptanceRow>();
                            var details = new Dictionary<string, AcceptanceDetail>();

                            async Task RestoreListAsync()
                            {
                                await page.GetByTestId("acceptance-list-page").WaitForAsync(new() { Timeout = 20000 });

                                // ถ้าตัวกรองวันที่ถูกรีเซ็ตหลังย้อนกลับ ให้กรองและค้นหาใหม่
                                if (dateFrom.HasValue &&
                                    string.IsNullOrEmpty(await page.GetByTestId("acceptance-search-date-from").InputValueAsync()))
                                {
                                    await SelectSearchDatesAsync(dateFrom, dateTo);
                                    await page.GetByTestId("acceptance-search-btn").ClickAsync();
                                    await RandomDelay(1500, 2500);
                                    await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                                }

                                await SelectTabAsync(tabTestId);

                                var cur = int.Parse(await page
                                    .Locator("[data-testid^='pagination-page-'][aria-current='page']").InnerTextAsync());
                                if (cur != pageNo)
                                {
                                    await page.GetByTestId($"pagination-page-{pageNo}").ClickAsync();
                                    await RandomDelay(800, 1500);
                                }
                                await rows.First.WaitForAsync();
                            }

                            while (done < total)
                            {
                                await rows.First.WaitForAsync();
                                var firstIdOnPage = await rows.First.GetAttributeAsync("data-id") ?? "";
                                var rowCount = await rows.CountAsync();

                                for (var i = 0; i < rowCount && done < total; i++)
                                {
                                    var row = rows.Nth(i);
                                    var data = await ReadRowAsync(row);
                                    results.Add(data);

                                    // คลิกปุ่มในคอลัมน์สุดท้าย (ไม่ผูกกับ testid เผื่อชื่อปุ่มต่างกันในแต่ละแท็บ)
                                    var listUrl = page.Url;
                                    await row.Locator("[data-field='actions'] button").First.ClickAsync();
                                    await page.WaitForURLAsync(u => u != listUrl, new() { Timeout = 15000 });
                                    await RandomDelay(1000, 2000);

                                    var detail = await DetailScraper.ScrapeAsync(page);
                                    details[data.Id] = detail;
                                    await _db.SaveAsync(data, detail, operatorRegNo, operatorName);
                                    done++;
                                    UpdateLog($"[{tabName}] {done}/{total} {detail.RequestNumber} สถานะ: {detail.Status}");

                                    await page.GoBackAsync();
                                    await RestoreListAsync();
                                    await RandomDelay(1500, 3000);
                                }

                                // หน้าถัดไป (ถ้ายังไม่ครบ)
                                var next = page.GetByTestId($"pagination-page-{pageNo + 1}");
                                if (done >= total || await next.CountAsync() == 0) break;

                                await next.ClickAsync();
                                pageNo++;
                                await Assertions.Expect(rows.First).Not.ToHaveAttributeAsync("data-id", firstIdOnPage);
                            }
                        }

                        // อ่านจำนวนของทุกแท็บก่อน (ตัวเลขในชื่อแท็บมองเห็นพร้อมกัน)
                        var counts = new Dictionary<string, int>();
                        foreach (var (tabId, tabName) in tabs)
                        {
                            counts[tabId] = await ReadTabCountAsync(tabId);
                            UpdateLog($"แท็บ {tabName}: {counts[tabId]} รายการ");
                        }

                        foreach (var (tabId, tabName) in tabs)
                        {
                            if (counts[tabId] == 0)
                            {
                                UpdateLog($"ข้ามแท็บ {tabName} (0 รายการ)");
                                continue;
                            }

                            await SelectTabAsync(tabId);
                            await ScrapeTabAsync(tabId, tabName, counts[tabId]);
                        }

                        //// ---------- 7) แท็บ "รายการรอตอบรับ" อ่านจำนวนจากชื่อแท็บ เช่น "(1)" ----------
                        //var pendingTab = page.GetByTestId("acceptance-tab-pending");
                        //await pendingTab.WaitForAsync();

                        //if (await pendingTab.GetAttributeAsync("aria-selected") != "true")
                        //{
                        //    await pendingTab.ClickAsync();
                        //    await Assertions.Expect(pendingTab).ToHaveAttributeAsync("aria-selected", "true");
                        //    await RandomDelay(800, 1500);
                        //}

                        //var countMatch = Regex.Match(await pendingTab.InnerTextAsync(), @"\((\d+)\)");
                        //var total = countMatch.Success ? int.Parse(countMatch.Groups[1].Value) : 0;
                        //UpdateLog($"รายการรอตอบรับทั้งหมด: {total}");

                        //// ---------- 8) วนอ่านทุกแถว ทุกหน้า ----------
                        //var results = new List<AcceptanceRow>();
                        //var details = new Dictionary<string, AcceptanceDetail>();   // key = id ของแถว   // key = id ของแถว
                        //                                                            // ข้อมูลโรงงานผู้รับดำเนินการ: ชื่อจากช่องเลือกโรงงานที่หัวเว็บ, เลขทะเบียนจากตอนเลือกโรงงาน (ขั้นตอน 2)
                        //UpdateLog($"ผู้รับดำเนินการ: {operatorName} ({operatorRegNo})");
                        //if (total > 0)
                        //{
                        //    var grid = page.GetByTestId("acceptance-list-table");
                        //    var rows = grid.Locator("div[role='row'][data-id]");

                        //    while (results.Count < total)
                        //    {
                        //        await rows.First.WaitForAsync();
                        //        var firstIdOnPage = await rows.First.GetAttributeAsync("data-id") ?? "";
                        //        var rowCount = await rows.CountAsync();

                        //        for (var i = 0; i < rowCount && results.Count < total; i++)
                        //        {
                        //            var row = rows.Nth(i);
                        //            var id = await row.GetAttributeAsync("data-id") ?? "";

                        //            async Task<string> CellAsync(string field) =>
                        //                (await row.Locator($"[data-field='{field}']").InnerTextAsync()).Trim();

                        //            var factoryCell = row.Locator("[data-field='generatorFactory']");

                        //            results.Add(new AcceptanceRow(
                        //                Id: id,
                        //                Sequence: await CellAsync("sequence"),
                        //                RequestNumber: await CellAsync("requestNumber"),
                        //                FactoryName: (await factoryCell.Locator("p").InnerTextAsync()).Trim(),
                        //                FactoryRegNo: (await factoryCell.Locator("span").InnerTextAsync()).Trim(),
                        //                SubmittedAt: await CellAsync("submittedAt"),
                        //                Status: await CellAsync("acceptanceStatus"),
                        //                RequestType: await CellAsync("requestType"),
                        //                Deadline: await row.GetByTestId($"acceptance-deadline-dot-{id}").GetAttributeAsync("aria-label") ?? ""));

                        //            // ---------- เก็บข้อมูลแถวนี้เสร็จ: คลิก "ตรวจสอบ" แล้วรอหน้า detail ----------
                        //            var listUrl = page.Url;
                        //            await row.GetByTestId($"acceptance-inspect-btn-{id}").ClickAsync();
                        //            await page.WaitForURLAsync(u => u != listUrl, new() { Timeout = 15000 });
                        //            UpdateLog($"เข้าหน้า detail ของคำขอ {results[^1].RequestNumber}: {page.Url}");

                        //            // TODO: logic หน้า detail ใส่ตรงนี้
                        //            await RandomDelay(1000, 2000);
                        //            var detail = await DetailScraper.ScrapeAsync(page);
                        //            details[id] = detail;
                        //            UpdateLog($"อ่านหน้า detail แล้ว: {detail.RequestNumber} มี {detail.Items.Count} รายการสิ่งปฏิกูล (หน้าเว็บระบุ {detail.ItemCount})");

                        //            await _db.SaveAsync(results[^1], detail, operatorRegNo, operatorName);
                        //            UpdateLog($"บันทึกลง DB แล้ว: {detail.RequestNumber} ({detail.Items.Count} รายการ)");


                        //            // กลับหน้ารายการ เพื่อวนแถวถัดไป
                        //            await page.GoBackAsync();
                        //            await RandomDelay(1500, 3000);
                        //            await grid.WaitForAsync();
                        //            await rows.First.WaitForAsync();
                        //        }

                        //        // หาเลขหน้าปัจจุบัน แล้วกดไปหน้า +1
                        //        var currentPage = int.Parse(await page
                        //            .Locator("[data-testid^='pagination-page-'][aria-current='page']")
                        //            .InnerTextAsync());
                        //        var nextPageBtn = page.GetByTestId($"pagination-page-{currentPage + 1}");

                        //        // ครบแล้ว หรือไม่มีปุ่มหน้าถัดไป ให้หยุด
                        //        if (results.Count >= total || await nextPageBtn.CountAsync() == 0) break;

                        //        await nextPageBtn.ClickAsync();
                        //        // รอจนแถวแรกเปลี่ยน แปลว่าโหลดหน้าใหม่แล้ว
                        //        await Assertions.Expect(rows.First).Not.ToHaveAttributeAsync("data-id", firstIdOnPage);
                        //    }
                        //}

                    }

                    UpdateStatus("✅ สกัดและจัดเก็บข้อมูลเข้าฐานข้อมูล เรียบร้อย 100%!", "#00FF99");

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

        private void BtnExportExcel_Click(object sender, RoutedEventArgs e)
        {
            ReportWastePermit report = new ReportWastePermit();
            report.WindowState = WindowState.Maximized;
            report.Show();
        }

    }
}

