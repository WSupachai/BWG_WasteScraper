using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Playwright;

namespace BWG_WasteScraper
{
    public record WasteItem(
    string ItemTitle, string ItemId, string Status,
    string WasteCode, string WasteName, string Quantity,
    string Description, string Hazard, string ManagementCodeText, string Response,
    Dictionary<string, string> Monthly);

    internal record AcceptanceDetail(
        string RequestNumber, string Factory, string RequestType, string SubmittedAt,
        string Status, string DeadlineAt, string DeadlineChip, string DeadlineInfo,
        string Year, string ItemCount, List<WasteItem> Items);

    internal static class DetailScraper
    {
        private static readonly string[] Months =
            { "ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค." };

        //private static async Task<string> TextAsync(ILocator l) => (await l.InnerTextAsync()).Trim();
        private static async Task<string> TextAsync(ILocator l) =>
        await l.CountAsync() > 0 ? (await l.First.InnerTextAsync()).Trim() : "";

        public static async Task<AcceptanceDetail> ScrapeAsync(IPage page)
        {
            var root = page.GetByTestId("acceptance-detail-page");
            await root.WaitForAsync(new() { Timeout = 15000 });

            var cards = root.Locator("[data-testid^='acceptance-item-card-']");
            await cards.First.WaitForAsync(new() { Timeout = 15000 });

            // ค่าที่อยู่ใน <span> หลังป้ายชื่อ เช่น "ปีที่ขออนุญาต : " -> "2570"
            async Task<string> AfterLabelAsync(string label) =>
                await TextAsync(root.Locator("p").Filter(new() { HasTextString = label }).First.Locator("span").Last);

            var requestNo = await TextAsync(root.Locator("h5").First);          // "เลขที่คำขอ : 102569-16"
            requestNo = requestNo.Contains(':') ? requestNo.Split(':', 2)[1].Trim() : requestNo;

            var items = new List<WasteItem>();
            var cardCount = await cards.CountAsync();

            for (var i = 0; i < cardCount; i++)
            {
                var card = cards.Nth(i);
                var testId = await card.GetAttributeAsync("data-testid") ?? "";
                var itemId = testId.Replace("acceptance-item-card-", "");

                // ป้ายชื่อเป็น <span> ตามด้วย <p> ที่เก็บค่า
                async Task<string> FieldAsync(string label) =>
                    await TextAsync(card.Locator($"xpath=.//span[normalize-space()='{label}']/following-sibling::p[1]"));

                var monthly = new Dictionary<string, string>();
                for (var m = 1; m <= 12; m++)
                    monthly[Months[m - 1]] = await TextAsync(page.GetByTestId($"acceptance-monthly-cell-{itemId}-{m}"));

                // ผลการตอบรับที่เลือกอยู่ตอนนี้ (อ่านอย่างเดียว ไม่คลิก)
                //var accepted = await page.GetByTestId($"acceptance-item-accept-{itemId}").Locator("input").IsCheckedAsync();
                //var rejected = await page.GetByTestId($"acceptance-item-reject-{itemId}").Locator("input").IsCheckedAsync();
                //var response = accepted ? "ตอบรับ" : rejected ? "ไม่ตอบรับ" : "ยังไม่เลือก";
                var acceptInput = page.GetByTestId($"acceptance-item-accept-{itemId}").Locator("input");
                var rejectInput = page.GetByTestId($"acceptance-item-reject-{itemId}").Locator("input");
                var accepted = await acceptInput.CountAsync() > 0 && await acceptInput.IsCheckedAsync();
                var rejected = await rejectInput.CountAsync() > 0 && await rejectInput.IsCheckedAsync();
                var response = accepted ? "ตอบรับ" : rejected ? "ไม่ตอบรับ" : "ยังไม่เลือก";

                items.Add(new WasteItem(
                    ItemTitle: await TextAsync(card.Locator("h6").First),
                    ItemId: itemId,
                    Status: await TextAsync(page.GetByTestId($"acceptance-item-status-{itemId}")),
                    WasteCode: await FieldAsync("รหัสประเภทหรือชนิด"),
                    WasteName: await FieldAsync("ชื่อสิ่งปฏิกูลหรือวัสดุที่ไม่ใช้แล้ว"),
                    Quantity: await FieldAsync("ปริมาณรับจัดการ (ตัน)"),
                    Description: await FieldAsync("รายละเอียดการก่อให้เกิดสิ่งปฏิกูลหรือวัสดุที่ไม่ใช้แล้ว"),
                    Hazard: await FieldAsync("ความเป็นอันตราย"),
                    ManagementCodeText: await TextAsync(page.GetByTestId($"acceptance-item-management-code-{itemId}")),
                    Response: response,
                    Monthly: monthly));
            }

            return new AcceptanceDetail(
                RequestNumber: requestNo,
                Factory: await AfterLabelAsync("โรงงานผู้ก่อกำเนิด"),
                RequestType: await TextAsync(page.GetByTestId("acceptance-detail-request-type-chip")),
                SubmittedAt: await TextAsync(root.Locator("p").Filter(new() { HasTextString = "วันที่ยื่นคำขอ" }).First.Locator("xpath=following-sibling::p[1]")),
                Status: await TextAsync(page.GetByTestId("acceptance-detail-status-chip")),
                DeadlineAt: await TextAsync(page.GetByTestId("acceptance-detail-deadline-at")),
                DeadlineChip: await TextAsync(page.GetByTestId("acceptance-deadline-chip")),
                DeadlineInfo: await TextAsync(page.GetByTestId("acceptance-detail-deadline-info")),
                Year: await AfterLabelAsync("ปีที่ขออนุญาต"),
                ItemCount: await AfterLabelAsync("จำนวนรายการสิ่งปฏิกูลฯ"),
                Items: items);
        }
    }
}
