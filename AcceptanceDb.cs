using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;

namespace BWG_WasteScraper
{
    internal class AcceptanceDb
    {
        private readonly string? _connString = AppSettings.ConnectionString;

        private SqlConnection CreateConnection()
        {
            if (string.IsNullOrWhiteSpace(_connString))
                throw new InvalidOperationException("ไม่พบ ConnectionString ใน AppSettings กรุณาตรวจสอบการตั้งค่า");

            return new SqlConnection(_connString);
        }

        // สร้างตารางถ้ายังไม่มี
        public async Task EnsureSchemaAsync()
        {
            var schemaPath = Path.Combine(AppContext.BaseDirectory, "schema.sql");
            var sql = await File.ReadAllTextAsync(schemaPath);

            await using var conn = CreateConnection();
            await conn.OpenAsync();
            await using var cmd = new SqlCommand(sql, conn);
            await cmd.ExecuteNonQueryAsync();
        }

        // บันทึก HD + DT ในทรานแซกชันเดียว ถ้ามีอยู่แล้วจะอัปเดต (รันซ้ำไม่เกิดข้อมูลซ้ำ)
        public async Task SaveAsync(AcceptanceRow row, AcceptanceDetail d, string operatorRegNo, string operatorName)
        {
            await using var conn = CreateConnection();
            await conn.OpenAsync();
            await using var tx = (SqlTransaction)await conn.BeginTransactionAsync();

            try
            {
                var acceptanceId = int.Parse(row.Id);

                // ---------- HD ----------
                const string hdSql = @"
UPDATE dbo.Acceptance_HD SET
    RequestNumber=@RequestNumber, FactoryName=@FactoryName, FactoryRegNo=@FactoryRegNo,
    OperatorFactoryRegNo=@OperatorFactoryRegNo, OperatorFactoryName=@OperatorFactoryName,
    RequestType=@RequestType, Status=@Status,
    SubmittedAt=@SubmittedAt, SubmittedAtText=@SubmittedAtText,
    DeadlineAt=@DeadlineAt, DeadlineAtText=@DeadlineAtText,
    DeadlineChip=@DeadlineChip, DeadlineInfo=@DeadlineInfo,
    RequestYear=@RequestYear, ItemCount=@ItemCount, UpdatedAt=SYSDATETIME()
WHERE AcceptanceId=@AcceptanceId;
IF @@ROWCOUNT = 0
INSERT INTO dbo.Acceptance_HD
    (AcceptanceId, RequestNumber, FactoryName, FactoryRegNo, OperatorFactoryRegNo, OperatorFactoryName,
     RequestType, Status, SubmittedAt, SubmittedAtText, DeadlineAt, DeadlineAtText, DeadlineChip, DeadlineInfo,
     RequestYear, ItemCount)
VALUES
    (@AcceptanceId, @RequestNumber, @FactoryName, @FactoryRegNo, @OperatorFactoryRegNo, @OperatorFactoryName,
     @RequestType, @Status, @SubmittedAt, @SubmittedAtText, @DeadlineAt, @DeadlineAtText, @DeadlineChip, @DeadlineInfo,
     @RequestYear, @ItemCount);";

                await using (var cmd = new SqlCommand(hdSql, conn, tx))
                {
                    cmd.Parameters.AddWithValue("@AcceptanceId", acceptanceId);
                    cmd.Parameters.AddWithValue("@RequestNumber", row.RequestNumber);
                    cmd.Parameters.AddWithValue("@FactoryName", row.FactoryName);
                    cmd.Parameters.AddWithValue("@FactoryRegNo", row.FactoryRegNo);
                    cmd.Parameters.AddWithValue("@OperatorFactoryRegNo", operatorRegNo);
                    cmd.Parameters.AddWithValue("@OperatorFactoryName", operatorName);
                    cmd.Parameters.AddWithValue("@RequestType", d.RequestType);
                    cmd.Parameters.AddWithValue("@Status", d.Status);
                    cmd.Parameters.AddWithValue("@SubmittedAt", Db(ParseThaiDate(row.SubmittedAt)));
                    cmd.Parameters.AddWithValue("@SubmittedAtText", row.SubmittedAt);
                    cmd.Parameters.AddWithValue("@DeadlineAt", Db(ParseThaiDate(d.DeadlineAt)));
                    cmd.Parameters.AddWithValue("@DeadlineAtText", d.DeadlineAt);
                    cmd.Parameters.AddWithValue("@DeadlineChip", d.DeadlineChip);
                    cmd.Parameters.AddWithValue("@DeadlineInfo", d.DeadlineInfo);
                    cmd.Parameters.AddWithValue("@RequestYear", Db(ParseInt(d.Year)));
                    cmd.Parameters.AddWithValue("@ItemCount", Db(ParseInt(d.ItemCount)));
                    await cmd.ExecuteNonQueryAsync();
                }

                // ---------- DT ----------
                var mCols = string.Join(", ", Enumerable.Range(1, 12).Select(i => $"M{i:00}"));
                var mVals = string.Join(", ", Enumerable.Range(1, 12).Select(i => $"@M{i:00}"));
                var mSet = string.Join(", ", Enumerable.Range(1, 12).Select(i => $"M{i:00}=@M{i:00}"));

                var dtSql = $@"
UPDATE dbo.Acceptance_DT SET
    AcceptanceId=@AcceptanceId, ItemNo=@ItemNo, ItemStatus=@ItemStatus,
    WasteCode=@WasteCode, WasteCodeText=@WasteCodeText, WasteName=@WasteName,
    QuantityTon=@QuantityTon, WasteDescription=@WasteDescription, Hazard=@Hazard,
    ManagementCode=@ManagementCode, ManagementCodeText=@ManagementCodeText,
    ResponseResult=@ResponseResult, {mSet}, UpdatedAt=SYSDATETIME()
WHERE AcceptanceItemId=@AcceptanceItemId;
IF @@ROWCOUNT = 0
INSERT INTO dbo.Acceptance_DT
    (AcceptanceItemId, AcceptanceId, ItemNo, ItemStatus, WasteCode, WasteCodeText, WasteName,
     QuantityTon, WasteDescription, Hazard, ManagementCode, ManagementCodeText, ResponseResult, {mCols})
VALUES
    (@AcceptanceItemId, @AcceptanceId, @ItemNo, @ItemStatus, @WasteCode, @WasteCodeText, @WasteName,
     @QuantityTon, @WasteDescription, @Hazard, @ManagementCode, @ManagementCodeText, @ResponseResult, {mVals});";

                foreach (var it in d.Items)
                {
                    var spaceAt = it.WasteCode.IndexOf(' ');
                    var code = spaceAt > 0 ? it.WasteCode[..spaceAt] : it.WasteCode;
                    var months = it.Monthly.Values.ToArray();   // เรียง ม.ค. → ธ.ค.

                    await using var cmd = new SqlCommand(dtSql, conn, tx);
                    cmd.Parameters.AddWithValue("@AcceptanceItemId", int.Parse(it.ItemId));
                    cmd.Parameters.AddWithValue("@AcceptanceId", acceptanceId);
                    cmd.Parameters.AddWithValue("@ItemNo", Db(ParseInt(it.ItemTitle)));
                    cmd.Parameters.AddWithValue("@ItemStatus", it.Status);
                    cmd.Parameters.AddWithValue("@WasteCode", code);
                    cmd.Parameters.AddWithValue("@WasteCodeText", it.WasteCode);
                    cmd.Parameters.AddWithValue("@WasteName", it.WasteName);
                    cmd.Parameters.AddWithValue("@QuantityTon", Db(ParseDecimal(it.Quantity)));
                    cmd.Parameters.AddWithValue("@WasteDescription", it.Description);
                    cmd.Parameters.AddWithValue("@Hazard", it.Hazard);
                    var mgmtCode = it.ManagementCodeText.Split('-', 2)[0].Trim();
                    cmd.Parameters.AddWithValue("@ManagementCode", mgmtCode);
                    cmd.Parameters.AddWithValue("@ManagementCodeText", it.ManagementCodeText);
                    cmd.Parameters.AddWithValue("@ResponseResult", it.Response);
                    for (var m = 0; m < 12; m++)
                        cmd.Parameters.AddWithValue($"@M{m + 1:00}", Db(ParseDecimal(months.ElementAtOrDefault(m) ?? "")));
                    await cmd.ExecuteNonQueryAsync();
                }

                await tx.CommitAsync();
            }
            catch
            {
                await tx.RollbackAsync();
                throw;
            }
        }

        // ---------- helpers ----------
        private static object Db(object? v) => v ?? DBNull.Value;

        private static int? ParseInt(string? s)
        {
            var m = Regex.Match(s ?? "", @"\d+");
            return m.Success ? int.Parse(m.Value) : null;
        }

        private static decimal? ParseDecimal(string? s) =>
            decimal.TryParse((s ?? "").Replace(",", "").Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : null;

        private static readonly Dictionary<string, int> ThaiMonths = new()
        {
            ["ม.ค."] = 1,
            ["ก.พ."] = 2,
            ["มี.ค."] = 3,
            ["เม.ย."] = 4,
            ["พ.ค."] = 5,
            ["มิ.ย."] = 6,
            ["ก.ค."] = 7,
            ["ส.ค."] = 8,
            ["ก.ย."] = 9,
            ["ต.ค."] = 10,
            ["พ.ย."] = 11,
            ["ธ.ค."] = 12
        };

        // "3 ต.ค. 2569 15:16 น." → 2026-10-03 15:16 (ลบ 543 จาก พ.ศ.)
        private static DateTime? ParseThaiDate(string? s)
        {
            var m = Regex.Match(s ?? "", @"(\d{1,2})\s+(\S+)\s+(\d{4})\s+(\d{1,2}):(\d{2})");
            if (!m.Success || !ThaiMonths.TryGetValue(m.Groups[2].Value, out var month)) return null;
            return new DateTime(int.Parse(m.Groups[3].Value) - 543, month, int.Parse(m.Groups[1].Value),
                                int.Parse(m.Groups[4].Value), int.Parse(m.Groups[5].Value), 0);
        }
    }
}
