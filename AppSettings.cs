using System;
using System.IO;
using Microsoft.Extensions.Configuration;

namespace BWG_WasteScraper
{
    public static class AppSettings
    {
        // 1. เติม ? ไว้ข้างหลังประเภทตัวแปร
        private static IConfiguration? _configuration;

        static AppSettings()
        {
            try
            {
                var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile("config.json", optional: false, reloadOnChange: true);

                _configuration = builder.Build();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"โหลดคอนฟิกไม่สำเร็จ: {ex.Message}");
            }
        }

        // 2. ใช้เครื่องหมาย ?. เพื่อป้องกัน App พังหากดันเกิด Error ในบล็อก catch ข้างบน
        public static string? ConnectionString => _configuration?.GetConnectionString("DefaultConnection");

        public static string? GetValue(string key) => _configuration?[key];
    }
}