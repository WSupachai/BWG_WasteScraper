using System.Windows;
using Microsoft.Data.SqlClient; // มั่นใจได้ว่าใช้ไดรเวอร์ตัวล่าสุดของ .NET

namespace BWG_WasteScraper
{
    public partial class MainWindow : Window
    {
        string? connString = AppSettings.ConnectionString;

        public MainWindow()
        {
            InitializeComponent();
        }

        /*
        private async void BtnLoginSubmit_Click(object sender, RoutedEventArgs e)
        {
            //DEV MODE
            MainScraperWindow scraperWindow = new MainScraperWindow("DEV MODE");
            scraperWindow.Show();

            this.Close();
        }
        */

        // 🎯 ฟังก์ชันดักจังหวะการกดปุ่ม "🔓 ตรวจสอบสิทธิ์เข้าใช้งาน"
        private async void BtnLoginSubmit_Click(object sender, RoutedEventArgs e)
        {        
            string user = TxtUsername.Text.Trim();
            string pass = TxtPassword.Password.Trim(); // ดึงความลับอย่างปลอดภัยจาก PasswordBox

            // ดักฟิลด์ว่างพื้นฐาน
            if (string.IsNullOrWhiteSpace(user) || string.IsNullOrWhiteSpace(pass))
            {
                MessageBox.Show("กรุณากรอก Username และ Password ให้ครบถ้วนก่อนเข้าสู่ระบบ", "ระบบความปลอดภัย", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // ล็อกปุ่มกดชั่วคราว ป้องกันการรัวปุ่มซ้ำขณะคิวรีฐานข้อมูลหมุนอยู่
            BtnLoginSubmit.IsEnabled = false;

            // รันการเช็กสิทธิ์ผ่านระบบมัลติเธรด (Async/Await) หน้าจอ Login จะได้ไม่กระตุกค้าง
            string loginResult = await CheckUserAccessAsync(user, pass);

            if (loginResult == "True")
            {
                // เปิดหน้ากาก Scraper หลัก โยนสัมภาระกรอกข้อมูลข้ามฟากไป
                MainScraperWindow scraperWindow = new MainScraperWindow(user);
                scraperWindow.Show();

                // สั่งทำลายและปิดบานหน้าต่าง Login ทิ้งเพื่อเคลียร์แรม
                this.Close();
            }
            else
            {
                // ถ้าผลลัพธ์เป็นอย่างอื่น (หรือ False) คืนชีพปุ่มให้Userลองกรอกใหม่
                BtnLoginSubmit.IsEnabled = true;
            }
        }

        private async Task<string> CheckUserAccessAsync(string userName, string password)
        {
            string result = "";
            string appId = "";

            // จำลองลอจิก Select Case จากโค้ดเดิมของพี่
            appId = "26";
 
            string version = "1.0.0"; 

            // ใช้ระบบจัดการ Context ด้วยการเปิดท่อเชื่อม SQL Connection 
            using (SqlConnection conn = new SqlConnection(connString))
            {
                try
                {
                    // เปิดท่อแบบอซิงโครนัสเพื่อไม่ให้หน้าจอ UI นิ่งค้าง
                    await conn.OpenAsync();

                    // 🎯 3. สั่งรันชุดคำสั่ง Stored Procedure (sp_CheckUserAccess) ตรงรุ่นของแท้
                    using (SqlCommand cmd = new SqlCommand("sp_CheckUserAccess", conn))
                    {
                        cmd.CommandType = System.Data.CommandType.StoredProcedure;

                        // ส่งพารามิเตอร์ 5 ตัวตรงสเปคเดิมของระบบพี่เป๊ะๆ
                        cmd.Parameters.AddWithValue("@UserName", userName);
                        cmd.Parameters.AddWithValue("@Password", password);
                        cmd.Parameters.AddWithValue("@VersionApp", version);
                        cmd.Parameters.AddWithValue("@AppID", appId);
                        cmd.Parameters.AddWithValue("@IPAddress", "");

                        // แอบเปิดอ่านข้อมูลส่งกลับแบบ Async จาก SQL Server
                        using (SqlDataReader reader = await cmd.ExecuteReaderAsync())
                        {
                            if (await reader.ReadAsync())
                            {
                                result = reader["Status"]?.ToString() ?? "No response from server";
                            }
                            else
                            {
                                result = "No response from server";
                            }
                        }

                        // ตรวจสอบค่า String ผลลัพธ์ตามกฎขององค์กรพี่
                        switch (result)
                        {
                            case "เข้าสู่ระบบสำเร็จ":
                                result = "True";
                                //MessageBox.Show("เข้าสู่ระบบสำเร็จ", "ตรวจสอบ", MessageBoxButton.OK, MessageBoxImage.Information);
                                break;
                            default:
                                // แสดง Pop-up เตือนคำสั่ง warning ตามที่ Stored Procedure ส่งกลับมาตรงๆ (เช่น รหัสผิด หรือเวอร์ชันไม่อัปเดต)
                                MessageBox.Show(result, "ตรวจสอบ", MessageBoxButton.OK, MessageBoxImage.Warning);
                                result = "False";
                                break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // ระบบดักจับการ Error เชื่อมต่อพัง พ่นรายละเอียดขึ้นเตือนทันที
                    MessageBox.Show($"Database Error: {ex.Message}", "ตรวจสอบ", MessageBoxButton.OK, MessageBoxImage.Error);
                    result = "False";
                }
            }

            return result;
        }

    }
}