using AutoConfirmOrder.Entities;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Diagnostics;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;

namespace AutoConfirmOrder.Screens
{
    partial class Main
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        // Thêm một WebBrowser control
        private System.Windows.Forms.WebBrowser webBrowser;

        // Khởi tạo ChromeDriver
        private IWebDriver driver;

        // Logs file
        private string LogFilePath = "Logs";
        private string LogFileName = "logs.txt";

        // Reload time
        private const int ReloadTime = 1000;

        // Loop
        private bool loop = true;

        // Default notification email
        private readonly string EMAIL = "votantaiclone08@gmail.com";
        private readonly string RECEIVE_EMAIL = "smurfernoti@gmail.com";
        private readonly string PASSWORD = "skxswehqlmchqmih";

        // Default data
        private readonly string VALID_DETAILS = "Diamond V, Diamond IV, Diamond III";
        private readonly string IGNORE_DETAILS = "Master, Challenger, Grand Master, Coaching, Paused, BR, LAS, LAN, TR, Iron, Normals, Unranked, Bronze, Diamond, Silver";
        private readonly string VALID_ADDONS = "";
        private readonly string IGNORE_ADDONS = "Duo with Booster";

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            emailLabel = new Label();
            emailTextBox = new TextBox();
            passwordLabel = new Label();
            passwordTextBox = new TextBox();
            receiveEmailLabel = new Label();
            receiveEmailTextBox = new TextBox();
            getDetailLabel = new Label();
            getDetailTextBox = new TextBox();
            ignoreDetailLabel = new Label();
            ignoreDetailTextbox = new TextBox();
            getAddonLabel = new Label();
            getAddonTextBox = new TextBox();
            ignoreAddonLabel = new Label();
            ignoreAddonTextBox = new TextBox();
            loginLinkLabel = new Label();
            loginLinkTextBox = new TextBox();
            runButton = new Button();
            openLogsButton = new Button();
            webBrowser = new WebBrowser();
            acceptDelayLabel = new Label();
            acceptDelayTextBox = new TextBox();
            SuspendLayout();
            // 
            // emailLabel
            // 
            emailLabel.BackColor = Color.White;
            emailLabel.Location = new Point(10, 10);
            emailLabel.Name = "emailLabel";
            emailLabel.Size = new Size(300, 20);
            emailLabel.TabIndex = 0;
            emailLabel.Text = "Email gửi thông báo *";
            // 
            // emailTextBox
            // 
            emailTextBox.Location = new Point(10, 30);
            emailTextBox.Name = "emailTextBox";
            emailTextBox.Text = EMAIL;
            emailTextBox.Size = new Size(300, 23);
            emailTextBox.TabIndex = 1;
            // 
            // passwordLabel
            // 
            passwordLabel.BackColor = Color.White;
            passwordLabel.Location = new Point(10, 64);
            passwordLabel.Name = "passwordLabel";
            passwordLabel.Size = new Size(300, 17);
            passwordLabel.TabIndex = 2;
            passwordLabel.Text = "Mật khẩu *";
            // 
            // passwordTextBox
            // 
            passwordTextBox.Location = new Point(10, 84);
            passwordTextBox.Name = "passwordTextBox";
            passwordTextBox.Text = PASSWORD;
            passwordTextBox.PasswordChar = '*';
            passwordTextBox.Size = new Size(300, 23);
            passwordTextBox.TabIndex = 3;
            // 
            // receiveEmailLabel
            // 
            receiveEmailLabel.BackColor = Color.White;
            receiveEmailLabel.Location = new Point(8, 600);
            receiveEmailLabel.Name = "receiveEmailLabel";
            receiveEmailLabel.Size = new Size(300, 20);
            receiveEmailLabel.TabIndex = 16;
            receiveEmailLabel.Text = "Email nhận thông báo *";
            // 
            // receiveEmailTextBox
            // 
            receiveEmailTextBox.Location = new Point(8, 620);
            receiveEmailTextBox.Name = "receiveEmailTextBox";
            receiveEmailTextBox.Text = RECEIVE_EMAIL;
            receiveEmailTextBox.Size = new Size(300, 23);
            receiveEmailTextBox.TabIndex = 17;
            // 
            // getDetailLabel
            // 
            getDetailLabel.BackColor = Color.White;
            getDetailLabel.Location = new Point(8, 122);
            getDetailLabel.Name = "getDetailLabel";
            getDetailLabel.Size = new Size(300, 17);
            getDetailLabel.TabIndex = 4;
            getDetailLabel.Text = "DETAILS - Lấy";
            // 
            // getDetailTextBox
            // 
            getDetailTextBox.Location = new Point(8, 142);
            getDetailTextBox.Multiline = true;
            getDetailTextBox.Name = "getDetailTextBox";
            getDetailTextBox.Text = VALID_DETAILS;
            getDetailTextBox.ScrollBars = ScrollBars.Vertical;
            getDetailTextBox.Size = new Size(300, 54);
            getDetailTextBox.TabIndex = 5;
            // 
            // ignoreDetailLabel
            // 
            ignoreDetailLabel.BackColor = Color.White;
            ignoreDetailLabel.Location = new Point(10, 211);
            ignoreDetailLabel.Name = "ignoreDetailLabel";
            ignoreDetailLabel.Size = new Size(300, 17);
            ignoreDetailLabel.TabIndex = 6;
            ignoreDetailLabel.Text = "DETAILS - Bỏ";
            // 
            // ignoreDetailTextbox
            // 
            ignoreDetailTextbox.Location = new Point(10, 231);
            ignoreDetailTextbox.Multiline = true;
            ignoreDetailTextbox.Name = "ignoreDetailTextbox";
            ignoreDetailTextbox.Text = IGNORE_DETAILS;
            ignoreDetailTextbox.ScrollBars = ScrollBars.Vertical;
            ignoreDetailTextbox.Size = new Size(300, 54);
            ignoreDetailTextbox.TabIndex = 7;
            // 
            // getAddonLabel
            // 
            getAddonLabel.BackColor = Color.White;
            getAddonLabel.Location = new Point(10, 307);
            getAddonLabel.Name = "getAddonLabel";
            getAddonLabel.Size = new Size(300, 17);
            getAddonLabel.TabIndex = 8;
            getAddonLabel.Text = "ADDONS - Lấy";
            // 
            // getAddonTextBox
            // 
            getAddonTextBox.Location = new Point(10, 327);
            getAddonTextBox.Multiline = true;
            getAddonTextBox.Name = "getAddonTextBox";
            getAddonTextBox.Text = VALID_ADDONS;
            getAddonTextBox.ScrollBars = ScrollBars.Vertical;
            getAddonTextBox.Size = new Size(300, 54);
            getAddonTextBox.TabIndex = 9;
            // 
            // ignoreAddonLabel
            // 
            ignoreAddonLabel.BackColor = Color.White;
            ignoreAddonLabel.Location = new Point(10, 394);
            ignoreAddonLabel.Name = "ignoreAddonLabel";
            ignoreAddonLabel.Size = new Size(300, 17);
            ignoreAddonLabel.TabIndex = 10;
            ignoreAddonLabel.Text = "ADDONS - Bỏ";
            // 
            // ignoreAddonTextBox
            // 
            ignoreAddonTextBox.Location = new Point(10, 414);
            ignoreAddonTextBox.Multiline = true;
            ignoreAddonTextBox.Name = "ignoreAddonTextBox";
            ignoreAddonTextBox.Text = IGNORE_ADDONS;
            ignoreAddonTextBox.ScrollBars = ScrollBars.Vertical;
            ignoreAddonTextBox.Size = new Size(300, 54);
            ignoreAddonTextBox.TabIndex = 11;
            // 
            // loginLinkLabel
            // 
            loginLinkLabel.BackColor = Color.White;
            loginLinkLabel.Location = new Point(8, 480);
            loginLinkLabel.Name = "loginLinkLabel";
            loginLinkLabel.Size = new Size(300, 20);
            loginLinkLabel.TabIndex = 12;
            loginLinkLabel.Text = "Link đăng nhập *";
            // 
            // loginLinkTextBox
            // 
            loginLinkTextBox.Location = new Point(8, 500);
            loginLinkTextBox.Name = "loginLinkTextBox";
            loginLinkTextBox.Size = new Size(300, 23);
            loginLinkTextBox.TabIndex = 13;
            // 
            // runButton
            // 
            runButton.Location = new Point(12, 665);
            runButton.Name = "runButton";
            runButton.Size = new Size(120, 23);
            runButton.TabIndex = 18;
            runButton.Text = "Run";
            runButton.Click += RunButton_Click;
            // 
            // openLogsButton
            // 
            openLogsButton.Location = new Point(188, 665);
            openLogsButton.Name = "openLogsButton";
            openLogsButton.Size = new Size(120, 23);
            openLogsButton.TabIndex = 19;
            openLogsButton.Text = "Logs";
            openLogsButton.Click += OpenLogsButton_Click;
            // 
            // webBrowser
            // 
            webBrowser.Dock = DockStyle.Fill;
            webBrowser.Location = new Point(0, 0);
            webBrowser.Name = "webBrowser";
            webBrowser.Size = new Size(320, 700);
            webBrowser.TabIndex = 20;
            // 
            // acceptDelayLabel
            // 
            acceptDelayLabel.BackColor = Color.White;
            acceptDelayLabel.Location = new Point(8, 538);
            acceptDelayLabel.Name = "acceptDelayLabel";
            acceptDelayLabel.Size = new Size(300, 20);
            acceptDelayLabel.TabIndex = 14;
            acceptDelayLabel.Text = "Thời gian delay khi Accept *";
            // 
            // acceptDelayTextBox
            // 
            acceptDelayTextBox.Location = new Point(8, 558);
            acceptDelayTextBox.Name = "acceptDelayTextBox";
            acceptDelayTextBox.Size = new Size(300, 23);
            acceptDelayTextBox.TabIndex = 15;
            acceptDelayTextBox.Text = "1000";
            // 
            // Main
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(320, 710);
            Controls.Add(acceptDelayLabel);
            Controls.Add(acceptDelayTextBox);
            Controls.Add(loginLinkLabel);
            Controls.Add(loginLinkTextBox);
            Controls.Add(ignoreAddonLabel);
            Controls.Add(ignoreAddonTextBox);
            Controls.Add(getAddonLabel);
            Controls.Add(getAddonTextBox);
            Controls.Add(ignoreDetailLabel);
            Controls.Add(ignoreDetailTextbox);
            Controls.Add(getDetailLabel);
            Controls.Add(getDetailTextBox);
            Controls.Add(passwordLabel);
            Controls.Add(passwordTextBox);
            Controls.Add(emailLabel);
            Controls.Add(emailTextBox);
            Controls.Add(receiveEmailLabel);
            Controls.Add(receiveEmailTextBox);
            Controls.Add(runButton);
            Controls.Add(openLogsButton);
            Controls.Add(webBrowser);
            Name = "Main";
            Text = "Auto Confirm Order 2.0";
            Load += MainForm_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        // Sự kiện được gọi khi form được tải
        private void MainForm_Load(object sender, EventArgs e)
        {
            // Create log file
            LogFileName = $"log-{DateTime.Now.ToString("yyyy-MM-dd-HH-mm")}.txt";

            // Hiển thị cửa sổ Console
            AllocConsole();
        }

        // Hàm để khởi tạo và thao tác với Selenium
        private void InitializeAndOperateSelenium()
        {
            try
            {
                // Lấy đường dẫn của chromedriver.exe trong thư mục đầu ra của ứng dụng
                string chromeDriverPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "chromedriver.exe");

                // Khởi động chrome driver
                driver = new ChromeDriver(Path.Combine(chromeDriverPath));

                // Đặt đường dẫn đến chromedriver.exe
                // Thay đổi đường dẫn tùy theo vị trí chromedriver.exe trên máy của bạn
                //string chromeDriverPath = "C:\\Users\\Janglee\\Desktop\\Driver\\chromedriver.exe";

                // Khởi tạo ChromeDriver
                //driver = new ChromeDriver(chromeDriverPath);

                // Mở trình duyệt và điều hướng đến trang web mong muốn
                string signInLink = loginLinkTextBox.Text.Trim();
                driver.Navigate().GoToUrl(signInLink);

                // Đợi trình duyệt load
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                // Tìm phần tử boosting order button theo XPath và Click 
                string boostingOrderButtonXpath = "/html[1]/body[1]/div[1]/div[1]/astro-island[1]/aside[1]/div[1]/ul[1]/li[4]/a[1]";
                IWebElement boostingOrderButton = driver.FindElement(By.XPath(boostingOrderButtonXpath));
                boostingOrderButton.Click();

                while (loop)
                {
                    int count = 0;

                    // Tìm phần tử table
                    string tableXpath = "/html[1]/body[1]/div[1]/div[1]/astro-island[2]/div[1]/div[2]/table[1]";
                    IWebElement table = driver.FindElement(By.XPath(tableXpath));

                    IWebElement tbody = table.FindElement(By.XPath(".//tbody"));
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(0);
                    var totalRows = tbody.FindElements(By.TagName("tr")).Count;

                    // Chọn ra những order thoả điều kiện
                    if (totalRows > 0)
                    {
                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                        // Lấy danh sách orders trong table
                        ICollection<IWebElement> orders = table.FindElements(By.XPath(".//tbody//tr"));
                        foreach (IWebElement order in orders)
                        {
                            count += 1;
                            // Thêm những detail vào 1 danh sách để dùng cho việc so sánh
                            ICollection<string> details = new List<string>();

                            // Lấy danh sách details của order
                            IWebElement server = order.FindElement(By.XPath(".//td[2]//div//div[2]"));
                            details.Add(server.GetAttribute("innerHTML"));
                            ICollection<IWebElement> conditions = order.FindElements(By.XPath(".//td[2]//div//div[1]"));
                            foreach (var condition in conditions)
                            {
                                details.Add(condition.GetAttribute("innerHTML"));
                            }

                            // Lấy giá của order
                            IWebElement price = order.FindElement(By.XPath(".//td[6]"));
                            string orderPrice = price.GetAttribute("innerHTML");

                            // Lấy danh sách addons của order
                            IWebElement addon = order.FindElement(By.XPath(".//td[3]"));
                            details = details.Concat(ConvertStringToList(addon.GetAttribute("innerHTML"))).ToList();

                            var now = DateTime.UtcNow.AddHours(7);
                            StringBuilder logBuilder = new StringBuilder();

                            // Logs
                            logBuilder.Append($"ORDER - {count} - {orderPrice}");
                            // Console
                            Console.Write(now + " - ");
                            Console.ForegroundColor = ConsoleColor.Yellow;
                            Console.Write("ORDER - " + count + " - " + orderPrice);
                            Console.ResetColor();
                            foreach (var detail in details)
                            {
                                // Logs
                                logBuilder.Append(" - ");
                                logBuilder.Append(detail);
                                // Console
                                Console.Write(" - ");
                                Console.Write(detail);
                            }

                            List<string> ignoreList = ConvertStringToList(ignoreDetailTextbox.Text);
                            List<string> validList = ConvertStringToList(getDetailTextBox.Text);

                            List<string> ignoreAddons = ConvertStringToList(ignoreAddonTextBox.Text);
                            List<string> validAddons = ConvertStringToList(getAddonTextBox.Text);

                            // Logs
                            logBuilder.AppendLine($" - VALID - {ContainsAnyPartialIgnoreCase(details.ToList(), ignoreList.Concat(ignoreAddons).ToList(), validList.Concat(validAddons).ToList())}");
                            // Console
                            Console.WriteLine($" - VALID - {ContainsAnyPartialIgnoreCase(details.ToList(), ignoreList.Concat(ignoreAddons).ToList(), validList.Concat(validAddons).ToList())}");
                            Console.WriteLine();
                            // ghi vào file logs.txt
                            WriteLog(logBuilder.ToString());

                            // Nếu không có từ khoá nào nằm trong danh sách ngoại lệ thì click accept button
                            if (ContainsAnyPartialIgnoreCase(details.ToList(), ignoreList.Concat(ignoreAddons).ToList(), validList.Concat(validAddons).ToList()))
                            {
                                IWebElement menuButton = order.FindElement(By.XPath(".//td[7]/button[1]"));
                                menuButton.Click();

                                string orderIdXpath = "/html[1]/body[1]/div[4]/div[2]/div[1]/h2[1]";
                                IWebElement orderId = driver.FindElement(By.XPath(orderIdXpath));
                                string orderIdValue = orderId.GetAttribute("innerHTML");

                                Order orderInfo = GetOrderInformation();
                                orderInfo.Price = orderPrice;

                                // Ấn 3 chấm xong đợi
                                System.Threading.Thread.Sleep(int.Parse(acceptDelayTextBox.Text));

                                try
                                {
                                    // Accept
                                    string acceptButtonXpath = "/html[1]/body[1]/div[4]/div[2]/div[3]/button[1]";
                                    IWebElement acceptButton = driver.FindElement(By.XPath(acceptButtonXpath));
                                    acceptButton.Click();

                                    // Done
                                    string doneButtonXpath = "/html[1]/body[1]/div[4]/div[1]/div[6]/button[1]";
                                    IWebElement doneButton = driver.FindElement(By.XPath(doneButtonXpath));
                                    doneButton.Click();

                                    // Lấy giá trị từ TextBoxes
                                    string fromEmail = emailTextBox.Text;
                                    string toEmail = receiveEmailTextBox.Text;
                                    string emailPassword = passwordTextBox.Text;

                                    // Gọi hàm để gửi email sử dụng SMTP
                                    SendEmail(fromEmail, toEmail, emailPassword, orderIdValue, orderInfo);

                                    break;
                                }
                                catch (NoSuchElementException ex)
                                {
                                    Console.WriteLine(ex.Message);
                                }
                            }
                        }
                    }
                    // Write log if order not found
                    else
                    {
                        var now = DateTime.UtcNow.AddHours(7);
                        StringBuilder logBuilder = new StringBuilder();

                        // Logs
                        logBuilder.Append($"{now} - NO ELEMENTS FOUND\n");

                        // Console
                        Console.Write($"{now} - ");
                        Console.ForegroundColor = ConsoleColor.Blue;
                        Console.Write("NO ELEMENTS FOUND\n");
                        Console.ResetColor();

                        WriteLog(logBuilder.ToString());
                    }

                    // Quét lại orders sau mỗi 0.5 giây
                    Thread.Sleep(500);
                }
            }
            catch (Exception ex)
            {
                // Xử lý exception nếu có
                runButton.Enabled = true;
                EnabledTextBox(true);
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("SOFTWARE EXCEPTION:");
                Console.WriteLine(ex.Message);
                Console.ResetColor();
            }
        }

        // Elements
        private Label emailLabel;
        private TextBox emailTextBox;
        private Label passwordLabel;
        private TextBox passwordTextBox;
        private Label receiveEmailLabel;
        private TextBox receiveEmailTextBox;
        private Label getDetailLabel;
        private TextBox getDetailTextBox;
        private Label ignoreDetailLabel;
        private TextBox ignoreDetailTextbox;
        private Label getAddonLabel;
        private TextBox getAddonTextBox;
        private Label ignoreAddonLabel;
        private TextBox ignoreAddonTextBox;
        private Label loginLinkLabel;
        private TextBox loginLinkTextBox;
        private Label acceptDelayLabel;
        private TextBox acceptDelayTextBox;
        private Button runButton;
        private Button openLogsButton;

        // Sự kiện được gọi khi nút "Run" được nhấn
        private async void RunButton_Click(object sender, EventArgs e)
        {
            runButton.Enabled = false;
            EnabledTextBox(false);
            WriteLog("Janglee - Linh Code Thue - 0348 04 08 99\n");
            await Task.Run(() => InitializeAndOperateSelenium());
        }

        // Sự kiện được gọi khi nút "Logs" được nhấn
        private void OpenLogsButton_Click(object sender, EventArgs e)
        {
            // Thực hiện các thao tác khi nút "Logs" được nhấn
            OpenLogs();
        }

        // Hàm để gửi email sử dụng SMTP
        private void SendEmail(string fromEmail, string toEmail, string emailPassword, string id, Order order)
        {
            try
            {
                // Tiêu đề email
                string subject = $"[{id}] - CONFIRMED";
                // Nội dung email
                string body = BuildHtmlMailBody(order);

                SmtpClient client = new SmtpClient("smtp.gmail.com");
                client.Port = 587;
                client.UseDefaultCredentials = false;
                client.Credentials = new NetworkCredential(fromEmail, emailPassword);
                client.EnableSsl = true;

                // Khởi tạo đối tượng MailMessage
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(fromEmail);
                mailMessage.To.Add(toEmail);
                mailMessage.Subject = subject;
                mailMessage.Body = body;
                mailMessage.IsBodyHtml = true;

                // Gửi email
                client.Send(mailMessage);

                Console.WriteLine("Email sent successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error sending email: " + ex.Message);
            }
        }

        private List<string> ConvertStringToList(string inputString)
        {
            // Chia chuỗi thành danh sách các chuỗi và loại bỏ khoảng trắng
            List<string> resultList = inputString.Split(',').Select(value => value.Trim()).ToList();
            return resultList;
        }

        private void WriteLog(string logMessage)
        {
            try
            {
                // Specify the path to the log folder
                string logFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, LogFilePath);

                // Create the log folder if it doesn't exist
                if (!Directory.Exists(logFolderPath))
                {
                    Directory.CreateDirectory(logFolderPath);
                }

                // Specify the path to the log file within the log folder
                string logFilePath = Path.Combine(logFolderPath, LogFileName);

                // Create the log file if it doesn't exist
                if (!File.Exists(logFilePath))
                {
                    File.Create(logFilePath).Close();
                }

                // Append the log message to the file
                File.AppendAllText(logFilePath, $"{DateTime.UtcNow.AddHours(7)} - {logMessage}\n");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error writing log: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private bool ContainsAnyPartialIgnoreCase(List<string> valuesToCheck, List<string> ignoreList, List<string> validList)
        {
            if (!ignoreList.Any() && !validList.Any())
            {
                return false;
            }

            if (ignoreList.Any() && !validList.Any())
            {
                // Kiểm tra xem có bất kỳ phần tử nào của valuesToCheck thuộc ignoreList không
                return valuesToCheck.Any(value => ignoreList.Contains(value));
            }

            if (ignoreList.Any() && validList.Any())
            {
                // Lọc những phần tử trong valuesToCheck thuộc ignoreList
                var ignoredValues = valuesToCheck.Where(valueToCheck => ignoreList.Any(ignoreItem => valueToCheck.Contains(ignoreItem)));

                // Kiểm tra xem có bất kỳ phần tử nào thuộc ignoreList mà không thuộc validList không
                return !ignoredValues.Any() || ignoredValues.All(value => validList.Contains(value));
            }

            return false;
        }

        private void OpenLogs()
        {
            try
            {
                // Specify the path to the log file
                string logFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, LogFilePath + "\\" + LogFileName);

                // Create a new ProcessStartInfo
                ProcessStartInfo psi = new ProcessStartInfo(logFilePath)
                {
                    UseShellExecute = true
                };

                // Start the process
                new Process { StartInfo = psi }.Start();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening logs: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private Order GetOrderInformation()
        {
            try
            {
                string gameXPath = "/html[1]/body[1]/div[4]/div[2]/div[2]/div[1]/span[1]/span[2]";
                IWebElement game = driver.FindElement(By.XPath(gameXPath));
                string gameValue = game.GetAttribute("innerHTML");

                string typeXPath = "/html[1]/body[1]/div[4]/div[2]/div[2]/div[1]/span[2]/span[2]";
                IWebElement type = driver.FindElement(By.XPath(typeXPath));
                string typeValue = type.GetAttribute("innerHTML");

                string regionXPath = "/html[1]/body[1]/div[4]/div[2]/div[2]/div[1]/span[3]/span[2]";
                IWebElement region = driver.FindElement(By.XPath(regionXPath));
                string regionValue = region.GetAttribute("innerHTML");

                string queueTypeXPath = "/html[1]/body[1]/div[4]/div[2]/div[2]/div[1]/span[4]/span[2]";
                IWebElement queueType = driver.FindElement(By.XPath(queueTypeXPath));
                string queueTypeValue = queueType.GetAttribute("innerHTML");

                string rankXPath = "/html[1]/body[1]/div[4]/div[2]/div[2]/div[1]/span[5]/span[2]";
                IWebElement rank = driver.FindElement(By.XPath(rankXPath));
                string rankValue = rank.GetAttribute("innerHTML");

                DateTime acceptAt = DateTime.UtcNow.AddHours(7);

                Order order = new Order
                {
                    AcceptAt = acceptAt,
                    Game = gameValue,
                    Type = typeValue,
                    Region = regionValue,
                    QueueType = queueTypeValue,
                    Rank = rankValue
                };

                return order;
            }
            catch (Exception)
            {
                throw;
            }
        }

        private string BuildHtmlMailBody(Order order)
        {
            StringBuilder htmlBuilder = new StringBuilder();

            htmlBuilder.AppendLine("<!DOCTYPE html>");
            htmlBuilder.AppendLine("<html lang=\"en\">");
            htmlBuilder.AppendLine("<head>");
            htmlBuilder.AppendLine("<meta charset=\"UTF-8\">");
            htmlBuilder.AppendLine("<meta http-equiv=\"X-UA-Compatible\" content=\"IE=edge\">");
            htmlBuilder.AppendLine("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
            htmlBuilder.AppendLine("<style>");
            htmlBuilder.AppendLine("table { width: 100%; border-collapse: collapse; }");
            htmlBuilder.AppendLine("th, td { border: 1px solid #dddddd; text-align: left; padding: 8px; }");
            htmlBuilder.AppendLine("th { background-color: #f2f2f2; }");
            htmlBuilder.AppendLine("</style>");
            htmlBuilder.AppendLine("</head>");
            htmlBuilder.AppendLine("<body>");

            htmlBuilder.AppendLine("<table>");
            htmlBuilder.AppendLine($"<tr><th>GAME</th><td>{order.Game}</td></tr>");
            htmlBuilder.AppendLine($"<tr><th>TYPE</th><td>{order.Type}</td></tr>");
            htmlBuilder.AppendLine($"<tr><th>REGION</th><td>{order.Region}</td></tr>");
            htmlBuilder.AppendLine($"<tr><th>QUEUE TYPE</th><td>{order.QueueType}</td></tr>");
            htmlBuilder.AppendLine($"<tr><th>RANK</th><td>{order.Rank}</td></tr>");
            htmlBuilder.AppendLine($"<tr><th>ACCEPT AT</th><td>{order.AcceptAt.ToString("yyyy-MM-dd HH:mm")} PM</td></tr>");
            htmlBuilder.AppendLine($"<tr><th>PRICE</th><td>{order.Price}</td></tr>");

            htmlBuilder.AppendLine("</table>");

            htmlBuilder.AppendLine("</body>");
            htmlBuilder.AppendLine("</html>");

            return htmlBuilder.ToString();
        }

        private void EnabledTextBox(bool isEnabled)
        {
            emailTextBox.Enabled = isEnabled;
            getAddonTextBox.Enabled = isEnabled;
            getDetailTextBox.Enabled = isEnabled;
            ignoreAddonTextBox.Enabled = isEnabled;
            ignoreDetailTextbox.Enabled = isEnabled;
            loginLinkTextBox.Enabled = isEnabled;
            receiveEmailTextBox.Enabled = isEnabled;
            acceptDelayTextBox.Enabled = isEnabled;
            passwordTextBox.Enabled = isEnabled;
        }
    }
}