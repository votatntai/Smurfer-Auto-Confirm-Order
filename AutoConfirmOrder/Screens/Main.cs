using System.Diagnostics;
using System.Reflection;
using System.Security.Principal;

namespace AutoConfirmOrder.Screens
{
    public partial class Main : Form
    {
        public Main()
        {
            // Check if user is NOT admin
            if (!IsRunningAsAdministrator())
            {
                ProcessStartInfo processStartInfo = new ProcessStartInfo(Assembly.GetEntryAssembly().CodeBase);
                processStartInfo.UseShellExecute = true;
                processStartInfo.Verb = "runas";
                Process.Start(processStartInfo);
                System.Windows.Forms.Application.Exit();
            }
            InitializeComponent();
        }

        private bool IsRunningAsAdministrator()
        {
            // Lấy tài khoản hiện tại
            WindowsIdentity windowsIdentity = WindowsIdentity.GetCurrent();
            // Sử dụng hệ thống tài khoản của hệ điều hành window hiện tại
            WindowsPrincipal windowsPrincipal = new WindowsPrincipal(windowsIdentity);
            // Kiểm tra quyền quan trị "Administrator"
            return windowsPrincipal.IsInRole(WindowsBuiltInRole.Administrator);
        }
    }
}
