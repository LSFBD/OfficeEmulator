///#define DEBUG
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Configuration;

namespace OfficeEmulator
{
    
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;
        const int SW_SHOWMINIMIZED = 2;
        static void Main(string[] args)
        {
            var handle = GetConsoleWindow();
            ShowWindow(handle, SW_HIDE);
#if DEBUG
            args = new string[] { Path.Combine(@"E:\Учёба\4КурсБезопасность\Лекции", "Лекция_№_1._Введение_в_информационную_безопасность.docx") };
#endif
            if (args.Length == 0)
                return;
            var cid = ConfigurationManager.AppSettings.Get("cid").ToLower();
            if (String.IsNullOrEmpty(cid))
            {
                ShowWindow(handle, SW_SHOW);
                cid = Console.ReadLine();
                UpdateSetting("cid", cid.ToLower());
            }
            var oneDrive = Environment.GetEnvironmentVariable("OneDriveConsumer");
            var folders = args[0].Split("\\");
            var fileName = folders[^1];
            var fileFolder = String.Join("\\", folders[0..^1]);
            try
            {
                File.Copy(args[0], Path.Combine(oneDrive, fileName));
            }
            catch (Exception)
            {
                //File.Exists return false, so I have to ignore FileAlreadyExistsException
            }
            var url = $@"https://onedrive.live.com/sync?ru=https://d.docs.live.net/{cid}/";

            OpenUrl(url + ToValidURL(fileName));
        }
        private static void OpenUrl(string url)
        {
            try
            {
                Process.Start(url);
            }
            catch
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    throw;
                }
            }
        }
        private static string ToValidURL(string filepath)
        {
            return Uri.EscapeDataString(filepath).ToLower(); ;
        }
        private static void UpdateSetting(string key, string value)
        {
            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            configuration.AppSettings.Settings[key].Value = value;
            configuration.Save();

            ConfigurationManager.RefreshSection("appSettings");
        }
    }
}
