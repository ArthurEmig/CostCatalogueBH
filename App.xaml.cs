using System.Runtime.InteropServices;
using System.Windows;

namespace CostsViewer
{
    public partial class App : Application
    {
        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

        protected override void OnStartup(StartupEventArgs e)
        {
            // Allocate a console for this GUI application
            AllocConsole();
            
            System.Console.WriteLine("=== CostsViewer Application Starting ===");
            System.Console.WriteLine("Debug console initialized - you can now see all debug output");
            
            base.OnStartup(e);
        }
    }
}


