using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MacroUploader {
    static class Program {
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args) {
            if (args != null && args.Length > 0) {
                string filePath = args[0];
                if (filePath.EndsWith(".mcro", StringComparison.Ordinal) || filePath.EndsWith(".ino", StringComparison.Ordinal)) {
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);

                    Directory.SetCurrentDirectory(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName));
                    Form1 main = new Form1(filePath);
                    Application.Run(main);
                    //Application.Run(new Form1(filePath));
                }
            }
            //Application.SetHighDpiMode(HighDpiMode.SystemAware);
            //Application.EnableVisualStyles();
            //Application.SetCompatibleTextRenderingDefault(false);
            //Application.Run(new Form1());
            Application.Exit();
        }
    }
}
