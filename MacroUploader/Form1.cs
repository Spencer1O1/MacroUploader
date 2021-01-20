using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MacroUploader {
    public partial class Form1 : Form {
        arduino connected;
        string path;
        public Form1(string path) {
            InitializeComponent();
            this.path = path;
            //LoadAndUpload(path, true);
        }
        
        public void LoadAndUpload(string path, bool nano = false) {
            connected = arduino_functions.info("SERIAL");
            connected.nano = nano;
            Console.WriteLine(connected.vid);
            Console.WriteLine(connected.pid);
            Console.WriteLine(connected.com);
            arduino_functions.CompileHex(connected, path);
            arduino_functions.UploadToArduino(connected, "SERIAL");
            
            Application.Exit();
        }

        private void Form1_Shown(object sender, EventArgs e) {
            LoadAndUpload(path, true);
        }
    }
}
