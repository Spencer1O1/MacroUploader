using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO.Ports;
using System.Management;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Reflection;

namespace MacroUploader {
    public class arduino_functions {
        public static arduino info(string search) {
            arduino info = new arduino();
            string[] ports = SerialPort.GetPortNames();
            try {
                //ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_SerialPort");
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\"");
                foreach (ManagementObject queryObj in searcher.Get()) {
                    //Console.WriteLine("-----------------------------------");
                    //Console.WriteLine("Win32_SerialPort instance");
                    //Console.WriteLine("-----------------------------------");
                    //Console.WriteLine("DeviceID: {0}", queryObj["DeviceID"]);
                    //Console.WriteLine("Name: {0}", queryObj["Name"]);

                    if (queryObj["Name"].ToString().Contains(search)) {
                        string arduinoVID = HelperClass.regex("VID_([0-9a-fA-F]+)", queryObj["DeviceID"].ToString());
                        string arduinoPID = HelperClass.regex("PID_([0-9a-fA-F]+)", queryObj["DeviceID"].ToString());
                        arduinoVID = arduinoVID.Remove(0, 4);
                        arduinoPID = arduinoPID.Remove(0, 4);

                        info.vid = arduinoVID;
                        info.pid = arduinoPID;
                        info.name = queryObj["Name"].ToString();
                    }

                }
                if (searcher.Get().Count == 0) {
                    MessageBox.Show("No arduino was found.");
                    info.error = true;
                }
            } catch (ManagementException e) {
                MessageBox.Show("An error occurred while querying for WMI data: " + e.Message);
                info.error = true;
            }

            foreach (string port in ports) {
                //System.Windows.Forms.MessageBox.Show(port);
                if (info.name.Contains(port)) {
                    info.com = port;
                }
            }

            if (info.vid == "" || info.pid == "" || info.com == "") {
                info.error = true;
            }

            return info;
        }

        public static void updateArduinoPort(arduino a, string search) {
            string[] ports = SerialPort.GetPortNames();
            try {
                ManagementObjectSearcher searcher = new ManagementObjectSearcher("root\\CIMV2", "SELECT * FROM Win32_PnPEntity WHERE ClassGuid=\"{4d36e978-e325-11ce-bfc1-08002be10318}\"");
                foreach (ManagementObject queryObj in searcher.Get()) {
                    if (queryObj["Name"].ToString().Contains(search)) {
                        a.name = queryObj["Name"].ToString();
                    }
                }
                if (searcher.Get().Count == 0) {
                    MessageBox.Show("No arduino was found.");
                    a.error = true;
                }
            } catch (ManagementException e) {
                MessageBox.Show("An error occurred while querying for WMI data: " + e.Message);
                a.error = true;
            }
            foreach (string port in ports) {
                //System.Windows.Forms.MessageBox.Show(port);
                if (a.name.Contains(port)) {
                    a.com = port;
                }
            }
        }

        public static bool UploadToArduino(arduino a, string search) {
            try {
                string command4;
                string command3 = "Mode " + a.com + ": Baud=1200 Parity=N Data=8 Stop=1";
                if (a.nano) {
                    command4 = ".\\hardware\\tools\\avr\\bin\\avrdude -C .\\hardware\\tools\\avr\\etc\\avrdude.conf -v -patmega328p -c arduino -D -P" + a.com + " -b115200 -Uflash:w:.\\build\\mcrobuild.ino.hex:i";
                } else {
                    command4 = ".\\hardware\\tools\\avr\\bin\\avrdude -C .\\hardware\\tools\\avr\\etc\\avrdude.conf -v -patmega32u4 -c avr109 -D -P" + a.com + " -b57600 -Uflash:w:.\\build\\mcrobuild.ino.hex:i";
                }

                Process p2 = new Process();
                ProcessStartInfo SI2 = new ProcessStartInfo();
                SI2.UseShellExecute = true;
                SI2.WindowStyle = ProcessWindowStyle.Hidden;
                SI2.FileName = "cmd.exe";
                SI2.Arguments = ("/C " + command3);
                p2.StartInfo = SI2;
                p2.Start();
                p2.WaitForExit();

                Thread.Sleep(2000);
                updateArduinoPort(a, search);

                Process p = new Process();
                ProcessStartInfo SI = new ProcessStartInfo();
                SI.UseShellExecute = true;
                SI.WindowStyle = ProcessWindowStyle.Hidden;
                SI.FileName = "cmd.exe";
                SI.Arguments = ("/C " + command4);
                p.StartInfo = SI;
                p.Start();
                p.WaitForExit();
            } catch (Exception e) {
                return false;
            }

            return true;
        }

        public static bool CompileHex(arduino a) {
            try {
                string command1;
                string command2;
                if (a.nano) {
                    command1 = "arduino-builder -dump-prefs -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose macro.ino";
                    command2 = "arduino-builder -compile -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose macro.ino";
                } else {
                    //command1 = "arduino-builder -dump-prefs -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose macro.ino";
                    //command2 = "arduino-builder -compile -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose macro.ino";
                    command1 = "";
                    command2 = "";
                }

                Process p1 = new Process();
                ProcessStartInfo SI1 = new ProcessStartInfo();
                SI1.UseShellExecute = true;
                SI1.WindowStyle = ProcessWindowStyle.Hidden;
                SI1.FileName = "cmd.exe";
                SI1.Arguments = ("/C " + command1);
                p1.StartInfo = SI1;
                p1.Start();
                p1.WaitForExit();

                Process p2 = new Process();
                ProcessStartInfo SI2 = new ProcessStartInfo();
                SI2.UseShellExecute = true;
                SI2.WindowStyle = ProcessWindowStyle.Hidden;
                SI2.FileName = "cmd.exe";
                SI2.Arguments = ("/C " + command2);
                p2.StartInfo = SI2;
                p2.Start();
                p2.WaitForExit();

            } catch (Exception e) {
                return false;
            }
            return true;
        }

        public static bool CompileHex(arduino a, string path) {
            //HelperClass.Empty(new DirectoryInfo(@".\build"));
            string build = Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName) + @"\build";
            
            HelperClass.EmptyDirectory(new DirectoryInfo(build));
            string inopath = Path.Combine(Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName), "tobuild\\mcrobuild.ino");

            File.Copy(path, inopath, true);
            try {
                string command1;
                string command2;
                if (a.nano) {
                    command1 = "arduino-builder.exe -dump-prefs -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose " + inopath;
                    command2 = "arduino-builder.exe -compile -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose " + inopath;
                    Console.WriteLine(inopath);
                    //command1 = "";
                    //command2 = "";
                } else {
                    //command1 = "arduino-builder -dump-prefs -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose " + path;
                    //command2 = "arduino-builder -compile -logger=machine -hardware ./hardware -tools ./tools-builder -tools ./hardware/tools/avr -built-in-libraries ./builtinlibraries -libraries ./libraries -fqbn=arduino:avr:nano:cpu=atmega328 -vid-pid=" + a.vid + "_" + a.pid + " -ide-version=10813 -build-path ./build -warnings=none -build-cache ./buildcache -prefs=build.warn_data_percentage=75 -prefs=runtime.tools.arduinoOTA.path=./hardware/tools/avr -prefs=runtime.tools.arduinoOTA-1.3.0.path=./hardware/tools/avr -prefs=runtime.tools.avrdude.path=./hardware/tools/avr -prefs=runtime.tools.avrdude-6.3.0-arduino17.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc.path=./hardware/tools/avr -prefs=runtime.tools.avr-gcc-7.3.0-atmel3.6.1-arduino7.path=./hardware/tools/avr -verbose " + path;
                    command1 = "";
                    command2 = "";
                }

                Process p1 = new Process();
                ProcessStartInfo SI1 = new ProcessStartInfo();
                SI1.UseShellExecute = true;
                SI1.WindowStyle = ProcessWindowStyle.Hidden;
                SI1.FileName = "cmd.exe";
                SI1.Arguments = ("/C " + command1);
                p1.StartInfo = SI1;
                p1.Start();
                p1.WaitForExit();

                Process p2 = new Process();
                ProcessStartInfo SI2 = new ProcessStartInfo();
                SI2.UseShellExecute = true;
                SI2.WindowStyle = ProcessWindowStyle.Hidden;
                SI2.FileName = "cmd.exe";
                SI2.Arguments = ("/C " + command2);
                p2.StartInfo = SI2;
                p2.Start();
                p2.WaitForExit();

            } catch (Exception e) {
                return false;
            }
            return true;
        }

    }

    public struct arduino {
        public string vid;
        public string pid;
        public string com;
        public string name;
        public bool error;
        public bool nano;
    }
}



