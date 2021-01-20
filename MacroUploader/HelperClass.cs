using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace MacroUploader {
    static class HelperClass {
        public static string regex(string pattern, string text) {
            Regex re = new Regex(pattern);
            Match m = re.Match(text);
            if (m.Success) {
                return m.Value;
            } else {
                return null;
            }
        }

        public static void EmptyDirectory(this System.IO.DirectoryInfo directory) {
            foreach (System.IO.FileInfo file in directory.GetFiles()) file.Delete();
            foreach (System.IO.DirectoryInfo subDirectory in directory.GetDirectories()) subDirectory.Delete(true);
        }
    }
}

