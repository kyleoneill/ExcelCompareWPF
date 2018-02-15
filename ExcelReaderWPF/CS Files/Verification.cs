using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReaderWPF.CS_Files
{
    public class Verification
    {
		public static bool VerifyFolders(string path1, string path2)
		{
			string[] firstAssess = Directory.GetFiles(path1, "*.xlsx")
				.Select(Path.GetFileNameWithoutExtension)
				.Select(p => p.Substring(0))
				.ToArray();
			string[] secondAssess = Directory.GetFiles(path2, "*.xlsx")
				.Select(Path.GetFileNameWithoutExtension)
				.Select(p => p.Substring(0))
				.ToArray();
			foreach (string file in firstAssess)
			{
				int pos = Array.IndexOf(secondAssess, file + " Second");
				if (pos < 0)
					return false;
			}
			return true;
		}
    }
}
