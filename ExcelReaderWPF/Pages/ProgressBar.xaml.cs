﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelReaderWPF.CS_Files;
using System.Threading;

namespace ExcelReaderWPF
{
	/// <summary>
	/// Interaction logic for ProgressBar.xaml
	/// </summary>
	
	// Two progress bars - one for current row/all rows in a sheet, one for current sheet(index)/all sheets in a book

	public partial class ProgressBar : Page
	{
		string _File1Path;
		string _File2Path;
		string _FileOutPath;
		int _FileOutIndex;

		Task _compareTask;

		public double Progress
		{
			get { return ProgressBarObj.Value; }
			set { Dispatcher.Invoke(() => ProgressBarObj.Value = value); } //Dispatcher sets the value of the progress bar on the same thread that owns the progress bar
		}

		public double ProgressSheets
		{
			get { return ProgressBarSheets.Value; }
			set { Dispatcher.Invoke(() => ProgressBarSheets.Value = value); }
		}

		public string SheetName
		{
			get { return ProgressTextBox.Text; }
			set { Dispatcher.Invoke(() => ProgressTextBox.Text = value); }
		}

		public ProgressBar()
		{
			InitializeComponent();
		}

		public ProgressBar(string fp1, string fp2, string fpo, int foi):this()
		{
			_File1Path = fp1;
			_File2Path = fp2;
			_FileOutPath = fpo;
			_FileOutIndex = foi;

			_compareTask = Task.Run(() =>
			{
				using (var comparer = new SheetComparer(_File1Path, _File2Path, _FileOutPath, _FileOutIndex)) //After the comparer obj is done being used, the dispose function is called
				{
					comparer.CompareSheet(this);
					comparer.Dispose();
				}
			});
		}
		public ProgressBar(string folderPath1, string folderPath2, string filePathOut):this()
		{
			string[] firstAssess = Directory.GetFiles(folderPath1, "*.xlsx")
						.Select(System.IO.Path.GetFileNameWithoutExtension)
						.Select(p => p.Substring(0))
						.ToArray();
			string[] secondAssess = Directory.GetFiles(folderPath2, "*.xlsx")
				.Select(System.IO.Path.GetFileNameWithoutExtension)
				.Select(p => p.Substring(0))
				.ToArray();
			_FileOutPath = filePathOut;
			_FileOutIndex = 1;
			foreach (string file in firstAssess)
			{
				_File1Path = folderPath1 + "\\" + file + ".xlsx";
				_File2Path = folderPath2 + "\\" + (secondAssess[Array.IndexOf(secondAssess, file + " Second")]) + ".xlsx";
				_compareTask = Task.Run(() =>
				{
					using (var comparer = new SheetComparer(_File1Path, _File2Path, _FileOutPath, _FileOutIndex))
					{
						comparer.CompareSheet(this);
						//comparer.Dispose(); Redundant, using should call dispose on its own
					}
				});
				_FileOutIndex++;
			}
		}
	}
}
