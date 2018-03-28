/**
 * SplitWindow.cs (c) 2017 by x01
 */
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Configuration;

using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace x01.ExcelHelper
{
	/// <summary>
	/// 生成免费申报表
	/// </summary>
	public partial class GenerateFreeSheetWindow : Window
	{
		#region Settings
		
		public string OriginPath
		{
			get {
				if (string.IsNullOrEmpty(tbxOriginPath.Text))
					throw new Exception("请选择原始文件！");
				return tbxOriginPath.Text;
			}
		}
		
		public string OriginSheet1Name
		{
			get {
				return tbxOriginSheet1Name.Text;
			}
		}
		public int OriginSheet1StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet1StartRow.Text, out row))
					throw new Exception("请在表1起始行中填充正确的数字！");
				return row;
			}
		}
		public int OriginSheet1EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet1EndRow.Text, out row))
					throw new Exception("请在表1结束行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet1StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet1StartCol.Text, out col))
					throw new Exception("请在表1起始列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet1EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet1EndCol.Text, out col))
					throw new Exception("请在表1结束列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet1CodeCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet1CodeCol.Text, out col))
					throw new Exception("请在表1代码列中填入正确的数字！");
				return col;
			}
		}
		
		public string OriginSheet2Name
		{
			get {
				return tbxOriginSheet2Name.Text;
			}
		}
		public int OriginSheet2StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet2StartRow.Text, out row))
					throw new Exception("请在表2起始行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet2EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet2EndRow.Text, out row))
					throw new Exception("请在表2结束行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet2StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet2StartCol.Text, out col))
					throw new Exception("请在表2开始列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet2EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet2EndCol.Text, out col))
					throw new Exception("请在表2结束列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet2CodeCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet2CodeCol.Text, out col))
					throw new Exception("请在表2代码列中填入正确的数字！");
				return col;
			}
		}
		
		public string TemplatePath
		{
			get {
				if (string.IsNullOrEmpty(tbxTemplatePath.Text))
					throw new Exception("请选择模板文件！");
				return tbxTemplatePath.Text;
			}
		}
		
		public string TemplateSheet1Name
		{
			get {
				return tbxTemplateSheet1Name.Text;
			}
		}
		public int TemplateSheet1StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxTemplateSheet1StartRow.Text, out row))
					throw new Exception("请在模板表1开始行中填入正确的数字！");
				return row;
			}
		}
		public int TemplateSheet1EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxTemplateSheet1EndRow.Text, out row))
					throw new Exception("请在模板表1结束行中填入正确的数字！");
				return row;
			}
		}
		public int TemplateSheet1StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet1StartCol.Text, out col))
					throw new Exception("请在模板表1开始列中填入正确的数字！");
				return col;
			}
		}
		public int TemplateSheet1EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet1EndCol.Text, out col))
					throw new Exception("请在模板表1结束列中填入正确的数字！");
				return col;
			}
		}
		public int TemplateSheet1CodeCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet1CodeCol.Text, out col))
					throw new Exception("请在模板表1代码列中填入正确的数字！");
				return col;
			}
		}
		
		public string TemplateSheet2Name
		{
			get {
				return tbxTemplateSheet2Name.Text;
			}
		}
		public int TemplateSheet2StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxTemplateSheet2StartRow.Text, out row))
					throw new Exception("请在模板表2开始行中填入正确的数字！");
				return row;
			}
		}
		public int TemplateSheet2EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxTemplateSheet2EndRow.Text, out row))
					throw new Exception("请在模板表2结束行中填入正确的数字！");
				return row;
			}
		}
		public int TemplateSheet2StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet2StartCol.Text, out col))
					throw new Exception("请在模板表2开始列中填入正确的数字！");
				return col;
			}
		}
		public int TemplateSheet2EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet2EndCol.Text, out col))
					throw new Exception("请在模板表2结束列中填入正确的数字！");
				return col;
			}
		}
		public int TemplateSheet2CodeCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet2CodeCol.Text, out col))
					throw new Exception("请在模板表2代码列中填入正确的数字！");
				return col;
			}
		}
		
		public string TemplateSheet3Name
		{
			get {
				return tbxTemplateSheet3Name.Text;
			}
		}
		public int TemplateSheet3StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxTemplateSheet3StartRow.Text, out row))
					throw new Exception("请在模板表3开始行中填入正确的数字！");
				return row;
			}
		}
		public int TemplateSheet3EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxTemplateSheet3EndRow.Text, out row))
					throw new Exception("请在模板表3结束行中填入正确的数字！");
				return row;
			}
		}
		public int TemplateSheet3StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet3StartCol.Text, out col))
					throw new Exception("请在模板表3开始列中填入正确的数字！");
				return col;
			}
		}
		public int TemplateSheet3EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet3EndCol.Text, out col))
					throw new Exception("请在模板表3结束列中填入正确的数字！");
				return col;
			}
		}
		public int TemplateSheet3CodeCol
		{
			get {
				int col;
				if (!int.TryParse(tbxTemplateSheet3CodeCol.Text, out col))
					throw new Exception("请在模板表3代码列中填入正确的数字！");
				return col;
			}
		}
		
		#endregion
		
		OpenFileDialog openDialog = new OpenFileDialog();
		SaveFileDialog saveDialog = new SaveFileDialog();
		public GenerateFreeSheetWindow()
		{
			InitializeComponent();
			
			openDialog.Filter = "Excel Files(*.xls)|*.xls|All Files(*.*)|*.*";
			saveDialog.Filter = "Excel Files(*.xls)| *.xls | All Files(*.*) | *.* ";
		}

		void OpenTemplateButton_Click(object sender, RoutedEventArgs e)
		{
			if ((bool)openDialog.ShowDialog()) {
				tbxTemplatePath.Text = openDialog.FileName;
			}
		}
		
		void OpenOriginButton_Click(object sender, RoutedEventArgs e)
		{
			if ((bool)openDialog.ShowDialog()) {
				tbxOriginPath.Text = openDialog.FileName;
			}
		}
		
		void GenerateFilesButton_Click(object sender, RoutedEventArgs e)
		{
			var orgBook = CreateWorkook(OriginPath);
			var orgSheet1 = GetSheet(orgBook,OriginSheet1Name);
			var orgSheet2 = GetSheet(orgBook,OriginSheet2Name);
			
			var tempBook = CreateWorkook(TemplatePath);
			var tempSheet1 = GetSheet(tempBook,TemplateSheet1Name);
			var tempSheet2 = GetSheet(tempBook,TemplateSheet2Name);
			var tempSheet3 = GetSheet(tempBook, TemplateSheet3Name);
			var tempSheetSum = GetSheet(tempBook, ConfigurationManager.AppSettings["TemplateSheet0Name"]);
			
			if (orgSheet1 != null) {
				for (int j = OriginSheet1StartCol - 1; j < OriginSheet1EndCol; j++) {
					string name = orgSheet1.GetRow(OriginSheet1StartRow-2).GetCell(j).StringCellValue;
					GenerateTemplateSheet(ref orgSheet1, ref tempSheet1, j, name, 
					                      OriginSheet1StartRow, OriginSheet1EndRow, OriginSheet1CodeCol,
					                     TemplateSheet1StartRow, TemplateSheet1EndRow, 
					                     TemplateSheet1StartCol, TemplateSheet1EndCol, TemplateSheet1CodeCol);
					GenerateTemplateSheet(ref orgSheet1, ref tempSheet2, j, name, 
					                      OriginSheet1StartRow, OriginSheet1EndRow, OriginSheet1CodeCol,
					                     TemplateSheet2StartRow, TemplateSheet2EndRow, 
					                     TemplateSheet2StartCol, TemplateSheet2EndCol, TemplateSheet2CodeCol);
					GenerateTemplateSheet(ref orgSheet1, ref tempSheet3, j, name, 
					                      OriginSheet1StartRow, OriginSheet1EndRow, OriginSheet1CodeCol,
					                     TemplateSheet3StartRow, TemplateSheet3EndRow, 
					                     TemplateSheet3StartCol, TemplateSheet3EndCol, TemplateSheet3CodeCol);
					tempSheetSum.ForceFormulaRecalculation = true;
					var fs = new FileStream(Path.Combine(Path.GetDirectoryName(TemplatePath),name+".xls"), FileMode.Create);
					tempBook.Write(fs);
					fs.Close();
					tempBook = CreateWorkook(TemplatePath);
					tempSheet1 = GetSheet(tempBook, TemplateSheet1Name);
					tempSheet2 = GetSheet(tempBook, TemplateSheet2Name);
					tempSheet3 = GetSheet(tempBook, TemplateSheet3Name);
					tempSheetSum = GetSheet(tempBook, ConfigurationManager.AppSettings["TemplateSheet0Name"]);
				}
			}
			if (orgSheet2 != null) {
				for (int j = OriginSheet2StartCol - 1; j < OriginSheet2EndCol; j++) {
					string name = orgSheet2.GetRow(OriginSheet1StartRow-2).GetCell(j).StringCellValue;
					GenerateTemplateSheet(ref orgSheet2, ref tempSheet1, j, name, 
					                      OriginSheet2StartRow, OriginSheet2EndRow, OriginSheet2CodeCol,
					                     TemplateSheet1StartRow, TemplateSheet1EndRow, 
					                     TemplateSheet1StartCol, TemplateSheet1EndCol, TemplateSheet1CodeCol);
					GenerateTemplateSheet(ref orgSheet2, ref tempSheet2, j, name, 
					                      OriginSheet2StartRow, OriginSheet2EndRow, OriginSheet2CodeCol,
					                     TemplateSheet2StartRow, TemplateSheet2EndRow, 
					                     TemplateSheet2StartCol, TemplateSheet2EndCol, TemplateSheet2CodeCol);
					GenerateTemplateSheet(ref orgSheet2, ref tempSheet3, j, name, 
					                      OriginSheet2StartRow, OriginSheet2EndRow, OriginSheet2CodeCol,
					                     TemplateSheet3StartRow, TemplateSheet3EndRow, 
					                     TemplateSheet3StartCol, TemplateSheet3EndCol, TemplateSheet3CodeCol);
					tempSheetSum.ForceFormulaRecalculation = true;
					var fs = new FileStream(Path.Combine(Path.GetDirectoryName(TemplatePath),name+".xls"), FileMode.Create);
					tempBook.Write(fs);
					fs.Close();
					tempBook = CreateWorkook(TemplatePath);
					tempSheet1 = GetSheet(tempBook, TemplateSheet1Name);
					tempSheet2 = GetSheet(tempBook, TemplateSheet2Name);
					tempSheet3 = GetSheet(tempBook, TemplateSheet3Name);
					tempSheetSum = GetSheet(tempBook, ConfigurationManager.AppSettings["TemplateSheet0Name"]);
				}
			}
			
			MessageBox.Show("OK!");
		}


		void GenerateTemplateSheet(ref ISheet orgSheet, ref ISheet tempSheet, 
		                           int orgCol, string name, 
		                           int orgStartRow, int orgEndRow, int orgCodeCol, 
		                           int tempStartRow, int tempEndRow, 
		                           int tempStartCol, int tempEndCol, int tempCodeCol)
		{
			if (tempSheet == null) return;
			
			for (int i = orgStartRow-1; i < orgEndRow; i++) {
				if (tempSheet != null) {
					for (int y = tempStartRow - 1; y < tempEndRow; y++) {
						for (int x = tempStartCol - 1; x < tempEndCol; x++) {
							if (tempSheet.GetRow(y) == null || orgSheet.GetRow(i) == null) 
								continue;
							var cell1 = tempSheet.GetRow(y).GetCell(tempCodeCol - 1);
							var cell2 = orgSheet.GetRow(i).GetCell(orgCodeCol - 1);
							if (cell1 == null || cell2 == null) continue;
							try {
								// 非数字有可能解析错误，优秀修改原始表的征订代码列。
								if (cell1.NumericCellValue == cell2.NumericCellValue) 
								{
									tempSheet.GetRow(y).GetCell(x).SetCellValue(orgSheet.GetRow(i).GetCell(orgCol).NumericCellValue);
								}
							} catch (Exception) {
								continue;
							}
						}
					}
				}
			}
			for (int y = tempStartRow - 1; y < tempEndRow; y++) {
				if (tempSheet == null)
					break;
				if (tempSheet.GetRow(y) == null) continue;
				var cell = tempSheet.GetRow(y).GetCell(tempCodeCol - 1);
				if (cell == null) continue;
				cell.SetCellType(CellType.Blank);
//				
//				var formula = tempSheet.GetRow(y).GetCell(tempCodeCol - 2);
//				if (formula == null) continue;
//				formula.SetCellType(CellType.Formula);
			}
			
			tempSheet.ForceFormulaRecalculation = true;
		}
		
		HSSFWorkbook CreateWorkook(string path)
		{
			var fs = new FileStream(path, FileMode.Open, FileAccess.Read);
			var book = new HSSFWorkbook(fs);
			fs.Close();
			return book;
		}
		
		ISheet GetSheet(HSSFWorkbook book, string sheetName)
		{
			if (string.IsNullOrEmpty(sheetName))
				return null;
			var sheet = book.GetSheet(sheetName);
			if (sheet == null)
				throw new Exception("GetSheet Error: " + sheetName);
			return sheet;
		}

		protected override void OnInitialized(EventArgs e)
		{
			base.OnInitialized(e);
			
			DefaultFill();
		}
		
		private void DefaultFill()
		{
			tbxOriginPath.Text = ConfigurationManager.AppSettings["OriginPath"];

			tbxOriginSheet1Name.Text = ConfigurationManager.AppSettings["OriginSheet1Name"];
			tbxOriginSheet1StartRow.Text = ConfigurationManager.AppSettings["OriginSheet1StartRow"];
			tbxOriginSheet1EndRow.Text = ConfigurationManager.AppSettings["OriginSheet1EndRow"];
			tbxOriginSheet1StartCol.Text = ConfigurationManager.AppSettings["OriginSheet1StartCol"];
			tbxOriginSheet1EndCol.Text = ConfigurationManager.AppSettings["OriginSheet1EndCol"];
			tbxOriginSheet1CodeCol.Text = ConfigurationManager.AppSettings["OriginSheet1CodeCol"];

			tbxOriginSheet2Name.Text = ConfigurationManager.AppSettings["OriginSheet2Name"];
			tbxOriginSheet2StartRow.Text = ConfigurationManager.AppSettings["OriginSheet2StartRow"];
			tbxOriginSheet2EndRow.Text = ConfigurationManager.AppSettings["OriginSheet2EndRow"];
			tbxOriginSheet2StartCol.Text = ConfigurationManager.AppSettings["OriginSheet2StartCol"];
			tbxOriginSheet2EndCol.Text = ConfigurationManager.AppSettings["OriginSheet2EndCol"];
			tbxOriginSheet2CodeCol.Text = ConfigurationManager.AppSettings["OriginSheet2CodeCol"];

			tbxTemplatePath.Text = ConfigurationManager.AppSettings["TemplatePath"];

			tbxTemplateSheet1Name.Text = ConfigurationManager.AppSettings["TemplateSheet1Name"];
			tbxTemplateSheet1StartRow.Text = ConfigurationManager.AppSettings["TemplateSheet1StartRow"];
			tbxTemplateSheet1EndRow.Text = ConfigurationManager.AppSettings["TemplateSheet1EndRow"];
			tbxTemplateSheet1StartCol.Text = ConfigurationManager.AppSettings["TemplateSheet1StartCol"];
			tbxTemplateSheet1EndCol.Text = ConfigurationManager.AppSettings["TemplateSheet1EndCol"];
			tbxTemplateSheet1CodeCol.Text = ConfigurationManager.AppSettings["TemplateSheet1CodeCol"];

			tbxTemplateSheet2Name.Text = ConfigurationManager.AppSettings["TemplateSheet2Name"];
			tbxTemplateSheet2StartRow.Text = ConfigurationManager.AppSettings["TemplateSheet2StartRow"];
			tbxTemplateSheet2EndRow.Text = ConfigurationManager.AppSettings["TemplateSheet2EndRow"];
			tbxTemplateSheet2StartCol.Text = ConfigurationManager.AppSettings["TemplateSheet2StartCol"];
			tbxTemplateSheet2EndCol.Text = ConfigurationManager.AppSettings["TemplateSheet2EndCol"];
			tbxTemplateSheet2CodeCol.Text = ConfigurationManager.AppSettings["TemplateSheet2CodeCol"];


			tbxTemplateSheet3Name.Text = ConfigurationManager.AppSettings["TemplateSheet3Name"];
			tbxTemplateSheet3StartRow.Text = ConfigurationManager.AppSettings["TemplateSheet3StartRow"];
			tbxTemplateSheet3EndRow.Text = ConfigurationManager.AppSettings["TemplateSheet3EndRow"];
			tbxTemplateSheet3StartCol.Text = ConfigurationManager.AppSettings["TemplateSheet3StartCol"];
			tbxTemplateSheet3EndCol.Text = ConfigurationManager.AppSettings["TemplateSheet3EndCol"];
			tbxTemplateSheet3CodeCol.Text = ConfigurationManager.AppSettings["TemplateSheet3CodeCol"];

		}
	}
}