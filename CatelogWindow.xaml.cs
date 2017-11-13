/**
 * CatelogWindow.cs (c) 2017 by x01
 */
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;

using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace x01.ExcelHelper
{
	/// <summary>
	/// 为系统销售销退添加分类信息
	/// </summary>
	public partial class CatelogWindow : Window
	{
		OpenFileDialog _openDialog = new OpenFileDialog();
		
		#region Settings
		
		public string OriginPath
		{
			get {
				if (string.IsNullOrEmpty(tbxOriginPath.Text))
					throw new Exception("请选择原始文件！");
				return tbxOriginPath.Text;
			}
		}
		
		// sheet1：高中
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
		
		// sheet2: 网点
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
		
		// sheet3：城区
		public string OriginSheet3Name
		{
			get {
				return tbxOriginSheet3Name.Text;
			}
		}
		public int OriginSheet3StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet3StartRow.Text, out row))
					throw new Exception("请在表3起始行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet3EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet3EndRow.Text, out row))
					throw new Exception("请在表3结束行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet3StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet3StartCol.Text, out col))
					throw new Exception("请在表3开始列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet3EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet3EndCol.Text, out col))
					throw new Exception("请在表2结束列中填入正确的数字！");
				return col;
			}
		}
		
		// sheet4：系统销售
		public string OriginSheet4Name
		{
			get {
				return tbxOriginSheet4Name.Text;
			}
		}
		public int OriginSheet4StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet4StartRow.Text, out row))
					throw new Exception("请在表4起始行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet4EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet4EndRow.Text, out row))
					throw new Exception("请在表4结束行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet4StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet4StartCol.Text, out col))
					throw new Exception("请在表4开始列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet4EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet4EndCol.Text, out col))
					throw new Exception("请在表4结束列中填入正确的数字！");
				return col;
			}
		}
		
		// sheet5：系统销退
		public string OriginSheet5Name
		{
			get {
				return tbxOriginSheet5Name.Text;
			}
		}
		public int OriginSheet5StartRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet5StartRow.Text, out row))
					throw new Exception("请在表5起始行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet5EndRow
		{
			get {
				int row;
				if (!int.TryParse(tbxOriginSheet5EndRow.Text, out row))
					throw new Exception("请在表2结束行中填入正确的数字！");
				return row;
			}
		}
		public int OriginSheet5StartCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet5StartCol.Text, out col))
					throw new Exception("请在表5开始列中填入正确的数字！");
				return col;
			}
		}
		public int OriginSheet5EndCol
		{
			get {
				int col;
				if (!int.TryParse(tbxOriginSheet5EndCol.Text, out col))
					throw new Exception("请在表5结束列中填入正确的数字！");
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
		
		#endregion
		
		public CatelogWindow()
		{
			InitializeComponent();
			
			_openDialog.Filter = "Excel Files(*.xls)|*.xls|All Files(*.*)|*.*";
		}
		
		protected override void OnInitialized(EventArgs e)
		{
			base.OnInitialized(e);
			
			this.tbxOriginPath.Text = ConfigurationManager.AppSettings["CategoryOriginPath"];
			tbxOriginSheet1Name.Text = ConfigurationManager.AppSettings["CategoryOriginSheet1Name"];
			tbxOriginSheet1StartRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet1StartRow"];
			tbxOriginSheet1EndRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet1EndRow"];
			tbxOriginSheet1StartCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet1StartCol"];
			tbxOriginSheet1EndCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet1EndCol"];
			tbxOriginSheet2Name.Text = ConfigurationManager.AppSettings["CategoryOriginSheet2Name"];
			tbxOriginSheet2StartRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet2StartRow"];
			tbxOriginSheet2EndRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet2EndRow"];
			tbxOriginSheet2StartCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet2StartCol"];
			tbxOriginSheet2EndCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet2EndCol"];
			tbxOriginSheet3Name.Text = ConfigurationManager.AppSettings["CategoryOriginSheet3Name"];
			tbxOriginSheet3StartRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet3StartRow"];
			tbxOriginSheet3EndRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet3EndRow"];
			tbxOriginSheet3StartCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet3StartCol"];
			tbxOriginSheet3EndCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet3EndCol"];
			tbxOriginSheet4Name.Text = ConfigurationManager.AppSettings["CategoryOriginSheet4Name"];
			tbxOriginSheet4StartRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet4StartRow"];
			tbxOriginSheet4EndRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet4EndRow"];
			tbxOriginSheet4StartCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet4StartCol"];
			tbxOriginSheet4EndCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet4EndCol"];
			tbxOriginSheet5Name.Text = ConfigurationManager.AppSettings["CategoryOriginSheet5Name"];
			tbxOriginSheet5StartRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet5StartRow"];
			tbxOriginSheet5EndRow.Text = ConfigurationManager.AppSettings["CategoryOriginSheet5EndRow"];
			tbxOriginSheet5StartCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet5StartCol"];
			tbxOriginSheet5EndCol.Text = ConfigurationManager.AppSettings["CategoryOriginSheet5EndCol"];
			tbxTemplatePath.Text = ConfigurationManager.AppSettings["CategoryTemplatePath"];
			tbxTemplateSheet1Name.Text = ConfigurationManager.AppSettings["CategoryTemplateSheet1Name"];
			tbxTemplateSheet1StartRow.Text = ConfigurationManager.AppSettings["CategoryTemplateSheet1StartRow"];
			tbxTemplateSheet1EndRow.Text = ConfigurationManager.AppSettings["CategoryTemplateSheet1EndRow"];
			tbxTemplateSheet1StartCol.Text = ConfigurationManager.AppSettings["CategoryTemplateSheet1StartCol"];
			tbxTemplateSheet1EndCol.Text = ConfigurationManager.AppSettings["CategoryTemplateSheet1EndCol"];
		}
		
		void OpenOriginButton_Click(object sender, RoutedEventArgs e)
		{
			if ((bool)_openDialog.ShowDialog()) {
				tbxOriginPath.Text = _openDialog.FileName;
			}
		}
		
		void OpenTemplateButton_Click(object sender, RoutedEventArgs e)
		{
			if ((bool)_openDialog.ShowDialog()) {
				tbxTemplatePath.Text = _openDialog.FileName;
			}
		}
		
		void FillButton_Click(object sender, RoutedEventArgs e)
		{
			var oldTitle = this.Title;
			this.Title = "正在运行，请等待...";
			
			var orgBook = CreateWorkbook(OriginPath);
			var orgSheet1 = GetSheet(orgBook, OriginSheet1Name);	// 高中
			var orgSheet2 = GetSheet(orgBook, OriginSheet2Name);	// 网点
			var orgSheet3 = GetSheet(orgBook, OriginSheet3Name);	// 城区
			var destSheet4 = GetSheet(orgBook, OriginSheet4Name);	// 销售
			var destSheet5 = GetSheet(orgBook, OriginSheet5Name);	// 销退
			
			
			// Fill Sell
			FillCategory(ref orgSheet1, ref destSheet4, 
			          OriginSheet1StartRow, OriginSheet1EndRow,	OriginSheet1StartCol, OriginSheet1EndCol,
					  OriginSheet4StartRow, OriginSheet4EndRow, OriginSheet4StartCol, OriginSheet4EndCol,
					  true);
			FillCategory(ref orgSheet2, ref destSheet4, 
			          OriginSheet2StartRow, OriginSheet2EndRow,	OriginSheet2StartCol, OriginSheet2EndCol,
					  OriginSheet4StartRow, OriginSheet4EndRow, OriginSheet4StartCol, OriginSheet4EndCol,
					  true);
			FillCategory(ref orgSheet3, ref destSheet4, 
			          OriginSheet3StartRow, OriginSheet3EndRow,	OriginSheet3StartCol, OriginSheet3EndCol,
					  OriginSheet4StartRow, OriginSheet4EndRow, OriginSheet4StartCol, OriginSheet4EndCol,
					  true);
			
			// Fill Back
			FillCategory(ref orgSheet1, ref destSheet5, 
			          OriginSheet1StartRow, OriginSheet1EndRow,	OriginSheet1StartCol, OriginSheet1EndCol,
					  OriginSheet5StartRow, OriginSheet5EndRow, OriginSheet5StartCol, OriginSheet5EndCol,
					  false);
			FillCategory(ref orgSheet2, ref destSheet5, 
			          OriginSheet2StartRow, OriginSheet2EndRow,	OriginSheet2StartCol, OriginSheet2EndCol,
					  OriginSheet5StartRow, OriginSheet5EndRow, OriginSheet5StartCol, OriginSheet5EndCol,
					  false);
			FillCategory(ref orgSheet3, ref destSheet5, 
			          OriginSheet3StartRow, OriginSheet3EndRow,	OriginSheet3StartCol, OriginSheet3EndCol,
					  OriginSheet5StartRow, OriginSheet5EndRow, OriginSheet5StartCol, OriginSheet5EndCol,
					  false);
			
			orgBook.ForceFormulaRecalculation = true;
			var fs = new FileStream(OriginPath, FileMode.Open, FileAccess.ReadWrite);
			orgBook.Write(fs);
			fs.Close();
			
			MessageBox.Show("OK!");
			this.Title = oldTitle;
		}
		
		void GenerateButton_Click(object sender, RoutedEventArgs e)
		{
			var oldTitle = this.Title;
			this.Title = "正在运行，请等待...";
			
			var orgBook = CreateWorkbook(OriginPath);
			var orgSheet1 = GetSheet(orgBook, OriginSheet4Name);	// xs
			var orgSheet2 = GetSheet(orgBook, OriginSheet5Name);	// xt
			
			var tempBook = CreateWorkbook(TemplatePath);
			var destSheet = GetSheet(tempBook, TemplateSheet1Name);
			
			GenerateDetails(ref tempBook, 
			                ref orgSheet1, OriginSheet4StartRow, OriginSheet4EndRow,
			                OriginSheet4StartCol, OriginSheet4EndCol,
			                ref orgSheet2, OriginSheet5StartRow, OriginSheet5EndRow, 
			                OriginSheet5StartCol, OriginSheet5EndCol,
			               	ref destSheet, TemplateSheet1StartRow, TemplateSheet1EndRow,
			               	TemplateSheet1StartCol, TemplateSheet1EndCol);
			
			MessageBox.Show("OK!");
			this.Title = oldTitle;
		}
		
		void FillCategory(ref ISheet org, ref ISheet dest, 
		              int orgStartRow, int orgEndRow, int orgStartCol, int orgEndCol,
		              int destStartRow, int destEndRow, int destStartCol, int destEndCol,
		              bool isSell = true)
		{
			for (int orgRow = orgStartRow - 1; orgRow < orgEndRow; orgRow++) {
				for (int orgCol = orgStartCol - 1; orgCol < orgEndCol; orgCol++) {
					var orgClientIdRowCell = org.GetRow(orgStartRow-1).GetCell(orgCol);
					if (orgClientIdRowCell == null) continue;
					double orgClientIdRowValue = orgClientIdRowCell.CellType == CellType.Numeric 
						? orgClientIdRowCell.NumericCellValue : 0;
					if (orgClientIdRowValue == 0) continue;
					for (int dr = destStartRow-1; dr < destEndRow; dr++) {
						var destClientIdCell = dest.GetRow(dr).GetCell(3-1); 
						if (destClientIdCell == null) continue;
						double destClientIdValue = destClientIdCell.CellType == CellType.Numeric
							? destClientIdCell.NumericCellValue : 0;
						if (destClientIdValue == 0) continue;
						if (orgClientIdRowValue == destClientIdValue) {
							var destNrCell = dest.GetRow(dr).GetCell(destStartCol - 1);
							if (destNrCell == null)
								continue;
							var destNrValue = destNrCell.CellType == CellType.String 
								? destNrCell.StringCellValue : string.Empty;
							if (string.IsNullOrEmpty(destNrValue))
								continue;
							var destDate = isSell ? GetDateFromSellNo(destNrValue) : GetDateFromBackNo(destNrValue);
							if (destDate == DateTime.MinValue)
								continue;
							for (int or = orgStartRow - 1; or < orgEndRow; or++) {
								var orgDateColCell = org.GetRow(or).GetCell(orgStartCol); // second col is date
								if (orgDateColCell == null)
									continue;
								var orgDate = orgDateColCell.CellType == CellType.Numeric && HSSFDateUtil.IsCellDateFormatted(orgDateColCell) 
									? orgDateColCell.DateCellValue : DateTime.MinValue;
								if (orgDate == DateTime.MinValue)
									continue;
								if (destDate == orgDate) {
									var orgTotalCell = org.GetRow(or).GetCell(orgCol);
									if (orgTotalCell == null)
										continue;
									double orgTotal = orgTotalCell.CellType == CellType.Numeric 
										? orgTotalCell.NumericCellValue : 0;
									if (orgTotal == 0)
										continue;
									var destTotalCell = dest.GetRow(dr).GetCell(destEndCol - 2);
									if (destTotalCell == null)
										continue;
									double destTotal = destTotalCell.CellType == CellType.Numeric 
										? destTotalCell.NumericCellValue : 0;
									if (destTotal == 0)
										continue;
									if (isSell == false) destTotal = -destTotal;
									if (orgTotal == destTotal) {
										var orgCategoryCell = org.GetRow(or).GetCell(orgStartCol - 1);
										if (orgCategoryCell == null)
											continue;
										var orgCategory = orgCategoryCell.CellType == CellType.String 
											? orgCategoryCell.StringCellValue : string.Empty;
										if (string.IsNullOrEmpty(orgCategory))
											continue;
										var destCategoryCell = dest.GetRow(dr).GetCell(destEndCol - 1);
										if (destCategoryCell == null) {
											destCategoryCell = dest.GetRow(dr).CreateCell(destEndCol - 1);
										}
										destCategoryCell.SetCellValue(orgCategory);
									}
								}
							}
						}
					}
				}
			}
		}
		
		void GenerateDetails(ref HSSFWorkbook destBook,  
		                    ref ISheet org1, int org1StartRow, int org1EndRow, int org1StartCol, int org1EndCol,
							ref ISheet org2, int org2StartRow, int org2EndRow, int org2StartCol, int org2EndCol,
		                    ref ISheet dest, int destStartRow, int destEndRow, int destStartCol, int destEndCol)
		{
			List<string> names = new List<string>();
			for (int or = org1StartRow - 1; or < org1EndRow; or++) {
				var orgNameCell = org1.GetRow(or).GetCell(4-1); 
				if (orgNameCell == null) continue;
				string orgNameValue = orgNameCell.CellType == CellType.String
					? orgNameCell.StringCellValue : string.Empty;
				if (string.IsNullOrEmpty(orgNameValue)) continue;
				if (!names.Contains(orgNameValue))
					names.Add(orgNameValue);
			}
			
			List<CategoryModel> categories = new List<CategoryModel>();
			foreach (var name in names) {
				categories.Clear();
				for (int or = org1StartRow - 1; or < org1EndRow; or++) {
					var m = new CategoryModel();
					var nameCell = org1.GetRow(or).GetCell(4-1);
					if (nameCell == null) continue;
					string nameValue = nameCell.CellType == CellType.String
						? nameCell.StringCellValue : string.Empty;
					if (string.IsNullOrEmpty(nameValue)) continue;
					if (name == nameValue) {
						if (org1.GetRow(or).GetCell(7) == null) continue;
						m.SellNr = org1.GetRow(or).GetCell(0).StringCellValue;
						m.SellMethod = org1.GetRow(or).GetCell(1).StringCellValue;
						m.ClientId = org1.GetRow(or).GetCell(2).NumericCellValue;
						m.ClientName = org1.GetRow(or).GetCell(3).StringCellValue;
						m.TypeNr = org1.GetRow(or).GetCell(4).NumericCellValue;
						m.Number = org1.GetRow(or).GetCell(5).NumericCellValue;
						m.Total = org1.GetRow(or).GetCell(6).NumericCellValue;
						m.Category = org1.GetRow(or).GetCell(7).StringCellValue;
						categories.Add(m);
					}
				}
				
				for (int r2 = org2StartRow - 1; r2 < org2EndRow; r2++) {
					var m = new CategoryModel();
					var nameCell = org2.GetRow(r2).GetCell(4-1);
					if (nameCell == null) continue;
					string nameValue = nameCell.CellType == CellType.String
						? nameCell.StringCellValue : string.Empty;
					if (string.IsNullOrEmpty(nameValue)) continue;
					if (name == nameValue) {
						if (org2.GetRow(r2).GetCell(7) == null) continue;
						m.SellNr = org2.GetRow(r2).GetCell(0).StringCellValue;
						m.SellMethod = org2.GetRow(r2).GetCell(1).StringCellValue;
						m.ClientId = org2.GetRow(r2).GetCell(2).NumericCellValue;
						m.ClientName = org2.GetRow(r2).GetCell(3).StringCellValue;
						m.TypeNr = org2.GetRow(r2).GetCell(4).NumericCellValue;
						m.Number = -org2.GetRow(r2).GetCell(5).NumericCellValue;
						m.Total = -org2.GetRow(r2).GetCell(6).NumericCellValue;
						m.Category = org2.GetRow(r2).GetCell(7).StringCellValue;
						categories.Add(m);
					}
				}
				
				var filters = categories.Where(c=>c.SellMethod == "非免");
				var groups = filters.GroupBy(c=>c.Category);
				int dr = destStartRow -1;
				foreach (var g in groups) {
					foreach (var c in g) {
						dest.GetRow(dr).GetCell(0).SetCellValue(c.SellNr);
						dest.GetRow(dr).GetCell(1).SetCellValue(c.SellMethod);
						dest.GetRow(dr).GetCell(2).SetCellValue(c.ClientId);
						dest.GetRow(dr).GetCell(3).SetCellValue(c.ClientName);
						dest.GetRow(dr).GetCell(4).SetCellValue(c.TypeNr);
						dest.GetRow(dr).GetCell(5).SetCellValue(c.Number);
						dest.GetRow(dr).GetCell(6).SetCellValue(c.Total);
						dest.GetRow(dr).GetCell(7).SetCellValue(c.Category);
						dr++;
					}
					dest.GetRow(dr).GetCell(5).SetCellValue(g.Sum(c=>c.Number));
					dest.GetRow(dr).GetCell(6).SetCellValue(g.Sum(c=>c.Total));
					dest.GetRow(dr).GetCell(3).SetCellValue(g.First().Category + "小计");
					dr++;
				}
				dest.GetRow(destEndRow+1).GetCell(5).SetCellValue(filters.Sum(c=>c.Number));
				dest.GetRow(destEndRow+1).GetCell(6).SetCellValue(filters.Sum(c=>c.Total));
				
				dest.ForceFormulaRecalculation = true;
				var fs = new FileStream(Path.Combine(Path.GetDirectoryName(TemplatePath), name + ".xls"),FileMode.Create);
				destBook.Write(fs);
				fs.Close();
				ClearSheet(ref dest, destStartRow, destEndRow, destStartCol, destEndCol);
			}
		}
		
		void ClearSheet(ref ISheet sheet, int startRow, int endRow, int startCol, int endCol)
		{
			for (int row = startRow-1; row < endRow; row++) {
				for (int col = startCol-1; col < endCol; col++) {
					var cell = sheet.GetRow(row).GetCell(col);
					if (cell == null) continue;
					cell.SetCellType(CellType.Blank);
				}
			}
		}
		
		DateTime GetDateFromSellNo(string sellNo)
		{
			if (sellNo.Length < 15) return DateTime.MinValue;
			int year, month, day;
			if (int.TryParse(sellNo.Substring(3,2), out year)
			    && int.TryParse(sellNo.Substring(5,2), out month)
			    && int.TryParse(sellNo.Substring(7,2), out day)) {
				return new DateTime(year+2000,month,day);
			}
			return DateTime.MinValue;
		}
		
		DateTime GetDateFromBackNo(string backNo)
		{
			if (backNo.Length < 15) return DateTime.MinValue;
			int year, month, day;
			if (int.TryParse(backNo.Substring(6,4), out year)
			    && int.TryParse(backNo.Substring(10,2), out month)
			    && int.TryParse(backNo.Substring(12,2), out day)) {
				return new DateTime(year,month,day);
			}
			return DateTime.MinValue;
		}
		
		HSSFWorkbook CreateWorkbook(string path)
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
	}
}