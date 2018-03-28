/**
 * ${ClassName}.cs (c) 2017 by x01
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

using Microsoft.Win32;

namespace x01.ExcelHelper
{
	/// <summary>
	/// Interaction logic for Window1.xaml
	/// </summary>
	public partial class MainWindow : Window
	{
		List<SplitModel> models = new List<SplitModel>();
		
		public MainWindow()
		{
			InitializeComponent();
			
			cbxModels.Items.Add("CSV");
			cbxModels.Items.Add("Text");
			cbxModels.Items.Add("Excel");
			cbxModels.SelectedIndex=0;
			
			cbxEncodings.Items.Add("gb2312");
			cbxEncodings.Items.Add("utf8");
			cbxEncodings.SelectedIndex=0;
		}
		
		void AddButton_Click(object sender, RoutedEventArgs e)
		{
			OpenFileDialog dlg = new OpenFileDialog();
			dlg.Filter = "CSV Files(*.csv)|*.csv|Text Files(*.txt)|*.txt|All Files(*.*)|*.*";
			dlg.Multiselect = true;
			if ((bool)dlg.ShowDialog()) {
				foreach (var f in dlg.FileNames) {
					lbxFiles.Items.Add(f);
				}
			}
		}
		
		void RemoveButton_Click(object sender, RoutedEventArgs e)
		{
			int count = lbxFiles.SelectedItems.Count;
			int lastIndex = count - 1;
			for (int i = 0; i < count; i++) {
				lbxFiles.Items.Remove(lbxFiles.SelectedItems[lastIndex--]);
			}
		}
		
		void CombineButton_Click(object sender, RoutedEventArgs e)
		{
			if ((string)cbxModels.SelectedValue == "Text")
				CombineText();
			else if ((string)cbxModels.SelectedValue == "CSV")
				CombineCSV();
			else 
				throw new NotImplementedException();
		}
		void CombineCSV()
		{
			int startLine;
			if (!int.TryParse(tbxStartLine.Text, out startLine)) {
				MessageBox.Show("Start Line must be number.");
				return;
			}
			
			Encoding encoding = Encoding.ASCII;
			if ((string)cbxEncodings.SelectedValue == "gb2312")
				encoding = Encoding.GetEncoding("gb2312");
			else if ((string)cbxEncodings.SelectedValue == "uft8")
				encoding = Encoding.UTF8;
			
			foreach (string f in lbxFiles.Items) {
				var text = File.ReadAllLines(f,encoding);
				for (int i = startLine-1; i < text.Length; i++) {
					var line = text[i];
					if (line != null) {
						var cols = line.Split(',');
						var m = new SplitModel();
						m.Danhao = cols[0];
						m.Bianhao = cols[2];
						m.Pinming = cols[4];
						m.Dingjia = cols[5];
						m.Zhekou = cols[6];
						m.Shuliang = cols[9];
						if (!string.IsNullOrEmpty(m.Shuliang))
							models.Add(m);
					}
				}
			}
			
			var path = "";
			var saveDlg = new SaveFileDialog();
			saveDlg.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*";
			if ((bool)saveDlg.ShowDialog()) {
				path = saveDlg.FileName;
			}
			if (string.IsNullOrEmpty(path)) {
				MessageBox.Show("File name cannot be empty.");
				return;
			}
			
			foreach (SplitModel m in models) {
				if (string.IsNullOrEmpty(m.Shuliang)) continue;
				string s = m.Bianhao + "," +  m.Pinming + "," + m.Danhao + "," 
					+ m.Dingjia + "," +  m.Shuliang+"," + m.Zhekou  + "\r\n";
				File.AppendAllText(path,s);
			}
			MessageBox.Show("Operate success!");			
		}
		void CombineText()
		{
			int startLine;
			if (!int.TryParse(tbxStartLine.Text, out startLine)) {
				MessageBox.Show("Start Line must be number.");
				return;
			}
			
			Encoding encoding = Encoding.ASCII;
			if ((string)cbxEncodings.SelectedValue == "gb2312")
				encoding = Encoding.GetEncoding("gb2312");
			else if ((string)cbxEncodings.SelectedValue == "utf8")
				encoding = Encoding.UTF8;
			
			foreach (string f in lbxFiles.Items) {
				var text = File.ReadAllLines(f, encoding);
				for (int i = startLine-1; i < text.Length; i++) {
					var line = text[i];
					if (line != null) {
						var cols = line.Split('\t');
						var m = new SplitModel();
						m.Danhao = cols[1];
						m.Bianhao = cols[3];
						m.Pinming = cols[5];
						m.Dingjia = cols[6];
						m.Zhekou = cols[7];
						m.Shuliang = cols[10];
						if (!string.IsNullOrEmpty(m.Shuliang))
							models.Add(m);
					}
				}
			}
			
			var path = "";
			var saveDlg = new SaveFileDialog();
			saveDlg.Filter = "Text Files(*.txt)|*.txt|All Files(*.*)|*.*";
			if ((bool)saveDlg.ShowDialog()) {
				path = saveDlg.FileName;
			}
			if (string.IsNullOrEmpty(path)) {
				MessageBox.Show("File name cannot be empty.");
				return;
			}
			
			foreach (SplitModel m in models) {
				if (string.IsNullOrEmpty(m.Shuliang)) continue;
				string s = m.Bianhao + "," +  m.Pinming + "," + m.Danhao + "," 
					+ m.Dingjia + "," +  m.Shuliang+"," + m.Zhekou  + "\r\n";
				File.AppendAllText(path,s);
			}
			MessageBox.Show("Operate success!");			
		}
		
		void UpButton_Click(object sender, RoutedEventArgs e)
		{
			var item = lbxFiles.SelectedItem;
			var index = lbxFiles.SelectedIndex;
			if (index <= 0) return;
			lbxFiles.Items.Remove(item);
			lbxFiles.Items.Insert(--index,item);
		}
		
		void DownButton_Click(object sender, RoutedEventArgs e)
		{
			var item = lbxFiles.SelectedItem;
			var index = lbxFiles.SelectedIndex;
			if (index >= lbxFiles.Items.Count - 1) return;
			lbxFiles.Items.Remove(item);
			lbxFiles.Items.Insert(++index, item);
		}
		
		void GenerateButton_Click(object sender, RoutedEventArgs e)
		{
			new GenerateFreeSheetWindow().ShowDialog();
		}
		
		void CategoryButton_Click(object sender, RoutedEventArgs e)
		{
			new FillSellCatelogWindow().ShowDialog();
		}
		
	}
}