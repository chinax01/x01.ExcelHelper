/**
 * ExcelModel.cs (c) 2017 by x01
 */
using System;

namespace x01.ExcelHelper
{
	/// <summary>
	/// Description of ExcelModel.
	/// </summary>
	public class SplitModel
	{
		public string Bianhao { get; set; }
		public string Danhao { get; set; }
		public string Pinming { get; set; }
		public string Dingjia { get; set; }
		public string Zhekou { get; set; }
		public string Shuliang { get; set; }
		
		public SplitModel()
		{
			
		}
	}
	
	
	public class CategoryModel
	{
		public string SellNr { get; set; }
		public string SellMethod { get; set; }
		public double ClientId { get; set; }
		public string ClientName { get; set; }
		public double TypeNr { get; set; }
		public double Number { get; set; }
		public double Total { get; set; }
		public string Category { get; set; }
	}
	
}
