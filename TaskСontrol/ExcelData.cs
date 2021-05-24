using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskСontrol
{
	internal class ExcelData
	{
		public static Collection<Detail> ExcelDataLoad(string dir, int startRow, int[] column)
		{
			Collection<Detail> details = new Collection<Detail>();

			Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
			Workbook wb = excel.Workbooks.Open(dir);
			Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet;

			int currentRow = startRow;
			bool isNotEmptyField = true;
			while (isNotEmptyField)
			{
				var b = excelSheet.Cells[currentRow, column[1]].Value;
				if (b != null
					&& Convert.ToString(b) != ""
					&& Convert.ToString(b) != " ")
				{
					Detail detail = new Detail((string)(excelSheet.Cells[currentRow, column[0]] as Microsoft.Office.Interop.Excel.Range)?.Value?.Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[2]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[3]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[1]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						Convert.ToString((excelSheet.Cells[currentRow, column[4]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
						new FileInfo(dir).Name);
					details.Add(detail);
					currentRow += 1;
				}
				else 
				{ 
					isNotEmptyField = false;					
				}				
			}
			wb.Close(0);
			excel.Quit();
			return details;
		}      
	}
}
