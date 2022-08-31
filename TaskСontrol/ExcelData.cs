using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TaskСontrol
{
	internal class ExcelData
	{
		public static Collection<Detail> ExcelDataLoad(string dir, int startRow, int[] column)
		{

			Collection<Detail> details = new Collection<Detail>();

			Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
			try
			{
				Workbook wb = excel.Workbooks.Open(dir);
				Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excel.ActiveSheet;

				int currentRow = startRow;
				bool isNotEmptyField = true;
				while (isNotEmptyField)
				{
					var b = excelSheet.Cells[currentRow, column[1]].Value;
					var c = excelSheet.Cells[currentRow, column[4]].Value;
					string tpr = "";
					int rt = 1;
					
					if (b != null
						&& Convert.ToString(b) != ""
						&& Convert.ToString(b) != " ")
					{
                        if (Convert.ToString(c) != "0")
                        {
							for (int i = 5; i <= 13; i++)
							{
								if (excelSheet.Cells[currentRow, column[i]].Value != "" && excelSheet.Cells[currentRow, column[i]].Value != null)
								{
									string temp = Convert.ToString((excelSheet.Cells[currentRow, column[i]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim();
									if (temp != "-")
									{
										if (tpr != "")
										{
											tpr += "->";
										}
										tpr += temp;
										switch (temp)
										{
											case string a when a.ToLower().Contains("гиб"):
												rt += 1;
												break;
											case string a when a.ToLower().Contains("покр"):
												rt += 1;
												break;
											case string a when a.ToLower().Contains("гор"):
												rt += 3;
												break;
											case string a when a.ToLower().Contains("эл"):
												rt += 2;
												break;
											case string a when a.ToLower().Contains("сверл"):
												rt += 3;
												break;
											case string a when a.ToLower().Contains("зенк"):
												rt += 3;
												break;
											case string a when a.ToLower().Contains("шлиф"):
												rt += 3;
												break;
											case string a when a.ToLower().Contains("тф"):
												rt += 3;
												break;
											case string a when a.ToLower().Contains("дроб"):
												rt += 1;
												break;
										}
									}
								}
							}
							if (!tpr.ToLower().Contains("эл") && !tpr.ToLower().Contains("гор") && !tpr.ToLower().Contains("покр") && !Convert.ToString((excelSheet.Cells[currentRow, column[2]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim().ToLower().Contains("оц"))
							{
								if (tpr != "")
								{
									tpr += "->";
								}
								tpr += "Сварка";
								rt += 4;
							}
							Detail detail = new Detail((string)(excelSheet.Cells[currentRow, column[0]] as Microsoft.Office.Interop.Excel.Range)?.Value?.Trim(),
								Convert.ToString((excelSheet.Cells[currentRow, column[2]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
								Convert.ToString((excelSheet.Cells[currentRow, column[3]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
								Convert.ToString((excelSheet.Cells[currentRow, column[1]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
								Convert.ToString((excelSheet.Cells[currentRow, column[4]] as Microsoft.Office.Interop.Excel.Range)?.Value).Trim(),
								new FileInfo(dir).Name,
								tpr, rt);
							details.Add(detail);
							
						}
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
			catch (Exception ex)
			{				
				excel.Quit();
				MessageBox.Show(Convert.ToString(ex));
				return details;
			}
		}
	}
}
