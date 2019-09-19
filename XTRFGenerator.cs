using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using Aspose.Cells;

namespace CoolTool
{
	internal class XTRFGenerator
	{
		private string file;
		private double exchangeRate;
		private List<string[]> dt = new List<string[]>();
		private Dictionary<string, string> headerDict = new Dictionary<string, string>();
		private Dictionary<int, ReportColumn> columns = new Dictionary<int, ReportColumn>();
		private int numberOfColumns;


		public XTRFGenerator(string file, double exchangeRate)
		{
			this.file = file;
			this.exchangeRate = exchangeRate;

			headerDict.Add("ID", "t");
			headerDict.Add("Name", "t");
			headerDict.Add("Languages", "t");
			headerDict.Add("Status", "t");
			headerDict.Add("Client > Legal Name", "t");
			headerDict.Add("Client > Name", "t");
			headerDict.Add("Client PO Number", "t");
			headerDict.Add("Project Manager", "t");
			headerDict.Add("Project Manager > Name", "t");
			headerDict.Add("Sales Person", "t");
			headerDict.Add("Total Agreed", "c");
			headerDict.Add("Total Cost", "c");
			headerDict.Add("Quote > Margin", "p");
			headerDict.Add("ROI", "p");
			headerDict.Add("PM cost", "c");
			headerDict.Add("Real profit", "c");
			headerDict.Add("Real margin", "p");
			headerDict.Add("Start Date and Time", "d");
			headerDict.Add("Deadline", "d");
			headerDict.Add("Project, Margin", "p");
			headerDict.Add("Time Spent in Minutes", "n");

			ProcessFile();
		}

		private void ProcessFile()
		{
			string[] lines = File.ReadAllLines(file);
			int lineCounter = 0;
			string line = "";
			char splitChar = '\t';

			do
			{
				line = lines[lineCounter];
				lineCounter++;
			} while (line.Split(splitChar).Length < 3 && lineCounter < 20);

			if (lineCounter == 20)
			{
				Log.AddLog("Az első 20 sor nem tartalmazott adatot.");
				return;
			}
			Program.mainWindow.updateProgress(10);

			parseHeader(line);
			Program.mainWindow.updateProgress(30);

			do
			{
				line = lines[lineCounter];
				string[] rl = line.Split(splitChar);
				if (rl.Length < numberOfColumns)
				{
					int tmplineCounter = lineCounter;
					int tmpLength = rl.Length;
					while (tmpLength < numberOfColumns && tmplineCounter < lines.Length - 1)
					{
						tmplineCounter++;
						line += "\r\n" + lines[tmplineCounter];
						tmpLength = line.Split(splitChar).Length;
					}

					if (tmpLength == numberOfColumns)
					{
						rl = line.Split(splitChar);
						dt.Add(rl);
						lineCounter = tmplineCounter;
					}
					else
					{
						Log.AddLog("A sor nem olvasható be: " + line, true);
					}
				}
				else if (rl.Length == numberOfColumns)
				{
					dt.Add(rl);
				}
				
				lineCounter++;
			}
			while (lineCounter < lines.Length);
			
			Log.AddLog("Az adatok beolvasása kész. Adatsorok száma: " + dt.Count);
			Program.mainWindow.updateProgress(50);
			
			RecheckHeader();

			ExportFile();
			Program.mainWindow.updateProgress(100);
		}

		private void RecheckHeader()
		{
			foreach (int column in columns.Keys)
			{
				if (columns[column].type == columnType.Unknown)
				{
					columns[column].type = columnType.Text;

					List<string> valuesToCheck = new List<string>();
					for (int i = 0; i < dt.Count && valuesToCheck.Count < 5; i++)
					{
						string[] rl = dt[i];
						if (!String.IsNullOrEmpty(rl[column].Trim()))
							valuesToCheck.Add(rl[column]);
					}

					if (valuesToCheck.Count > 0)
					{
						columnType cTbase;
						string valueToCheck = valuesToCheck[0];
						cTbase = ParseFormat(valueToCheck);

						for (int i = 1; i < valuesToCheck.Count; i++)
						{
							if (cTbase != ParseFormat(valuesToCheck[i]))
							{
								cTbase = columnType.Text;
								break;
							}
						}

						columns[column].type = cTbase;
					}

					if (columns[column].type == columnType.Text)
					{
						Log.AddLog("Az oszlopot a CoolTool nem ismeri, ezért szövegként kezeli: " + columns[column].header);
					}
				}
			}
		}

		private columnType ParseFormat(string valueToCheck)
		{
			double valueD;
			DateTime valueDT;
			columnType cT;

			if (DateTime.TryParse(valueToCheck.Replace("CET", "").Replace("CEST", "").Trim(), out valueDT))
			{
				cT = columnType.Date;
			}
			else if (valueToCheck.Contains("%")
				&& Double.TryParse(valueToCheck.Replace("%", "").Replace(".", ",").Replace(" ", "").Trim(), out valueD))
			{
				cT = columnType.Percent;
			}
			else if ((valueToCheck.Contains("€") || valueToCheck.ToLower().Contains("Ft"))
				&& Double.TryParse(valueToCheck.Replace(".", ","), out valueD))
			{
				cT = columnType.Currency;
			}
			else if (Double.TryParse(valueToCheck.Replace(".", ",").Replace(" ", "").Trim(), out valueD))
			{
				cT = columnType.Number;
			}
			else cT = columnType.Text;

			return cT;
		}

		private void parseHeader(string line)
		{
			string[] headerStrings = line.Split('\t');   // beolvasott header-ek
			numberOfColumns = headerStrings.Length;
			int columnCounter = 0;
			foreach (string headerItem in headerStrings)    // a beolvasott headerek sorban
			{

				ReportColumn rc = new ReportColumn();   // új report oszlop létrehozása
				rc.header = headerItem;     // a címe a beolvasott címe

				if (headerDict.ContainsKey(headerItem.Trim()))    // ha értelmezett a beolvasott header
				{
					switch (headerDict[headerItem.Trim()])    // az adott oszlop típusa
					{
						case "t":
							rc.type = columnType.Text;
							break;
						case "p":
							rc.type = columnType.Percent;
							break;
						case "d":
							rc.type = columnType.Date;
							break;
						case "c":
							rc.type = columnType.Currency;
							break;
						case "n":
							rc.type = columnType.Number;
							break;
					}
				}
				else
				{
					rc.type = columnType.Unknown;  // ha nem ismeri az adott oszlopot, akkor szövegként rakja ki
				}

				columns.Add(columnCounter, rc);
				columnCounter++;

			}

			Log.AddLog("A fejléc beolvasása kész. Oszlopok száma: " + columns.Count);

		}

		private void ExportFile()
		{
			string exportFilePath = Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".xlsx");
			Workbook wb = new Workbook();
			Worksheet ws = wb.Worksheets[0];

			for (int column = 0; column < columns.Count; column++)
			{
				ws.Cells[0, column].PutValue(columns[column].header);
			}

			int row = 1;
			Cell c;
			Style style;

			foreach (string[] rl in dt)
			{

				for (int column = 0; column < rl.Length; column++)
				{
					switch (columns[column].type)
					{
						case columnType.Currency:
							double valueC;
							string valueString = rl[column];
							double exchRate = 1;
							if (valueString.Contains("€"))
							{
								exchRate = exchangeRate;
								valueString = valueString.Replace("€", "").Trim();
							}
							else
							{
								valueString = valueString.Replace("Ft", "").Trim();
							}

							if (!Double.TryParse(valueString.Replace(".", ","), out valueC))
							{
								valueC = 0;
							}

							valueC = valueC * exchRate;
							ws.Cells[row, column].PutValue(valueC);
							c = ws.Cells[row, column];
							style = c.GetStyle();
							style.Custom = "# ### ##0.00 Ft";
							c.SetStyle(style);
							break;
						case columnType.Date:
							DateTime valueDT;
							if (!DateTime.TryParse(rl[column].Replace(" CET", "").Replace(" CEST", ""), out valueDT))
							{
								valueDT = DateTime.Now;
							}

							ws.Cells[row, column].PutValue(valueDT);
							c = ws.Cells[row, column];
							style = c.GetStyle();
							style.Custom = "yyyy-mm-dd";
							c.SetStyle(style);
							break;
						case columnType.Number:
							double valueN;
							if (!Double.TryParse(rl[column].Replace(".", ",").Replace(" ", "").Trim(), out valueN))
							{
								valueN = 0;
							}

							ws.Cells[row, column].PutValue(valueN);
							c = ws.Cells[row, column];
							style = c.GetStyle();
							style.Custom = "# ### ##0.00";
							c.SetStyle(style);
							break;
						case columnType.Percent:
							double valuePc;
							if (!Double.TryParse(rl[column].Replace("%", "").Replace(".", ",").Replace(" ", "").Trim(), out valuePc))
							{
								valuePc = 0;
							}

							valuePc = valuePc / 100;

							ws.Cells[row, column].PutValue(valuePc);
							c = ws.Cells[row, column];
							style = c.GetStyle();
							style.Custom = "0%";
							c.SetStyle(style);
							break;
						case columnType.Text:
							ws.Cells[row, column].PutValue(rl[column]);
							break;

					}

				}
				row++;
				Program.mainWindow.updateProgress(50 + 40 * row / dt.Count);
			}

			ws.AutoFitColumns(0, ws.Cells.MaxDataColumn);
			wb.Save(exportFilePath);
			Log.AddLog("A CSV konvertálása befejeződött. Exportfájl: " + exportFilePath);
		}

		class ReportColumn
		{
			public string header;
			public columnType type;


		}
		public enum columnType
		{
			Percent, Date, Text, Number, Currency, Unknown
		}

	}
}