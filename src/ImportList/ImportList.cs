using System;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using System.Text.RegularExpressions;

using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.ApplicationServices;

using SpreadsheetGear;

namespace SweeneyControls
{
	public class Entry
	{
		public string PartNumber;
		public string Manufacturer;
		public string Assembly;
		public string Symbol;
		public double? X;
		public double? Y;

		public string Description1;
		public string Description2;
		public string Description3;
	}

	public partial class ImportList : IExtensionApplication
	{
		public static CultureInfo ci = new CultureInfo ("en");

		public void Initialize ()
		{
		}

		public void Terminate ()
		{
		}

		protected static TypedValue[] RunCommand (params TypedValue[] Values)
		{
			TypedValue[] Result;

			using (ResultBuffer args = new ResultBuffer (Values))
			using (ResultBuffer InvokeResult = Application.Invoke (args))
			{
				Result = InvokeResult.AsArray ();
			}

			return Result;
		}

		[CommandMethod("impsc")]
		public void ImportSchema ()
		{
			OpenFileDialog ofd = new OpenFileDialog ("Select Excel spreadsheet to import",
					  null,
					  "xls; xlsx",
					  "ExcelFileToLink",
					  OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles
					);

			var dr = ofd.ShowDialog ();

			if (dr != System.Windows.Forms.DialogResult.OK)
			{
				return;
			}

			//
			IWorkbook wb = SpreadsheetGear.Factory.GetWorkbook (ofd.Filename);
			IWorksheet ws = wb.Worksheets[0];
			IRange UsedRange = ws.UsedRange;
			int RowsCount = UsedRange.RowCount;
			Entry[] Entries = Enumerable.Range (0, RowsCount)
				.Skip (1)
				.Select (n =>
				{
					Entry Entry = new Entry
					{
						PartNumber = ws.Cells[n, 0].Text.Trim (),
						Manufacturer = ws.Cells[n, 1].Text.Trim (),
						Assembly = ws.Cells[n, 2].Text.Trim (),
						Symbol = ws.Cells[n, 3].Text.Trim (),
						Description1 = ws.Cells[n, 5].Text.Trim (),
						Description2 = ws.Cells[n, 6].Text.Trim (),
						Description3 = ws.Cells[n, 7].Text.Trim ()
					};

					string LocStr = ws.Cells[n, 4].Text;
					Match mLoc = Regex.Match (LocStr ?? "", @"([\d\.]+)\s*,\s*([\d\.]+)");
					if (mLoc.Success
						&& double.TryParse (mLoc.Groups[1].Value, NumberStyles.AllowDecimalPoint, ci, out double X)
						&& double.TryParse (mLoc.Groups[2].Value, NumberStyles.AllowDecimalPoint, ci, out double Y)
						)
					{
						Entry.X = X;
						Entry.Y = Y;
					}

					return Entry;
				}
				)
				.Where (e => !string.IsNullOrWhiteSpace (e.Symbol) && e.X.HasValue && e.Y.HasValue)
				.ToArray ()
				;
			wb.Close ();

			//
			var doc = Application.DocumentManager.MdiActiveDocument;
			var Editor = Application.DocumentManager.MdiActiveDocument.Editor;
			using (Database db = doc.Database)
			{
				string InstallationCode = "";
				var PropEnum = db.SummaryInfo.CustomProperties;
				while (PropEnum.MoveNext ())
				{
					if (PropEnum.Key.ToString () == "Installation Code")
					{
						InstallationCode = PropEnum.Value.ToString ();
						break;
					}
				}

				foreach (var Entry in Entries)
				{
					Dictionary<string, string> Attributes = new Dictionary<string, string>
						{
							//["TAG1"] = Entry.Symbol,
							["MFG"] = Entry.Manufacturer,
							["CAT"] = Entry.PartNumber,
							["ASSYCODE"] = Entry.Assembly,
							["DESC1"] = Entry.Description1,
							["DESC2"] = Entry.Description2,
							["DESC3"] = Entry.Description3,
							["INST"] = InstallationCode
						};

					TypedValue[] RatingsRaw = RunCommand (
							new TypedValue ((int)LispDataType.Text, "c:ace_get_textvals"),
							new TypedValue ((int)LispDataType.Nil),
							new TypedValue ((int)LispDataType.Text, Entry.Symbol),
							new TypedValue ((int)LispDataType.Text, Entry.Manufacturer),
							new TypedValue ((int)LispDataType.Text, Entry.PartNumber),
							new TypedValue ((int)LispDataType.Text, Entry.Assembly)
						);
					for (int i = 0; i + 3 < RatingsRaw.Length; i += 4)
					{
						if (RatingsRaw[i].TypeCode != (int)LispDataType.ListBegin
							|| RatingsRaw[i+1].TypeCode != (int)LispDataType.Text
							|| RatingsRaw[i+2].TypeCode != (int)LispDataType.Text
							|| RatingsRaw[i+3].TypeCode != (int)LispDataType.ListEnd
							)
						{
							break;
						}

						string Key = ((RatingsRaw[i + 1].Value as string) ?? "").ToUpper ().Trim ();
						string Value = RatingsRaw[i + 2].Value as string;

						if (!Regex.IsMatch (Key, @"^RATING\d*"))
						{
							continue;
						}

						Attributes[Key] = Value;
					}

					//
					string BlockHandle = RunCommand (
							new TypedValue ((int)LispDataType.Text, "c:wd_insym2"),
							new TypedValue ((int)LispDataType.Text, Entry.Symbol),
							new TypedValue ((int)LispDataType.ListBegin),
							new TypedValue ((int)LispDataType.Double, Entry.X),
							new TypedValue ((int)LispDataType.Double, Entry.Y),
							new TypedValue ((int)LispDataType.ListEnd),
							new TypedValue ((int)LispDataType.Nil),
							new TypedValue ((int)LispDataType.Int32, 2)
						)[0].Value as string;

					//
					long lh = Convert.ToInt64 (BlockHandle, 16);
					Handle h = new Handle (lh);
					ObjectId id = db.GetObjectId (false, h, 0);

					//
					using (Transaction tran = db.TransactionManager.StartTransaction ())
					{
						var Block = tran.GetObject (id, OpenMode.ForWrite) as BlockReference;

						foreach (var aid in Block.AttributeCollection.OfType<ObjectId> ())
						{
							using (var objAttr = tran.GetObject (aid, OpenMode.ForWrite))
							{
								if (objAttr is AttributeReference aref && Attributes.ContainsKey (aref.Tag))
								{
									aref.TextString = Attributes[aref.Tag];
									Attributes.Remove (aref.Tag);
								}
							}
						}

						foreach (var kvp in Attributes.ToArray ())
						{
							Block.AttributeCollection.AppendAttribute (new AttributeReference
								{
									Tag = kvp.Key,
									TextString = kvp.Value
								});
							Attributes.Remove (kvp.Key);
						}

						tran.Commit ();
						Block.Dispose ();
					}

					//
					RunCommand (
						new TypedValue ((int)LispDataType.Text, "c:wd_pinlist_attach"),
						new TypedValue ((int)LispDataType.ObjectId, id),
						new TypedValue ((int)LispDataType.Int32, 1)
					);
				}
			}
		}
	}
}
