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
	public enum ColumnIndex : int
	{
		PartNumber = 0,
		Manufacturer,
		Assembly,
		Symbol,
		Location,
		Description1,
		Description2,
		Description3
	}

	public partial class Controller : IExtensionApplication
	{
		public static CultureInfo ci = new CultureInfo ("en");

		public void Initialize ()
		{
		}

		public void Terminate ()
		{
		}

		protected static TypedValue[] RunApplicationCommand (params TypedValue[] Values)
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
		public void ImportList ()
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

			// document
			var ActiveDocument = Application.DocumentManager.MdiActiveDocument;
			var Editor = ActiveDocument.Editor;

			// read input
			IWorkbook wb = null;
			try
			{
				wb = SpreadsheetGear.Factory.GetWorkbook (ofd.Filename);
			}
			catch
			{
				Editor.WriteMessage ("Failed to open the spreadsheet");
				return;
			}

			Entry[] Entries;
			try
			{
				IWorksheet ws = wb.Worksheets[0];
				IRange UsedRange = ws.UsedRange;
				int RowsCount = UsedRange.RowCount;
				Entries = Enumerable.Range (0, RowsCount)
					.Skip (1)       // skip header row
					.Select (n =>
						{
							Entry Entry = new Entry
							{
								PartNumber = ws.Cells[n, (int)ColumnIndex.PartNumber].Text.Trim (),
								Manufacturer = ws.Cells[n, (int)ColumnIndex.Manufacturer].Text.Trim (),
								Assembly = ws.Cells[n, (int)ColumnIndex.Assembly].Text.Trim (),
								Symbol = ws.Cells[n, (int)ColumnIndex.Symbol].Text.Trim (),
								Description1 = ws.Cells[n, (int)ColumnIndex.Description1].Text.Trim (),
								Description2 = ws.Cells[n, (int)ColumnIndex.Description2].Text.Trim (),
								Description3 = ws.Cells[n, (int)ColumnIndex.Description3].Text.Trim ()
							};

							string LocStr = ws.Cells[n, (int)ColumnIndex.Location].Text;
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
			}
			finally
			{
				wb.Close ();
			}

			// edit the drawing
			using (Database db = ActiveDocument.Database)
			{
				// find Installation Code
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

				// render items
				foreach (var Entry in Entries)
				{
					// due attributes
					Dictionary<string, string> Attributes = new Dictionary<string, string>
						{
							["MFG"] = Entry.Manufacturer,
							["CAT"] = Entry.PartNumber,
							["ASSYCODE"] = Entry.Assembly,
							["DESC1"] = Entry.Description1,
							["DESC2"] = Entry.Description2,
							["DESC3"] = Entry.Description3,
							["INST"] = InstallationCode
						};

					// find 'ratings'
					TypedValue[] RatingsRaw = RunApplicationCommand (
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

					// insert the symbol
					string BlockHandle = RunApplicationCommand (
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

					// populate attributes
					using (Transaction tran = db.TransactionManager.StartTransaction ())
					{
						var Block = tran.GetObject (id, OpenMode.ForWrite) as BlockReference;

						// update existing
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

						// add remaining
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

					// attach pins
					RunApplicationCommand (
						new TypedValue ((int)LispDataType.Text, "c:wd_pinlist_attach"),
						new TypedValue ((int)LispDataType.ObjectId, id),
						new TypedValue ((int)LispDataType.Int32, 1)
					);
				}
			}
		}
	}
}
