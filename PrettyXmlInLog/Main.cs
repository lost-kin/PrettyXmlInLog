using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using NppPluginNET;

namespace PrettyXmlInLog
{
	class Main
	{
		#region Fields

		internal const string PluginName = "Pretty Xml In Log";
		private static readonly Regex XmlRegex = new Regex(@"(?:<\?xml.*\?>)?\s*<(?<tag>[\w:]+)[^<]*>.*</\k<tag>>");

		#endregion

		#region StartUp/CleanUp

		internal static void CommandMenuInit()
		{
			PluginBase.SetCommand(0, "Format Selection as XML", FormatSelection, new ShortcutKey(true, true, true, Keys.S));
			PluginBase.SetCommand(1, "Format XML in Line", FormatLine, new ShortcutKey(true, true, true, Keys.L));
			PluginBase.SetCommand(2, "Format XML in All Lines", FormatAllLines, new ShortcutKey(true, true, true, Keys.A));
		}

		internal static void SetToolBarIcon()
		{
		}

		internal static void PluginCleanUp()
		{
		}

		#endregion

		#region Menu functions

		internal static void FormatLine()
		{
			try
			{
				Cursor.Current = Cursors.WaitCursor;

				var hCurrentEditView = PluginBase.GetCurrentScintilla();

				var currentPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETCURRENTPOS, 0, 0);
				var lineNum = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_LINEFROMPOSITION, currentPos, 0);

				FormatXmlInLine(lineNum);
			}
			catch (Exception ignore)
			{ }
			finally
			{
				Cursor.Current = Cursors.Default;
			}
		}

		internal static void FormatSelection()
		{
			try
			{
				Cursor.Current = Cursors.WaitCursor;

				var hCurrentEditView = PluginBase.GetCurrentScintilla();

				var startPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETSELECTIONNSTART, 0, 0);
				var endPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETSELECTIONNEND, 0, 0);

				if (startPos >= endPos) { return; }

				var selectionText = GetTextRange(startPos, endPos);
				selectionText = FormatAsXml(selectionText);

				ReplaceTextBetween(startPos, endPos, selectionText);
			}
			catch (Exception ignore)
			{ }
			finally
			{
				Cursor.Current = Cursors.Default;
			}
		}

		internal static void FormatAllLines()
		{
			try
			{
				Cursor.Current = Cursors.WaitCursor;

				int currentLine = 0;

				var lineCount = GetLineCount();

				while (currentLine < lineCount)
				{
					FormatXmlInLine(currentLine);
					currentLine++;

					var newLineCount = GetLineCount();
					var numLinesAdded = newLineCount - lineCount;
					currentLine += numLinesAdded;
					lineCount = newLineCount;
				}
			}
			catch (Exception ignore)
			{
			}
			finally
			{
				Cursor.Current = Cursors.Default;
			}
		}

		private static void FormatXmlInLine(int lineNum)
		{
			var hCurrentEditView = PluginBase.GetCurrentScintilla();
			var startPos = (int) Win32.SendMessage(hCurrentEditView, SciMsg.SCI_POSITIONFROMLINE, lineNum, 0);
			var endPos = (int) Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETLINEENDPOSITION, lineNum, 0);

			if (startPos >= endPos) { return; }

			var lineText = GetTextRange(startPos, endPos);

			var match = XmlRegex.Match(lineText);

			if (!match.Success) { return; }

			var formattedXml = FormatAsXml(match.Value);

			if (match.Index != 0)
			{
				formattedXml = Environment.NewLine + formattedXml;
			}

			if ((endPos - startPos) != match.Length)
			{
				formattedXml = formattedXml + Environment.NewLine;
			}

			var xmlStartPos = startPos + match.Index;
			var xmlEndPos = xmlStartPos + match.Length;

			ReplaceTextBetween(xmlStartPos, xmlEndPos, formattedXml);
		}

		private static void ReplaceTextBetween(int startPos, int endPos, string newText)
		{
			var hCurrentEditView = PluginBase.GetCurrentScintilla();
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_SETCURRENTPOS, startPos, 0);
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_DELETERANGE, startPos, endPos - startPos);
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_INSERTTEXT, startPos, newText);
		}

		private static string FormatAsXml(string text)
		{
			using (var ms = new MemoryStream())
			using (var xtw = new XmlTextWriter(ms, Encoding.Unicode))
			{
				var doc = new XmlDocument();
				doc.LoadXml(text);
				xtw.Formatting = Formatting.Indented;

				doc.WriteContentTo(xtw);

				xtw.Flush();
				ms.Seek(0, SeekOrigin.Begin);
				using (var sr = new StreamReader(ms))
				{
					return sr.ReadToEnd();
				}
			}
		}

		private static string GetTextRange(int startPos, int endPos)
		{
			using (var sciTextRange = new Sci_TextRange(startPos, endPos, endPos - startPos + 1))
			{
				var hCurrentEditView = PluginBase.GetCurrentScintilla();
				Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETTEXTRANGE, 0, sciTextRange.NativePointer);
				return sciTextRange.lpstrText;
			}
		}

		private static String GetCurrentLine()
		{
			return GetText(SciMsg.SCI_GETCURLINE);
		}

		private static string GetSelection()
		{
			return GetText(SciMsg.SCI_GETSELTEXT);
		}

		private static string GetText(SciMsg sciMsg)
		{
			var hCurrentEditView = PluginBase.GetCurrentScintilla();
			var length = (int)Win32.SendMessage(hCurrentEditView, sciMsg, 0, 0);

			if (length > 0)
			{
				var lineStringBuilder = new StringBuilder(length + 1);
				Win32.SendMessage(hCurrentEditView, sciMsg, length, lineStringBuilder);

				return lineStringBuilder.ToString();
			}
			else
			{
				return String.Empty;
			}
		}

		private static void ReplaceSelection(string newText)
		{
			var hCurrentEditView = PluginBase.GetCurrentScintilla();
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_REPLACESEL, 0, newText);
		}

		private static int GetLineCount()
		{
			var hCurrentEditView = PluginBase.GetCurrentScintilla();
			return (int) Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETLINECOUNT, 0, 0);
		}

		#endregion
	}
}