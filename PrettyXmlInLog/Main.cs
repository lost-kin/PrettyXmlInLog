using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using NppPluginNET;

namespace PrettyXmlInLog
{
	class Main
	{
		#region Fields

		internal const string PluginName = "Pretty Xml In Log";

		#endregion

		#region StartUp/CleanUp

		internal static void CommandMenuInit()
		{
			PluginBase.SetCommand(0, "Format Selection", FormatSelection, new ShortcutKey(true, true, true, Keys.S));
			PluginBase.SetCommand(1, "Format Line", FormatLine, new ShortcutKey(true, true, true, Keys.L));
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
				var hCurrentEditView = PluginBase.GetCurrentScintilla();

				var currentPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETCURRENTPOS, 0, 0);
				var lineNum = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_LINEFROMPOSITION, currentPos, 0);

				var startPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_POSITIONFROMLINE, lineNum, 0);
				var endPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETLINEENDPOSITION, lineNum, 0);

				FormatAsXmlTextBetween(startPos, endPos);
			}
			catch (Exception ignore)
			{ }
		}

		internal static void FormatSelection()
		{
			try
			{
				var hCurrentEditView = PluginBase.GetCurrentScintilla();

				var startPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETSELECTIONNSTART, 0, 0);
				var endPos = (int)Win32.SendMessage(hCurrentEditView, SciMsg.SCI_GETSELECTIONNEND, 0, 0);

				FormatAsXmlTextBetween(startPos, endPos);
			}
			catch (Exception ignore)
			{ }
		}

		private static void FormatAsXmlTextBetween(int startPos, int endPos)
		{
			if (startPos >= endPos) { return; }

			var lineText = GetTextRange(startPos, endPos);
			lineText = FormatAsXml(lineText);

			var hCurrentEditView = PluginBase.GetCurrentScintilla();
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_SETCURRENTPOS, startPos, 0);
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_DELETERANGE, startPos, endPos - startPos);
			Win32.SendMessage(hCurrentEditView, SciMsg.SCI_INSERTTEXT, startPos, lineText);
		}

		private static string FormatAsXml(string text)
		{
			var ms = new MemoryStream();
			var xtw = new XmlTextWriter(ms, Encoding.Unicode);
			var doc = new XmlDocument();

			doc.LoadXml(text);
			xtw.Formatting = Formatting.Indented;
			doc.WriteContentTo(xtw);
			xtw.Flush();
			ms.Seek(0, SeekOrigin.Begin);
			var sr = new StreamReader(ms);

			return sr.ReadToEnd();
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

		#endregion
	}
}