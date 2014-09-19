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
			//PluginBase.SetCommand(1, "Format Line", FormatLine, new ShortcutKey(true, true, true, Keys.L));
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
				var line = GetCurrentLine();
				MessageBox.Show(line);
			}
			catch (Exception ignore)
			{ }
		}

		internal static void FormatSelection()
		{
			try
			{
				var selection = GetSelection();

				selection = FormatAsXml(selection);

				ReplaceSelection(selection);
			}
			catch (Exception ignore)
			{ }
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
				var lineStringBuilder = new StringBuilder(length);
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