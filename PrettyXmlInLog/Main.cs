using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
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
			PluginBase.SetCommand(0, "Format Line", FormatLine, new ShortcutKey(true, true, true, Keys.L));
			PluginBase.SetCommand(1, "Format Selection", FormatSelection, new ShortcutKey(true, true, true, Keys.S));
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
				MessageBox.Show(selection);
			}
			catch (Exception ignore)
			{ }
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

		#endregion
	}
}