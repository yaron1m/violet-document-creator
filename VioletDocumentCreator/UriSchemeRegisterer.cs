using System;
using System.IO;
using Microsoft.Win32;

// Credit - https://www.brad-smith.info/blog/archives/842
namespace VioletDocumentCreator
{
	public class UriSchemeRegisterer
	{
		private const string UriScheme = "violet";
		private const string UriKey = "URL:Violet Protocol";

		public static void Install()
		{
			// install
			var appPath = typeof(Program).Assembly.Location;

			Console.Write($"Attempting to register URI scheme for:\n{appPath}");

			try
			{
				if (!File.Exists(appPath))
				{
					throw new InvalidOperationException($"Application not found at: {appPath}");
				}

				RegisterUriScheme(appPath);
				Console.WriteLine(" Success.");
			}
			catch (Exception ex)
			{
				Console.WriteLine(" Failed!");
				Console.WriteLine("{0}: {1}", ex.GetType().Name, ex.Message);
			}
		}

		public static void Uninstall()
		{
			// uninstall
			Console.Write("Attempting to unregister URI scheme...");

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree(UriScheme);
				Console.WriteLine(" Success.");
			}
			catch (Exception ex)
			{
				Console.WriteLine(" Failed!");
				Console.WriteLine("{0}: {1}", ex.GetType().Name, ex.Message);
			}
		}

		public static void RegisterUriScheme(string appPath)
		{
			// HKEY_CLASSES_ROOT\myscheme
			using (RegistryKey hkcrClass = Registry.ClassesRoot.CreateSubKey(UriScheme))
			{
				hkcrClass.SetValue(null, UriKey);
				hkcrClass.SetValue("URL Protocol", String.Empty, RegistryValueKind.String);

				// use the application's icon as the URI scheme icon
				using (RegistryKey defaultIcon = hkcrClass.CreateSubKey("DefaultIcon"))
				{
					string iconValue = String.Format("\"{0}\",0", appPath);
					defaultIcon.SetValue(null, iconValue);
				}

				// open the application and pass the URI to the command-line
				using (RegistryKey shell = hkcrClass.CreateSubKey("shell"))
				{
					using (RegistryKey open = shell.CreateSubKey("open"))
					{
						using (RegistryKey command = open.CreateSubKey("command"))
						{
							string cmdValue = String.Format("\"{0}\" \"%1\"", appPath);
							command.SetValue(null, cmdValue);
						}
					}
				}
			}
		}
	}
}
