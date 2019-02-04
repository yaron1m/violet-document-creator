using System;
using System.IO;
using Microsoft.Win32;

// Credit - https://www.brad-smith.info/blog/archives/842
namespace VioletDocumentCreator
{
	public interface IUriSchemeRegisterer
	{
		void Install();
		void Uninstall();
	}

	public class UriSchemeRegisterer : IUriSchemeRegisterer
	{
		private const string UriScheme = "violet";
		private const string UriKey = "URL:Violet Protocol";

		public void Install()
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

		public void Uninstall()
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
			using (var hkcrClass = Registry.ClassesRoot.CreateSubKey(UriScheme))
			{
				hkcrClass.SetValue(null, UriKey);
				hkcrClass.SetValue("URL Protocol", String.Empty, RegistryValueKind.String);

				// use the application's icon as the URI scheme icon
				using (var defaultIcon = hkcrClass.CreateSubKey("DefaultIcon"))
				{
					string iconValue = $"\"{appPath}\",0";
					defaultIcon.SetValue(null, iconValue);
				}

				// open the application and pass the URI to the command-line
				using (var shell = hkcrClass.CreateSubKey("shell"))
				{
					using (var open = shell.CreateSubKey("open"))
					{
						using (var command = open.CreateSubKey("command"))
						{
							var cmdValue = $"\"{appPath}\" \"%1\"";
							command.SetValue(null, cmdValue);
						}
					}
				}
			}
		}
	}
}