using System;
using System.Globalization;
using System.IO;
using System.Threading;

namespace VioletDocumentCreator
{
	public class VioletDocumentCreator
	{
		private readonly IDocumentCreator _documentCreator;
		private readonly IUriSchemeRegisterer _uriSchemeRegisterer;

		public VioletDocumentCreator(IDocumentCreator documentCreator, IUriSchemeRegisterer uriSchemeRegisterer)
		{
			_documentCreator = documentCreator;
			_uriSchemeRegisterer = uriSchemeRegisterer;
		}

		public void Run(string[] args)
		{
			Thread.CurrentThread.CurrentUICulture = new CultureInfo("en-us");

			if (args.Length == 0)
			{
				Console.WriteLine("Missing arguments.\n");
				Console.WriteLine("Press any key to continue...");
				return;
			}

			if (args[0].StartsWith(Consts.VioletSchemePrefix))
			{
				try
				{
					_documentCreator.CreateDocument(args);
				}
				catch (FileNotFoundException e)
				{
					Console.WriteLine();
					Console.WriteLine("##############################################################");
					Console.WriteLine("#                   WORD file not found                      #");
					Console.WriteLine("#        Please make sure the lecture name is correct        #");
					Console.WriteLine("##############################################################");
					Console.WriteLine("\n\n" + e.Message);
					Console.Read();
					return;
				}
				catch (IOException e)
				{
					Console.WriteLine();
					Console.WriteLine("##############################################################");
					if (e.Message.ToLower().Contains("doc"))
						Console.WriteLine("#                  Error editing WORD file                   #");
					else if (e.Message.ToLower().Contains("pdf"))
						Console.WriteLine("#                   Error opening PDF file                   #");
					else
						Console.WriteLine("#                     Error editin file                      #");
					Console.WriteLine("#         Please make sure file is not already open          #");
					Console.WriteLine("##############################################################");
					Console.WriteLine("\n\n" + e.Message);
					Console.Read();
					return;
				}
				catch (ArgumentException e)
				{
					Console.WriteLine();
					Console.WriteLine("##############################################################");
					Console.WriteLine("#                  Missing field in order                    #");
					Console.WriteLine("#       Please make sure the following fields are valid      #");
					Console.WriteLine("#     Organization name, Contact name, Date, phone number    #");
					Console.WriteLine("##############################################################");
					Console.WriteLine("\n\n" + e.Message);
					Console.Read();
				}
				catch (Exception e)
				{
					Console.WriteLine();
					Console.WriteLine("##############################################################");
					Console.WriteLine("#                        Error found                         #");
					Console.WriteLine("#          Please contact your local support agent           #");
					Console.WriteLine("##############################################################");
					Console.WriteLine("\n\n" + e.Message);
					Console.Read();
					return;
				}
				return;
			}

			switch (args[0])
			{
				case "install":
					_uriSchemeRegisterer.Install();
					break;

				case "uninstall":
					_uriSchemeRegisterer.Uninstall();
					break;

				default:
					Console.WriteLine("Invalid argument.");
					Console.WriteLine("Valid arguments: install, uninstall");
					break;
			}

			Console.WriteLine();
			Console.WriteLine("Press any key to continue...");
		}
	}
}