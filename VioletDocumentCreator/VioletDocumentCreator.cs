using System;
using System.Globalization;
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
				_documentCreator.CreateDocument(args);
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