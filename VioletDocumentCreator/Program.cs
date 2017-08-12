﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VioletDocumentCreator
{
	public class Program
	{
		public static void Main(string[] args)
		{
			if (args.Length == 0)
			{
				Console.WriteLine("Missing arguments.");
				Console.WriteLine();
				Console.WriteLine("Press any key to continue...");
				Console.ReadKey();
			}

			if (args[0].StartsWith("violet:"))
			{
				DocumentCreator.CreateDocument(args);
				return;
			}

			switch (args[0])
			{
				case "install":
					UriSchemeRegisterer.Install();
					break;

				case "uninstall":
					UriSchemeRegisterer.Uninstall();
					break;				

				default:
					Console.WriteLine("Invalid argument.");
					Console.WriteLine("Valid arguments: install, uninstall");
					break;
			}

			Console.WriteLine();
			Console.WriteLine("Press any key to continue...");
			Console.ReadKey();
		}
	}
}