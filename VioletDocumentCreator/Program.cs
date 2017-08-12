using System;
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
			if (args[1].Equals("install"))
			{
				UriSchemeRegisterer.Install(args);
			}
			else
			{
				
			}

		}
	}
}
