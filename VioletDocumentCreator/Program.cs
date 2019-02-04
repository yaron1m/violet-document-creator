namespace VioletDocumentCreator
{
	public class Program
	{
		public static void Main(string[] args)
		{
			var documentCreator = new DocumentCreator();
			var uriSchemeRegisterer = new UriSchemeRegisterer();
			var violetDocumentCreator = new VioletDocumentCreator(documentCreator, uriSchemeRegisterer);
			violetDocumentCreator.Run(args);
		}
	}
}
