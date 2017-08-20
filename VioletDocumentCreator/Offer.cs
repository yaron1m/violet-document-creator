using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace VioletDocumentCreator
{
	public class Offer
	{
		private const string Docx = ".docx";
		private const string Pdf = ".pdf";

		public string[] Topic { get; set; }
		public string Email { get; set; }
		public string OrganizationName { get; set; }
		public string ContactFirstName { get; set; }
		public string ContactLastName { get; set; }
		public string Id { get; set; }
		public DateTime OrderCreationDate { get; set; }
		public string ContactPhone1 { get; set; }
		public string ContactPhone2 { get; set; }

		public Offer(string[] args)
		{
			var joinArguments = string.Join(" ", args);
			joinArguments = HttpUtility.UrlDecode(joinArguments);
			var rawData = joinArguments.Substring("violet:".Length);

			var data = new Dictionary<string, string>();

			foreach (var item in rawData.Split('&'))
			{
				var keyValue = item.Split('=');
				data[keyValue[0]] = keyValue[1];
			}

			Id = GetValueOrNull(data, "id");
			Email = GetValueOrNull(data, "email");
			OrganizationName = GetValueOrNull(data, "organizationName");
			ContactFirstName = GetValueOrNull(data, "contactFirstName");
			ContactLastName = GetValueOrNull(data, "contactLastName");
			OrderCreationDate = GetValueOrNull(data, "orderCreationDate") == null ? new DateTime() : DateTime.Parse(data["orderCreationDate"]);
			ContactPhone1 = GetValueOrNull(data, "contactPhone1");
			ContactPhone2 = GetValueOrNull(data, "contactPhone2");

			Topic = data.ContainsKey("topic") ? data["topic"].Split('#') : null;


		}
		private static string GetValueOrNull(Dictionary<string, string> data, string key) => data.ContainsKey(key) ? data[key] : null;


		public string PhoneNumbers => string.Join(", ", new List<string>() { ContactPhone1, ContactPhone2 }.Where(x => !string.IsNullOrWhiteSpace(x)));

		public string ContactName => ContactFirstName + " " + ContactLastName;

		private readonly string _directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

		public string GetTamplatePath(int topicIndex) => _directory + @"\word\" + Topic[topicIndex] + Docx;
		public string GetSavingDirectory(int topicIndex) => _directory + @"\הצעות מחיר\" + Topic[topicIndex];

		private string GetSavingName(int topicIndex) => Topic[topicIndex] + " - " + OrganizationName + " - " + ContactName + " " + Id;

		public string GetPdfFileName(int topicIndex) => GetSavingName(topicIndex) + Pdf;
		public string GetDocSavingPath(int topicIndex) => _directory + @"\הצעות מחיר\" + Topic[topicIndex] + @"\" + GetSavingName(topicIndex) + Docx;
		public string GetPdfSavingPath(int topicIndex) => _directory + @"\הצעות מחיר\" + Topic[topicIndex] + @"\" + GetPdfFileName(topicIndex);
	}
}
