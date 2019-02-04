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

		public Offer(string offerData)
		{
			var data = offerData.Split('&')
				.Select(x => x.Split('='))
				.ToDictionary(x => x[0], x => x[1]);

			Id = GetValueOrNull(data, "id");
			Email = GetValueOrNull(data, "email");
			OrganizationName = GetValueOrNull(data, "organizationName");
			ContactFirstName = GetValueOrNull(data, "contactFirstName");
			ContactLastName = GetValueOrNull(data, "contactLastName");
			OrderCreationDate = GetValueOrNull(data, "orderCreationDate") == null
				? new DateTime()
				: DateTime.Parse(data["orderCreationDate"]);
			ContactPhone1 = GetValueOrNull(data, "contactPhone1");
			ContactPhone2 = GetValueOrNull(data, "contactPhone2");

			Topic = data.ContainsKey("topic") ? data["topic"].Split('#') : null;
		}

		private static string GetValueOrNull(Dictionary<string, string> data, string key) => data.ContainsKey(key)
			? data[key]
			: null;

		public string PhoneNumbers => string.Join(", ",
			new List<string>() {ContactPhone1, ContactPhone2}.Where(x => !string.IsNullOrWhiteSpace(x)));

		public string ContactName => ContactFirstName + " " + ContactLastName;

		private readonly string _directory =
			Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

		public string GetTamplatePath(int topicIndex) => _directory + @"\word\" + Topic[topicIndex] + Docx;
		public string GetSavingDirectory(int topicIndex) => _directory + @"\הצעות מחיר\" + Topic[topicIndex];

		private string GetSavingName(int topicIndex) => Topic[topicIndex] + " - " + OrganizationName + " - " + ContactName +
		                                                " " + Id;

		public string GetPdfFileName(int topicIndex) => GetSavingName(topicIndex) + Pdf;

		public string GetDocSavingPath(int topicIndex)
		{
			return _directory
			       + @"\הצעות מחיר\"
			       + RemoveInvalid(Topic[topicIndex]) + @"\"
			       + RemoveInvalid(GetSavingName(topicIndex))
			       + Docx;
		}

		public string GetPdfSavingPath(int topicIndex)
		{
			return _directory
			       + @"\הצעות מחיר\"
			       + RemoveInvalid(Topic[topicIndex]) + @"\" +
			       RemoveInvalid(GetPdfFileName(topicIndex));
		}

		private string RemoveInvalid(string str)
		{
			string invalid = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());

			foreach (char c in invalid)
			{
				str = str.Replace(c.ToString(), "");
			}

			return str;
		}
	}
}