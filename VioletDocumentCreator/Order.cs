using System;
using System.Collections.Generic;
using System.IO;

namespace VioletDocumentCreator
{
	public class Order
	{
		public string topic { get; set; }
		public string email { get; set; }
		public string organizationName { get; set; }
		public string contactFirstName { get; set; }
		public string contactLastName { get; set; }
		public string id { get; set; }
		public DateTime orderCreationDate { get; set; }
		public string contactPhone1 { get; set; }
		public string contactPhone2 { get; set; }

		public string ContactName => contactFirstName + " " + contactLastName;

		private readonly string _directory = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

//		public Order(string[] args)
//		{
//			var joinArguments = string.Join(" ", args);
//			var rawData = joinArguments.Substring("violet:".Length);
//
//			var data = new Dictionary<string, string>();
//
//			foreach (var item in rawData.Split('&'))
//			{
//				var keyValue = item.Split('=');
//				data[keyValue[0]] = keyValue[1];
//			}
//
//			topic = GetValueOrNull(data, "topic");
//			email = GetValueOrNull(data, "email");
//			organizationName = GetValueOrNull(data, "organizationName");
//			contactFirstName = GetValueOrNull(data, "contactFirstName");
//			contactLastName = GetValueOrNull(data, "contactLastName");
//			id = GetValueOrNull(data, "orderId");
//			orderCreationDate = GetValueOrNull(data, "orderCreationDate") == null ? new DateTime() : DateTime.Parse(GetValueOrNull(data, "orderCreationDate"));
//			contactPhone1 = GetValueOrNull(data, "contactPhone1");
//			contactPhone2 = GetValueOrNull(data, "contactPhone2");
//		}

		private static string GetValueOrNull(Dictionary<string, string> data, string key) => data.ContainsKey(key) ? data[key] : null;

		public string GetTamplatePath() => _directory + @"\word\" + topic + ".docx";
		public string GetSavingDirectory() => _directory + @"\הצעות מחיר\" + topic;

		private string GetSavingName() => organizationName + " " + ContactName + " " + id;

		public string GetPdfFileName() => GetSavingName() + ".pdf";
		public string GetDocSavingPath() => _directory + @"\הצעות מחיר\" + topic + @"\" + GetSavingName() + ".docx";
		public string GetPdfSavingPath() => _directory + @"\הצעות מחיר\" + topic + @"\" + GetPdfFileName();


	}
}
