using System;
using System.Collections.Generic;

namespace VioletDocumentCreator
{
	public class Order
	{
		public string Topic { get; set; }
		public string Email { get; set; }
		public string OrganizationName { get; set; }
		public string ContactFirstName { get; set; }
		public string ContactLastName { get; set; }
		public string OrderId { get; set; }
		public DateTime OrderCreationDate { get; set; }
		public string ContactPhone1 { get; set; }
		public string ContactPhone2 { get; set; }

		public string ContactName => ContactFirstName + " " + ContactLastName;

		public Order(string[] args)
		{
			var joinArguments = string.Join(" ", args);
			var rawData = joinArguments.Substring("violet:".Length);

			var data = new Dictionary<string, string>();

			foreach (var item in rawData.Split('&'))
			{
				var keyValue = item.Split('=');
				data[keyValue[0]] = keyValue[1];
			}

			Topic = GetValueOrNull(data, "topic");
			Email = GetValueOrNull(data, "email");
			OrganizationName = GetValueOrNull(data, "organizationName");
			ContactFirstName = GetValueOrNull(data, "contactFirstName");
			ContactLastName = GetValueOrNull(data, "contactLastName");
			OrderId = GetValueOrNull(data, "orderId");
			OrderCreationDate = GetValueOrNull(data, "orderCreationDate") == null ? new DateTime() : DateTime.Parse(GetValueOrNull(data, "orderCreationDate"));
			ContactPhone1 = GetValueOrNull(data, "contactPhone1");
			ContactPhone2 = GetValueOrNull(data, "contactPhone2");
		}

		private static string GetValueOrNull(Dictionary<string, string> data, string key) => data.ContainsKey(key) ? data[key] : null;
	}
}
