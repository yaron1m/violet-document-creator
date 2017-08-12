using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;

namespace VioletDocumentCreator
{
	public class DocumentCreator
	{
		public static void CreateDocument(string[] args)
		{
			var joinArguments = string.Join(" ", args);
			var rawData = joinArguments.Substring("violet:".Length);

			var data = new Dictionary<string, string>();

			/*
				var subject = "ISO-14001";
				var email = "email";
				var orgName = "orgName";
				var contactName = "conatctName";
				var id = "id";
				var recieveDate = "2008-05-01T07:34:42-5:00";
				var cell1 = "cell1";
				var cell2 = "cell2";
				var phone = "phone";
			*/

			foreach (var item in rawData.Split('&'))
			{
				var keyValue = item.Split('=');
				data[keyValue[0]] = keyValue[1];
			}
		}

		private static void CreateWordAndPDF(string subject, string contactName, string orgName, string id, string recieveDate, string cell1, string cell2, string phone)
		{
			using (var docX = DocX.Load(Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\word\" + subject + ".docx"))
			{
				docX.ReplaceText("<שם איש קשר>", contactName);
				docX.ReplaceText("<שם ארגון>", orgName);
				docX.ReplaceText("<סימוכין>", id);
				DateTime dt = DateTime.Parse(recieveDate);
				docX.ReplaceText("<תאריך>", dt.ToString("dddd dd בMMMM yyyy"));


				var phones = new List<string>();
				phones.Add(cell1);
				phones.Add(cell2);
				phones.Add(phone);
				var phoneNumbersList = phones.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
				var phoneNumbers = String.Join(", ", phoneNumbersList);

				docX.ReplaceText("<נייד>", phoneNumbers);


				var saveDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\הצעות מחיר\" +
				              subject;
				if (!System.IO.Directory.Exists(saveDir))
				{
					System.IO.Directory.CreateDirectory(saveDir);
				}

				string savePath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) + @"\הצעות מחיר\" + subject + @"\" + orgName + " " + contactName + " " + id + ".docx";

				docX.SaveAs(savePath);

				string PdfPath = savePath.Substring(0, savePath.Length - 4) + "pdf";// Path.GetPathRoot(savePath) + Path.GetFileNameWithoutExtension(savePath) + ".pdf";
				SaveWordToPDF.Convert(savePath, PdfPath);
			}
		}
	}
}
