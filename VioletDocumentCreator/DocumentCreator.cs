using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Novacode;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;


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
	public class SendOutlookMail
	{

		public static void CreateMailItem(string filePath, string email)
		{
			//יוצר מייל חדש עם הקובץ של הזמנת ההרצאה 
			//מען המייל הוא הכתובת מהטופס
			//ישלח מחנן סי-פויינט

			var outlookApp = new Outlook.Application();
			var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
			foreach (Outlook.Account account in outlookApp.Session.Accounts)
			{
				// When the e-mail address matches, send the mail.
				if (account.SmtpAddress == "hanan@c-point.co.il")
				{
					mailItem.SendUsingAccount = account;
					mailItem.Subject = "הצעת מחיר מחנן מלין";
					mailItem.To = email;
					mailItem.Importance = Outlook.OlImportance.olImportanceLow;
					mailItem.Attachments.Add(filePath, Outlook.OlAttachmentType.olByValue, 1, Path.GetFileName(filePath));
					mailItem.Display(false);
				}
			}

		}
	}
	public class SaveWordToPDF
	{
		//שומר וורד כפידיאף

		public static void Convert(string input, string output)
		{
			// Create an instance of Word.exe
			Word._Application oWord = new Word.Application();

			// Make this instance of word invisible (Can still see it in the taskmgr).
			oWord.Visible = false;

			// Interop requires objects.
			object oMissing = System.Reflection.Missing.Value;
			object isVisible = true;
			object readOnly = false;
			object oInput = input;
			object oOutput = output;
			object oFormat = Word.WdSaveFormat.wdFormatPDF;

			// Load a document into our instance of word.exe
			Word._Document oDoc = oWord.Documents.Open(ref oInput, ref oMissing, ref readOnly, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref isVisible, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

			// Make this document the active document.
			oDoc.Activate();

			// Save this document in Word 2003 format.
			oDoc.SaveAs(ref oOutput, ref oFormat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

			// Always close Word.exe.
			object saveChanges = false;
			oWord.Quit(ref saveChanges, ref oMissing, ref oMissing);
		}


	}
}
