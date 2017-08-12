using System.Collections.Generic;
using System.IO;
using System.Linq;
using Novacode;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace VioletDocumentCreator
{
	public class DocumentCreator
	{
		private const string EmailAddress = "yaron1m@gmail.com";

		public static void CreateDocument(string[] args)
		{
			var order = new Order(args);

			CreateWordOrder(order);

			ConvertDocxToPdf(order.GetDocSavingPath(), order.GetPdfSavingPath());

			CreateMailItem(order);
		}

		private static void CreateWordOrder(Order order)
		{
			using (var docX = DocX.Load(order.GetTamplatePath()))
			{
				docX.ReplaceText("<שם איש קשר>", order.ContactName);
				docX.ReplaceText("<שם ארגון>", order.OrganizationName);
				docX.ReplaceText("<סימוכין>", order.OrderId);
				docX.ReplaceText("<תאריך>", order.OrderCreationDate.ToString("dddd dd בMMMM yyyy"));


				var phones = new List<string>() { order.ContactPhone1, order.ContactPhone2 };
				var phoneNumbersList = phones.Where(x => !string.IsNullOrWhiteSpace(x)).ToList();
				var phoneNumbers = string.Join(", ", phoneNumbersList);

				docX.ReplaceText("<נייד>", phoneNumbers);


				if (!Directory.Exists(order.GetSavingDirectory()))
					Directory.CreateDirectory(order.GetSavingDirectory());

				docX.SaveAs(order.GetDocSavingPath());
			}
		}

		public static void CreateMailItem(Order order)
		{
			//יוצר מייל חדש עם הקובץ של הזמנת ההרצאה 
			//מען המייל הוא הכתובת מהטופס
			//ישלח מחנן סי-פויינט

			var outlookApp = new Outlook.Application();
			var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);
			foreach (Outlook.Account account in outlookApp.Session.Accounts)
			{
				// When the e-mail address matches, send the mail.
				if (account.SmtpAddress == EmailAddress)
				{
					mailItem.SendUsingAccount = account;
					mailItem.Subject = "חנן מלין - הצעת מחיר מספר " + order.OrderId;
					mailItem.To = order.Email;
					mailItem.Importance = Outlook.OlImportance.olImportanceLow;
					mailItem.Attachments.Add(order.GetPdfSavingPath(), Outlook.OlAttachmentType.olByValue, 1, order.GetPdfFileName());
					mailItem.Display(false);
					break;
				}
			}
		}

		//שומר וורד כפידיאף

		public static void ConvertDocxToPdf(string input, string output)
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
