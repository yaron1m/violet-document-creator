using System;
using System.Diagnostics;
using System.IO;
using System.Web;
using Novacode;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace VioletDocumentCreator
{
	public interface IDocumentCreator
	{
		void CreateDocument(string[] args);
	}

	public class DocumentCreator: IDocumentCreator
	{
		private const string EmailAddress = "hanan@c-point.co.il";

		public void CreateDocument(string[] args)
		{
			var openWordForEdit = args[0].StartsWith(Consts.FileOpenSchemePrefix);

			var offerData = ExtractOfferData(args, openWordForEdit);
			var offer = new Offer(offerData);

			Console.WriteLine("Creating offers:");
			for (var topicIndex = 0; topicIndex < offer.Topic.Length; topicIndex++)
			{
				Console.WriteLine("Creating offer - " + offer.Topic[topicIndex]);
				CreateWordOrder(offer, topicIndex, openWordForEdit);

                Console.WriteLine("Converting to PDF - " + offer.Topic[topicIndex]);
				ConvertDocxToPdf(offer.GetDocSavingPathTemp(topicIndex), offer.GetPdfSavingPathTemp(topicIndex));

                Console.WriteLine("Copying word file - " + offer.Topic[topicIndex]);
                File.Copy(offer.GetDocSavingPathTemp(topicIndex), offer.GetDocSavingPath(topicIndex), true);

                Console.WriteLine("Copying PDF file - " + offer.Topic[topicIndex]);
                File.Copy(offer.GetPdfSavingPathTemp(topicIndex), offer.GetPdfSavingPath(topicIndex), true);

                Console.WriteLine("Cleanup of file - " + offer.Topic[topicIndex]);
                File.Delete(offer.GetDocSavingPathTemp(topicIndex));
                File.Delete(offer.GetPdfSavingPathTemp(topicIndex));

            }

            Console.WriteLine("Creating email");
			CreateMailItem(offer);
		}

		private string ExtractOfferData(string[] args, bool openWordForEdit)
		{
			var joinArguments = string.Join(" ", args);
			joinArguments = HttpUtility.UrlDecode(joinArguments);
			var rawData = joinArguments.Substring(openWordForEdit
				? Consts.FileOpenSchemePrefix.Length
				: Consts.VioletSchemePrefix.Length);
			return rawData;
		}

		private void CreateWordOrder(Offer offer, int topicIndex, bool openWordForEdit)
		{
			using (var docX = DocX.Load(offer.GetTamplatePath(topicIndex)))
			{
				docX.ReplaceText("<שם איש קשר>", offer.ContactName);
				docX.ReplaceText("<שם ארגון>", offer.OrganizationName);
				docX.ReplaceText("<סימוכין>", offer.Id);
				docX.ReplaceText("<תאריך>", offer.OrderCreationDate.ToString("dd בMMMM yyyy"));
				docX.ReplaceText("<נייד>", offer.PhoneNumbers);

				if (!Directory.Exists(offer.GetSavingDirectoryTemp()))
					Directory.CreateDirectory(offer.GetSavingDirectoryTemp());

				docX.SaveAs(offer.GetDocSavingPathTemp(topicIndex));

				if (openWordForEdit)
				{
					var process = Process.Start(offer.GetDocSavingPathTemp(topicIndex));
					process.WaitForExit();
				}
			}
		}

		public void CreateMailItem(Offer offer)
		{
			var outlookApp = new Outlook.Application();
			var mailItem = (Outlook.MailItem)outlookApp.CreateItem(Outlook.OlItemType.olMailItem);

			foreach (Outlook.Account account in outlookApp.Session.Accounts)
			{
				// When the e-mail address matches, send an email.
				if (account.SmtpAddress == EmailAddress)
				{
					mailItem.SendUsingAccount = account;
					mailItem.Subject = "חנן מלין - הצעת מחיר מספר " + offer.Id;
					mailItem.To = offer.Email;
					mailItem.Importance = Outlook.OlImportance.olImportanceLow;
					mailItem.Display(false);

					for (var topicIndex = 0; topicIndex < offer.Topic.Length; topicIndex++)
					{
						mailItem.Attachments.Add(offer.GetPdfSavingPath(topicIndex),
							Outlook.OlAttachmentType.olByValue, 1, offer.GetPdfFileName(topicIndex));
					}
					return;
				}
			}
		}

		public void ConvertDocxToPdf(string input, string output)
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
