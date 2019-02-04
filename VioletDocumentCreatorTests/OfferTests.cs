using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using VioletDocumentCreator;

namespace VioletDocumentCreatorTests
{
	[TestClass]
	public class OfferTests
	{
		private Offer _offer;

		[TestInitialize]
		public void Setup()
		{
			var data =
				"id=10&" +
				"topic=ISO - 14001#חשיבה? יצירתית&" +
				"email=yaron1m@gmail.com&" +
				"organizationName=שיא האיכות בע\"מ&" +
				"contactFirstName=ירון&" +
				"contactLastName=מלין&" +
				"contactPhone1=052-5415049&" +
				"contactPhone2=03-5354045&" +
				"orderCreationDate=2017-08-20T16:00:23.361Z";

			_offer = new Offer(data);
		}

		[TestMethod]
		public void Offer_CreateOffer_FieldsAreParsedCorrectly()
		{
			Assert.AreEqual(_offer.ContactFirstName, "ירון");
			Assert.AreEqual(_offer.ContactLastName, "מלין");
			Assert.AreEqual(_offer.ContactName, "ירון מלין");
			Assert.AreEqual(_offer.ContactPhone1, "052-5415049");
			Assert.AreEqual(_offer.ContactPhone2, "03-5354045");
			Assert.AreEqual(_offer.Email, "yaron1m@gmail.com");
			Assert.AreEqual(_offer.Id, "10");
			Assert.IsTrue(_offer.OrderCreationDate.Year == 2017);
			Assert.IsTrue(_offer.OrderCreationDate.Month == 8);
			Assert.IsTrue(_offer.OrderCreationDate.Day == 20);
			Assert.AreEqual(_offer.Topic.Length, 2);
			Assert.AreEqual(_offer.Topic[0], "ISO - 14001");
			Assert.AreEqual(_offer.Topic[1], "חשיבה? יצירתית");
			Assert.AreEqual(_offer.OrganizationName, "שיא האיכות בע\"מ");
		}

		[TestMethod]
		public void GetDocSavingPath_InvalidCharacters_ValidPath()
		{
			var expected0 = Directory.GetCurrentDirectory() + 
				"\\הצעות מחיר\\ISO - 14001\\ISO - 14001 - שיא האיכות בעמ - ירון מלין 10.docx";
			Assert.AreEqual(_offer.GetDocSavingPath(0), expected0);

			var expected1 = Directory.GetCurrentDirectory() + 
				"\\הצעות מחיר\\חשיבה יצירתית\\חשיבה יצירתית - שיא האיכות בעמ - ירון מלין 10.docx";
			Assert.AreEqual(_offer.GetDocSavingPath(1), expected1);
		}

		[TestMethod]
		public void GetPdfFileName_InvalidCharacters_ValidPath()
		{
			var expected0 = Directory.GetCurrentDirectory() + 
				"\\הצעות מחיר\\ISO - 14001\\ISO - 14001 - שיא האיכות בעמ - ירון מלין 10.pdf";
			Assert.AreEqual(_offer.GetPdfSavingPath(0), expected0);

			var expected1 = Directory.GetCurrentDirectory() + 
				"\\הצעות מחיר\\חשיבה יצירתית\\חשיבה יצירתית - שיא האיכות בעמ - ירון מלין 10.pdf";
			Assert.AreEqual(_offer.GetPdfSavingPath(1), expected1);
		}
	}
}