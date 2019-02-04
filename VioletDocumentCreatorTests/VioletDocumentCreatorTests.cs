using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Moq;
using VioletDocumentCreator;

namespace VioletDocumentCreatorTests
{
	[TestClass]
	public class VioletDocumentCreatorTests
	{
		private VioletDocumentCreator.VioletDocumentCreator _target;
		private Mock<IDocumentCreator> _documentCreatorMock;
		private Mock<IUriSchemeRegisterer> _uriSchemeRegistererMock;

		[TestInitialize]
		public void Setup()
		{
			_documentCreatorMock = new Mock<IDocumentCreator>();
			_uriSchemeRegistererMock = new Mock<IUriSchemeRegisterer>();

			_target = new VioletDocumentCreator.VioletDocumentCreator(_documentCreatorMock.Object, _uriSchemeRegistererMock.Object);
		}

		[TestMethod]
		public void Run_NoArguments_DocumentCreatorIsNotCalled()
		{
			_target.Run(new string[0]);

			_documentCreatorMock.Verify(m => m.CreateDocument(It.IsAny<string[]>()), Times.Never);
		}

		[TestMethod]
		public void Run_ArgumentStartsWithVioletPrefix_DocumentCreatorIsCalled()
		{
			var arg0 = "violet:Bla";

			_target.Run(new[]{ arg0 });

			_documentCreatorMock.Verify(m => m.CreateDocument(It.Is<string[]>(args => args[0].Equals(arg0))));
		}

		[TestMethod]
		public void Run_InstallCommand_UriSchemeInstallerIsCalled()
		{
			var arg0 = "install";

			_target.Run(new[]{ arg0 });

			_uriSchemeRegistererMock.Verify(m => m.Install());
		}

		[TestMethod]
		public void Run_UninstallCommand_UriSchemeInstallerIsCalled()
		{
			var arg0 = "uninstall";

			_target.Run(new[]{ arg0 });

			_uriSchemeRegistererMock.Verify(m => m.Uninstall());
		}

	}
}