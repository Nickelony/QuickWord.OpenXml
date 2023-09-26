// Ignore Spelling: Uncropped

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;

namespace QuickWord.Tests.Drawings;

[TestClass]
public class CroppingTests
{
	private WordprocessingDocument _wordDocument = null!;
	private Drawing _drawing = null!;

	[TestInitialize]
	public void Initialize()
	{
		_wordDocument = WordprocessingDocument.Create("Test.docx", WordprocessingDocumentType.Document);
		Body body = _wordDocument.CreateBody();

		_drawing = _wordDocument.MainDocumentPart!.CreateImage("Assets/Icon.png", ImagePartType.Png);
		body.AppendChild(new Paragraph(_drawing));
	}

	[TestMethod]
	public void Cropping()
	{
		_drawing.Cropping(new Cropping { LeftFactor = 0.45, TopFactor = 0.33, RightFactor = 0.25, BottomFactor = 0.1 });
		Cropping? actualCropping = _drawing.GetCropping();

		Assert.AreEqual(0.45, actualCropping!.LeftFactor);
		Assert.AreEqual(0.33, actualCropping!.TopFactor);
		Assert.AreEqual(0.25, actualCropping!.RightFactor);
		Assert.AreEqual(0.1, actualCropping!.BottomFactor);

		_drawing.ResetCropping();
		Assert.IsNull(_drawing.GetCropping());
	}

	[TestMethod]
	public void UncroppedSizes()
	{
		// Image size is 1024x1024
		_drawing.Resize(512, 512);
		// Now the image size is 512x512
		_drawing.Cropping(new Cropping { LeftFactor = 0.25, TopFactor = 0.25, RightFactor = 0.25, BottomFactor = 0.25 });
		// Cropped image size is now 256x256

		Assert.AreEqual(256, _drawing.GetWidth());
		Assert.AreEqual(256, _drawing.GetWidth());

		Assert.AreEqual(512, _drawing.GetUncroppedWidth());
		Assert.AreEqual(512, _drawing.GetUncroppedHeight());

		Assert.AreEqual(1024, _drawing.GetOriginalWidth());
		Assert.AreEqual(1024, _drawing.GetOriginalHeight());
	}

	[TestCleanup]
	public void Cleanup()
		=> _wordDocument.Dispose();
}
