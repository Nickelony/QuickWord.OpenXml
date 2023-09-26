using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;

namespace QuickWord.Tests.Drawings;

[TestClass]
public class EffectTests
{
	private WordprocessingDocument _wordDocument = null!;
	private Drawing _drawing = null!;

	[TestInitialize]
	public void Initialize()
	{
		_wordDocument = WordprocessingDocument.Create("Test.docx", WordprocessingDocumentType.Document);
		_wordDocument.CreateBody();

		_drawing = _wordDocument.MainDocumentPart!.CreateImage("Assets/Icon.png", ImagePartType.Png);
	}

	[TestMethod]
	public void Opacity()
	{
		_drawing.Opacity(0.66);
		Assert.AreEqual(0.66, _drawing.OpacityValue());
	}

	// TODO: Add border tests

	[TestCleanup]
	public void Cleanup()
		=> _wordDocument.Dispose();
}
