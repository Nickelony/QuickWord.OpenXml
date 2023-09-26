using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;

namespace QuickWord.Tests.Drawings;

[TestClass]
public class ConversionTests
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
	public void ToAnchoredAndBack()
	{
		_drawing = _drawing.ToAnchoredDrawing();
		Assert.IsTrue(_drawing.IsAnchored());

		_drawing = _drawing.ToInlinedDrawing();
		Assert.IsTrue(_drawing.IsInlined());
	}

	[TestCleanup]
	public void Cleanup()
		=> _wordDocument.Dispose();
}
