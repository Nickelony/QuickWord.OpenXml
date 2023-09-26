using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace QuickWord.Tests.Drawings;

[TestClass]
public class WrappingTests
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
	public void Wrapping()
	{
		_drawing.ToAnchoredDrawing();

		_drawing.SquareWrapping(0.5, 0.5, 0.5, 0.5, ImageMeasuringUnits.Inches);
		Assert.AreEqual(WrappingType.Square, _drawing.GetWrappingType());

		_drawing.TightWrapping(new DW.WrapPolygon(), 0.5, 0.5, ImageMeasuringUnits.Inches);
		Assert.AreEqual(WrappingType.Tight, _drawing.GetWrappingType());

		_drawing.ThroughWrapping(new DW.WrapPolygon(), 0.5, 0.5, ImageMeasuringUnits.Inches);
		Assert.AreEqual(WrappingType.Through, _drawing.GetWrappingType());

		_drawing.TopAndBottomWrapping(0.5, 0.5, ImageMeasuringUnits.Inches);
		Assert.AreEqual(WrappingType.TopAndBottom, _drawing.GetWrappingType());

		_drawing.NoTextWrapping();
		Assert.AreEqual(WrappingType.None, _drawing.GetWrappingType());
	}

	[TestCleanup]
	public void Cleanup()
		=> _wordDocument.Dispose();
}
