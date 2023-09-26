using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;

namespace QuickWord.Tests.Drawings;

[TestClass]
public class AnchorSpecificTests
{
	private WordprocessingDocument _wordDocument = null!;
	private Drawing _anchoredDrawing = null!;

	[TestInitialize]
	public void Initialize()
	{
		_wordDocument = WordprocessingDocument.Create("Test.docx", WordprocessingDocumentType.Document);
		_wordDocument.CreateBody();

		_anchoredDrawing = _wordDocument.MainDocumentPart!.CreateImage("Assets/Icon.png", ImagePartType.Png).ToAnchoredDrawing();
	}

	[TestMethod]
	public void AllowOverlapping()
	{
		_anchoredDrawing.AllowOverlapping(false);

		Assert.IsFalse(_anchoredDrawing.AllowOverlappingValue());
		Assert.IsFalse(_anchoredDrawing.Anchor!.AllowOverlap!.Value);
	}

	[TestMethod]
	public void BehindText()
	{
		_anchoredDrawing.BehindText(false);

		Assert.IsFalse(_anchoredDrawing.BehindTextValue());
		Assert.IsFalse(_anchoredDrawing.Anchor!.BehindDoc!.Value);
	}

	[TestMethod]
	public void LayoutInTableCell()
	{
		_anchoredDrawing.LayoutInTableCell(false);

		Assert.IsFalse(_anchoredDrawing.LayoutInTableCellValue());
		Assert.IsFalse(_anchoredDrawing.Anchor!.LayoutInCell!.Value);
	}

	[TestMethod]
	public void Locked()
	{
		_anchoredDrawing.Locked(false);

		Assert.IsFalse(_anchoredDrawing.LockedValue());
		Assert.IsFalse(_anchoredDrawing.Anchor!.Locked!.Value);
	}

	[TestCleanup]
	public void Cleanup()
		=> _wordDocument.Dispose();
}
