// Ignore Spelling: Kinsoku

using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;

namespace QuickWord.Tests;

[TestClass]
public class ParagraphTests
{
	[TestMethod]
	public void AdjustRightIndent()
	{
		Paragraph paragraph = new Paragraph().AdjustRightIndent(false);

		Assert.IsFalse(paragraph.AdjustRightIndentValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.AdjustRightIndent!.Val!.Value);

		paragraph.AdjustRightIndent(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void AutoSpaceDE()
	{
		Paragraph paragraph = new Paragraph().AutoSpaceDE(false);

		Assert.IsFalse(paragraph.AutoSpaceDEValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.AutoSpaceDE!.Val!.Value);

		paragraph.AutoSpaceDE(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void AutoSpaceDN()
	{
		Paragraph paragraph = new Paragraph().AutoSpaceDN(false);

		Assert.IsFalse(paragraph.AutoSpaceDNValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.AutoSpaceDN!.Val!.Value);

		paragraph.AutoSpaceDN(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void BiDirectional()
	{
		Paragraph paragraph = new Paragraph().BiDirectional(false);

		Assert.IsFalse(paragraph.BiDirectionalValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.BiDi!.Val!.Value);

		paragraph.BiDirectional(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void ConditionalFormatStyle()
	{
		Paragraph paragraph = new Paragraph().ConditionalFormatStyle(new ConditionalFormatStyle { FirstRow = true });

		Assert.IsTrue(paragraph.GetConditionalFormatStyle()!.FirstRow!.Value);
		Assert.IsTrue(paragraph.ParagraphProperties!.ConditionalFormatStyle!.FirstRow!.Value);

		paragraph.ConditionalFormatStyle(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void ContextualSpacing()
	{
		Paragraph paragraph = new Paragraph().ContextualSpacing(false);

		Assert.IsFalse(paragraph.ContextualSpacingValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.ContextualSpacing!.Val!.Value);

		paragraph.ContextualSpacing(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void DivId()
	{
		Paragraph paragraph = new Paragraph().DivId("id");

		Assert.AreEqual("id", paragraph.DivIdValue());
		Assert.AreEqual("id", paragraph.ParagraphProperties!.DivId!.Val!.Value);

		paragraph.DivId(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void FrameProperties()
	{
		Paragraph paragraph = new Paragraph().FrameProperties(new FrameProperties { Width = "100" });

		Assert.AreEqual("100", paragraph.GetFrameProperties()!.Width!.Value);
		Assert.AreEqual("100", paragraph.ParagraphProperties!.FrameProperties!.Width!.Value);

		paragraph.FrameProperties(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Indentation()
	{
		Paragraph paragraph = new Paragraph().Indentation(new Indentation { Left = "100" });

		Assert.AreEqual("100", paragraph.GetIndentation()!.Left!.Value);
		Assert.AreEqual("100", paragraph.ParagraphProperties!.Indentation!.Left!.Value);

		paragraph.Indentation(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Justification()
	{
		Paragraph paragraph = new Paragraph().Justification(JustificationValues.Center);

		Assert.AreEqual(JustificationValues.Center, paragraph.JustificationValue());
		Assert.AreEqual(JustificationValues.Center, paragraph.ParagraphProperties!.Justification!.Val!.Value);

		paragraph.Justification(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void KeepLinesTogether()
	{
		Paragraph paragraph = new Paragraph().KeepLinesTogether(false);

		Assert.IsFalse(paragraph.KeepLinesTogetherValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.KeepLines!.Val!.Value);

		paragraph.KeepLinesTogether(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void KeepWithNext()
	{
		Paragraph paragraph = new Paragraph().KeepWithNext(false);

		Assert.IsFalse(paragraph.KeepWithNextValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.KeepNext!.Val!.Value);

		paragraph.KeepWithNext(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Kinsoku()
	{
		Paragraph paragraph = new Paragraph().Kinsoku(false);

		Assert.IsFalse(paragraph.KinsokuValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.Kinsoku!.Val!.Value);

		paragraph.Kinsoku(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void MirrorIndents()
	{
		Paragraph paragraph = new Paragraph().MirrorIndents(false);

		Assert.IsFalse(paragraph.MirrorIndentsValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.MirrorIndents!.Val!.Value);

		paragraph.MirrorIndents(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void NumberingProperties()
	{
		Paragraph paragraph = new Paragraph()
			.NumberingProperties(new NumberingProperties { NumberingId = new NumberingId { Val = 2 } });

		Assert.AreEqual(2, paragraph.GetNumberingProperties()!.NumberingId!.Val!.Value);
		Assert.AreEqual(2, paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value);

		paragraph.NumberingProperties(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void OutlineLevel()
	{
		Paragraph paragraph = new Paragraph().OutlineLevel(OutlineLevelValues.Level3);

		Assert.AreEqual(OutlineLevelValues.Level3, paragraph.OutlineLevelValue());
		Assert.AreEqual(OutlineLevelValues.Level3, (OutlineLevelValues)paragraph.ParagraphProperties!.OutlineLevel!.Val!.Value);

		paragraph.OutlineLevel(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void OverflowPunctuation()
	{
		Paragraph paragraph = new Paragraph().OverflowPunctuation(false);

		Assert.IsFalse(paragraph.OverflowPunctuationValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.OverflowPunctuation!.Val!.Value);

		paragraph.OverflowPunctuation(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void PageBreakBefore()
	{
		Paragraph paragraph = new Paragraph().PageBreakBefore(false);

		Assert.IsFalse(paragraph.PageBreakBeforeValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.PageBreakBefore!.Val!.Value);

		paragraph.PageBreakBefore(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Borders()
	{
		Paragraph paragraph = new Paragraph().Borders(new ParagraphBorders(new TopBorder { Val = BorderValues.Single }));

		Assert.AreEqual(BorderValues.Single, paragraph.GetBorders()!.TopBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.Single, paragraph.ParagraphProperties!.ParagraphBorders!.TopBorder!.Val!.Value);

		paragraph.Borders(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Style()
	{
		Paragraph paragraph = new Paragraph().Style("Heading1");

		Assert.AreEqual("Heading1", paragraph.StyleValue());
		Assert.AreEqual("Heading1", paragraph.ParagraphProperties!.ParagraphStyleId!.Val!.Value);

		paragraph.Style(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void MarkRunProperties()
	{
		Paragraph paragraph = new Paragraph().MarkRunProperties(new ParagraphMarkRunProperties());

		Assert.IsNotNull(paragraph.GetMarkRunProperties());
		Assert.IsNotNull(paragraph.ParagraphProperties!.ParagraphMarkRunProperties);

		paragraph.MarkRunProperties(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void SectionProperties()
	{
		Paragraph paragraph = new Paragraph().SectionProperties(new SectionProperties());

		Assert.IsNotNull(paragraph.GetSectionProperties());
		Assert.IsNotNull(paragraph.ParagraphProperties!.SectionProperties);

		paragraph.SectionProperties(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Shading()
	{
		Paragraph paragraph = new Paragraph().Shading(new Shading { Fill = "Red" });

		Assert.AreEqual("Red", paragraph.GetShading()!.Fill!.Value);
		Assert.AreEqual("Red", paragraph.ParagraphProperties!.Shading!.Fill!.Value);

		paragraph.Shading(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void SnapToGrid()
	{
		Paragraph paragraph = new Paragraph().SnapToGrid(false);

		Assert.IsFalse(paragraph.SnapToGridValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.SnapToGrid!.Val!.Value);

		paragraph.SnapToGrid(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Spacing()
	{
		Paragraph paragraph = new Paragraph().Spacing(new SpacingBetweenLines { Line = "100" });

		Assert.AreEqual("100", paragraph.GetSpacing()!.Line!.Value);
		Assert.AreEqual("100", paragraph.ParagraphProperties!.SpacingBetweenLines!.Line!.Value);

		paragraph.Spacing(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void SuppressAutoHyphenation()
	{
		Paragraph paragraph = new Paragraph().SuppressAutoHyphenation(false);

		Assert.IsFalse(paragraph.SuppressAutoHyphenationValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.SuppressAutoHyphens!.Val!.Value);

		paragraph.SuppressAutoHyphenation(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void SuppressLineNumbers()
	{
		Paragraph paragraph = new Paragraph().SuppressLineNumbers(false);

		Assert.IsFalse(paragraph.SuppressLineNumbersValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.SuppressLineNumbers!.Val!.Value);

		paragraph.SuppressLineNumbers(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void SuppressOverlapping()
	{
		Paragraph paragraph = new Paragraph().SuppressOverlapping(false);

		Assert.IsFalse(paragraph.SuppressOverlappingValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.SuppressOverlap!.Val!.Value);

		paragraph.SuppressOverlapping(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void Tabs()
	{
		Paragraph paragraph = new Paragraph().Tabs(new Tabs());

		Assert.IsNotNull(paragraph.GetTabs());
		Assert.IsNotNull(paragraph.ParagraphProperties!.Tabs);

		paragraph.Tabs(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void VerticalTextAlignment()
	{
		Paragraph paragraph = new Paragraph().VerticalTextAlignment(VerticalTextAlignmentValues.Center);

		Assert.AreEqual(VerticalTextAlignmentValues.Center, paragraph.VerticalTextAlignmentValue());
		Assert.AreEqual(VerticalTextAlignmentValues.Center, paragraph.ParagraphProperties!.TextAlignment!.Val!.Value);

		paragraph.VerticalTextAlignment(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void TextBoxTightWrap()
	{
		Paragraph paragraph = new Paragraph().TextBoxTightWrap(TextBoxTightWrapValues.AllLines);

		Assert.AreEqual(TextBoxTightWrapValues.AllLines, paragraph.TextBoxTightWrapValue());
		Assert.AreEqual(TextBoxTightWrapValues.AllLines, paragraph.ParagraphProperties!.TextBoxTightWrap!.Val!.Value);

		paragraph.TextBoxTightWrap(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void TextDirection()
	{
		Paragraph paragraph = new Paragraph().TextDirection(TextDirectionValues.BottomToTopLeftToRight);

		Assert.AreEqual(TextDirectionValues.BottomToTopLeftToRight, paragraph.TextDirectionValue());
		Assert.AreEqual(TextDirectionValues.BottomToTopLeftToRight, paragraph.ParagraphProperties!.TextDirection!.Val!.Value);

		paragraph.TextDirection(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void TopLinePunctuation()
	{
		Paragraph paragraph = new Paragraph().TopLinePunctuation(false);

		Assert.IsFalse(paragraph.TopLinePunctuationValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.TopLinePunctuation!.Val!.Value);

		paragraph.TopLinePunctuation(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void WidowControl()
	{
		Paragraph paragraph = new Paragraph().WidowControl(false);

		Assert.IsFalse(paragraph.WidowControlValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.WidowControl!.Val!.Value);

		paragraph.WidowControl(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}

	[TestMethod]
	public void WordWrap()
	{
		Paragraph paragraph = new Paragraph().WordWrap(false);

		Assert.IsFalse(paragraph.WordWrapValue());
		Assert.IsFalse(paragraph.ParagraphProperties!.WordWrap!.Val!.Value);

		paragraph.WordWrap(null); // Removes ParagraphProperties since it's the last element
		Assert.IsNull(paragraph.ParagraphProperties);
	}
}
