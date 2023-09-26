using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.Extras;

namespace QuickWord.Tests.Extras;

[TestClass]
public class ParagraphExtrasTests
{
	[TestMethod]
	public void Formatting()
	{
		Paragraph paragraph = new Paragraph().AdjustRightIndent().LineSpacing(3).Shading(new Shading { Fill = "Red" });
		ParagraphFormatting formatting = paragraph.CloneFormatting();

		Assert.AreEqual(3, paragraph.ParagraphProperties!.ChildElements.Count);

		paragraph.ResetFormatting();
		Assert.AreEqual(0, paragraph.ChildElements.Count);

		paragraph.ApplyFormatting(formatting);
		Assert.IsTrue(paragraph.AdjustRightIndentValue());
		Assert.AreEqual(3, paragraph.LineSpacingValue(LineMeasuringUnits.WholeLines));
		Assert.AreEqual("Red", paragraph.GetShading()!.Fill!.Value);
		Assert.AreEqual(3, paragraph.ParagraphProperties!.ChildElements.Count);

		var anotherFormatting = new ParagraphFormatting { SuppressAutoHyphenation = true, Shading = new Shading { Fill = "Blue" } };

		paragraph.ApplyFormatting(anotherFormatting, true);
		Assert.IsTrue(paragraph.AdjustRightIndentValue());
		Assert.AreEqual(3, paragraph.LineSpacingValue(LineMeasuringUnits.WholeLines));
		Assert.IsTrue(paragraph.SuppressAutoHyphenationValue());
		Assert.AreEqual("Blue", paragraph.GetShading()!.Fill!.Value);
		Assert.AreEqual(4, paragraph.ParagraphProperties!.ChildElements.Count);

		paragraph.ApplyFormatting(anotherFormatting);
		Assert.IsNull(paragraph.AdjustRightIndentValue());
		Assert.IsNull(paragraph.LineSpacingValue(LineMeasuringUnits.WholeLines));
		Assert.IsTrue(paragraph.SuppressAutoHyphenationValue());
		Assert.AreEqual("Blue", paragraph.GetShading()!.Fill!.Value);
		Assert.AreEqual(2, paragraph.ParagraphProperties!.ChildElements.Count);

		paragraph.ResetFormatting();
	}

	[TestMethod]
	public void GetText()
	{
		var paragraph = new Paragraph(
			new Run().Text("This is a"),
			new Run().Text(" single paragraph "),
			new Run().Text("with 3 different"),
			new Run().Text(" Runs in it."),
			new Run().Text("\nThe 5th Run is on another line.", true)
		);

		Assert.AreEqual(
			$"This is a single paragraph with 3 different Runs in it.{Environment.NewLine}The 5th Run is on another line.",
			paragraph.GetText());
	}

	[TestMethod]
	public void Spacing()
	{
		Paragraph paragraph = new Paragraph()
			.LineSpacing(3, LineMeasuringUnits.WholeLines)
			.SpacingBefore(12, LineMeasuringUnits.Points)
			.SpacingAfter(6, LineMeasuringUnits.Points);

		Assert.AreEqual(36, paragraph.LineSpacingValue(LineMeasuringUnits.Points));
		Assert.AreEqual("720", paragraph.GetSpacing()!.Line!.Value);

		Assert.AreEqual(1, paragraph.SpacingBeforeValue(LineMeasuringUnits.WholeLines));
		Assert.AreEqual("240", paragraph.GetSpacing()!.Before!.Value);

		Assert.AreEqual(0.5, paragraph.SpacingAfterValue(LineMeasuringUnits.WholeLines));
		Assert.AreEqual("120", paragraph.GetSpacing()!.After!.Value);
	}

	[TestMethod]
	public void Indentation()
	{
		Paragraph paragraph = new Paragraph()
			.LeftIndentation(2.54, IndentationUnits.Centimeters)
			.RightIndentation(1, IndentationUnits.Inches);

		Assert.AreEqual(1, paragraph.LeftIndentationValue(IndentationUnits.Inches));
		Assert.AreEqual("1440", paragraph.GetIndentation()!.Left!.Value);

		Assert.AreEqual(2.54, paragraph.RightIndentationValue(IndentationUnits.Centimeters));
		Assert.AreEqual("1440", paragraph.GetIndentation()!.Right!.Value);

		paragraph
			.LeftIndentation(3, IndentationUnits.Characters)
			.RightIndentation(5, IndentationUnits.Characters);

		Assert.AreEqual(3, paragraph.LeftIndentationValue(IndentationUnits.Characters));
		Assert.AreEqual(3, paragraph.GetIndentation()!.LeftChars!.Value);

		Assert.AreEqual(5, paragraph.RightIndentationValue(IndentationUnits.Characters));
		Assert.AreEqual(5, paragraph.GetIndentation()!.RightChars!.Value);
	}

	[TestMethod]
	public void FillColor()
	{
		Paragraph paragraph = new Paragraph().FillColor("Red");
		Assert.AreEqual("Red", paragraph.GetShading()!.Fill!.Value);
	}

	[TestMethod]
	public void Borders()
	{
		Paragraph paragraph = new Paragraph()
			.LeftBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.TopBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.RightBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.BottomBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.BarBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.BetweenBorder(1.5, BorderValues.DashDotStroked, "Red", 1);

		Assert.AreEqual(9U, paragraph.GetBorders()!.LeftBorder!.Size!.Value);
		Assert.AreEqual(9U, paragraph.GetBorders()!.TopBorder!.Size!.Value);
		Assert.AreEqual(9U, paragraph.GetBorders()!.RightBorder!.Size!.Value);
		Assert.AreEqual(9U, paragraph.GetBorders()!.BottomBorder!.Size!.Value);
		Assert.AreEqual(9U, paragraph.GetBorders()!.BarBorder!.Size!.Value);
		Assert.AreEqual(9U, paragraph.GetBorders()!.BetweenBorder!.Size!.Value);

		Assert.AreEqual(BorderValues.DashDotStroked, paragraph.GetBorders()!.LeftBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, paragraph.GetBorders()!.TopBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, paragraph.GetBorders()!.RightBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, paragraph.GetBorders()!.BottomBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, paragraph.GetBorders()!.BarBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, paragraph.GetBorders()!.BetweenBorder!.Val!.Value);

		Assert.AreEqual("Red", paragraph.GetBorders()!.LeftBorder!.Color!.Value);
		Assert.AreEqual("Red", paragraph.GetBorders()!.TopBorder!.Color!.Value);
		Assert.AreEqual("Red", paragraph.GetBorders()!.RightBorder!.Color!.Value);
		Assert.AreEqual("Red", paragraph.GetBorders()!.BottomBorder!.Color!.Value);
		Assert.AreEqual("Red", paragraph.GetBorders()!.BarBorder!.Color!.Value);
		Assert.AreEqual("Red", paragraph.GetBorders()!.BetweenBorder!.Color!.Value);

		Assert.AreEqual(1U, paragraph.GetBorders()!.LeftBorder!.Space!.Value);
		Assert.AreEqual(1U, paragraph.GetBorders()!.TopBorder!.Space!.Value);
		Assert.AreEqual(1U, paragraph.GetBorders()!.RightBorder!.Space!.Value);
		Assert.AreEqual(1U, paragraph.GetBorders()!.BottomBorder!.Space!.Value);
		Assert.AreEqual(1U, paragraph.GetBorders()!.BarBorder!.Space!.Value);
		Assert.AreEqual(1U, paragraph.GetBorders()!.BetweenBorder!.Space!.Value);

		paragraph.ResetFormatting();

		paragraph
			.LeftBorder(new LeftBorder { Size = 6 })
			.TopBorder(new TopBorder { Size = 6 })
			.RightBorder(new RightBorder { Size = 6 })
			.BottomBorder(new BottomBorder { Size = 6 })
			.BarBorder(new BarBorder { Size = 6 })
			.BetweenBorder(new BetweenBorder { Size = 6 });

		Assert.AreEqual(6U, paragraph.GetBorders()!.LeftBorder!.Size!.Value);
		Assert.AreEqual(6U, paragraph.GetBorders()!.TopBorder!.Size!.Value);
		Assert.AreEqual(6U, paragraph.GetBorders()!.RightBorder!.Size!.Value);
		Assert.AreEqual(6U, paragraph.GetBorders()!.BottomBorder!.Size!.Value);
		Assert.AreEqual(6U, paragraph.GetBorders()!.BarBorder!.Size!.Value);
		Assert.AreEqual(6U, paragraph.GetBorders()!.BetweenBorder!.Size!.Value);
	}
}
