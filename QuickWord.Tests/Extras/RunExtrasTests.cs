using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.Extras;

namespace QuickWord.Tests.Extras;

[TestClass]
public class RunExtrasTests
{
	[TestMethod]
	public void Formatting()
	{
		Run run = new Run().Bold().Italic().Color(new Color { Val = "Red" });
		RunFormatting formatting = run.CloneFormatting();

		Assert.AreEqual(3, run.RunProperties!.ChildElements.Count);

		run.ResetFormatting();
		Assert.AreEqual(0, run.ChildElements.Count);

		run.ApplyFormatting(formatting);
		Assert.IsTrue(run.BoldValue());
		Assert.IsTrue(run.ItalicValue());
		Assert.AreEqual("Red", run.GetColor()!.Val!.Value);
		Assert.AreEqual(3, run.RunProperties!.ChildElements.Count);

		var anotherFormatting = new RunFormatting { AllCaps = true, Color = new Color { Val = "Blue" } };

		run.ApplyFormatting(anotherFormatting, true);
		Assert.IsTrue(run.BoldValue());
		Assert.IsTrue(run.ItalicValue());
		Assert.IsTrue(run.AllCapsValue());
		Assert.AreEqual("Blue", run.GetColor()!.Val!.Value);
		Assert.AreEqual(4, run.RunProperties!.ChildElements.Count);

		run.ApplyFormatting(anotherFormatting);
		Assert.IsNull(run.BoldValue());
		Assert.IsNull(run.ItalicValue());
		Assert.IsTrue(run.AllCapsValue());
		Assert.AreEqual("Blue", run.GetColor()!.Val!.Value);
		Assert.AreEqual(2, run.RunProperties!.ChildElements.Count);

		run.ResetFormatting();
	}

	[TestMethod]
	public void Text()
	{
		Run run = new Run().Text($"Text\nwith\r\nline{Environment.NewLine}breaks", true);

		Assert.AreEqual(
			$"Text{Environment.NewLine}with{Environment.NewLine}line{Environment.NewLine}breaks",
			run.Text());

		run.Text(null);
		Assert.AreEqual(0, run.ChildElements.Count);
	}

	[TestMethod]
	public void Border()
	{
		Run run = new Run().Border(0.5);
		Assert.AreEqual(3U, run.GetBorder()!.Size!.Value);
	}

	[TestMethod]
	public void FillColor()
	{
		Run run = new Run().FillColor("Red");
		Assert.AreEqual("Red", run.GetShading()!.Fill!.Value);
	}

	[TestMethod]
	public void FontColor()
	{
		Run run = new Run().FontColor("Red");
		Assert.AreEqual("Red", run.GetColor()!.Val!.Value);
	}

	[TestMethod]
	public void FontFace()
	{
		Run run = new Run().FontFace("Comic Sans MS");
		Assert.AreEqual("Comic Sans MS", run.GetFonts()!.Ascii!.Value);
	}

	[TestMethod]
	public void Language()
	{
		Run run = new Run().Language("de-DE");
		Assert.AreEqual("de-DE", run.GetLanguages()!.Val!.Value);
	}

	[TestMethod]
	public void ManualWidth()
	{
		Run run = new Run().ManualWidth(100);
		Assert.AreEqual(100, run.ManualWidthValue(MeasuringUnits.Points));
		Assert.AreEqual(2000U, run.GetFitText()!.Val!.Value); // Twips

		run.ManualWidth(1, MeasuringUnits.Inches);
		Assert.AreEqual(1, run.ManualWidthValue(MeasuringUnits.Inches));
		Assert.AreEqual(2.54, run.ManualWidthValue(MeasuringUnits.Centimeters));

		run.ManualWidth(5.08, MeasuringUnits.Centimeters);
		Assert.AreEqual(5.08, run.ManualWidthValue(MeasuringUnits.Centimeters));
		Assert.AreEqual(2, run.ManualWidthValue(MeasuringUnits.Inches));
	}

	[TestMethod]
	public void Underline()
	{
		Run run = new Run().Underline(UnderlineValues.Single, "Red");
		Assert.AreEqual(UnderlineValues.Single, run.GetUnderline()!.Val!.Value);
		Assert.AreEqual("Red", run.GetUnderline()!.Color!.Value);
	}
}
