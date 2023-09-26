using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;

namespace QuickWord.Tests;

[TestClass]
public class RunTests
{
	[TestMethod]
	public void Bold()
	{
		Run run = new Run().Bold(false);

		Assert.IsFalse(run.BoldValue());
		Assert.IsFalse(run.RunProperties!.Bold!.Val!.Value);

		run.Bold(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void BoldComplexScript()
	{
		Run run = new Run().BoldComplexScript(false);

		Assert.IsFalse(run.BoldComplexScriptValue());
		Assert.IsFalse(run.RunProperties!.BoldComplexScript!.Val!.Value);

		run.BoldComplexScript(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Border()
	{
		Run run = new Run().Border(new Border { Val = BorderValues.Single });

		Assert.AreEqual(BorderValues.Single, run.GetBorder()!.Val!.Value);
		Assert.AreEqual(BorderValues.Single, run.RunProperties!.Border!.Val!.Value);

		run.Border(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void AllCaps()
	{
		Run run = new Run().AllCaps(false);

		Assert.IsFalse(run.AllCapsValue());
		Assert.IsFalse(run.RunProperties!.Caps!.Val!.Value);

		run.AllCaps(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Color()
	{
		Run run = new Run().Color(new Color { Val = "Red" });

		Assert.AreEqual("Red", run.GetColor()!.Val!.Value);
		Assert.AreEqual("Red", run.RunProperties!.Color!.Val!.Value);

		run.Color(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void ComplexScript()
	{
		Run run = new Run().ComplexScript(false);

		Assert.IsFalse(run.ComplexScriptValue());
		Assert.IsFalse(run.RunProperties!.ComplexScript!.Val!.Value);

		run.ComplexScript(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void DoubleStrike()
	{
		Run run = new Run().DoubleStrike(false);

		Assert.IsFalse(run.DoubleStrikeValue());
		Assert.IsFalse(run.RunProperties!.DoubleStrike!.Val!.Value);

		run.DoubleStrike(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void EastAsianLayout()
	{
		Run run = new Run().EastAsianLayout(new EastAsianLayout { Vertical = true });

		Assert.IsTrue(run.GetEastAsianLayout()!.Vertical!.Value);
		Assert.IsTrue(run.RunProperties!.EastAsianLayout!.Vertical!.Value);

		run.EastAsianLayout(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void TextEffect()
	{
		Run run = new Run().TextEffect(TextEffectValues.Lights);

		Assert.AreEqual(TextEffectValues.Lights, run.TextEffectValue());
		Assert.AreEqual(TextEffectValues.Lights, run.RunProperties!.TextEffect!.Val!.Value);

		run.TextEffect(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void EmphasisMark()
	{
		Run run = new Run().EmphasisMark(EmphasisMarkValues.Circle);

		Assert.AreEqual(EmphasisMarkValues.Circle, run.EmphasisMarkValue());
		Assert.AreEqual(EmphasisMarkValues.Circle, run.RunProperties!.Emphasis!.Val!.Value);

		run.EmphasisMark(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Emboss()
	{
		Run run = new Run().Emboss(false);

		Assert.IsFalse(run.EmbossValue());
		Assert.IsFalse(run.RunProperties!.Emboss!.Val!.Value);

		run.Emboss(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void FitText()
	{
		Run run = new Run().FitText(new FitText { Val = 100U });

		Assert.AreEqual(100U, run.GetFitText()!.Val!.Value);
		Assert.AreEqual(100U, run.RunProperties!.FitText!.Val!.Value);

		run.FitText(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void HighlightColor()
	{
		Run run = new Run().HighlightColor(HighlightColorValues.Cyan);

		Assert.AreEqual(HighlightColorValues.Cyan, run.HighlightColorValue());
		Assert.AreEqual(HighlightColorValues.Cyan, run.RunProperties!.Highlight!.Val!.Value);

		run.HighlightColor(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Italic()
	{
		Run run = new Run().Italic(false);

		Assert.IsFalse(run.ItalicValue());
		Assert.IsFalse(run.RunProperties!.Italic!.Val!.Value);

		run.Italic(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void ItalicComplexScript()
	{
		Run run = new Run().ItalicComplexScript(false);

		Assert.IsFalse(run.ItalicComplexScriptValue());
		Assert.IsFalse(run.RunProperties!.ItalicComplexScript!.Val!.Value);

		run.ItalicComplexScript(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Imprint()
	{
		Run run = new Run().Imprint(false);

		Assert.IsFalse(run.ImprintValue());
		Assert.IsFalse(run.RunProperties!.Imprint!.Val!.Value);

		run.Imprint(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Kerning()
	{
		Run run = new Run().Kerning(16, TextMeasuringUnits.HalfPoints);
		Assert.AreEqual(16, run.KerningValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual(8, run.KerningValue(TextMeasuringUnits.Points));
		Assert.AreEqual(16U, run.RunProperties!.Kern!.Val!.Value);

		run.Kerning(32, TextMeasuringUnits.Points);
		Assert.AreEqual(32, run.KerningValue(TextMeasuringUnits.Points));
		Assert.AreEqual(64, run.KerningValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual(64U, run.RunProperties!.Kern!.Val!.Value);

		run.Kerning(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Languages()
	{
		Run run = new Run().Languages(new Languages { Val = "de-DE" });

		Assert.AreEqual("de-DE", run.GetLanguages()!.Val!.Value);
		Assert.AreEqual("de-DE", run.RunProperties!.Languages!.Val!.Value);

		run.Languages(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void NoProofing()
	{
		Run run = new Run().NoProofing(false);

		Assert.IsFalse(run.NoProofingValue());
		Assert.IsFalse(run.RunProperties!.NoProof!.Val!.Value);

		run.NoProofing(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void OfficeMath()
	{
		Run run = new Run().OfficeMath(false);

		Assert.IsFalse(run.OfficeMathValue());
		Assert.IsFalse(run.RunProperties!.GetFirstChild<OfficeMath>()!.Val!.Value);

		run.OfficeMath(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Outline()
	{
		Run run = new Run().Outline(false);

		Assert.IsFalse(run.OutlineValue());
		Assert.IsFalse(run.RunProperties!.Outline!.Val!.Value);

		run.Outline(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void VerticalPosition()
	{
		Run run = new Run().VerticalPosition(16, TextMeasuringUnits.HalfPoints);
		Assert.AreEqual(16, run.VerticalPositionValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual(8, run.VerticalPositionValue(TextMeasuringUnits.Points));
		Assert.AreEqual("16", run.RunProperties!.Position!.Val!.Value);

		run.VerticalPosition(32, TextMeasuringUnits.Points);
		Assert.AreEqual(32, run.VerticalPositionValue(TextMeasuringUnits.Points));
		Assert.AreEqual(64, run.VerticalPositionValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual("64", run.RunProperties!.Position!.Val!.Value);

		run.VerticalPosition(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Fonts()
	{
		Run run = new Run().Fonts(new RunFonts { Ascii = "Comic Sans MS" });

		Assert.AreEqual("Comic Sans MS", run.GetFonts()!.Ascii!.Value);
		Assert.AreEqual("Comic Sans MS", run.RunProperties!.RunFonts!.Ascii!.Value);

		run.Fonts(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Style()
	{
		Run run = new Run().Style("Heading1Char");

		Assert.AreEqual("Heading1Char", run.StyleValue());
		Assert.AreEqual("Heading1Char", run.RunProperties!.RunStyle!.Val!.Value);

		run.Style(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void RightToLeft()
	{
		Run run = new Run().RightToLeft(false);

		Assert.IsFalse(run.RightToLeftValue());
		Assert.IsFalse(run.RunProperties!.RightToLeftText!.Val!.Value);

		run.RightToLeft(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Shadow()
	{
		Run run = new Run().Shadow(false);

		Assert.IsFalse(run.ShadowValue());
		Assert.IsFalse(run.RunProperties!.Shadow!.Val!.Value);

		run.Shadow(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Shading()
	{
		Run run = new Run().Shading(new Shading { Val = ShadingPatternValues.DiagonalCross });

		Assert.AreEqual(ShadingPatternValues.DiagonalCross, run.GetShading()!.Val!.Value);
		Assert.AreEqual(ShadingPatternValues.DiagonalCross, run.RunProperties!.Shading!.Val!.Value);

		run.Shading(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void SmallCaps()
	{
		Run run = new Run().SmallCaps(false);

		Assert.IsFalse(run.SmallCapsValue());
		Assert.IsFalse(run.RunProperties!.SmallCaps!.Val!.Value);

		run.SmallCaps(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void SnapToGrid()
	{
		Run run = new Run().SnapToGrid(false);

		Assert.IsFalse(run.SnapToGridValue());
		Assert.IsFalse(run.RunProperties!.SnapToGrid!.Val!.Value);

		run.SnapToGrid(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void CharacterSpacing()
	{
		Run run = new Run().CharacterSpacing(1, MeasuringUnits.Points);
		Assert.AreEqual(1, run.CharacterSpacingValue(MeasuringUnits.Points));
		Assert.AreEqual(20, run.RunProperties!.Spacing!.Val!.Value);

		run.CharacterSpacing(1, MeasuringUnits.Inches);
		Assert.AreEqual(1, run.CharacterSpacingValue(MeasuringUnits.Inches));
		Assert.AreEqual(1440, run.RunProperties!.Spacing!.Val!.Value);

		run.CharacterSpacing(1, MeasuringUnits.Centimeters);
		Assert.AreEqual(1, run.CharacterSpacingValue(MeasuringUnits.Centimeters));
		Assert.AreEqual(567, run.RunProperties!.Spacing!.Val!.Value);

		run.CharacterSpacing(1, MeasuringUnits.Twips);
		Assert.AreEqual(1, run.CharacterSpacingValue(MeasuringUnits.Twips));
		Assert.AreEqual(1, run.RunProperties!.Spacing!.Val!.Value);

		run.CharacterSpacing(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void SpecVanish()
	{
		Run run = new Run().SpecVanish(false);

		Assert.IsFalse(run.SpecVanishValue());
		Assert.IsFalse(run.RunProperties!.SpecVanish!.Val!.Value);

		run.SpecVanish(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Strike()
	{
		Run run = new Run().Strike(false);

		Assert.IsFalse(run.StrikeValue());
		Assert.IsFalse(run.RunProperties!.Strike!.Val!.Value);

		run.Strike(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void FontSize()
	{
		Run run = new Run().FontSize(16, TextMeasuringUnits.HalfPoints);
		Assert.AreEqual(16, run.FontSizeValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual(8, run.FontSizeValue(TextMeasuringUnits.Points));
		Assert.AreEqual("16", run.RunProperties!.FontSize!.Val!.Value);

		run.FontSize(32, TextMeasuringUnits.Points);
		Assert.AreEqual(32, run.FontSizeValue(TextMeasuringUnits.Points));
		Assert.AreEqual(64, run.FontSizeValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual("64", run.RunProperties!.FontSize!.Val!.Value);

		run.FontSize(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void ComplexScriptFontSize()
	{
		Run run = new Run().ComplexScriptFontSize(16, TextMeasuringUnits.HalfPoints);
		Assert.AreEqual(16, run.ComplexScriptFontSizeValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual(8, run.ComplexScriptFontSizeValue(TextMeasuringUnits.Points));
		Assert.AreEqual("16", run.RunProperties!.FontSizeComplexScript!.Val!.Value);

		run.ComplexScriptFontSize(32, TextMeasuringUnits.Points);
		Assert.AreEqual(32, run.ComplexScriptFontSizeValue(TextMeasuringUnits.Points));
		Assert.AreEqual(64, run.ComplexScriptFontSizeValue(TextMeasuringUnits.HalfPoints));
		Assert.AreEqual("64", run.RunProperties!.FontSizeComplexScript!.Val!.Value);

		run.ComplexScriptFontSize(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Underline()
	{
		Run run = new Run().Underline(new Underline { Val = UnderlineValues.Single });

		Assert.AreEqual(UnderlineValues.Single, run.GetUnderline()!.Val!.Value);
		Assert.AreEqual(UnderlineValues.Single, run.RunProperties!.Underline!.Val!.Value);

		run.Underline(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void Hidden()
	{
		Run run = new Run().Hidden(false);

		Assert.IsFalse(run.HiddenValue());
		Assert.IsFalse(run.RunProperties!.Vanish!.Val!.Value);

		run.Hidden(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void VerticalAlignment()
	{
		Run run = new Run().VerticalAlignment(VerticalPositionValues.Superscript);

		Assert.AreEqual(VerticalPositionValues.Superscript, run.VerticalAlignmentValue());
		Assert.AreEqual(VerticalPositionValues.Superscript, run.RunProperties!.VerticalTextAlignment!.Val!.Value);

		run.VerticalAlignment(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void CharacterScale()
	{
		Run run = new Run().CharacterScale(150);

		Assert.AreEqual(150, run.CharacterScaleValue());
		Assert.AreEqual(150, run.RunProperties!.CharacterScale!.Val!.Value);

		run.CharacterScale(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}

	[TestMethod]
	public void WebHidden()
	{
		Run run = new Run().WebHidden(false);

		Assert.IsFalse(run.WebHiddenValue());
		Assert.IsFalse(run.RunProperties!.WebHidden!.Val!.Value);

		run.WebHidden(null); // Removes RunProperties since it's the last element
		Assert.IsNull(run.RunProperties);
	}
}
