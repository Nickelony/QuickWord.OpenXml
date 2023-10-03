using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="Run" /> class.
/// </summary>
public static class RunExtensions
{
	#region Get property methods

	/// <summary>
	/// Specifies whether the bold property shall be applied to all non-complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Bold" /></para>
	/// </summary>
	public static bool? BoldValue(this Run run)
		=> run.RunProperties?.Bold?.Val?.Value;

	/// <summary>
	/// Specifies whether the bold property shall be applied to all complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.BoldComplexScript" /></para>
	/// </summary>
	public static bool? BoldComplexScriptValue(this Run run)
		=> run.RunProperties?.BoldComplexScript?.Val?.Value;

	/// <summary>
	/// Specifies information about the border applied to the text in the current run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Border" /></para>
	/// </summary>
	public static Border? GetBorder(this Run run)
		=> run.RunProperties?.Border;

	/// <summary>
	/// Specifies that any lowercase characters in this text run shall
	/// be formatted for display only as their capital letter character equivalents.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Caps" /></para>
	/// </summary>
	public static bool? AllCapsValue(this Run run)
		=> run.RunProperties?.Caps?.Val?.Value;

	/// <summary>
	/// Specifies the color which shall be used to display
	/// the contents of this run in the document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Color" /></para>
	/// </summary>
	public static Color? GetColor(this Run run)
		=> run.RunProperties?.Color;

	/// <summary>
	/// Specifies whether the contents of this run shall be treated
	/// as complex script text regardless of their Unicode character values when
	/// determining the formatting for this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ComplexScript" /></para>
	/// </summary>
	public static bool? ComplexScriptValue(this Run run)
		=> run.RunProperties?.ComplexScript?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run shall be displayed with
	/// two horizontal lines through each character displayed on the line.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.DoubleStrike" /></para>
	/// </summary>
	public static bool? DoubleStrikeValue(this Run run)
		=> run.RunProperties?.DoubleStrike?.Val?.Value;

	/// <summary>
	/// Specifies any East Asian typography settings which shall
	/// be applied to the contents of the run.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.EastAsianLayout" /></para>
	/// </summary>
	public static EastAsianLayout? GetEastAsianLayout(this Run run)
		=> run.RunProperties?.EastAsianLayout;

	/// <summary>
	/// Specifies an animated text effect which should be displayed when
	/// rendering the contents of this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextEffect" /></para>
	/// </summary>
	public static TextEffectValues? TextEffectValue(this Run run)
		=> run.RunProperties?.TextEffect?.Val?.Value;

	/// <summary>
	/// Specifies the emphasis mark which shall be displayed for each
	/// non-space character in this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Emphasis" /></para>
	/// </summary>
	public static EmphasisMarkValues? EmphasisMarkValue(this Run run)
		=> run.RunProperties?.Emphasis?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run should be displayed as if embossed,
	/// which makes text appear as if it is raised off the page in relief.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Emboss" /></para>
	/// </summary>
	public static bool? EmbossValue(this Run run)
		=> run.RunProperties?.Emboss?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run shall not be automatically displayed based
	/// on the width of its contents, rather its contents shall be resized
	/// to fit the width specified.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FitText" /></para>
	/// </summary>
	public static FitText? GetFitText(this Run run)
		=> run.RunProperties?.FitText;

	/// <summary>
	/// Specifies a highlighting color which is applied as
	/// a background behind the contents of this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Highlight" /></para>
	/// </summary>
	public static HighlightColorValues? HighlightColorValue(this Run run)
		=> run.RunProperties?.Highlight?.Val?.Value;

	/// <summary>
	/// Specifies whether the italic property should be applied to all non-complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Italic" /></para>
	/// </summary>
	public static bool? ItalicValue(this Run run)
		=> run.RunProperties?.Italic?.Val?.Value;

	/// <summary>
	/// Specifies whether the italic property should be applied to all complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ItalicComplexScript" /></para>
	/// </summary>
	public static bool? ItalicComplexScriptValue(this Run run)
		=> run.RunProperties?.ItalicComplexScript?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run should be displayed as if imprinted,
	/// which makes text appear to be imprinted or pressed into page
	/// (also referred to as 'engrave').
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Imprint" /></para>
	/// </summary>
	public static bool? ImprintValue(this Run run)
		=> run.RunProperties?.Imprint?.Val?.Value;

	/// <summary>
	/// Specifies whether font kerning shall be applied to the contents of this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Kern" /></para>
	/// </summary>
	public static double? KerningValue(this Run run, TextMeasuringUnits desiredUnits = TextMeasuringUnits.Points)
	{
		uint? value = run.RunProperties?.Kern?.Val?.Value;
		return value.HasValue ? HalfPoints.ToOther((int)value.Value, desiredUnits) : null;
	}

	/// <summary>
	/// Specifies the languages which shall be used to check
	/// spelling and grammar (if requested) when processing the contents of this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Languages" /></para>
	/// </summary>
	public static Languages? GetLanguages(this Run run)
		=> run.RunProperties?.Languages;

	/// <summary>
	/// Specifies that the contents of this run shall not report
	/// any errors when the document is scanned for spelling and grammar.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NoProof" /></para>
	/// </summary>
	public static bool? NoProofingValue(this Run run)
		=> run.RunProperties?.NoProof?.Val?.Value;

	/// <summary>
	/// Specifies that this run contains WordprocessingML which shall
	/// be handled as though it was Office Open XML Math.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.OfficeMath" /></para>
	/// </summary>
	public static bool? OfficeMathValue(this Run run)
		=> run.RunProperties?.GetFirstChild<OfficeMath>()?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run should be displayed as if they have
	/// an outline, by drawing a one pixel wide border around the inside and outside
	/// borders of each character glyph in the run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Outline" /></para>
	/// </summary>
	public static bool? OutlineValue(this Run run)
		=> run.RunProperties?.Outline?.Val?.Value;

	/// <summary>
	/// Specifies the amount by which text shall be raised or lowered for this run in relation
	/// to the default baseline of the surrounding non-positioned text.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Position" /></para>
	/// </summary>
	public static double? VerticalPositionValue(this Run run, TextMeasuringUnits desiredUnits = TextMeasuringUnits.Points)
	{
		int? value = int.TryParse(run.RunProperties?.Position?.Val, out int result) ? result : null;
		return value.HasValue ? HalfPoints.ToOther(value.Value, desiredUnits) : null;
	}

	/// <summary>
	/// Specifies the fonts which shall be used to display the text contents of this run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RunFonts" /></para>
	/// </summary>
	public static RunFonts? GetFonts(this Run run)
		=> run.RunProperties?.RunFonts;

	/// <summary>
	/// Specifies the style ID of the character style which shall be used to
	/// format the contents of this paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RunStyle" /></para>
	/// </summary>
	public static string? StyleValue(this Run run)
		=> run.RunProperties?.RunStyle?.Val?.Value;

	/// <summary>
	/// Specifies whether the contents of this run shall have right-to-left characteristics.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RightToLeftText" /></para>
	/// </summary>
	public static bool? RightToLeftValue(this Run run)
		=> run.RunProperties?.RightToLeftText?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run shall be displayed as if
	/// each character has a shadow.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shadow" /></para>
	/// </summary>
	public static bool? ShadowValue(this Run run)
		=> run.RunProperties?.Shadow?.Val?.Value;

	/// <summary>
	/// Specifies the shading applied to the contents of the run.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Shading? GetShading(this Run run)
		=> run.RunProperties?.Shading;

	/// <summary>
	/// Specifies that all small letter characters in this text run shall be formatted for display only
	/// as their capital letter character equivalents in a font size two points smaller than the actual
	/// font size specified for this text.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SmallCaps" /></para>
	/// </summary>
	public static bool? SmallCapsValue(this Run run)
		=> run.RunProperties?.SmallCaps?.Val?.Value;

	/// <summary>
	/// Specifies whether the current paragraph should use the document grid lines per
	/// page settings defined in the docGrid element (§17.6.5) when laying out
	/// the contents in the paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SnapToGrid" /></para>
	/// </summary>
	public static bool? SnapToGridValue(this Run run)
		=> run.RunProperties?.SnapToGrid?.Val?.Value;

	/// <summary>
	/// Specifies the amount of character pitch which shall be added or removed after
	/// each character in this run before the following character is rendered in the document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Spacing" /></para>
	/// </summary>
	public static double? CharacterSpacingValue(this Run run, MeasuringUnits desiredUnits = MeasuringUnits.Points)
	{
		int? value = run.RunProperties?.Spacing?.Val?.Value;
		return value.HasValue ? Twips.ToOther(value.Value, desiredUnits) : null;
	}

	/// <summary>
	/// Specifies that the given run shall always behave as if it is hidden,
	/// even when hidden text is being displayed in the current document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpecVanish" /></para>
	/// </summary>
	public static bool? SpecVanishValue(this Run run)
		=> run.RunProperties?.SpecVanish?.Val?.Value;

	/// <summary>
	/// Specifies that the contents of this run shall be displayed
	/// with a single horizontal line through the center of the line.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Strike" /></para>
	/// </summary>
	public static bool? StrikeValue(this Run run)
		=> run.RunProperties?.Strike?.Val?.Value;

	/// <summary>
	/// Specifies the font size which shall be applied to all non complex
	/// script characters in the contents of this run when displayed.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FontSize" /></para>
	/// </summary>
	public static double? FontSizeValue(this Run run, TextMeasuringUnits desiredUnits = TextMeasuringUnits.Points)
	{
		int? value = int.TryParse(run.RunProperties?.FontSize?.Val, out int result) ? result : null;
		return value.HasValue ? HalfPoints.ToOther(value.Value, desiredUnits) : null;
	}

	/// <summary>
	/// Specifies the font size which shall be applied to all complex
	/// script characters in the contents of this run when displayed.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FontSizeComplexScript" /></para>
	/// </summary>
	public static double? ComplexScriptFontSizeValue(this Run run, TextMeasuringUnits desiredUnits = TextMeasuringUnits.Points)
	{
		int? value = int.TryParse(run.RunProperties?.FontSizeComplexScript?.Val, out int result) ? result : null;
		return value.HasValue ? HalfPoints.ToOther(value.Value, desiredUnits) : null;
	}

	/// <summary>
	/// Specifies that the contents of this run should be displayed along with
	/// an underline appearing directly below the character height
	/// (less all spacing above and below the characters on the line).
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Underline" /></para>
	/// </summary>
	public static Underline? GetUnderline(this Run run)
		=> run.RunProperties?.Underline;

	/// <summary>
	/// Specifies whether the contents of this run shall be hidden from
	/// display at display time in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Vanish" /></para>
	/// </summary>
	public static bool? HiddenValue(this Run run)
		=> run.RunProperties?.Vanish?.Val?.Value;

	/// <summary>
	/// Specifies the alignment which shall be applied to the contents of this run in relation
	/// to the default appearance of the run's text.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignment" /></para>
	/// </summary>
	public static VerticalPositionValues? VerticalAlignmentValue(this Run run)
		=> run.RunProperties?.VerticalTextAlignment?.Val?.Value;

	/// <summary>
	/// Specifies the amount by which each character shall be stretched or compressed.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.CharacterScale" /></para>
	/// </summary>
	public static long? CharacterScaleValue(this Run run)
		=> run.RunProperties?.CharacterScale?.Val?.Value;

	/// <summary>
	/// Specifies whether the contents of this run shall be hidden from
	/// display at display time in a document when the document is being
	/// displayed in a web page view.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WebHidden" /></para>
	/// </summary>
	public static bool? WebHiddenValue(this Run run)
		=> run.RunProperties?.WebHidden?.Val?.Value;

	#endregion Get property methods

	#region Set property methods

	/// <summary>
	/// Specifies whether the bold property shall be applied to all non-complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Bold" /></para>
	/// </summary>
	public static Run Bold(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Bold>(value);
		return run;
	}

	/// <summary>
	/// Specifies whether the bold property shall be applied to all complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.BoldComplexScript" /></para>
	/// </summary>
	public static Run BoldComplexScript(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<BoldComplexScript>(value);
		return run;
	}

	/// <summary>
	/// Specifies information about the border applied to the text in the current run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Border" /></para>
	/// </summary>
	public static Run Border(this Run run, Border? border)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(border);
		return run;
	}

	/// <summary>
	/// Specifies that any lowercase characters in this text run shall
	/// be formatted for display only as their capital letter character equivalents.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Caps" /></para>
	/// </summary>
	public static Run AllCaps(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Caps>(value);
		return run;
	}

	/// <summary>
	/// Specifies the color which shall be used to display
	/// the contents of this run in the document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Color" /></para>
	/// </summary>
	public static Run Color(this Run run, Color? color)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(color);
		return run;
	}

	/// <summary>
	/// Specifies whether the contents of this run shall be treated
	/// as complex script text regardless of their Unicode character values when
	/// determining the formatting for this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ComplexScript" /></para>
	/// </summary>
	public static Run ComplexScript(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<ComplexScript>(value);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run shall be displayed with
	/// two horizontal lines through each character displayed on the line.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.DoubleStrike" /></para>
	/// </summary>
	public static Run DoubleStrike(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<DoubleStrike>(value);
		return run;
	}

	/// <summary>
	/// Specifies any East Asian typography settings which shall
	/// be applied to the contents of the run.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.EastAsianLayout" /></para>
	/// </summary>
	public static Run EastAsianLayout(this Run run, EastAsianLayout? layout)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(layout);
		return run;
	}

	/// <summary>
	/// Specifies an animated text effect which should be displayed when
	/// rendering the contents of this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextEffect" /></para>
	/// </summary>
	public static Run TextEffect(this Run run, TextEffectValues? effect)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<TextEffect>(effect);
		return run;
	}

	/// <summary>
	/// Specifies the emphasis mark which shall be displayed for each
	/// non-space character in this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Emphasis" /></para>
	/// </summary>
	public static Run EmphasisMark(this Run run, EmphasisMarkValues? emphasis)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Emphasis>(emphasis);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run should be displayed as if embossed,
	/// which makes text appear as if it is raised off the page in relief.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Emboss" /></para>
	/// </summary>
	public static Run Emboss(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Emboss>(value);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run shall not be automatically displayed based
	/// on the width of its contents, rather its contents shall be resized
	/// to fit the width specified.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FitText" /></para>
	/// </summary>
	public static Run FitText(this Run run, FitText? fitText)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(fitText);
		return run;
	}

	/// <summary>
	/// Specifies a highlighting color which is applied as
	/// a background behind the contents of this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Highlight" /></para>
	/// </summary>
	public static Run HighlightColor(this Run run, HighlightColorValues? color)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Highlight>(color);
		return run;
	}

	/// <summary>
	/// Specifies whether the italic property should be applied to all non-complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Italic" /></para>
	/// </summary>
	public static Run Italic(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Italic>(value);
		return run;
	}

	/// <summary>
	/// Specifies whether the italic property should be applied to all complex
	/// script characters in the contents of this run when displayed in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ItalicComplexScript" /></para>
	/// </summary>
	public static Run ItalicComplexScript(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<ItalicComplexScript>(value);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run should be displayed as if imprinted,
	/// which makes text appear to be imprinted or pressed into page
	/// (also referred to as 'engrave').
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Imprint" /></para>
	/// </summary>
	public static Run Imprint(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Imprint>(value);
		return run;
	}

	/// <summary>
	/// Specifies whether font kerning shall be applied to the contents of this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Kern" /></para>
	/// </summary>
	public static Run Kerning(this Run run, double? size, TextMeasuringUnits units = TextMeasuringUnits.Points)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Kern>(
			size.HasValue ? (uint)HalfPoints.FromOther(size.Value, units) : null);

		return run;
	}

	/// <summary>
	/// Specifies the languages which shall be used to check
	/// spelling and grammar (if requested) when processing the contents of this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Languages" /></para>
	/// </summary>
	public static Run Languages(this Run run, Languages? languages)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(languages);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run shall not report
	/// any errors when the document is scanned for spelling and grammar.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NoProof" /></para>
	/// </summary>
	public static Run NoProofing(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<NoProof>(value);
		return run;
	}

	/// <summary>
	/// Specifies that this run contains WordprocessingML which shall
	/// be handled as though it was Office Open XML Math.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.OfficeMath" /></para>
	/// </summary>
	public static Run OfficeMath(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<OfficeMath>(value);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run should be displayed as if they have
	/// an outline, by drawing a one pixel wide border around the inside and outside
	/// borders of each character glyph in the run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Outline" /></para>
	/// </summary>
	public static Run Outline(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Outline>(value);
		return run;
	}

	/// <summary>
	/// Specifies the amount by which text shall be raised or lowered for this run in relation
	/// to the default baseline of the surrounding non-positioned text.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Position" /></para>
	/// </summary>
	public static Run VerticalPosition(this Run run, double? shift, TextMeasuringUnits units = TextMeasuringUnits.Points)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Position>(
			shift.HasValue ? HalfPoints.FromOther(shift.Value, units).ToString() : null);

		return run;
	}

	/// <summary>
	/// Specifies the fonts which shall be used to display the text contents of this run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RunFonts" /></para>
	/// </summary>
	public static Run Fonts(this Run run, RunFonts? fonts)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(fonts);
		return run;
	}

	/// <summary>
	/// Specifies the style ID of the character style which shall be used to
	/// format the contents of this paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RunStyle" /></para>
	/// </summary>
	public static Run Style(this Run run, string? styleId)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<RunStyle>(styleId);
		return run;
	}

	/// <summary>
	/// Specifies whether the contents of this run shall have right-to-left characteristics.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RightToLeftText" /></para>
	/// </summary>
	public static Run RightToLeft(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<RightToLeftText>(value);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run shall be displayed as if
	/// each character has a shadow.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shadow" /></para>
	/// </summary>
	public static Run Shadow(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Shadow>(value);
		return run;
	}

	/// <summary>
	/// Specifies the shading applied to the contents of the run.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Run Shading(this Run run, Shading? shading)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(shading);
		return run;
	}

	/// <summary>
	/// Specifies that all small letter characters in this text run shall be formatted for display only
	/// as their capital letter character equivalents in a font size two points smaller than the actual
	/// font size specified for this text.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SmallCaps" /></para>
	/// </summary>
	public static Run SmallCaps(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<SmallCaps>(value);
		return run;
	}

	/// <summary>
	/// Specifies whether the current paragraph should use the document grid lines per
	/// page settings defined in the docGrid element (§17.6.5) when laying out
	/// the contents in the paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SnapToGrid" /></para>
	/// </summary>
	public static Run SnapToGrid(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<SnapToGrid>(value);
		return run;
	}

	/// <summary>
	/// Specifies the amount of character pitch which shall be added or removed after
	/// each character in this run before the following character is rendered in the document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Spacing" /></para>
	/// </summary>
	public static Run CharacterSpacing(this Run run, double? size, MeasuringUnits units = MeasuringUnits.Points)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Spacing>(
			size.HasValue ? Twips.FromOther(size.Value, units) : null);

		return run;
	}

	/// <summary>
	/// Specifies that the given run shall always behave as if it is hidden,
	/// even when hidden text is being displayed in the current document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpecVanish" /></para>
	/// </summary>
	public static Run SpecVanish(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<SpecVanish>(value);
		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run shall be displayed
	/// with a single horizontal line through the center of the line.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Strike" /></para>
	/// </summary>
	public static Run Strike(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Strike>(value);
		return run;
	}

	/// <summary>
	/// Specifies the font size which shall be applied to all non complex
	/// script characters in the contents of this run when displayed.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FontSize" /></para>
	/// </summary>
	public static Run FontSize(this Run run, double? size, TextMeasuringUnits units = TextMeasuringUnits.Points)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<FontSize>(
			size.HasValue ? HalfPoints.FromOther(size.Value, units).ToString() : null);

		return run;
	}

	/// <summary>
	/// Specifies the font size which shall be applied to all complex
	/// script characters in the contents of this run when displayed.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FontSizeComplexScript" /></para>
	/// </summary>
	public static Run ComplexScriptFontSize(this Run run, double? size, TextMeasuringUnits units = TextMeasuringUnits.Points)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<FontSizeComplexScript>(
			size.HasValue ? HalfPoints.FromOther(size.Value, units).ToString() : null);

		return run;
	}

	/// <summary>
	/// Specifies that the contents of this run should be displayed along with
	/// an underline appearing directly below the character height
	/// (less all spacing above and below the characters on the line).
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Underline" /></para>
	/// </summary>
	public static Run Underline(this Run run, Underline? underline)
	{
		run.GetOrInit<RunProperties>().SetPropertyClassOrRemove(underline);
		return run;
	}

	/// <summary>
	/// Specifies whether the contents of this run shall be hidden from
	/// display at display time in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Vanish" /></para>
	/// </summary>
	public static Run Hidden(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<Vanish>(value);
		return run;
	}

	/// <summary>
	/// Specifies the alignment which shall be applied to the contents of this run in relation
	/// to the default appearance of the run's text.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.VerticalTextAlignment" /></para>
	/// </summary>
	public static Run VerticalAlignment(this Run run, VerticalPositionValues? position)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<VerticalTextAlignment>(position);
		return run;
	}

	/// <summary>
	/// Specifies the amount by which each character shall be stretched or compressed.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.CharacterScale" /></para>
	/// </summary>
	public static Run CharacterScale(this Run run, long? percentage)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<CharacterScale>(percentage);
		return run;
	}

	/// <summary>
	/// Specifies whether the contents of this run shall be hidden from
	/// display at display time in a document when the document is being
	/// displayed in a web page view.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WebHidden" /></para>
	/// </summary>
	public static Run WebHidden(this Run run, bool? value = true)
	{
		run.GetOrInit<RunProperties>().SetValOrRemove<WebHidden>(value);
		return run;
	}

	#endregion Set property methods
}
