using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using System.Text;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml.Extras;

public static class RunExtraExtensions
{
	#region Formatting

	/// <summary>
	/// Clones all formatting properties from the run.
	/// </summary>
	public static RunFormatting CloneFormatting(this Run run) => new()
	{
		Bold = run.BoldValue(),
		BoldComplexScript = run.BoldComplexScriptValue(),
		Border = run.GetBorder()?.CloneNode(true) as Border,
		AllCaps = run.AllCapsValue(),
		Color = run.GetColor()?.CloneNode(true) as Color,
		ComplexScript = run.ComplexScriptValue(),
		DoubleStrike = run.DoubleStrikeValue(),
		EastAsianLayout = run.GetEastAsianLayout()?.CloneNode(true) as EastAsianLayout,
		TextEffect = run.TextEffectValue(),
		EmphasisMark = run.EmphasisMarkValue(),
		Emboss = run.EmbossValue(),
		FitText = run.GetFitText()?.CloneNode(true) as FitText,
		HighlightColor = run.HighlightColorValue(),
		Italic = run.ItalicValue(),
		ItalicComplexScript = run.ItalicComplexScriptValue(),
		Imprint = run.ImprintValue(),
		Kerning = run.KerningValue(TextMeasuringUnits.Points),
		Languages = run.GetLanguages()?.CloneNode(true) as Languages,
		NoProofing = run.NoProofingValue(),
		OfficeMath = run.OfficeMathValue(),
		Outline = run.OutlineValue(),
		VerticalPosition = run.VerticalPositionValue(TextMeasuringUnits.Points),
		Fonts = run.GetFonts()?.CloneNode(true) as RunFonts,
		Style = run.StyleValue(),
		RightToLeft = run.RightToLeftValue(),
		Shadow = run.ShadowValue(),
		Shading = run.GetShading()?.CloneNode(true) as Shading,
		SmallCaps = run.SmallCapsValue(),
		SnapToGrid = run.SnapToGridValue(),
		CharacterSpacing = run.CharacterSpacingValue(MeasuringUnits.Points),
		SpecVanish = run.SpecVanishValue(),
		Strike = run.StrikeValue(),
		FontSize = run.FontSizeValue(),
		ComplexScriptFontSize = run.ComplexScriptFontSizeValue(),
		Underline = run.GetUnderline()?.CloneNode(true) as Underline,
		Hidden = run.HiddenValue(),
		VerticalAlignment = run.VerticalAlignmentValue(),
		CharacterScale = run.CharacterScaleValue(),
		WebHidden = run.WebHiddenValue()
	};

	/// <summary>
	/// Applies the given formatting properties to the run (replaces every possible property of the run unless <c>ignoreNulls</c> is set to true).
	/// </summary>
	public static Run ApplyFormatting(this Run run, RunFormatting formatting, bool ignoreNulls = false)
	{
		if (formatting.Bold is not null || (formatting.Bold is null && !ignoreNulls))
			run.Bold(formatting.Bold);

		if (formatting.BoldComplexScript is not null || (formatting.BoldComplexScript is null && !ignoreNulls))
			run.BoldComplexScript(formatting.BoldComplexScript);

		if (formatting.Border is not null || (formatting.Border is null && !ignoreNulls))
			run.Border(formatting.Border?.CloneNode(true) as Border);

		if (formatting.AllCaps is not null || (formatting.AllCaps is null && !ignoreNulls))
			run.AllCaps(formatting.AllCaps);

		if (formatting.Color is not null || (formatting.Color is null && !ignoreNulls))
			run.Color(formatting.Color?.CloneNode(true) as Color);

		if (formatting.ComplexScript is not null || (formatting.ComplexScript is null && !ignoreNulls))
			run.ComplexScript(formatting.ComplexScript);

		if (formatting.DoubleStrike is not null || (formatting.DoubleStrike is null && !ignoreNulls))
			run.DoubleStrike(formatting.DoubleStrike);

		if (formatting.EastAsianLayout is not null || (formatting.EastAsianLayout is null && !ignoreNulls))
			run.EastAsianLayout(formatting.EastAsianLayout?.CloneNode(true) as EastAsianLayout);

		if (formatting.TextEffect is not null || (formatting.TextEffect is null && !ignoreNulls))
			run.TextEffect(formatting.TextEffect);

		if (formatting.EmphasisMark is not null || (formatting.EmphasisMark is null && !ignoreNulls))
			run.EmphasisMark(formatting.EmphasisMark);

		if (formatting.Emboss is not null || (formatting.Emboss is null && !ignoreNulls))
			run.Emboss(formatting.Emboss);

		if (formatting.FitText is not null || (formatting.FitText is null && !ignoreNulls))
			run.FitText(formatting.FitText?.CloneNode(true) as FitText);

		if (formatting.HighlightColor is not null || (formatting.HighlightColor is null && !ignoreNulls))
			run.HighlightColor(formatting.HighlightColor);

		if (formatting.Italic is not null || (formatting.Italic is null && !ignoreNulls))
			run.Italic(formatting.Italic);

		if (formatting.ItalicComplexScript is not null || (formatting.ItalicComplexScript is null && !ignoreNulls))
			run.ItalicComplexScript(formatting.ItalicComplexScript);

		if (formatting.Imprint is not null || (formatting.Imprint is null && !ignoreNulls))
			run.Imprint(formatting.Imprint);

		if (formatting.Kerning is not null || (formatting.Kerning is null && !ignoreNulls))
			run.Kerning(formatting.Kerning, TextMeasuringUnits.Points);

		if (formatting.Languages is not null || (formatting.Languages is null && !ignoreNulls))
			run.Languages(formatting.Languages?.CloneNode(true) as Languages);

		if (formatting.NoProofing is not null || (formatting.NoProofing is null && !ignoreNulls))
			run.NoProofing(formatting.NoProofing);

		if (formatting.OfficeMath is not null || (formatting.OfficeMath is null && !ignoreNulls))
			run.OfficeMath(formatting.OfficeMath);

		if (formatting.Outline is not null || (formatting.Outline is null && !ignoreNulls))
			run.Outline(formatting.Outline);

		if (formatting.VerticalPosition is not null || (formatting.VerticalPosition is null && !ignoreNulls))
			run.VerticalPosition(formatting.VerticalPosition, TextMeasuringUnits.Points);

		if (formatting.Fonts is not null || (formatting.Fonts is null && !ignoreNulls))
			run.Fonts(formatting.Fonts?.CloneNode(true) as RunFonts);

		if (formatting.Style is not null || (formatting.Style is null && !ignoreNulls))
			run.Style(formatting.Style);

		if (formatting.RightToLeft is not null || (formatting.RightToLeft is null && !ignoreNulls))
			run.RightToLeft(formatting.RightToLeft);

		if (formatting.Shadow is not null || (formatting.Shadow is null && !ignoreNulls))
			run.Shadow(formatting.Shadow);

		if (formatting.Shading is not null || (formatting.Shading is null && !ignoreNulls))
			run.Shading(formatting.Shading?.CloneNode(true) as Shading);

		if (formatting.SmallCaps is not null || (formatting.SmallCaps is null && !ignoreNulls))
			run.SmallCaps(formatting.SmallCaps);

		if (formatting.SnapToGrid is not null || (formatting.SnapToGrid is null && !ignoreNulls))
			run.SnapToGrid(formatting.SnapToGrid);

		if (formatting.CharacterSpacing is not null || (formatting.CharacterSpacing is null && !ignoreNulls))
			run.CharacterSpacing(formatting.CharacterSpacing, MeasuringUnits.Points);

		if (formatting.SpecVanish is not null || (formatting.SpecVanish is null && !ignoreNulls))
			run.SpecVanish(formatting.SpecVanish);

		if (formatting.Strike is not null || (formatting.Strike is null && !ignoreNulls))
			run.Strike(formatting.Strike);

		if (formatting.FontSize is not null || (formatting.FontSize is null && !ignoreNulls))
			run.FontSize(formatting.FontSize);

		if (formatting.ComplexScriptFontSize is not null || (formatting.ComplexScriptFontSize is null && !ignoreNulls))
			run.ComplexScriptFontSize(formatting.ComplexScriptFontSize);

		if (formatting.Underline is not null || (formatting.Underline is null && !ignoreNulls))
			run.Underline(formatting.Underline?.CloneNode(true) as Underline);

		if (formatting.Hidden is not null || (formatting.Hidden is null && !ignoreNulls))
			run.Hidden(formatting.Hidden);

		if (formatting.VerticalAlignment is not null || (formatting.VerticalAlignment is null && !ignoreNulls))
			run.VerticalAlignment(formatting.VerticalAlignment);

		if (formatting.CharacterScale is not null || (formatting.CharacterScale is null && !ignoreNulls))
			run.CharacterScale(formatting.CharacterScale);

		if (formatting.WebHidden is not null || (formatting.WebHidden is null && !ignoreNulls))
			run.WebHidden(formatting.WebHidden);

		return run;
	}

	/// <summary>
	/// Resets every possible property of the run.
	/// </summary>
	public static Run ResetFormatting(this Run run)
	{
		run.RemoveAllChildren<RunProperties>();
		return run;
	}

	#endregion Formatting

	#region Get methods

	/// <summary>
	/// Gets the text of the run, while converting each <see cref="Break" /> into a <c><see cref="Environment.NewLine"/></c>.
	/// </summary>
	public static string Text(this Run run)
	{
		var builder = new StringBuilder();

		foreach (OpenXmlElement child in run.ChildElements.Where<OpenXmlElement>(c => c is DocumentFormat.OpenXml.Wordprocessing.Text or Break))
		{
			if (child is Text textElement)
				builder.Append(textElement.Text);
			else if (child is Break)
				builder.Append(Environment.NewLine);
		}

		return builder.ToString();
	}

	/// <summary>
	/// Specifies the width that the run shall be scaled to fit into.
	/// <para>Property of <see cref="FitText" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FitText" /></para>
	/// </summary>
	public static double? ManualWidthValue(this Run run, MeasuringUnits desiredUnits)
	{
		uint? value = run.RunProperties?.FitText?.Val?.Value;
		return value.HasValue ? Twips.ToOther((int)value.Value, desiredUnits) : null;
	}

	#endregion Get methods

	#region Set methods

	/// <summary>
	/// Sets the text of the run.
	/// </summary>
	/// <param name="parseNewLineChars">
	///	Specifies whether <c><see cref="Environment.NewLine"/></c> chars should be parsed into <see cref="Break"/> objects.
	///	</param>
	public static Run Text(this Run run, string? text, bool parseNewLineChars = true)
	{
		run.RemoveAllChildren<Text>();
		run.RemoveAllChildren<Break>();

		if (text is not null)
			run.AppendText(text, parseNewLineChars);

		return run;
	}

	/// <summary>
	/// Sets the border of the run.
	/// <para><c>Size</c>, <c>Val</c> and <c>Color</c> properties of <see cref="DocumentFormat.OpenXml.Wordprocessing.Border" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Border" /></para>
	/// </summary>
	public static Run Border(this Run run, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		run.GetOrInit<RunProperties>().Border = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing,
		};

		return run;
	}

	/// <summary>
	/// Sets the background fill color of the run.
	/// <para><c>Fill</c> property of <see cref="Shading" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Run FillColor(this Run run, string htmlColor)
	{
		Shading shading = run.GetOrInit<RunProperties>().GetOrInit<Shading>();
		shading.Fill = htmlColor;
		shading.Val = shading.Val?.Value is null or ShadingPatternValues.Nil
			? ShadingPatternValues.Clear
			: shading.Val.Value; // Preserve original pattern

		return run;
	}

	/// <summary>
	/// Sets the color of the font.
	/// <para><c>Val</c> property of <see cref="Color" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Color" /></para>
	/// </summary>
	public static Run FontColor(this Run run, string htmlColor)
	{
		run.GetOrInit<RunProperties>().GetOrInit<Color>().Val = htmlColor;
		return run;
	}

	/// <summary>
	/// Sets the font face of the ASCII characters.
	/// <para><c>Ascii</c> property of <see cref="RunFonts" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.RunFonts" /></para>
	/// </summary>
	/// <param name="faceName">Font face, such as "Comic Sans MS" or "Times New Roman".</param>
	public static Run FontFace(this Run run, string faceName)
	{
		run.GetOrInit<RunProperties>().GetOrInit<RunFonts>().Ascii = faceName;
		return run;
	}

	/// <summary>
	/// Sets the language of the run.
	/// <para><c>Val</c> property of <see cref="Languages" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Languages" /></para>
	/// </summary>
	/// <param name="latinLanguageCode">Latin language code, such as "en-GB", "en-US" or "de-DE".</param>
	public static Run Language(this Run run, string latinLanguageCode)
	{
		run.GetOrInit<RunProperties>().GetOrInit<Languages>().Val = latinLanguageCode;
		return run;
	}

	/// <inheritdoc cref="ManualWidthValue" />
	public static Run ManualWidth(this Run run, double width, MeasuringUnits units = MeasuringUnits.Points)
	{
		run.GetOrInit<RunProperties>().GetOrInit<FitText>().Val = (uint)Twips.FromOther(width, units);
		return run;
	}

	/// <inheritdoc cref="RunExtensions.Underline" />
	public static Run Underline(this Run run, UnderlineValues style = UnderlineValues.Single, string htmlColor = "auto")
	{
		Underline underline = run.GetOrInit<RunProperties>().GetOrInit<Underline>();
		underline.Val = style;
		underline.Color = htmlColor;

		return run;
	}

	#endregion Set methods

	#region Additional methods

	/// <summary>
	/// Appends text at the end of the run.
	/// </summary>
	/// <param name="parseNewLineChars">
	///	Specifies whether <c><see cref="Environment.NewLine"/></c> chars should be parsed into <see cref="Break"/> objects.
	///	</param>
	public static Run AppendText(this Run run, string text, bool parseNewLineChars = true)
	{
		if (parseNewLineChars)
		{
			string[] textArray = text.Replace("\r", "").Split('\n');
			bool first = true;

			foreach (string line in textArray)
			{
				if (!first)
					run.AppendChild(new Break());

				run.AppendChild(new Text(line) { Space = SpaceProcessingModeValues.Preserve });
				first = false;
			}
		}
		else
			run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });

		return run;
	}

	#endregion Additional methods
}
