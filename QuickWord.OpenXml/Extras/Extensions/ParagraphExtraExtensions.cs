using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace QuickWord.OpenXml.Extras;

/// <summary>
/// Additional extension methods for the <see cref="Paragraph"/> class.
/// </summary>
public static class ParagraphExtraExtensions
{
	public static IEnumerable<Run> Runs(this Paragraph paragraph)
		=> paragraph.Elements<Run>();

	public static Run? Runs(this Paragraph paragraph, int index)
		=> paragraph.Elements<Run>().ElementAtOrDefault(index);

	#region Formatting

	/// <summary>
	/// Clones all formatting properties from the paragraph. This does not include any Run formatting.
	/// </summary>
	public static ParagraphFormatting CloneFormatting(this Paragraph paragraph) => new()
	{
		AdjustRightIndent = paragraph.AdjustRightIndentValue(),
		AutoSpaceDE = paragraph.AutoSpaceDEValue(),
		AutoSpaceDN = paragraph.AutoSpaceDNValue(),
		BiDirectional = paragraph.BiDirectionalValue(),
		ConditionalFormatStyle = paragraph.GetConditionalFormatStyle()?.CloneNode(true) as ConditionalFormatStyle,
		ContextualSpacing = paragraph.ContextualSpacingValue(),
		DivId = paragraph.DivIdValue(),
		FrameProperties = paragraph.GetFrameProperties()?.CloneNode(true) as FrameProperties,
		Indentation = paragraph.GetIndentation()?.CloneNode(true) as Indentation,
		Justification = paragraph.JustificationValue(),
		KeepLinesTogether = paragraph.KeepLinesTogetherValue(),
		KeepWithNext = paragraph.KeepWithNextValue(),
		Kinsoku = paragraph.KinsokuValue(),
		MirrorIndents = paragraph.MirrorIndentsValue(),
		NumberingProperties = paragraph.GetNumberingProperties()?.CloneNode(true) as NumberingProperties,
		OutlineLevel = paragraph.OutlineLevelValue(),
		OverflowPunctuation = paragraph.OverflowPunctuationValue(),
		PageBreakBefore = paragraph.PageBreakBeforeValue(),
		Borders = paragraph.GetBorders()?.CloneNode(true) as ParagraphBorders,
		Style = paragraph.StyleValue(),
		MarkRunProperties = paragraph.GetMarkRunProperties()?.CloneNode(true) as ParagraphMarkRunProperties,
		SectionProperties = paragraph.GetSectionProperties()?.CloneNode(true) as SectionProperties,
		Shading = paragraph.GetShading()?.CloneNode(true) as Shading,
		SnapToGrid = paragraph.SnapToGridValue(),
		Spacing = paragraph.GetSpacing()?.CloneNode(true) as SpacingBetweenLines,
		SuppressAutoHyphenation = paragraph.SuppressAutoHyphenationValue(),
		SuppressLineNumbers = paragraph.SuppressLineNumbersValue(),
		SuppressOverlapping = paragraph.SuppressOverlappingValue(),
		Tabs = paragraph.GetTabs()?.CloneNode(true) as Tabs,
		VerticalTextAlignment = paragraph.VerticalTextAlignmentValue(),
		TextBoxTightWrap = paragraph.TextBoxTightWrapValue(),
		TextDirection = paragraph.TextDirectionValue(),
		TopLinePunctuation = paragraph.TopLinePunctuationValue(),
		WidowControl = paragraph.WidowControlValue(),
		WordWrap = paragraph.WordWrapValue()
	};

	/// <summary>
	/// Applies the given formatting properties to the paragraph (replaces every possible property of the paragraph unless <c>ignoreNulls</c> is set to true).
	/// </summary>
	public static Paragraph ApplyFormatting(this Paragraph paragraph, ParagraphFormatting formatting, bool ignoreNulls = false)
	{
		if (formatting.AdjustRightIndent is not null || (formatting.AdjustRightIndent is null && !ignoreNulls))
			paragraph.AdjustRightIndent(formatting.AdjustRightIndent);

		if (formatting.AutoSpaceDE is not null || (formatting.AutoSpaceDE is null && !ignoreNulls))
			paragraph.AutoSpaceDE(formatting.AutoSpaceDE);

		if (formatting.AutoSpaceDN is not null || (formatting.AutoSpaceDN is null && !ignoreNulls))
			paragraph.AutoSpaceDN(formatting.AutoSpaceDN);

		if (formatting.BiDirectional is not null || (formatting.BiDirectional is null && !ignoreNulls))
			paragraph.BiDirectional(formatting.BiDirectional);

		if (formatting.ConditionalFormatStyle is not null || (formatting.ConditionalFormatStyle is null && !ignoreNulls))
			paragraph.ConditionalFormatStyle(formatting.ConditionalFormatStyle?.CloneNode(true) as ConditionalFormatStyle);

		if (formatting.ContextualSpacing is not null || (formatting.ContextualSpacing is null && !ignoreNulls))
			paragraph.ContextualSpacing(formatting.ContextualSpacing);

		if (formatting.DivId is not null || (formatting.DivId is null && !ignoreNulls))
			paragraph.DivId(formatting.DivId);

		if (formatting.FrameProperties is not null || (formatting.FrameProperties is null && !ignoreNulls))
			paragraph.FrameProperties(formatting.FrameProperties?.CloneNode(true) as FrameProperties);

		if (formatting.Indentation is not null || (formatting.Indentation is null && !ignoreNulls))
			paragraph.Indentation(formatting.Indentation?.CloneNode(true) as Indentation);

		if (formatting.Justification is not null || (formatting.Justification is null && !ignoreNulls))
			paragraph.Justification(formatting.Justification);

		if (formatting.KeepLinesTogether is not null || (formatting.KeepLinesTogether is null && !ignoreNulls))
			paragraph.KeepLinesTogether(formatting.KeepLinesTogether);

		if (formatting.KeepWithNext is not null || (formatting.KeepWithNext is null && !ignoreNulls))
			paragraph.KeepWithNext(formatting.KeepWithNext);

		if (formatting.Kinsoku is not null || (formatting.Kinsoku is null && !ignoreNulls))
			paragraph.Kinsoku(formatting.Kinsoku);

		if (formatting.MirrorIndents is not null || (formatting.MirrorIndents is null && !ignoreNulls))
			paragraph.MirrorIndents(formatting.MirrorIndents);

		if (formatting.NumberingProperties is not null || (formatting.NumberingProperties is null && !ignoreNulls))
			paragraph.NumberingProperties(formatting.NumberingProperties?.CloneNode(true) as NumberingProperties);

		if (formatting.OutlineLevel is not null || (formatting.OutlineLevel is null && !ignoreNulls))
			paragraph.OutlineLevel(formatting.OutlineLevel);

		if (formatting.OverflowPunctuation is not null || (formatting.OverflowPunctuation is null && !ignoreNulls))
			paragraph.OverflowPunctuation(formatting.OverflowPunctuation);

		if (formatting.PageBreakBefore is not null || (formatting.PageBreakBefore is null && !ignoreNulls))
			paragraph.PageBreakBefore(formatting.PageBreakBefore);

		if (formatting.Borders is not null || (formatting.Borders is null && !ignoreNulls))
			paragraph.Borders(formatting.Borders?.CloneNode(true) as ParagraphBorders);

		if (formatting.Style is not null || (formatting.Style is null && !ignoreNulls))
			paragraph.Style(formatting.Style);

		if (formatting.MarkRunProperties is not null || (formatting.MarkRunProperties is null && !ignoreNulls))
			paragraph.MarkRunProperties(formatting.MarkRunProperties?.CloneNode(true) as ParagraphMarkRunProperties);

		if (formatting.SectionProperties is not null || (formatting.SectionProperties is null && !ignoreNulls))
			paragraph.SectionProperties(formatting.SectionProperties?.CloneNode(true) as SectionProperties);

		if (formatting.Shading is not null || (formatting.Shading is null && !ignoreNulls))
			paragraph.Shading(formatting.Shading?.CloneNode(true) as Shading);

		if (formatting.SnapToGrid is not null || (formatting.SnapToGrid is null && !ignoreNulls))
			paragraph.SnapToGrid(formatting.SnapToGrid);

		if (formatting.Spacing is not null || (formatting.Spacing is null && !ignoreNulls))
			paragraph.Spacing(formatting.Spacing?.CloneNode(true) as SpacingBetweenLines);

		if (formatting.SuppressAutoHyphenation is not null || (formatting.SuppressAutoHyphenation is null && !ignoreNulls))
			paragraph.SuppressAutoHyphenation(formatting.SuppressAutoHyphenation);

		if (formatting.SuppressLineNumbers is not null || (formatting.SuppressLineNumbers is null && !ignoreNulls))
			paragraph.SuppressLineNumbers(formatting.SuppressLineNumbers);

		if (formatting.SuppressOverlapping is not null || (formatting.SuppressOverlapping is null && !ignoreNulls))
			paragraph.SuppressOverlapping(formatting.SuppressOverlapping);

		if (formatting.Tabs is not null || (formatting.Tabs is null && !ignoreNulls))
			paragraph.Tabs(formatting.Tabs?.CloneNode(true) as Tabs);

		if (formatting.VerticalTextAlignment is not null || (formatting.VerticalTextAlignment is null && !ignoreNulls))
			paragraph.VerticalTextAlignment(formatting.VerticalTextAlignment);

		if (formatting.TextBoxTightWrap is not null || (formatting.TextBoxTightWrap is null && !ignoreNulls))
			paragraph.TextBoxTightWrap(formatting.TextBoxTightWrap);

		if (formatting.TextDirection is not null || (formatting.TextDirection is null && !ignoreNulls))
			paragraph.TextDirection(formatting.TextDirection);

		if (formatting.TopLinePunctuation is not null || (formatting.TopLinePunctuation is null && !ignoreNulls))
			paragraph.TopLinePunctuation(formatting.TopLinePunctuation);

		if (formatting.WidowControl is not null || (formatting.WidowControl is null && !ignoreNulls))
			paragraph.WidowControl(formatting.WidowControl);

		if (formatting.WordWrap is not null || (formatting.WordWrap is null && !ignoreNulls))
			paragraph.WordWrap(formatting.WordWrap);

		return paragraph;
	}

	/// <summary>
	/// Applies formatting to every run in the paragraph (replaces every possible property of each run unless <c>ignoreNulls</c> is set to true).
	/// </summary>
	public static Paragraph ApplyRunFormatting(this Paragraph paragraph, RunFormatting formatting, bool ignoreNulls = false)
	{
		paragraph.Elements<Run>().ToList().ForEach(r => r.ApplyFormatting(formatting, ignoreNulls));
		return paragraph;
	}

	/// <summary>
	/// Resets every possible property of the paragraph.
	/// </summary>
	public static Paragraph ResetFormatting(this Paragraph paragraph, bool includingRuns = false)
	{
		paragraph.RemoveAllChildren<ParagraphProperties>();

		if (includingRuns)
			paragraph.ResetRunFormatting();

		return paragraph;
	}

	/// <summary>
	/// Resets every possible property of each run in the paragraph.
	/// </summary>
	public static Paragraph ResetRunFormatting(this Paragraph paragraph)
	{
		paragraph.Elements<Run>().ToList().ForEach(r => r.ResetFormatting());
		return paragraph;
	}

	#endregion Formatting

	#region Get methods

	/// <summary>
	/// Gets the text that is a combination of all the runs in the paragraph.
	/// </summary>
	public static string GetText(this Paragraph paragraph)
	{
		var builder = new StringBuilder();

		foreach (Run run in paragraph.Elements<Run>())
			builder.Append(run.GetText());

		return builder.ToString();
	}

	/// <summary>
	/// Specifies the spacing that should be inserted between lines of text in the paragraph.
	/// <para>Child of <see cref="SpacingBetweenLines" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines" /></para>
	/// </summary>
	public static double? LineSpacingValue(this Paragraph paragraph, LineMeasuringUnits desiredUnits)
	{
		string? lineSpacing = paragraph.ParagraphProperties?.SpacingBetweenLines?.Line;

		return int.TryParse(lineSpacing, out int result)
			? Twips.ToOther(result, desiredUnits)
			: null;
	}

	/// <summary>
	/// Specifies the spacing that should be inserted before the paragraph.
	/// <para>Child of <see cref="SpacingBetweenLines" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines" /></para>
	/// </summary>
	public static double? SpacingBeforeValue(this Paragraph paragraph, LineMeasuringUnits desiredUnits)
	{
		string? spacingBefore = paragraph.ParagraphProperties?.SpacingBetweenLines?.Before;

		return int.TryParse(spacingBefore, out int result)
			? Twips.ToOther(result, desiredUnits)
			: null;
	}

	/// <summary>
	/// Specifies the spacing that should be inserted after the paragraph.
	/// <para>Child of <see cref="SpacingBetweenLines" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines" /></para>
	/// </summary>
	public static double? SpacingAfterValue(this Paragraph paragraph, LineMeasuringUnits desiredUnits)
	{
		string? spacingAfter = paragraph.ParagraphProperties?.SpacingBetweenLines?.After;

		return int.TryParse(spacingAfter, out int result)
			? Twips.ToOther(result, desiredUnits)
			: null;
	}

	/// <summary>
	/// Specifies the indentation of the left edge of the paragraph.
	/// <para>Child of <see cref="Indentation" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Indentation" /></para>
	/// </summary>
	public static double? LeftIndentationValue(this Paragraph paragraph, IndentationUnits desiredUnits)
	{
		Indentation? indentationProperty = paragraph.ParagraphProperties?.Indentation;

		return desiredUnits is IndentationUnits.Characters
			? indentationProperty?.LeftChars?.Value
			: int.TryParse(indentationProperty?.Left, out int result)
				? Twips.ToOther(result, desiredUnits)
				: null;
	}

	/// <summary>
	/// Specifies the indentation of the right edge of the paragraph.
	/// <para>Child of <see cref="Indentation" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Indentation" /></para>
	/// </summary>
	public static double? RightIndentationValue(this Paragraph paragraph, IndentationUnits desiredUnits)
	{
		Indentation? indentationProperty = paragraph.ParagraphProperties?.Indentation;

		return desiredUnits is IndentationUnits.Characters
			? indentationProperty?.RightChars?.Value
			: int.TryParse(indentationProperty?.Right, out int result)
				? Twips.ToOther(result, desiredUnits)
				: null;
	}

	#endregion Get methods

	#region Set methods

	/// <inheritdoc cref="LineSpacingValue" />
	public static Paragraph LineSpacing(this Paragraph paragraph, double spacing, LineMeasuringUnits units = LineMeasuringUnits.WholeLines)
	{
		paragraph
			.GetOrInit<ParagraphProperties>()
			.GetOrInit<SpacingBetweenLines>()
			.Line = Twips.FromOther(spacing, units)
				.ToString();

		return paragraph;
	}

	/// <inheritdoc cref="SpacingBeforeValue" />
	public static Paragraph SpacingBefore(this Paragraph paragraph, double spacing, LineMeasuringUnits units = LineMeasuringUnits.Points)
	{
		paragraph
			.GetOrInit<ParagraphProperties>()
			.GetOrInit<SpacingBetweenLines>()
			.Before = Twips.FromOther(spacing, units)
				.ToString();

		return paragraph;
	}

	/// <inheritdoc cref="SpacingAfterValue" />
	public static Paragraph SpacingAfter(this Paragraph paragraph, double spacing, LineMeasuringUnits units = LineMeasuringUnits.Points)
	{
		paragraph
			.GetOrInit<ParagraphProperties>()
			.GetOrInit<SpacingBetweenLines>()
			.After = Twips.FromOther(spacing, units)
				.ToString();

		return paragraph;
	}

	/// <inheritdoc cref="LeftIndentationValue" />
	public static Paragraph LeftIndentation(this Paragraph paragraph, double indentation, IndentationUnits units)
	{
		Indentation indentationProperty = paragraph
			.GetOrInit<ParagraphProperties>()
			.GetOrInit<Indentation>();

		if (units is IndentationUnits.Characters)
		{
			indentationProperty.LeftChars = (int)indentation;
			indentationProperty.Left = null;
		}
		else
		{
			indentationProperty.Left = Twips.FromOther(indentation, units).ToString();
			indentationProperty.LeftChars = null;
		}

		return paragraph;
	}

	/// <inheritdoc cref="RightIndentationValue" />
	public static Paragraph RightIndentation(this Paragraph paragraph, double indentation, IndentationUnits units)
	{
		Indentation indentationProperty = paragraph
			.GetOrInit<ParagraphProperties>()
			.GetOrInit<Indentation>();

		if (units is IndentationUnits.Characters)
		{
			indentationProperty.RightChars = (int)indentation;
			indentationProperty.Right = null;
		}
		else
		{
			indentationProperty.Right = Twips.FromOther(indentation, units).ToString();
			indentationProperty.RightChars = null;
		}

		return paragraph;
	}

	/// <summary>
	/// Sets the background fill color of the paragraph.
	/// <para><c>Fill</c> property of <see cref="Shading" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Paragraph FillColor(this Paragraph paragraph, string htmlColor)
	{
		Shading shading = paragraph.GetOrInit<ParagraphProperties>().GetOrInit<Shading>();
		shading.Fill = htmlColor;
		shading.Val = shading.Val?.Value is null or ShadingPatternValues.Nil
			? ShadingPatternValues.Clear
			: shading.Val.Value; // Preserve original pattern

		return paragraph;
	}

	#endregion Set methods

	#region Borders

	public static Paragraph LeftBorder(this Paragraph paragraph, LeftBorder? border)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().SetPropertyClassOrRemove(border);
		return paragraph;
	}

	public static Paragraph LeftBorder(this Paragraph paragraph, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().LeftBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return paragraph;
	}

	public static Paragraph TopBorder(this Paragraph paragraph, TopBorder? border)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().SetPropertyClassOrRemove(border);
		return paragraph;
	}

	public static Paragraph TopBorder(this Paragraph paragraph, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().TopBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return paragraph;
	}

	public static Paragraph RightBorder(this Paragraph paragraph, RightBorder? border)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().SetPropertyClassOrRemove(border);
		return paragraph;
	}

	public static Paragraph RightBorder(this Paragraph paragraph, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().RightBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return paragraph;
	}

	public static Paragraph BottomBorder(this Paragraph paragraph, BottomBorder? border)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().SetPropertyClassOrRemove(border);
		return paragraph;
	}

	public static Paragraph BottomBorder(this Paragraph paragraph, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().BottomBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return paragraph;
	}

	public static Paragraph BarBorder(this Paragraph paragraph, BarBorder? border)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().SetPropertyClassOrRemove(border);
		return paragraph;
	}

	public static Paragraph BarBorder(this Paragraph paragraph, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().BarBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return paragraph;
	}

	public static Paragraph BetweenBorder(this Paragraph paragraph, BetweenBorder? border)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().SetPropertyClassOrRemove(border);
		return paragraph;
	}

	public static Paragraph BetweenBorder(this Paragraph paragraph, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		paragraph.GetOrInit<ParagraphProperties>().GetOrInit<ParagraphBorders>().BetweenBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return paragraph;
	}

	#endregion Borders
}
