// Ignore Spelling: Kinsoku

using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="Paragraph" /> class.
/// </summary>
public static class ParagraphExtensions
{
	#region Get property methods

	/// <summary>
	/// Specifies whether the right indent shall be automatically adjusted for the given paragraph when
	/// a document grid has been defined for the current section using the docGrid element (§17.6.5),
	/// modifying of the current right indent used on this paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.AdjustRightIndent" /></para>
	/// </summary>
	public static bool? AdjustRightIndentValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.AdjustRightIndent?.Val?.Value;

	/// <summary>
	/// Specifies whether inter-character spacing shall automatically be adjusted between
	/// regions of Latin text and regions of East Asian text in the current paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.AutoSpaceDE" /></para>
	/// </summary>
	public static bool? AutoSpaceDEValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.AutoSpaceDE?.Val?.Value;

	/// <summary>
	/// Specifies whether inter-character spacing shall automatically be adjusted between
	/// regions of numbers and regions of East Asian text in the current paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.AutoSpaceDN" /></para>
	/// </summary>
	public static bool? AutoSpaceDNValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.AutoSpaceDN?.Val?.Value;

	/// <summary>
	/// Specifies that this paragraph shall be displayed from right to left.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.BiDi" /></para>
	/// </summary>
	public static bool? BiDirectionalValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.BiDi?.Val?.Value;

	/// <summary>
	/// Specifies the set of conditional table style formatting properties which
	/// have been applied to this paragraph, if this paragraph is contained within a table cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static ConditionalFormatStyle? GetConditionalFormatStyle(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.ConditionalFormatStyle;

	/// <summary>
	/// Specifies that any space specified before or after this paragraph, specified using
	/// the spacing element (§17.3.1.33), should not be applied when
	/// the preceding and following paragraphs are of the same paragraph style,
	/// affecting the top and bottom spacing respectively.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ContextualSpacing" /></para>
	/// </summary>
	public static bool? ContextualSpacingValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.ContextualSpacing?.Val?.Value;

	/// <summary>
	/// Specifies that this paragraph should be located within the specified
	/// HTML <i>div</i> tag when this document is saved in HTML format.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.DivId" /></para>
	/// </summary>
	public static string? DivIdValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.DivId?.Val;

	/// <summary>
	/// Specifies information about the current paragraph with regard to text frames.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FrameProperties" /></para>
	/// </summary>
	public static FrameProperties? GetFrameProperties(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.FrameProperties;

	/// <summary>
	/// Specifies the set of indentation properties applied to the current paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Indentation" /></para>
	/// </summary>
	public static Indentation? GetIndentation(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.Indentation;

	/// <summary>
	/// Specifies the paragraph alignment which shall be applied to text in this paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Justification" /></para>
	/// </summary>
	public static JustificationValues? JustificationValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.Justification?.Val?.Value;

	/// <summary>
	/// Specifies that when rendering this document in a page view, all lines
	/// of this paragraph are maintained on a single page whenever possible.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.KeepLines" /></para>
	/// </summary>
	public static bool? KeepLinesTogetherValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.KeepLines?.Val?.Value;

	/// <summary>
	/// Specifies that when rendering this document in a paginated view, the contents
	/// of this paragraph are at least partly rendered on the same page as
	/// the following paragraph whenever possible.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.KeepNext" /></para>
	/// </summary>
	public static bool? KeepWithNextValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.KeepNext?.Val?.Value;

	/// <summary>
	/// Specifies whether East Asian typography and line-breaking rules shall be applied
	/// to text in this paragraph to determine which characters can begin and end each line.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Kinsoku" /></para>
	/// </summary>
	public static bool? KinsokuValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.Kinsoku?.Val?.Value;

	/// <summary>
	/// Specifies whether the paragraph indents should be interpreted as mirrored indents.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.MirrorIndents" /></para>
	/// </summary>
	public static bool? MirrorIndentsValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.MirrorIndents?.Val?.Value;

	/// <summary>
	/// Specifies that the current paragraph references a numbering definition instance in the current document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NumberingProperties" /></para>
	/// </summary>
	public static NumberingProperties? GetNumberingProperties(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.NumberingProperties;

	/// <summary>
	/// Specifies the outline level which shall be associated with the current paragraph in the document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.OutlineLevel" /></para>
	/// </summary>
	public static OutlineLevelValues? OutlineLevelValue(this Paragraph paragraph)
		=> (OutlineLevelValues?)paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;

	/// <summary>
	/// Specifies that the text in this paragraph shall be allowed to extend one character
	/// beyond the extents applied by any indents/margins when the character that extends
	/// past those extents is a punctuation character.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.OverflowPunctuation" /></para>
	/// </summary>
	public static bool? OverflowPunctuationValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.OverflowPunctuation?.Val?.Value;

	/// <summary>
	/// Specifies that when rendering this document in a paginated view, the contents
	/// of this paragraph are rendered on the start of a new page in the document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore" /></para>
	/// </summary>
	public static bool? PageBreakBeforeValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.PageBreakBefore?.Val?.Value;

	/// <summary>
	/// Specifies the borders for the paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ParagraphBorders" /></para>
	/// </summary>
	public static ParagraphBorders? GetBorders(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.ParagraphBorders;

	/// <summary>
	/// Specifies the style ID of the paragraph style which shall be used to format the contents of this paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId" /></para>
	/// </summary>
	public static string? StyleValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.ParagraphStyleId?.Val;

	/// <summary>
	/// Specifies the set of run properties applied to the glyph used
	/// to represent the physical location of the paragraph mark for this paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ParagraphMarkRunProperties" /></para>
	/// </summary>
	public static ParagraphMarkRunProperties? GetMarkRunProperties(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.ParagraphMarkRunProperties;

	/// <summary>
	/// Specifies the section properties for the final section of the document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SectionProperties" /></para>
	/// </summary>
	public static SectionProperties? GetSectionProperties(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.SectionProperties;

	/// <summary>
	/// Specifies the shading applied to the contents of the paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Shading? GetShading(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.Shading;

	/// <summary>
	/// Specifies whether the current paragraph should use the document grid lines per page settings defined
	/// in the docGrid element (§17.6.5) when laying out the contents in the paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SnapToGrid" /></para>
	/// </summary>
	public static bool? SnapToGridValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.SnapToGrid?.Val?.Value;

	/// <summary>
	/// Specifies the inter-line and inter-paragraph spacing which shall be applied
	/// to the contents of this paragraph when it is displayed by a consumer.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines" /></para>
	/// </summary>
	public static SpacingBetweenLines? GetSpacing(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.SpacingBetweenLines;

	/// <summary>
	/// Specifies whether any hyphenation shall be performed on this paragraph by the consumer when
	/// requested using the autoHyphenation element (§17.15.1.10) in the document's settings.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SuppressAutoHyphens" /></para>
	/// </summary>
	public static bool? SuppressAutoHyphenationValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.SuppressAutoHyphens?.Val?.Value;

	/// <summary>
	/// Specifies whether line numbers shall be calculated for lines in this paragraph by the consumer when
	/// line numbering is requested using the lnNumType element (§17.6.8) in the paragraph's
	/// parent section settings.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SuppressLineNumbers" /></para>
	/// </summary>
	public static bool? SuppressLineNumbersValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.SuppressLineNumbers?.Val?.Value;

	/// <summary>
	/// Specifies whether a text frame which intersects another text frame at display time shall
	/// be allowed to overlap the contents of the other text frame.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SuppressOverlap" /></para>
	/// </summary>
	public static bool? SuppressOverlappingValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.SuppressOverlap?.Val?.Value;

	/// <summary>
	/// Specifies a sequence of custom tab stops which shall be used for
	/// any tab characters in the current paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Tabs" /></para>
	/// </summary>
	public static Tabs? GetTabs(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.Tabs;

	/// <summary>
	/// Specifies the vertical alignment of all text on each line displayed within the paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextAlignment" /></para>
	/// </summary>
	public static VerticalTextAlignmentValues? VerticalTextAlignmentValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.TextAlignment?.Val?.Value;

	/// <summary>
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextBoxTightWrap" /></para>
	/// </summary>
	public static TextBoxTightWrapValues? TextBoxTightWrapValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.TextBoxTightWrap?.Val?.Value;

	/// <summary>
	/// Specifies the direction of the text flow for this paragraph.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextDirection" /></para>
	/// </summary>
	public static TextDirectionValues? TextDirectionValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.TextDirection?.Val?.Value;

	/// <summary>
	/// Specifies whether punctuation shall be compressed when it appears as the first
	/// character in a line, allowing subsequent characters on the line to be move in accordingly.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TopLinePunctuation" /></para>
	/// </summary>
	public static bool? TopLinePunctuationValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.TopLinePunctuation?.Val?.Value;

	/// <summary>
	/// Specifies whether a consumer shall prevent a single line of this paragraph from
	/// being displayed on a separate page from the remaining content at display time by moving
	/// the line onto the following page.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidowControl" /></para>
	/// </summary>
	public static bool? WidowControlValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.WidowControl?.Val?.Value;

	/// <summary>
	/// Specifies whether a consumer shall break text which exceeds the text extents of
	/// a line by breaking the word across two lines (breaking on the character level) or by
	/// moving the word to the following line (breaking on the word level).
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WordWrap" /></para>
	/// </summary>
	public static bool? WordWrapValue(this Paragraph paragraph)
		=> paragraph.ParagraphProperties?.WordWrap?.Val?.Value;

	#endregion Get property methods

	#region Set property methods

	/// <summary>
	/// Specifies whether the right indent shall be automatically adjusted for the given paragraph when
	/// a document grid has been defined for the current section using the docGrid element (§17.6.5),
	/// modifying of the current right indent used on this paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.AdjustRightIndent" /></para>
	/// </summary>
	public static Paragraph AdjustRightIndent(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<AdjustRightIndent>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether inter-character spacing shall automatically be adjusted between
	/// regions of Latin text and regions of East Asian text in the current paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.AutoSpaceDE" /></para>
	/// </summary>
	public static Paragraph AutoSpaceDE(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<AutoSpaceDE>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether inter-character spacing shall automatically be adjusted between
	/// regions of numbers and regions of East Asian text in the current paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.AutoSpaceDN" /></para>
	/// </summary>
	public static Paragraph AutoSpaceDN(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<AutoSpaceDN>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies that this paragraph shall be displayed from right to left.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.BiDi" /></para>
	/// </summary>
	public static Paragraph BiDirectional(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<BiDi>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies the set of conditional table style formatting properties which
	/// have been applied to this paragraph, if this paragraph is contained within a table cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static Paragraph ConditionalFormatStyle(this Paragraph paragraph, ConditionalFormatStyle? style)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(style);
		return paragraph;
	}

	/// <summary>
	/// Specifies that any space specified before or after this paragraph, specified using
	/// the spacing element (§17.3.1.33), should not be applied when
	/// the preceding and following paragraphs are of the same paragraph style,
	/// affecting the top and bottom spacing respectively.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ContextualSpacing" /></para>
	/// </summary>
	public static Paragraph ContextualSpacing(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<ContextualSpacing>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies that this paragraph should be located within the specified
	/// HTML <i>div</i> tag when this document is saved in HTML format.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.DivId" /></para>
	/// </summary>
	public static Paragraph DivId(this Paragraph paragraph, string? id)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<DivId>(id);
		return paragraph;
	}

	/// <summary>
	/// Specifies information about the current paragraph with regard to text frames.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.FrameProperties" /></para>
	/// </summary>
	public static Paragraph FrameProperties(this Paragraph paragraph, FrameProperties? properties)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(properties);
		return paragraph;
	}

	/// <summary>
	/// Specifies the set of indentation properties applied to the current paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Indentation" /></para>
	/// </summary>
	public static Paragraph Indentation(this Paragraph paragraph, Indentation? indentation)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(indentation);
		return paragraph;
	}

	/// <summary>
	/// Specifies the paragraph alignment which shall be applied to text in this paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Justification" /></para>
	/// </summary>
	public static Paragraph Justification(this Paragraph paragraph, JustificationValues? justification)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<Justification>(justification);
		return paragraph;
	}

	/// <summary>
	/// Specifies that when rendering this document in a page view, all lines
	/// of this paragraph are maintained on a single page whenever possible.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.KeepLines" /></para>
	/// </summary>
	public static Paragraph KeepLinesTogether(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<KeepLines>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies that when rendering this document in a paginated view, the contents
	/// of this paragraph are at least partly rendered on the same page as
	/// the following paragraph whenever possible.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.KeepNext" /></para>
	/// </summary>
	public static Paragraph KeepWithNext(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<KeepNext>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether East Asian typography and line-breaking rules shall be applied
	/// to text in this paragraph to determine which characters can begin and end each line.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Kinsoku" /></para>
	/// </summary>
	public static Paragraph Kinsoku(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<Kinsoku>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether the paragraph indents should be interpreted as mirrored indents.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.MirrorIndents" /></para>
	/// </summary>
	public static Paragraph MirrorIndents(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<MirrorIndents>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies that the current paragraph references a numbering definition instance in the current document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NumberingProperties" /></para>
	/// </summary>
	public static Paragraph NumberingProperties(this Paragraph paragraph, NumberingProperties? properties)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(properties);
		return paragraph;
	}

	/// <summary>
	/// Specifies the outline level which shall be associated with the current paragraph in the document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.OutlineLevel" /></para>
	/// </summary>
	public static Paragraph OutlineLevel(this Paragraph paragraph, OutlineLevelValues? level)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<OutlineLevel>((int?)level);
		return paragraph;
	}

	/// <summary>
	/// Specifies that the text in this paragraph shall be allowed to extend one character
	/// beyond the extents applied by any indents/margins when the character that extends
	/// past those extents is a punctuation character.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.OverflowPunctuation" /></para>
	/// </summary>
	public static Paragraph OverflowPunctuation(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<OverflowPunctuation>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies that when rendering this document in a paginated view, the contents
	/// of this paragraph are rendered on the start of a new page in the document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.PageBreakBefore" /></para>
	/// </summary>
	public static Paragraph PageBreakBefore(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<PageBreakBefore>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies the borders for the paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ParagraphBorders" /></para>
	/// </summary>
	public static Paragraph Borders(this Paragraph paragraph, ParagraphBorders? borders)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(borders);
		return paragraph;
	}

	/// <summary>
	/// Specifies the style ID of the paragraph style which shall be used to format the contents of this paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ParagraphStyleId" /></para>
	/// </summary>
	public static Paragraph Style(this Paragraph paragraph, string? styleId)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<ParagraphStyleId>(styleId);
		return paragraph;
	}

	/// <summary>
	/// Specifies the set of run properties applied to the glyph used
	/// to represent the physical location of the paragraph mark for this paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ParagraphMarkRunProperties" /></para>
	/// </summary>
	public static Paragraph MarkRunProperties(this Paragraph paragraph, ParagraphMarkRunProperties? properties)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(properties);
		return paragraph;
	}

	/// <summary>
	/// Specifies the section properties for the final section of the document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SectionProperties" /></para>
	/// </summary>
	public static Paragraph SectionProperties(this Paragraph paragraph, SectionProperties? properties)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(properties);
		return paragraph;
	}

	/// <summary>
	/// Specifies the shading applied to the contents of the paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Paragraph Shading(this Paragraph paragraph, Shading? shading)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(shading);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether the current paragraph should use the document grid lines per page settings defined
	/// in the docGrid element (§17.6.5) when laying out the contents in the paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SnapToGrid" /></para>
	/// </summary>
	public static Paragraph SnapToGrid(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<SnapToGrid>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies the inter-line and inter-paragraph spacing which shall be applied
	/// to the contents of this paragraph when it is displayed by a consumer.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SpacingBetweenLines" /></para>
	/// </summary>
	public static Paragraph Spacing(this Paragraph paragraph, SpacingBetweenLines? spacing)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(spacing);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether any hyphenation shall be performed on this paragraph by the consumer when
	/// requested using the autoHyphenation element (§17.15.1.10) in the document's settings.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SuppressAutoHyphens" /></para>
	/// </summary>
	public static Paragraph SuppressAutoHyphenation(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<SuppressAutoHyphens>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether line numbers shall be calculated for lines in this paragraph by the consumer when
	/// line numbering is requested using the lnNumType element (§17.6.8) in the paragraph's
	/// parent section settings.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SuppressLineNumbers" /></para>
	/// </summary>
	public static Paragraph SuppressLineNumbers(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<SuppressLineNumbers>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether a text frame which intersects another text frame at display time shall
	/// be allowed to overlap the contents of the other text frame.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.SuppressOverlap" /></para>
	/// </summary>
	public static Paragraph SuppressOverlapping(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<SuppressOverlap>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies a sequence of custom tab stops which shall be used for
	/// any tab characters in the current paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Tabs" /></para>
	/// </summary>
	public static Paragraph Tabs(this Paragraph paragraph, Tabs? tabs)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetPropertyClassOrRemove(tabs);
		return paragraph;
	}

	/// <summary>
	/// Specifies the vertical alignment of all text on each line displayed within the paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextAlignment" /></para>
	/// </summary>
	public static Paragraph VerticalTextAlignment(this Paragraph paragraph, VerticalTextAlignmentValues? alignment)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<TextAlignment>(alignment);
		return paragraph;
	}

	/// <summary>
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextBoxTightWrap" /></para>
	/// </summary>
	public static Paragraph TextBoxTightWrap(this Paragraph paragraph, TextBoxTightWrapValues? tightWrap)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<TextBoxTightWrap>(tightWrap);
		return paragraph;
	}

	/// <summary>
	/// Specifies the direction of the text flow for this paragraph.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextDirection" /></para>
	/// </summary>
	public static Paragraph TextDirection(this Paragraph paragraph, TextDirectionValues? direction)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<TextDirection>(direction);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether punctuation shall be compressed when it appears as the first
	/// character in a line, allowing subsequent characters on the line to be move in accordingly.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TopLinePunctuation" /></para>
	/// </summary>
	public static Paragraph TopLinePunctuation(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<TopLinePunctuation>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether a consumer shall prevent a single line of this paragraph from
	/// being displayed on a separate page from the remaining content at display time by moving
	/// the line onto the following page.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidowControl" /></para>
	/// </summary>
	public static Paragraph WidowControl(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<WidowControl>(value);
		return paragraph;
	}

	/// <summary>
	/// Specifies whether a consumer shall break text which exceeds the text extents of
	/// a line by breaking the word across two lines (breaking on the character level) or by
	/// moving the word to the following line (breaking on the word level).
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WordWrap" /></para>
	/// </summary>
	public static Paragraph WordWrap(this Paragraph paragraph, bool? value = true)
	{
		paragraph.GetOrInit<ParagraphProperties>().SetValOrRemove<WordWrap>(value);
		return paragraph;
	}

	#endregion Set property methods
}
