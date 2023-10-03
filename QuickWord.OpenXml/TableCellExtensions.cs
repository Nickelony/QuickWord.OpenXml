using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="TableCell" /> class.
/// </summary>
public static class TableCellExtensions
{
	#region Get property methods

	/// <summary>
	/// Specifies the set of conditional table style formatting properties which
	/// have been applied to this paragraph, if this paragraph is contained
	/// within a table cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static ConditionalFormatStyle? GetConditionalFormatStyle(this TableCell cell)
		=> cell.TableCellProperties?.ConditionalFormatStyle;

	/// <summary>
	/// Specifies the number of grid columns in the parent table's
	/// table grid which shall be spanned by the current cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridSpan" /></para>
	/// </summary>
	public static int? GridSpanValue(this TableCell cell)
		=> cell.TableCellProperties?.GridSpan?.Val?.Value;

	/// <summary>
	/// Specifies whether the end of cell glyph shall influence
	/// the height of the given table row in the table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.HideMark" /></para>
	/// </summary>
	public static bool? HideEndOfCellMarkerValue(this TableCell cell)
	{
		OnOffOnlyValues? value = cell.TableCellProperties?.HideMark?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies that this cell is part of a horizontally merged set of cells in a table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.HorizontalMerge" /></para>
	/// </summary>
	public static MergedCellValues? HorizontalMergeValue(this TableCell cell)
		=> cell.TableCellProperties?.HorizontalMerge?.Val?.Value;

	/// <summary>
	/// Specifies how this table cell shall be laid out when the parent
	/// table is displayed in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NoWrap" /></para>
	/// </summary>
	public static bool? NoContentWrappingValue(this TableCell cell)
	{
		OnOffOnlyValues? value = cell.TableCellProperties?.NoWrap?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies the shading applied to the contents of the cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Shading? GetShading(this TableCell cell)
		=> cell.TableCellProperties?.Shading;

	/// <summary>
	/// Specifies the set of borders for the edges of the current
	/// table cell, using the eight border types defined by its child elements.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellBorders" /></para>
	/// </summary>
	public static TableCellBorders? GetBorders(this TableCell cell)
		=> cell.TableCellProperties?.TableCellBorders;

	/// <summary>
	/// Specifies that the contents of the current cell shall have their
	/// inter-character spacing increased or reduced as necessary to fit
	/// the width of the text extents of the current cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellFitText" /></para>
	/// </summary>
	public static bool? FitTextValue(this TableCell cell)
	{
		OnOffOnlyValues? value = cell.TableCellProperties?.TableCellFitText?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies a set of cell margins for a single table cell in the parent table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellMargin" /></para>
	/// </summary>
	public static TableCellMargin? GetMargins(this TableCell cell)
		=> cell.TableCellProperties?.TableCellMargin;

	/// <summary>
	/// Specifies the preferred width for this table cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellWidth" /></para>
	/// </summary>
	public static TableCellWidth? GetWidth(this TableCell cell)
		=> cell.TableCellProperties?.TableCellWidth;

	/// <summary>
	/// Specifies the direction of the text flow for this cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextDirection" /></para>
	/// </summary>
	public static TextDirectionValues? TextDirectionValue(this TableCell cell)
		=> cell.TableCellProperties?.TextDirection?.Val?.Value;

	/// <summary>
	/// Specifies the vertical alignment of the contents of the current cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellVerticalAlignment" /></para>
	/// </summary>
	public static TableVerticalAlignmentValues? VerticalContentAlignmentValue(this TableCell cell)
		=> cell.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value;

	/// <summary>
	/// Specifies that this cell is part of a vertically merged set of cells in a table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.VerticalMerge" /></para>
	/// </summary>
	public static MergedCellValues? VerticalMergeValue(this TableCell cell)
		=> cell.TableCellProperties?.VerticalMerge?.Val?.Value;

	#endregion Get property methods

	#region Set property methods

	/// <summary>
	/// Specifies the set of conditional table style formatting properties which
	/// have been applied to this paragraph, if this paragraph is contained
	/// within a table cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static TableCell ConditionalFormatStyle(this TableCell cell, ConditionalFormatStyle? style)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(style);
		return cell;
	}

	/// <summary>
	/// Specifies the number of grid columns in the parent table's
	/// table grid which shall be spanned by the current cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridSpan" /></para>
	/// </summary>
	public static TableCell GridSpan(this TableCell cell, int? span)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<GridSpan>(span);
		return cell;
	}

	/// <summary>
	/// Specifies whether the end of cell glyph shall influence
	/// the height of the given table row in the table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.HideMark" /></para>
	/// </summary>
	public static TableCell HideEndOfCellMarker(this TableCell cell, bool? value = true)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<HideMark>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return cell;
	}

	/// <summary>
	/// Specifies that this cell is part of a horizontally merged set of cells in a table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.HorizontalMerge" /></para>
	/// </summary>
	public static TableCell HorizontalMerge(this TableCell cell, MergedCellValues? merge)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<HorizontalMerge>(merge);
		return cell;
	}

	/// <summary>
	/// Specifies how this table cell shall be laid out when the parent
	/// table is displayed in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NoWrap" /></para>
	/// </summary>
	public static TableCell NoContentWrapping(this TableCell cell, bool? value = true)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<NoWrap>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return cell;
	}

	/// <summary>
	/// Specifies the shading applied to the contents of the cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static TableCell Shading(this TableCell cell, Shading? shading)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(shading);
		return cell;
	}

	/// <summary>
	/// Specifies the set of borders for the edges of the current
	/// table cell, using the eight border types defined by its child elements.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellBorders" /></para>
	/// </summary>
	public static TableCell Borders(this TableCell cell, TableCellBorders? borders)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(borders);
		return cell;
	}

	/// <summary>
	/// Specifies that the contents of the current cell shall have their
	/// inter-character spacing increased or reduced as necessary to fit
	/// the width of the text extents of the current cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellFitText" /></para>
	/// </summary>
	public static TableCell FitText(this TableCell cell, bool? value = true)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<TableCellFitText>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return cell;
	}

	/// <summary>
	/// Specifies a set of cell margins for a single table cell in the parent table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellMargin" /></para>
	/// </summary>
	public static TableCell Margin(this TableCell cell, TableCellMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Specifies the preferred width for this table cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellWidth" /></para>
	/// </summary>
	public static TableCell Width(this TableCell cell, TableCellWidth? width)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(width);
		return cell;
	}

	/// <summary>
	/// Specifies the direction of the text flow for this cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextDirection" /></para>
	/// </summary>
	public static TableCell TextDirection(this TableCell cell, TextDirectionValues? direction)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<TextDirection>(direction);
		return cell;
	}

	/// <summary>
	/// Specifies the vertical alignment of the contents of the current cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellVerticalAlignment" /></para>
	/// </summary>
	public static TableCell VerticalContentAlignment(this TableCell cell, TableVerticalAlignmentValues? alignment)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<TableCellVerticalAlignment>(alignment);
		return cell;
	}

	/// <summary>
	/// Specifies that this cell is part of a vertically merged set of cells in a table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.VerticalMerge" /></para>
	/// </summary>
	public static TableCell VerticalMerge(this TableCell cell, MergedCellValues? merge)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<VerticalMerge>(merge);
		return cell;
	}

	#endregion Set property methods
}
