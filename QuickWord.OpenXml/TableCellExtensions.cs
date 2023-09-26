using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

public static class TableCellExtensions
{
	#region Get property methods

	/// <summary>
	/// Specifies the set of conditional table style formatting properties which
	/// have been applied to this paragraph, if this paragraph is contained
	/// within a table cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static ConditionalFormatStyle? GetConditionalFormatStyle(this TableCell cell)
		=> cell.TableCellProperties?.ConditionalFormatStyle;

	/// <summary>
	/// Specifies the number of grid columns in the parent table's
	/// table grid which shall be spanned by the current cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridSpan" /></para>
	/// </summary>
	public static int? GridSpanValue(this TableCell cell)
		=> cell.TableCellProperties?.GridSpan?.Val?.Value;

	/// <summary>
	/// Specifies whether the end of cell glyph shall influence
	/// the height of the given table row in the table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.HideMark" /></para>
	/// </summary>
	public static bool? HideEndOfCellMarkerValue(this TableCell cell)
	{
		OnOffOnlyValues? value = cell.TableCellProperties?.HideMark?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies that this cell is part of a horizontally merged set of cells in a table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.HorizontalMerge" /></para>
	/// </summary>
	public static MergedCellValues? HorizontalMergeValue(this TableCell cell)
		=> cell.TableCellProperties?.HorizontalMerge?.Val?.Value;

	/// <summary>
	/// Specifies how this table cell shall be laid out when the parent
	/// table is displayed in a document.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.NoWrap" /></para>
	/// </summary>
	public static bool? NoContentWrappingValue(this TableCell cell)
	{
		OnOffOnlyValues? value = cell.TableCellProperties?.NoWrap?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies the shading applied to the contents of the cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Shading? GetShading(this TableCell cell)
		=> cell.TableCellProperties?.Shading;

	/// <summary>
	/// Specifies the set of borders for the edges of the current
	/// table cell, using the eight border types defined by its child elements.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellBorders" /></para>
	/// </summary>
	public static TableCellBorders? GetBorders(this TableCell cell)
		=> cell.TableCellProperties?.TableCellBorders;

	/// <summary>
	/// Specifies that the contents of the current cell shall have their
	/// inter-character spacing increased or reduced as necessary to fit
	/// the width of the text extents of the current cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellFitText" /></para>
	/// </summary>
	public static bool? FitTextValue(this TableCell cell)
	{
		OnOffOnlyValues? value = cell.TableCellProperties?.TableCellFitText?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies a set of cell margins for a single table cell in the parent table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellMargin" /></para>
	/// </summary>
	public static TableCellMargin? GetMargins(this TableCell cell)
		=> cell.TableCellProperties?.TableCellMargin;

	/// <summary>
	/// Specifies the preferred width for this table cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellWidth" /></para>
	/// </summary>
	public static TableCellWidth? GetWidth(this TableCell cell)
		=> cell.TableCellProperties?.TableCellWidth;

	/// <summary>
	/// Specifies the direction of the text flow for this cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TextDirection" /></para>
	/// </summary>
	public static TextDirectionValues? TextDirectionValue(this TableCell cell)
		=> cell.TableCellProperties?.TextDirection?.Val?.Value;

	/// <summary>
	/// Specifies the vertical alignment of the contents of the current cell.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellVerticalAlignment" /></para>
	/// </summary>
	public static TableVerticalAlignmentValues? VerticalContentAlignmentValue(this TableCell cell)
		=> cell.TableCellProperties?.TableCellVerticalAlignment?.Val?.Value;

	/// <summary>
	/// Specifies that this cell is part of a vertically merged set of cells in a table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.VerticalMerge" /></para>
	/// </summary>
	public static MergedCellValues? VerticalMergeValue(this TableCell cell)
		=> cell.TableCellProperties?.VerticalMerge?.Val?.Value;

	#endregion Get property methods

	#region Set property methods

	/// <inheritdoc cref="GetConditionalFormatStyle" />
	public static TableCell ConditionalFormatStyle(this TableCell cell, ConditionalFormatStyle? style)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(style);
		return cell;
	}

	/// <inheritdoc cref="GridSpanValue" />
	public static TableCell GridSpan(this TableCell cell, int? span)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<GridSpan>(span);
		return cell;
	}

	/// <inheritdoc cref="HideEndOfCellMarkerValue" />
	public static TableCell HideEndOfCellMarker(this TableCell cell, bool? value = true)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<HideMark>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return cell;
	}

	/// <inheritdoc cref="HorizontalMergeValue" />
	public static TableCell HorizontalMerge(this TableCell cell, MergedCellValues? merge)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<HorizontalMerge>(merge);
		return cell;
	}

	/// <inheritdoc cref="NoContentWrappingValue" />
	public static TableCell NoContentWrapping(this TableCell cell, bool? value = true)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<NoWrap>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return cell;
	}

	/// <inheritdoc cref="GetShading" />
	public static TableCell Shading(this TableCell cell, Shading? shading)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(shading);
		return cell;
	}

	/// <inheritdoc cref="GetBorders" />
	public static TableCell Borders(this TableCell cell, TableCellBorders? borders)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(borders);
		return cell;
	}

	/// <inheritdoc cref="FitTextValue" />
	public static TableCell FitText(this TableCell cell, bool? value = true)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<TableCellFitText>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return cell;
	}

	/// <inheritdoc cref="GetMargins" />
	public static TableCell Margin(this TableCell cell, TableCellMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <inheritdoc cref="GetWidth" />
	public static TableCell Width(this TableCell cell, TableCellWidth? width)
	{
		cell.GetOrInit<TableCellProperties>().SetPropertyClassOrRemove(width);
		return cell;
	}

	/// <inheritdoc cref="TextDirectionValue" />
	public static TableCell TextDirection(this TableCell cell, TextDirectionValues? direction)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<TextDirection>(direction);
		return cell;
	}

	/// <inheritdoc cref="VerticalContentAlignmentValue" />
	public static TableCell VerticalContentAlignment(this TableCell cell, TableVerticalAlignmentValues? alignment)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<TableCellVerticalAlignment>(alignment);
		return cell;
	}

	/// <inheritdoc cref="VerticalMergeValue" />
	public static TableCell VerticalMerge(this TableCell cell, MergedCellValues? merge)
	{
		cell.GetOrInit<TableCellProperties>().SetValOrRemove<VerticalMerge>(merge);
		return cell;
	}

	#endregion Set property methods
}
