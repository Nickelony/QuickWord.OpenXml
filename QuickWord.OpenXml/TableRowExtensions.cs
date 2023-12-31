﻿using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="TableRow" /> class.
/// </summary>
public static class TableRowExtensions
{
	#region Get property methods

	/// <summary>
	/// Specifies whether the contents within the current cell shall
	/// be rendered on a single page.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.CantSplit" /></para>
	/// </summary>
	public static bool? CantSplitValue(this TableRow row)
	{
		OnOffOnlyValues? value = row.TableRowProperties?.GetFirstChild<CantSplit>()?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies the set of conditional table style formatting properties
	/// which have been applied to this paragraph, if this paragraph
	/// is contained within a table cell.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static ConditionalFormatStyle? GetConditionalFormatStyle(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<ConditionalFormatStyle>();

	/// <summary>
	/// Specifies that this paragraph should be located within the specified
	/// HTML <i>div</i> tag when this document is saved in HTML format.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.DivId" /></para>
	/// </summary>
	public static string? DivIdValue(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<DivId>()?.Val?.Value;

	/// <summary>
	/// Specifies the number of grid columns in the parent table's table grid
	/// (§17.4.49; §17.4.48) which shall be left after the last cell in the table row.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridAfter" /></para>
	/// </summary>
	public static int? GridAfterValue(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<GridAfter>()?.Val?.Value;

	/// <summary>
	/// Specifies the number of grid columns in the parent table's table grid
	/// (§17.4.49; §17.4.48) which must be skipped before the contents
	/// of this table row (its table cells) are added to the parent table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridBefore" /></para>
	/// </summary>
	public static int? GridBeforeValue(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<GridBefore>()?.Val?.Value;

	/// <summary>
	/// Specifies that the glyph representing the end character of current
	/// table row shall not be displayed in the current document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Hidden" /></para>
	/// </summary>
	public static bool? HideEndOfRowMarkerValue(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<Hidden>()?.Val?.Value;

	/// <summary>
	/// Specifies the alignment of the set of rows which are part
	/// of the current table properties exception list with respect
	/// to the text margins in the current section.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableJustification" /></para>
	/// </summary>
	public static TableRowAlignmentValues? JustificationValue(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<TableJustification>()?.Val?.Value;

	/// <summary>
	/// Specifies that the current table row shall be repeated at the top
	/// of each new page on which part of this table is displayed.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableHeader" /></para>
	/// </summary>
	public static bool? IsHeaderValue(this TableRow row)
	{
		OnOffOnlyValues? value = row.TableRowProperties?.GetFirstChild<TableHeader>()?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies the height of the current table row within the current table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableRowHeight" /></para>
	/// </summary>
	public static TableRowHeight? GetHeight(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<TableRowHeight>();

	/// <summary>
	/// Specifies the preferred width for the total number of grid columns
	/// after this table row as specified in the gridAfter element (§17.4.14).
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthAfterTableRow" /></para>
	/// </summary>
	public static WidthAfterTableRow? GetWidthAfter(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<WidthAfterTableRow>();

	/// <summary>
	/// Specifies the preferred width for the total number of grid columns
	/// before this table row as specified in the gridAfter element (§17.4.14).
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthBeforeTableRow" /></para>
	/// </summary>
	public static WidthBeforeTableRow? GetWidthBefore(this TableRow row)
		=> row.TableRowProperties?.GetFirstChild<WidthBeforeTableRow>();

	#endregion Get property methods

	#region Set property methods

	/// <summary>
	/// Specifies whether the contents within the current cell shall
	/// be rendered on a single page.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.CantSplit" /></para>
	/// </summary>
	public static TableRow CantSplit(this TableRow row, bool? value = true)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<CantSplit>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return row;
	}

	/// <summary>
	/// Specifies the set of conditional table style formatting properties
	/// which have been applied to this paragraph, if this paragraph
	/// is contained within a table cell.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.ConditionalFormatStyle" /></para>
	/// </summary>
	public static TableRow ConditionalFormatStyle(this TableRow row, ConditionalFormatStyle? style)
	{
		row.GetOrInit<TableRowProperties>().SetPropertyClassOrRemove(style);
		return row;
	}

	/// <summary>
	/// Specifies that this paragraph should be located within the specified
	/// HTML <i>div</i> tag when this document is saved in HTML format.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.DivId" /></para>
	/// </summary>
	public static TableRow DivId(this TableRow row, string? id)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<DivId>(id);
		return row;
	}

	/// <summary>
	/// Specifies the number of grid columns in the parent table's table grid
	/// (§17.4.49; §17.4.48) which shall be left after the last cell in the table row.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridAfter" /></para>
	/// </summary>
	public static TableRow GridAfter(this TableRow row, int? columns)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<GridAfter>(columns);
		return row;
	}

	/// <summary>
	/// Specifies the number of grid columns in the parent table's table grid
	/// (§17.4.49; §17.4.48) which must be skipped before the contents
	/// of this table row (its table cells) are added to the parent table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.GridBefore" /></para>
	/// </summary>
	public static TableRow GridBefore(this TableRow row, int? columns)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<GridBefore>(columns);
		return row;
	}

	/// <summary>
	/// Specifies that the glyph representing the end character of current
	/// table row shall not be displayed in the current document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Hidden" /></para>
	/// </summary>
	public static TableRow HideEndOfRowMarker(this TableRow row, bool? value = true)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<Hidden>(value);
		return row;
	}

	/// <summary>
	/// Specifies the alignment of the set of rows which are part
	/// of the current table properties exception list with respect
	/// to the text margins in the current section.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableJustification" /></para>
	/// </summary>
	public static TableRow Justification(this TableRow row, TableRowAlignmentValues? alignment)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<TableJustification>(alignment);
		return row;
	}

	/// <summary>
	/// Specifies that the current table row shall be repeated at the top
	/// of each new page on which part of this table is displayed.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableHeader" /></para>
	/// </summary>
	public static TableRow IsHeader(this TableRow row, bool? value = true)
	{
		row.GetOrInit<TableRowProperties>().SetValOrRemove<TableHeader>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return row;
	}

	/// <summary>
	/// Specifies the height of the current table row within the current table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableRowHeight" /></para>
	/// </summary>
	public static TableRow Height(this TableRow row, TableRowHeight? height)
	{
		row.GetOrInit<TableRowProperties>().SetPropertyClassOrRemove(height);
		return row;
	}

	/// <summary>
	/// Specifies the preferred width for the total number of grid columns
	/// after this table row as specified in the gridAfter element (§17.4.14).
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthAfterTableRow" /></para>
	/// </summary>
	public static TableRow WidthAfter(this TableRow row, WidthAfterTableRow? width)
	{
		row.GetOrInit<TableRowProperties>().SetPropertyClassOrRemove(width);
		return row;
	}

	/// <summary>
	/// Specifies the preferred width for the total number of grid columns
	/// before this table row as specified in the gridAfter element (§17.4.14).
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthBeforeTableRow" /></para>
	/// </summary>
	public static TableRow WidthBefore(this TableRow row, WidthBeforeTableRow? width)
	{
		row.GetOrInit<TableRowProperties>().SetPropertyClassOrRemove(width);
		return row;
	}

	#endregion Set property methods
}
