using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="Table" /> class.
/// </summary>
public static class TableExtensions
{
	/// <summary>
	/// Gets the <see cref="TableProperties" /> object of the <see cref="Table" />.
	/// <para>Returns <see langword="null" /> if the node doesn't exist.</para>
	/// </summary>
	public static TableProperties? GetTableProperties(this Table table)
		=> table.GetFirstChild<TableProperties>();

	#region Get property methods

	/// <summary>
	/// Specifies that the cells with this table shall be visually
	/// represented in a right to left direction.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.BiDiVisual" /></para>
	/// </summary>
	public static bool? VisuallyBiDirectionalValue(this Table table)
	{
		OnOffOnlyValues? value = table.GetTableProperties()?.BiDiVisual?.Val?.Value;
		return value is null ? null : value is OnOffOnlyValues.On;
	}

	/// <summary>
	/// Specifies the alignment of the set of rows which are part of the current
	/// table properties exception list with respect to the text margins in
	/// the current section.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableJustification" /></para>
	/// </summary>
	public static TableRowAlignmentValues? JustificationValue(this Table table)
		=> table.GetTableProperties()?.TableJustification?.Val?.Value;

	/// <summary>
	/// Specifies the shading applied to the contents of the table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Shading? GetShading(this Table table)
		=> table.GetTableProperties()?.Shading;

	/// <summary>
	/// Specifies the set of borders for the edges of the current table,
	/// using the six border types defined by its child elements.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableBorders" /></para>
	/// </summary>
	public static TableBorders? GetBorders(this Table table)
		=> table.GetTableProperties()?.TableBorders;

	/// <summary>
	/// Specifies the caption for the table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCaption" /></para>
	/// </summary>
	public static string? CaptionValue(this Table table)
		=> table.GetTableProperties()?.TableCaption?.Val;

	/// <summary>
	/// Specifies a set of cell margins for all cells in the parent table row
	/// via a set of table-level property exceptions.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault" /></para>
	/// </summary>
	public static TableCellMarginDefault? GetDefaultCellMargins(this Table table)
		=> table.GetTableProperties()?.TableCellMarginDefault;

	/// <summary>
	/// Specifies the default table cell spacing (the spacing between adjacent
	/// cells and the edges of the table) for all cells in the parent row.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellSpacing" /></para>
	/// </summary>
	public static TableCellSpacing? GetCellSpacing(this Table table)
		=> table.GetTableProperties()?.TableCellSpacing;

	/// <summary>
	/// Specifies the description for the table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableDescription" /></para>
	/// </summary>
	public static string? DescriptionValue(this Table table)
		=> table.GetTableProperties()?.TableDescription?.Val;

	/// <summary>
	/// Specifies the indentation which shall be added before the leading edge of the current
	/// table in the document (the left edge in a left-to-right table, and the right edge
	/// in a right-to-left table).
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableIndentation" /></para>
	/// </summary>
	public static TableIndentation? GetIndentation(this Table table)
		=> table.GetTableProperties()?.TableIndentation;

	/// <summary>
	/// Specifies the algorithm which shall be used to lay out
	/// the contents of this table within the document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableLayout" /></para>
	/// </summary>
	public static TableLayoutValues? LayoutValue(this Table table)
		=> table.GetTableProperties()?.TableLayout?.Type?.Value;

	/// <summary>
	/// Specifies the components of the conditional formatting of the referenced
	/// table style (if one exists) which shall be applied to the set of table rows
	/// with the current table-level property exceptions.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableLook" /></para>
	/// </summary>
	public static TableLook? GetLook(this Table table)
		=> table.GetTableProperties()?.TableLook;

	/// <summary>
	/// Specifies whether the current table shall allow other floating tables to overlap
	/// its extents when the tables are displayed in a document.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableOverlap" /></para>
	/// </summary>
	public static TableOverlapValues? OverlapValue(this Table table)
		=> table.GetTableProperties()?.TableOverlap?.Val?.Value;

	/// <summary>
	/// Specifies information about the current table with regard to floating tables.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TablePositionProperties" /></para>
	/// </summary>
	public static TablePositionProperties? GetPositionProperties(this Table table)
		=> table.GetTableProperties()?.TablePositionProperties;

	/// <summary>
	/// Specifies the style ID of the table style which shall be used
	/// to format the contents of this table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyle" /></para>
	/// </summary>
	public static string? StyleValue(this Table table)
		=> table.GetTableProperties()?.TableStyle?.Val;

	/// <summary>
	/// Specifies the number of columns which shall
	/// comprise each a table style column band for this table style.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyleColumnBandSize" /></para>
	/// </summary>
	public static int? StyleColumnBandSizeValue(this Table table)
		=> table.GetTableProperties()?.GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value;

	/// <summary>
	/// Specifies the number of rows which shall
	/// comprise each a table style row band for this table style.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyleRowBandSize" /></para>
	/// </summary>
	public static int? StyleRowBandSizeValue(this Table table)
		=> table.GetTableProperties()?.GetFirstChild<TableStyleRowBandSize>()?.Val?.Value;

	/// <summary>
	/// Specifies the preferred width for this table.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableWidth" /></para>
	/// </summary>
	public static TableWidth? GetWidth(this Table table)
		=> table.GetTableProperties()?.TableWidth;

	#endregion Get property methods

	#region Set property methods

	/// <summary>
	/// Specifies that the cells with this table shall be visually
	/// represented in a right to left direction.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.BiDiVisual" /></para>
	/// </summary>
	public static Table VisuallyBiDirectional(this Table table, bool? value = true)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<BiDiVisual>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return table;
	}

	/// <summary>
	/// Specifies the alignment of the set of rows which are part of the current
	/// table properties exception list with respect to the text margins in
	/// the current section.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableJustification" /></para>
	/// </summary>
	public static Table Justification(this Table table, TableRowAlignmentValues? alignment)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableJustification>(alignment);
		return table;
	}

	/// <summary>
	/// Specifies the shading applied to the contents of the table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Table Shading(this Table table, Shading? shading)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(shading);
		return table;
	}

	/// <summary>
	/// Specifies the set of borders for the edges of the current table,
	/// using the six border types defined by its child elements.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableBorders" /></para>
	/// </summary>
	public static Table Borders(this Table table, TableBorders? borders)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(borders);
		return table;
	}

	/// <summary>
	/// Specifies the caption for the table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCaption" /></para>
	/// </summary>
	public static Table Caption(this Table table, string? caption)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableCaption>(caption);
		return table;
	}

	/// <summary>
	/// Specifies a set of cell margins for all cells in the parent table row
	/// via a set of table-level property exceptions.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault" /></para>
	/// </summary>
	public static Table DefaultCellMargins(this Table table, TableCellMarginDefault? margin)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Specifies the default table cell spacing (the spacing between adjacent
	/// cells and the edges of the table) for all cells in the parent row.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellSpacing" /></para>
	/// </summary>
	public static Table CellSpacing(this Table table, TableCellSpacing? spacing)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(spacing);
		return table;
	}

	/// <summary>
	/// Specifies the description for the table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableDescription" /></para>
	/// </summary>
	public static Table Description(this Table table, string? description)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableDescription>(description);
		return table;
	}

	/// <summary>
	/// Specifies the indentation which shall be added before the leading edge of the current
	/// table in the document (the left edge in a left-to-right table, and the right edge
	/// in a right-to-left table).
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableIndentation" /></para>
	/// </summary>
	public static Table Indentation(this Table table, TableIndentation? indentation)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(indentation);
		return table;
	}

	/// <summary>
	/// Specifies the algorithm which shall be used to lay out
	/// the contents of this table within the document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableLayout" /></para>
	/// </summary>
	public static Table Layout(this Table table, TableLayoutValues? layout)
	{
		table.GetOrInit<TableProperties>(true).SetFieldOrRemove<TableLayout>("Type", layout);
		return table;
	}

	/// <summary>
	/// Specifies the components of the conditional formatting of the referenced
	/// table style (if one exists) which shall be applied to the set of table rows
	/// with the current table-level property exceptions.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableLook" /></para>
	/// </summary>
	public static Table Look(this Table table, TableLook? look)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(look);
		return table;
	}

	/// <summary>
	/// Specifies whether the current table shall allow other floating tables to overlap
	/// its extents when the tables are displayed in a document.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableOverlap" /></para>
	/// </summary>
	public static Table Overlap(this Table table, TableOverlapValues? overlap)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableOverlap>(overlap);
		return table;
	}

	/// <summary>
	/// Specifies information about the current table with regard to floating tables.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TablePositionProperties" /></para>
	/// </summary>
	public static Table PositionProperties(this Table table, TablePositionProperties? properties)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(properties);
		return table;
	}

	/// <summary>
	/// Specifies the style ID of the table style which shall be used
	/// to format the contents of this table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyle" /></para>
	/// </summary>
	public static Table Style(this Table table, string? styleId)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableStyle>(styleId);
		return table;
	}

	/// <summary>
	/// Specifies the number of columns which shall
	/// comprise each a table style column band for this table style.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyleColumnBandSize" /></para>
	/// </summary>
	public static Table StyleColumnBandSize(this Table table, int? bandSize)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableStyleColumnBandSize>(bandSize);
		return table;
	}

	/// <summary>
	/// Specifies the number of rows which shall
	/// comprise each a table style row band for this table style.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyleRowBandSize" /></para>
	/// </summary>
	public static Table StyleRowBandSize(this Table table, int? bandSize)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableStyleRowBandSize>(bandSize);
		return table;
	}

	/// <summary>
	/// Specifies the preferred width for this table.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableWidth" /></para>
	/// </summary>
	public static Table Width(this Table table, TableWidth? width)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(width);
		return table;
	}

	#endregion Set property methods
}
