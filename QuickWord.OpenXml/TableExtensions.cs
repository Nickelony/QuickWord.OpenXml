using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

public static class TableExtensions
{
	public static TableProperties? GetTableProperties(this Table table)
		=> table.GetFirstChild<TableProperties>();

	#region Get property methods

	/// <summary>
	/// Specifies that the cells with this table shall be visually
	/// represented in a right to left direction.
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
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableJustification" /></para>
	/// </summary>
	public static TableRowAlignmentValues? JustificationValue(this Table table)
		=> table.GetTableProperties()?.TableJustification?.Val?.Value;

	/// <summary>
	/// Specifies the shading applied to the contents of the table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Shading? GetShading(this Table table)
		=> table.GetTableProperties()?.Shading;

	/// <summary>
	/// Specifies the set of borders for the edges of the current table,
	/// using the six border types defined by its child elements.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableBorders" /></para>
	/// </summary>
	public static TableBorders? GetBorders(this Table table)
		=> table.GetTableProperties()?.TableBorders;

	/// <summary>
	/// Specifies the caption for the table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCaption" /></para>
	/// </summary>
	public static string? CaptionValue(this Table table)
		=> table.GetTableProperties()?.TableCaption?.Val;

	/// <summary>
	/// Specifies a set of cell margins for all cells in the parent table row
	/// via a set of table-level property exceptions.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellMarginDefault" /></para>
	/// </summary>
	public static TableCellMarginDefault? GetDefaultCellMargins(this Table table)
		=> table.GetTableProperties()?.TableCellMarginDefault;

	/// <summary>
	/// Specifies the default table cell spacing (the spacing between adjacent
	/// cells and the edges of the table) for all cells in the parent row.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableCellSpacing" /></para>
	/// </summary>
	public static TableCellSpacing? GetCellSpacing(this Table table)
		=> table.GetTableProperties()?.TableCellSpacing;

	/// <summary>
	/// Specifies the description for the table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableDescription" /></para>
	/// </summary>
	public static string? DescriptionValue(this Table table)
		=> table.GetTableProperties()?.TableDescription?.Val;

	/// <summary>
	/// Specifies the indentation which shall be added before the leading edge of the current
	/// table in the document (the left edge in a left-to-right table, and the right edge
	/// in a right-to-left table).
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableIndentation" /></para>
	/// </summary>
	public static TableIndentation? GetIndentation(this Table table)
		=> table.GetTableProperties()?.TableIndentation;

	/// <summary>
	/// Specifies the algorithm which shall be used to lay out
	/// the contents of this table within the document.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableLayout" /></para>
	/// </summary>
	public static TableLayoutValues? LayoutValue(this Table table)
		=> table.GetTableProperties()?.TableLayout?.Type?.Value;

	/// <summary>
	/// Specifies the components of the conditional formatting of the referenced
	/// table style (if one exists) which shall be applied to the set of table rows
	/// with the current table-level property exceptions.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableLook" /></para>
	/// </summary>
	public static TableLook? GetLook(this Table table)
		=> table.GetTableProperties()?.TableLook;

	/// <summary>
	/// Specifies whether the current table shall allow other floating tables to overlap
	/// its extents when the tables are displayed in a document.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableOverlap" /></para>
	/// </summary>
	public static TableOverlapValues? OverlapValue(this Table table)
		=> table.GetTableProperties()?.TableOverlap?.Val?.Value;

	/// <summary>
	/// Specifies information about the current table with regard to floating tables.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TablePositionProperties" /></para>
	/// </summary>
	public static TablePositionProperties? GetPositionProperties(this Table table)
		=> table.GetTableProperties()?.TablePositionProperties;

	/// <summary>
	/// Specifies the style ID of the table style which shall be used
	/// to format the contents of this table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyle" /></para>
	/// </summary>
	public static string? StyleValue(this Table table)
		=> table.GetTableProperties()?.TableStyle?.Val;

	/// <summary>
	/// Specifies the number of columns which shall
	/// comprise each a table style column band for this table style.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyleColumnBandSize" /></para>
	/// </summary>
	public static int? StyleColumnBandSizeValue(this Table table)
		=> table.GetTableProperties()?.GetFirstChild<TableStyleColumnBandSize>()?.Val?.Value;

	/// <summary>
	/// Specifies the number of rows which shall
	/// comprise each a table style row band for this table style.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableStyleRowBandSize" /></para>
	/// </summary>
	public static int? StyleRowBandSizeValue(this Table table)
		=> table.GetTableProperties()?.GetFirstChild<TableStyleRowBandSize>()?.Val?.Value;

	/// <summary>
	/// Specifies the preferred width for this table.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableWidth" /></para>
	/// </summary>
	public static TableWidth? GetWidth(this Table table)
		=> table.GetTableProperties()?.TableWidth;

	#endregion Get property methods

	#region Set property methods

	/// <inheritdoc cref="VisuallyBiDirectionalValue" />
	public static Table VisuallyBiDirectional(this Table table, bool? value = true)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<BiDiVisual>(value is null
			? null
			: value is true ? OnOffOnlyValues.On : OnOffOnlyValues.Off);

		return table;
	}

	/// <inheritdoc cref="JustificationValue" />
	public static Table Justification(this Table table, TableRowAlignmentValues? alignment)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableJustification>(alignment);
		return table;
	}

	/// <inheritdoc cref="GetShading" />
	public static Table Shading(this Table table, Shading? shading)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(shading);
		return table;
	}

	/// <inheritdoc cref="GetBorders" />
	public static Table Borders(this Table table, TableBorders? borders)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(borders);
		return table;
	}

	/// <inheritdoc cref="CaptionValue" />
	public static Table Caption(this Table table, string? caption)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableCaption>(caption);
		return table;
	}

	/// <inheritdoc cref="GetDefaultCellMargins" />
	public static Table DefaultCellMargins(this Table table, TableCellMarginDefault? margin)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <inheritdoc cref="GetCellSpacing" />
	public static Table CellSpacing(this Table table, TableCellSpacing? spacing)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(spacing);
		return table;
	}

	/// <inheritdoc cref="DescriptionValue" />
	public static Table Description(this Table table, string? description)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableDescription>(description);
		return table;
	}

	/// <inheritdoc cref="GetIndentation" />
	public static Table Indentation(this Table table, TableIndentation? indentation)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(indentation);
		return table;
	}

	/// <inheritdoc cref="LayoutValue" />
	public static Table Layout(this Table table, TableLayoutValues? layout)
	{
		table.GetOrInit<TableProperties>(true).SetFieldOrRemove<TableLayout>("Type", layout);
		return table;
	}

	/// <inheritdoc cref="GetLook" />
	public static Table Look(this Table table, TableLook? look)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(look);
		return table;
	}

	/// <inheritdoc cref="OverlapValue" />
	public static Table Overlap(this Table table, TableOverlapValues? overlap)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableOverlap>(overlap);
		return table;
	}

	/// <inheritdoc cref="GetPositionProperties" />
	public static Table PositionProperties(this Table table, TablePositionProperties? properties)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(properties);
		return table;
	}

	/// <inheritdoc cref="StyleValue" />
	public static Table Style(this Table table, string? styleId)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableStyle>(styleId);
		return table;
	}

	/// <inheritdoc cref="StyleColumnBandSizeValue" />
	public static Table StyleColumnBandSize(this Table table, int? bandSize)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableStyleColumnBandSize>(bandSize);
		return table;
	}

	/// <inheritdoc cref="StyleRowBandSizeValue" />
	public static Table StyleRowBandSize(this Table table, int? bandSize)
	{
		table.GetOrInit<TableProperties>(true).SetValOrRemove<TableStyleRowBandSize>(bandSize);
		return table;
	}

	/// <inheritdoc cref="GetWidth" />
	public static Table Width(this Table table, TableWidth? width)
	{
		table.GetOrInit<TableProperties>(true).SetPropertyClassOrRemove(width);
		return table;
	}

	#endregion Set property methods
}
