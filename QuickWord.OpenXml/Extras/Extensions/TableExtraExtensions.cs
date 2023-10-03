using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;
using System.Collections.Generic;
using System.Linq;

namespace QuickWord.OpenXml.Extras;

/// <summary>
/// Additional extension methods for the <see cref="Table"/> class.
/// </summary>
public static class TableExtraExtensions
{
	public static IEnumerable<TableRow> Rows(this Table table)
		=> table.Elements<TableRow>();

	public static TableRow? Rows(this Table table, int index)
		=> table.Elements<TableRow>().ElementAtOrDefault(index);

	public static IEnumerable<TableCell?> GetColumnOfCells(this Table table, int cellIndex)
		=> table.Rows().Select(r => r.Cells(cellIndex));

	#region Formatting

	/// <summary>
	/// Clones all formatting properties from the table.
	/// </summary>
	public static TableFormatting CloneFormatting(this Table table) => new()
	{
		VisuallyBiDirectional = table.VisuallyBiDirectionalValue(),
		Justification = table.JustificationValue(),
		Shading = table.GetShading()?.CloneNode(true) as Shading,
		Borders = table.GetBorders()?.CloneNode(true) as TableBorders,
		Caption = table.CaptionValue(),
		DefaultCellMargins = table.GetDefaultCellMargins()?.CloneNode(true) as TableCellMarginDefault,
		CellSpacing = table.GetCellSpacing()?.CloneNode(true) as TableCellSpacing,
		Description = table.DescriptionValue(),
		Indentation = table.GetIndentation()?.CloneNode(true) as TableIndentation,
		Layout = table.LayoutValue(),
		Look = table.GetLook()?.CloneNode(true) as TableLook,
		Overlap = table.OverlapValue(),
		PositionProperties = table.GetPositionProperties()?.CloneNode(true) as TablePositionProperties,
		Style = table.StyleValue(),
		StyleColumnBandSize = table.StyleColumnBandSizeValue(),
		StyleRowBandSize = table.StyleRowBandSizeValue(),
		Width = table.GetWidth()?.CloneNode(true) as TableWidth
	};

	/// <summary>
	/// Applies the given formatting properties to the table (replaces every possible property of the table unless <c>ignoreNulls</c> is set to true).
	/// </summary>
	public static Table ApplyFormatting(this Table table, TableFormatting formatting, bool ignoreNulls = false)
	{
		if (formatting.VisuallyBiDirectional is not null || (formatting.VisuallyBiDirectional is null && !ignoreNulls))
			table.VisuallyBiDirectional(formatting.VisuallyBiDirectional);

		if (formatting.Justification is not null || (formatting.Justification is null && !ignoreNulls))
			table.Justification(formatting.Justification);

		if (formatting.Shading is not null || (formatting.Shading is null && !ignoreNulls))
			table.Shading(formatting.Shading?.CloneNode(true) as Shading);

		if (formatting.Borders is not null || (formatting.Borders is null && !ignoreNulls))
			table.Borders(formatting.Borders?.CloneNode(true) as TableBorders);

		if (formatting.Caption is not null || (formatting.Caption is null && !ignoreNulls))
			table.Caption(formatting.Caption);

		if (formatting.DefaultCellMargins is not null || (formatting.DefaultCellMargins is null && !ignoreNulls))
			table.DefaultCellMargins(formatting.DefaultCellMargins?.CloneNode(true) as TableCellMarginDefault);

		if (formatting.CellSpacing is not null || (formatting.CellSpacing is null && !ignoreNulls))
			table.CellSpacing(formatting.CellSpacing?.CloneNode(true) as TableCellSpacing);

		if (formatting.Description is not null || (formatting.Description is null && !ignoreNulls))
			table.Description(formatting.Description);

		if (formatting.Indentation is not null || (formatting.Indentation is null && !ignoreNulls))
			table.Indentation(formatting.Indentation?.CloneNode(true) as TableIndentation);

		if (formatting.Layout is not null || (formatting.Layout is null && !ignoreNulls))
			table.Layout(formatting.Layout);

		if (formatting.Look is not null || (formatting.Look is null && !ignoreNulls))
			table.Look(formatting.Look?.CloneNode(true) as TableLook);

		if (formatting.Overlap is not null || (formatting.Overlap is null && !ignoreNulls))
			table.Overlap(formatting.Overlap);

		if (formatting.PositionProperties is not null || (formatting.PositionProperties is null && !ignoreNulls))
			table.PositionProperties(formatting.PositionProperties?.CloneNode(true) as TablePositionProperties);

		if (formatting.Style is not null || (formatting.Style is null && !ignoreNulls))
			table.Style(formatting.Style);

		if (formatting.StyleColumnBandSize is not null || (formatting.StyleColumnBandSize is null && !ignoreNulls))
			table.StyleColumnBandSize(formatting.StyleColumnBandSize);

		if (formatting.StyleRowBandSize is not null || (formatting.StyleRowBandSize is null && !ignoreNulls))
			table.StyleRowBandSize(formatting.StyleRowBandSize);

		if (formatting.Width is not null || (formatting.Width is null && !ignoreNulls))
			table.Width(formatting.Width?.CloneNode(true) as TableWidth);

		return table;
	}

	/// <summary>
	/// Resets every possible property of the table.
	/// </summary>
	public static Table ResetFormatting(this Table table)
	{
		table.RemoveAllChildren<TableProperties>();
		return table;
	}

	#endregion Formatting

	#region Borders

	/// <summary>
	/// Sets the borders of the table.
	/// </summary>
	public static Table Borders(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table
			.LeftBorder(width, type, htmlColor, spacing)
			.TopBorder(width, type, htmlColor, spacing)
			.RightBorder(width, type, htmlColor, spacing)
			.BottomBorder(width, type, htmlColor, spacing)
			.InsideHorizontalBorder(width, type, htmlColor, spacing)
			.InsideVerticalBorder(width, type, htmlColor, spacing);

		return table;
	}

	/// <summary>
	/// Sets the left border of the table.
	/// </summary>
	public static Table LeftBorder(this Table table, LeftBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="LeftBorder" />
	public static Table LeftBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().LeftBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the top border of the table.
	/// </summary>
	public static Table TopBorder(this Table table, TopBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="TopBorder" />
	public static Table TopBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().TopBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the right border of the table.
	/// </summary>
	public static Table RightBorder(this Table table, RightBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="RightBorder" />
	public static Table RightBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().RightBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the bottom border of the table.
	/// </summary>
	public static Table BottomBorder(this Table table, BottomBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="BottomBorder" />
	public static Table BottomBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().BottomBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the inside-horizontal border of the table.
	/// </summary>
	public static Table InsideHorizontalBorder(this Table table, InsideHorizontalBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="InsideHorizontalBorder" />
	public static Table InsideHorizontalBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().InsideHorizontalBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the inside-vertical border of the table.
	/// </summary>
	public static Table InsideVerticalBorder(this Table table, InsideVerticalBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="InsideVerticalBorder" />
	public static Table InsideVerticalBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().InsideVerticalBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the start border of the table.
	/// </summary>
	public static Table StartBorder(this Table table, StartBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="StartBorder" />
	public static Table StartBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().StartBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	/// <summary>
	/// Sets the end border of the table.
	/// </summary>
	public static Table EndBorder(this Table table, EndBorder? border)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().SetPropertyClassOrRemove(border);
		return table;
	}

	/// <inheritdoc cref="EndBorder" />
	public static Table EndBorder(this Table table, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableBorders>().EndBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return table;
	}

	#endregion Borders

	#region DefaultMarginsOfCells

	// Get:

	/// <summary>
	/// Gets the default left margin of the table's cells (in points).
	/// </summary>
	public static double? DefaultLeftMarginOfCellsValue(this Table table)
	{
		short? marginWidth = table.GetDefaultCellMargins()?.TableCellLeftMargin?.Width?.Value;
		return marginWidth is null ? null : marginWidth / 20.0;
	}

	/// <summary>
	/// Gets the default top margin of the table's cells and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? DefaultTopMarginOfCellsValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetDefaultCellMargins()?.TopMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the default right margin of the table's cells (in points).
	/// </summary>
	public static double? DefaultRightMarginOfCellsValue(this Table table)
	{
		short? marginWidth = table.GetDefaultCellMargins()?.TableCellRightMargin?.Width?.Value;
		return marginWidth is null ? null : marginWidth / 20.0;
	}

	/// <summary>
	/// Gets the default bottom margin of the table's cells and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? DefaultBottomMarginOfCellsValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetDefaultCellMargins()?.BottomMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the default start margin of the table's cells and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? DefaultStartMarginOfCellsValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetDefaultCellMargins()?.StartMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the default end margin of the table's cells and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? DefaultEndMarginOfCellsValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetDefaultCellMargins()?.EndMargin?.GetExactWidth(out units);
	}

	// Set:

	/// <summary>
	/// Sets the default left margin of the table's cells.
	/// </summary>
	public static Table DefaultLeftMarginOfCells(this Table table, TableCellLeftMargin? margin)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Sets the default left margin of the table's cells (in points).
	/// </summary>
	public static Table DefaultLeftMarginOfCells(this Table table, double width)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().TableCellLeftMargin = new()
		{
			Width = (short)(width * 20),
			Type = TableWidthValues.Dxa
		};

		return table;
	}

	/// <summary>
	/// Sets the default top margin of the table's cells.
	/// </summary>
	public static Table DefaultTopMarginOfCells(this Table table, TopMargin? margin)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Sets the default top margin of the table's cells in the given units.
	/// </summary>
	public static Table DefaultTopMarginOfCells(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().GetOrInit<TopMargin>().SetExactWidth(width, units);
		return table;
	}

	/// <summary>
	/// Sets the default right margin of the table's cells.
	/// </summary>
	public static Table DefaultRightMarginOfCells(this Table table, TableCellRightMargin? margin)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Sets the default right margin of the table's cells (in points).
	/// </summary>
	public static Table DefaultRightMarginOfCells(this Table table, double width)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().TableCellRightMargin = new()
		{
			Width = (short)(width * 20),
			Type = TableWidthValues.Dxa
		};

		return table;
	}

	/// <summary>
	/// Sets the default bottom margin of the table's cells.
	/// </summary>
	public static Table DefaultBottomMarginOfCells(this Table table, BottomMargin? margin)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Sets the default bottom margin of the table's cells in the given units.
	/// </summary>
	public static Table DefaultBottomMarginOfCells(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().GetOrInit<BottomMargin>().SetExactWidth(width, units);
		return table;
	}

	/// <summary>
	/// Sets the default start margin of the table's cells.
	/// </summary>
	public static Table DefaultStartMarginOfCells(this Table table, StartMargin? margin)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Sets the default start margin of the table's cells in the given units.
	/// </summary>
	public static Table DefaultStartMarginOfCells(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().GetOrInit<StartMargin>().SetExactWidth(width, units);
		return table;
	}

	/// <summary>
	/// Sets the default end margin of the table's cells.
	/// </summary>
	public static Table DefaultEndMarginOfCells(this Table table, EndMargin? margin)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().SetPropertyClassOrRemove(margin);
		return table;
	}

	/// <summary>
	/// Sets the default end margin of the table's cells in the given units.
	/// </summary>
	public static Table DefaultEndMarginOfCells(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellMarginDefault>().GetOrInit<EndMargin>().SetExactWidth(width, units);
		return table;
	}

	#endregion DefaultMarginsOfCells

	#region Other

	// Get:

	/// <summary>
	/// Gets the cell spacing of the table and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? CellSpacingValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetCellSpacing()?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the indentation of the table and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? IndentationValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetIndentation()?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the width of the table and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? WidthValue(this Table table, out WidthUnits? units)
	{
		units = null;
		return table.GetWidth()?.GetExactWidth(out units);
	}

	// Set:

	/// <summary>
	/// Sets the cell spacing of the table in the given units.
	/// </summary>
	public static Table CellSpacing(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableCellSpacing>().SetExactWidth(width, units);
		return table;
	}

	/// <summary>
	/// Sets the indentation of the table in the given units.
	/// </summary>
	public static Table Indentation(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableIndentation>().SetExactWidth(width, units);
		return table;
	}

	/// <summary>
	/// Sets the width of the table in the given units.
	/// </summary>
	public static Table Width(this Table table, double width, WidthUnits? units)
	{
		table.GetOrInit<TableProperties>(true).GetOrInit<TableWidth>().SetExactWidth(width, units);
		return table;
	}

	/// <summary>
	/// Sets the background fill color of the table.
	/// <para><c>Fill</c> property of <see cref="Shading" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static Table FillColor(this Table table, string htmlColor)
	{
		Shading shading = table.GetOrInit<TableProperties>(true).GetOrInit<Shading>();
		shading.Fill = htmlColor;
		shading.Val = shading.Val?.Value is null or ShadingPatternValues.Nil
			? ShadingPatternValues.Clear
			: shading.Val.Value; // Preserve original pattern

		return table;
	}

	#endregion Other
}
