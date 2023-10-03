using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;
using System.Collections.Generic;
using System.Linq;

namespace QuickWord.OpenXml.Extras;

/// <summary>
/// Additional extension methods for the <see cref="TableCell"/> class.
/// </summary>
public static class TableCellExtraExtensions
{
	public static IEnumerable<Paragraph> Paragraphs(this TableCell cell)
		=> cell.Elements<Paragraph>();

	public static Paragraph? Paragraphs(this TableCell cell, int index)
		=> cell.Elements<Paragraph>().ElementAtOrDefault(index);

	#region Formatting

	/// <summary>
	/// Clones all formatting properties from the table cell.
	/// </summary>
	public static TableCellFormatting CloneFormatting(this TableCell cell) => new()
	{
		ConditionalFormatStyle = cell.GetConditionalFormatStyle()?.CloneNode(true) as ConditionalFormatStyle,
		GridSpan = cell.GridSpanValue(),
		HideEndOfCellMarker = cell.HideEndOfCellMarkerValue(),
		HorizontalMerge = cell.HorizontalMergeValue(),
		NoContentWrapping = cell.NoContentWrappingValue(),
		Shading = cell.GetShading()?.CloneNode(true) as Shading,
		Borders = cell.GetBorders()?.CloneNode(true) as TableCellBorders,
		FitText = cell.FitTextValue(),
		Margins = cell.GetMargins()?.CloneNode(true) as TableCellMargin,
		Width = cell.GetWidth()?.CloneNode(true) as TableCellWidth,
		TextDirection = cell.TextDirectionValue(),
		VerticalContentAlignment = cell.VerticalContentAlignmentValue(),
		VerticalMerge = cell.VerticalMergeValue()
	};

	/// <summary>
	/// Applies the given formatting properties to the table cell (replaces every possible property of the table cell unless <c>ignoreNulls</c> is set to true).
	/// </summary>
	public static TableCell ApplyFormatting(this TableCell cell, TableCellFormatting formatting, bool ignoreNulls = false)
	{
		if (formatting.ConditionalFormatStyle is not null || (formatting.ConditionalFormatStyle is null && !ignoreNulls))
			cell.ConditionalFormatStyle(formatting.ConditionalFormatStyle);

		if (formatting.GridSpan is not null || (formatting.GridSpan is null && !ignoreNulls))
			cell.GridSpan(formatting.GridSpan);

		if (formatting.HideEndOfCellMarker is not null || (formatting.HideEndOfCellMarker is null && !ignoreNulls))
			cell.HideEndOfCellMarker(formatting.HideEndOfCellMarker);

		if (formatting.HorizontalMerge is not null || (formatting.HorizontalMerge is null && !ignoreNulls))
			cell.HorizontalMerge(formatting.HorizontalMerge);

		if (formatting.NoContentWrapping is not null || (formatting.NoContentWrapping is null && !ignoreNulls))
			cell.NoContentWrapping(formatting.NoContentWrapping);

		if (formatting.Shading is not null || (formatting.Shading is null && !ignoreNulls))
			cell.Shading(formatting.Shading?.CloneNode(true) as Shading);

		if (formatting.Borders is not null || (formatting.Borders is null && !ignoreNulls))
			cell.Borders(formatting.Borders?.CloneNode(true) as TableCellBorders);

		if (formatting.FitText is not null || (formatting.FitText is null && !ignoreNulls))
			cell.FitText(formatting.FitText);

		if (formatting.Margins is not null || (formatting.Margins is null && !ignoreNulls))
			cell.Margin(formatting.Margins?.CloneNode(true) as TableCellMargin);

		if (formatting.Width is not null || (formatting.Width is null && !ignoreNulls))
			cell.Width(formatting.Width?.CloneNode(true) as TableCellWidth);

		if (formatting.TextDirection is not null || (formatting.TextDirection is null && !ignoreNulls))
			cell.TextDirection(formatting.TextDirection);

		if (formatting.VerticalContentAlignment is not null || (formatting.VerticalContentAlignment is null && !ignoreNulls))
			cell.VerticalContentAlignment(formatting.VerticalContentAlignment);

		if (formatting.VerticalMerge is not null || (formatting.VerticalMerge is null && !ignoreNulls))
			cell.VerticalMerge(formatting.VerticalMerge);

		return cell;
	}

	/// <summary>
	/// Resets every possible property of the table cell.
	/// </summary>
	public static TableCell ResetFormatting(this TableCell cell)
	{
		cell.RemoveAllChildren<TableCellProperties>();
		return cell;
	}

	#endregion Formatting

	#region Borders

	/// <summary>
	/// Sets the left border of the cell.
	/// </summary>
	public static TableCell LeftBorder(this TableCell cell, LeftBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="LeftBorder" />
	public static TableCell LeftBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().LeftBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the top border of the cell.
	/// </summary>
	public static TableCell TopBorder(this TableCell cell, TopBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="TopBorder" />
	public static TableCell TopBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().TopBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the right border of the cell.
	/// </summary>
	public static TableCell RightBorder(this TableCell cell, RightBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="RightBorder" />
	public static TableCell RightBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().RightBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the bottom border of the cell.
	/// </summary>
	public static TableCell BottomBorder(this TableCell cell, BottomBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="BottomBorder" />
	public static TableCell BottomBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().BottomBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the inside-horizontal border of the cell.
	/// </summary>
	public static TableCell InsideHorizontalBorder(this TableCell cell, InsideHorizontalBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="InsideHorizontalBorder" />
	public static TableCell InsideHorizontalBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().InsideHorizontalBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the inside-vertical border of the cell.
	/// </summary>
	public static TableCell InsideVerticalBorder(this TableCell cell, InsideVerticalBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="InsideVerticalBorder" />
	public static TableCell InsideVerticalBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().InsideVerticalBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the top-left to bottom-right diagonal border of the cell.
	/// </summary>
	public static TableCell TopLeftToBottomRightCellBorder(this TableCell cell, TopLeftToBottomRightCellBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="TopLeftToBottomRightCellBorder" />
	public static TableCell TopLeftToBottomRightCellBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().TopLeftToBottomRightCellBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the top-right to bottom-left diagonal border of the cell.
	/// </summary>
	public static TableCell TopRightToBottomLeftCellBorder(this TableCell cell, TopRightToBottomLeftCellBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="TopRightToBottomLeftCellBorder" />
	public static TableCell TopRightToBottomLeftCellBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().TopRightToBottomLeftCellBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the start border of the cell.
	/// </summary>
	public static TableCell StartBorder(this TableCell cell, StartBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="StartBorder" />
	public static TableCell StartBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().StartBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	/// <summary>
	/// Sets the end border of the cell.
	/// </summary>
	public static TableCell EndBorder(this TableCell cell, EndBorder? border)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().SetPropertyClassOrRemove(border);
		return cell;
	}

	/// <inheritdoc cref="EndBorder" />
	public static TableCell EndBorder(this TableCell cell, double width, BorderValues type = BorderValues.Single, string htmlColor = "auto", uint spacing = 0)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellBorders>().EndBorder = new()
		{
			Size = BorderSize.ToSixth(width),
			Val = type,
			Color = htmlColor,
			Space = spacing
		};

		return cell;
	}

	#endregion Borders

	#region Margins

	// Get:

	/// <summary>
	/// Gets the left margin size of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? LeftMarginValue(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.GetMargins()?.LeftMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the top margin size of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? TopMarginValue(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.GetMargins()?.TopMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the right margin size of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? RightMarginValue(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.GetMargins()?.RightMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the bottom margin size of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? BottomMargin(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.GetMargins()?.BottomMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the start margin size of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? StartMarginValue(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.GetMargins()?.StartMargin?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the end margin size of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? EndMarginValue(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.GetMargins()?.EndMargin?.GetExactWidth(out units);
	}

	// Set:

	/// <summary>
	/// Sets the left margin of the cell.
	/// </summary>
	public static TableCell LeftMargin(this TableCell cell, LeftMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Sets the left margin of the cell in the given units.
	/// </summary>
	public static TableCell LeftMargin(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().GetOrInit<LeftMargin>().SetExactWidth(width, units);
		return cell;
	}

	/// <summary>
	/// Sets the top margin of the cell.
	/// </summary>
	public static TableCell TopMargin(this TableCell cell, TopMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Sets the top margin of the cell in the given units.
	/// </summary>
	public static TableCell TopMargin(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().GetOrInit<TopMargin>().SetExactWidth(width, units);
		return cell;
	}

	/// <summary>
	/// Sets the right margin of the cell.
	/// </summary>
	public static TableCell RightMargin(this TableCell cell, RightMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Sets the right margin of the cell in the given units.
	/// </summary>
	public static TableCell RightMargin(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().GetOrInit<RightMargin>().SetExactWidth(width, units);
		return cell;
	}

	/// <summary>
	/// Sets the bottom margin of the cell.
	/// </summary>
	public static TableCell BottomMargin(this TableCell cell, BottomMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Sets the bottom margin of the cell in the given units.
	/// </summary>
	public static TableCell BottomMargin(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().GetOrInit<BottomMargin>().SetExactWidth(width, units);
		return cell;
	}

	/// <summary>
	/// Sets the start margin of the cell.
	/// </summary>
	public static TableCell StartMargin(this TableCell cell, StartMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Sets the start margin of the cell in the given units.
	/// </summary>
	public static TableCell StartMargin(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().GetOrInit<StartMargin>().SetExactWidth(width, units);
		return cell;
	}

	/// <summary>
	/// Sets the end margin of the cell.
	/// </summary>
	public static TableCell EndMargin(this TableCell cell, EndMargin? margin)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().SetPropertyClassOrRemove(margin);
		return cell;
	}

	/// <summary>
	/// Sets the end margin of the cell in the given units.
	/// </summary>
	public static TableCell EndMargin(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellMargin>().GetOrInit<EndMargin>().SetExactWidth(width, units);
		return cell;
	}

	#endregion Margins

	/// <summary>
	/// Gets the width of the cell and the units of the result (e.g. 12pt or 66.6%).
	/// </summary>
	public static double? WidthValue(this TableCell cell, out WidthUnits? units)
	{
		units = null;
		return cell.TableCellProperties?.TableCellWidth?.GetExactWidth(out units);
	}

	/// <summary>
	/// Sets the width of the cell in the given units.
	/// </summary>
	public static TableCell Width(this TableCell cell, double width, WidthUnits? units)
	{
		cell.GetOrInit<TableCellProperties>().GetOrInit<TableCellWidth>().SetExactWidth(width, units);
		return cell;
	}

	/// <summary>
	/// Sets the background fill color of the cell.
	/// <para><c>Fill</c> property of <see cref="Shading" />.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.Shading" /></para>
	/// </summary>
	public static TableCell FillColor(this TableCell cell, string htmlColor)
	{
		Shading shading = cell.GetOrInit<TableCellProperties>().GetOrInit<Shading>();
		shading.Fill = htmlColor;
		shading.Val = shading.Val?.Value is null or ShadingPatternValues.Nil
			? ShadingPatternValues.Clear
			: shading.Val.Value; // Preserve original pattern

		return cell;
	}
}
