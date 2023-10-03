using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;
using System.Collections.Generic;
using System.Linq;

namespace QuickWord.OpenXml.Extras;

/// <summary>
/// Additional extension methods for the <see cref="TableRow"/> class.
/// </summary>
public static class TableRowExtraExtensions
{
	public static IEnumerable<TableCell> Cells(this TableRow row)
		=> row.Elements<TableCell>();

	public static TableCell? Cells(this TableRow row, int index)
		=> row.Elements<TableCell>().ElementAtOrDefault(index);

	#region Formatting

	/// <summary>
	/// Clones all formatting properties from the table row.
	/// </summary>
	public static TableRowFormatting CloneFormatting(this TableRow row) => new()
	{
		CantSplit = row.CantSplitValue(),
		ConditionalFormatStyle = row.GetConditionalFormatStyle()?.CloneNode(true) as ConditionalFormatStyle,
		DivId = row.DivIdValue(),
		GridAfter = row.GridAfterValue(),
		GridBefore = row.GridBeforeValue(),
		HideEndOfRowMarker = row.HideEndOfRowMarkerValue(),
		Justification = row.JustificationValue(),
		IsHeader = row.IsHeaderValue(),
		Height = row.GetHeight(),
		WidthAfter = row.GetWidthAfter()?.CloneNode(true) as WidthAfterTableRow,
		WidthBefore = row.GetWidthBefore()?.CloneNode(true) as WidthBeforeTableRow
	};

	/// <summary>
	/// Applies the given formatting properties to the table row (replaces every possible property of the table row unless <c>ignoreNulls</c> is set to true).
	/// </summary>
	public static TableRow ApplyFormatting(this TableRow row, TableRowFormatting formatting, bool ignoreNulls = false)
	{
		if (formatting.CantSplit is not null || (formatting.CantSplit is null && !ignoreNulls))
			row.CantSplit(formatting.CantSplit);

		if (formatting.ConditionalFormatStyle is not null || (formatting.ConditionalFormatStyle is null && !ignoreNulls))
			row.ConditionalFormatStyle(formatting.ConditionalFormatStyle?.CloneNode(true) as ConditionalFormatStyle);

		if (formatting.DivId is not null || (formatting.DivId is null && !ignoreNulls))
			row.DivId(formatting.DivId);

		if (formatting.GridAfter is not null || (formatting.GridAfter is null && !ignoreNulls))
			row.GridAfter(formatting.GridAfter);

		if (formatting.GridBefore is not null || (formatting.GridBefore is null && !ignoreNulls))
			row.GridBefore(formatting.GridBefore);

		if (formatting.HideEndOfRowMarker is not null || (formatting.HideEndOfRowMarker is null && !ignoreNulls))
			row.HideEndOfRowMarker(formatting.HideEndOfRowMarker);

		if (formatting.Justification is not null || (formatting.Justification is null && !ignoreNulls))
			row.Justification(formatting.Justification);

		if (formatting.IsHeader is not null || (formatting.IsHeader is null && !ignoreNulls))
			row.IsHeader(formatting.IsHeader);

		if (formatting.Height is not null || (formatting.Height is null && !ignoreNulls))
			row.Height(formatting.Height);

		if (formatting.WidthAfter is not null || (formatting.WidthAfter is null && !ignoreNulls))
			row.WidthAfter(formatting.WidthAfter?.CloneNode(true) as WidthAfterTableRow);

		if (formatting.WidthBefore is not null || (formatting.WidthBefore is null && !ignoreNulls))
			row.WidthBefore(formatting.WidthBefore?.CloneNode(true) as WidthBeforeTableRow);

		return row;
	}

	/// <summary>
	/// Resets every possible property of the table row.
	/// </summary>
	public static TableRow ResetFormatting(this TableRow row)
	{
		row.RemoveAllChildren<TableRowProperties>();
		return row;
	}

	#endregion Formatting

	/// <summary>
	/// Gets the preferred width for the total number of grid columns before this table row and the units of the result (e.g. 12pt or 66.6%).
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthBeforeTableRow" /></para>
	/// </summary>
	public static double? WidthBeforeValue(this TableRow row, out WidthUnits? units)
	{
		units = null;
		return row.GetWidthBefore()?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the preferred width for the total number of grid columns after this table row and the units of the result (e.g. 12pt or 66.6%).
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthAfterTableRow" /></para>
	/// </summary>
	public static double? WidthAfterValue(this TableRow row, out WidthUnits? units)
	{
		units = null;
		return row.GetWidthAfter()?.GetExactWidth(out units);
	}

	/// <summary>
	/// Gets the height of the row in the desired units and the rule that is applied to the height.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableRowHeight" /></para>
	/// </summary>
	public static double? HeightValue(this TableRow row, MeasuringUnits desiredUnits, out HeightRuleValues? rule)
	{
		rule = null;
		TableRowHeight? tableRowHeight = row.GetHeight();

		if (tableRowHeight is null || tableRowHeight.Val is null)
			return null;

		rule = tableRowHeight.HeightType?.Value;
		return Twips.ToOther((int)tableRowHeight.Val.Value, desiredUnits);
	}

	/// <summary>
	/// Sets the preferred width for the total number of grid columns before this table row in the given units.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthBeforeTableRow" /></para>
	/// </summary>
	public static TableRow WidthBefore(this TableRow row, double width, WidthUnits units)
	{
		row.GetOrInit<TableRowProperties>().GetOrInit<WidthBeforeTableRow>().SetExactWidth(width, units);
		return row;
	}

	/// <summary>
	/// Sets the preferred width for the total number of grid columns after this table row in the given units.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.WidthAfterTableRow" /></para>
	/// </summary>
	public static TableRow WidthAfter(this TableRow row, double width, WidthUnits units)
	{
		row.GetOrInit<TableRowProperties>().GetOrInit<WidthAfterTableRow>().SetExactWidth(width, units);
		return row;
	}

	/// <summary>
	/// Sets the height of the row in the given units and the given rule.
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.TableRowHeight" /></para>
	/// </summary>
	public static TableRow Height(this TableRow row, double height, MeasuringUnits units, HeightRuleValues rule)
	{
		TableRowHeight tableRowHeight = row.GetOrInit<TableRowProperties>().GetOrInit<TableRowHeight>();
		tableRowHeight.HeightType = rule;
		tableRowHeight.Val = (uint)Twips.FromOther(height, units);

		return row;
	}
}
