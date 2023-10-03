using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;

namespace QuickWord.OpenXml.Extras;

public class TableRowFormatting
{
	public bool? CantSplit { get; set; }
	public ConditionalFormatStyle? ConditionalFormatStyle { get; set; }
	public string? DivId { get; set; }
	public int? GridAfter { get; set; }
	public int? GridBefore { get; set; }
	public bool? HideEndOfRowMarker { get; set; }
	public TableRowAlignmentValues? Justification { get; set; }
	public bool? IsHeader { get; set; }
	public TableRowHeight? Height { get; set; }
	public WidthAfterTableRow? WidthAfter { get; set; }
	public WidthBeforeTableRow? WidthBefore { get; set; }

	public QRowHeight QHeight
	{
		set => Height = new TableRowHeight()
		{
			Val = (uint)Twips.FromOther(value.Height, value.Units),
			HeightType = value.Rule
		};
	}

	public QWidth QWidthBefore
	{
		set => WidthBefore = new WidthBeforeTableRow().SetExactWidth(value.Width, value.Units) as WidthBeforeTableRow;
	}

	public QWidth QWidthAfter
	{
		set => WidthAfter = new WidthAfterTableRow().SetExactWidth(value.Width, value.Units) as WidthAfterTableRow;
	}
}
