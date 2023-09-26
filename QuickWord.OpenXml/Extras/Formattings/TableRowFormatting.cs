using DocumentFormat.OpenXml.Wordprocessing;

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
}
