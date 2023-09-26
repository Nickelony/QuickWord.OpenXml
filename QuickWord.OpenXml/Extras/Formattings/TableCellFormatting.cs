using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class TableCellFormatting
{
	public ConditionalFormatStyle? ConditionalFormatStyle { get; set; }
	public int? GridSpan { get; set; }
	public bool? HideEndOfCellMarker { get; set; }
	public MergedCellValues? HorizontalMerge { get; set; }
	public bool? NoContentWrapping { get; set; }
	public Shading? Shading { get; set; }
	public TableCellBorders? Borders { get; set; }
	public bool? FitText { get; set; }
	public TableCellMargin? Margins { get; set; }
	public TableCellWidth? Width { get; set; }
	public TextDirectionValues? TextDirection { get; set; }
	public TableVerticalAlignmentValues? VerticalContentAlignment { get; set; }
	public MergedCellValues? VerticalMerge { get; set; }
}
