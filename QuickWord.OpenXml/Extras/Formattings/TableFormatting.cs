using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml.Extras;

public class TableFormatting
{
	public bool? VisuallyBiDirectional { get; set; }
	public TableRowAlignmentValues? Justification { get; set; }
	public Shading? Shading { get; set; }
	public TableBorders? Borders { get; set; }
	public string? Caption { get; set; }
	public TableCellMarginDefault? DefaultCellMargins { get; set; }
	public TableCellSpacing? CellSpacing { get; set; }
	public string? Description { get; set; }
	public TableIndentation? Indentation { get; set; }
	public TableLayoutValues? Layout { get; set; }
	public TableLook? Look { get; set; }
	public TableOverlapValues? Overlap { get; set; }
	public TablePositionProperties? PositionProperties { get; set; }
	public string? Style { get; set; }
	public int? StyleColumnBandSize { get; set; }
	public int? StyleRowBandSize { get; set; }
	public TableWidth? Width { get; set; }
}
