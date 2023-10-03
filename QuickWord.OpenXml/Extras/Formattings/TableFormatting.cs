using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;

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

	public QBorder QBorders
	{
		set => QLeftBorder
			 = QTopBorder
			 = QRightBorder
			 = QBottomBorder
			 = QInsideHorizontalBorder
			 = QInsideVerticalBorder
			 = QStartBorder
			 = QEndBorder
			 = value;
	}

	public QBorder QLeftBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.LeftBorder = new LeftBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QTopBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.TopBorder = new TopBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QRightBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.RightBorder = new RightBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QBottomBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.BottomBorder = new BottomBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QInsideHorizontalBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.InsideHorizontalBorder = new InsideHorizontalBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QInsideVerticalBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.InsideVerticalBorder = new InsideVerticalBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QStartBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.StartBorder = new StartBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QEndBorder
	{
		set
		{
			Borders ??= new TableBorders();
			Borders.EndBorder = new EndBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public double QDefaultLeftMarginOfCells
	{
		set
		{
			DefaultCellMargins ??= new TableCellMarginDefault();
			DefaultCellMargins.TableCellLeftMargin = new TableCellLeftMargin
			{
				Width = (short)(value * 20),
				Type = TableWidthValues.Dxa
			};
		}
	}

	public QWidth QDefaultTopMarginOfCells
	{
		set
		{
			DefaultCellMargins ??= new TableCellMarginDefault();
			DefaultCellMargins.TopMargin = new TopMargin().SetExactWidth(value.Width, value.Units) as TopMargin;
		}
	}

	public QWidth QDefaultRightMarginOfCells
	{
		set
		{
			DefaultCellMargins ??= new TableCellMarginDefault();
			DefaultCellMargins.TableCellRightMargin = new TableCellRightMargin
			{
				Width = (short)(value.Width * 20),
				Type = TableWidthValues.Dxa
			};
		}
	}

	public QWidth QDefaultBottomMarginOfCells
	{
		set
		{
			DefaultCellMargins ??= new TableCellMarginDefault();
			DefaultCellMargins.BottomMargin = new BottomMargin().SetExactWidth(value.Width, value.Units) as BottomMargin;
		}
	}

	public QWidth QDefaultStartMarginOfCells
	{
		set
		{
			DefaultCellMargins ??= new TableCellMarginDefault();
			DefaultCellMargins.StartMargin = new StartMargin().SetExactWidth(value.Width, value.Units) as StartMargin;
		}
	}

	public QWidth QDefaultEndMarginOfCells
	{
		set
		{
			DefaultCellMargins ??= new TableCellMarginDefault();
			DefaultCellMargins.EndMargin = new EndMargin().SetExactWidth(value.Width, value.Units) as EndMargin;
		}
	}

	public QWidth QCellSpacing
	{
		set => CellSpacing = new TableCellSpacing().SetExactWidth(value.Width, value.Units) as TableCellSpacing;
	}

	public QWidth QIndentation
	{
		set => Indentation = new TableIndentation().SetExactWidth(value.Width, value.Units);
	}

	public QWidth QWidth
	{
		set => Width = new TableWidth().SetExactWidth(value.Width, value.Units) as TableWidth;
	}

	public string? FillColor
	{
		get => Shading?.Fill;
		set
		{
			Shading ??= new Shading();
			Shading.Fill = value;
		}
	}
}
