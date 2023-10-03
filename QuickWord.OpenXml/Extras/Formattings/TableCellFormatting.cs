using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;

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

	public QBorder QLeftBorder
	{
		set
		{
			Borders ??= new TableCellBorders();
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
			Borders ??= new TableCellBorders();
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
			Borders ??= new TableCellBorders();
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
			Borders ??= new TableCellBorders();
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
			Borders ??= new TableCellBorders();
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
			Borders ??= new TableCellBorders();
			Borders.InsideVerticalBorder = new InsideVerticalBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QTopLeftToBottomRightCellBorder
	{
		set
		{
			Borders ??= new TableCellBorders();
			Borders.TopLeftToBottomRightCellBorder = new TopLeftToBottomRightCellBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QTopRightToBottomLeftCellBorder
	{
		set
		{
			Borders ??= new TableCellBorders();
			Borders.TopRightToBottomLeftCellBorder = new TopRightToBottomLeftCellBorder
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
			Borders ??= new TableCellBorders();
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
			Borders ??= new TableCellBorders();
			Borders.EndBorder = new EndBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QWidth QLeftMargin
	{
		set
		{
			Margins ??= new TableCellMargin();
			Margins.LeftMargin = new LeftMargin().SetExactWidth(value.Width, value.Units) as LeftMargin;
		}
	}

	public QWidth QTopMargin
	{
		set
		{
			Margins ??= new TableCellMargin();
			Margins.TopMargin = new TopMargin().SetExactWidth(value.Width, value.Units) as TopMargin;
		}
	}

	public QWidth QRightMargin
	{
		set
		{
			Margins ??= new TableCellMargin();
			Margins.RightMargin = new RightMargin().SetExactWidth(value.Width, value.Units) as RightMargin;
		}
	}

	public QWidth QBottomMargin
	{
		set
		{
			Margins ??= new TableCellMargin();
			Margins.BottomMargin = new BottomMargin().SetExactWidth(value.Width, value.Units) as BottomMargin;
		}
	}

	public QWidth QStartMargin
	{
		set
		{
			Margins ??= new TableCellMargin();
			Margins.StartMargin = new StartMargin().SetExactWidth(value.Width, value.Units) as StartMargin;
		}
	}

	public QWidth QEndMargin
	{
		set
		{
			Margins ??= new TableCellMargin();
			Margins.EndMargin = new EndMargin().SetExactWidth(value.Width, value.Units) as EndMargin;
		}
	}

	public QWidth QWidth
	{
		set => Width = new TableCellWidth().SetExactWidth(value.Width, value.Units) as TableCellWidth;
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
