// Ignore Spelling: Kinsoku

using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;

namespace QuickWord.OpenXml.Extras;

public class ParagraphFormatting
{
	public bool? AdjustRightIndent { get; set; }
	public bool? AutoSpaceDE { get; set; }
	public bool? AutoSpaceDN { get; set; }
	public bool? BiDirectional { get; set; }
	public ConditionalFormatStyle? ConditionalFormatStyle { get; set; }
	public bool? ContextualSpacing { get; set; }
	public string? DivId { get; set; }
	public FrameProperties? FrameProperties { get; set; }
	public Indentation? Indentation { get; set; }
	public JustificationValues? Justification { get; set; }
	public bool? KeepLinesTogether { get; set; }
	public bool? KeepWithNext { get; set; }
	public bool? Kinsoku { get; set; }
	public bool? MirrorIndents { get; set; }
	public NumberingProperties? NumberingProperties { get; set; }
	public OutlineLevelValues? OutlineLevel { get; set; }
	public bool? OverflowPunctuation { get; set; }
	public bool? PageBreakBefore { get; set; }
	public ParagraphBorders? Borders { get; set; }
	public string? Style { get; set; }
	public ParagraphMarkRunProperties? MarkRunProperties { get; set; }
	public SectionProperties? SectionProperties { get; set; }
	public Shading? Shading { get; set; }
	public bool? SnapToGrid { get; set; }
	public SpacingBetweenLines? Spacing { get; set; }
	public bool? SuppressAutoHyphenation { get; set; }
	public bool? SuppressLineNumbers { get; set; }
	public bool? SuppressOverlapping { get; set; }
	public Tabs? Tabs { get; set; }
	public VerticalTextAlignmentValues? VerticalTextAlignment { get; set; }
	public TextBoxTightWrapValues? TextBoxTightWrap { get; set; }
	public TextDirectionValues? TextDirection { get; set; }
	public bool? TopLinePunctuation { get; set; }
	public bool? WidowControl { get; set; }
	public bool? WordWrap { get; set; }

	// Extras:

	public QParagraphSpacing QLineSpacing
	{
		set
		{
			Spacing ??= new SpacingBetweenLines();
			Spacing.Line = Twips.FromOther(value.Spacing, value.Units).ToString();
		}
	}

	public QParagraphSpacing QSpacingBefore
	{
		set
		{
			Spacing ??= new SpacingBetweenLines();
			Spacing.Before = Twips.FromOther(value.Spacing, value.Units).ToString();
		}
	}

	public QParagraphSpacing QSpacingAfter
	{
		set
		{
			Spacing ??= new SpacingBetweenLines();
			Spacing.After = Twips.FromOther(value.Spacing, value.Units).ToString();
		}
	}

	public QParagraphIndentation QLeftIndentation
	{
		set
		{
			Indentation ??= new Indentation();

			if (value.Units is IndentationUnits.Characters)
			{
				Indentation.Left = null;
				Indentation.LeftChars = (int)value.Indentation;
			}
			else
			{
				Indentation.Left = Twips.FromOther(value.Indentation, value.Units).ToString();
				Indentation.LeftChars = null;
			}
		}
	}

	public QParagraphIndentation QRightIndentation
	{
		set
		{
			Indentation ??= new Indentation();

			if (value.Units is IndentationUnits.Characters)
			{
				Indentation.Right = null;
				Indentation.RightChars = (int)value.Indentation;
			}
			else
			{
				Indentation.Right = Twips.FromOther(value.Indentation, value.Units).ToString();
				Indentation.RightChars = null;
			}
		}
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

	public QBorder QLeftBorder
	{
		set
		{
			Borders ??= new ParagraphBorders();
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
			Borders ??= new ParagraphBorders();
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
			Borders ??= new ParagraphBorders();
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
			Borders ??= new ParagraphBorders();
			Borders.BottomBorder = new BottomBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QBarBorder
	{
		set
		{
			Borders ??= new ParagraphBorders();
			Borders.BarBorder = new BarBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}

	public QBorder QBetweenBorder
	{
		set
		{
			Borders ??= new ParagraphBorders();
			Borders.BetweenBorder = new BetweenBorder
			{
				Val = value.Border,
				Size = BorderSize.ToSixth(value.Width),
				Color = value.Color,
				Space = value.Spacing
			};
		}
	}
}
