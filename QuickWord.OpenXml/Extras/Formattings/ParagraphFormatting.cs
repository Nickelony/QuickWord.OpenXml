﻿// Ignore Spelling: Kinsoku

using DocumentFormat.OpenXml.Wordprocessing;

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
}
