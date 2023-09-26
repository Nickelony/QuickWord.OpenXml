// Ignore Spelling: Inlined

using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingConversionExtensions
{
	/// <summary>
	/// Determines whether the drawing is inlined. This means that the drawing is in line with the text.
	/// </summary>
	public static bool IsInlined(this Drawing drawing)
		=> drawing.Inline is not null;

	/// <summary>
	/// Determines whether the drawing is anchored. Anchored drawings allow for more control over positioning and enables text wrapping.
	/// </summary>
	public static bool IsAnchored(this Drawing drawing)
		=> drawing.Anchor is not null;

	/// <summary>
	/// Converts the drawing into an inlined drawing (if applicable). This means that the drawing will be in line with the text.
	/// </summary>
	public static Drawing ToInlinedDrawing(this Drawing drawing)
	{
		if (drawing.IsInlined())
			return drawing; // Already inlined

		DW.Extent? anchorExtent = drawing.Anchor?.Extent;
		DW.EffectExtent? anchorEffectExtent = drawing.Anchor?.EffectExtent;
		DW.DocProperties? anchorDocProperties = drawing.Anchor?.GetFirstChild<DW.DocProperties>();
		DW.NonVisualGraphicFrameDrawingProperties? anchorProperties = drawing.Anchor?.GetFirstChild<DW.NonVisualGraphicFrameDrawingProperties>();
		A.Graphic? anchorGraphic = drawing.Anchor?.GetFirstChild<A.Graphic>();

		if (anchorExtent is null || anchorEffectExtent is null || anchorDocProperties is null || anchorProperties is null || anchorGraphic is null)
			return drawing; // Should generally never happen

		drawing.Inline = new DW.Inline
		{
			Extent = anchorExtent.CloneNode(true) as DW.Extent,
			EffectExtent = anchorEffectExtent.CloneNode(true) as DW.EffectExtent,
			DocProperties = anchorDocProperties.CloneNode(true) as DW.DocProperties,
			NonVisualGraphicFrameDrawingProperties = anchorProperties.CloneNode(true) as DW.NonVisualGraphicFrameDrawingProperties,
			Graphic = anchorGraphic.CloneNode(true) as A.Graphic
		};

		drawing.Anchor?.Remove();
		return drawing;
	}

	/// <summary>
	/// Converts the drawing into an anchored drawing (if applicable). Anchored drawings allow for more control over positioning and enables text wrapping.
	/// </summary>
	public static Drawing ToAnchoredDrawing(this Drawing drawing)
	{
		if (drawing.IsAnchored())
			return drawing; // Already anchored

		DW.Extent? inlineExtent = drawing.Inline?.Extent;
		DW.EffectExtent? inlineEffectExtent = drawing.Inline?.EffectExtent;
		DW.DocProperties? inlineDocProperties = drawing.Inline?.DocProperties;
		DW.NonVisualGraphicFrameDrawingProperties? inlineProperties = drawing.Inline?.NonVisualGraphicFrameDrawingProperties;
		A.Graphic? inlineGraphic = drawing.Inline?.Graphic;

		if (inlineExtent is null || inlineEffectExtent is null || inlineDocProperties is null || inlineProperties is null || inlineGraphic is null)
			return drawing; // Should generally never happen

		drawing.Anchor = new DW.Anchor()
		{
			RelativeHeight = 0,
			AllowOverlap = true,
			BehindDoc = false,
			LayoutInCell = false,
			Locked = false,
			SimplePos = false
		};

		drawing.Anchor.AppendChild(new DW.SimplePosition() { X = 0, Y = 0 });

		drawing.Anchor.AppendChild(new DW.HorizontalPosition(new DW.HorizontalAlignment("left"))
		{ RelativeFrom = DW.HorizontalRelativePositionValues.Margin });

		drawing.Anchor.AppendChild(new DW.VerticalPosition(new DW.VerticalAlignment("top"))
		{ RelativeFrom = DW.VerticalRelativePositionValues.Margin });

		drawing.Anchor.AppendChild(inlineExtent.CloneNode(true));
		drawing.Anchor.AppendChild(inlineEffectExtent.CloneNode(true));
		drawing.Anchor.AppendChild(new DW.WrapSquare { WrapText = DW.WrapTextValues.BothSides });
		drawing.Anchor.AppendChild(inlineDocProperties.CloneNode(true));
		drawing.Anchor.AppendChild(inlineProperties.CloneNode(true));
		drawing.Anchor.AppendChild(inlineGraphic.CloneNode(true));

		drawing.Inline?.Remove();
		return drawing;
	}
}
