using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;
using System;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingUtilityExtensions
{
	private static OpenXmlElement GetDrawingTypeObject(this Drawing drawing)
	{
		if (drawing.IsInlined())
			return drawing.Inline!;

		if (drawing.IsAnchored())
			return drawing.Anchor!;

		throw new InvalidOperationException(CONSTS.UNKNOWN_DRAWING_TYPE);
	}

	internal static DW.Extent? GetExtent(this Drawing drawing)
		=> drawing.GetDrawingTypeObject().GetFirstChild<DW.Extent>();

	internal static PIC.Picture? GetPicture(this Drawing drawing)
		=> drawing.GetDrawingTypeObject().GetFirstChild<A.Graphic>()?.GraphicData?.GetFirstChild<PIC.Picture>();

	/// <summary>
	/// Gets the source rectangle responsible for cropping the image. Returns <see langword="null" /> if the image is not cropped.
	/// </summary>
	internal static A.SourceRectangle? GetSourceRectangle(this Drawing drawing)
		=> drawing.GetPicture()?.BlipFill?.SourceRectangle;

	internal static Drawing SetEffectExtent(this Drawing drawing, long xExtent, long yExtent)
	{
		long newCx = (long)(xExtent * CONSTS.EMU_PER_PIXEL);
		long newCy = (long)(yExtent * CONSTS.EMU_PER_PIXEL);

		DW.EffectExtent effectExtent = drawing.GetOrInitEffectExtent();
		effectExtent.LeftEdge = effectExtent.RightEdge = newCx;
		effectExtent.TopEdge = effectExtent.BottomEdge = newCy;

		return drawing;
	}

	internal static DW.Extent GetOrInitExtent(this Drawing drawing)
	{
		return drawing.GetDrawingTypeObject()
			.GetOrInit<DW.Extent>();
	}

	internal static DW.EffectExtent GetOrInitEffectExtent(this Drawing drawing)
	{
		return drawing.GetDrawingTypeObject()
			.GetOrInit<DW.EffectExtent>();
	}

	internal static PIC.Picture GetOrInitPicture(this Drawing drawing)
	{
		return drawing.GetDrawingTypeObject()
			.GetOrInit<A.Graphic>()
			.GetOrInit<A.GraphicData>()
			.GetOrInit<PIC.Picture>();
	}

	internal static A.SourceRectangle GetOrInitSourceRectangle(this Drawing drawing)
	{
		return drawing.GetOrInitPicture()
			.GetOrInit<PIC.BlipFill>()
			.GetOrInit<A.SourceRectangle>();
	}

	internal static PIC.ShapeProperties GetOrInitShapeProperties(this Drawing drawing)
	{
		return drawing.GetOrInitPicture()
			.GetOrInit<PIC.ShapeProperties>();
	}

	internal static A.Transform2D GetOrInitTransform2D(this Drawing drawing)
	{
		return drawing.GetOrInitShapeProperties()
			.GetOrInit<A.Transform2D>();
	}

	internal static A.Extents GetOrInitTransform2DExtents(this Drawing drawing)
	{
		return drawing.GetOrInitTransform2D()
			.GetOrInit<A.Extents>();
	}

	internal static A.AlphaModulationFixed GetOrInitAlphaModulationFixed(this Drawing drawing)
	{
		return drawing.GetOrInitPicture()
			.GetOrInit<PIC.BlipFill>()
			.GetOrInit<A.Blip>()
			.GetOrInit<A.AlphaModulationFixed>();
	}
}
