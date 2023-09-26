using DocumentFormat.OpenXml.Wordprocessing;
using System;
using QuickWord.OpenXml.Measurements;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingWrappingExtensions
{
	/// <summary>
	/// Gets the wrapping type of the drawing.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	/// <exception cref="Exception">Occurs when there is no wrapping property in the anchor.</exception>
	public static WrappingType GetWrappingType(this Drawing drawing)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		if (drawing.Anchor!.GetFirstChild<DW.WrapNone>() is not null)
			return WrappingType.None;

		if (drawing.Anchor!.GetFirstChild<DW.WrapSquare>() is not null)
			return WrappingType.Square;

		if (drawing.Anchor!.GetFirstChild<DW.WrapTight>() is not null)
			return WrappingType.Tight;

		if (drawing.Anchor!.GetFirstChild<DW.WrapThrough>() is not null)
			return WrappingType.Through;

		if (drawing.Anchor!.GetFirstChild<DW.WrapTopBottom>() is not null)
			return WrappingType.TopAndBottom;

		throw new Exception(CONSTS.WRAPPING_NOT_FOUND);
	}

	/// <summary>
	/// Makes the drawing appear either in front or behind the text (if BehindText() is set).
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing NoTextWrapping(this Drawing drawing)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.ResetAllWrapping();
		drawing.Anchor!.InsertAfter(new DW.WrapNone(), drawing.GetOrInitEffectExtent());
		return drawing;
	}

	/// <summary>
	/// Sets "Square" wrapping with the specified distances from the closest text (basically the margin around the drawing).
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing SquareWrapping(this Drawing drawing, double distanceFromLeft, double distanceFromTop,
		double distanceFromRight, double distanceFromBottom, ImageMeasuringUnits units,
		DW.WrapTextValues wrapValue = DW.WrapTextValues.BothSides)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.ResetAllWrapping();

		drawing.Anchor!.DistanceFromLeft = (uint)(Pixels.FromOther(distanceFromLeft, units) * CONSTS.EMU_PER_PIXEL);
		drawing.Anchor!.DistanceFromTop = (uint)(Pixels.FromOther(distanceFromTop, units) * CONSTS.EMU_PER_PIXEL);
		drawing.Anchor!.DistanceFromRight = (uint)(Pixels.FromOther(distanceFromRight, units) * CONSTS.EMU_PER_PIXEL);
		drawing.Anchor!.DistanceFromBottom = (uint)(Pixels.FromOther(distanceFromBottom, units) * CONSTS.EMU_PER_PIXEL);

		drawing.Anchor!.InsertAfter(new DW.WrapSquare
		{
			WrapText = wrapValue
		}, drawing.GetOrInitEffectExtent());

		return drawing;
	}

	/// <summary>
	/// Sets "Tight" wrapping with the specified shape and distances from the closest text (basically a left and right margin).
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing TightWrapping(this Drawing drawing, DW.WrapPolygon polygon, double distanceFromLeft, double distanceFromRight,
		ImageMeasuringUnits units, DW.WrapTextValues wrapValue = DW.WrapTextValues.BothSides)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.ResetAllWrapping();

		drawing.Anchor!.DistanceFromLeft = (uint)(Pixels.FromOther(distanceFromLeft, units) * CONSTS.EMU_PER_PIXEL);
		drawing.Anchor!.DistanceFromRight = (uint)(Pixels.FromOther(distanceFromRight, units) * CONSTS.EMU_PER_PIXEL);

		drawing.Anchor!.InsertAfter(new DW.WrapTight
		{
			WrapText = wrapValue,
			WrapPolygon = polygon
		}, drawing.GetOrInitEffectExtent());

		return drawing;
	}

	/// <summary>
	/// Sets "Through" wrapping with the specified shape and distances from the closest text (basically a left and right margin).
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing ThroughWrapping(this Drawing drawing, DW.WrapPolygon polygon, double distanceFromLeft, double distanceFromRight,
		ImageMeasuringUnits units, DW.WrapTextValues wrapValue = DW.WrapTextValues.BothSides)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.ResetAllWrapping();

		drawing.Anchor!.DistanceFromLeft = (uint)(Pixels.FromOther(distanceFromLeft, units) * CONSTS.EMU_PER_PIXEL);
		drawing.Anchor!.DistanceFromRight = (uint)(Pixels.FromOther(distanceFromRight, units) * CONSTS.EMU_PER_PIXEL);

		drawing.Anchor!.InsertAfter(new DW.WrapThrough
		{
			WrapText = wrapValue,
			WrapPolygon = polygon
		}, drawing.GetOrInitEffectExtent());

		return drawing;
	}

	/// <summary>
	/// Sets "Top and bottom" wrapping with the specified distances from the closest text (basically a top and bottom margin).
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing TopAndBottomWrapping(this Drawing drawing, double distanceFromTop, double distanceFromBottom,
		ImageMeasuringUnits units)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.ResetAllWrapping();

		drawing.Anchor!.DistanceFromTop = (uint)(Pixels.FromOther(distanceFromTop, units) * CONSTS.EMU_PER_PIXEL);
		drawing.Anchor!.DistanceFromBottom = (uint)(Pixels.FromOther(distanceFromBottom, units) * CONSTS.EMU_PER_PIXEL);

		drawing.Anchor!.InsertAfter(new DW.WrapTopBottom(), drawing.GetOrInitEffectExtent());

		return drawing;
	}

	/// <summary>
	/// Removes and resets all wrapping properties of the drawing.
	/// </summary>
	private static void ResetAllWrapping(this Drawing drawing)
	{
		drawing.Anchor!.RemoveAllChildren<DW.WrapNone>();
		drawing.Anchor!.RemoveAllChildren<DW.WrapSquare>();
		drawing.Anchor!.RemoveAllChildren<DW.WrapTight>();
		drawing.Anchor!.RemoveAllChildren<DW.WrapThrough>();
		drawing.Anchor!.RemoveAllChildren<DW.WrapTopBottom>();

		drawing.Anchor!.DistanceFromTop
			= drawing.Anchor!.DistanceFromBottom
			= drawing.Anchor!.DistanceFromLeft
			= drawing.Anchor!.DistanceFromRight
			= 0;
	}
}
