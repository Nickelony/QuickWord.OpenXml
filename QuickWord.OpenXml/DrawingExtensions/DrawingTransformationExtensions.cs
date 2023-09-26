using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using System.Linq;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingTransformationExtensions
{
	#region Get methods

	/// <summary>
	/// Gets the width of the image in the desired units.
	/// </summary>
	public static double GetWidth(this Drawing drawing, ImageMeasuringUnits desiredUnits = ImageMeasuringUnits.Pixels)
	{
		DW.Extent? extent = drawing.GetExtent()
			?? throw new InvalidOperationException(CONSTS.INVALID_EXTENT);

		double width = extent.Cx!.Value / CONSTS.EMU_PER_PIXEL;
		return Pixels.ToOther(width, desiredUnits);
	}

	/// <summary>
	/// Gets the height of the image in the desired units.
	/// </summary>
	public static double GetHeight(this Drawing drawing, ImageMeasuringUnits desiredUnits = ImageMeasuringUnits.Pixels)
	{
		DW.Extent? extent = drawing.GetExtent()
			?? throw new InvalidOperationException(CONSTS.INVALID_EXTENT);

		double height = extent.Cy!.Value / CONSTS.EMU_PER_PIXEL;
		return Pixels.ToOther(height, desiredUnits);
	}

	/// <summary>
	/// Gets the original width of the image before any scaling or cropping.
	/// </summary>
	[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
	public static double GetOriginalWidth(this Drawing drawing, ImageMeasuringUnits desiredUnits = ImageMeasuringUnits.Pixels)
	{
		Document? parentDocument = drawing.Ancestors<Document>().FirstOrDefault();
		string? blipEmbed = drawing.GetPicture()?.BlipFill?.Blip?.Embed;

		if (parentDocument is null || blipEmbed is null ||
			parentDocument.MainDocumentPart?.GetPartById(blipEmbed) is not ImagePart imagePart)
			throw new InvalidOperationException(CONSTS.COULD_NOT_FETCH_IMAGE_DATA);

		using Stream stream = imagePart.GetStream();
		using var bitmap = Image.FromStream(stream);
		return Pixels.ToOther(bitmap.Width, desiredUnits);
	}

	/// <summary>
	/// Gets the original height of the image before any scaling or cropping.
	/// </summary>
	[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
	public static double GetOriginalHeight(this Drawing drawing, ImageMeasuringUnits desiredUnits = ImageMeasuringUnits.Pixels)
	{
		Document? parentDocument = drawing.Ancestors<Document>().FirstOrDefault();
		string? blipEmbed = drawing.GetPicture()?.BlipFill?.Blip?.Embed;

		if (parentDocument is null || blipEmbed is null ||
			parentDocument.MainDocumentPart?.GetPartById(blipEmbed) is not ImagePart imagePart)
			throw new InvalidOperationException(CONSTS.COULD_NOT_FETCH_IMAGE_DATA);

		using Stream stream = imagePart.GetStream();
		using var bitmap = Image.FromStream(stream);
		return Pixels.ToOther(bitmap.Height, desiredUnits);
	}

	/// <summary>
	/// Gets the rotation angle of the image in degrees.
	/// <para>Returns <see langword="null" /> if the property is not set.</para>
	/// </summary>
	public static double? RotationValue(this Drawing drawing)
		=> drawing.GetOrInitTransform2D()?.Rotation?.Value / CONSTS.ANGLE_MULTIPLIER;

	/// <summary>
	/// Specifies whether the image is horizontally flipped.
	/// </summary>
	public static bool FlipHorizontallyValue(this Drawing drawing)
		=> drawing.GetOrInitTransform2D()?.HorizontalFlip?.Value ?? false;

	/// <summary>
	/// Specifies whether the image is vertically flipped.
	/// </summary>
	public static bool FlipVerticallyValue(this Drawing drawing)
		=> drawing.GetOrInitTransform2D()?.VerticalFlip?.Value ?? false;

	/// <summary>
	/// Gets the absolute horizontal position of the image in the desired units.
	/// <para>Returns <see langword="null" /> if the property is not set.</para>
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static double? AbsoluteHorizontalPositionValue(this Drawing drawing, ImageMeasuringUnits desiredUnits, out DW.HorizontalRelativePositionValues? toTheRightOf)
	{
		toTheRightOf = null;

		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.HorizontalPosition? horizontalPosition = drawing.Anchor!.HorizontalPosition;

		if (horizontalPosition is null)
			return null;

		toTheRightOf = horizontalPosition.RelativeFrom!.Value;
		long positionOffset = long.Parse(horizontalPosition.PositionOffset!.Text);
		return Pixels.ToOther((long)(positionOffset / CONSTS.EMU_PER_PIXEL), desiredUnits);
	}

	/// <summary>
	/// Gets the absolute vertical position of the image in the desired units.
	/// <para>Returns <see langword="null" /> if the property is not set.</para>
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static double? AbsoluteVerticalPositionValue(this Drawing drawing, ImageMeasuringUnits desiredUnits, out DW.VerticalRelativePositionValues? below)
	{
		below = null;

		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.VerticalPosition? verticalPosition = drawing.Anchor!.VerticalPosition;

		if (verticalPosition is null)
			return null;

		below = verticalPosition.RelativeFrom!.Value;
		long positionOffset = long.Parse(verticalPosition.PositionOffset!.Text);
		return Pixels.ToOther((long)(positionOffset / CONSTS.EMU_PER_PIXEL), desiredUnits);
	}

	/// <summary>
	/// Gets the horizontal position of the image relative to the page margins.
	/// <para>Returns <see langword="null" /> if the property is not set.</para>
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static DW.HorizontalAlignmentValues? HorizontalAlignmentValue(this Drawing drawing, out DW.HorizontalRelativePositionValues? relativeTo)
	{
		relativeTo = null;

		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.HorizontalPosition? horizontalPosition = drawing.Anchor!.HorizontalPosition;

		if (!Enum.TryParse(horizontalPosition?.HorizontalAlignment?.Text, true, out DW.HorizontalAlignmentValues result))
			return null;

		relativeTo = horizontalPosition!.RelativeFrom!.Value;
		return result;
	}

	/// <summary>
	/// Gets the vertical position of the image relative to the page margins.
	/// <para>Returns <see langword="null" /> if the property is not set.</para>
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static DW.VerticalAlignmentValues? VerticalAlignmentValue(this Drawing drawing, out DW.VerticalRelativePositionValues? relativeTo)
	{
		relativeTo = null;

		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.VerticalPosition? verticalPosition = drawing.Anchor!.VerticalPosition;

		if (!Enum.TryParse(verticalPosition?.VerticalAlignment?.Text, true, out DW.VerticalAlignmentValues result))
			return null;

		relativeTo = verticalPosition!.RelativeFrom!.Value;
		return result;
	}

	#endregion Get methods

	#region Set methods

	/// <summary>
	/// Sets the width of the image in the desired units.
	/// </summary>
	/// <param name="keepAspectRatio">Causes the height of the image to automatically adjust in order to preserve the aspect ratio.</param>
	public static Drawing SetWidth(this Drawing drawing, double width, ImageMeasuringUnits units = ImageMeasuringUnits.Pixels, bool keepAspectRatio = false)
	{
		DW.Extent extent = drawing.GetOrInitExtent();
		A.Extents transform2dExtents = drawing.GetOrInitTransform2DExtents();

		long newCx = (long)(Pixels.FromOther(width, units) * CONSTS.EMU_PER_PIXEL);
		long lastCx = extent.Cx?.Value ?? 1;
		double scaleDifference = (double)newCx / lastCx;

		extent.Cx = newCx;
		transform2dExtents.Cx = newCx;

		if (keepAspectRatio && extent.Cy is not null && transform2dExtents.Cy is not null)
		{
			extent.Cy = (long)(extent.Cy * scaleDifference);
			transform2dExtents.Cy = (long)(transform2dExtents.Cy * scaleDifference);
		}

		return drawing;
	}

	/// <summary>
	/// Sets the height of the image in the desired units.
	/// </summary>
	/// <param name="keepAspectRatio">Causes the width of the image to automatically adjust in order to preserve the aspect ratio.</param>
	public static Drawing SetHeight(this Drawing drawing, double height, ImageMeasuringUnits units = ImageMeasuringUnits.Pixels, bool keepAspectRatio = false)
	{
		DW.Extent extent = drawing.GetOrInitExtent();
		A.Extents transform2dExtents = drawing.GetOrInitTransform2DExtents();

		long newCy = (long)(Pixels.FromOther(height, units) * CONSTS.EMU_PER_PIXEL);
		long lastCy = extent.Cy?.Value ?? 1;
		double scaleDifference = (double)newCy / lastCy;

		extent.Cy = newCy;
		transform2dExtents.Cy = newCy;

		if (keepAspectRatio && extent.Cx is not null && transform2dExtents.Cx is not null)
		{
			extent.Cx = (long)(extent.Cx * scaleDifference);
			transform2dExtents.Cx = (long)(transform2dExtents.Cx * scaleDifference);
		}

		return drawing;
	}

	/// <summary>
	/// Resizes the image to the given width and height and in the desired units.
	/// </summary>
	public static Drawing Resize(this Drawing drawing, double width, double height, ImageMeasuringUnits units = ImageMeasuringUnits.Pixels)
		=> drawing.SetWidth(width, units).SetHeight(height, units);

	/// <summary>
	/// Resets the image to its original size.
	/// <para>Note: This method does not affect cropping. Any performed cropping will still be preserved.</para>
	/// </summary>
	public static Drawing ResetSize(this Drawing drawing)
	{
		double originalWidth = drawing.GetOriginalWidth();
		double originalHeight = drawing.GetOriginalHeight();

		A.SourceRectangle? sourceRectangle = drawing.GetSourceRectangle();

		if (sourceRectangle is null)
			return drawing.Resize(originalWidth, originalHeight); // Image is not cropped

		var currentCropping = sourceRectangle.ToCropping();

		double leftDifference = originalWidth * currentCropping.LeftFactor;
		double rightDifference = originalWidth * currentCropping.RightFactor;
		double newWidth = originalWidth - leftDifference - rightDifference;
		newWidth = Math.Round(newWidth);

		double topDifference = originalHeight * currentCropping.TopFactor;
		double bottomDifference = originalHeight * currentCropping.BottomFactor;
		double newHeight = originalHeight - topDifference - bottomDifference;
		newHeight = Math.Round(newHeight);

		return drawing.Resize(newWidth, newHeight);
	}

	/// <summary>
	/// Scales the image horizontally by the given factor.
	/// </summary>
	public static Drawing ScaleHorizontally(this Drawing drawing, double factor, bool keepAspectRatio = false)
		=> drawing.SetWidth(drawing.GetWidth() * factor, ImageMeasuringUnits.Pixels, keepAspectRatio);

	/// <summary>
	/// Scales the image vertically by the given factor.
	/// </summary>
	public static Drawing ScaleVertically(this Drawing drawing, double factor, bool keepAspectRatio = false)
		=> drawing.SetHeight(drawing.GetHeight() * factor, ImageMeasuringUnits.Pixels, keepAspectRatio);

	/// <summary>
	/// Scales the image by the given factors.
	/// </summary>
	public static Drawing Scale(this Drawing drawing, double xFactor, double yFactor)
		=> drawing.ScaleHorizontally(xFactor).ScaleVertically(yFactor);

	/// <summary>
	/// Sets the rotation of the image to the given angle in degrees.
	/// </summary>
	/// <param name="angle">Rotation angle in degrees.</param>
	public static Drawing Rotation(this Drawing drawing, double angle)
	{
		drawing.GetOrInitTransform2D().Rotation = (int)(angle * CONSTS.ANGLE_MULTIPLIER);
		return drawing;
	}

	/// <inheritdoc cref="FlipHorizontallyValue" />
	public static Drawing FlipHorizontally(this Drawing drawing, bool value = true)
	{
		drawing.GetOrInitTransform2D().HorizontalFlip = value;
		return drawing;
	}

	/// <inheritdoc cref="FlipVerticallyValue" />
	public static Drawing FlipVertically(this Drawing drawing, bool value = true)
	{
		drawing.GetOrInitTransform2D().VerticalFlip = value;
		return drawing;
	}

	/// <summary>
	/// Sets the absolute horizontal position of the image in the given units.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing AbsoluteHorizontalPosition(this Drawing drawing, double x, ImageMeasuringUnits units, DW.HorizontalRelativePositionValues toTheRightOf)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.HorizontalPosition horizontalPosition = drawing.Anchor!.GetOrInit<DW.HorizontalPosition>();
		horizontalPosition.RemoveAllChildren<DW.HorizontalAlignment>();
		horizontalPosition.PositionOffset = new DW.PositionOffset((Pixels.FromOther(x, units) * CONSTS.EMU_PER_PIXEL).ToString());
		horizontalPosition.RelativeFrom = toTheRightOf;

		return drawing;
	}

	/// <summary>
	/// Sets the absolute vertical position of the image in the given units.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing AbsoluteVerticalPosition(this Drawing drawing, double y, ImageMeasuringUnits units, DW.VerticalRelativePositionValues below)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.VerticalPosition verticalPosition = drawing.Anchor!.GetOrInit<DW.VerticalPosition>();
		verticalPosition.RemoveAllChildren<DW.VerticalAlignment>();
		verticalPosition.PositionOffset = new DW.PositionOffset((Pixels.FromOther(y, units) * CONSTS.EMU_PER_PIXEL).ToString());
		verticalPosition.RelativeFrom = below;

		return drawing;
	}

	/// <summary>
	/// Sets the horizontal alignment of the image.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing HorizontalAlignment(this Drawing drawing, DW.HorizontalAlignmentValues alignment, DW.HorizontalRelativePositionValues relativeTo)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.HorizontalPosition horizontalPosition = drawing.Anchor!.GetOrInit<DW.HorizontalPosition>();
		horizontalPosition.RemoveAllChildren<DW.PositionOffset>();
		horizontalPosition.HorizontalAlignment = new DW.HorizontalAlignment(alignment.ToString().ToLower());
		horizontalPosition.RelativeFrom = relativeTo;

		return drawing;
	}

	/// <summary>
	/// Sets the vertical alignment of the image.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static Drawing VerticalAlignment(this Drawing drawing, DW.VerticalAlignmentValues alignment, DW.VerticalRelativePositionValues relativeTo)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		DW.VerticalPosition verticalPosition = drawing.Anchor!.GetOrInit<DW.VerticalPosition>();
		verticalPosition.RemoveAllChildren<DW.PositionOffset>();
		verticalPosition.VerticalAlignment = new DW.VerticalAlignment(alignment.ToString().ToLower());
		verticalPosition.RelativeFrom = relativeTo;

		return drawing;
	}

	#endregion Set methods
}
