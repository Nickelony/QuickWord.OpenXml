// Ignore Spelling: Uncropped

using DocumentFormat.OpenXml.Wordprocessing;
using System;
using QuickWord.OpenXml.Utilities;
using A = DocumentFormat.OpenXml.Drawing;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingCroppingExtensions
{
	#region Get methods

	/// <summary>
	/// Gets the cropping of the image on each side. Factor values are between 0 and 1 where 0.0 is no cropping and 1.0 is full cropping.
	/// <para>Returns <see langword="null" /> if cropping is not set.</para>
	/// </summary>
	public static Cropping? GetCropping(this Drawing drawing)
	{
		A.SourceRectangle? sourceRectangle = drawing.GetSourceRectangle();

		if (sourceRectangle is null)
			return null; // Image is not cropped

		return sourceRectangle.ToCropping();
	}

	/// <summary>
	/// Gets the uncropped width of the image.
	/// </summary>
	public static double GetUncroppedWidth(this Drawing drawing, ImageMeasuringUnits desiredUnits = ImageMeasuringUnits.Pixels)
	{
		double currentWidth = drawing.GetWidth(desiredUnits);
		A.SourceRectangle? sourceRectangle = drawing.GetSourceRectangle();

		if (sourceRectangle is null)
			return currentWidth; // Image is not cropped

		var cropping = sourceRectangle.ToCropping();
		return currentWidth * (1 / (1 - (cropping.LeftFactor + cropping.RightFactor)));
	}

	/// <summary>
	/// Gets the uncropped height of the image.
	/// </summary>
	public static double GetUncroppedHeight(this Drawing drawing, ImageMeasuringUnits desiredUnits = ImageMeasuringUnits.Pixels)
	{
		double currentHeight = drawing.GetHeight(desiredUnits);
		A.SourceRectangle? sourceRectangle = drawing.GetSourceRectangle();

		if (sourceRectangle is null)
			return currentHeight; // Image is not cropped

		var cropping = sourceRectangle.ToCropping();
		return currentHeight * (1 / (1 - (cropping.TopFactor + cropping.BottomFactor)));
	}

	#endregion Get methods

	#region Set methods

	/// <summary>
	/// Resets the cropping of the image.
	/// </summary>
	public static Drawing ResetCropping(this Drawing drawing)
	{
		A.SourceRectangle? sourceRectangle = drawing.GetSourceRectangle();

		if (sourceRectangle is null)
			return drawing; // Image is not cropped

		drawing.Resize(drawing.GetUncroppedWidth(), drawing.GetUncroppedHeight());
		sourceRectangle.Remove();
		return drawing;
	}

	/// <summary>
	/// Sets the cropping of the image.
	/// </summary>
	public static Drawing Cropping(this Drawing drawing, Cropping? cropping)
	{
		if (cropping is null)
		{
			drawing.ResetCropping();
			return drawing;
		}

		return drawing.Cropping(cropping.LeftFactor, cropping.TopFactor, cropping.RightFactor, cropping.BottomFactor);
	}

	/// <summary>
	/// Sets the cropping of the image on each side. Factor values are between 0 and 1 where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public static Drawing Cropping(this Drawing drawing,
		double leftFactor, double topFactor, double rightFactor, double bottomFactor)
	{
		return drawing
			.LeftCropping(leftFactor)
			.TopCropping(topFactor)
			.RightCropping(rightFactor)
			.BottomCropping(bottomFactor);
	}

	/// <summary>
	/// Sets the left cropping of the image, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public static Drawing LeftCropping(this Drawing drawing, double cropFactor)
	{
		A.SourceRectangle sourceRectangle = drawing.GetOrInitSourceRectangle();

		double originalWidth = drawing.GetUncroppedWidth();
		var currentCropping = sourceRectangle.ToCropping();

		double leftDifference = originalWidth * cropFactor;
		double rightDifference = originalWidth * currentCropping.RightFactor;
		double newWidth = originalWidth - leftDifference - rightDifference;
		newWidth = Math.Round(newWidth);

		sourceRectangle.Left = (int)(cropFactor * CONSTS.PERCENTAGE_MULTIPLIER);
		drawing.Resize(newWidth, drawing.GetHeight());

		return drawing;
	}

	/// <summary>
	/// Sets the right cropping of the image, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public static Drawing RightCropping(this Drawing drawing, double cropFactor)
	{
		A.SourceRectangle sourceRectangle = drawing.GetOrInitSourceRectangle();

		double originalWidth = drawing.GetUncroppedWidth();
		var currentCropping = sourceRectangle.ToCropping();

		double leftDifference = originalWidth * currentCropping.LeftFactor;
		double rightDifference = originalWidth * cropFactor;
		double newWidth = originalWidth - leftDifference - rightDifference;
		newWidth = Math.Round(newWidth);

		sourceRectangle.Right = (int)(cropFactor * CONSTS.PERCENTAGE_MULTIPLIER);
		drawing.Resize(newWidth, drawing.GetHeight());

		return drawing;
	}

	/// <summary>
	/// Sets the top cropping of the image, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public static Drawing TopCropping(this Drawing drawing, double cropFactor)
	{
		A.SourceRectangle sourceRectangle = drawing.GetOrInitSourceRectangle();

		double originalHeight = drawing.GetUncroppedHeight();
		var currentCropping = sourceRectangle.ToCropping();

		double topDifference = originalHeight * cropFactor;
		double bottomDifference = originalHeight * currentCropping.BottomFactor;
		double newHeight = originalHeight - topDifference - bottomDifference;
		newHeight = Math.Round(newHeight);

		sourceRectangle.Top = (int)(cropFactor * CONSTS.PERCENTAGE_MULTIPLIER);
		drawing.Resize(drawing.GetWidth(), newHeight);

		return drawing;
	}

	/// <summary>
	/// Sets the bottom cropping of the image, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public static Drawing BottomCropping(this Drawing drawing, double cropFactor)
	{
		A.SourceRectangle sourceRectangle = drawing.GetOrInitSourceRectangle();

		double originalHeight = drawing.GetUncroppedHeight();
		var currentCropping = sourceRectangle.ToCropping();

		double topDifference = originalHeight * currentCropping.TopFactor;
		double bottomDifference = originalHeight * cropFactor;
		double newHeight = originalHeight - topDifference - bottomDifference;
		newHeight = Math.Round(newHeight);

		sourceRectangle.Bottom = (int)(cropFactor * CONSTS.PERCENTAGE_MULTIPLIER);
		drawing.Resize(drawing.GetWidth(), newHeight);

		return drawing;
	}

	#endregion Set methods
}
