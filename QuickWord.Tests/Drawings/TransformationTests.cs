using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.DrawingExtensions;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace QuickWord.Tests.Drawings;

[TestClass]
public class TransformationTests
{
	private WordprocessingDocument _wordDocument = null!;
	private Drawing _drawing = null!;

	[TestInitialize]
	public void Initialize()
	{
		_wordDocument = WordprocessingDocument.Create("Test.docx", WordprocessingDocumentType.Document);
		Body body = _wordDocument.CreateBody();

		_drawing = _wordDocument.MainDocumentPart!.CreateImage("Assets/Icon.png", ImagePartType.Png);
		body.AppendChild(new Paragraph(_drawing));
	}

	[TestMethod]
	public void Width()
	{
		_drawing.SetWidth(512);
		Assert.AreEqual(512, _drawing.GetWidth());

		_drawing.SetWidth(1, ImageMeasuringUnits.Inches);
		Assert.AreEqual(1, _drawing.GetWidth(ImageMeasuringUnits.Inches));
		Assert.AreEqual(2.54, _drawing.GetWidth(ImageMeasuringUnits.Centimeters));

		_drawing.SetWidth(5.08, ImageMeasuringUnits.Centimeters);
		Assert.AreEqual(5.08, _drawing.GetWidth(ImageMeasuringUnits.Centimeters));
		Assert.AreEqual(2, _drawing.GetWidth(ImageMeasuringUnits.Inches));
	}

	[TestMethod]
	public void Height()
	{
		_drawing.SetHeight(512);
		Assert.AreEqual(512, _drawing.GetHeight());

		_drawing.SetHeight(1, ImageMeasuringUnits.Inches);
		Assert.AreEqual(1, _drawing.GetHeight(ImageMeasuringUnits.Inches));
		Assert.AreEqual(2.54, _drawing.GetHeight(ImageMeasuringUnits.Centimeters));

		_drawing.SetHeight(5.08, ImageMeasuringUnits.Centimeters);
		Assert.AreEqual(5.08, _drawing.GetHeight(ImageMeasuringUnits.Centimeters));
		Assert.AreEqual(2, _drawing.GetHeight(ImageMeasuringUnits.Inches));
	}

	[TestMethod]
	public void Resize()
	{
		_drawing.Resize(512, 512);
		Assert.AreEqual(512, _drawing.GetWidth());
		Assert.AreEqual(512, _drawing.GetHeight());

		_drawing.Resize(1, 1, ImageMeasuringUnits.Inches);
		Assert.AreEqual(1, _drawing.GetWidth(ImageMeasuringUnits.Inches));
		Assert.AreEqual(1, _drawing.GetHeight(ImageMeasuringUnits.Inches));
		Assert.AreEqual(2.54, _drawing.GetWidth(ImageMeasuringUnits.Centimeters));
		Assert.AreEqual(2.54, _drawing.GetHeight(ImageMeasuringUnits.Centimeters));

		_drawing.Resize(5.08, 5.08, ImageMeasuringUnits.Centimeters);
		Assert.AreEqual(5.08, _drawing.GetWidth(ImageMeasuringUnits.Centimeters));
		Assert.AreEqual(5.08, _drawing.GetHeight(ImageMeasuringUnits.Centimeters));
		Assert.AreEqual(2, _drawing.GetWidth(ImageMeasuringUnits.Inches));
		Assert.AreEqual(2, _drawing.GetHeight(ImageMeasuringUnits.Inches));
	}

	[TestMethod]
	public void ResetSize()
	{
		// Image size is 1024x1024
		_drawing.Resize(512, 512);
		Assert.AreEqual(512, _drawing.GetWidth());
		Assert.AreEqual(512, _drawing.GetHeight());

		_drawing.ResetSize();
		Assert.AreEqual(1024, _drawing.GetWidth());
		Assert.AreEqual(1024, _drawing.GetHeight());

		// Scale + Cropping
		_drawing.Scale(0.5, 0.5); // Image size is now 512x512
		_drawing.Cropping(0.25, 0.25, 0.25, 0.25); // Image size is now 256x256 (Cropped) and 512x512 (Uncropped)

		Assert.AreEqual(256, _drawing.GetWidth());
		Assert.AreEqual(256, _drawing.GetHeight());
		Assert.AreEqual(512, _drawing.GetUncroppedWidth());
		Assert.AreEqual(512, _drawing.GetUncroppedHeight());

		_drawing.ResetSize();
		Assert.AreEqual(512, _drawing.GetWidth());
		Assert.AreEqual(512, _drawing.GetHeight());
		Assert.AreEqual(1024, _drawing.GetUncroppedWidth());
		Assert.AreEqual(1024, _drawing.GetUncroppedHeight());

		_drawing.ResetCropping();
		Assert.AreEqual(1024, _drawing.GetWidth());
		Assert.AreEqual(1024, _drawing.GetHeight());
		Assert.AreEqual(1024, _drawing.GetUncroppedWidth());
		Assert.AreEqual(1024, _drawing.GetUncroppedHeight());
	}

	[TestMethod]
	public void Scale()
	{
		_drawing.Scale(0.5, 0.25);
		Assert.AreEqual(512, _drawing.GetWidth());
		Assert.AreEqual(256, _drawing.GetHeight());

		_drawing.ResetSize();
		_drawing.ScaleHorizontally(0.5, true);
		Assert.AreEqual(512, _drawing.GetWidth());
		Assert.AreEqual(512, _drawing.GetHeight());

		_drawing.ResetSize();
		_drawing.ScaleVertically(0.5, true);
		Assert.AreEqual(512, _drawing.GetWidth());
		Assert.AreEqual(512, _drawing.GetHeight());
	}

	[TestMethod]
	public void Rotation()
	{
		_drawing.Rotation(90);
		Assert.AreEqual(90, _drawing.RotationValue());

		_drawing.Rotation(180);
		Assert.AreEqual(180, _drawing.RotationValue());

		_drawing.Rotation(270);
		Assert.AreEqual(270, _drawing.RotationValue());
	}

	[TestMethod]
	public void Flip()
	{
		_drawing.FlipHorizontally();
		Assert.IsTrue(_drawing.FlipHorizontallyValue());

		_drawing.FlipVertically();
		Assert.IsTrue(_drawing.FlipVerticallyValue());

		_drawing.FlipHorizontally(false);
		Assert.IsFalse(_drawing.FlipHorizontallyValue());

		_drawing.FlipVertically(false);
		Assert.IsFalse(_drawing.FlipVerticallyValue());
	}

	[TestMethod]
	public void AbsolutePositions()
	{
		_drawing.ToAnchoredDrawing();

		_drawing.AbsoluteHorizontalPosition(96, ImageMeasuringUnits.Pixels, DW.HorizontalRelativePositionValues.Margin);
		Assert.AreEqual(96, _drawing.AbsoluteHorizontalPositionValue(ImageMeasuringUnits.Pixels, out DW.HorizontalRelativePositionValues? toTheRightOf));
		Assert.AreEqual(DW.HorizontalRelativePositionValues.Margin, toTheRightOf);
		Assert.AreEqual(1, _drawing.AbsoluteHorizontalPositionValue(ImageMeasuringUnits.Inches, out _));
		Assert.AreEqual(2.54, _drawing.AbsoluteHorizontalPositionValue(ImageMeasuringUnits.Centimeters, out _));

		_drawing.AbsoluteVerticalPosition(96, ImageMeasuringUnits.Pixels, DW.VerticalRelativePositionValues.Margin);
		Assert.AreEqual(96, _drawing.AbsoluteVerticalPositionValue(ImageMeasuringUnits.Pixels, out DW.VerticalRelativePositionValues? below));
		Assert.AreEqual(DW.VerticalRelativePositionValues.Margin, below);
		Assert.AreEqual(1, _drawing.AbsoluteVerticalPositionValue(ImageMeasuringUnits.Inches, out _));
		Assert.AreEqual(2.54, _drawing.AbsoluteVerticalPositionValue(ImageMeasuringUnits.Centimeters, out _));
	}

	[TestMethod]
	public void Alignments()
	{
		_drawing.ToAnchoredDrawing();

		_drawing.HorizontalAlignment(DW.HorizontalAlignmentValues.Center, DW.HorizontalRelativePositionValues.Margin);
		Assert.AreEqual(DW.HorizontalAlignmentValues.Center, _drawing.HorizontalAlignmentValue(out DW.HorizontalRelativePositionValues? hRelativeTo));
		Assert.AreEqual(DW.HorizontalRelativePositionValues.Margin, hRelativeTo);

		_drawing.VerticalAlignment(DW.VerticalAlignmentValues.Center, DW.VerticalRelativePositionValues.Margin);
		Assert.AreEqual(DW.VerticalAlignmentValues.Center, _drawing.VerticalAlignmentValue(out DW.VerticalRelativePositionValues? vRelativeTo));
		Assert.AreEqual(DW.VerticalRelativePositionValues.Margin, vRelativeTo);
	}

	[TestCleanup]
	public void Cleanup()
		=> _wordDocument.Dispose();
}
