using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class MainDocumentPartExtensions
{
	/// <summary>
	/// Extracts all the images from the document.
	/// </summary>
	/// <returns>A collection of <see cref="Stream" /> objects which hold the image data.</returns>
	public static IEnumerable<Stream> ExtractAllImages(this MainDocumentPart mainDocumentPart)
	{
		foreach (IdPartPair part in mainDocumentPart.Parts)
		{
			if (part.OpenXmlPart is ImagePart imagePart)
				yield return imagePart.GetStream();
		}
	}

	/// <summary>
	/// Creates a new <see cref="Drawing" /> object from the specified image file, while automatically determining the image dimensions.
	/// <para>This method uses a <see cref="Bitmap" /> object to measure the dimensions of the image, which is only supported on Windows.</para>
	/// </summary>
	/// <param name="fileName">File path of the image file.</param>
	/// <param name="type">File format of the image.</param>
	/// <returns>The <see cref="Drawing" /> object which can be inserted into a <see cref="Run" /> object.</returns>
	[SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
	public static Drawing CreateImage(this MainDocumentPart mainDocumentPart, string fileName, ImagePartType type)
	{
		int width, height;

		using (var bitmap = new Bitmap(fileName))
		{
			width = bitmap.Width;
			height = bitmap.Height;
		}

		return mainDocumentPart.CreateImage(fileName, type, width, height);
	}

	/// <summary>
	/// Creates a new <see cref="Drawing" /> object from the specified image file.
	/// </summary>
	/// <param name="fileName">File path of the image file.</param>
	/// <param name="type">File format of the image.</param>
	/// <param name="width">Desired width of the image.</param>
	/// <param name="height">Desired height of the image.</param>
	/// <returns>The <see cref="Drawing" /> object which can be inserted into a <see cref="Run" /> object.</returns>
	public static Drawing CreateImage(this MainDocumentPart mainDocumentPart,
		string fileName, ImagePartType type, int width, int height)
	{
		ImagePart imagePart = mainDocumentPart.AddImagePart(type);
		string imagePartId = mainDocumentPart.GetIdOfPart(imagePart);

		using (var stream = new FileStream(fileName, FileMode.Open))
			imagePart.FeedData(stream);

		return new Drawing
		(
			new DW.Inline
			(
				new DW.Extent()
				{
					Cx = (long)(width * CONSTS.EMU_PER_PIXEL),
					Cy = (long)(height * CONSTS.EMU_PER_PIXEL)
				},

				new DW.EffectExtent()
				{
					LeftEdge = 0,
					TopEdge = 0,
					RightEdge = 0,
					BottomEdge = 0
				},

				new DW.DocProperties()
				{
					Id = 1,
					Name = Path.GetFileName(fileName)
				},

				new DW.NonVisualGraphicFrameDrawingProperties
				(
					new A.GraphicFrameLocks() { NoChangeAspect = true }
				),

				new A.Graphic
				(
					new A.GraphicData
					(
						new PIC.Picture
						(
							new PIC.NonVisualPictureProperties
							(
								new PIC.NonVisualDrawingProperties()
								{
									Id = 0,
									Name = Path.GetFileName(fileName)
								},

								new PIC.NonVisualPictureDrawingProperties()
							),

							new PIC.BlipFill
							(
								new A.Blip() { Embed = imagePartId },

								new A.Stretch
								(
									new A.FillRectangle()
								)
							),

							new PIC.ShapeProperties
							(
								new A.Transform2D
								(
									new A.Offset()
									{
										X = 0,
										Y = 0
									},

									new A.Extents()
									{
										Cx = (long)(width * CONSTS.EMU_PER_PIXEL),
										Cy = (long)(height * CONSTS.EMU_PER_PIXEL)
									}
								),

								new A.PresetGeometry
								(
									new A.AdjustValueList()
								)
								{ Preset = A.ShapeTypeValues.Rectangle }
							)
							{ BlackWhiteMode = A.BlackWhiteModeValues.Auto }
						)
					)
					{ Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" }
				)
			)
		);
	}
}
