using DocumentFormat.OpenXml.Wordprocessing;
using System.Drawing;
using QuickWord.OpenXml.Utilities;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingEffectExtensions
{
	/// <summary>
	/// Gets the opacity of the image as a value between 0.0 (fully invisible) and 1.0 (fully visible).
	/// <para>Returns <see langword="null" /> if the property is not set.</para>
	/// </summary>
	public static double? OpacityValue(this Drawing drawing)
		=> drawing.GetOrInitAlphaModulationFixed().Amount?.Value / CONSTS.PERCENTAGE_MULTIPLIER;

	/// <summary>
	/// Sets the opacity of the image with a value between 0.0 (fully invisible) and 1.0 (fully visible).
	/// </summary>
	public static Drawing Opacity(this Drawing drawing, double opacity)
	{
		drawing.GetOrInitAlphaModulationFixed().Amount = (int)(opacity * CONSTS.PERCENTAGE_MULTIPLIER);
		return drawing;
	}

	// TODO: Add more border customization options. Possibly allow passing A.Outline objects and write extensions for it?
	public static Drawing Border(this Drawing drawing, double? width, string htmlColor = "black")
	{
		PIC.ShapeProperties shapeProperties = drawing.GetOrInitShapeProperties();
		shapeProperties.RemoveAllChildren<A.Outline>();

		if (width is null)
		{
			if (shapeProperties.ChildElements.Count == 0)
				shapeProperties.Remove();

			return drawing;
		}

		string parsedColor = ColorTranslator.FromHtml(htmlColor).ToHex().TrimStart('#');

		var outline = new A.Outline
		(
			new A.SolidFill() { RgbColorModelHex = new A.RgbColorModelHex() { Val = parsedColor } },
			new A.Miter()
		)
		{ Width = (int)(width * 12700) }; // EMU

		shapeProperties.Append(outline);
		drawing.SetEffectExtent((long)width + 1, (long)width + 1);

		return drawing;
	}
}
