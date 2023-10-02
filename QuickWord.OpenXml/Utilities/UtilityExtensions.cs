using DocumentFormat.OpenXml.Drawing;
using QuickWord.OpenXml.DrawingExtensions;
using System.Drawing;

namespace QuickWord.OpenXml.Utilities;

public static class UtilityExtensions
{
	public static string ToHex(this Color c)
		=> $"#{c.R:X2}{c.G:X2}{c.B:X2}";

	/// <summary>
	/// Creates a new <see cref="Cropping" /> object, based on the values of the given <see cref="SourceRectangle"/>.
	/// </summary>
	/// <returns>The <see cref="Cropping" /> object.</returns>
	public static Cropping ToCropping(this SourceRectangle rectangle) => new()
	{
		LeftFactor = rectangle.Left?.Value / CONSTS.PERCENTAGE_MULTIPLIER ?? 0,
		TopFactor = rectangle.Top?.Value / CONSTS.PERCENTAGE_MULTIPLIER ?? 0,
		RightFactor = rectangle.Right?.Value / CONSTS.PERCENTAGE_MULTIPLIER ?? 0,
		BottomFactor = rectangle.Bottom?.Value / CONSTS.PERCENTAGE_MULTIPLIER ?? 0
	};
}
