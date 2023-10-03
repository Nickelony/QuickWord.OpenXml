using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;
using QuickWord.OpenXml.Utilities;
using System.Collections.Generic;
using System.Linq;

namespace QuickWord.OpenXml.Extras.Extensions;

/// <summary>
/// Additional extension methods for the <see cref="Body"/> class.
/// </summary>
public static class BodyExtraExtensions
{
	public static IEnumerable<Paragraph> Paragraphs(this Body body)
		=> body.Elements<Paragraph>();

	public static Paragraph? Paragraphs(this Body body, int index)
		=> body.Elements<Paragraph>().ElementAtOrDefault(index);

	#region Margins

	public static double? LeftMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Left?.Value is uint value ? Twips.ToOther((int)value, desiredUnits) : null;

	public static double? TopMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Top?.Value is int value ? Twips.ToOther(value, desiredUnits) : null;

	public static double? RightMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Right?.Value is uint value ? Twips.ToOther((int)value, desiredUnits) : null;

	public static double? BottomMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Bottom?.Value is int value ? Twips.ToOther(value, desiredUnits) : null;

	public static double? HeaderMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Header?.Value is uint value ? Twips.ToOther((int)value, desiredUnits) : null;

	public static double? FooterMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Footer?.Value is uint value ? Twips.ToOther((int)value, desiredUnits) : null;

	public static double? GutterMarginValue(this Body body, MeasuringUnits desiredUnits)
		=> body.GetPageMargin()?.Gutter?.Value is uint value ? Twips.ToOther((int)value, desiredUnits) : null;

	public static Body LeftMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Left = (uint)Twips.FromOther(size, units);
		return body;
	}

	public static Body TopMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Top = Twips.FromOther(size, units);
		return body;
	}

	public static Body RightMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Right = (uint)Twips.FromOther(size, units);
		return body;
	}

	public static Body BottomMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Bottom = Twips.FromOther(size, units);
		return body;
	}

	public static Body HeaderMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Header = (uint)Twips.FromOther(size, units);
		return body;
	}

	public static Body FooterMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Footer = (uint)Twips.FromOther(size, units);
		return body;
	}

	public static Body GutterMargin(this Body body, double size, MeasuringUnits units = MeasuringUnits.Points)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageMargin>().Gutter = (uint)Twips.FromOther(size, units);
		return body;
	}

	#endregion Margins

	#region Page sizes

	private const int LETTER_TWIPS_WIDTH = 12240;
	private const int LETTER_TWIPS_HEIGHT = 15840;

	public static double PageWidthValue(this Body body, MeasuringUnits desiredUnits)
	{
		PageSize? pageSize = body.GetSectionProperties()?.GetFirstChild<PageSize>();

		if (pageSize is null)
			return Twips.ToOther(LETTER_TWIPS_WIDTH, desiredUnits);

		int value = (int?)pageSize.Width?.Value ?? LETTER_TWIPS_WIDTH;
		return Twips.ToOther(value, desiredUnits);
	}

	public static double PageHeightValue(this Body body, MeasuringUnits desiredUnits)
	{
		PageSize? pageSize = body.GetSectionProperties()?.GetFirstChild<PageSize>();

		if (pageSize is null)
			return Twips.ToOther(LETTER_TWIPS_HEIGHT, desiredUnits);

		int value = (int?)pageSize.Height?.Value ?? LETTER_TWIPS_HEIGHT;
		return Twips.ToOther(value, desiredUnits);
	}

	public static Body PageWidth(this Body body, double size, MeasuringUnits units)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageSize>().Width = (uint)Twips.FromOther(size, units);
		return body;
	}

	public static Body PageHeight(this Body body, double size, MeasuringUnits units)
	{
		body.GetOrInit<SectionProperties>().GetOrInit<PageSize>().Height = (uint)Twips.FromOther(size, units);
		return body;
	}

	#endregion Page sizes
}
