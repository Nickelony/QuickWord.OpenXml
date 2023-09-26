namespace QuickWord.OpenXml.Measurements;

public static class Pixels
{
	public static double ToOther(double pixels, ImageMeasuringUnits desiredUnits) => desiredUnits switch
	{
		ImageMeasuringUnits.Centimeters => pixels / (96 / 2.54), // ~37.8px => 1cm
		ImageMeasuringUnits.Inches => pixels / 96, // 96px => 1"
		_ => pixels // Same unit
	};

	public static double FromOther(double other, ImageMeasuringUnits fromUnits) => fromUnits switch
	{
		ImageMeasuringUnits.Centimeters => other * (96 / 2.54), // 1cm => ~37.8px
		ImageMeasuringUnits.Inches => other * 96, // 1" => 96px
		_ => other // Same unit
	};
}
