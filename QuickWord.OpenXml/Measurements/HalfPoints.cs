namespace QuickWord.OpenXml.Measurements;

public static class HalfPoints
{
	public static double ToOther(int halfPoints, TextMeasuringUnits unit) => unit switch
	{
		TextMeasuringUnits.Points => halfPoints / 2.0, // 1hp => 0.5pt
		_ => halfPoints // Same unit
	};

	public static int FromOther(double other, TextMeasuringUnits unit) => unit switch
	{
		TextMeasuringUnits.Points => (int)(other * 2), // 1pt => 2hp
		_ => (int)other // Same unit
	};
}
