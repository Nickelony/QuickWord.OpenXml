namespace QuickWord.OpenXml.Measurements;

public static class BorderSize
{
	/// <summary>
	/// <c>1px => 6</c>
	/// </summary>
	public static uint ToSixth(double borderSize)
		 => (uint)(borderSize * 6);

	/// <summary>
	/// <c>6 => 1px</c>
	/// </summary>
	public static double FromSixth(uint sixthBorderSize)
		 => sixthBorderSize / 6.0;
}
