namespace QuickWord.OpenXml.DrawingExtensions;

public class Cropping
{
	/// <summary>
	/// Left cropping factor of the drawing, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public double LeftFactor { get; set; }

	/// <summary>
	/// Top cropping factor of the drawing, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public double TopFactor { get; set; }

	/// <summary>
	/// Right cropping factor of the drawing, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public double RightFactor { get; set; }

	/// <summary>
	/// Bottom cropping factor of the drawing, where 0.0 is no cropping and 1.0 is full cropping.
	/// </summary>
	public double BottomFactor { get; set; }
}
