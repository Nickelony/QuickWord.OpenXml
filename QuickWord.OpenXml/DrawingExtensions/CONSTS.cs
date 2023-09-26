namespace QuickWord.OpenXml.DrawingExtensions;

internal static class CONSTS
{
	public const double ANGLE_MULTIPLIER = 60000;
	public const double EMU_PER_PIXEL = 9525;
	public const double PERCENTAGE_MULTIPLIER = 100000;

	public const string DRAWING_NOT_ANCHORED = "Drawing is not anchored.";
	public const string UNKNOWN_DRAWING_TYPE = "Drawing is neither inlined nor anchored.";
	public const string INVALID_EXTENT = "Couldn't get extent. Drawing is broken.";
	public const string COULD_NOT_FETCH_IMAGE_DATA = "Couldn't fetch original image data. Picture reference embed might be missing or the Drawing object is not inside a Document's Body.";
	public const string WRAPPING_NOT_FOUND = "No suitable wrapping property has been found. The document is broken.";
}
