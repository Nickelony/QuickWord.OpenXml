using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

public static class BodyExtensions
{
	public static SectionProperties? GetSectionProperties(this Body body)
		=> body.GetFirstChild<SectionProperties>();

	public static PageMargin? GetPageMargin(this Body body)
		=> body.GetSectionProperties()?.GetFirstChild<PageMargin>();

	public static PageSize? GetPageSize(this Body body)
		=> body.GetSectionProperties()?.GetFirstChild<PageSize>();

	public static Body PageMargin(this Body body, PageMargin? margin)
	{
		body.GetOrInit<SectionProperties>().SetPropertyClassOrRemove(margin);
		return body;
	}

	public static Body PageSize(this Body body, PageSize? size)
	{
		body.GetOrInit<SectionProperties>().SetPropertyClassOrRemove(size);
		return body;
	}
}
