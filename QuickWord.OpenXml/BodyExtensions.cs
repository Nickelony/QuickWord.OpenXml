using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Utilities;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="Body"/> class.
/// </summary>
public static class BodyExtensions
{
	/// <summary>
	/// Gets the <see cref="SectionProperties" /> object of the <see cref="Body" />.
	/// <para>Returns <see langword="null" /> if the node doesn't exist.</para>
	/// </summary>
	/// <param name="body"></param>
	/// <returns></returns>
	public static SectionProperties? GetSectionProperties(this Body body)
		=> body.GetFirstChild<SectionProperties>();

	/// <summary>
	/// Specifies the page margins for all pages in this section.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.PageMargin" /></para>
	/// </summary>
	public static PageMargin? GetPageMargin(this Body body)
		=> body.GetSectionProperties()?.GetFirstChild<PageMargin>();

	/// <summary>
	/// Specifies the properties (size and orientation) for all pages in the current section.
	/// <para>Returns <see langword="null" /> if the property node doesn't exist.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.PageSize" /></para>
	/// </summary>
	public static PageSize? GetPageSize(this Body body)
		=> body.GetSectionProperties()?.GetFirstChild<PageSize>();

	/// <summary>
	/// Specifies the page margins for all pages in this section.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.PageMargin" /></para>
	/// </summary>
	public static Body PageMargin(this Body body, PageMargin? margin)
	{
		body.GetOrInit<SectionProperties>().SetPropertyClassOrRemove(margin);
		return body;
	}

	/// <summary>
	/// Specifies the properties (size and orientation) for all pages in the current section.
	/// <para>Setting this to <see langword="null" /> will remove the property node from the document.</para>
	/// <para><see href="https://learn.microsoft.com/en-us/dotnet/api/DocumentFormat.OpenXml.Wordprocessing.PageSize" /></para>
	/// </summary>
	public static Body PageSize(this Body body, PageSize? size)
	{
		body.GetOrInit<SectionProperties>().SetPropertyClassOrRemove(size);
		return body;
	}
}
