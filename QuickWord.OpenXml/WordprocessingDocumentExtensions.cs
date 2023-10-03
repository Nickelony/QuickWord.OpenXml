// Ignore Spelling: Wordprocessing

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace QuickWord.OpenXml;

/// <summary>
/// A set of extension methods for the <see cref="WordprocessingDocument" /> class.
/// </summary>
public static class WordprocessingDocumentExtensions
{
	/// <summary>
	/// Creates a new body and adds it to the document.
	/// </summary>
	/// <returns>The <see cref="Body" /> object</returns>
	public static Body CreateBody(this WordprocessingDocument wordDocument)
	{
		MainDocumentPart mainDocumentPart = wordDocument.AddMainDocumentPart();
		var body = new Body();

		mainDocumentPart.Document = new Document(body);
		return body;
	}
}
