using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Linq;
using QuickWord.OpenXml.DrawingExtensions;

namespace QuickWord.OpenXml.QuickObjects;

public static class QDrawing
{
	private const string INVALID_DOCUMENT = "Couldn't locate MainDocumentPart.";

	public static Drawing FromImage(Body body, string fileName, ImagePartType type)
	{
		MainDocumentPart mainDocumentPart = body.Ancestors<Document>().FirstOrDefault()?.MainDocumentPart
			?? throw new InvalidOperationException(INVALID_DOCUMENT);

		return mainDocumentPart.CreateImage(fileName, type);
	}

	public static Drawing FromImage(Body body, string fileName, ImagePartType type, int width, int height)
	{
		MainDocumentPart mainDocumentPart = body.Ancestors<Document>().FirstOrDefault()?.MainDocumentPart
			?? throw new InvalidOperationException(INVALID_DOCUMENT);

		return mainDocumentPart.CreateImage(fileName, type, width, height);
	}
}
