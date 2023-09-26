using DocumentFormat.OpenXml.Wordprocessing;
using System;

namespace QuickWord.OpenXml.DrawingExtensions;

public static class DrawingAnchorSpecificExtensions
{
	/// <summary>
	/// Specifies whether objects shall be allowed to overlap the drawing.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static bool? AllowOverlappingValue(this Drawing drawing) => !drawing.IsAnchored()
		? throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED)
		: drawing.Anchor?.AllowOverlap?.Value;

	/// <summary>
	/// Specifies whether the drawing shall be displayed behind the document text.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static bool? BehindTextValue(this Drawing drawing) => !drawing.IsAnchored()
		? throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED)
		: drawing.Anchor?.BehindDoc?.Value;

	/// <summary>
	/// The "Layout in table cell" option is supplied for backward compatibility with
	/// older versions of Word and older file formats. For example, in a *.doc file,
	/// you can't put an image whose Text Wrapping is set to "Square" inside
	/// a table cell unless the "Layout in table cell" option is selected.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static bool? LayoutInTableCellValue(this Drawing drawing) => !drawing.IsAnchored()
		? throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED)
		: drawing.Anchor?.LayoutInCell?.Value;

	/// <summary>
	/// Specifies whether the drawing's anchor shall be locked.
	/// </summary>
	/// <exception cref="InvalidOperationException">Occurs when the drawing is not anchored.</exception>
	public static bool? LockedValue(this Drawing drawing) => !drawing.IsAnchored()
		? throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED)
		: drawing.Anchor?.Locked?.Value;

	/// <inheritdoc cref="AllowOverlappingValue" />
	public static Drawing AllowOverlapping(this Drawing drawing, bool value = true)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.Anchor!.AllowOverlap = value;
		return drawing;
	}

	/// <inheritdoc cref="BehindTextValue" />
	public static Drawing BehindText(this Drawing drawing, bool value = true)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.Anchor!.BehindDoc = value;
		return drawing;
	}

	/// <inheritdoc cref="LayoutInTableCellValue" />
	public static Drawing LayoutInTableCell(this Drawing drawing, bool value = true)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.Anchor!.LayoutInCell = value;
		return drawing;
	}

	/// <inheritdoc cref="LockedValue" />
	public static Drawing Locked(this Drawing drawing, bool value = true)
	{
		if (!drawing.IsAnchored())
			throw new InvalidOperationException(CONSTS.DRAWING_NOT_ANCHORED);

		drawing.Anchor!.Locked = value;
		return drawing;
	}
}
