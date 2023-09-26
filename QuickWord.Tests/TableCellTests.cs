using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;

namespace QuickWord.Tests;

[TestClass]
public class TableCellTests
{
	[TestMethod]
	public void ConditionalFormatStyle()
	{
		TableCell cell = new TableCell().ConditionalFormatStyle(new ConditionalFormatStyle { FirstRow = true });

		Assert.IsTrue(cell.GetConditionalFormatStyle()!.FirstRow!.Value);
		Assert.IsTrue(cell.TableCellProperties!.ConditionalFormatStyle!.FirstRow!.Value);

		cell.ConditionalFormatStyle(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void GridSpan()
	{
		TableCell cell = new TableCell().GridSpan(5);

		Assert.AreEqual(5, cell.GridSpanValue());
		Assert.AreEqual(5, cell.TableCellProperties!.GridSpan!.Val!.Value);

		cell.GridSpan(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void HideEndOfCellMarker()
	{
		TableCell cell = new TableCell().HideEndOfCellMarker(false);

		Assert.IsFalse(cell.HideEndOfCellMarkerValue());
		Assert.IsFalse(cell.TableCellProperties!.HideMark!.Val!.Value is OnOffOnlyValues.On);

		cell.HideEndOfCellMarker(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void HorizontalMerge()
	{
		TableCell cell = new TableCell().HorizontalMerge(MergedCellValues.Restart);

		Assert.AreEqual(MergedCellValues.Restart, cell.HorizontalMergeValue());
		Assert.AreEqual(MergedCellValues.Restart, cell.TableCellProperties!.HorizontalMerge!.Val!.Value);

		cell.HorizontalMerge(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void NoContentWrapping()
	{
		TableCell cell = new TableCell().NoContentWrapping(false);

		Assert.IsFalse(cell.NoContentWrappingValue());
		Assert.IsFalse(cell.TableCellProperties!.NoWrap!.Val!.Value is OnOffOnlyValues.On);

		cell.NoContentWrapping(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void Shading()
	{
		TableCell cell = new TableCell().Shading(new Shading { Fill = "Red" });

		Assert.AreEqual("Red", cell.GetShading()!.Fill!.Value);
		Assert.AreEqual("Red", cell.TableCellProperties!.Shading!.Fill!.Value);

		cell.Shading(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void Borders()
	{
		TableCell cell = new TableCell().Borders(new TableCellBorders { TopBorder = new TopBorder { Val = BorderValues.Single } });

		Assert.AreEqual(BorderValues.Single, cell.GetBorders()!.TopBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.Single, cell.TableCellProperties!.TableCellBorders!.TopBorder!.Val!.Value);

		cell.Borders(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void FitText()
	{
		TableCell cell = new TableCell().FitText(false);

		Assert.IsFalse(cell.FitTextValue());
		Assert.IsFalse(cell.TableCellProperties!.TableCellFitText!.Val!.Value is OnOffOnlyValues.On);

		cell.FitText(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void Margins()
	{
		TableCell cell = new TableCell().Margin(new TableCellMargin { TopMargin = new TopMargin { Width = "10" } });

		Assert.AreEqual("10", cell.GetMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("10", cell.TableCellProperties!.TableCellMargin!.TopMargin!.Width!.Value);

		cell.Margin(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void Width()
	{
		TableCell cell = new TableCell().Width(new TableCellWidth { Width = "10" });

		Assert.AreEqual("10", cell.GetWidth()!.Width!.Value);
		Assert.AreEqual("10", cell.TableCellProperties!.TableCellWidth!.Width!.Value);

		cell.Width(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void TextDirection()
	{
		TableCell cell = new TableCell().TextDirection(TextDirectionValues.BottomToTopLeftToRight);

		Assert.AreEqual(TextDirectionValues.BottomToTopLeftToRight, cell.TextDirectionValue());
		Assert.AreEqual(TextDirectionValues.BottomToTopLeftToRight, cell.TableCellProperties!.TextDirection!.Val!.Value);

		cell.TextDirection(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void VerticalContentAlignment()
	{
		TableCell cell = new TableCell().VerticalContentAlignment(TableVerticalAlignmentValues.Center);

		Assert.AreEqual(TableVerticalAlignmentValues.Center, cell.VerticalContentAlignmentValue());
		Assert.AreEqual(TableVerticalAlignmentValues.Center, cell.TableCellProperties!.TableCellVerticalAlignment!.Val!.Value);

		cell.VerticalContentAlignment(null);
		Assert.IsNull(cell.TableCellProperties);
	}

	[TestMethod]
	public void VerticalMerge()
	{
		TableCell cell = new TableCell().VerticalMerge(MergedCellValues.Restart);

		Assert.AreEqual(MergedCellValues.Restart, cell.VerticalMergeValue());
		Assert.AreEqual(MergedCellValues.Restart, cell.TableCellProperties!.VerticalMerge!.Val!.Value);

		cell.VerticalMerge(null);
		Assert.IsNull(cell.TableCellProperties);
	}
}
