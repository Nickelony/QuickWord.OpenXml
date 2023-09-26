using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;

namespace QuickWord.Tests;

[TestClass]
public class TableRowTests
{
	[TestMethod]
	public void CantSplit()
	{
		TableRow row = new TableRow().CantSplit(false);

		Assert.IsFalse(row.CantSplitValue());
		Assert.IsFalse(row.TableRowProperties!.GetFirstChild<CantSplit>()!.Val!.Value is OnOffOnlyValues.On);

		row.CantSplit(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void ConditionalFormatStyle()
	{
		TableRow row = new TableRow().ConditionalFormatStyle(new ConditionalFormatStyle { FirstRow = true });

		Assert.IsTrue(row.GetConditionalFormatStyle()!.FirstRow!.Value);
		Assert.IsTrue(row.TableRowProperties!.GetFirstChild<ConditionalFormatStyle>()!.FirstRow!.Value);

		row.ConditionalFormatStyle(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void DivId()
	{
		TableRow row = new TableRow().DivId("Test");

		Assert.AreEqual("Test", row.DivIdValue());
		Assert.AreEqual("Test", row.TableRowProperties!.GetFirstChild<DivId>()!.Val!.Value);

		row.DivId(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void GridAfter()
	{
		TableRow row = new TableRow().GridAfter(5);

		Assert.AreEqual(5, row.GridAfterValue());
		Assert.AreEqual(5, row.TableRowProperties!.GetFirstChild<GridAfter>()!.Val!.Value);

		row.GridAfter(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void GridBefore()
	{
		TableRow row = new TableRow().GridBefore(5);

		Assert.AreEqual(5, row.GridBeforeValue());
		Assert.AreEqual(5, row.TableRowProperties!.GetFirstChild<GridBefore>()!.Val!.Value);

		row.GridBefore(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void HideEndOfRowMarker()
	{
		TableRow row = new TableRow().HideEndOfRowMarker(false);

		Assert.IsFalse(row.HideEndOfRowMarkerValue());
		Assert.IsFalse(row.TableRowProperties!.GetFirstChild<Hidden>()!.Val!.Value);

		row.HideEndOfRowMarker(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void Justification()
	{
		TableRow row = new TableRow().Justification(TableRowAlignmentValues.Center);

		Assert.AreEqual(TableRowAlignmentValues.Center, row.JustificationValue());
		Assert.AreEqual(TableRowAlignmentValues.Center, row.TableRowProperties!.GetFirstChild<TableJustification>()!.Val!.Value);

		row.Justification(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void IsHeader()
	{
		TableRow row = new TableRow().IsHeader(false);

		Assert.IsFalse(row.IsHeaderValue());
		Assert.IsFalse(row.TableRowProperties!.GetFirstChild<TableHeader>()!.Val!.Value is OnOffOnlyValues.On);

		row.IsHeader(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void Height()
	{
		TableRow row = new TableRow().Height(new TableRowHeight { Val = 5 });

		Assert.AreEqual(5U, row.GetHeight()!.Val!.Value);
		Assert.AreEqual(5U, row.TableRowProperties!.GetFirstChild<TableRowHeight>()!.Val!.Value);

		row.Height(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void WidthAfter()
	{
		TableRow row = new TableRow().WidthAfter(new WidthAfterTableRow { Width = "100" });

		Assert.AreEqual("100", row.GetWidthAfter()!.Width!.Value);
		Assert.AreEqual("100", row.TableRowProperties!.GetFirstChild<WidthAfterTableRow>()!.Width!.Value);

		row.WidthAfter(null);
		Assert.IsNull(row.TableRowProperties);
	}

	[TestMethod]
	public void WidthBefore()
	{
		TableRow row = new TableRow().WidthBefore(new WidthBeforeTableRow { Width = "100" });

		Assert.AreEqual("100", row.GetWidthBefore()!.Width!.Value);
		Assert.AreEqual("100", row.TableRowProperties!.GetFirstChild<WidthBeforeTableRow>()!.Width!.Value);

		row.WidthBefore(null);
		Assert.IsNull(row.TableRowProperties);
	}
}
