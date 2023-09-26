using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;

namespace QuickWord.Tests;

[TestClass]
public class TableTests
{
	[TestMethod]
	public void VisuallyBiDirectional()
	{
		Table table = new Table().VisuallyBiDirectional(false);

		Assert.IsFalse(table.VisuallyBiDirectionalValue());
		Assert.IsFalse(table.GetTableProperties()!.BiDiVisual!.Val!.Value is OnOffOnlyValues.On);

		table.VisuallyBiDirectional(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Justification()
	{
		Table table = new Table().Justification(TableRowAlignmentValues.Center);

		Assert.AreEqual(TableRowAlignmentValues.Center, table.JustificationValue());
		Assert.AreEqual(TableRowAlignmentValues.Center, table.GetTableProperties()!.TableJustification!.Val!.Value);

		table.Justification(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Shading()
	{
		Table table = new Table().Shading(new Shading { Fill = "Red" });

		Assert.AreEqual("Red", table.GetShading()!.Fill!.Value);
		Assert.AreEqual("Red", table.GetTableProperties()!.Shading!.Fill!.Value);

		table.Shading(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Borders()
	{
		Table table = new Table().Borders(new TableBorders(new LeftBorder { Color = "Red" }));

		Assert.AreEqual("Red", table.GetBorders()!.LeftBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetTableProperties()!.TableBorders!.LeftBorder!.Color!.Value);

		table.Borders(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Caption()
	{
		Table table = new Table().Caption("Test");

		Assert.AreEqual("Test", table.CaptionValue());
		Assert.AreEqual("Test", table.GetTableProperties()!.TableCaption!.Val!.Value);

		table.Caption(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void DefaultCellMargins()
	{
		Table table = new Table().DefaultCellMargins(new TableCellMarginDefault(new TableCellLeftMargin { Width = 66 }));

		Assert.AreEqual(66, table.GetDefaultCellMargins()!.TableCellLeftMargin!.Width!.Value);
		Assert.AreEqual(66, table.GetTableProperties()!.TableCellMarginDefault!.TableCellLeftMargin!.Width!.Value);

		table.DefaultCellMargins(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void CellSpacing()
	{
		Table table = new Table().CellSpacing(new TableCellSpacing { Width = "66" });

		Assert.AreEqual("66", table.GetCellSpacing()!.Width!.Value);
		Assert.AreEqual("66", table.GetTableProperties()!.TableCellSpacing!.Width!.Value);

		table.CellSpacing(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Description()
	{
		Table table = new Table().Description("Test");

		Assert.AreEqual("Test", table.DescriptionValue());
		Assert.AreEqual("Test", table.GetTableProperties()!.TableDescription!.Val!.Value);

		table.Description(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Indentation()
	{
		Table table = new Table().Indentation(new TableIndentation { Width = 66 });

		Assert.AreEqual(66, table.GetIndentation()!.Width!.Value);
		Assert.AreEqual(66, table.GetTableProperties()!.TableIndentation!.Width!.Value);

		table.Indentation(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Layout()
	{
		Table table = new Table().Layout(TableLayoutValues.Fixed);

		Assert.AreEqual(TableLayoutValues.Fixed, table.LayoutValue());
		Assert.AreEqual(TableLayoutValues.Fixed, table.GetTableProperties()!.TableLayout!.Type!.Value);

		table.Layout(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Look()
	{
		Table table = new Table().Look(new TableLook { FirstColumn = true });

		Assert.IsTrue(table.GetLook()!.FirstColumn!.Value);
		Assert.IsTrue(table.GetTableProperties()!.TableLook!.FirstColumn!.Value);

		table.Look(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Overlap()
	{
		Table table = new Table().Overlap(TableOverlapValues.Overlap);

		Assert.AreEqual(TableOverlapValues.Overlap, table.OverlapValue());
		Assert.AreEqual(TableOverlapValues.Overlap, table.GetTableProperties()!.TableOverlap!.Val!.Value);

		table.Overlap(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void PositionProperties()
	{
		Table table = new Table().PositionProperties(new TablePositionProperties { BottomFromText = 66 });

		Assert.AreEqual(66, table.GetPositionProperties()!.BottomFromText!.Value);
		Assert.AreEqual(66, table.GetTableProperties()!.TablePositionProperties!.BottomFromText!.Value);

		table.PositionProperties(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Style()
	{
		Table table = new Table().Style("Heading1");

		Assert.AreEqual("Heading1", table.StyleValue());
		Assert.AreEqual("Heading1", table.GetTableProperties()!.TableStyle!.Val!.Value);

		table.Style(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void StyleColumnBandSize()
	{
		Table table = new Table().StyleColumnBandSize(66);

		Assert.AreEqual(66, table.StyleColumnBandSizeValue());
		Assert.AreEqual(66, table.GetTableProperties()!.GetFirstChild<TableStyleColumnBandSize>()!.Val!.Value);

		table.StyleColumnBandSize(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void StyleRowBandSize()
	{
		Table table = new Table().StyleRowBandSize(66);

		Assert.AreEqual(66, table.StyleRowBandSizeValue());
		Assert.AreEqual(66, table.GetTableProperties()!.GetFirstChild<TableStyleRowBandSize>()!.Val!.Value);

		table.StyleRowBandSize(null);
		Assert.IsNull(table.GetTableProperties());
	}

	[TestMethod]
	public void Width()
	{
		Table table = new Table().Width(new TableWidth { Width = "66" });

		Assert.AreEqual("66", table.GetWidth()!.Width!.Value);
		Assert.AreEqual("66", table.GetTableProperties()!.TableWidth!.Width!.Value);

		table.Width(null);
		Assert.IsNull(table.GetTableProperties());
	}
}
