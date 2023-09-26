using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.Extras;

namespace QuickWord.Tests.Extras;

[TestClass]
public class TableExtrasTests
{
	[TestMethod]
	public void Formatting()
	{
		Table table = new Table()
			.VisuallyBiDirectional()
			.Justification(TableRowAlignmentValues.Center)
			.Shading(new Shading { Fill = "Red" });

		TableFormatting formatting = table.CloneFormatting();

		Assert.AreEqual(3, table.GetTableProperties()!.ChildElements.Count);

		table.ResetFormatting();
		Assert.AreEqual(0, table.ChildElements.Count);

		table.ApplyFormatting(formatting);
		Assert.IsTrue(table.VisuallyBiDirectionalValue());
		Assert.AreEqual(TableRowAlignmentValues.Center, table.JustificationValue());
		Assert.AreEqual("Red", table.GetShading()!.Fill!.Value);
		Assert.AreEqual(3, table.GetTableProperties()!.ChildElements.Count);

		var anotherFormatting = new TableFormatting { Caption = "Test", Shading = new Shading { Fill = "Blue" } };

		table.ApplyFormatting(anotherFormatting, true);
		Assert.IsTrue(table.VisuallyBiDirectionalValue());
		Assert.AreEqual(TableRowAlignmentValues.Center, table.JustificationValue());
		Assert.AreEqual("Test", table.CaptionValue());
		Assert.AreEqual("Blue", table.GetShading()!.Fill!.Value);
		Assert.AreEqual(4, table.GetTableProperties()!.ChildElements.Count);

		table.ApplyFormatting(anotherFormatting);
		Assert.IsNull(table.VisuallyBiDirectionalValue());
		Assert.IsNull(table.JustificationValue());
		Assert.AreEqual("Test", table.CaptionValue());
		Assert.AreEqual("Blue", table.GetShading()!.Fill!.Value);
		Assert.AreEqual(2, table.GetTableProperties()!.ChildElements.Count);

		table.ResetFormatting();
	}

	[TestMethod]
	public void Borders()
	{
		Table table = new Table()
			.LeftBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.TopBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.RightBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.BottomBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.InsideHorizontalBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.InsideVerticalBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.StartBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.EndBorder(1.5, BorderValues.DashDotStroked, "Red", 1);

		Assert.AreEqual(9U, table.GetBorders()!.LeftBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.TopBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.RightBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.BottomBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.InsideHorizontalBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.InsideVerticalBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.StartBorder!.Size!.Value);
		Assert.AreEqual(9U, table.GetBorders()!.EndBorder!.Size!.Value);

		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.LeftBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.TopBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.RightBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.BottomBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.InsideHorizontalBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.InsideVerticalBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.StartBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, table.GetBorders()!.EndBorder!.Val!.Value);

		Assert.AreEqual("Red", table.GetBorders()!.LeftBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.TopBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.RightBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.BottomBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.InsideHorizontalBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.InsideVerticalBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.StartBorder!.Color!.Value);
		Assert.AreEqual("Red", table.GetBorders()!.EndBorder!.Color!.Value);

		Assert.AreEqual(1U, table.GetBorders()!.LeftBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.TopBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.RightBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.BottomBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.InsideHorizontalBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.InsideVerticalBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.StartBorder!.Space!.Value);
		Assert.AreEqual(1U, table.GetBorders()!.EndBorder!.Space!.Value);

		table.ResetFormatting();

		table
			.LeftBorder(new LeftBorder { Size = 6 })
			.TopBorder(new TopBorder { Size = 6 })
			.RightBorder(new RightBorder { Size = 6 })
			.BottomBorder(new BottomBorder { Size = 6 })
			.InsideHorizontalBorder(new InsideHorizontalBorder { Size = 6 })
			.InsideVerticalBorder(new InsideVerticalBorder { Size = 6 })
			.StartBorder(new StartBorder { Size = 6 })
			.EndBorder(new EndBorder { Size = 6 });

		Assert.AreEqual(6U, table.GetBorders()!.LeftBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.TopBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.RightBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.BottomBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.InsideHorizontalBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.InsideVerticalBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.StartBorder!.Size!.Value);
		Assert.AreEqual(6U, table.GetBorders()!.EndBorder!.Size!.Value);
	}

	[TestMethod]
	public void DefaultCellMargins()
	{
		Table table = new Table()
			.DefaultTopMarginOfCells(6, WidthUnits.Points)
			.DefaultBottomMarginOfCells(6, WidthUnits.Points)
			.DefaultLeftMarginOfCells(6)
			.DefaultRightMarginOfCells(6)
			.DefaultStartMarginOfCells(6, WidthUnits.Points)
			.DefaultEndMarginOfCells(6, WidthUnits.Points);

		Assert.AreEqual("120", table.GetDefaultCellMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("120", table.GetDefaultCellMargins()!.BottomMargin!.Width!.Value);
		Assert.AreEqual(120, table.GetDefaultCellMargins()!.TableCellLeftMargin!.Width!.Value);
		Assert.AreEqual(120, table.GetDefaultCellMargins()!.TableCellRightMargin!.Width!.Value);
		Assert.AreEqual("120", table.GetDefaultCellMargins()!.StartMargin!.Width!.Value);
		Assert.AreEqual("120", table.GetDefaultCellMargins()!.EndMargin!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetDefaultCellMargins()!.TopMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetDefaultCellMargins()!.BottomMargin!.Type!.Value);
		Assert.AreEqual(TableWidthValues.Dxa, table.GetDefaultCellMargins()!.TableCellLeftMargin!.Type!.Value);
		Assert.AreEqual(TableWidthValues.Dxa, table.GetDefaultCellMargins()!.TableCellRightMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetDefaultCellMargins()!.StartMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetDefaultCellMargins()!.EndMargin!.Type!.Value);

		Assert.AreEqual(6, table.DefaultTopMarginOfCellsValue(out WidthUnits? topUnits));
		Assert.AreEqual(6, table.DefaultBottomMarginOfCellsValue(out WidthUnits? bottomUnits));
		Assert.AreEqual(6, table.DefaultLeftMarginOfCellsValue());
		Assert.AreEqual(6, table.DefaultRightMarginOfCellsValue());
		Assert.AreEqual(6, table.DefaultStartMarginOfCellsValue(out WidthUnits? startUnits));
		Assert.AreEqual(6, table.DefaultEndMarginOfCellsValue(out WidthUnits? endUnits));
		Assert.AreEqual(WidthUnits.Points, topUnits);
		Assert.AreEqual(WidthUnits.Points, bottomUnits);
		Assert.AreEqual(WidthUnits.Points, startUnits);
		Assert.AreEqual(WidthUnits.Points, endUnits);

		table.ResetFormatting();

		table
			.DefaultTopMarginOfCells(66.6, WidthUnits.Percentage)
			.DefaultBottomMarginOfCells(66.6, WidthUnits.Percentage)
			.DefaultStartMarginOfCells(66.6, WidthUnits.Percentage)
			.DefaultEndMarginOfCells(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", table.GetDefaultCellMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("3330", table.GetDefaultCellMargins()!.BottomMargin!.Width!.Value);
		Assert.AreEqual("3330", table.GetDefaultCellMargins()!.StartMargin!.Width!.Value);
		Assert.AreEqual("3330", table.GetDefaultCellMargins()!.EndMargin!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetDefaultCellMargins()!.TopMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetDefaultCellMargins()!.BottomMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetDefaultCellMargins()!.StartMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetDefaultCellMargins()!.EndMargin!.Type!.Value);

		Assert.AreEqual(66.6, table.DefaultTopMarginOfCellsValue(out topUnits));
		Assert.AreEqual(66.6, table.DefaultBottomMarginOfCellsValue(out bottomUnits));
		Assert.AreEqual(66.6, table.DefaultStartMarginOfCellsValue(out startUnits));
		Assert.AreEqual(66.6, table.DefaultEndMarginOfCellsValue(out endUnits));
		Assert.AreEqual(WidthUnits.Percentage, topUnits);
		Assert.AreEqual(WidthUnits.Percentage, bottomUnits);
		Assert.AreEqual(WidthUnits.Percentage, startUnits);
		Assert.AreEqual(WidthUnits.Percentage, endUnits);

		table.ResetFormatting();

		table
			.DefaultTopMarginOfCells(new TopMargin { Width = "240" })
			.DefaultBottomMarginOfCells(new BottomMargin { Width = "240" })
			.DefaultLeftMarginOfCells(new TableCellLeftMargin { Width = 240 })
			.DefaultRightMarginOfCells(new TableCellRightMargin { Width = 240 })
			.DefaultStartMarginOfCells(new StartMargin { Width = "240" })
			.DefaultEndMarginOfCells(new EndMargin { Width = "240" });

		Assert.AreEqual("240", table.GetDefaultCellMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("240", table.GetDefaultCellMargins()!.BottomMargin!.Width!.Value);
		Assert.AreEqual(240, table.GetDefaultCellMargins()!.TableCellLeftMargin!.Width!.Value);
		Assert.AreEqual(240, table.GetDefaultCellMargins()!.TableCellRightMargin!.Width!.Value);
		Assert.AreEqual("240", table.GetDefaultCellMargins()!.StartMargin!.Width!.Value);
		Assert.AreEqual("240", table.GetDefaultCellMargins()!.EndMargin!.Width!.Value);
	}

	[TestMethod]
	public void CellSpacing()
	{
		Table table = new Table()
			.CellSpacing(6, WidthUnits.Points);

		Assert.AreEqual("120", table.GetTableProperties()!.TableCellSpacing!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetTableProperties()!.TableCellSpacing!.Type!.Value);

		Assert.AreEqual(6, table.CellSpacingValue(out WidthUnits? units));
		Assert.AreEqual(WidthUnits.Points, units);

		table.ResetFormatting();

		table.CellSpacing(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", table.GetTableProperties()!.TableCellSpacing!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetTableProperties()!.TableCellSpacing!.Type!.Value);

		Assert.AreEqual(66.6, table.CellSpacingValue(out units));
		Assert.AreEqual(WidthUnits.Percentage, units);

		table.ResetFormatting();

		table.CellSpacing(new TableCellSpacing { Width = "240" });
		Assert.AreEqual("240", table.GetTableProperties()!.TableCellSpacing!.Width!.Value);
	}

	[TestMethod]
	public void Indentation()
	{
		Table table = new Table()
			.Indentation(6, WidthUnits.Points);

		Assert.AreEqual(120, table.GetTableProperties()!.TableIndentation!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetTableProperties()!.TableIndentation!.Type!.Value);

		Assert.AreEqual(6, table.IndentationValue(out WidthUnits? units));
		Assert.AreEqual(WidthUnits.Points, units);

		table.ResetFormatting();

		table.Indentation(66.6, WidthUnits.Percentage);

		Assert.AreEqual(3330, table.GetTableProperties()!.TableIndentation!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetTableProperties()!.TableIndentation!.Type!.Value);

		Assert.AreEqual(66.6, table.IndentationValue(out units));
		Assert.AreEqual(WidthUnits.Percentage, units);

		table.ResetFormatting();

		table.Indentation(new TableIndentation { Width = 240 });
		Assert.AreEqual(240, table.GetTableProperties()!.TableIndentation!.Width!.Value);
	}

	[TestMethod]
	public void Width()
	{
		Table table = new Table()
			.Width(6, WidthUnits.Points);

		Assert.AreEqual("120", table.GetTableProperties()!.TableWidth!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, table.GetTableProperties()!.TableWidth!.Type!.Value);

		Assert.AreEqual(6, table.WidthValue(out WidthUnits? units));
		Assert.AreEqual(WidthUnits.Points, units);

		table.ResetFormatting();

		table.Width(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", table.GetTableProperties()!.TableWidth!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, table.GetTableProperties()!.TableWidth!.Type!.Value);

		Assert.AreEqual(66.6, table.WidthValue(out units));
		Assert.AreEqual(WidthUnits.Percentage, units);

		table.ResetFormatting();

		table.Width(new TableWidth { Width = "240" });
		Assert.AreEqual("240", table.GetTableProperties()!.TableWidth!.Width!.Value);
	}

	[TestMethod]
	public void FillColor()
	{
		Table table = new Table().FillColor("Red");
		Assert.AreEqual("Red", table.GetShading()!.Fill!.Value);
	}
}
