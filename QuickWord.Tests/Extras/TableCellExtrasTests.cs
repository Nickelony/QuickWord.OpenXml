using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.Extras;

namespace QuickWord.Tests.Extras;

[TestClass]
public class TableCellExtrasTests
{
	[TestMethod]
	public void Formatting()
	{
		TableCell cell = new TableCell().FitText().NoContentWrapping().Shading(new Shading { Fill = "Red" });
		TableCellFormatting formatting = cell.CloneFormatting();

		Assert.AreEqual(3, cell.TableCellProperties!.ChildElements.Count);

		cell.ResetFormatting();
		Assert.AreEqual(0, cell.ChildElements.Count);

		cell.ApplyFormatting(formatting);
		Assert.IsTrue(cell.FitTextValue());
		Assert.IsTrue(cell.NoContentWrappingValue());
		Assert.AreEqual("Red", cell.GetShading()!.Fill!.Value);
		Assert.AreEqual(3, cell.TableCellProperties!.ChildElements.Count);

		var anotherFormatting = new TableCellFormatting { HideEndOfCellMarker = true, Shading = new Shading { Fill = "Blue" } };

		cell.ApplyFormatting(anotherFormatting, true);
		Assert.IsTrue(cell.FitTextValue());
		Assert.IsTrue(cell.NoContentWrappingValue());
		Assert.IsTrue(cell.HideEndOfCellMarkerValue());
		Assert.AreEqual("Blue", cell.GetShading()!.Fill!.Value);
		Assert.AreEqual(4, cell.TableCellProperties!.ChildElements.Count);

		cell.ApplyFormatting(anotherFormatting);
		Assert.IsNull(cell.FitTextValue());
		Assert.IsNull(cell.NoContentWrappingValue());
		Assert.IsTrue(cell.HideEndOfCellMarkerValue());
		Assert.AreEqual("Blue", cell.GetShading()!.Fill!.Value);
		Assert.AreEqual(2, cell.TableCellProperties!.ChildElements.Count);

		cell.ResetFormatting();
	}

	[TestMethod]
	public void Borders()
	{
		TableCell cell = new TableCell()
			.LeftBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.TopBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.RightBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.BottomBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.InsideHorizontalBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.InsideVerticalBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.StartBorder(1.5, BorderValues.DashDotStroked, "Red", 1)
			.EndBorder(1.5, BorderValues.DashDotStroked, "Red", 1);

		Assert.AreEqual(9U, cell.GetBorders()!.LeftBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.TopBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.RightBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.BottomBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.InsideHorizontalBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.InsideVerticalBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.StartBorder!.Size!.Value);
		Assert.AreEqual(9U, cell.GetBorders()!.EndBorder!.Size!.Value);

		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.LeftBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.TopBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.RightBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.BottomBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.InsideHorizontalBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.InsideVerticalBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.StartBorder!.Val!.Value);
		Assert.AreEqual(BorderValues.DashDotStroked, cell.GetBorders()!.EndBorder!.Val!.Value);

		Assert.AreEqual("Red", cell.GetBorders()!.LeftBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.TopBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.RightBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.BottomBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.InsideHorizontalBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.InsideVerticalBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.StartBorder!.Color!.Value);
		Assert.AreEqual("Red", cell.GetBorders()!.EndBorder!.Color!.Value);

		Assert.AreEqual(1U, cell.GetBorders()!.LeftBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.TopBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.RightBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.BottomBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.InsideHorizontalBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.InsideVerticalBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.StartBorder!.Space!.Value);
		Assert.AreEqual(1U, cell.GetBorders()!.EndBorder!.Space!.Value);

		cell.ResetFormatting();

		cell.LeftBorder(new LeftBorder { Size = 6 })
			.TopBorder(new TopBorder { Size = 6 })
			.RightBorder(new RightBorder { Size = 6 })
			.BottomBorder(new BottomBorder { Size = 6 })
			.InsideHorizontalBorder(new InsideHorizontalBorder { Size = 6 })
			.InsideVerticalBorder(new InsideVerticalBorder { Size = 6 })
			.StartBorder(new StartBorder { Size = 6 })
			.EndBorder(new EndBorder { Size = 6 });

		Assert.AreEqual(6U, cell.GetBorders()!.LeftBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.TopBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.RightBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.BottomBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.InsideHorizontalBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.InsideVerticalBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.StartBorder!.Size!.Value);
		Assert.AreEqual(6U, cell.GetBorders()!.EndBorder!.Size!.Value);
	}

	[TestMethod]
	public void Margins()
	{
		TableCell cell = new TableCell()
			.TopMargin(6, WidthUnits.Points)
			.BottomMargin(6, WidthUnits.Points)
			.LeftMargin(6, WidthUnits.Points)
			.RightMargin(6, WidthUnits.Points)
			.StartMargin(6, WidthUnits.Points)
			.EndMargin(6, WidthUnits.Points);

		Assert.AreEqual("120", cell.GetMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("120", cell.GetMargins()!.BottomMargin!.Width!.Value);
		Assert.AreEqual("120", cell.GetMargins()!.LeftMargin!.Width!.Value);
		Assert.AreEqual("120", cell.GetMargins()!.RightMargin!.Width!.Value);
		Assert.AreEqual("120", cell.GetMargins()!.StartMargin!.Width!.Value);
		Assert.AreEqual("120", cell.GetMargins()!.EndMargin!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetMargins()!.TopMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetMargins()!.BottomMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetMargins()!.LeftMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetMargins()!.RightMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetMargins()!.StartMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetMargins()!.EndMargin!.Type!.Value);

		Assert.AreEqual(6, cell.TopMarginValue(out WidthUnits? topUnits));
		Assert.AreEqual(6, cell.BottomMargin(out WidthUnits? bottomUnits));
		Assert.AreEqual(6, cell.LeftMarginValue(out WidthUnits? leftUnits));
		Assert.AreEqual(6, cell.RightMarginValue(out WidthUnits? rightUnits));
		Assert.AreEqual(6, cell.StartMarginValue(out WidthUnits? startUnits));
		Assert.AreEqual(6, cell.EndMarginValue(out WidthUnits? endUnits));
		Assert.AreEqual(WidthUnits.Points, topUnits);
		Assert.AreEqual(WidthUnits.Points, bottomUnits);
		Assert.AreEqual(WidthUnits.Points, leftUnits);
		Assert.AreEqual(WidthUnits.Points, rightUnits);
		Assert.AreEqual(WidthUnits.Points, startUnits);
		Assert.AreEqual(WidthUnits.Points, endUnits);

		cell.ResetFormatting();

		cell.TopMargin(66.6, WidthUnits.Percentage)
			.BottomMargin(66.6, WidthUnits.Percentage)
			.LeftMargin(66.6, WidthUnits.Percentage)
			.RightMargin(66.6, WidthUnits.Percentage)
			.StartMargin(66.6, WidthUnits.Percentage)
			.EndMargin(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", cell.GetMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("3330", cell.GetMargins()!.BottomMargin!.Width!.Value);
		Assert.AreEqual("3330", cell.GetMargins()!.LeftMargin!.Width!.Value);
		Assert.AreEqual("3330", cell.GetMargins()!.RightMargin!.Width!.Value);
		Assert.AreEqual("3330", cell.GetMargins()!.StartMargin!.Width!.Value);
		Assert.AreEqual("3330", cell.GetMargins()!.EndMargin!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetMargins()!.TopMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetMargins()!.BottomMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetMargins()!.LeftMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetMargins()!.RightMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetMargins()!.StartMargin!.Type!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetMargins()!.EndMargin!.Type!.Value);

		Assert.AreEqual(66.6, cell.TopMarginValue(out topUnits));
		Assert.AreEqual(66.6, cell.BottomMargin(out bottomUnits));
		Assert.AreEqual(66.6, cell.LeftMarginValue(out leftUnits));
		Assert.AreEqual(66.6, cell.RightMarginValue(out rightUnits));
		Assert.AreEqual(66.6, cell.StartMarginValue(out startUnits));
		Assert.AreEqual(66.6, cell.EndMarginValue(out endUnits));
		Assert.AreEqual(WidthUnits.Percentage, topUnits);
		Assert.AreEqual(WidthUnits.Percentage, bottomUnits);
		Assert.AreEqual(WidthUnits.Percentage, leftUnits);
		Assert.AreEqual(WidthUnits.Percentage, rightUnits);
		Assert.AreEqual(WidthUnits.Percentage, startUnits);
		Assert.AreEqual(WidthUnits.Percentage, endUnits);

		cell.ResetFormatting();

		cell.TopMargin(new TopMargin { Width = "240" })
			.BottomMargin(new BottomMargin { Width = "240" })
			.LeftMargin(new LeftMargin { Width = "240" })
			.RightMargin(new RightMargin { Width = "240" })
			.StartMargin(new StartMargin { Width = "240" })
			.EndMargin(new EndMargin { Width = "240" });

		Assert.AreEqual("240", cell.GetMargins()!.TopMargin!.Width!.Value);
		Assert.AreEqual("240", cell.GetMargins()!.BottomMargin!.Width!.Value);
		Assert.AreEqual("240", cell.GetMargins()!.LeftMargin!.Width!.Value);
		Assert.AreEqual("240", cell.GetMargins()!.RightMargin!.Width!.Value);
		Assert.AreEqual("240", cell.GetMargins()!.StartMargin!.Width!.Value);
		Assert.AreEqual("240", cell.GetMargins()!.EndMargin!.Width!.Value);
	}

	[TestMethod]
	public void Width()
	{
		TableCell cell = new TableCell()
			.Width(6, WidthUnits.Points);

		Assert.AreEqual("120", cell.GetWidth()!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, cell.GetWidth()!.Type!.Value);
		Assert.AreEqual(6, cell.WidthValue(out WidthUnits? units));
		Assert.AreEqual(WidthUnits.Points, units);

		cell.ResetFormatting();
		cell.Width(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", cell.GetWidth()!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, cell.GetWidth()!.Type!.Value);
		Assert.AreEqual(66.6, cell.WidthValue(out units));
		Assert.AreEqual(WidthUnits.Percentage, units);

		cell.ResetFormatting();
		cell.Width(new TableCellWidth { Width = "240" });

		Assert.AreEqual("240", cell.GetWidth()!.Width!.Value);
	}

	[TestMethod]
	public void FillColor()
	{
		TableCell cell = new TableCell().FillColor("Red");
		Assert.AreEqual("Red", cell.GetShading()!.Fill!.Value);
	}
}
