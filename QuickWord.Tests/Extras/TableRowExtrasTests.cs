using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml;
using QuickWord.OpenXml.Extras;

namespace QuickWord.Tests.Extras;

[TestClass]
public class TableRowExtrasTests
{
	[TestMethod]
	public void Formatting()
	{
		TableRow row = new TableRow().CantSplit().IsHeader().WidthAfter(new WidthAfterTableRow { Width = "66" });
		TableRowFormatting formatting = row.CloneFormatting();

		Assert.AreEqual(3, row.TableRowProperties!.ChildElements.Count);

		row.ResetFormatting();
		Assert.AreEqual(0, row.ChildElements.Count);

		row.ApplyFormatting(formatting);
		Assert.IsTrue(row.CantSplitValue());
		Assert.IsTrue(row.IsHeaderValue());
		Assert.AreEqual("66", row.GetWidthAfter()!.Width!.Value);
		Assert.AreEqual(3, row.TableRowProperties!.ChildElements.Count);

		var anotherFormatting = new TableRowFormatting { HideEndOfRowMarker = true, WidthAfter = new WidthAfterTableRow { Width = "120" } };

		row.ApplyFormatting(anotherFormatting, true);
		Assert.IsTrue(row.CantSplitValue());
		Assert.IsTrue(row.IsHeaderValue());
		Assert.IsTrue(row.HideEndOfRowMarkerValue());
		Assert.AreEqual("120", row.GetWidthAfter()!.Width!.Value);
		Assert.AreEqual(4, row.TableRowProperties!.ChildElements.Count);

		row.ApplyFormatting(anotherFormatting);
		Assert.IsNull(row.CantSplitValue());
		Assert.IsNull(row.IsHeaderValue());
		Assert.IsTrue(row.HideEndOfRowMarkerValue());
		Assert.AreEqual("120", row.GetWidthAfter()!.Width!.Value);
		Assert.AreEqual(2, row.TableRowProperties!.ChildElements.Count);

		row.ResetFormatting();
	}

	[TestMethod]
	public void WidthBefore()
	{
		TableRow row = new TableRow()
			.WidthBefore(6, WidthUnits.Points);

		Assert.AreEqual("120", row.GetWidthBefore()!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, row.GetWidthBefore()!.Type!.Value);
		Assert.AreEqual(6, row.WidthBeforeValue(out WidthUnits? units));
		Assert.AreEqual(WidthUnits.Points, units);

		row.ResetFormatting();
		row.WidthBefore(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", row.GetWidthBefore()!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, row.GetWidthBefore()!.Type!.Value);
		Assert.AreEqual(66.6, row.WidthBeforeValue(out units));
		Assert.AreEqual(WidthUnits.Percentage, units);

		row.ResetFormatting();
		row.WidthBefore(new WidthBeforeTableRow { Width = "240" });

		Assert.AreEqual("240", row.GetWidthBefore()!.Width!.Value);
	}

	[TestMethod]
	public void WidthAfter()
	{
		TableRow row = new TableRow()
			.WidthAfter(6, WidthUnits.Points);

		Assert.AreEqual("120", row.GetWidthAfter()!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Dxa, row.GetWidthAfter()!.Type!.Value);
		Assert.AreEqual(6, row.WidthAfterValue(out WidthUnits? units));
		Assert.AreEqual(WidthUnits.Points, units);

		row.ResetFormatting();
		row.WidthAfter(66.6, WidthUnits.Percentage);

		Assert.AreEqual("3330", row.GetWidthAfter()!.Width!.Value);
		Assert.AreEqual(TableWidthUnitValues.Pct, row.GetWidthAfter()!.Type!.Value);
		Assert.AreEqual(66.6, row.WidthAfterValue(out units));
		Assert.AreEqual(WidthUnits.Percentage, units);

		row.ResetFormatting();
		row.WidthAfter(new WidthAfterTableRow { Width = "240" });

		Assert.AreEqual("240", row.GetWidthAfter()!.Width!.Value);
	}

	[TestMethod]
	public void Height()
	{
		TableRow row = new TableRow()
			.Height(100, MeasuringUnits.Points, HeightRuleValues.AtLeast);

		Assert.AreEqual(2000U, row.GetHeight()!.Val!.Value); // Twips
		Assert.AreEqual(HeightRuleValues.AtLeast, row.GetHeight()!.HeightType!.Value);

		Assert.AreEqual(100, row.HeightValue(MeasuringUnits.Points, out HeightRuleValues? rule));
		Assert.AreEqual(HeightRuleValues.AtLeast, rule);

		row.Height(1, MeasuringUnits.Inches, HeightRuleValues.Auto);
		Assert.AreEqual(1, row.HeightValue(MeasuringUnits.Inches, out _));
		Assert.AreEqual(2.54, row.HeightValue(MeasuringUnits.Centimeters, out _));

		row.Height(5.08, MeasuringUnits.Centimeters, HeightRuleValues.Exact);
		Assert.AreEqual(5.08, row.HeightValue(MeasuringUnits.Centimeters, out _));
		Assert.AreEqual(2, row.HeightValue(MeasuringUnits.Inches, out _));
	}
}
