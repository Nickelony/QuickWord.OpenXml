// Ignore Spelling: Strikethrough

using DocumentFormat.OpenXml.Wordprocessing;
using QuickWord.OpenXml.Measurements;

namespace QuickWord.OpenXml.Extras;

public class RunFormatting
{
	public bool? Bold { get; set; }
	public bool? BoldComplexScript { get; set; }
	public Border? Border { get; set; }
	public bool? AllCaps { get; set; }
	public Color? Color { get; set; }
	public bool? ComplexScript { get; set; }
	public bool? DoubleStrike { get; set; }
	public EastAsianLayout? EastAsianLayout { get; set; }
	public TextEffectValues? TextEffect { get; set; }
	public EmphasisMarkValues? EmphasisMark { get; set; }
	public bool? Emboss { get; set; }
	public FitText? FitText { get; set; }
	public HighlightColorValues? HighlightColor { get; set; }
	public bool? Italic { get; set; }
	public bool? ItalicComplexScript { get; set; }
	public bool? Imprint { get; set; }
	public double? Kerning { get; set; }
	public Languages? Languages { get; set; }
	public bool? NoProofing { get; set; }
	public bool? OfficeMath { get; set; }
	public bool? Outline { get; set; }
	public double? VerticalPosition { get; set; }
	public RunFonts? Fonts { get; set; }
	public string? Style { get; set; }
	public bool? RightToLeft { get; set; }
	public bool? Shadow { get; set; }
	public Shading? Shading { get; set; }
	public bool? SmallCaps { get; set; }
	public bool? SnapToGrid { get; set; }
	public double? CharacterSpacing { get; set; }
	public bool? SpecVanish { get; set; }
	public bool? Strike { get; set; }
	public double? FontSize { get; set; }
	public double? ComplexScriptFontSize { get; set; }
	public Underline? Underline { get; set; }
	public bool? Hidden { get; set; }
	public VerticalPositionValues? VerticalAlignment { get; set; }
	public long? CharacterScale { get; set; }
	public bool? WebHidden { get; set; }

	// Extras:

	public string? FillColor
	{
		get => Shading?.Fill;
		set
		{
			Shading ??= new Shading();
			Shading.Fill = value;
		}
	}

	public string? FontColor
	{
		get => Color?.Val;
		set
		{
			Color ??= new Color();
			Color.Val = value;
		}
	}

	public string? FontFace
	{
		get => Fonts?.Ascii;
		set
		{
			Fonts ??= new RunFonts();
			Fonts.Ascii = value;
		}
	}

	public string? Language
	{
		get => Languages?.Val;
		set
		{
			Languages ??= new Languages();
			Languages.Val = value;
		}
	}

	public QBorder QBorder
	{
		set => Border = new Border
		{
			Val = value.Border,
			Size = BorderSize.ToSixth(value.Width),
			Color = value.Color,
			Space = value.Spacing
		};
	}

	public QManualWidth QManualWidth
	{
		set
		{
			FitText ??= new FitText();
			FitText.Val = (uint)Twips.FromOther(value.Width, value.Units);
		}
	}

	public QUnderline QUnderline
	{
		set => Underline = new Underline
		{
			Val = value.Style,
			Color = value.Color
		};
	}
}
