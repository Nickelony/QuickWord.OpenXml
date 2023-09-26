using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Reflection;

namespace QuickWord.OpenXml.Utilities;

internal static class OpenXmlElementExtensions
{
	/// <summary>
	/// Gets the child of the specified type, or initializes / appends a new one if it doesn't exist.
	/// </summary>
	/// <typeparam name="TProperty">OpenXml property class, such as <see cref="Shading" />, <see cref="RunFonts" /> or <see cref="Bold" />.</typeparam>
	/// <param name="propertyHost">The property host, such as <see cref="RunProperties" />, <see cref="TableBorders" /> or even just a <see cref="Paragraph" />.</param>
	public static TProperty GetOrInit<TProperty>(this OpenXmlElement propertyHost, bool topElement = false) where TProperty : OpenXmlElement, new()
	{
		TProperty? targetChild = propertyHost.GetFirstChild<TProperty>();

		if (targetChild is null)
		{
			var newChild = new TProperty();
			PropertyInfo? targetPropertyField = propertyHost.GetType().GetProperty(typeof(TProperty).Name);

			if (targetPropertyField is null)
			{
				if (topElement)
					propertyHost.InsertAt(newChild, 0);
				else
					propertyHost.AppendChild(newChild);
			}
			else
				targetPropertyField.SetValue(propertyHost, newChild);

			return newChild;
		}
		else
			return targetChild;
	}

	/// <summary>
	/// Sets the value of the specified field of an OpenXml property class, such as <see cref="Bold" /> or <see cref="FontSize" />.
	/// <para>Removes the property if <see langword="null" /> is passed.</para>
	/// </summary>
	/// <typeparam name="TProperty">OpenXml property class, such as <see cref="Bold" /> or <see cref="FontSize" />.</typeparam>
	/// <param name="fieldName">The name of the field, such as "Val".</param>
	/// <param name="propertyHost">The property host, such as <see cref="RunProperties" />, <see cref="ParagraphProperties" /> or <see cref="TableBorders" />.</param>
	public static void SetFieldOrRemove<TProperty>(this OpenXmlElement propertyHost, string fieldName, object? value)
		where TProperty : OpenXmlElement, new()
	{
		if (value is null) // Setting to null should remove the property completely
			propertyHost.RemoveAllChildren<TProperty>();
		else // Get reference to property or create it if it doesn't exist, then set the value of the specified field
		{
			TProperty openXmlProperty = propertyHost.GetOrInit<TProperty>();
			openXmlProperty.SetExactField(fieldName, value);
		}

		if (propertyHost.ChildElements.Count == 0)
			propertyHost.Remove();
	}

	/// <summary>
	/// Sets the value of the "Val" field of an OpenXml property class, such as <see cref="Bold" /> or <see cref="FontSize" />.
	/// <para>Removes the property if <see langword="null" /> is passed.</para>
	/// </summary>
	/// <typeparam name="TProperty">OpenXml property class, such as <see cref="Bold" /> or <see cref="FontSize" />.</typeparam>
	/// <param name="propertyHost">The property host, such as <see cref="RunProperties" />, <see cref="ParagraphProperties" /> or <see cref="TableBorders" />.</param>
	public static void SetValOrRemove<TProperty>(this OpenXmlElement propertyHost, object? value)
		where TProperty : OpenXmlElement, new()
		=> SetFieldOrRemove<TProperty>(propertyHost, "Val", value);

	/// <summary>
	/// Sets the value of an OpenXml property class child, such as <see cref="Shading" /> or <see cref="RunFonts" />.
	/// <para>Removes the property if <see langword="null" /> is passed.</para>
	/// </summary>
	/// <typeparam name="TProperty">OpenXml property class, such as <see cref="Shading" /> or <see cref="RunFonts" />.</typeparam>
	/// <param name="propertyHost">The property host, such as <see cref="RunProperties" />, <see cref="ParagraphProperties" /> or <see cref="TableBorders" />.</param>
	public static void SetPropertyClassOrRemove<TProperty>(this OpenXmlElement propertyHost, TProperty? property)
		where TProperty : OpenXmlElement
	{
		propertyHost.RemoveAllChildren<TProperty>();

		if (property is not null)
			propertyHost.AppendChild(property);

		if (propertyHost.ChildElements.Count == 0)
			propertyHost.Remove();
	}

	/// <summary>
	/// Sets the value of the specified field of an OpenXml property class, such as <see cref="Shading" />, <see cref="RunFonts" /> or <see cref="Bold" />.
	/// </summary>
	/// <param name="propertyElement">The property element, such as <see cref="Shading" />, <see cref="RunFonts" /> or <see cref="Bold" />.</param>
	/// <param name="fieldName">Name of the target field, such as <c>"Fill"</c> from <see cref="Shading" />, <c>"Ascii"</c> from <see cref="RunFonts" /> or <c>"Val"</c> from <see cref="Bold" />.</param>
	private static void SetExactField(this OpenXmlElement propertyElement, string fieldName, object? value)
	{
		PropertyInfo? exactField = propertyElement.GetType().GetProperty(fieldName);

		if (exactField is null)
			return; // This should generally never happen, unless fieldName is wrong

		// Get the value of the "Value class" field (such as StringValue, UInt32Value, OnOffValue etc.)
		object? exactFieldValue = exactField.GetValue(propertyElement);

		if (exactFieldValue is null) // Field exists, but its value is null,
									 // so we need to create a new "Value class" instance (such as StringValue),
									 // then we assign the new instance to the field with its value set through the constructor
		{
			object? newValue = Activator.CreateInstance(exactField.PropertyType, value);
			exactField.SetValue(propertyElement, newValue);
		}
		else // Field exists and is already set, so we just need to set the .Value field of the "Value class" (e.g. StringValue.Value)
		{
			PropertyInfo? valueProperty = exactFieldValue!.GetType().GetProperty("Value");
			valueProperty?.SetValue(exactFieldValue, value);
		}
	}
}
