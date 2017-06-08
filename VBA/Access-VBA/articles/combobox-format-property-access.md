---
title: ComboBox.Format Property (Access)
keywords: vbaac10.chm11375
f1_keywords:
- vbaac10.chm11375
ms.prod: access
api_name:
- Access.ComboBox.Format
ms.assetid: 9bb18f6a-0a25-9bbf-88ba-adf603c11826
ms.date: 06/08/2017
---


# ComboBox.Format Property (Access)

You can use the  **Format** property to customize the way numbers, dates, times, and text are displayed and printed. Read/write **String**.


## Syntax

 _expression_. **Format**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

You can use one of the predefined formats or you can create a custom format by using formatting symbols.

The  **Format** property uses different settings for different data types. For information about settings for a specific data type, see one of the following topics:


- [Date/Time Data Type](format-propertydate-time-data-type.md)
    
- [Number and Currency Data Types](format-propertynumber-and-currency-data-types.md)
    
- [Text and Memo Data Types](format-propertytext-and-memo-data-types.md)
    
- [Yes/No Data Type](format-propertyyes-no-data-type.md)
    
In Visual Basic, enter a string expression that corresponds to one of the predefined formats or enter a custom format.

The  **Format** property affects only how data is displayed. It doesn't affect how data is stored.

Microsoft Access provides predefined formats for Date/Time, Number and Currency, Text and Memo, and Yes/No data types. The predefined formats depend on the country/region specified by double-clicking Regional Options in Windows Control Panel. Microsoft Access displays formats appropriate for the country/region selected. For example, with  **English (United States)** selected on the **General** tab, 1234.56 in the Currency format appears as $1,234.56, but when **English (British)** is selected on the **General** tab, the number appears as ?1,234.56.

If you set a field's  **Format** property in table Design view, Microsoft Access uses that format to display data in datasheets. It also applies the field's **Format** property to new controls on forms and reports.

You can use the following symbols in custom formats for any data type.



|**Symbol**|**Meaning**|
|:-----|:-----|
|(space)|Display spaces as literal characters.|
|"ABC"|Display anything inside quotation marks as literal characters.|
|!|Force left alignment instead of right alignment.|
|*|Fill available space with the next character.|
|\|Display the next character as a literal character. You can also display literal characters by placing quotation marks around them.|
|[ _color_ ]|Display the formatted data in the color specified between the brackets. Available colors: Black, Blue, Green, Cyan, Red, Magenta, Yellow, White.|
You can't mix custom formatting symbols for the Number and Currency data types with Date/Time, Yes/No, or Text and Memo formatting symbols.

When you have defined an input mask and set the  **Format** property for the same data, the **Format** property takes precedence when the data is displayed and the input mask is ignored. For example, if you create a Password input mask in table Design view and also set the **Format** property for the same field, either in the table or in a control on a form, the Password input mask is ignored and the data is displayed according to the **Format** property.


## Example

The following three examples set the  **Format** property by using a predefined format:


```vb
Me!Date.Format = "Medium Date" 
 
Me!Time.Format = "Long Time" 
 
Me!Registered.Format = "Yes/No"
```

The next example sets the  **Format** property by using a custom format. This format displays a date as: Jan 2006.




```vb
Forms!Employees!HireDate.Format = "mmm yyyy"
```

The following example demonstrates a Visual Basic function that formats numeric data by using the Currency format and formats text data entirely in capital letters. The function is called from the  **OnLostFocus** event of an unbound control named TaxRefund.




```vb
Function FormatValue() As Integer 
    Dim varEnteredValue As Variant 
 
    varEnteredValue = Forms!Survey!TaxRefund.Value 
    If IsNumeric(varEnteredValue) = True Then 
        Forms!Survey!TaxRefund.Format = "Currency" 
    Else 
        Forms!Survey!TaxRefund.Format = ">" 
    End If 
End Function
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

