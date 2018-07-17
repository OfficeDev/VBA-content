---
title: XPath.SetValue Method (Excel)
keywords: vbaxl10.chm760076
f1_keywords:
- vbaxl10.chm760076
ms.prod: excel
api_name:
- Excel.XPath.SetValue
ms.assetid: 9d7e9eea-0962-cff8-6909-b31d349eb78a
ms.date: 06/08/2017
---


# XPath.SetValue Method (Excel)

Maps the specified  **[XPath](xpath-object-excel.md)** object to a **[ListColumn](listcolumn-object-excel.md)** object or **[Range](range-object-excel.md)** collection. If the **XPath** object has previously been mapped to the **ListColumn** object or **Range** collection, the **SetValue** method sets the properties of the **XPath** object.


## Syntax

 _expression_ . **SetValue**( **_Map_** , **_XPath_** , **_SelectionNamespace_** , **_Repeating_** )

 _expression_ A variable that represents a **XPath** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Map_|Required| **[XmlMap](xmlmap-object-excel.md)**|The map info that the mapped range will be associated with.|
| _XPath_|Required| **String**|A valid XPath expression that tells Excel what XML data should appear in this mapped range. The XPath string can also contain valid filters, in which case, only a subset of the data that the XPath points to will ever appear in this mapped range.|
| _SelectionNamespace_|Optional| **Variant**|Specifies any namespace prefixes used in the XPath argument. This argument can be omitted if the  **XPath** object doesn't contain any prefixes, or if the **XPath** object uses the Microsoft Excel internal prefixes.|
| _Repeating_|Optional| **Variant**|Specifies whether the  **XPath** object is to be bound to a column in an XML list, or mapped to a single cell. Set to **True** to bind the **XPath** object to a column in an XML list. **False** forces a non-repeating cell to be created. If the range is greater than a single cell and **False** is specified, a runtime error will occur.|

## Remarks

See  **[IsExportable](xmlmap-isexportable-property-excel.md)** Property for a discussion on XPath support in Excel. If the XPath expression is invalid or if the XPath specified has already been mapped, a runtime error will occur.

If Excel cannot resolve the namespace, a runtime error will occur.

This method will produce an error if any of the following conditions are true:


- The range spans multiple columns in the grid.
    
- Part of the range spans already mapped cells and the rest spans unmapped cells.
    
- Part of the range spans one mapping, and another part of the range spans a different mapping or different XPath from the same mapping.
    


If the range is a single cell then Excel defaults to creating a single-mapped, non-repeating mapped cell. The non-repeating cell is given no header.

The exception to the above statement occurs when the single-cell range lies within a ListObject, in which case, mapping information is applied to the entire column.

If the range spans multiple cells then Excel creates a repeating XML List. Excel treats the selected range as all data values, so when the XML List is created, the range is shifted down by one row and the header is placed in the cell that the top of the range occupied. The insert row lies at the bottom of the shifted range.

|**Note**|
|:-----|  
|<ul><li>Excel's header detection algorithm is not used in the object model. The assumption is that no headers exist in the grid.</li><li>Auto-merge and auto-grow are disabled when creating mapped ranges in the object model.</li></ul>|

## Example

The following example creates an XML list based on the "Contacts" schema map that is attached to the workbook, and then uses the  **SetValue** method to bind each column to an **XPath** object.


```vb
Sub CreateXMLList() 
    Dim mapContact As XmlMap 
    Dim strXPath As String 
    Dim lstContacts As ListObject 
    Dim objNewCol As ListColumn 
 
    ' Specify the schema map to use. 
    Set mapContact = ActiveWorkbook.XmlMaps("Contacts") 
     
    ' Create a new list. 
    Set lstContacts = ActiveSheet.ListObjects.Add 
         
    ' Specify the first element to map. 
    strXPath = "/Root/Person/FirstName" 
    ' Map the element. 
    lstContacts.ListColumns(1).XPath.SetValue mapContact, strXPath 
 
    ' Specify the second element to map. 
    strXPath = "/Root/Person/LastName" 
    ' Add a column to the list. 
    Set objNewCol = lstContacts.ListColumns.Add 
    ' Map the element. 
    objNewCol.XPath.SetValue mapContact, strXPath 
 
    strXPath = "/Root/Person/Address/Zip" 
    Set objNewCol = lstContacts.ListColumns.Add 
    objNewCol.XPath.SetValue mapContact, strXPath 
End Sub
```


## See also


#### Concepts


[XPath Object](xpath-object-excel.md)

