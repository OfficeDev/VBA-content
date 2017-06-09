---
title: XPath Object (Excel)
keywords: vbaxl10.chm759072
f1_keywords:
- vbaxl10.chm759072
ms.prod: excel
api_name:
- Excel.XPath
ms.assetid: e13f2b3e-cef2-4e3c-f942-5347cf722e2d
ms.date: 06/08/2017
---


# XPath Object (Excel)

Represents an XPath that has been mapped to a  **[Range](range-object-excel.md)** or **[ListColumn](listcolumn-object-excel.md)** object.


## Example

Use the  **[SetValue](xpath-setvalue-method-excel.md)** method to map an XPath to a range or list column. The **SetValue** method is also used to change the properties of an existing XPath.

The following example creates an XML list based on the "Contacts" schema map that is attached to the workbook, then uses the  **SetValue** method to bind each column to an XPath.

Use the  **[Clear](xpath-clear-method-excel.md)** method to remove an XPath that has been mapped to a range or list column.




```vb
Sub CreateXMLList() 
 Dim mapContact As XmlMap 
 Dim strXPath As String 
 Dim lstContacts As ListObject 
 Dim lcNewCol As ListColumn 
 
 ' Specify the schema map to use. 
 Set mapContact = ActiveWorkbook.XmlMaps("Contacts") 
 
 ' Create a new list. 
 Set lstContacts = ActiveSheet.ListObjects.Add 
 
 ' Specify the first element to map. 
 strXPath = "/Root/Person/FirstName" 
 ' Map the element. 
 lstContacts.ListColumns(1).XPath.SetValue mapContact, strXPath 
 
 ' Specify the element to map. 
 strXPath = "/Root/Person/LastName" 
 ' Add a column to the list. 
 Set lcNewCol = lstContacts.ListColumns.Add 
 ' Map the element. 
 lcNewCol.XPath.SetValue mapContact, strXPath 
 
 strXPath = "/Root/Person/Address/Zip" 
 Set lcNewCol = lstContacts.ListColumns.Add 
 lcNewCol.XPath.SetValue mapContact, strXPath 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

