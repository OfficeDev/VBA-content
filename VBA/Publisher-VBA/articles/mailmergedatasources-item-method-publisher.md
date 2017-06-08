---
title: MailMergeDataSources.Item Method (Publisher)
keywords: vbapb10.chm7143427
f1_keywords:
- vbapb10.chm7143427
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSources.Item
ms.assetid: a65fedf6-aae5-64ef-e7d0-6bbc3d5b733c
ms.date: 06/08/2017
---


# MailMergeDataSources.Item Method (Publisher)

Returns the  **[MailMergeDataSource](mailmergedatasource-object-publisher.md)** object at the specified index position in the **MailMergeDataSources** collection.


## Syntax

 _expression_. **Item**( **_varIndex_**)

 _expression_A variable that represents a  **MailMergeDataSources** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|varIndex|Required| **Variant**|The index number or name of the object to return.|

### Return Value

MailMergeDataSource


## Remarks

The  **Item** method is the default member of the **MailMergeDataSources** collection

If there is only a single  **MailMergeDataSource** object in the active document, the **MailMergeDataSources** collection is empty. In that case, if you try to use the **DataSources** property of the **MailMergeDataSource** object to get the data sources collection, Microsoft Publisher returns an error.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to get the names of all the connected data sources in the  **MailMergeDataSources** collection in the active document. It uses the **IsDataSourceConnected** property of the active document to determine if a data source is connected.

If one or more data sources is connected, the macro uses the  **Count** property of the **MailMergeDataSources** collection to determine how many data sources are connected.

If just one data source is connected, the macro prints the name of that data source in the  **Immediate** window; if more than one data source is connected, the macro uses the **Item** method of the **MailMergeDataSources** collection to iterate through the collection and the **Name** property of the **MailMergeDataSource** object to print the name of each connected data source in the **Immediate** window.




```vb
Public Sub Item_Example() 
 
 Dim pubMailMergeDataSources As Publisher.MailMergeDataSources 
 Dim pubMailMergeDataSource As Publisher.MailMergeDataSource 
 Dim lngCount As Long 
 Dim intCounter As Integer 
 
 If ThisDocument.IsDataSourceConnected Then 
 
 Set pubMailMergeDataSources = ThisDocument.MailMerge.DataSource.DataSources 
 
 lngCount = pubMailMergeDataSources.Count 
 
 If lngCount > 1 Then 
 
 ' More than one data source is connected. 
 For intCounter = 1 To lngCount 
 Debug.Print pubMailMergeDataSources.Item(intCounter).Name 
 Next 
 
 Else 
 
 ' Only one data source is connected. 
 Set pubMailMergeDataSource = ThisDocument.MailMerge.DataSource 
 Debug.Print "Only one data source ("; pubMailMergeDataSource.Name; ") is connected." 
 
 End If 
 
 Else 
 
 Debug.Print "No data sources are connected." 
 
 End If 
 
End Sub
```


