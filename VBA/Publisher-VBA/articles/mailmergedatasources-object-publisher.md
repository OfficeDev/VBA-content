---
title: MailMergeDataSources Object (Publisher)
keywords: vbapb10.chm7274495
f1_keywords:
- vbapb10.chm7274495
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSources
ms.assetid: 9eff8354-fbc3-7f55-ba6e-738a60f41259
ms.date: 06/08/2017
---


# MailMergeDataSources Object (Publisher)

Represents the collection of all  **MailMergeDataSource** objects in the active Microsoft Publisher document, each of which represents one of the data sources in a mail merge operation.
 


## Remarks

The default member of the  **MailMergeDataSources** collection is the **Item** method, which returns the **MailMergeDataSource** object at the index position you specify.
 

 
If there is only a single  **MailMergeDataSource** object in the active document, the **MailMergeDataSources** collection is empty. In that case, if you attempt to get the value of the **DataSources** property of the **MailMergeDataSource** object, Publisher returns an error.
 

 

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to get the names of all the connected data sources in the  **MailMergeDataSources** collection in the active document. It uses the **IsDataSourceConnected** property of the active document to determine if a data source is connected.
 

 
If one or more data sources is connected, the macro uses the  **Count** property of the **MailMergeDataSources** collection to determine how many data sources are connected.
 

 
If just one data source is connected, the macro prints the name of that data source in the  **Immediate** window; if more than one data source is connected, it uses the **Item** method of the **MailMergeDataSources** collection to iterate through the collection and the **Name** property of the **MailMergeDataSource** object to print the name of each connected data source in the **Immediate** window.
 

 



```
Public Sub MailMergeDataSources_Example() 
 
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
 Debug.Print "Only one data source ("; pubMailMergeDataSource.Name; ") is connected!" 
 
 End If 
 
 Else 
 
 Debug.Print "No data sources are connected!" 
 
 End If 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Item](mailmergedatasources-item-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](mailmergedatasources-application-property-publisher.md)|
|[Count](mailmergedatasources-count-property-publisher.md)|
|[Creator](mailmergedatasources-creator-property-publisher.md)|
|[Parent](mailmergedatasources-parent-property-publisher.md)|

