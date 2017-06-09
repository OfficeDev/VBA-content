---
title: Fields Object (Publisher)
keywords: vbapb10.chm6094847
f1_keywords:
- vbapb10.chm6094847
ms.prod: publisher
api_name:
- Publisher.Fields
ms.assetid: fd7c95d9-bc34-95ee-180d-b99f3629eb33
ms.date: 06/08/2017
---


# Fields Object (Publisher)

A collection of  **[Field](field-object-publisher.md)** objects that represent all the fields in a text range.
 


## Remarks

The  **[Count](fields-count-property-publisher.md)** property for this collection in a publication returns the number of items in a specified shape or selection.
 

 

## Example

Use the  **[Fields](textrange-fields-property-publisher.md)** property to return the **Fields** collection. Use **Fields** (index), where index is the index number, to return a single **Field** object. The index number represents the position of the field in the selection, range, or publication. The following example displays the field code and the result of the first field in each text box in the active publication.
 

 

```
Sub ShowFieldCodes() 
 Dim pagPage As Page 
 Dim shpShape As Shape 
 
 For Each pagPage In ActiveDocument.Pages 
 For Each shpShape In pagPage.Shapes 
 If shpShape.Type = pbTextFrame Then 
 With shpShape.TextFrame.TextRange 
 If .Fields.Count > 0 Then 
 MsgBox "Code = " &amp; .Fields(1).Code &amp; vbLf _ 
 &amp; "Result = " &amp; .Fields(1).Result &amp; vbLf 
 End If 
 End With 
 End If 
 Next 
 Next 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddHorizontalInVertical](fields-addhorizontalinvertical-method-publisher.md)|
|[AddPhoneticGuide](fields-addphoneticguide-method-publisher.md)|
|[Item](fields-item-method-publisher.md)|
|[Unlink](fields-unlink-method-publisher.md)|

## Properties



|**Name**|
|:-----|
|[Application](fields-application-property-publisher.md)|
|[Count](fields-count-property-publisher.md)|
|[Parent](fields-parent-property-publisher.md)|

