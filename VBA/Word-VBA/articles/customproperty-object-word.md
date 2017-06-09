---
title: CustomProperty Object (Word)
keywords: vbawd10.chm3552
f1_keywords:
- vbawd10.chm3552
ms.prod: word
api_name:
- Word.CustomProperty
ms.assetid: 1c4aa1ba-ad56-54d1-6e0d-2a82f7b9f4a9
ms.date: 06/08/2017
---


# CustomProperty Object (Word)

Represents a single instance of a custom property for a smart tag. The  **CustomProperty** object is a member of the **[CustomProperties](customproperties-object-word.md)** collection.


## Remarks

Use the  **[Item](customproperties-item-method-word.md)** method—or **[Properties](http://msdn.microsoft.com/library/c9f81907-e257-85cd-bc65-5b614e905738%28Office.15%29.aspx)** (Index), where Index is the number of the property—of the **CustomProperties** collection to return a **CustomProperty** object.

Use the  **[Name](customproperty-name-property-word.md)** and **[Value](customproperty-value-property-word.md)** properties to return the information related to a custom property for a smart tag. This example displays a message containing the name and value of the first custom property of the first smart tag in the current document. This example assumes that the current document contains at least one smart tag and that the first smart tag has at least one custom property.




```vb
Sub SmartTagsProps() 
 With ActiveDocument.SmartTags(Index:=1).Properties.Item(Index:=1) 
 MsgBox "Smart Tag Name: " &; .Name &; vbLf &; _ 
 "Smart Tag Value: " &; .Value 
 End With 
End Sub
```


## See also


#### Other resources


[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)


