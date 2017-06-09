---
title: PreviewPane.WordEditor Property (Outlook)
keywords: vbaol11.chm3640
f1_keywords:
- vbaol11.chm3640
ms.assetid: 8c50e511-99ed-a691-352e-ae8f0942dbe5
ms.date: 06/08/2017
ms.prod: outlook
---


# PreviewPane.WordEditor Property (Outlook)

Returns the Microsoft Word Document Object Model of the message being displayed. Read-only.


## Syntax

 _expression_ . **WordEditor**

 _expression_ A variable that represents a **PreviewPane** object.


## Remarks

The  **WordEditor** property is only valid if[IsWordMail](inspector-iswordmail-method-outlook.md) returns True and the[EditorType](inspector-editortype-property-outlook.md) is **olEditorWord** . The returned **WordDocument** object provides access to most of the Word object model, except for the following members:


- Tables.Add
    
- Range.ConvertToTable
    
- InlineShapes.AddChart
    
- Shapes.AddChart
    
- Range.InsertXML
    
- Selection.InsertXML
    
- Range.ImportFragment
    

## See also


#### Other resources



[PreviewPane Object (Outlook)](previewpane-object-outlook.md)

