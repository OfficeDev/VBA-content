---
title: Explorer.ActiveInlineResponseWordEditor Property (Outlook)
keywords: vbaol11.chm3597
f1_keywords:
- vbaol11.chm3597
ms.assetid: b9058694-ab8f-4962-ab7d-afac1704dd29
ms.date: 06/08/2017
ms.prod: outlook
---


# Explorer.ActiveInlineResponseWordEditor Property (Outlook)
Returns the Word [Document](http://msdn.microsoft.com/library/8d83487a-2345-a036-a916-971c9db5b7fb%28Office.15%29.aspx) object of the active inline response that is displayed in the explorer Reading Pane. Read-only.

## Syntax

 _expression_ . **ActiveInlineResponseWordEditor**

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


## Remarks

This property returns  **Null** ( **Nothing** in Visual Basic) if no inline response is visible in the Reading Pane. The returned Word **Document** object provides access to most of the Word object model except for the following members:


- [InlineShapes.AddChart2](http://msdn.microsoft.com/library/108899b6-24bb-cf4c-db95-066219536c19%28Office.15%29.aspx)
    
- [Range.ConvertToTable](http://msdn.microsoft.com/library/a7d005ec-774e-151c-ff38-64df3ea36646%28Office.15%29.aspx)
    
- [Range.ImportFragment](http://msdn.microsoft.com/library/d9feca50-6370-c1c2-00c0-e64ff7a5adb9%28Office.15%29.aspx)
    
- [Range.InsertXML](http://msdn.microsoft.com/library/daee0fee-01cb-5ad7-f61d-ea6ebec1d04a%28Office.15%29.aspx)
    
- [Shapes.AddChart2](http://msdn.microsoft.com/library/54b1e65b-57ad-4824-2acf-2e1e0a22f085%28Office.15%29.aspx)
    
- [Selection.InsertXML](http://msdn.microsoft.com/library/7a9e52b5-9b05-f939-6fd0-33a923989f48%28Office.15%29.aspx)
    
- [Tables.Add](http://msdn.microsoft.com/library/127b5f74-876f-1307-5d25-a04c99debd6b%28Office.15%29.aspx)
    

## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

