---
title: Font.ExpandUsingKashida Property (Publisher)
keywords: vbapb10.chm5374004
f1_keywords:
- vbapb10.chm5374004
ms.prod: publisher
api_name:
- Publisher.Font.ExpandUsingKashida
ms.assetid: ecf3a170-5f07-379e-ff56-504beb770308
ms.date: 06/08/2017
---


# Font.ExpandUsingKashida Property (Publisher)

Returns or sets an  **MsoTriState** constant indicating whether to apply kashida rules while applying tracking to Arabic text. Read/write.


## Syntax

 _expression_. **ExpandUsingKashida**

 _expression_A variable that represents an  **Font** object.


### Return Value

MsoTriState


## Remarks

The  **ExpandUsingKashida** property value can be one of the **MsoTriState** constants declared in the Microsoft Office type library and shown in the following table.



|**Constant**|**Description**|
|:-----|:-----|
| **msoFalse**| Microsoft Publisher does not apply kashida rules while applying tracking to Arabic text.|
| **msoTriStateMixed**|Return value indicating a combination of  **msoTrue** and **msoFalse** for the specified text range.|
| **msoTriStateToggle**|Set value that switches between  **msoTrue** and **msoFalse**.|
| **msoTrue**| Publisher does apply kashida rules while applying tracking to Arabic text.|

## Example

The following example sets Publisher to apply kashida rules while applying tracking to Arabic text for all text ranges on page one of the active publication.


```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 If shpLoop.HasTextFrame Then 
 shpLoop.TextFrame.TextRange _ 
 .Font.ExpandUsingKashida = msoTrue 
 End If 
Next shpLoop
```


