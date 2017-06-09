---
title: View.RevisionsBalloonShowConnectingLines Property (Word)
keywords: vbawd10.chm161808428
f1_keywords:
- vbawd10.chm161808428
ms.prod: word
api_name:
- Word.View.RevisionsBalloonShowConnectingLines
ms.assetid: 78c1cf42-93a7-eec9-84f6-40c6e7de036c
ms.date: 06/08/2017
---


# View.RevisionsBalloonShowConnectingLines Property (Word)

 **True** for Microsoft Word to display connecting lines from the text to the revision and comment balloons. Read/write **Boolean** .


## Syntax

 _expression_ . **RevisionsBalloonShowConnectingLines**

 _expression_ A variable that represents a **[View](view-object-word.md)** object.


## Example

This example hides the lines connecting the document text with the corresponding revision or comment balloons.


```vb
Sub ShowConnectingLines() 
 ActiveWindow.View _ 
 .RevisionsBalloonShowConnectingLines = False 
End Sub
```


## See also


#### Concepts


[View Object](view-object-word.md)

