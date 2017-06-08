---
title: Presentation.Close Method (PowerPoint)
keywords: vbapp10.chm583039
f1_keywords:
- vbapp10.chm583039
ms.prod: powerpoint
api_name:
- PowerPoint.Presentation.Close
ms.assetid: 0227528a-4693-dd1a-bb5c-cd31384014b0
ms.date: 06/08/2017
---


# Presentation.Close Method (PowerPoint)

Closes the specified presentation.


## Syntax

 _expression_. **Close**

 _expression_ A variable that represents a **Presentation** object.


## Remarks

When you use this method, PowerPoint will close an open presentation without prompting the user to save their work. To prevent the loss of work, use the  **Save** method or the **SaveAs** method before you use the **Close** method.


## Example

This example closes Pres1.ppt without saving changes.


```vb
With Application.Presentations("pres1.ppt")

    .Saved = True

    .Close

End With
```

This example closes all open presentations.




```vb
With Application.Presentations

    For i = .Count To 1 Step -1

        .Item(i).Close

    Next

End With
```


## See also


#### Concepts


[Presentation Object](presentation-object-powerpoint.md)

