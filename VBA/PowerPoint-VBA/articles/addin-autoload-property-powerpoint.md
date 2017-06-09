---
title: AddIn.AutoLoad Property (PowerPoint)
keywords: vbapp10.chm521007
f1_keywords:
- vbapp10.chm521007
ms.prod: powerpoint
api_name:
- PowerPoint.AddIn.AutoLoad
ms.assetid: ba8eca66-6d94-62ca-0270-85f2a508299f
ms.date: 06/08/2017
---


# AddIn.AutoLoad Property (PowerPoint)

Determines whether the specified add-in is automatically loaded each time PowerPoint is started. Read/write.


## Syntax

 _expression_. **AutoLoad**

 _expression_ A variable that represents an **AddIn** object.


### Return Value

MsoTriState


## Remarks

Setting this property to  **msoTrue** automatically sets the **[Registered](addin-registered-property-powerpoint.md)** property to **msoTrue**.

The value of the  **AutoLoad** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified add-in is not automatically loaded each time PowerPoint is started. |
|**msoTrue**| The specified add-in is automatically loaded each time PowerPoint is started.|

## Example

This example displays the name of each add-in that's automatically loaded each time PowerPoint is started.


```vb
For Each myAddIn In AddIns

    If myAddIn.AutoLoad Then

        MsgBox myAddIn.Name

        afound = True

    End If

Next myAddIn

If afound <> True Then 

    MsgBox "No add-ins were loaded automatically."

End If
```

This example specifies that the add-in named "myTools" be loaded automatically each time PowerPoint is started.




```vb
Application.AddIns("mytools").AutoLoad = msoTrue
```


## See also


#### Concepts


[AddIn Object](addin-object-powerpoint.md)

