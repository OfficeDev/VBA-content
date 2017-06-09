---
title: COMAddIn.Object Property (Office)
keywords: vbaof11.chm219007
f1_keywords:
- vbaof11.chm219007
ms.prod: office
api_name:
- Office.COMAddIn.Object
ms.assetid: 20dd8eca-6f8e-7445-ec0c-a29b29409c58
ms.date: 06/08/2017
---


# COMAddIn.Object Property (Office)

Gets or sets an object reference. Read/write.


## Syntax

 _expression_. **Object**

 _expression_ A variable that represents a **COMAddIn** object.


## Remarks

The  **Object** property is a read/write property in which any object reference can be stored. In this regard, it is similar to the general-purpose **Tag** property of certain ActiveX controls.

In some cases, the  **Object** property returns the object represented by the specified **COMAddIn** object; otherwise, it returns **Nothing** by default.


## Example

The following example returns the object represented by the COM add-in  **msodraa9.ShapeSelect**.


```
Dim objBaseObject As Object 
Set objBaseObject = _ 
 Application.COMAddIns.Item("msodraa9.ShapeSelect").Object
```


## See also


#### Concepts


[COMAddIn Object](comaddin-object-office.md)
#### Other resources


[COMAddIn Object Members](comaddin-members-office.md)

