---
title: Application.LoadPicture Method (Access)
keywords: vbaac10.chm12555
f1_keywords:
- vbaac10.chm12555
ms.prod: access
api_name:
- Access.Application.LoadPicture
ms.assetid: d7e64367-c8f2-22c3-6e6e-18eaae9ed07a
ms.date: 06/08/2017
---


# Application.LoadPicture Method (Access)

The  **LoadPicture** method loads a graphic into an ActiveX control.


## Syntax

 _expression_. **LoadPicture**( ** _FileName_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The file name of the graphic to be loaded. The graphic can be a bitmap file (.bmp), icon file (.ico), run-length encoded file (.rle), or metafile (.wmf).|

### Return Value

Object


## Remarks

Assign the return value of the  **LoadPicture** method to the **Picture** property of an ActiveX control to dynamically load a graphic into the control. The following example loads a bitmap into a control called OLECustomControl on an Orders form:


```vb
Set Forms!Orders!OLECustomControl.Picture = _ 
 LoadPicture("Stars.bmp")
```


 **Note**  You can't use the  **LoadPicture** method to set the **Picture** property of an image control. This method works with ActiveX controls only. To set the **Picture** property of an image control, simply assign to it a string specifying the file name and path of the desired graphic.


## See also


#### Concepts


[Application Object](application-object-access.md)

