---
title: Document.PreviewPicture Property (Visio)
keywords: vis_sdr.chm10550775
f1_keywords:
- vis_sdr.chm10550775
ms.prod: visio
api_name:
- Visio.Document.PreviewPicture
ms.assetid: 4354f66b-6f0b-1511-3c77-fc7cd58f539e
ms.date: 06/08/2017
---


# Document.PreviewPicture Property (Visio)

Gets or sets the preview picture shown in the  **Open** dialog box and when you click the **File** tab, and then click **New**. Read/write.


## Syntax

 _expression_ . **PreviewPicture**

 _expression_ A variable that represents a **Document** object.


### Return Value

IPictureDisp


## Remarks

The  **PreviewPicture** property returns and accepts only EMF files (enhanced metafiles). Microsoft Visio will raise an exception if the file you pass to **PreviewPicture** is a non-EMF file.

To delete an existing preview, set the  **PreviewPicture** property to **Nothing** .

You can use the  **PreviewPicture** property to include a preview picture in a template that does not have any diagrams stored in it.

COM provides a standard implementation of a picture object that has the  **IPictureDisp** interface on top of the underlying system picture support. The **IPictureDisp** interface exposes a picture object's properties and is implemented in the **stdole** type library as a **StdPicture** object creatable within Microsoft Visual Basic. The **stdole** type library is automatically referenced from all Visual Basic for Applications (VBA) projects in Visio.

To get information about the  **StdPicture** object that supports the **IPictureDisp** interface:




1. In the  **Code** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab, click **Visual Basic**.
    
2. On the  **View** menu, click **Object Browser**.
    
3. In the  **Project/Library** list, click **stdole**.
    
4. Under  **Classes**, examine the class named  **StdPicture** .
    
For details about the  **IPictureDisp** interface, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

Currently, only in-process solutions can use the  **PreviewPicture** property because the **IPictureDisp** interface cannot be marshaled.


