---
title: MasterShortcut.Icon Property (Visio)
keywords: vis_sdr.chm16050665
f1_keywords:
- vis_sdr.chm16050665
ms.prod: visio
api_name:
- Visio.MasterShortcut.Icon
ms.assetid: c0dafec6-ca3b-f791-c5b1-3da19f0d98ee
ms.date: 06/08/2017
---


# MasterShortcut.Icon Property (Visio)

Returns the icon contained in a master shortcut. Read/write.


## Syntax

 _expression_ . **Icon**

 _expression_ A variable that represents a **MasterShortcut** object.


### Return Value

IPictureDisp


## Remarks

The  **Icon** property returns and accepts only HICON files. Microsoft Visio raises an exception if the object returned is a non-HICON file.

COM provides a standard implementation of a picture object that has the  **IPictureDisp** interface on top of the underlying system picture support. The **IPictureDisp** interface exposes a picture object's properties and is implemented in the stdole type library as a **StdPicture** object creatable within Microsoft Visual Basic. The stdole type library is automatically referenced from all Visual Basic for Applications (VBA) projects in Visio.

To get information about the  **StdPicture** object that supports the **IPictureDisp** interface:




1. In the  **Code** group on the[Developer](http://msdn.microsoft.com/library/1bdc55f5-8fc7-7257-03d5-c049eceb29ff%28Office.15%29.aspx) tab, click **Visual Basic**.
    
2. On the  **View** menu, click **Object Browser**.
    
3. In the  **Project/Library** list, click **stdole**.
    
4. Under  **Classes**, examine the class named  **StdPicture** .
    


For details about the  **IPictureDisp** interface, see the Microsoft Platform SDK on MSDN, the Microsoft Developer Network.

Currently, only in-process solutions can use the  **Icon** property because the **IPictureDisp** interface cannot be marshaled.


