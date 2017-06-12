---
title: DrawingControl.Src Property (Visio)
keywords: vis_sdr.chm51015
f1_keywords:
- vis_sdr.chm51015
ms.prod: visio
api_name:
- Visio.Src
ms.assetid: 396fd4cc-c408-4fa5-7089-e5658593cae5
ms.date: 06/08/2017
---


# DrawingControl.Src Property (Visio)

The full path, including file name, of the Microsoft Visio file used to initialize a  **DrawingControl** object. Read/write.


## Syntax

 _expression_ . **Src**

 _expression_ A variable that represents a **DrawingControl** object.


### Return Value

String


## Remarks

You can set the  **Src** property either at design time (for example, in the **Properties** window in Microsoft Visual Basic 6.0), or at run time, typically in the **Form_Load()** procedure, as shown in the following example. It is recommended that you set **Src** at design time.

The source file can be any valid Visio file type, including .vsd, .vdx, .vss, .vsx, .vst, and .vtx. You can load a file that you distribute with your program or any other file that users have access to, including files on their computer, on their local network, or on the Web.

When it attempts to open a file you specify in the  **Src** property, the Visio Drawing Control behaves in much the same way Visio does when it attempts to open a file; it first searches all the paths listed in the **DrawingPaths** property string. If the path to the file you want to open is listed in that string, you can specify just the file name. If not, you must specify a fully-qualified path and file name.

When you set the  **Src** property to load a file into the Visio Drawing Control, the control opens a copy of the file, but does not keep the original file open for writing. As a result, you cannot use the ** Document.Save** method to save changes to a file loaded into the Visio Drawing Control. To save changes in a file, first use the **Src** property to load the file into the control, and then set **Src** to an empty string (""). To save the modified file to disk, use the **Document.SaveAs** method.

If you do not set the  **Src** property to an empty string after loading a drawing into the Visio Drawing Control, each time you close and reopen your application, the original drawing will be loaded, and any modifications you or your users have made will be lost.


## Example

The following example shows how to set the  **Src** property at run time in the **Form_Load()** sub procedure of your Visual Basic program. Before running this example, replace _Fullpath\filename_ with the full path to and file name of the Visio file you want to display in the Visio Drawing Control window.


```vb
Private Sub Form_Load() 
 
 vsoDrawingControl.Src = "Fullpath\filename " 
 
End Sub
```


