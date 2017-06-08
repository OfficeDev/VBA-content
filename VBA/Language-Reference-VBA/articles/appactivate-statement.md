---
title: AppActivate Statement
keywords: vblr6.chm1008855
f1_keywords:
- vblr6.chm1008855
ms.prod: office
ms.assetid: 8af4340f-e249-6806-044e-a68bf06ff3f6
ms.date: 06/08/2017
---


# AppActivate Statement

Activates an application window.

 **Syntax**

 **AppActivate** **_title_** [, **_wait_** ]

The  **AppActivate** statement syntax has these[named arguments](vbe-glossary.md):


|**Part**|**Description**|
|:-----|:-----|
|**_title_**|Required. [String expression](vbe-glossary.md) specifying the title in the title bar of the application window you want to activate. The task ID returned by the **Shell** function can be used in place of **_title_** to activate an application.|
|**_wait_**|Optional. [Boolean](vbe-glossary.md) value specifying whether the calling application has the focus before activating another. If **False** (default), the specified application is immediately activated, even if the calling application does not have the focus. If **True**, the calling application waits until it has the focus, then activates the specified application.|
 **Remarks**
The  **AppActivate** statement changes the focus to the named application or window but does not affect whether it is maximized or minimized. Focus moves from the activated application window when the user takes some action to change the focus or close the window. Use the **Shell** function to start an application and set the window style.
In determining which application to activate,  **_title_** is compared to the title string of each running application. If there is no exact match, any application whose title string begins with **_title_** is activated. If there is more than one instance of the application named by **_title_**, one instance is arbitrarily activated.

## Example

This example illustrates various uses of the  **AppActivate** statement to activate an application window. The **Shell** statements assume the applications are in the paths specified. On the Macintosh, the default drive name is "HD" and portions of the pathname are separated by colons instead of backslashes.


```vb
Dim MyAppID, ReturnValue 
AppActivate "Microsoft Word" ' Activate Microsoft 
 ' Word. 
 
' AppActivate can also use the return value of the Shell function. 
MyAppID = Shell("C:\WORD\WINWORD.EXE", 1) ' Run Microsoft Word. 
AppActivate MyAppID ' Activate Microsoft 
 ' Word. 
 
 
 
' You can also use the return value of the Shell function. 
ReturnValue = Shell("c:\EXCEL\EXCEL.EXE",1) ' Run Microsoft Excel. 
AppActivate ReturnValue ' Activate Microsoft 
 ' Excel. 

```


