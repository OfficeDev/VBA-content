---
title: PublishObject.SlideShowName Property (PowerPoint)
keywords: vbapp10.chm635007
f1_keywords:
- vbapp10.chm635007
ms.prod: powerpoint
api_name:
- PowerPoint.PublishObject.SlideShowName
ms.assetid: 8555cc11-e221-4bcf-3ea7-84e242985814
ms.date: 06/08/2017
---


# PublishObject.SlideShowName Property (PowerPoint)

Returns or sets the name of the custom slide show to be published as a Web presentation. Read/write.


## Syntax

 _expression_. **SlideShowName**

 _expression_ A variable that represents a **PublishObject** object.


### Return Value

String


## Example

The following example saves the current presentation as an HTML version 4.0 file with the name "mallard.htm." It then displays a message indicating that the current named presentation is being saved in both PowerPoint and HTML formats.


```vb
With Pres.PublishObjects(1)
    PresName = .SlideShowName
    .SourceType = ppPublishAll
    .FileName = "C:\HTMLPres\mallard.htm"
    .HTMLVersion = ppHTMLVersion4
    MsgBox ("Saving presentation " &; "'" _
        &; PresName &; "'" &; " in PowerPoint" _
        &; Chr(10) &; Chr(13) _
        &; " format and HTML version 4.0 format")
    .Publish
End With
```


## See also


#### Concepts


[PublishObject Object](publishobject-object-powerpoint.md)

