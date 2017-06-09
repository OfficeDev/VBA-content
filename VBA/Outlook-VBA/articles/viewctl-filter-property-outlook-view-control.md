---
title: ViewCtl.Filter Property (Outlook View Control)
ms.prod: outlook
ms.assetid: 4074d1d3-e3b5-810f-3ba9-3cf5bd1507ab
ms.date: 06/08/2017
---


# ViewCtl.Filter Property (Outlook View Control)

Returns or sets a  **String**that represents the Distributed Authoring and Versioning (DAV) Searching and Locating (DASL) statement used to restrict the display to a specified subset of data. Read/write.


## Syntax

 _expression_. **Filter**

 _expression_A variable that represents a  **ViewCtl** object.


## Remarks

To reset a filter back to its default value, type the following line of code. 


```vb
object.Filter = " ""DAV:isfolder"" = False And ""DAV:ishidden"" = False "
```


