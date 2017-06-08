---
title: Comments.Add2 Method (PowerPoint)
keywords: vbapp10.chm641005
f1_keywords:
- vbapp10.chm641005
ms.assetid: 4add4727-0193-061b-da71-793a4d6b3aa9
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Comments.Add2 Method (PowerPoint)

Returns a  **[Comment](comment-object-powerpoint.md)** object that represents a new comment added to a slide.


## Syntax

 _expression_. **Add2**_(Left,_ _Top,_ _Author,_ _AuthorInitials,_ _Text,_ _ProviderID,_ _UserID)_

 _expression_ A variable that represents a **Comments** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Float**|The position, measured in points, of the left edge of the comment, relative to the left edge of the presentation.|
| _Top_|Required|**Float**|The position, measured in points, of the top edge of the comment, relative to the top edge of the presentation.|
| _Author_|Required|**String**|The author of the comment.|
| _AuthorInitials_|Required|**String**|The author's initials.|
| _Text_|Required|**String**|The comment's text. |
| _ProviderID_|Required|**String**|The service that provides contact information.Example: ?AD? (Active Directory)|
| _UserID_|Required|**String**|The ID of the user providing the comment.|
| _Left_|Required|FLOAT||
| _Top_|Required|FLOAT||
| _Author_|Required|STRING||
| _AuthorInitials_|Required|STRING||
| _Text_|Required|STRING||
| _ProviderID_|Required|STRING||
| _UserID_|Required|STRING||

### Return value

 **COMMENT**


