---
title: Comment.Replies Property (PowerPoint)
keywords: vbapp10.chm642014
f1_keywords:
- vbapp10.chm642014
ms.assetid: 3af06afb-e507-bb3b-901b-30bf6bbfa0ef
ms.date: 06/08/2017
ms.prod: powerpoint
---


# Comment.Replies Property (PowerPoint)

Returns a [Comments](comments-object-powerpoint.md) collection of **Comment** objects that are children of the specified comment. Read-only.


## Syntax

 _expression_. **Replies**

 _expression_ A variable that represents a **Comment** object.


## Remarks

Calling the [Add](comments-add-method-powerpoint.md) method on the returned collection of replies adds a new reply, unless the collection was accessed from a reply to a reply.


## Property value

 **COMMENTS**


