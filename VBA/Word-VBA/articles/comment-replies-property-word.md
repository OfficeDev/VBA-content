---
title: Comment.Replies Property (Word)
keywords: vbawd10.chm154993657
f1_keywords:
- vbawd10.chm154993657
ms.prod: word
ms.assetid: a52838be-d6ca-c4e0-56c4-0faf6e86f748
ms.date: 06/08/2017
---


# Comment.Replies Property (Word)

Returns a [Comments](comments-object-word.md) collection of **Comment** objects that are children of the specified comment. Read-only.


## Syntax

 _expression_ . **Replies**

 _expression_ A variable that represents a **Comment** object.


## Remarks

Calling the [Add](comments-add-method-word.md) method on the returned collection of replies adds a new reply, unless the collection was accessed from a reply to a reply.

The [Comments.ShowBy](comments-showby-property-word.md) property fails when called on the **Comments** collection returned by the **Replies** property.


## Property value

 **COMMENTS**


## See also


#### Concepts


[Comment Object](comment-object-word.md)

