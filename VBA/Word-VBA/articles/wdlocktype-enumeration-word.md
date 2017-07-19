---
title: WdLockType Enumeration (Word)
ms.prod: word
api_name:
- Word.WdLockType
ms.assetid: 2c861165-b6b7-5518-1569-628469b099cd
ms.date: 06/08/2017
---


# WdLockType Enumeration (Word)

Specifies the type of lock for a  **[CoAuthLock](coauthlock-object-word.md)** object.



|**Name**|**Value**|**Description**|
|:-----|:-----|:-----|
| **wdLockChanged**|3|Specifies a placeholder lock. A placeholder lock indicates that another user has removed their lock from the range, but the current user has not updated their view of the document by saving.|
| **wdLockEphemeral**|2|Specifies an ephemeral lock. Word implicitly places an ephemeral lock on a range when a user begins editing a range in a document with coauthoring enabled.|
| **wdLockNone**|0|Reserved for future use.|
| **wdLockReservation**|1|Specifies a reservation lock. A reservation lock is explicitly created by a user through the  **Block Authors** button on the **Review** tab in Word.|

