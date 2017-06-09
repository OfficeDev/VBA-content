---
title: What is the difference between the DataObject and the Clipboard?
keywords: fm20.chm5225203
f1_keywords:
- fm20.chm5225203
ms.prod: office
ms.assetid: 8f3e37d4-1f31-ed57-d6d8-d579d9cd22b9
ms.date: 06/08/2017
---


# What is the difference between the DataObject and the Clipboard?

The  **DataObject** and the Clipboard both provide a means to move data from one place to another. As an application developer, there are several important points to remember when you use either a **DataObject** or the Clipboard:



- You can store more than one piece of data at a time on either a  **DataObject** or the Clipboard as long as each piece of data has a different[data format](glossary-vba.md). If you store data with a format that is already in use, the new data is saved and the old data is discarded.
    
- The Clipboard supports picture formats and text formats. A  **DataObject** currently supports only text formats.
    
- A  **DataObject** exists only while your application is running; the Clipboard exists as long as the operating system is running. This means you can put data on the Clipboard and close an application without losing the data. The same is not true with the **DataObject**. If you close the application that put data on a **DataObject**, you lose the data.
    
- A  **DataObject** is a standard OLE object, while the Clipboard is not. This means the Clipboard can support standard move operations (copy, cut, and paste) but not drag-and-drop operations. You must use the **DataObject** if you want your application to support drag-and-drop operations.
    


 **Tip**  You can define your own data format names when you use the  **SetText** method to move data to the Clipboard or a **DataObject**. This can help distinguish between text that your application moves and text that the user moves.


