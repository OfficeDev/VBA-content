---
title: "Свойство PictureFormat.OriginalFileSize (издатель)"
keywords: vbapb10.chm3604772
f1_keywords: vbapb10.chm3604772
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalFileSize
ms.assetid: 30704f2a-d739-7f14-d69a-73ab1f5ab8f3
ms.date: 06/08/2017
ms.openlocfilehash: d5802183ba531b424aa37cdabacc2b55edac58d9
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalfilesize-property-publisher"></a>Свойство PictureFormat.OriginalFileSize (издатель)

Возвращает значение типа **Long** , представляющее размер, в байтах, рисунка или объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalFileSize**

 переменная _expression_A, представляющий объект **PictureFormat** .


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанных рисунков. Возвращает значение «Отказано в разрешении» для фигуры, представляющие внедренные или вставлять рисунки.

Используйте один из следующих свойств для определения, является ли фигура представляет связанного рисунка:


-  Свойство **[Type](shape-type-property-publisher.md)** объекта **[фигуры](shape-object-publisher.md)**
    
- Свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)**
    



## <a name="example"></a>Пример

В следующем примере проверяется каждого изображения в активной публикации и печатает свойства выбранного изображения для изображений, которые связаны.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Original File Size: " &; .OriginalFileSize &; " bytes" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


