---
title: "Свойство PictureFormat.OriginalWidth (издатель)"
keywords: vbapb10.chm3604777
f1_keywords: vbapb10.chm3604777
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalWidth
ms.assetid: 3c418f3f-b2af-3176-9a37-a548b15fb4bc
ms.date: 06/08/2017
ms.openlocfilehash: 0f6699f4a5843fbf62daa2a906d3d2e98dfb2350
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalwidth-property-publisher"></a>Свойство PictureFormat.OriginalWidth (издатель)

Возвращает **Variant** , который представляет, в точках, ширина указанного связанного рисунка или объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalWidth**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанных рисунков. Возвращает значение «Отказано в разрешении» для фигуры, представляющие внедренные или вставлять рисунки.

Чтобы определить, является ли фигура представляет связанного рисунка, используйте свойство **[Type](shape-type-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** или свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** .


## <a name="example"></a>Пример

В следующем примере проверяется каждого изображения в активной публикации и возвращает свойства выбранного изображения для изображений, которые связаны.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Vertical Scaling: " &; .VerticalScale &; "%" 
 Debug.Print "Original Image Width: " &; .OriginalWidth &; " points" 
 Debug.Print "Width in publication: " &; .Width &; " points" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


