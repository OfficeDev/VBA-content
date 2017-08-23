---
title: "Свойство PictureFormat.OriginalHeight (издатель)"
keywords: vbapb10.chm3604774
f1_keywords: vbapb10.chm3604774
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalHeight
ms.assetid: 0bf97bb1-d333-a7ed-686c-da2f3cce97c5
ms.date: 06/08/2017
ms.openlocfilehash: 87ff07a137dc91072524c7de793732378f0bc40b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalheight-property-publisher"></a>Свойство PictureFormat.OriginalHeight (издатель)

Возвращает **Variant** представляющее высота в пунктах указанного связанного рисунка или объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalHeight**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Variant


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанные рисунки или объекты OLE. Возвращает значение «Отказано в разрешении» для фигуры, представляющие внедренные или вставлять рисунки.

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
 Debug.Print "Horizontal Scaling: " &; .HorizontalScale &; "%" 
 Debug.Print "Original Image Height: " &; .OriginalHeight &; " points" 
 Debug.Print "Height in publication: " &; .Height &; " points" 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop 

```


