---
title: "Свойство PictureFormat.Filename (издатель)"
keywords: vbapb10.chm3604756
f1_keywords: vbapb10.chm3604756
ms.prod: publisher
api_name: Publisher.PictureFormat.Filename
ms.assetid: 73e2a224-f15a-50cc-462e-10ccf9478122
ms.date: 06/08/2017
ms.openlocfilehash: 842ea9c09db35f5abfa00d60bf61fdf27f0b226a
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatfilename-property-publisher"></a>Свойство PictureFormat.Filename (издатель)

Возвращает **строку** , представляющую имя файла, указанного изображения или объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Имя файла**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="remarks"></a>Заметки

Связанные рисунки и объекты OLE возвращаемая строка представляет полный путь и имя рисунка. Внедренные изображения и объекты OLE возвращаемая строка представляет только имя файла.

Чтобы определить, является ли фигура представляет связанного рисунка, используйте свойство **[Type](shape-type-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** или свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** .


## <a name="example"></a>Пример

Следующий пример возвращает свойства выбранного изображения для каждого изображения в активной публикации.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 
 If .IsEmpty = msoFalse Then 
 
 Debug.Print "File Name: " &; .Filename 
 Debug.Print "Horizontal Scaling: " &; .HorizontalScale &; "%" 
 Debug.Print "Vertical Scaling: " &; .VerticalScale &; "%" 
 Debug.Print "File size in publication: " &; .FileSize &; " bytes" 
 
 End If 
 
 End With 
 End If 
 Next shpLoop 
Next pgLoop
```


