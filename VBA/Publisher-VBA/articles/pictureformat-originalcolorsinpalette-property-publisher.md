---
title: "Свойство PictureFormat.OriginalColorsInPalette (издатель)"
keywords: vbapb10.chm3604771
f1_keywords: vbapb10.chm3604771
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalColorsInPalette
ms.assetid: 87c67430-1a5a-47f7-822f-6af8783f73b3
ms.date: 06/08/2017
ms.openlocfilehash: 7b3b5a82c58a6f10610a1b69dcd94276d14128c4
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalcolorsinpalette-property-publisher"></a>Свойство PictureFormat.OriginalColorsInPalette (издатель)

Возвращает значение типа **Long** , представляющее номер цветов в палитре указанного связанного рисунка. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalColorsInPalette**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанных рисунков или объектов, которые не являются TrueColor OLE (то есть, они содержат данные цвета менее 24 бита на канал.) Возвращает фигуры, представляющие внедренных или вставленного изображения и объекты OLE или связанные рисунки, которые являются TrueColor «Отказано в разрешении».

Используйте один из следующих свойств для определения, является ли фигура представляет связанного рисунка:


-  Свойство **[Type](shape-type-property-publisher.md)** объекта **[фигуры](shape-object-publisher.md)**
    
- Свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)**
    


Свойство **[OriginalIsTrueColor](pictureformat-originalistruecolor-property-publisher.md)** определяет, содержит ли связанного рисунка данные цвета 24 бита на канал или более высокой версии.


## <a name="example"></a>Пример

Следующий пример возвращает список всех картинок в активной публикации, которые не являются TrueColor. Возвращается количество цветов в палитре каждого изображения, и если связь изображения и связанных рисунков не TrueColor, также возвращается количество цветов в палитру его.


```vb
Sub PictureColorInformation() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .IsTrueColor = msoFalse Then 
 Debug.Print .Filename 
 Debug.Print "This picture has " &; .ColorsInPalette &; " colors." 
 If .IsLinked = msoTrue Then 
 If .OriginalIsTrueColor = msoFalse Then 
 Debug.Print "The linked picture has " &; _ 
 .OriginalColorsInPalette &; " colors." 
 End If 
 End If 
 End If 
 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 
 
End Sub
```


