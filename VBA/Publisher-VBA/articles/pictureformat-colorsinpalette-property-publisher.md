---
title: "Свойство PictureFormat.ColorsInPalette (издатель)"
keywords: vbapb10.chm3604754
f1_keywords: vbapb10.chm3604754
ms.prod: publisher
api_name: Publisher.PictureFormat.ColorsInPalette
ms.assetid: 34e671b1-af0e-0dac-1429-246facae975b
ms.date: 06/08/2017
ms.openlocfilehash: 5c783a8582639ffede535cbe858ea4a8ee2f2614
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatcolorsinpalette-property-publisher"></a>Свойство PictureFormat.ColorsInPalette (издатель)

 Возвращает значение типа **Long** , представляющее номер цветов в палитре рисунков. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ColorsInPalette**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Это свойство применяется только к изображений, которые не являются TrueColor (изображения, содержащие данные цвет менее 24 бита на канал). Возвращает фигуры, представляющие изображения, TrueColor «Отказано в разрешении».

Используйте свойство **[IsTrueColor](pictureformat-istruecolor-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** для определения того, содержит ли изображение данных цвета 24 бита на канал или более высокой версии.


## <a name="example"></a>Пример

В следующем примере проверяется каждого изображения в активном документе и печатает ли изображен TrueColor. Если изображение не TrueColor, в примере выводится сколько цветов находятся в палитры рисунка.


```vb
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 Debug.Print .Filename 
 If .IsTrueColor = msoTrue Then 
 Debug.Print "This picture is TrueColor" 
 Else 
 Debug.Print "This picture contains " &; .ColorsInPalette &; " colors." 
 End If 
 End If 
 End With 
 
 End If 
 Next shpLoop 
Next pgLoop 

```


