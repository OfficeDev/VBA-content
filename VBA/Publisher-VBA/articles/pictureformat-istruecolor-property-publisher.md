---
title: "Свойство PictureFormat.IsTrueColor (издатель)"
keywords: vbapb10.chm3604770
f1_keywords: vbapb10.chm3604770
ms.prod: publisher
api_name: Publisher.PictureFormat.IsTrueColor
ms.assetid: 63708d40-996a-67ca-b4eb-dd53c83d1764
ms.date: 06/08/2017
ms.openlocfilehash: 2396596660768d645998afd57e0fbaf33c9ae36b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatistruecolor-property-publisher"></a>Свойство PictureFormat.IsTrueColor (издатель)

Возвращает **MsoTriState** константа, указывающее, содержит ли указанный рисунок или объект OLE данные цвета 24 бита на канал или более высокой версии. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsTrueColor**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Для изображений, которые не являются TrueColor используйте свойство **[ColorsInPalette](pictureformat-colorsinpalette-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** для определения количества цветов в палитре рисунков.

Значение свойства **IsTrueColor** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указанный рисунок не содержит данных цвета 24 бита на канал или более высокой версии.|
| **msoTriStateMixed**|Возвращает значение, указывающее, сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**| Указанный рисунок содержит данные цвета 24 бита на канал или более высокой версии.|

## <a name="example"></a>Пример

В следующем примере проверяется каждого изображения в активном документе и печатает ли изображен TrueColor. Если он не установлен TrueColor, в примере выводится сколько цветов находятся в палитры рисунка.


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


