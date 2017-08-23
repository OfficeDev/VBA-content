---
title: "Свойство PictureFormat.OriginalHasAlphaChannel (издатель)"
keywords: vbapb10.chm3604773
f1_keywords: vbapb10.chm3604773
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalHasAlphaChannel
ms.assetid: e58a97d2-4ced-d3cf-56b2-6a89df02bcdf
ms.date: 06/08/2017
ms.openlocfilehash: e56fdb2f07e3a201954d69c8bb5935e3acb0e590
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalhasalphachannel-property-publisher"></a>Свойство PictureFormat.OriginalHasAlphaChannel (издатель)

Возвращает константу **MsoTriState** в зависимости от того, содержит ли исходный, связанных рисунков альфа-канал. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalHasAlphaChannel**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанных рисунков. Возвращает значение «Отказано в разрешении» для фигуры, представляющие внедренные или вставлять рисунки.

Используйте один из следующих свойств для определения, является ли фигура представляет связанного рисунка:


-  Свойство **[Type](shape-type-property-publisher.md)** объекта **[фигуры](shape-object-publisher.md)**
    
- Свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)**
    


Альфа-канал — это специальные 8-разрядный канал, используемых обработки программного обеспечения изображения для хранения дополнительных данных, таких как маскирование сведения или прозрачность сведений.

Значение свойства **OriginalHasAlphaChannel** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Исходный, связанных рисунков не содержит альфа-канал.|
| **msoTriStateMixed**| Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Исходный, связанных рисунков содержит альфа-канал.|

## <a name="example"></a>Пример

В следующем примере возвращается, содержит ли первую фигуру на первой странице активная публикация альфа-канал. Если связь изображения, а исходный рисунок содержит альфа-канал, также возвращаются. В этом примере предполагается, что фигурой является рисунок.


```vb
With ActiveDocument.Pages(1).Shapes(1).PictureFormat 
 If .HasAlphaChannel = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture contains an alpha channel." 
 
 If .IsLinked = msoTrue Then 
 If .OriginalHasAlphaChannel = msoTrue Then 
 Debug.Print "The linked picture " &; _ 
 "also contains an alpha channel." 
 End If 
 End If 
 End If 
End With 

```


