---
title: "Свойство PictureFormat.IsLinked (издатель)"
keywords: vbapb10.chm3604769
f1_keywords: vbapb10.chm3604769
ms.prod: publisher
api_name: Publisher.PictureFormat.IsLinked
ms.assetid: 2215cee8-864d-7228-8692-a428385d2be2
ms.date: 06/08/2017
ms.openlocfilehash: 02b24287a943c2175358ac8dcd2c50fb297293f6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatislinked-property-publisher"></a>Свойство PictureFormat.IsLinked (издатель)

Возвращает константу **MsoTriState** , указывающее, является ли указанный рисунок рисунка или объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsLinked**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Возвращает **msoFalse** для вставленных или внедренного изображения и объекты OLE.

Если связанного рисунка или объекта OLE несколько дополнительных свойств **[PictureFormat](pictureformat-object-publisher.md)** объект реагирования на исходный рисунок (например, ** [OriginalFileSize](pictureformat-originalfilesize-property-publisher.md)**) доступны.

Значение свойства **IsLinked** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указанный рисунок не связанного рисунка.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Указанный рисунок является связанного рисунка.|

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


