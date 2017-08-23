---
title: "Свойство PictureFormat.HasAlphaChannel (издатель)"
keywords: vbapb10.chm3604758
f1_keywords: vbapb10.chm3604758
ms.prod: publisher
api_name: Publisher.PictureFormat.HasAlphaChannel
ms.assetid: 97739201-cd0d-cc78-a28e-935fb11da5b3
ms.date: 06/08/2017
ms.openlocfilehash: 67cfdf3743933d6f46eacf4ca940a0141bd19cdf
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformathasalphachannel-property-publisher"></a>Свойство PictureFormat.HasAlphaChannel (издатель)

Возвращает константу **MsoTriState** , указывающее, содержит ли указанный рисунок альфа-канал. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HasAlphaChannel**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Альфа-канал — это специальные 8-разрядный канал, используемых обработки программного обеспечения изображения для хранения дополнительных данных, например маскирование или сведения о прозрачности.

Значение свойства **HasAlphaChannel** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указанный рисунок не содержит альфа-канал.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Указанный рисунок содержит альфа-канал.|

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


