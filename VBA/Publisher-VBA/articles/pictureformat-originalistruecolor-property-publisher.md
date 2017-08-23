---
title: "Свойство PictureFormat.OriginalIsTrueColor (издатель)"
keywords: vbapb10.chm3604775
f1_keywords: vbapb10.chm3604775
ms.prod: publisher
api_name: Publisher.PictureFormat.OriginalIsTrueColor
ms.assetid: 837109d4-3479-2500-a1fa-b4c00e0f8672
ms.date: 06/08/2017
ms.openlocfilehash: 99c3808681bce1bc96e0daac855e8518067f7dd1
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatoriginalistruecolor-property-publisher"></a>Свойство PictureFormat.OriginalIsTrueColor (издатель)

Возвращает **MsoTriState** константа, указывающее, является ли указанный связанных рисунков или объекта OLE содержит данные цвета 24 бита на канал или более высокой версии. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **OriginalIsTrueColor**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Это свойство применяется только к связанные рисунки или объекты OLE. Возвращает «Отказано в разрешении» для фигуры, представляющие внедренных или вставленного изображения и объекты OLE.

Чтобы определить, является ли фигура представляет связанного рисунка, используйте свойство **[Type](shape-type-property-publisher.md)** объекта **[Shape](shape-object-publisher.md)** или свойство **[IsLinked](pictureformat-islinked-property-publisher.md)** объекта **[PictureFormat](pictureformat-object-publisher.md)** .

Значение свойства **OriginalIsTrueColor** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указанный связанного рисунка не содержит данных цвета 24 бита на канал или более высокой версии.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Указанный связанный рисунок содержит данные цвета 24 бита на канал или более высокой версии.|

## <a name="example"></a>Пример

Следующий пример возвращает список изображений в активном документе, которые TrueColor. Если связанный рисунок и связанного рисунка является также TrueColor, эти сведения также возвращается.


```vb
Sub PictureColorInformation() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 If shpLoop.Type = pbLinkedPicture Or shpLoop.Type = pbPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .IsTrueColor = msoTrue Then 
 Debug.Print .Filename 
 Debug.Print "This picture is TrueColor" 
 If .IsLinked = msoTrue Then 
 If .OriginalIsTrueColor = msoTrue Then 
 Debug.Print "The linked picture is also TrueColor." 
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


