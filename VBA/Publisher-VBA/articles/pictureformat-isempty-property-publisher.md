---
title: "Свойство PictureFormat.IsEmpty (издатель)"
keywords: vbapb10.chm3604788
f1_keywords: vbapb10.chm3604788
ms.prod: publisher
api_name: Publisher.PictureFormat.IsEmpty
ms.assetid: 493cbb8f-e069-14a9-a827-7f7631eb3a09
ms.date: 06/08/2017
ms.openlocfilehash: 638be845247e67e27eb5848935af0bf8558032cc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatisempty-property-publisher"></a>Свойство PictureFormat.IsEmpty (издатель)

Возвращает константу **MsoTriState** , которое указывает, находится ли указанный фигуры пустой рамки. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Функция IsEmpty**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **IsEmpty** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Указанный фигура не является пустой рамки.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Указанный фигуры является пустой рамки.|

## <a name="example"></a>Пример

В следующем примере проверяет каждого изображения в активной публикации и если это еще не пустой рамки, печатает свойства выбранного изображения для рисунка.


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


