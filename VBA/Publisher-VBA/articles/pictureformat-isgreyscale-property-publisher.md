---
title: "Свойство PictureFormat.IsGreyScale (издатель)"
keywords: vbapb10.chm3604768
f1_keywords: vbapb10.chm3604768
ms.prod: publisher
api_name: Publisher.PictureFormat.IsGreyScale
ms.assetid: 1f8308c1-353e-2aac-9b4b-fad300a89b97
ms.date: 06/08/2017
ms.openlocfilehash: 6a972affe0b4b9263c05e5400a767a8952e3a041
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatisgreyscale-property-publisher"></a>Свойство PictureFormat.IsGreyScale (издатель)

Возвращает константу **MsoTriState** , которое указывает, является ли рисунок оттенки серого изображения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IsGreyScale**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **IsGreyScale** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**|Изображение не оттенки серого изображения.|
| **msoTriStateMixed**|Указывает сочетание **msoTrue** и **msoFalse** для диапазона указанной фигуры.|
| **msoTrue**|Указанный рисунок является оттенки серого изображения.|

## <a name="example"></a>Пример

Следующий пример возвращает список оттенки серого изображения, содержащиеся в активной публикации.


```vb
Sub ListGreyScalePictures() 
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
 For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse And .IsGreyScale = msoCTrue Then 
 
 Debug.Print .Filename 
 Debug.Print "Page " &; pgLoop.PageNumber 
 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
 Next pgLoop 
 
End Sub
```


