---
title: "Метод ShapeRange.Cut (издатель)"
keywords: vbapb10.chm2293777
f1_keywords: vbapb10.chm2293777
ms.prod: publisher
api_name: Publisher.ShapeRange.Cut
ms.assetid: 961d4646-8318-d2ff-ed98-649583d36115
ms.date: 06/08/2017
ms.openlocfilehash: d44d8999f91f0bbbf8b35c4f26f58fff2966f0b6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shaperangecut-method-publisher"></a>Метод ShapeRange.Cut (издатель)

Удаляет указанный объект и помещает его в буфер обмена.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вырезание**

 переменная _expression_A, представляющий объект **ShapeRange** .


## <a name="remarks"></a>Заметки

Используйте метод **вставьте**Вставка содержимого буфера обмена.

Метод **Copy** можно использовать на **фигуры** , но не удается метод **Paste** .


## <a name="example"></a>Пример

В этом примере удаляется фигуры одно и фигуры двух со страницы, один из активных публикации помещает их копии в буфер обмена и вставляет копии на второй страницы.


```vb
With ActiveDocument 
 .Pages(1).Shapes.Range(Array(1, 2)).Cut 
 .Pages(2).Shapes.Paste 
End With
```

В этом примере удаляется один фигуры на странице один активный публикации и помещает его копию в буфер обмена.




```vb
ActiveDocument
```




```
.Pages(1).Shapes(1).Cut
```

В этом примере удаляет текст в фигуре одно на странице один активный публикации и помещает его копию в буфер обмена.




```vb
ActiveDocument
```




```
.Pages(1).Shapes(1).TextFrame.TextRange.Cut
```


