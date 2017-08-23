---
title: "Метод TextRange.Cut (издатель)"
keywords: vbapb10.chm5308473
f1_keywords: vbapb10.chm5308473
ms.prod: publisher
api_name: Publisher.TextRange.Cut
ms.assetid: c9b8b896-26e7-ac58-0e1a-a66ef789f397
ms.date: 06/08/2017
ms.openlocfilehash: dfc27f397ead93d979c646e75f4e43b241e5ee52
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangecut-method-publisher"></a>Метод TextRange.Cut (издатель)

Удаляет указанный объект и помещает его в буфер обмена.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Вырезание**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Используйте метод **[вставьте](textrange-paste-method-publisher.md)** Вставка содержимого буфера обмена.

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


