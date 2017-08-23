---
title: "Метод TextRange.Copy (издатель)"
keywords: vbapb10.chm5308480
f1_keywords: vbapb10.chm5308480
ms.prod: publisher
api_name: Publisher.TextRange.Copy
ms.assetid: e0d92492-fa0e-9424-471d-09866402702c
ms.date: 06/08/2017
ms.openlocfilehash: 3d8480a953b8aaabafe379c55df72e4483d5c228
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="textrangecopy-method-publisher"></a>Метод TextRange.Copy (издатель)

Копирует указанный объект в буфер обмена.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Копия**

 переменная _expression_A, представляющий объект **TextRange** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Используйте метод **вставьте**Вставка содержимого буфера обмена.

Метод **Copy** можно использовать на **фигуры** , но не удается метод **Paste** .


## <a name="example"></a>Пример

В этом примере копируется фигур первый и второй на странице один из активных публикации в буфер обмена и вставляет копии на второй страницы.


```vb
With ActiveDocument 
 .Pages(1).Shapes.Range(Array(1, 2)).Copy 
 .Pages(2).Shapes.Paste 
End With
```

В этом примере копирует один фигуры на странице один из активных публикации в буфер обмена.




```vb
ActiveDocument.Pages(1).Shapes(1).Copy
```

В этом примере копирует текст в фигуре одно на странице один из активных публикации в буфер обмена.




```vb
ActiveDocument.Pages(1).Shapes(1).TextFrame.TextRange.Copy
```


