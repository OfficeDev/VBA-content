---
title: "Метод Shape.Copy (издатель)"
keywords: vbapb10.chm2228242
f1_keywords: vbapb10.chm2228242
ms.prod: publisher
api_name: Publisher.Shape.Copy
ms.assetid: cfec06d8-9f9b-4d88-eb28-e9e29fb1aed1
ms.date: 06/08/2017
ms.openlocfilehash: 25b822cda35efaa8f279435479c3bced79a1032d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="shapecopy-method-publisher"></a>Метод Shape.Copy (издатель)

Копирует указанный объект в буфер обмена.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Копия**

 переменная _expression_A, представляющий объект **фигуры** .


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


