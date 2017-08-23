---
title: "Свойство ReaderSpread.PageCount (издатель)"
keywords: vbapb10.chm524294
f1_keywords: vbapb10.chm524294
ms.prod: publisher
api_name: Publisher.ReaderSpread.PageCount
ms.assetid: 39d26cd7-f4b8-bbf3-a2a8-32a4c9362e30
ms.date: 06/08/2017
ms.openlocfilehash: 15831653d26be41a3dc952fc814635acb8a5a4f6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="readerspreadpagecount-property-publisher"></a>Свойство ReaderSpread.PageCount (издатель)

Возвращает значение типа **Long** , указывающее количество страниц в указанном ширина чтения. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageCount**

 переменная _expression_A, представляет собой объект- **ReaderSpread** .


### <a name="return-value"></a>Возвращаемое значение

Длинный


## <a name="remarks"></a>Заметки

Распространение читатель может содержать не более двух страниц.


## <a name="example"></a>Пример

В следующем примере проверяется распространения чтения третьей страницы в активной публикации для просмотра, если он содержит несколько страниц, а затем отображает общее число страниц в распространении.


```vb
Sub NumberOfPagesInSpread() 
 If ActiveDocument.Pages(3).ReaderSpread.PageCount > 1 Then 
 MsgBox "The spread has two pages." 
 Else 
 MsgBox "The spread has only one page." 
 End If 
End Sub
```


