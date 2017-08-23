---
title: "Свойство Page.PageID (издатель)"
keywords: vbapb10.chm393223
f1_keywords: vbapb10.chm393223
ms.prod: publisher
api_name: Publisher.Page.PageID
ms.assetid: 07a87780-fb97-93ff-6f7d-1f1b72d3cb6a
ms.date: 06/08/2017
ms.openlocfilehash: cf5dfa1f8ebcbfdca49fdb2e694c208b38ef8d69
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagepageid-property-publisher"></a>Свойство Page.PageID (издатель)

Возвращает значение типа **Long** , указывающее, уникальный идентификатор для страницы в публикации. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageID**

 переменная _expression_A, представляющий объект **Page** .


## <a name="remarks"></a>Заметки

 **PageID** значения — случайных чисел, назначенных на страницы после их добавления. При добавлении или удалении страниц этих уникальных значений не изменяется. Кроме того эти номера не начинаются с 1, они не являются непрерывными.


## <a name="example"></a>Пример

Следующий пример отображает свойства **PageIndex**, **PageNumber**и **PageID** для всех страниц в активной публикации.


```vb
Dim lngLoop As Long 
 
With ActiveDocument.Pages 
 For lngLoop = 1 To .Count 
 With .Item(lngLoop) 
 Debug.Print "PageIndex = " &; .PageIndex _ 
 &; " / PageNumber = " &; .PageNumber _ 
 &; " / PageID = " &; .PageID 
 End With 
 Next lngLoop 
End With
```


