---
title: "Свойство Page.PageIndex (издатель)"
keywords: vbapb10.chm393224
f1_keywords: vbapb10.chm393224
ms.prod: publisher
api_name: Publisher.Page.PageIndex
ms.assetid: f64cc275-0474-7b97-d840-22e1e576d6f5
ms.date: 06/08/2017
ms.openlocfilehash: 5e313ad0950e89bfe2c654dadaa1f109936385cd
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pagepageindex-property-publisher"></a>Свойство Page.PageIndex (издатель)

Получает индекс страницы в коллекции **страниц** активных документов. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PageIndex**

 переменная _expression_A, представляющий объект **Page** .


## <a name="remarks"></a>Заметки

Значение свойства **PageIndex** назначается каждой странице при его добавлении, а значение увеличивается для каждой дополнительной страницы. При добавлении или удалении страниц, номера страниц индекса изменяется таким образом, что первая страница всегда равно 1 и номера индекса страницы всегда увеличивается на 1. К примеру в публикации с пять страниц, номера страниц индекса будет 1-5, независимо от того, номер страницы, отображаемые на страницах сами.


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


