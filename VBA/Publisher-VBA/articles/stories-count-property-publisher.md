---
title: "Свойство Stories.Count (издатель)"
keywords: vbapb10.chm5701635
f1_keywords: vbapb10.chm5701635
ms.prod: publisher
api_name: Publisher.Stories.Count
ms.assetid: 3380c5fc-cfd7-98d6-9c19-4a2fe9877166
ms.date: 06/08/2017
ms.openlocfilehash: 4c50c2bf314fb93bd9b863d5280c5d8af21d5d02
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="storiescount-property-publisher"></a>Свойство Stories.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **функциональности** .


## <a name="example"></a>Пример

В этом примере отображается число страниц в активный документ.


```vb
Sub CountNumberOfPages() 
 MsgBox "Your publication contains " &; _ 
 ActiveDocument.Pages.Count &; " page(s)." 
End Sub
```

В этом примере отображается количество фигур в активном документе.




```vb
Sub CountNumberOfShapes() 
 Dim intShapes As Integer 
 Dim pg As Page 
 
 For Each pg In ActiveDocument.Pages 
 intShapes = intShapes + pg.Shapes.Count 
 Next 
 
 MsgBox "Your publication contains " &; intShapes &; " shape(s)." 
End Sub
```


