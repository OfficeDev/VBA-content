---
title: "Свойство BorderArts.Count (издатель)"
keywords: vbapb10.chm7733251
f1_keywords: vbapb10.chm7733251
ms.prod: publisher
api_name: Publisher.BorderArts.Count
ms.assetid: 024cd14d-80f7-7372-c550-ef804661bbae
ms.date: 06/08/2017
ms.openlocfilehash: ace0ec3b02105a8080808635dfbde2c949dc63a3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="borderartscount-property-publisher"></a>Свойство BorderArts.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **BorderArts** .


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


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект BorderArts](borderarts-object-publisher.md)

