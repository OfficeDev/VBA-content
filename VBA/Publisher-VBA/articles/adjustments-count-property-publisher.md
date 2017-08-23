---
title: "Свойство Adjustments.Count (издатель)"
keywords: vbapb10.chm2424835
f1_keywords: vbapb10.chm2424835
ms.prod: publisher
api_name: Publisher.Adjustments.Count
ms.assetid: 1b32f1c3-0bbc-a175-4f59-36cc76df12fd
ms.date: 06/08/2017
ms.openlocfilehash: 3ef89bc5dde4500ed3b9116575fe9fcf6696c1e6
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="adjustmentscount-property-publisher"></a>Свойство Adjustments.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляющий объект **корректировки** .


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


 [Объект корректировки](adjustments-object-publisher.md)

