---
title: "Свойство ObjectVerbs.Count (издатель)"
keywords: vbapb10.chm4521987
f1_keywords: vbapb10.chm4521987
ms.prod: publisher
api_name: Publisher.ObjectVerbs.Count
ms.assetid: 0d868be0-f46d-d8bb-2af1-47e2d1a3a388
ms.date: 06/08/2017
ms.openlocfilehash: 2dbcfbc336346cb82b31a86306d6d966aaf83037
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="objectverbscount-property-publisher"></a>Свойство ObjectVerbs.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляющий объект **ObjectVerbs** .


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


