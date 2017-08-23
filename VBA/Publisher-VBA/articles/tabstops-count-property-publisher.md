---
title: "Свойство TabStops.Count (издатель)"
keywords: vbapb10.chm5570563
f1_keywords: vbapb10.chm5570563
ms.prod: publisher
api_name: Publisher.TabStops.Count
ms.assetid: 5ba876e2-b1c0-4de9-6942-02e6688aa169
ms.date: 06/08/2017
ms.openlocfilehash: feccc15c938eefb06c04fcdc42c656b4201e4149
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="tabstopscount-property-publisher"></a>Свойство TabStops.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **TabStops** .


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


