---
title: "Свойство Hyperlinks.Count (издатель)"
keywords: vbapb10.chm6881283
f1_keywords: vbapb10.chm6881283
ms.prod: publisher
api_name: Publisher.Hyperlinks.Count
ms.assetid: 36747f3e-b365-11ca-9cbe-f6148f7da235
ms.date: 06/08/2017
ms.openlocfilehash: bc6b32849b16894c8c3a70508073e03f8dbca470
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinkscount-property-publisher"></a>Свойство Hyperlinks.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляющий объект **гиперссылки** .


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


