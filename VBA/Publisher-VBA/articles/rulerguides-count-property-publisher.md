---
title: "Свойство RulerGuides.Count (издатель)"
keywords: vbapb10.chm720899
f1_keywords: vbapb10.chm720899
ms.prod: publisher
api_name: Publisher.RulerGuides.Count
ms.assetid: 92a93b1a-80c1-7a41-cb94-ac0859a4a470
ms.date: 06/08/2017
ms.openlocfilehash: 5a2b2d4abbafe5e3ba0be78606ccff715830af7d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="rulerguidescount-property-publisher"></a>Свойство RulerGuides.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **RulerGuides** .


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


