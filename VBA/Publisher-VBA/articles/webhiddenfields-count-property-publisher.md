---
title: "Свойство WebHiddenFields.Count (издатель)"
keywords: vbapb10.chm3997699
f1_keywords: vbapb10.chm3997699
ms.prod: publisher
api_name: Publisher.WebHiddenFields.Count
ms.assetid: 167c4c58-10cf-4dbb-5dfc-d60ab3856357
ms.date: 06/08/2017
ms.openlocfilehash: 9babbb65b175d7242a8534c65e56c504287e0f88
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webhiddenfieldscount-property-publisher"></a>Свойство WebHiddenFields.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **WebHiddenFields** .


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


