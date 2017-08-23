---
title: "Свойство MailMergeFilters.Count (издатель)"
keywords: vbapb10.chm6750209
f1_keywords: vbapb10.chm6750209
ms.prod: publisher
api_name: Publisher.MailMergeFilters.Count
ms.assetid: 6ed658be-d3d0-ae5c-548d-ea724c9a8434
ms.date: 06/08/2017
ms.openlocfilehash: 2e0c0be505d80b1ed647f4b37e0a64f716f0258e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefilterscount-property-publisher"></a>Свойство MailMergeFilters.Count (издатель)

Возвращает значение типа **Long** , представляющее количество элементов в указанном семействе сайтов.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Count**

 переменная _expression_A, представляет собой объект- **MailMergeFilters** .


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


