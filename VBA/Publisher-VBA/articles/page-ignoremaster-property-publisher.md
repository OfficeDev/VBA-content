---
title: "Свойство Page.IgnoreMaster (издатель)"
keywords: vbapb10.chm393233
f1_keywords: vbapb10.chm393233
ms.prod: publisher
api_name: Publisher.Page.IgnoreMaster
ms.assetid: 53cd7b4b-4164-c6d3-766f-885a056d9b2b
ms.date: 06/08/2017
ms.openlocfilehash: a3f713347b6da5015a237be0057c35c1a3c3ca3b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pageignoremaster-property-publisher"></a>Свойство Page.IgnoreMaster (издатель)

 **Значение true** для Microsoft Publisher для форматирования для указанной странице главную страницу. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **IgnoreMaster**

 переменная _expression_A, представляющий объект **страницы** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере добавляется красной звездочкой в верхнем левом углу главной страницы, чтобы он отображается на каждой странице; затем добавляет несколько новых страниц и задает один из страниц, чтобы не использовать главную страницу, чтобы фигуры не отображается над ним.


```vb
Sub AddNewPageIgnoreMaster() 
 Dim pgNew As Page 
 
 With ActiveDocument 
 .MasterPages(1).Shapes.AddShape(Type:=msoShape5pointStar, _ 
 Left:=50, Top:=50, Width:=50, Height:=50).Fill.ForeColor _ 
 .CMYK.SetCMYK Cyan:=0, Magenta:=255, Yellow:=255, Black:=0 
 .Pages.Add Count:=1, After:=1 
 Set pgNew = .Pages.Add(Count:=1, After:=1) 
 pgNew.IgnoreMaster = True 
 End With 
End Sub
```


