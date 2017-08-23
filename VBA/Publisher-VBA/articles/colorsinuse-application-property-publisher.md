---
title: "Свойство ColorsInUse.Application (издатель)"
keywords: vbapb10.chm2949121
f1_keywords: vbapb10.chm2949121
ms.prod: publisher
api_name: Publisher.ColorsInUse.Application
ms.assetid: 726ef603-5f79-ac5b-e399-cbfe85773ba2
ms.date: 06/08/2017
ms.openlocfilehash: 85fcc746822831d746c382f7001263dd795326d3
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorsinuseapplication-property-publisher"></a>Свойство ColorsInUse.Application (издатель)

При использовании без квалификатор объекта, данное свойство возвращает объект **[приложения](application-object-publisher.md)** , который представляет текущего экземпляра Publisher. Используется квалификатор объекта, данное свойство возвращает объект **приложения** , представляющего создателя указанный объект. При использовании с помощью объекта OLE-автоматизации возвращает объект приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Приложения**

 переменная _expression_A, представляет собой объект- **ColorsInUse** .


## <a name="example"></a>Пример

В этом примере отображаются сведения о версии и построения для Publisher.


```vb
With Application 
 MsgBox "Current Publisher: version " _ 
 &; .Version &; " build " &; .Build 
End With
```

В этом примере отображается имя приложения, создавшего каждого связанного объекта на странице один активный публикации.




```vb
Dim shpOle As Shape 
 
For Each shpOle In ActiveDocument.Pages(1).Shapes 
 If shpOle.Type = pbLinkedOLEObject Then 
 MsgBox shpOle.OLEFormat.Application.Name 
 End If 
Next
```


