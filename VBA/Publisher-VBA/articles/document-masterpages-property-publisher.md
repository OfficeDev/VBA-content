---
title: "Свойство Document.MasterPages (издатель)"
keywords: vbapb10.chm196629
f1_keywords: vbapb10.chm196629
ms.prod: publisher
api_name: Publisher.Document.MasterPages
ms.assetid: 26e5342b-94f0-4fd5-2743-92cfd2d43a01
ms.date: 06/08/2017
ms.openlocfilehash: f645844b6b6ea133d1fd1750aaf9f26af0a64e78
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentmasterpages-property-publisher"></a>Свойство Document.MasterPages (издатель)

Возвращает коллекцию **[макетом](masterpages-object-publisher.md)** для указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Макетом**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Макетом


## <a name="example"></a>Пример

В следующем примере задается текст в первой текстовой рамке на главную страницу на втором квартале.


```vb
Dim mp As MasterPages 
 
Set mp = ActiveDocument.MasterPages 
 
With mp.Item(1) 
 .Shapes(1).TextFrame.TextRange.Text = "Second Quarter" 
End With
```


