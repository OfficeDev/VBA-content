---
title: "Свойство Document.LayoutGuides (издатель)"
keywords: vbapb10.chm196626
f1_keywords: vbapb10.chm196626
ms.prod: publisher
api_name: Publisher.Document.LayoutGuides
ms.assetid: 0c45366d-6b7a-7cf3-a566-bb945ff32ba4
ms.date: 06/08/2017
ms.openlocfilehash: 4efc546b5a7ec64446a4630e7c77844c617fa1b0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentlayoutguides-property-publisher"></a>Свойство Document.LayoutGuides (издатель)

Возвращает объект **[LayoutGuides](layoutguides-object-publisher.md)** , состоящий из полей и сетки направляющие разметки для всех страниц, включая главные страницы в публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LayoutGuides**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В следующем примере изменяется направляющие сетки, чтобы существует три столбца и пять строк.


```vb
Dim layTemp As LayoutGuides 
 
Set layTemp = ActiveDocument.LayoutGuides 
 
With layTemp 
 .Rows = 5 
 .Columns = 3 
End With 

```


