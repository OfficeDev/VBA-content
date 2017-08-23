---
title: "Свойство Hyperlink.Range (издатель)"
keywords: vbapb10.chm4587526
f1_keywords: vbapb10.chm4587526
ms.prod: publisher
api_name: Publisher.Hyperlink.Range
ms.assetid: ff105ffe-cb48-0f6a-99ff-eaac0500938f
ms.date: 06/08/2017
ms.openlocfilehash: b1b26b42ef426bc15fb3b88a7fdbf154a74d2a48
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="hyperlinkrange-property-publisher"></a>Свойство Hyperlink.Range (издатель)

Возвращает объект **[TextRange](textrange-object-publisher.md)** , представляющий основной текст, к которому был применен указанного гиперссылки.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Диапазон**

 переменная _expression_A, представляющий объект **гиперссылки** .


## <a name="remarks"></a>Заметки

Если свойство **Тип** указанного объекта **гиперссылки** имеет значение, отличное от **msoHyperlinkRange**, свойство **диапазон** возвращает значение nothing.


## <a name="example"></a>Пример

В следующем примере возвращает диапазон текста, связанный с первого гиперссылки на странице один активный публикации и изменяет основной текст для «См.»


```vb
Dim txtHyperlink As TextRange 
 
txtHyperlink = ActiveDocument.Pages(1) _ 
 .Shapes(1).Hyperlink.Range 
 
txtHyperlink.Text = "Go here"
```


