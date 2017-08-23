---
title: "Свойство ParagraphFormat.LockToBaseLine (издатель)"
keywords: vbapb10.chm5439540
f1_keywords: vbapb10.chm5439540
ms.prod: publisher
api_name: Publisher.ParagraphFormat.LockToBaseLine
ms.assetid: 4430bab6-a338-e61d-681c-6063d4a5c3b3
ms.date: 06/08/2017
ms.openlocfilehash: dc0b09dc1a681290b2e70c27a5e26078dd1d16fe
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatlocktobaseline-property-publisher"></a>Свойство ParagraphFormat.LockToBaseLine (издатель)

Возвращает **MsoTristate** , представляющий текст будет расположена по направляющие или нет. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **LockToBaseLine**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTristate


## <a name="remarks"></a>Заметки

Значение свойства **LockToBaseLine** может иметь одно из **MsoTriState** константы объявляются в библиотеке типов Microsoft Office и показаны в следующей таблице.



|**Константы**|**Описание**|
|:-----|:-----|
| **msoFalse**| Текст не выравнивается по исходных значений.|
| **msoTriStateMixed**|Указанный абзацы содержат текст, который выравнивается по исходных значений и текст, который не выравнивается по исходных значений.|
| **msoTrue**|Текст выравнивается по исходных значений.|

## <a name="example"></a>Пример

В следующем примере задается свойство **LockToBaseLine** значение **True**.


```vb
Dim objParaForm As ParagraphFormat 
Set objParaForm = ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.ParagraphFormat 
objParaForm.LockToBaseLine = msoTrue 

```


