---
title: "Свойство ParagraphFormat.StartInNextTextBox (издатель)"
keywords: vbapb10.chm5439539
f1_keywords: vbapb10.chm5439539
ms.prod: publisher
api_name: Publisher.ParagraphFormat.StartInNextTextBox
ms.assetid: 96b34fa8-04ef-e472-16f0-15f82e7912ba
ms.date: 06/08/2017
ms.openlocfilehash: 34130a0d081cdf822da9e41736ce8e4dcb3673ba
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="paragraphformatstartinnexttextbox-property-publisher"></a>Свойство ParagraphFormat.StartInNextTextBox (издатель)

Возвращает или задает константой **MsoTriState** , представляющий необходимость всегда запускать выделенного абзаца в следующей связанной надписи. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **StartInNextTextBox**

 переменная _expression_A, представляет собой объект- **ParagraphFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoTriState


## <a name="remarks"></a>Заметки

Значение свойства **StartInNextTextBox** может иметь одно из ** [MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.

Если текст добавляется в предыдущем текстовое поле, что приводит к переполнению в текстовое поле, содержащий указанный текст, указанный текст (и любой текст, следующий его) перемещаются в верхней части следующего доступные текстовое поле. Если нет связанного текстового поля, указанный текст (и любой текст, следующий его) помещаются в буфер переполнение текста. Она останется в буфере до другого связанного текстового поля добавляется к публикации или изменении свойства **StartInNextTextBox** .

Это свойство соответствует управления **Пуск в следующем текстовом поле** в диалоговом окне **Абзац** .


