function classifyEmails() {
  const QUERY = 'is:unread newer_than:1d';
  const threads = GmailApp.search(QUERY);

  threads.forEach(thread => {
    try {
      const messages = thread.getMessages();
      if (messages.length === 0) return;

      // 🔹 Usamos SOLO el primer mail del thread (el que abre el caso)
      const message = messages[0];

      const subject = safeNormalize(message.getSubject());
      const body = safeNormalize(message.getPlainBody());
      const from = safeNormalize(message.getFrom());
      const attachments = message.getAttachments({ includeInlineImages: false }) || [];

      const labels = getLabelsByRules(subject, body, from, attachments);

      labels.forEach(labelName => {
        const label = GmailApp.getUserLabelByName(labelName) || GmailApp.createLabel(labelName);
        label.addToThread(thread);
      });

      //thread.markRead();

    } catch (e) {
      console.error('Error procesando thread:', e);
    }
  });
}
function getLabelsByRules(subject, body, from, attachments) {

  subject = safeString(subject);
  body = safeString(body);
  from = safeString(from);
  attachments = Array.isArray(attachments) ? attachments : [];

  const text = subject + ' ' + body;
  const labels = [];

  /* ============================
     SISTEMAS (multi)
  ============================ */

  if (containsAny(text, [
    'erp','factura','comprobante','nota','credito','debito',
    'cliente','proveedor','afip','impuesto','pago','cobro',
    'tarifa','precio','costo','liquidacion','contabilidad'
  ])) labels.push('ERP');

}

  /* ============================
     SUBTIPO: FACTURACIÓN
  ============================ */

  if (containsAny(text, [
    'factura','recibo','nota de credito','nota de debito',
    'comprobante','iva','monto','importe','liquidacion'
  ])) {
    labels.push('Facturación');
  }

  /* ============================
     INTENCIÓN (solo una)
     Orden de prioridad:
     Resuelto → Incidente → Requerimiento → Información
  ============================ */

  if (containsAny(text, [
    'resuelto','solucionado','ok','pueden cerrar','listo','confirmado'
  ])) {
    labels.push('Resuelto');


  } else if (containsAny(text, [
    'necesito','solicito','quiero','podrian','por favor','requiero',
    'cargar','crear','modificar','dar de alta','agregar','anular','cancelar'
  ])) {
    labels.push('Modificaciones');

  } else if (containsAny(text, [
    'consulta','duda','me confirman','me indican','me pasan',
    'necesito saber','estado de','cuando','donde','cual'
  ])) {
    labels.push('Informacion');
  }

  /* ============================
     DOCUMENTOS (PDF, Excel, Word, etc)
  ============================ */

  if (
    attachments.length > 0 &&
    attachments.some(a => {
      const type = (a.getContentType && a.getContentType().toLowerCase()) || '';
      const name = (a.getName && a.getName().toLowerCase()) || '';

      return (
        type.includes('image') ||
        type.includes('pdf') ||
        type.includes('spreadsheet') ||      // Google Sheets
        type.includes('excel') ||
        type.includes('word') ||
        type.includes('officedocument') ||   // .xlsx .docx
        type.includes('text') ||
        name.match(/\.(pdf|xls|xlsx|doc|docx|csv|txt)$/)
      );
    })
  ) {
    labels.push('Contiene Adjunto');
  }

  /* ============================
     FALLBACK
  ============================ */

  if (labels.length === 0) {
    labels.push('Soporte/General');
  }

  return [...new Set(labels)];
}
function safeString(value) {
  return typeof value === 'string' ? value : '';
}

function safeNormalize(value) {
  return safeString(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

function containsAny(text, words) {
  return words.some(w => text.includes(w));
}
