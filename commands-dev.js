(function () {
  'use strict';

  var ENDPOINT = 'https://dana-dev-api.ar-tech.cloud/dana/draft-reply';
  var NOTIFICATION_ID = 'dana-drafting';
  var PLACEHOLDER = /short context-specific instruction|insert your name here|\[?human[_ -]?input:?\s*\.\.\.?\]?/i;

  function textFromHtml(html) {
    var doc = new DOMParser().parseFromString(html || '', 'text/html');
    return (doc.body && doc.body.textContent || '').replace(/\u00a0/g, ' ').replace(/\r\n?/g, '\n').trim();
  }

  function escapeHtml(value) {
    return String(value || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  function addProgress(item) {
    return new Promise(function (resolve) {
      if (!item.notificationMessages || typeof item.notificationMessages.addAsync !== 'function') return resolve();
      item.notificationMessages.addAsync(NOTIFICATION_ID, {
        type: Office.MailboxEnums.ItemNotificationMessageType.ProgressIndicator,
        message: 'Dana is drafting your reply…'
      }, function () { resolve(); });
    });
  }

  function removeProgress(item) {
    if (!item.notificationMessages || typeof item.notificationMessages.removeAsync !== 'function') return;
    item.notificationMessages.removeAsync(NOTIFICATION_ID, function () {});
  }

  function readBody(item) {
    return new Promise(function (resolve, reject) {
      item.body.getAsync(Office.CoercionType.Html, function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return reject(result.error);
        resolve(textFromHtml(result.value));
      });
    });
  }

  function contextFor(item, body) {
    var from = item.from || {};
    return {
      messageId: item.itemId || '',
      conversationId: item.conversationId || '',
      sender: {
        displayName: from.displayName || '',
        emailAddress: from.emailAddress || '',
        name: from.displayName || '',
        address: from.emailAddress || ''
      },
      subject: item.subject || '',
      body: body || '',
      to: Array.isArray(item.to) ? item.to : [],
      cc: Array.isArray(item.cc) ? item.cc : [],
      replyMode: 'replyAll'
    };
  }

  function notes(proposal) {
    var raw = Array.isArray(proposal.reviewNotes) ? proposal.reviewNotes : [];
    if (!raw.length && proposal.needsHumanInput === true && proposal.uncertaintyReason) raw = [proposal.uncertaintyReason];
    return raw.map(function (note) { return String(note || '').trim(); })
      .filter(function (note) { return note && !PLACEHOLDER.test(note); });
  }

  function bodyHtml(proposal) {
    var reply = String(proposal.proposedReply || '').replace(/\[\[HUMAN_INPUT:[\s\S]*?\]\]/gi, '').trim();
    var review = notes(proposal);
    var html = '<div>' + escapeHtml(reply).replace(/\n/g, '<br>') + '</div>';
    if (proposal.needsHumanInput === true && review.length) {
      var isHebrew = /[\u0590-\u05ff]/.test(review.join(' '));
      var heading = isHebrew ? 'לבדיקה לפני שליחה:' : 'Review before sending:';
      html += '<div style="margin-top:14px;padding-top:8px;border-top:1px solid #d9d9d9;color:#c00000;font-weight:600">' +
        heading + '<br>' + review.map(function (note) { return '• ' + escapeHtml(note); }).join('<br>') +
        '</div>';
    }
    return html;
  }

  function openReply(item, proposal) {
    if (typeof item.displayReplyAllFormAsync !== 'function') throw new Error('Native Outlook Reply is unavailable in this client.');
    return new Promise(function (resolve, reject) {
      item.displayReplyAllFormAsync(bodyHtml(proposal), function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) return reject(result.error || new Error('Native reply could not be opened.'));
        resolve();
      });
    });
  }

  function draftReply(event) {
    var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
    if (!item) { event.completed(); return; }
    addProgress(item)
      .then(function () { return readBody(item); })
      .then(function (body) {
        var context = contextFor(item, body);
        return fetch(ENDPOINT, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(context)
        });
      })
      .then(function (response) { if (!response.ok) throw new Error('Dana could not draft this reply.'); return response.json(); })
      .then(function (proposal) { if (!proposal || typeof proposal.proposedReply !== 'string' || !proposal.proposedReply) throw new Error('Dana returned an incomplete draft.'); return openReply(item, proposal); })
      .then(function () { removeProgress(item); event.completed(); })
      .catch(function () { removeProgress(item); event.completed(); });
  }

  Office.onReady(function () { Office.actions.associate('draftReply', draftReply); });
}());
