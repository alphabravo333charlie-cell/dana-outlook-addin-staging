Office.onReady(function (info) {
  var status = document.getElementById('status');
  if (status) status.textContent = 'Dana Outlook POC — running (' + (info.host || 'Outlook') + ')';
});
