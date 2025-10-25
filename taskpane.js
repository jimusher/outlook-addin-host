(function () {
  function log() {
    try { console.log.apply(console, ['[Extract URLs]'].concat([].slice.call(arguments))); } catch (e) {}
  }
  function warn() { try { console.warn.apply(console, ['[Extract URLs]'].concat([].slice.call(arguments))); } catch (e) {} }
  function err()  { try { console.error.apply(console, ['[Extract URLs]'].concat([].slice.call(arguments))); } catch (e) {} }

  // Ensure code only runs after Office has initialized
  Office.onReady(function (info) {
    log('Office.onReady fired. host =', info && info.host);

    // If opened directly in a browser (not inside Outlook), show message and exit
    if (!info || info.host !== Office.HostType.Outlook) {
      var outside = document.getElementById('links');
      if (outside) {
        outside.textContent = 'This page is open outside Outlook. Use the ribbon button in Outlook to run the add-in.';
        outside.classList.add('muted');
      }
      warn('Not running in Outlook host; exiting.');
      return;
    }

    // ---- DOM references (guarded) ----
    var linksDiv = document.getElementById('links');
    var copyBtn = document.getElementById('copyAll');
    var decodeToggle = document.getElementById('decodeSafeLinks');

    if (!linksDiv) {
      warn('#links not found in taskpane.html; exiting.');
      return;
    }

    // ---- Helpers ----
    function prettyUrl(u) {
      if (!decodeToggle || !decodeToggle.checked) return u;
      try {
        var url = new URL(u);
        var host = url.hostname || '';
        if (host.indexOf('safelinks.protection.outlook.com') >= 0 || host.indexOf('safelinks.office.com') >= 0) {
          var orig = url.searchParams.get('url');
          if (orig) return decodeURIComponent(orig);
        }
      } catch (e) { /* ignore */ }
      return u;
    }

    function render(urls) {
      if (!urls || !urls.size) {
        linksDiv.textContent = 'No URLs found in this message.';
        linksDiv.classList.add('muted');
        return;
      }
      linksDiv.classList.remove('muted');
      linksDiv.innerHTML = Array.from(urls)
        .map(function (u) { return '<a href="' + u + '" target="_blank" rel="noreferrer noopener">' + u + '</a>'; })
        .join('');
    }

    function extract(html) {
      var urls = new Set();
      try {
        var parser = new DOMParser();
        var doc = parser.parseFromString(html || '', 'text/html');

        var anchors = doc.querySelectorAll('a[href]');
        for (var i = 0; i < anchors.length; i++) {
          var href = anchors[i].getAttribute('href');
          if (href) urls.add(prettyUrl(href));
        }

        var txt = (doc.body && doc.body.textContent) ? doc.body.textContent : '';
        var matches = txt.match(/https?:\/\/[^\s"'<>]+/gi) || [];
        for (var j = 0; j < matches.length; j++) {
          urls.add(prettyUrl(matches[j]));
        }
      } catch (e) {
        err('Failed to parse HTML', e);
      }
      return urls;
    }

    function refresh() {
      linksDiv.textContent = 'Reading message…';
      linksDiv.classList.add('muted');

      var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
      log('Refresh called. Office.context.mailbox.item exists =', !!item);

      if (!item || !item.body || !item.body.getAsync) {
        linksDiv.textContent = 'Outlook item not available yet. Please try again.';
        warn('item/body/getAsync not available; exiting refresh.');
        return;
      }

      try {
        log('Calling item.body.getAsync(HTML)…');
        item.body.getAsync(Office.CoercionType.Html, function (res) {
          log('getAsync callback status =', res && res.status, 'error =', res && res.error);
          if (res.status !== Office.AsyncResultStatus.Succeeded) {
            linksDiv.textContent = "Couldn't read the message body: " + (res.error && res.error.message ? res.error.message : res.status);
            return;
          }
          var urls = extract(res.value);
          log('Extracted URL count =', urls.size);
          render(urls);

          if (copyBtn) {
            copyBtn.onclick = function () {
              var all = Array.from(urls).join('\n');
              if (!all) return;
              if (navigator.clipboard && navigator.clipboard.writeText) {
                navigator.clipboard.writeText(all)
                  .then(function () { copyBtn.textContent = 'Copied!'; setTimeout(function () { copyBtn.textContent = 'Copy all'; }, 1200); })
                  .catch(function (e) { warn('Async clipboard failed', e); fallbackCopy(); });
              } else {
                fallbackCopy();
              }
              function fallbackCopy() {
                try {
                  var ta = document.createElement('textarea');
                  ta.value = all; ta.style.position = 'fixed'; ta.style.opacity = '0';
                  document.body.appendChild(ta); ta.select();
                  document.execCommand('copy');
                  document.body.removeChild(ta);
                  copyBtn.textContent = 'Copied!';
                  setTimeout(function () { copyBtn.textContent = 'Copy all'; }, 1200);
                } catch (e) { warn('Fallback copy failed', e); }
              }
            };
          }
        });
      } catch (e) {
        err('getAsync threw', e);
      }
    }

    if (decodeToggle) decodeToggle.addEventListener('change', refresh);

    // Initial load
    refresh();
  });
})();