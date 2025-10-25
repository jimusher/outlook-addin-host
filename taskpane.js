(function () {
  // Ensure code only runs after Office has initialized
  Office.onReady(function (info) {
    // If this page is opened directly in a browser (not inside Outlook), show a friendly message and exit
    if (!info || info.host !== Office.HostType.Outlook) {
      var linksDiv = document.getElementById('links');
      if (linksDiv) {
        linksDiv.textContent = 'This page is open outside Outlook. Use the ribbon button in Outlook to run the add-in.';
        linksDiv.classList.add('muted');
      }
      return;
    }

    // ---- DOM references (guarded) ----
    var linksDiv = document.getElementById('links');
    var copyBtn = document.getElementById('copyAll');
    var decodeToggle = document.getElementById('decodeSafeLinks');

    // Defensive: if critical container is missing, bail gracefully
    if (!linksDiv) {
      console.warn('[Extract URLs] #links element not found in taskpane.html');
      return;
    }

    // ---- Helpers ----
    function prettyUrl(u) {
      // Optionally decode Microsoft Defender Safe Links
      if (!decodeToggle || !decodeToggle.checked) return u;
      try {
        var url = new URL(u);
        var host = url.hostname || '';
        if (host.indexOf('safelinks.protection.outlook.com') >= 0 || host.indexOf('safelinks.office.com') >= 0) {
          var orig = url.searchParams.get('url');
          if (orig) return decodeURIComponent(orig);
        }
      } catch (e) { /* ignore parse errors */ }
      return u;
    }

    function render(urls) {
      if (!urls || !urls.size) {
        linksDiv.textContent = 'No URLs found in this message.';
        linksDiv.classList.add('muted');
        return;
      }
      linksDiv.classList.remove('muted');
      // Render each URL on its own line as a clickable link
      linksDiv.innerHTML = Array.from(urls)
        .map(function (u) { return '<a href="' + u + '" target="_blank" rel="noreferrer noopener">' + u + '</a>'; })
        .join('');
    }

    function extract(html) {
      var urls = new Set();
      try {
        var parser = new DOMParser();
        var doc = parser.parseFromString(html || '', 'text/html');

        // 1) All anchor hrefs
        var anchors = doc.querySelectorAll('a[href]');
        for (var i = 0; i < anchors.length; i++) {
          var href = anchors[i].getAttribute('href');
          if (href) urls.add(prettyUrl(href));
        }

        // 2) Plain-text URLs in the body
        var txt = (doc.body && doc.body.textContent) ? doc.body.textContent : '';
        var matches = txt.match(/https?:\/\/[^\s"'<>]+/gi) || [];
        for (var j = 0; j < matches.length; j++) {
          urls.add(prettyUrl(matches[j]));
        }
      } catch (e) {
        console.error('[Extract URLs] Failed to parse HTML', e);
      }
      return urls;
    }

    function refresh() {
      linksDiv.textContent = 'Reading message…';
      linksDiv.classList.add('muted');

      // Outlook host guarantee: Office.context.mailbox.item is available here
      var item = Office.context && Office.context.mailbox && Office.context.mailbox.item;
      if (!item || !item.body || !item.body.getAsync) {
        linksDiv.textContent = 'Outlook item not available yet. Please try again.';
        return;
      }

      item.body.getAsync(Office.CoercionType.Html, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          linksDiv.textContent = "Couldn't read the message body: " + (res.error && res.error.message ? res.error.message : res.status);
          return;
        }
        var urls = extract(res.value);
        render(urls);

        if (copyBtn) {
          copyBtn.onclick = function () {
            var all = Array.from(urls).join('\n');
            if (!all) return;
            // Try asynchronous clipboard first
            if (navigator.clipboard && navigator.clipboard.writeText) {
              navigator.clipboard.writeText(all)
                .then(function () { copyBtn.textContent = 'Copied!'; setTimeout(function () { copyBtn.textContent = 'Copy all'; }, 1200); })
                .catch(fallbackCopy);
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
              } catch (e) {
                console.warn('[Extract URLs] Clipboard copy failed', e);
              }
            }
          };
        }
      });
    }

    if (decodeToggle) decodeToggle.addEventListener('change', refresh);

    // Initial load
    refresh();
  });
})();
