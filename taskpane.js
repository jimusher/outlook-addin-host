(function(){
  Office.onReady(function(){
    const linksDiv = document.getElementById('links');
    const copyBtn = document.getElementById('copyAll');
    const decodeToggle = document.getElementById('decodeSafeLinks');

    function prettyUrl(u){
      if (!decodeToggle.checked) return u;
      try {
        const url = new URL(u);
        const host = url.hostname || '';
        if (host.includes('safelinks.protection.outlook.com') || host.includes('safelinks.office.com')){
          const orig = url.searchParams.get('url');
          if (orig) return decodeURIComponent(orig);
        }
      } catch(e) {}
      return u;
    }

    function render(urls){
      if (!urls.size){
        linksDiv.textContent = 'No URLs found in this message.';
        linksDiv.classList.add('muted');
        return;
      }
      linksDiv.classList.remove('muted');
      linksDiv.innerHTML = Array.from(urls).map(u => `<a href="${u}" target="_blank" rel="noreferrer noopener">${u}</a>`).join('');
    }

    function extract(html){
      const urls = new Set();
      try {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html || '', 'text/html');
        doc.querySelectorAll('a[href]').forEach(a => {
          const href = a.getAttribute('href');
          if (href) urls.add(prettyUrl(href));
        });
        const txt = doc.body ? (doc.body.textContent || '') : '';
        const matches = txt.match(/https?:\/\/[^\s"'<>]+/gi) || [];
        matches.forEach(m => urls.add(prettyUrl(m)));
      } catch(e){}
      return urls;
    }

    function refresh(){
      linksDiv.textContent = 'Reading message…';
      linksDiv.classList.add('muted');
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, function(res){
        if (res.status !== Office.AsyncResultStatus.Succeeded){
          linksDiv.textContent = "Couldn't read the message body: " + res.error.message;
          return;
        }
        const urls = extract(res.value);
        render(urls);
        copyBtn.onclick = function(){
          const all = Array.from(urls).join('\n');
          navigator.clipboard.writeText(all).then(() => {
            copyBtn.textContent = 'Copied!';
            setTimeout(()=> copyBtn.textContent = 'Copy all', 1200);
          }).catch(()=>{
            // Fallback
            const ta = document.createElement('textarea');
            ta.value = all; document.body.appendChild(ta); ta.select();
            try { document.execCommand('copy'); } catch(e){}
            document.body.removeChild(ta);
          });
        };
      });
    }

    decodeToggle.addEventListener('change', refresh);
    refresh();
  });
})();
