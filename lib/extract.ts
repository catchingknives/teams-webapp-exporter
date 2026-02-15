/**
 * Browser-side extraction logic for Teams chat messages.
 * Adapted from payload.js — runs inside page.evaluate() as a string.
 *
 * IMPORTANT: The helper strings use plain JS (no TS, no const/let, no arrow functions)
 * because tsx/esbuild injects __name helpers that break in browser evaluate context.
 */

export interface ExtractedMessage {
  id: number;
  author: string;
  timestamp: string; // ISO 8601
  content: string;
}

export interface ExtractionResult {
  messages: ExtractedMessage[];
  error?: string;
}

// ── Shared browser-side helpers (plain JS strings) ──────────────────

/** sleep, waitForChange, findScrollContainer, getOldestTimestamp */
const HELPERS_CORE = `
    var sleep = function(ms) { return new Promise(function(r) { setTimeout(r, ms); }); };

    var waitForChange = function(target, maxMs) {
      return new Promise(function(resolve) {
        var done = false;
        var finish = function() { if (!done) { done = true; obs.disconnect(); resolve(); } };
        var obs = new MutationObserver(function() { setTimeout(finish, 200); });
        obs.observe(target, { childList: true, subtree: true });
        setTimeout(finish, maxMs);
      });
    };

    var findScrollContainer = function(el) {
      var current = el;
      while (current && current !== document.documentElement) {
        var style = getComputedStyle(current);
        var ov = style.overflowY;
        if ((ov === 'auto' || ov === 'scroll') && current.scrollHeight > current.clientHeight) {
          return current;
        }
        current = current.parentElement;
      }
      current = el;
      while (current && current !== document.documentElement) {
        if (current.scrollTop > 0 && current.scrollHeight > current.clientHeight) {
          return current;
        }
        current = current.parentElement;
      }
      if (el.scrollHeight > el.clientHeight) return el;
      return null;
    };

    var getOldestTimestamp = function(nodes) {
      var oldest = null;
      for (var i = 0; i < nodes.length; i++) {
        var timeEl = nodes[i].querySelector('[id^="timestamp-"]');
        if (timeEl) {
          var dt = new Date(timeEl.getAttribute('datetime'));
          if (!oldest || dt < oldest) oldest = dt;
        }
      }
      return oldest;
    };
`;

/** Dedup collected nodes by message-body ID, filter GIFs, sort chronologically */
const HELPERS_DEDUP = `
    var dedupAndSort = function(collected, label) {
      if (label) console.log('[' + label + '] Deduplicating ' + collected.length + ' collected nodes...');
      var map = new Map();
      collected.forEach(function(n) {
        var msg = n.querySelector('[id^="message-body-"]');
        if (msg && !n.querySelector('[aria-label="Animated GIF"]')) {
          var id = parseInt(msg.id.replace('message-body-', ''), 10);
          if (!map.has(id)) map.set(id, n);
        }
      });
      var entries = Array.from(map.entries());
      entries.sort(function(a, b) { return a[0] - b[0]; });
      var nodes = entries.map(function(entry) { return entry[1]; });
      if (label) console.log('[' + label + '] ' + nodes.length + ' unique messages after dedup');
      return nodes;
    };
`;

/** DOM cleanup + structured message extraction from sorted nodes */
const HELPERS_EXTRACT = `
    var replaceEmojiImages = function(node) {
      node.querySelectorAll('img[itemtype*="Emoji"]').forEach(function(img) {
        var span = document.createElement('span');
        span.innerText = img.alt || '';
        img.parentNode.replaceChild(span, img);
      });
    };

    var replaceMentions = function(node) {
      try {
        node.querySelectorAll('div[aria-label*="Mention"]').forEach(function(div) {
          var span = document.createElement('span');
          while (div.firstChild) span.appendChild(div.firstChild);
          span.style.fontWeight = 'bold';
          div.parentNode.insertBefore(span, div);
          div.parentNode.removeChild(div);
        });
      } catch(e) {}
    };

    var replaceQuotedReplies = function(node) {
      try {
        node.querySelectorAll('div[data-track-module-name="messageQuotedReply"]').forEach(function(div) {
          var blockquote = document.createElement('blockquote');
          while (div.firstChild) blockquote.appendChild(div.firstChild);
          div.parentNode.insertBefore(blockquote, div);
          div.parentNode.removeChild(div);
        });
      } catch(e) {}
    };

    var extractMessages = function(nodes, label) {
      if (label) console.log('[' + label + '] Extracting text from ' + nodes.length + ' messages...');
      var messages = [];
      for (var i = 0; i < nodes.length; i++) {
        try {
          var n = nodes[i];
          var clone = n.cloneNode(true);
          replaceEmojiImages(clone);
          replaceMentions(clone);
          replaceQuotedReplies(clone);

          var authorEl = clone.querySelector('[data-tid="message-author-name"]');
          var timeEl = clone.querySelector('[id^="timestamp-"]');
          var bodyEl = clone.querySelector('[id^="message-body-"] [id^="content-"]');

          if (!authorEl || !timeEl || !bodyEl) continue;

          var msgBody = clone.querySelector('[id^="message-body-"]');
          var msgId = msgBody ? parseInt(msgBody.id.replace('message-body-', ''), 10) : 0;

          messages.push({
            id: msgId,
            author: authorEl.innerText.trim(),
            timestamp: timeEl.getAttribute('datetime'),
            content: bodyEl.innerText
          });
        } catch(e) { /* skip malformed message */ }
        if (label && i > 0 && i % 500 === 0) {
          console.log('[' + label + '] Processed ' + i + '/' + nodes.length + ' messages...');
        }
      }
      if (label) console.log('[' + label + '] Done — ' + messages.length + ' messages extracted');
      return messages;
    };
`;

/** All shared helpers combined */
const HELPERS = HELPERS_CORE + HELPERS_DEDUP + HELPERS_EXTRACT;

// ── Script generators ───────────────────────────────────────────────

/**
 * Returns a JS string IIFE that scrolls UP to load history.
 * days: -1 = no scroll, 0 = all history, >0 = N days back.
 */
export function getExtractionScript(days: number, sinceTimestamp?: string): string {
  return `(async function() {
    ${HELPERS}

    var list = document.getElementById('chat-pane-list');
    if (!list) {
      return { messages: [], error: 'No chat pane found. Make sure a chat is open in Teams.' };
    }

    var needsScroll = ${days} >= 0;
    var cutoffDate = ${days} > 0 ? new Date(Date.now() - ${days} * 86400000) : null;
    var sinceDate = ${sinceTimestamp ? `new Date("${sinceTimestamp}")` : 'null'};

    var seenIds = new Map();
    var collect = function() {
      var children = Array.from(list.children);
      for (var j = 0; j < children.length; j++) {
        var el = children[j];
        var mb = el.querySelector('[id^="message-body-"]');
        if (mb && !el.querySelector('[aria-label="Animated GIF"]')) {
          var mid = parseInt(mb.id.replace('message-body-', ''), 10);
          if (!seenIds.has(mid)) seenIds.set(mid, el);
        }
      }
    };
    collect();

    if (needsScroll) {
      var scrollContainer = findScrollContainer(list);
      if (scrollContainer) {
        scrollContainer.scrollTop = scrollContainer.scrollHeight;
        await waitForChange(list, 2000);
        collect();

        var obs = new MutationObserver(function() { collect(); });
        obs.observe(list, { childList: true, subtree: true, characterData: true });

        var noChangeCount = 0;
        var prevOldest = null;
        var scrollIter = 0;

        while (true) {
          scrollIter++;
          scrollContainer.scrollTop = 0;
          list.dispatchEvent(new KeyboardEvent('keydown', {
            key: 'Home', code: 'Home', bubbles: true, cancelable: true
          }));

          await waitForChange(list, 1500);
          collect();

          var snapshot = Array.from(seenIds.values());
          var effectiveCutoff = cutoffDate || sinceDate;
          if (effectiveCutoff) {
            var oldest = getOldestTimestamp(snapshot);
            if (oldest && oldest <= effectiveCutoff) {
              var reason = sinceDate && !cutoffDate ? 'last-exported timestamp' : 'cutoff date';
              console.log('[extract] Hit ' + reason + ' after ' + scrollIter + ' scrolls');
              break;
            }
          }

          var currentOldest = getOldestTimestamp(snapshot);
          var changed = !prevOldest || !currentOldest
            || currentOldest.getTime() !== prevOldest.getTime();
          if (!changed) {
            noChangeCount++;
            if (noChangeCount >= 3) {
              console.log('[extract] Scrolling done — no new messages after ' + scrollIter + ' iterations');
              break;
            }
          } else {
            noChangeCount = 0;
            if (scrollIter % 5 === 0) {
              console.log('[extract] Scroll ' + scrollIter + ', oldest so far: ' + (currentOldest ? currentOldest.toISOString() : 'unknown'));
            }
          }
          prevOldest = currentOldest;
        }

        obs.disconnect();
      }
    }

    var nodes = dedupAndSort(Array.from(seenIds.values()), 'extract');

    if (cutoffDate) {
      nodes = nodes.filter(function(n) {
        var timeEl = n.querySelector('[id^="timestamp-"]');
        if (!timeEl) return true;
        return new Date(timeEl.getAttribute('datetime')) >= cutoffDate;
      });
      console.log('[extract] ' + nodes.length + ' messages after date filter');
    }

    if (nodes.length === 0) {
      return { messages: [], error: 'No messages could be extracted.' };
    }

    return { messages: extractMessages(nodes, 'extract') };
  })()`;
}

/**
 * Returns a JS string IIFE that scrolls DOWN from the current position.
 * Use when the user has manually scrolled to the top of a chat.
 */
export function getScrollDownExtractionScript(): string {
  return `(async function() {
    ${HELPERS}

    var list = document.getElementById('chat-pane-list');
    if (!list) {
      return { messages: [], error: 'No chat pane found. Make sure a chat is open in Teams.' };
    }

    var seenIds = new Map();
    var collect = function() {
      var children = Array.from(list.children);
      for (var j = 0; j < children.length; j++) {
        var el = children[j];
        var mb = el.querySelector('[id^="message-body-"]');
        if (mb && !el.querySelector('[aria-label="Animated GIF"]')) {
          var mid = parseInt(mb.id.replace('message-body-', ''), 10);
          if (!seenIds.has(mid)) seenIds.set(mid, el);
        }
      }
    };
    collect();

    var scrollContainer = findScrollContainer(list);
    if (scrollContainer) {
      var obs = new MutationObserver(function() { collect(); });
      obs.observe(list, { childList: true, subtree: true, characterData: true });

      var noChangeCount = 0;
      var prevCount = 0;

      for (var step = 0; step < 5000; step++) {
        scrollContainer.scrollTop += Math.floor(scrollContainer.clientHeight * 0.6);
        scrollContainer.dispatchEvent(new KeyboardEvent('keydown', {
          key: 'PageDown', code: 'PageDown', bubbles: true, cancelable: true
        }));

        await waitForChange(list, 2000);
        collect();

        var currentCount = seenIds.size;

        if (step % 10 === 0) {
          console.log('[scroll-down] step ' + step + ', unique messages: ' + currentCount);
        }

        if (currentCount === prevCount) {
          noChangeCount++;
          if (noChangeCount === 4) {
            scrollContainer.scrollTop += scrollContainer.clientHeight * 2;
            await waitForChange(list, 3000);
            collect();
          }
          if (noChangeCount >= 8) break;
        } else {
          noChangeCount = 0;
        }
        prevCount = currentCount;

        if (scrollContainer.scrollTop + scrollContainer.clientHeight >= scrollContainer.scrollHeight - 10) {
          await waitForChange(list, 2000);
          collect();
          noChangeCount++;
          if (noChangeCount >= 5) break;
        }
      }

      obs.disconnect();
    }

    var nodes = dedupAndSort(Array.from(seenIds.values()), 'scroll-down');

    if (nodes.length === 0) {
      return { messages: [], error: 'No messages could be extracted.' };
    }

    return { messages: extractMessages(nodes, 'scroll-down') };
  })()`;
}
