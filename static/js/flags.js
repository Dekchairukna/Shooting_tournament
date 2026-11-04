/* ธงชาติหน้าชื่อสังกัด (รายการนานาชาติ)
 * อ่านข้อความในช่องสังกัด แล้วเติมรูปธงไว้ข้างหน้า ถ้าข้อความตรงกับประเทศในตาราง nations.py
 * ใช้กับ .aff-flag, .cell-affiliation (Overview), .team-cell (Bracket) และอัปเดตเองเมื่อหน้ารีเฟรชข้อมูลสด
 */
(function () {
  const INDEX = window.NATION_FLAGS || {};
  const BASE = window.NATION_FLAG_BASE || '/static/flags/';
  if (!Object.keys(INDEX).length) return;
  const SELECTOR = '.aff-flag, .cell-affiliation, .team-cell';

  function norm(text) {
    const t = String(text || '').normalize('NFKD').replace(/[̀-ͯ]/g, '').toLowerCase();
    const ascii = t.replace(/[^\x00-\x7f]/g, '').replace(/[^a-z0-9]/g, '');
    return ascii;
  }
  function lookup(text) {
    const raw = String(text || '').trim();
    if (!raw) return null;
    const code = raw.match(/\(([A-Za-z]{3})\)\s*$/);
    if (code && INDEX[code[1].toLowerCase()]) return INDEX[code[1].toLowerCase()];
    const lower = raw.toLowerCase();
    if (INDEX[lower]) return INDEX[lower];
    const n = norm(raw);
    if (n && INDEX[n]) return INDEX[n];
    return null;
  }
  function textOf(el) {
    let s = '';
    el.childNodes.forEach(node => {
      if (node.nodeType === 3) s += node.nodeValue;
      else if (node.nodeType === 1 && !node.classList.contains('nflag')) s += node.textContent;
    });
    return s.trim();
  }
  function decorate(el) {
    const text = textOf(el);
    const old = el.querySelector(':scope > img.nflag');
    if (el.dataset.flagFor === text && (old || el.dataset.flagNone === '1')) return;
    el.dataset.flagFor = text;
    const file = lookup(text);
    if (!file) { el.dataset.flagNone = '1'; if (old) old.remove(); return; }
    el.dataset.flagNone = '';
    const src = BASE + file + '.svg';
    if (old) { if (old.getAttribute('src') !== src) old.setAttribute('src', src); return; }
    const img = document.createElement('img');
    img.className = 'nflag';
    img.src = src;
    img.alt = '';
    img.setAttribute('aria-hidden', 'true');
    el.insertBefore(img, el.firstChild);
  }
  function scan(root) {
    if (root.nodeType !== 1) return;
    if (root.matches && root.matches(SELECTOR)) decorate(root);
    root.querySelectorAll && root.querySelectorAll(SELECTOR).forEach(decorate);
  }
  const style = document.createElement('style');
  style.textContent = 'img.nflag{display:inline-block;width:1.45em;height:1.09em;object-fit:cover;vertical-align:-.18em;margin-right:.4em;border-radius:2px;box-shadow:0 0 0 1px rgba(0,0,0,.18);}';
  document.head.appendChild(style);
  const start = () => {
    scan(document.body);
    let pending = false;
    new MutationObserver(muts => {
      if (pending) return;
      pending = true;
      requestAnimationFrame(() => {
        pending = false;
        const targets = new Set();
        muts.forEach(m => {
          let n = m.target.nodeType === 3 ? m.target.parentElement : m.target;
          const host = n && n.closest ? n.closest(SELECTOR) : null;
          if (host) targets.add(host);
          m.addedNodes.forEach(a => { if (a.nodeType === 1) targets.add(a); });
        });
        targets.forEach(scan);
      });
    }).observe(document.body, { childList: true, subtree: true, characterData: true });
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start); else start();
})();
