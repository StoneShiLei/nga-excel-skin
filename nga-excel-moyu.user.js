// ==UserScript==
// @name         NGA Excel 摸鱼皮肤
// @namespace    nga-excel-moyu
// @version      1.10.23
// @charset      UTF-8
// @description  把 NGA 伪装成 CW3 联调 Excel。Alt+Q 老板键切到接口核对，F10 显示/恢复原版。
// @author       moyu
// @match        https://ngabbs.com/*
// @match        https://*.ngabbs.com/*
// @match        https://bbs.nga.cn/*
// @match        https://nga.cn/*
// @match        https://www.nga.cn/*
// @match        https://*.nga.cn/*
// @match        https://nga.178.com/*
// @match        https://bbs.ngacn.cc/*
// @exclude      https://img.nga.cn/*
// @exclude      https://img*.nga.cn/*
// @exclude      https://pic.nga.cn/*
// @exclude      https://*.nga.cn/attachments/*
// @exclude      https://img.nga.178.com/*
// @exclude      https://img*.nga.178.com/*
// @exclude      https://img.ngacn.cc/*
// @noframes
// @run-at       document-start
// @grant        GM_xmlhttpRequest
// @grant        unsafeWindow
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_openInTab
// @connect      *
// ==/UserScript==

(function () {
  'use strict';
  if (window.top !== window) return;
  var _host = (location.hostname || '').toLowerCase();
  var _path = (location.pathname || '').toLowerCase();
  if (/^(img\d*|pic|att)\./.test(_host)) return;
  if (/\/attachments\//.test(_path) && /\.(jpe?g|png|gif|webp|bmp|webm|mp4)(\/|$)/i.test(_path)) return;
  if (/\.(jpe?g|png|gif|webp|bmp)$/i.test(_path)) return;

  function isAdPath(path) {
    return /\/misc\/adpage/i.test(path || '');
  }
  function unwrapAdHref(href) {
    var raw = String(href || '');
    if (!raw) return '';
    try {
      var u = new URL(raw, location.href);
      if (!isAdPath(u.pathname)) return u.href;
      var q = u.search || '';
      if (/^\?(nojump|test)$/i.test(q)) return '';
      q = q.replace(/^\?\d*/, '');
      try { q = decodeURIComponent(q); } catch (e1) {}
      q = String(q || '').replace(/^(?:url|u|to|target|jump)=/i, '');
      if (!/^https?:\/\//i.test(q) && /^(?:\/)?(?:thread|read|post|nuke)\.php/i.test(q)) {
        q = u.origin + (q.charAt(0) === '/' ? q : '/' + q);
      }
      if (!/^https?:\/\/([a-z0-9-]+\.)?(nga\.cn|nga\.donews\.com|ngacn\.cc|178\.com|ngabbs\.com|bigccq\.cn)(:\d+)?([/?#]|$)/i.test(q)) return '';
      return q;
    } catch (e2) { return ''; }
  }
  function looksLikeAdHtml(html) {
    return /function\s+getJump\s*\(|adpage_insert|\u70b9\u6b64\u8df3\u8fc7\u5e7f\u544a/.test(String(html || ''));
  }
  if (isAdPath(_path)) {
    var _adDest = unwrapAdHref(location.href);
    if (_adDest) {
      try {
        var _st = document.createElement('style');
        _st.textContent = 'html,body{opacity:0!important;background:#fff!important}';
        (document.documentElement || document.head).appendChild(_st);
      } catch (eSt) {}
      try { location.replace(_adDest); } catch (eAd) { location.href = _adDest; }
    }
    return;
  }

  var FILENAME = 'CW3_Sprint_联调清单_202609.xlsx';
  var HOTKEY_HINT_MS = 5000;
  var COLS = 10;
  var MIN_ROWS = 50;
  var LS_MODE = 'nga-xl-mode';
  var LS_COLS = 'nga-xl-cols';
  var LS_PANE = 'nga-xl-pane';
  var LS_COLORD = 'nga-xl-colord';
  var LS_DRAFT = 'nga-xl-draft';
  var LS_SUB = 'nga-xl-subpick';

  var CSS = '';
  CSS += ':host{font-family:"Segoe UI","Microsoft YaHei",DengXian,sans-serif;color:#252423;display:block;width:100%;height:100%;}';
  CSS += '*{box-sizing:border-box;}';
  CSS += '.app{display:flex;flex-direction:column;width:100%;height:100%;background:#fff;user-select:none;position:relative;}';
  CSS += '.appbar{height:36px;display:flex;align-items:center;gap:8px;padding:0 10px;background:#fff;border-bottom:1px solid #e1dfdd;}';
  CSS += '.icon{width:20px;height:20px;border-radius:3px;background:#185c37;color:#fff;font:700 12px/20px Arial;text-align:center;flex:none;}';
  CSS += '.filebtn{font:12px "Microsoft YaHei";padding:2px 8px;border-radius:3px;}';
  CSS += '.filebtn:hover,.filebtn.on{background:#f3f2f1;}';
  CSS += '.filemenu{position:absolute;top:36px;left:28px;z-index:50;width:200px;background:#fff;border:1px solid #c8c6c4;box-shadow:0 6px 18px rgba(0,0,0,.18);padding:4px 0;font:12px "Microsoft YaHei";}';
  CSS += '.filemenu.off{display:none;} .filemenu button{display:block;width:100%;text-align:left;border:0;background:transparent;padding:6px 14px;font:12px "Microsoft YaHei";color:#252423;cursor:pointer;} .filemenu button:hover{background:#c5e0b4;} .filemenu .sep{height:1px;background:#e1dfdd;margin:4px 0;}';
  CSS += '.fname{margin-left:6px;font:600 12px/22px "Microsoft YaHei";padding:0 10px;border-radius:4px;border:1px solid transparent;}';
  CSS += '.fname:hover{background:#f3f2f1;border-color:#e1dfdd;}';
  CSS += '.appbar .sp{flex:1;}';
  CSS += '.search{width:220px;height:24px;border:1px solid #d2d0ce;border-radius:4px;background:#f3f2f1;font:12px "Microsoft YaHei";padding:0 8px;color:#605e5c;}';
  CSS += '.share{margin-left:8px;height:24px;padding:0 12px;border:none;border-radius:4px;background:#185c37;color:#fff;font:12px "Microsoft YaHei";}';
  CSS += '.rtabs{height:28px;display:flex;align-items:flex-end;padding:0 8px;gap:2px;background:#fff;}';
  CSS += '.rtab{font:12px "Microsoft YaHei";padding:4px 12px;border-radius:4px 4px 0 0;color:#252423;cursor:default;}';
  CSS += '.rtab.on{color:#185c37;font-weight:600;border-bottom:3px solid #185c37;padding-bottom:3px;}';
  CSS += '.ribbon{height:36px;background:#f3f2f1;border-bottom:1px solid #e1dfdd;display:flex;align-items:center;padding:0 8px;gap:0;overflow:hidden;white-space:nowrap;flex:none;}';
  CSS += '.rg{display:flex;flex-direction:row;align-items:center;padding:0 10px;border-right:1px solid #e1dfdd;gap:6px;height:24px;}';
  CSS += '.rg .row{display:flex;align-items:center;gap:4px;}';
  CSS += '.rg .cap{font:11px "Microsoft YaHei";color:#605e5c;}';
  CSS += '.btn{height:22px;min-width:22px;padding:0 6px;border:1px solid transparent;border-radius:3px;background:transparent;font:12px "Microsoft YaHei";color:#252423;}';
  CSS += '.btn:hover{background:#fff;border-color:#d2d0ce;}';
  CSS += '.btn.b{font-weight:700;} .btn.i{font-style:italic;} .btn.u{text-decoration:underline;}';
  CSS += '.sel{height:22px;border:1px solid #d2d0ce;background:#fff;font:12px "Segoe UI";border-radius:2px;}';
  CSS += '.fontsel{width:92px;} .sizesel{width:40px;}';
  CSS += '.sw{width:16px;height:16px;border:1px solid #c8c6c4;display:inline-block;vertical-align:middle;}';
  CSS += '.formula{height:24px;display:flex;align-items:stretch;border-bottom:1px solid #d4d4d4;background:#fff;}';
  CSS += '.namebox{width:72px;border-right:1px solid #d4d4d4;font:12px Consolas,"Segoe UI";display:flex;align-items:center;justify-content:center;color:#252423;}';
  CSS += '.fx{width:28px;font:italic 13px "Times New Roman";color:#217346;display:flex;align-items:center;justify-content:center;border-right:1px solid #d4d4d4;}';
  CSS += '.fxinput{flex:1;font:12px Calibri,DengXian,"Microsoft YaHei";padding:0 8px;display:flex;align-items:center;overflow:hidden;white-space:nowrap;user-select:text;}';
  CSS += '.main{flex:1;display:flex;min-height:0;background:#fff;}';
  CSS += '.gridwrap{flex:1;min-width:0;overflow:auto;position:relative;background:#fff;}';
  CSS += 'table.grid{border-collapse:collapse;table-layout:fixed;font:12px Calibri,DengXian,"Microsoft YaHei";}';
  CSS += 'table.grid th,table.grid td{border:1px solid #d0d0d0;height:18px;padding:0 4px;overflow:hidden;min-width:0;max-width:0;text-overflow:ellipsis;white-space:nowrap;vertical-align:middle;}';
  CSS += 'table.grid th{background:#f8f8f8;font:11px "Segoe UI";color:#333;text-align:center;position:sticky;top:0;z-index:2;}';
  CSS += 'table.grid th.rh{width:36px;left:0;z-index:3;color:#666;font-weight:400;}';
  CSS += 'table.grid td.rh{width:36px;background:#f8f8f8;text-align:center;color:#666;position:sticky;left:0;z-index:1;font:11px "Segoe UI";padding:0;}';
  CSS += 'table.grid td.sel{outline:2px solid #217346;outline-offset:-2px;background:#e2f0d9 !important;position:relative;z-index:1;}';
  CSS += 'table.grid tr.selrow td:not(.rh){background:#e2f0d9;}';
  CSS += 'table.grid td.link{color:#0563c1;text-decoration:underline;cursor:pointer;}';
  CSS += 'table.grid td.num{text-align:right;font-variant-numeric:tabular-nums;}';
  CSS += 'table.grid td.pct{text-align:right;color:#1f4e79;}';
  CSS += 'table.grid td.neg{color:#c00000;}';
  CSS += 'table.grid td.head{background:#375623;color:#fff;font-weight:700;}';
  CSS += 'table.grid td.sec{background:#c6e0b4;font-weight:700;}';
  CSS += 'table.grid td.tot{background:#e2efda;font-weight:700;}';
  CSS += 'table.grid td.title{font-size:12px;font-weight:700;color:#375623;}';
  CSS += 'table.grid td.sub{color:#833c0c;font-style:italic;}';
  CSS += 'table.grid td.warn{background:#fff2cc;}';
  CSS += 'table.grid tr:nth-child(even) td:not(.rh):not(.head):not(.sec):not(.tot):not(.title):not(.warn){background:#fafafa;}';
  CSS += 'table.grid col{min-width:0;} .resz{position:absolute;right:0;top:0;width:8px;height:100%;cursor:col-resize;z-index:30;} .resz:hover{background:rgba(33,115,70,.45);} .gridwrap.colresize,.gridwrap.colresize *{cursor:col-resize!important;} table.grid th[data-col]{cursor:grab;} table.grid th[data-col].dragging{opacity:.4;} table.grid th[data-col].drop-before{box-shadow:inset 3px 0 0 #185c37;} table.grid th[data-col].drop-after{box-shadow:inset -3px 0 0 #185c37;} .gridwrap.colmove,.gridwrap.colmove *{cursor:grabbing!important;}';
  CSS += '.pane{width:400px;min-width:180px;flex:none;position:relative;border-left:1px solid #d4d4d4;display:flex;flex-direction:column;background:#fff;}';
  CSS += '.pane.hide{display:none;} .pane-resizer{position:absolute;left:-3px;top:0;bottom:0;width:6px;cursor:col-resize;z-index:6;} .pane-resizer:hover{background:rgba(33,115,70,.35);}';
  CSS += '.pane h3{margin:0;height:28px;display:flex;align-items:center;padding:0 10px;font:12px "Microsoft YaHei";background:#f3f2f1;border-bottom:1px solid #e1dfdd;color:#185c37;}';
  CSS += '.pane .body{flex:1;overflow:auto;padding:10px 12px;font:13px Calibri,DengXian,"Microsoft YaHei";line-height:1.55;user-select:text;white-space:normal;word-break:break-word;color:#252423;}';
  CSS += '.pane .body .qbox{border-left:3px solid #a19f9d;background:#f3f2f1;padding:6px 8px;margin:0 0 10px;color:#605e5c;}';
  CSS += '.pane .body .rbox{border-left:3px solid #217346;padding:0 0 0 8px;margin:0 0 10px;}';
  CSS += '.pane .body .qlab,.pane .body .rlab{font:11px "Microsoft YaHei";color:#217346;margin:0 0 4px;}';
  CSS += '.pane .body .qtxt,.pane .body .rtxt{white-space:pre-wrap;word-break:break-word;}';
  CSS += '.pane .body .imgls a{display:block;margin:2px 0;}';
  CSS += '.pane .body .votes{display:flex;gap:18px;margin:0 0 10px;padding:0 0 6px;border-bottom:1px solid #eee;font:11px Consolas,"Segoe UI";color:#605e5c;} .pane .body .votes b{font-weight:400;color:#a19f9d;margin-right:6px;}';
  CSS += '.imgmenu{position:fixed;z-index:80;width:220px;background:#fff;border:1px solid #c8c6c4;box-shadow:0 6px 18px rgba(0,0,0,.2);padding:10px;font:12px "Microsoft YaHei";}';
  CSS += '.imgmenu img.tiny{width:2cm;height:2cm;object-fit:cover;display:block;margin:0 auto 8px;border:1px solid #e1dfdd;background:#f3f2f1;}';
  CSS += '.imgmenu .url{font:11px Consolas,monospace;color:#605e5c;word-break:break-all;margin:0 0 8px;max-height:2.6em;overflow:hidden;}';
  CSS += '.imgmenu button{display:block;width:100%;text-align:left;margin:0 0 4px;padding:6px 8px;border:1px solid #d2d0ce;background:#f3f2f1;border-radius:3px;font:12px "Microsoft YaHei";cursor:pointer;}';
  CSS += '.imgmenu button:hover{background:#fff;border-color:#217346;}';
  CSS += 'table.grid td.wrap{white-space:normal;height:auto;padding:1px 4px;overflow:hidden;vertical-align:top;}';
  CSS += 'table.grid td.wrap .clamp{display:-webkit-box;-webkit-box-orient:vertical;-webkit-line-clamp:3;overflow:hidden;white-space:pre-wrap;word-break:break-word;line-height:18px;max-height:54px;}';
  CSS += 'table.grid tr.wraprow td{height:auto;vertical-align:top;}';
  CSS += '.pane .body a{color:#0563c1;text-decoration:underline;cursor:pointer;word-break:break-all;} .pane .body .imgls b{display:block;margin:8px 0 4px;font:12px "Microsoft YaHei";color:#185c37;}';
  CSS += '.pane .meta{padding:6px 12px;font:11px "Microsoft YaHei";color:#605e5c;border-bottom:1px solid #eee;}';
  CSS += '.compose{border-top:1px solid #d4d4d4;background:#fafafa;padding:8px 10px;flex:none;}';
  CSS += '.compose .chd{font:11px "Microsoft YaHei";color:#185c37;margin:0 0 6px;}';
  CSS += '.compose .ctitle{width:100%;height:22px;margin:0 0 6px;border:1px solid #d2d0ce;font:12px "Microsoft YaHei";padding:0 6px;display:block;}';
  CSS += '.compose .ctitle.off{display:none;}';
  CSS += '.compose .cbody{width:100%;height:72px;border:1px solid #d2d0ce;font:12px Calibri,"Microsoft YaHei";padding:6px;resize:vertical;min-height:56px;display:block;}';
  CSS += '.compose .crow{display:flex;align-items:center;gap:6px;margin-top:6px;}';
  CSS += '.compose .cstat{font:11px "Microsoft YaHei";color:#605e5c;}';
  CSS += '.compose .sp{flex:1;}';
  CSS += '.compose .send{background:#185c37;color:#fff;border:none;height:22px;padding:0 12px;border-radius:3px;}';
  CSS += '.compose input,.compose textarea{user-select:text;} .compose .cfile{width:0;height:0;opacity:0;position:absolute;}';
  CSS += '.compose [data-compose="quote"].off{display:none;}';
  CSS += '.sheets{height:30px;display:flex;align-items:stretch;background:#f3f2f1;border-top:1px solid #d0d0d0;}';
  CSS += '.stab{min-width:84px;padding:0 14px;display:flex;align-items:center;justify-content:center;font:12px "Microsoft YaHei";border-right:1px solid #e1dfdd;cursor:pointer;color:#252423;}';
  CSS += '.stab:hover{background:#fff;}';
  CSS += '.stab.on{background:#fff;border-top:2px solid #217346;font-weight:600;color:#217346;}';
  CSS += '.sstat{flex:1;display:flex;align-items:center;justify-content:flex-end;gap:16px;padding:0 12px;font:11px "Segoe UI","Microsoft YaHei";color:#fff;background:#217346;}';
  CSS += '.hint{position:absolute;right:16px;bottom:40px;background:#217346;color:#fff;font:12px "Microsoft YaHei";padding:8px 12px;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.2);z-index:9;pointer-events:none;opacity:.92;}';
  CSS += '.hint b{color:#fff2cc;}';
  CSS += '.navrow{height:28px;display:flex;align-items:center;gap:6px;padding:0 8px;background:#fafafa;border-bottom:1px solid #d4d4d4;font:12px "Microsoft YaHei";flex:none;position:relative;z-index:20;overflow:visible;}';
  CSS += '.navrow .lab{color:#605e5c;}';
  CSS += '.navrow select{height:22px;width:160px;max-width:180px;border:1px solid #d2d0ce;background:#fff;font:12px "Microsoft YaHei";}';
  CSS += '.navrow select.off,.navrow .lab.off,.navrow .subwrap.off{display:none;}';
  CSS += '.navrow .subwrap{position:relative;display:inline-flex;flex:none;z-index:40;}';
  CSS += '#xl-subbtn{appearance:none;-webkit-appearance:none;height:22px;min-width:88px;max-width:168px;padding:0 8px;border:1px solid #8a8886;border-radius:2px;background:#fff;color:#252423;font:12px "Microsoft YaHei";cursor:pointer;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}';
  CSS += '#xl-subbtn:hover{border-color:#217346;}';
  CSS += '.subpop{position:absolute;z-index:80;top:24px;left:0;width:220px;max-height:260px;overflow:auto;background:#fff;border:1px solid #c8c6c4;box-shadow:0 6px 16px rgba(0,0,0,.16);padding:4px 8px;}';
  CSS += '.subpop.off{display:none;} .subpop label{display:flex;gap:6px;align-items:center;padding:3px 0;cursor:pointer;}';
  CSS += 'table.grid td.pinhead{cursor:pointer;background:#e2efda;color:#185c37;font-weight:600;}';
  CSS += '.navrow .btn[disabled]{opacity:.4;}';
  CSS += '.navrow #xl-title{flex:1;min-width:40px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;color:#185c37;font-weight:600;}';
  CSS += 'table.grid td a.cella{color:inherit;text-decoration:inherit;display:block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:100%;}';
  CSS += '.pane h3{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}';
  CSS += 'table.grid td.findhit{background:#fff2cc !important;}';
  CSS += '.findbar{height:28px;display:flex;align-items:center;gap:6px;padding:0 8px;background:#fff8e1;border-bottom:1px solid #d4d4d4;font:12px Microsoft YaHei;flex:none;}';
  CSS += '.findbar.off{display:none;} .findbar input{height:22px;width:180px;border:1px solid #d2d0ce;padding:0 6px;font:12px Segoe UI;}';
  CSS += '.dlg{position:fixed;inset:0;background:rgba(0,0,0,.25);z-index:90;display:flex;align-items:center;justify-content:center;} .dlg.off{display:none;}';
  CSS += '.dlgbox{width:320px;background:#fff;border:1px solid #c8c6c4;box-shadow:0 8px 24px rgba(0,0,0,.2);padding:12px;} .dlgbox h4{margin:0 0 10px;font:13px Microsoft YaHei;color:#185c37;}';
  CSS += '.dlgf{display:flex;justify-content:flex-end;gap:6px;margin-top:10px;}';
  CSS += '.actrow{display:flex;flex-wrap:wrap;gap:4px;padding:6px 10px;border-bottom:1px solid #eee;}';
  CSS += '.pane .body .votes button{border:1px solid #d2d0ce;background:#f3f2f1;padding:2px 8px;font:11px Consolas,Segoe UI;cursor:pointer;} .pane .body .votes button:hover{border-color:#217346;background:#fff;}';
  CSS += '.cmtls{margin:10px 0 0;padding-top:8px;border-top:1px solid #eee;} .cmtls .chd{font:12px Microsoft YaHei;color:#185c37;margin:0 0 8px;font-weight:600;} .cmtls .cmt{border-left:3px solid #c8c6c4;background:#f7f7f5;padding:6px 8px;margin:0 0 8px;} .cmtls .cwho{font:11px Microsoft YaHei;color:#217346;margin:0 0 4px;} .cmtls .cwho span{color:#a19f9d;font-weight:400;margin-left:8px;} .cmtls .cmeta{font:11px Microsoft YaHei;color:#a19f9d;margin:0 0 4px;} .cmtls .ctxt{white-space:pre-wrap;word-break:break-word;font:12px Calibri,Microsoft YaHei;color:#252423;line-height:1.5;}';
  CSS += '.poll{margin:8px 0;padding-top:6px;border-top:1px solid #eee;} .poll b{display:block;margin:0 0 4px;font:12px Microsoft YaHei;color:#185c37;} .poll label{display:block;margin:3px 0;font:12px Microsoft YaHei;}';
  CSS += '.compose .vrow{display:flex;gap:6px;align-items:center;margin-top:6px;} .compose .vrow.off{display:none;} .compose .vimg{height:28px;cursor:pointer;border:1px solid #d2d0ce;} .compose .vcode{height:22px;width:90px;border:1px solid #d2d0ce;padding:0 6px;font:12px Segoe UI;}';
  CSS += '.navrow .mini{width:52px;height:22px;border:1px solid #d2d0ce;font:12px Segoe UI;padding:0 4px;}';
  CSS += '.search,.findbar input,.navrow input,.dlg input,.compose input,.compose textarea{user-select:text;}';
  CSS += '.search{color:#252423;}'

  var COL_LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  var state = {
    mode: localStorage.getItem(LS_MODE) || 'skin',
    sheet: 'pl',
    selR: 1,
    selC: 1,
    cells: [],
    meta: { hrefs: {}, preview: {} },
    nav: { boards: [], prev: '', next: '', back: '', page: 1, maxPage: 1, pages: [], collections: [] },
    isRead: false,
    hintOn: true,
    find: { hits: [], i: 0, q: '' },
    userMap: {},
    lastFid: ''
  };

  var host, shadow, root, bootStyle;

  function $(sel, el) { return (el || shadow).querySelector(sel); }
  function $$(sel, el) { return Array.prototype.slice.call((el || shadow).querySelectorAll(sel)); }

  function colName(i) { return COL_LETTERS.charAt(i - 1) || 'A'; }
  function cellAddr(r, c) { return colName(c) + r; }
  function isReadPage() {
    var p = (location.pathname || '').toLowerCase();
    return p === '/read.php' || /\/read\.php$/i.test(p);
  }
  function layoutKey() {
    var sh = (state && state.sheet) || 'tk';
    if (sh && sh !== 'tk') return sh;
    try { if (isReadPage()) return 'read'; } catch (eLk) {}
    return 'list';
  }
  var COL_DEFAULTS = {
    list: [32, 72, 360, 68, 44, 86, 52, 70, 70, 48, 48],
    read: [32, 44, 90, 96, 380, 56, 56, 48, 48, 48, 48],
    pl: [32, 80, 280, 72, 64, 72, 80, 64, 80, 48, 48],
    hc: [32, 72, 200, 80, 80, 80, 80, 80, 80, 48, 48],
    as: [32, 72, 280, 80, 80, 80, 80, 80, 48, 48, 48]
  };
  var COL_DEFAULT = COL_DEFAULTS.list;
  function defaultWidths(key) {
    return (COL_DEFAULTS[key] || COL_DEFAULTS.list).slice();
  }
  function ensureWidthArr(key, arr) {
    var def = defaultWidths(key);
    var out = (arr && arr.length) ? arr.slice() : def.slice();
    while (out.length < def.length) out.push(def[out.length] || 48);
    return out;
  }
  function pageLS() {
    try { if (typeof unsafeWindow !== 'undefined' && unsafeWindow.localStorage) return unsafeWindow.localStorage; } catch (e0) {}
    try { return localStorage; } catch (e1) { return null; }
  }
  function storeGet(key) {
    var out = [];
    try { if (typeof GM_getValue === 'function') out.push(GM_getValue(key, null)); } catch (eG) {}
    try {
      var ls = pageLS();
      if (ls) out.push(ls.getItem(key));
    } catch (eL) {}
    return out;
  }
  function storeSet(key, val) {
    var s = typeof val === 'string' ? val : JSON.stringify(val);
    try { if (typeof GM_setValue === 'function') GM_setValue(key, s); } catch (eG) {}
    try {
      var ls = pageLS();
      if (ls) ls.setItem(key, s);
    } catch (eL) {}
  }
  var subPickStore = null;
  function readSubStore() {
    if (subPickStore) return subPickStore;
    var cands = storeGet(LS_SUB);
    var best = {};
    var i, raw, o;
    for (i = 0; i < cands.length; i++) {
      raw = cands[i];
      if (raw == null || raw === '') continue;
      try {
        o = typeof raw === 'string' ? JSON.parse(raw) : raw;
        if (o && typeof o === 'object') best = o;
      } catch (eS) {}
    }
    subPickStore = best;
    return best;
  }
  function copyPick(src) {
    var out = {};
    var k;
    src = src || {};
    for (k in src) if (src.hasOwnProperty(k) && src[k]) out[k] = 1;
    return out;
  }
  function useSubPick(fid) {
    fid = fid || '';
    if (!fid || state.subParent === fid) return;
    state.subParent = fid;
    state.subPick = copyPick(readSubStore()[fid]);
  }
  function saveSubPick() {
    var fid = state.subParent || '';
    if (!fid) return;
    var map = readSubStore();
    var pick = copyPick(state.subPick);
    var n = 0;
    var k;
    for (k in pick) if (pick.hasOwnProperty(k)) n++;
    if (n) map[fid] = pick;
    else delete map[fid];
    subPickStore = map;
    storeSet(LS_SUB, map);
  }
  function parseColsRaw(raw) {
    if (raw == null || raw === '') return null;
    try {
      var a = typeof raw === 'string' ? JSON.parse(raw) : raw;
      if (!Array.isArray(a) || a.length < 3) return null;
      return a.map(function (x) { return parseInt(x, 10) || 48; });
    } catch (eP) { return null; }
  }
  function parseColsMap(raw) {
    if (raw == null || raw === '') return null;
    var v = raw;
    try { if (typeof raw === 'string') v = JSON.parse(raw); } catch (eM) { return null; }
    if (Array.isArray(v) && v.length >= 3) {
      var one = parseColsRaw(v);
      if (!one) return null;
      return { list: one.slice() };
    }
    if (!v || typeof v !== 'object') return null;
    var out = {}, k, one2;
    for (k in v) {
      if (!v.hasOwnProperty(k)) continue;
      one2 = parseColsRaw(v[k]);
      if (one2) out[k] = one2;
    }
    if (out.tk && !out.list) out.list = out.tk;
    return out;
  }
  function colsScore(a) {
    if (!Array.isArray(a) || a.length < 3) return -1;
    var s = 0, i, n;
    for (i = 0; i < a.length; i++) {
      n = Number(a[i]) || 0;
      if (n > 0) s += 1;
      if (i < COL_DEFAULT.length && n !== COL_DEFAULT[i]) s += 3;
    }
    return s;
  }
  function readColsMap() {
    var cands = storeGet(LS_COLS), merged = {}, i, m, k;
    for (i = 0; i < cands.length; i++) {
      m = parseColsMap(cands[i]);
      if (!m) continue;
      for (k in m) {
        if (!m.hasOwnProperty(k)) continue;
        if (!merged[k] || colsScore(m[k]) >= colsScore(merged[k])) merged[k] = m[k];
      }
    }
    return merged;
  }
  function loadCols() {
    return ensureWidthArr(layoutKey(), readColsMap()[layoutKey()]);
  }
  function loadPaneW() {
    var cands = storeGet(LS_PANE), i, n;
    for (i = 0; i < cands.length; i++) {
      n = parseInt(cands[i], 10);
      if (!isNaN(n)) return Math.max(180, Math.min(900, n));
    }
    return 400;
  }
  function parseOrdersRaw(raw) {
    if (raw == null || raw === '') return null;
    try {
      var o = typeof raw === 'string' ? JSON.parse(raw) : raw;
      if (o && typeof o === 'object' && !Array.isArray(o)) return o;
    } catch (eO) {}
    return null;
  }
  function readColOrders() {
    var cands = storeGet(LS_COLORD), i, o, best = null, n = 0, k, c;
    for (i = 0; i < cands.length; i++) {
      o = parseOrdersRaw(cands[i]);
      if (!o) continue;
      c = 0;
      for (k in o) if (o.hasOwnProperty(k) && Array.isArray(o[k])) c += o[k].length;
      if (c >= n) { n = c; best = o; }
    }
    if (best && best.tk && !best.list) best.list = best.tk.slice ? best.tk.slice() : best.tk;
    return best || {};
  }
  var colWidthMap = readColsMap();
  var colWidths = defaultWidths('list');
  var layoutView = '';
  function snapshotLayout(k) {
    k = k || layoutView;
    if (!k) return;
    if (colWidths && colWidths.length >= 3) colWidthMap[k] = colWidths.slice();
    if (typeof colOrderMap !== 'undefined' && colOrderMap && colOrderMap[k] && colOrderMap[k].slice) colOrderMap[k] = colOrderMap[k].slice();
  }
  function applyLayoutForView() {
    var k = layoutKey();
    if (layoutView && layoutView !== k) snapshotLayout(layoutView);
    layoutView = k;
    colWidthMap[k] = ensureWidthArr(k, colWidthMap[k]);
    colWidths = colWidthMap[k];
  }
  applyLayoutForView();
  function defaultColOrder(n) {
    var a = [], i;
    for (i = 1; i <= n; i++) a.push(i);
    return a;
  }
  var colOrderMap = readColOrders();
  function saveColOrders() {
    storeSet(LS_COLORD, colOrderMap);
  }
  function normalizeOrder(arr, n) {
    n = n || COLS;
    var used = {}, out = [], i, v;
    for (i = 0; i < (arr || []).length; i++) {
      v = parseInt(arr[i], 10);
      if (v >= 1 && v <= 26 && !used[v]) { used[v] = 1; out.push(v); }
    }
    for (i = 1; i <= n; i++) if (!used[i]) out.push(i);
    out = out.filter(function (x) { return x <= n; });
    for (i = 1; i <= n; i++) if (out.indexOf(i) < 0) out.push(i);
    return out;
  }
  function ensureColOrder() {
    var key = layoutView || layoutKey();
    colOrderMap[key] = normalizeOrder(colOrderMap[key], COLS);
    return colOrderMap[key];
  }
  function visToLog(v) {
    var o = ensureColOrder();
    return o[v - 1] || v;
  }
  function logToVis(c) {
    var o = ensureColOrder();
    var i = o.indexOf(c);
    return i < 0 ? c : i + 1;
  }
  function stepCol(c, delta) {
    var v = logToVis(c) + delta;
    if (v < 1) v = 1;
    if (v > COLS) v = COLS;
    return visToLog(v);
  }
  function moveColOrder(fromVis, toVis) {
    if (!fromVis || !toVis || fromVis === toVis) return false;
    var o = ensureColOrder().slice();
    if (fromVis < 1 || toVis < 1 || fromVis > o.length || toVis > o.length) return false;
    var item = o.splice(fromVis - 1, 1)[0];
    o.splice(toVis - 1, 0, item);
    colOrderMap[layoutView || layoutKey()] = o;
    saveColOrders();
    return true;
  }
  function resetColOrder() {
    colOrderMap[layoutView || layoutKey()] = defaultColOrder(COLS);
    saveColOrders();
    render();
    composeStat('Cols reset');
  }
  function clearColDropUi() {
    if (!shadow) return;
    Array.prototype.forEach.call(shadow.querySelectorAll('th.dragging, th.drop-before, th.drop-after'), function (el) {
      el.classList.remove('dragging', 'drop-before', 'drop-after');
    });
    var w = shadow.querySelector('.gridwrap');
    if (w) w.classList.remove('colmove');
  }
  var paneWidth = loadPaneW();
  var layoutDrag = null;
  function applyColWidths() {
    if (!shadow) return;
    var cols = shadow.querySelectorAll('table.grid col');
    var ths = shadow.querySelectorAll('table.grid thead th');
    var total = 0;
    for (var i = 0; i < cols.length; i++) {
      var ci = parseInt(cols[i].getAttribute('data-ci'), 10);
      if (isNaN(ci)) ci = i;
      var w = colWidths[ci] || 48;
      cols[i].style.width = w + 'px';
      cols[i].style.minWidth = w + 'px';
      cols[i].style.maxWidth = w + 'px';
      if (ths[i]) {
        ths[i].style.width = w + 'px';
        ths[i].style.minWidth = w + 'px';
        ths[i].style.maxWidth = w + 'px';
      }
      total += w;
    }
    var table = shadow.querySelector('table.grid');
    if (table) {
      table.style.width = total + 'px';
      table.style.minWidth = total + 'px';
      table.style.maxWidth = total + 'px';
    }
  }
  function applyPaneWidth() {
    var pane = shadow && shadow.querySelector('.pane');
    if (pane) pane.style.width = paneWidth + 'px';
  }
  function saveCols() {
    var k = layoutView || layoutKey();
    if (!colWidths || colWidths.length < 3) return;
    var storedMap = readColsMap();
    var storedOne = storedMap[k];
    var next = colWidths.slice();
    if (!(storedOne && colsScore(next) < colsScore(storedOne))) colWidthMap[k] = next;
    var key;
    for (key in storedMap) {
      if (!storedMap.hasOwnProperty(key)) continue;
      if (key === k) continue;
      if (!colWidthMap[key] && storedMap[key]) colWidthMap[key] = storedMap[key].slice ? storedMap[key].slice() : storedMap[key];
    }
    storeSet(LS_COLS, colWidthMap);
  }
  function savePane() {
    storeSet(LS_PANE, String(paneWidth));
  }
  function hydrateLayout() {
    var m = readColsMap(), k;
    for (k in m) {
      if (!m.hasOwnProperty(k)) continue;
      if (!colWidthMap[k] || colsScore(m[k]) >= colsScore(colWidthMap[k])) colWidthMap[k] = m[k];
    }
    var o = readColOrders();
    if (o) {
      if (o.tk && !o.list) o.list = o.tk.slice ? o.tk.slice() : o.tk;
      colOrderMap = o;
    }
    paneWidth = loadPaneW();
    applyLayoutForView();
    try { applyColWidths(); applyPaneWidth(); } catch (eH) {}
  }
  function fillGridToView() {
    if (!shadow) return false;
    var wrap = shadow.querySelector('.gridwrap');
    if (!wrap) return false;
    var w = wrap.clientWidth;
    var h = wrap.clientHeight;
    if (w < 80 || h < 80) return false;
    var sum = colWidths[0] || 32;
    var i;
    var ord = normalizeOrder(colOrderMap[layoutView || layoutKey()], 26);
    for (i = 1; i <= 10; i++) sum += colWidths[ord[i - 1]] || 56;
    var cols = 10;
    for (i = 11; i <= 26; i++) {
      if (sum >= w - 8) break;
      sum += colWidths[ord[i - 1]] || 56;
      cols = i;
    }
    var rows = Math.max(40, Math.floor(h / 18) + 8);
    if (cols === COLS && rows === MIN_ROWS) return false;
    COLS = cols;
    MIN_ROWS = rows;
    return true;
  }
  function colFromPoint(e) {
    if (!shadow) return -1;
    var ths = shadow.querySelectorAll('table.grid thead th');
    if (!ths.length) return -1;
    var head = ths[0].getBoundingClientRect();
    var last = ths[ths.length - 1].getBoundingClientRect();
    if (e.clientY < head.top - 2 || e.clientY > head.bottom + 2) return -1;
    if (e.clientX < head.left - 2 || e.clientX > last.right + 8) return -1;
    var edge = 8;
    var hit = -1;
    for (var i = 0; i < ths.length; i++) {
      var r = ths[i].getBoundingClientRect();
      if (Math.abs(e.clientX - r.right) <= edge) hit = i;
    }
    return hit;
  }
  function autoFit(col) {
    if (!shadow) return;
    var sel = col === 0 ? 'table.grid td.rh' : 'table.grid td[data-c="' + col + '"]';
    var nodes = shadow.querySelectorAll(sel);
    var probe = document.createElement('span');
    probe.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font:12px Calibri,DengXian,"Microsoft YaHei";padding:0 6px;';
    shadow.appendChild(probe);
    var max = 40;
    for (var i = 0; i < nodes.length; i++) {
      probe.textContent = nodes[i].textContent || '';
      var w = probe.offsetWidth + 8;
      if (w > max) max = w;
    }
    if (probe.parentNode) probe.parentNode.removeChild(probe);
    colWidths[col] = Math.min(560, Math.max(36, max));
    applyColWidths();
    saveCols();
  }

  function nfmt(v) {
    if (v === '' || v == null) return '';
    var n = Number(v);
    if (isNaN(n)) return String(v);
    var neg = n < 0;
    var s = Math.abs(Math.round(n)).toString().replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    return neg ? '(' + s + ')' : s;
  }
  function pfmt(v) {
    if (v === '' || v == null || isNaN(v)) return '';
    return (Number(v) * 100).toFixed(1) + '%';
  }
  function rng(seed) {
    var x = seed * 1103515245 + 12345;
    return function () {
      x = (x * 16807) % 2147483647;
      return (x & 2147483647) / 2147483647;
    };
  }

  function emptyGrid() {
    var g = [];
    for (var r = 1; r <= MIN_ROWS; r++) {
      g[r] = [];
      for (var c = 1; c <= COLS; c++) g[r][c] = { t: '' };
    }
    return g;
  }
  function rowOf(g, r) {
    if (!g[r]) {
      g[r] = [];
      for (var c = 1; c <= COLS; c++) g[r][c] = { t: '' };
    }
    return g[r];
  }
  function put(g, r, c, cell) {
    rowOf(g, r)[c] = cell;
  }
  var navBusy = false;
  var navPending = '';
  function sameOriginHref(href) {
    try { return new URL(href, location.href).origin === location.origin; } catch (e) { return false; }
  }
  function stripScripts(node) {
    if (!node || !node.querySelectorAll) return;
    var list = node.querySelectorAll("script");
    var i;
    for (i = list.length - 1; i >= 0; i--) {
      if (list[i].parentNode) list[i].parentNode.removeChild(list[i]);
    }
  }
  function applyFetchedHtml(html, href, opt) {
    opt = opt || {};
    state.nav.pageMeta = null;
    mergeUserInfo(html);
    var doc;
    try { doc = new DOMParser().parseFromString(html, "text/html"); } catch (eP) { throw eP; }
    if (!doc || !doc.body) throw new Error("parse");
    stripScripts(doc);
    var body = document.body;
    if (!body || !host) throw new Error("nobody");
    var n = body.firstChild;
    while (n) {
      var nx = n.nextSibling;
      if (n !== host) body.removeChild(n);
      n = nx;
    }
    var kid = doc.body.firstChild;
    while (kid) {
      var next = kid.nextSibling;
      try { body.insertBefore(document.importNode(kid, true), host); } catch (eI) {}
      kid = next;
    }
    stripScripts(body);
    try {
      layoutDrag = null;
      clearColDropUi();
      snapshotLayout(layoutView);
    } catch (eSnap) {}
    try {
      if (opt.history === "none") {}
      else if (opt.replace) history.replaceState({ xl: 1 }, "", href);
      else history.pushState({ xl: 1 }, "", href);
    } catch (eH) {}
    try { if (state) state.isRead = isReadPage(); } catch (eIs) {}
    try { applyLayoutForView(); } catch (eAl) {}
    lastSig = "";
    state.feedKey = "";
    try { resetFeedIfNeeded(); } catch (eR) {}
    navBusy = false;
    try { lockTitle(); setFavicon(); } catch (eL) {}
    scheduleRefresh();
    setTimeout(scheduleRefresh, 80);
    setTimeout(releaseIfLogin, 200);
    if (navPending && navPending !== href) {
      var p = navPending;
      navPending = '';
      silentGo(p);
    }
  }
  function silentGo(href, opt) {
    opt = opt || {};
    if (!href) return;
    var adDest = unwrapAdHref(href);
    if (adDest) href = adDest;
    if (navBusy && !opt.force) { navPending = href; return; }
    navBusy = true;
    navPending = '';
    composeStat(opt.reload ? "Refresh..." : "Opening...");
    httpReq({ url: href, html: 1 }).then(function (r) {
      if (!r || r.status >= 400) throw new Error("http");
      var landed = r.url || href;
      var dest = unwrapAdHref(landed);
      var adHit = looksLikeAdHtml(r.text);
      if (!adHit) {
        try { adHit = isAdPath(new URL(landed, location.href).pathname); } catch (eA) { adHit = isAdPath(landed); }
      }
      if (adHit) {
        dest = destFromAdHtml(r.text) || dest;
        if (dest && dest !== landed && !opt._adSkip) {
          opt._adSkip = 1;
          navBusy = false;
          silentGo(dest, opt);
          return;
        }
        if (dest) {
          navBusy = false;
          try { location.replace(dest); } catch (eR) { location.href = dest; }
          return;
        }
      }
      applyFetchedHtml(r.text, dest || landed, opt);
      if (!navBusy) composeStat("");
    }).catch(function () {
      navBusy = false;
      try { window.location.assign(href); } catch (e3) { window.location.href = href; }
    });
  }
  function destFromAdHtml(html) {
    var m = String(html || '').match(/\?https?:\/\/[^"'\s<>]+/i);
    if (m) return unwrapAdHref(originRoot() + '/misc/adpage_insert_2.html' + m[0]);
    m = String(html || '').match(/https?:\/\/[^"'\s<>]+\/(?:thread|read)\.php[^"'\s<>]*/i);
    return m ? unwrapAdHref(m[0]) : '';
  }
  function goTo(href) {
    if (!href) return;
    try { href = new URL(href, location.href).href; } catch (e0) {}
    var ad = unwrapAdHref(href);
    if (ad) href = ad;
    if (/^javascript:/i.test(href)) return;
    if (state.mode === "off" || !sameOriginHref(href)) {
      try { window.location.assign(href); } catch (e) { window.location.href = href; }
      return;
    }
    silentGo(href);
  }
  function originRoot() {
    return location.protocol + '//' + location.host;
  }

  function buildPL() {
    var g = emptyGrid();
    g[1][1] = { t: 'CW3 API 联调核对 · Sprint 2026-09', k: 'title' };
    g[2][1] = { t: 'MCT.CareCloud.Worker.MyCCN.CareWait3 · localhost:5048 · 更新 2026-09-20 18:10', k: 'sub' };
    var hs = ['CWM','PrefixRoute','Method','Auth','状态','环境','耗时ms','备注'];
    for (var i = 0; i < hs.length; i++) g[3][1 + i] = { t: hs[i], k: 'head' };
    var apis = [
      ['CWM-6503','/cw3/workspace/training/getcontactproviderinfobyemail','POST','Bearer','开发中','local:5048',86,'Email lookup，只读，不写库'],
      ['CWM-6503','/cw3/workspace/training/roster/preview','POST','Bearer','已合 develop','SHTest',142,'Contact/Provider 匹配口径对齐'],
      ['CWM-6503','/cw3/workspace/training/roster/commit','POST','Bearer','联调中','SHTest',210,'Add/Edit 与上传共用 Rows[]'],
      ['CWM-6488','/cw3/workspace/training/attendee/status','POST','Bearer','已合 develop','SHTest',54,'registered / waitlisted / attended / no_show'],
      ['CWM-6481','/cw3/workspace/training/attendee/remove','POST','Bearer','已合 develop','SHTest',67,'软删除，重算 Event Roster'],
      ['CWM-6460','/cw3/workspace/training/proof-of-hours','POST','Bearer','待前端','SHTest',188,'结构化 JSON，不生成服务端 PDF'],
      ['CWM-6442','/cw3/workspace/training/attendee-with-event','POST','Bearer','契约绿','local:5048',231,'Fields 投影 + event $in'],
      ['CWM-6410','/cw3/workspace/{Entity}/query','POST','Bearer','已上线','SHTest',96,'必带 datasource_code / deleted:null'],
      ['CWM-6388','/cw3/workspace/referral-enrollment/save','POST','Bearer','回归中','SHTest',175,'subsidy 与机构同一实体'],
      ['CWM-6362','/cw3/workspace/funding/rank/recalc','POST','Bearer','待验收','SHTest',640,'收入<0 当 0，清资格缓存'],
      ['CWM-6347','/cw3/workspace/family-application/save','POST','Bearer','已上线','SHTest',128,'FamilyFee.Range'],
      ['CWM-6298','/cw3/workspace/provider/profile/publish','POST','Bearer','已上线','SHTest',310,'publish 后推 SearchDB'],
      ['CWM-6248','/cw3/workspace/provider/duedate','POST','Bearer','已上线','SHTest',73,'DueDate 批量'],
      ['CWM-6187','/cw3/workspace/options/merge','POST','Bearer','已上线','SHTest',41,'系统级+租户级合并'],
      ['CWM-6154','/cw3/workspace/family-application/assign-org','POST','Bearer','已上线','SHTest',119,'首次分配 Org 时间'],
      ['MCCN-2381','/myccn/connection/refresh','POST','Bearer','已上线','Staging',155,'刷新 Connection 顺序']
    ];
    for (var r = 0; r < apis.length; r++) {
      var row = 4 + r, x = apis[r];
      var warn = /开发中|联调|回归|待验收|待前端/.test(x[4]);
      g[row][1] = { t: x[0], k: warn ? 'warn' : '' };
      g[row][2] = { t: x[1] };
      g[row][3] = { t: x[2] };
      g[row][4] = { t: x[3] };
      g[row][5] = { t: x[4], k: warn ? 'warn' : (x[4] === '契约绿' ? 'tot' : '') };
      g[row][6] = { t: x[5] };
      g[row][7] = { t: x[6], num: 1 };
      g[row][8] = { t: x[7] };
    }
    g[21][1] = { t: '核对要点', k: 'sec' };
    g[22][1] = { t: '1. 入口缺 WorkspaceId / BucketId 直接抛，不要默默查全库。' };
    g[23][1] = { t: '2. 列表/批量/Job 必须设 Fields，禁止整份 data 拉回。' };
    g[24][1] = { t: '3. Training 改完跑 ContractTests；family-application / referral / funding 只在 SHTest 验。' };
    g[25][1] = { t: '4. 本表是联调备忘，不含连接串 / token。' };
    return g;
  }

  function buildHC() {
    var g = emptyGrid();
    g[1][1] = { t: 'Training Contract / SHTest 用例', k: 'title' };
    g[2][1] = { t: 'CareWait3.ContractTests · run-contract-tests.sh · 2026-09-20', k: 'sub' };
    var hs = ['#','Case','Given','When','Then','结果','工单','耗时ms'];
    for (var i = 0; i < hs.length; i++) g[3][1 + i] = { t: hs[i], k: 'head' };
    var cases = [
      ['01','lookup 已有 Contact','emailNormalized 命中','getcontactproviderinfobyemail','返回 contact + providers','PASS','CWM-6503',42],
      ['02','lookup 新人 Email','无 Contact','同上','contact=null，表单手填','PASS','CWM-6503',38],
      ['03','lookup 多 Provider','同一 Email 2 条','同上','列出全部，不擅自选','TODO','CWM-6503',0],
      ['04','lookup 无 Email','email=""','同上','No Email','PASS','CWM-6503',12],
      ['05','lookup 非法 Email','格式不合法','同上','Invalid Email','PASS','CWM-6503',13],
      ['06','缺 WorkspaceId','header 不传','任意 /cw3','入口直接抛','PASS','—',9],
      ['07','roster preview 只读','有文件行','roster/preview','不写库，返回 diff','PASS','CWM-6503',155],
      ['08','Event=Complete 拒绝 preview','status=completed','roster/preview','拒绝','PASS','CWM-6503',28],
      ['09','commit 满员转 waitlisted','Registered+Attended 满','roster/commit','默认 waitlisted','PASS','CWM-6488',98],
      ['10','commit 无 FileName','手工 Add','roster/commit','不校验 fingerprint','PASS','CWM-6503',77],
      ['11','status 非法值','status=foo','attendee/status','拒绝','PASS','CWM-6488',16],
      ['12','remove 重算计数','删 1 条 attendee','attendee/remove','Roster + First/Last Seen','PASS','CWM-6481',81],
      ['13','proof-of-hours Contact','ContactId','proof-of-hours','出席史 JSON','PASS','CWM-6460',120],
      ['14','attendee-with-event 投影','Fields 只传 name','attendee-with-event','不整份拉取','PASS','CWM-6442',203],
      ['15','email 不可被覆盖','contactFields 含 email','roster/commit','忽略 email 写入','FAIL','CWM-6503',66],
      ['16','Provider ID 优先于 Email','同时有 ID 和 Email','roster/commit','走 entityId','PASS','CWM-6503',74]
    ];
    var pass = 0, fail = 0, todo = 0;
    for (var d = 0; d < cases.length; d++) {
      var row = 4 + d, x = cases[d];
      var rk = x[5] === 'FAIL' ? 'warn' : (x[5] === 'TODO' ? 'sub' : (x[5] === 'PASS' ? 'tot' : ''));
      g[row][1] = { t: x[0] };
      g[row][2] = { t: x[1] };
      g[row][3] = { t: x[2] };
      g[row][4] = { t: x[3] };
      g[row][5] = { t: x[4] };
      g[row][6] = { t: x[5], k: rk };
      g[row][7] = { t: x[6] };
      g[row][8] = { t: x[7], num: 1 };
      if (x[5] === 'PASS') pass++;
      else if (x[5] === 'FAIL') fail++;
      else todo++;
    }
    var tot = 21;
    g[tot][1] = { t: '合计', k: 'tot' };
    g[tot][2] = { t: cases.length + ' cases', k: 'tot' };
    g[tot][6] = { t: 'PASS ' + pass + ' / FAIL ' + fail + ' / TODO ' + todo, k: fail ? 'warn' : 'tot' };
    g[23][1] = { t: '核心链路 family-application / referral-enrollment / funding 无自动化测试，不进本表，SHTest 手验。', k: 'sub' };
    return g;
  }

  function buildAssumptions() {
    var g = emptyGrid();
    g[1][1] = { t: 'CW3 Entity / 查询口径', k: 'title' };
    g[2][1] = { t: 'CW3EntityRequest + FixRequest · 禁止直连写库 · 更新 2026-09-20', k: 'sub' };
    var hs = ['Entity','datasource_code','entityId','读写','查询必带','坑'];
    for (var i = 0; i < hs.length; i++) g[3][1 + i] = { t: hs[i], k: 'head' };
    var rows = [
      ['family-application','family-application','AP_','FixRequest','WorkspaceId,BucketId,datasource_code,deleted:null','status 在根级，不要 data.status'],
      ['referral-request','referral-request','RR_','FixRequest','同上','和 enrollment 不是一张单'],
      ['referral-enrollment','referral-enrollment','RF_/ER_','FixRequest','同上','type=subsidy 与机构同一实体'],
      ['provider','provider','PV_','FixRequest','同上','profile.email / data.email 双路径'],
      ['training-event','training-event','EV_','FixRequest','同上','Complete 后 roster preview 拒绝'],
      ['training-attendee','training-attendee','AT_','FixRequest','同上','列表必须 Fields 投影'],
      ['training-contact','training-contact','CON_','FixRequest','同上','匹配用 data.emailNormalized'],
      ['family-fee','family-fee','FF_','FixRequest','同上','Range 功能，CWM-6347'],
      ['options','options','—','options/merge','workspace + code','系统级与租户级合并'],
      ['metadata/status','workspace-metadata','—','只读配置','先判断能否改配置','别先改代码加状态']
    ];
    for (var r = 0; r < rows.length; r++) {
      for (var c = 0; c < rows[r].length; c++) g[4 + r][1 + c] = { t: rows[r][c] };
    }
    g[15][1] = { t: '硬约束', k: 'sec' };
    g[16][1] = { t: '同事给的 AP_/RR_/RF_/ER_/PV_ 是 entityId，不是 _id。' };
    g[17][1] = { t: 'Mongo 里 workspace_id / bucket_id 是 ObjectId，用字符串查会空结果且不报错。' };
    g[18][1] = { t: '批量修历史数据写 Exec/{版本}/ExecVersion*.cs 并注册，不要直接改库。' };
    g[19][1] = { t: '出站回调写在业务方法里同步 await，不要另开内部 notify Any。' };
    g[21][1] = { t: '本表给联调备忘用，不含连接串 / token / 真实 WorkspaceId。', k: 'sub' };
    return g;
  }

  function textOf(el) {
    return (el && (el.textContent || '').replace(/\s+/g, ' ').trim()) || '';
  }

  function linkHref(a) {
    var href = '';
    try { href = a.href || ''; } catch (e) { href = ''; }
    if (!href) href = a.getAttribute('href') || '';
    if (!href || href.indexOf('javascript:') === 0) return '';
    if (href.charAt(0) === '/') href = originRoot() + href;
    return href;
  }
  function isJunkTitle(title) {
    if (!title) return true;
    var s = String(title).replace(/\s+/g, ' ').trim();
    if (s.length < 2) return true;
    if (/^(上一页|下一页|末页|首页|尾页|\d+)$/.test(s)) return true;
    if (/^https?:\/\//i.test(s)) return true;
    if (/^(www\.|ngabbs\.|bbs\.nga)/i.test(s)) return true;
    if (/read\.php\?tid=/i.test(s) && !/[\u4e00-\u9fff]/.test(s)) return true;
    return false;
  }
  function tidKey(href) {
    var m = (href || '').match(/[?&]tid=(\d+)/i);
    return m ? m[1] : '';
  }
  function linkTitle(a) {
    var t = textOf(a);
    if (!isJunkTitle(t)) return t;
    t = (a.getAttribute('title') || '').trim();
    if (/打开新窗口/.test(t)) t = '';
    if (!isJunkTitle(t)) return t;
    var box = a.closest && a.closest('td, .c3');
    if (box) {
      var t2 = textOf(box).replace(/https?:\/\/\S+/g, ' ').replace(/\s+/g, ' ').trim();
      if (!isJunkTitle(t2)) return t2.slice(0, 80);
    }
    return '';
  }
  function titleScore(title, a) {
    var s = 0;
    if (a.classList && a.classList.contains('topic')) s += 50;
    if (/[\u4e00-\u9fff]/.test(title)) s += 30;
    if (!/^https?:/i.test(title)) s += 20;
    s += Math.min(title.length, 40);
    return s;
  }
  function pageWin() {
    try { return (typeof unsafeWindow !== 'undefined' && unsafeWindow) || window; } catch (e) { return window; }
  }
  function cookieGet(n) {
    var parts = ('; ' + document.cookie).split('; ' + n + '=');
    if (parts.length < 2) return '';
    return decodeURIComponent(parts.pop().split(';').shift());
  }
  function currentUid() {
    return cookieGet('ngaPassportUid') || cookieGet('_178c') || '';
  }
  function qsGet(name, href) {
    var src = href || location.search || '';
    var m = String(src).match(new RegExp('[?&#]' + name + '=([^&#]*)', 'i'));
    return m ? decodeURIComponent(m[1].replace(/\+/g, ' ')) : '';
  }
  function currentPageN() {
    return parseInt(qsGet('page') || '1', 10) || 1;
  }
  function currentStid() {
    return qsGet('stid') || qsGet('colid') || '';
  }
  function draftKey() {
    var tid = currentTid();
    if (tid) return 'tid:' + tid;
    var fid = currentFid();
    if (fid) return 'fid:' + fid;
    return (location.pathname || '') + (location.search || '');
  }
  function saveDraft() {
    if (!shadow) return;
    var ta = shadow.querySelector('.compose .cbody');
    var ti = shadow.querySelector('.compose .ctitle');
    var body = ta ? ta.value : '';
    var title = ti ? ti.value : '';
    if (!(body || title)) { clearDraft(); return; }
    try {
      localStorage.setItem(LS_DRAFT, JSON.stringify({
        key: draftKey(),
        body: body,
        title: title,
        comment: composeHold.comment ? 1 : 0,
        ts: Date.now()
      }));
    } catch (e) {}
  }
  function clearDraft() {
    try {
      var raw = localStorage.getItem(LS_DRAFT);
      if (!raw) return;
      var d = JSON.parse(raw);
      if (!d || d.key === draftKey()) localStorage.removeItem(LS_DRAFT);
    } catch (e2) {}
  }
  function restoreDraft() {
    if (!shadow) return;
    try {
      var raw = localStorage.getItem(LS_DRAFT);
      if (!raw) return;
      var d = JSON.parse(raw);
      if (!d || d.key !== draftKey()) return;
      if (d.ts && (Date.now() - d.ts) > 7 * 86400000) { localStorage.removeItem(LS_DRAFT); return; }
      var ta = shadow.querySelector('.compose .cbody');
      var ti = shadow.querySelector('.compose .ctitle');
      if (ta && d.body && !ta.value) ta.value = d.body;
      if (ti && d.title && !ti.value) ti.value = d.title;
      if (d.comment) composeHold.comment = 1;
      syncCompose();
    } catch (e3) {}
  }
  function lockTitle() {
    if (state.mode === 'off') return;
    var want = FILENAME + ' - Excel';
    if (document.title !== want) document.title = want;
    try { setFavicon(); } catch (eF) {}
  }
  function hideFileMenu() {
    var m = shadow && shadow.querySelector('.filemenu');
    var b = shadow && shadow.querySelector('.filebtn');
    if (m) m.classList.add('off');
    if (b) b.classList.remove('on');
  }
  function toggleFileMenu() {
    var m = shadow && shadow.querySelector('.filemenu');
    var b = shadow && shadow.querySelector('.filebtn');
    if (!m) return;
    var on = m.classList.contains('off');
    m.classList.toggle('off', !on);
    if (b) b.classList.toggle('on', on);
  }
  function copySelCell() {
    var cell = state.cells && state.cells[state.selR] && state.cells[state.selR][state.selC];
    var t = cell && cell.t != null ? String(cell.t) : '';
    if (!t) t = rowHref(state.selR) || '';
    copyText(t);
    composeStat('Copied');
  }
  function copyRowLink() {
    var pv = selPreview();
    var href = '';
    if (pv && pv.pid) href = originRoot() + '/read.php?pid=' + encodeURIComponent(pv.pid);
    if (!href) href = rowHref(state.selR) || location.href;
    var t = (pv && pv.title) ? (pv.title + '\n' + href) : href;
    copyText(t);
    composeStat('Copied');
  }
  function eventInPaneText(e) {
    var path = e.composedPath ? e.composedPath() : [];
    var i, n;
    for (i = 0; i < path.length; i++) {
      n = path[i];
      if (n && n.classList && n.classList.contains('body') && n.closest && n.closest('.pane')) return true;
    }
    try {
      var s = shadow && shadow.getSelection && shadow.getSelection();
      if (s && String(s).length) return true;
    } catch (e2) {}
    return false;
  }
  function hardReload() {
    silentGo(location.href, { replace: 1, reload: 1, force: 1 });
  }
  function gotoLast() {
    var mx = (state.nav && state.nav.maxPage) || 0;
    var href = location.href || '';
    if (isReadPage()) {
      if (mx > 1) gotoPage(mx);
      else gotoPage(9999);
      return;
    }
    if (mx > 100 && /thread\.php/i.test(href) && !/[?&]nounion=1/.test(href)) mx = 100;
    if (mx > 1) gotoPage(mx);
    else composeStat('Last');
  }
  function toggleDigest() {
    var fid = currentFid();
    if (!fid) { composeStat('No board'); return; }
    if (/[?&]recommend=1/.test(location.search || '')) {
      goTo(originRoot() + '/thread.php?fid=' + encodeURIComponent(fid));
      return;
    }
    goTo(originRoot() + '/thread.php?fid=' + encodeURIComponent(fid) + '&recommend=1&order_by=postdatedesc');
  }
  function actOwnerFilter() {
    var pv = selPreview();
    var uid = (pv && pv.uid) || '';
    var tid = currentTid();
    var fid = currentFid();
    if (!uid) { composeStat('No assignee'); return; }
    if (tid) {
      if (String(qsGet('authorid') || '') === String(uid)) {
        goTo(originRoot() + '/read.php?tid=' + encodeURIComponent(tid));
        return;
      }
      goTo(originRoot() + '/read.php?tid=' + encodeURIComponent(tid) + '&authorid=' + encodeURIComponent(uid));
      return;
    }
    if (fid) {
      if (String(qsGet('authorid') || '') === String(uid)) {
        goTo(originRoot() + '/thread.php?fid=' + encodeURIComponent(fid));
        return;
      }
      goTo(originRoot() + '/thread.php?fid=' + encodeURIComponent(fid) + '&authorid=' + encodeURIComponent(uid));
      return;
    }
    composeStat('No board');
  }
  function actMyTopics() {
    var uid = currentUid();
    if (!uid) { composeStat('Sign-in'); return; }
    goTo(originRoot() + '/thread.php?authorid=' + encodeURIComponent(uid));
  }
  function runFile(act) {
    if (act === 'reload') hardReload();
    else if (act === 'copy') copyRowLink();
    else if (act === 'mine') actMyTopics();
    else if (act === 'back') goTo(state.nav && state.nav.back);
    else if (act === 'owner') actOwnerFilter();
    else if (act === 'digest') toggleDigest();
    else if (act === 'last') gotoLast();
    else if (act === 'cols') resetColOrder();
    else if (act === 'off') setMode('off');
  }
  var composeHold = { attachments: [], checks: [], urls: [], edit: null, comment: 0, vcodeUrl: '', vcodeName: 'vcode' };
  function gmXhr() {
    if (typeof GM_xmlhttpRequest === 'function') return GM_xmlhttpRequest;
    try { if (typeof GM !== 'undefined' && GM.xmlHttpRequest) return GM.xmlHttpRequest; } catch (e) {}
    return null;
  }
  function sniffCharset(u8, header) {
    var cs = '';
    var m = String(header || '').match(/charset\s*=\s*["']?([a-z0-9_\-]+)/i);
    if (m) cs = m[1];
    if (!cs && u8 && u8.length) {
      var head = '';
      var n = Math.min(u8.length, 4096);
      var i;
      for (i = 0; i < n; i++) head += String.fromCharCode(u8[i]);
      m = head.match(/charset\s*=\s*["']?([a-z0-9_\-]+)/i);
      if (m) cs = m[1];
    }
    cs = String(cs || '').toLowerCase();
    if (/gbk|gb2312|gb18030|gb_2312/.test(cs)) return 'gb18030';
    if (/utf-?8/.test(cs)) return 'utf-8';
    if (/big5/.test(cs)) return 'big5';
    return cs;
  }
  function decodeBytes(u8, cs) {
    try { return new TextDecoder(cs).decode(u8); } catch (e) {
      try { return new TextDecoder(cs === 'gb18030' ? 'gbk' : 'utf-8').decode(u8); } catch (e2) { return ''; }
    }
  }
  function cjkCount(s) { return (String(s).match(/[\u4e00-\u9fff]/g) || []).length; }
  function replCount(s) { return (String(s).match(/\uFFFD/g) || []).length; }
  function decodeNgaHtml(buf, header) {
    var u8 = buf instanceof Uint8Array ? buf : new Uint8Array(buf || 0);
    if (!u8.length) return '';
    var hinted = sniffCharset(u8, header);
    var utf = decodeBytes(u8, 'utf-8');
    var gbk = decodeBytes(u8, 'gb18030');
    var utfBad = replCount(utf);
    var utfCjk = cjkCount(utf);
    var gbkCjk = cjkCount(gbk);
    var pick = utf;
    if (hinted === 'gb18030') pick = gbk;
    else if (hinted === 'utf-8') pick = utf;
    else if (utfBad > 2 && gbkCjk >= utfCjk) pick = gbk;
    else if (utfBad === 0 && utfCjk > 0 && utfCjk >= gbkCjk) pick = utf;
    else if (gbkCjk > utfCjk) pick = gbk;
    if (hinted === 'gb18030' && utfBad === 0 && utfCjk > gbkCjk + 10) pick = utf;
    var out = String(pick || utf || gbk || '');
    if (/<html|<head|<!doctype/i.test(out.slice(0, 800))) {
      out = out.replace(/charset\s*=\s*["']?[\w\-]+/ig, 'charset=UTF-8');
    }
    return out;
  }
  function httpReq(opts) {
    var url = opts.url;
    var method = (opts.method || 'GET').toUpperCase();
    var body = opts.body || null;
    var headers = opts.headers || {};
    var gm = gmXhr();
    function viaGM() {
      if (!gm) return Promise.reject(new Error('no gm'));
      return new Promise(function (resolve, reject) {
        gm({
          method: method,
          url: url,
          data: body,
          headers: headers,
          anonymous: false,
          withCredentials: true,
          responseType: 'arraybuffer',
          onload: function (r) {
            var hdr = r.responseHeaders || '';
            var text = r.response ? decodeNgaHtml(r.response, hdr) : (r.responseText || '');
            resolve({ status: r.status, text: text, url: r.finalUrl || url });
          },
          onerror: function () { reject(new Error('net')); }
        });
      });
    }
    if (opts.forceGM) return viaGM();
    return fetch(url, { method: method, credentials: 'include', body: body, headers: headers, redirect: 'follow' }).then(function (r) {
      var ct = '';
      try { ct = r.headers.get('content-type') || ''; } catch (eH) {}
      return r.arrayBuffer().then(function (buf) {
        return { status: r.status, text: decodeNgaHtml(buf, ct), url: r.url || url };
      });
    }).catch(function (e) {
      if (gm) return viaGM();
      throw e;
    });
  }
  function parseNuke(text) {
    var s = String(text || '').replace(/^\uFEFF/, '');
    var idx = s.indexOf('window.script_muti_get_var_store');
    if (idx >= 0) {
      var eq = s.indexOf('=', idx);
      if (eq >= 0) {
        var json = s.slice(eq + 1).replace(/^\s*/, '');
        json = json.replace(/<\/script>[\s\S]*$/i, '').replace(/;?\s*$/, '');
        try { return JSON.parse(json); } catch (e) {
          try { return Function('return (' + json + ')')(); } catch (e2) {}
        }
      }
    }
    try { return JSON.parse(s); } catch (e3) {}
    var brace = s.indexOf('{');
    if (brace >= 0) {
      try { return JSON.parse(s.slice(brace).replace(/<\/script>[\s\S]*$/i, '').replace(/;?\s*$/, '')); } catch (e4) {}
    }
    return { raw: s };
  }
  function nukeMsg(obj, raw) {
    var t = raw || '';
    if (obj) {
      if (obj.error) return typeof obj.error === 'string' ? obj.error : (obj.error[0] || JSON.stringify(obj.error));
      if (obj.data && obj.data[0]) return String(obj.data[0]);
    }
    var m = t.match(/ERROR[:：]\s*([^\n<]+)/i);
    if (m) return m[1].trim();
    if (/发贴完毕|发帖完毕|操作成功|完成/.test(t)) return 'ok';
    return '';
  }
  function formBody(fields) {
    var p = new URLSearchParams();
    Object.keys(fields || {}).forEach(function (k) {
      if (fields[k] == null) return;
      p.set(k, String(fields[k]));
    });
    return p.toString();
  }
  function nukeFail(obj, msg) {
    if (obj && obj.error) {
      var er = obj.error;
      return typeof er === 'string' ? er : (er[0] || JSON.stringify(er));
    }
    var t = String(msg || '');
    if (/ERROR[:：]\s*([^\n<]+)/i.test(t)) return t.match(/ERROR[:：]\s*([^\n<]+)/i)[1].trim();
    if (/未登录|请先登录|验证码/.test(t) && !/支持|反对|操作成功/.test(t)) return t;
    if (/失败|错误/.test(t) && !/支持|反对|操作成功|表示/.test(t)) return t;
    return '';
  }
  function nukePost(params) {
    var merged = { __output: 11 };
    Object.keys(params || {}).forEach(function (k) { merged[k] = params[k]; });
    var u = originRoot() + '/nuke.php?';
    return httpReq({ method: 'POST', url: u, body: formBody(merged), headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' } }).then(function (r) {
      var obj = parseNuke(r.text);
      var msg = nukeMsg(obj, r.text);
      var fail = nukeFail(obj, msg);
      if (fail) throw new Error(fail);
      if (!obj.data && msg !== 'ok' && !msg) throw new Error('fail');
      return { obj: obj, text: r.text, msg: msg };
    });
  }
  function callPage(path, args) {
    try {
      var cur = pageWin();
      var parts = path.split('.');
      for (var i = 0; i < parts.length; i++) {
        if (!cur) return false;
        cur = cur[parts[i]];
      }
      if (typeof cur !== 'function') return false;
      cur.apply(null, args || []);
      return true;
    } catch (e) { return false; }
  }
  function pageHrefFrom(base, n) {
    try {
      var u = new URL(base || location.href, location.href);
      n = parseInt(n, 10);
      if (!n || n <= 1) u.searchParams.delete('page');
      else u.searchParams.set('page', String(n));
      return u.href;
    } catch (e) { return ''; }
  }
  function scrapeThreadTitle() {
    var sels = ['#postsubject0', '#postsubject', '#toptpc', 'h1#toptpc', '.topicbox h1', '#m_posts h3'];
    var i, el, t;
    for (i = 0; i < sels.length; i++) {
      el = document.querySelector(sels[i]);
      t = textOf(el);
      if (!t || t.indexOf(FILENAME) >= 0) continue;
      if (t.length < 2) continue;
      return t.slice(0, 140);
    }
    t = (document.title || '').replace(/\s*[-_].*NGA.*$/i, '').replace(/\s*NGA.*$/i, '').trim();
    if (t && t.indexOf(FILENAME) < 0 && t.length > 1) return t.slice(0, 140);
    return '';
  }
  function resetFeedIfNeeded() {
    var key = location.pathname + (location.search || '').replace(/#.*$/, '');
    if (state.feedKey !== key) {
      state.feedKey = key;
      state.feed = { posts: [], seen: {}, next: '', loading: false, done: false, tried: {} };
    }
    if (!state.feed) state.feed = { posts: [], seen: {}, next: '', loading: false, done: false, tried: {} };
  }
  function mergeFeed(posts) {
    resetFeedIfNeeded();
    var i, k, p;
    for (i = 0; i < (posts || []).length; i++) {
      p = posts[i];
      k = p.pid ? ('pid' + p.pid) : ('f' + p.floor + ':' + (p.time || '') + ':' + (p.author || ''));
      if (state.feed.seen[k]) continue;
      state.feed.seen[k] = 1;
      state.feed.posts.push(p);
    }
    state.feed.posts.sort(function (a, b) { return (a.floor || 0) - (b.floor || 0); });
  }
  function loadMorePosts() {
    if (state.mode === 'off' || state.sheet !== 'tk' || !state.isRead) return;
    resetFeedIfNeeded();
    if (state.feed.loading || state.feed.done) return;
    var next = state.feed.next || (state.nav && state.nav.next) || '';
    if (!next) {
      var cur = (state.nav && state.nav.page) || currentPageN();
      next = pageHrefFrom(location.href, cur + 1);
    }
    if (!next || next === location.href) { state.feed.done = true; return; }
    state.feed.tried = state.feed.tried || {};
    if (state.feed.tried[next]) { state.feed.done = true; return; }
    state.feed.tried[next] = 1;
    state.feed.loading = true;
    httpReq({ url: next, html: 1 }).then(function (r) {
      mergeUserInfo(r.text);
      var doc = new DOMParser().parseFromString(r.text, 'text/html');
      var posts = scrapePosts(doc);
      var before = state.feed.posts.length;
      mergeFeed(posts);
      var pg = scrapePager(doc, next);
      state.feed.next = pg.next || '';
      if (pg.max && state.nav && pg.max > (state.nav.maxPage || 0)) state.nav.maxPage = pg.max;
      if (state.feed.posts.length === before) state.feed.done = true;
      var wrap = shadow && shadow.querySelector('.gridwrap');
      var sl = wrap ? wrap.scrollLeft : 0;
      var st = wrap ? wrap.scrollTop : 0;
      render();
      wrap = shadow && shadow.querySelector('.gridwrap');
      if (wrap) { wrap.scrollLeft = sl; wrap.scrollTop = st; }
    }).catch(function () {
      state.feed.done = true;
    }).then(function () {
      state.feed.loading = false;
    });
  }
  function isFloorLabel(s) {
    s = String(s || '').replace(/\s+/g, '');
    return /^#?\d+$/.test(s);
  }
  function extractJsonAfter(s, key) {
    s = String(s || '');
    var i = s.indexOf(key);
    if (i < 0) return null;
    i = s.indexOf('{', i);
    if (i < 0) return null;
    var depth = 0, j, q = '', esc = false, ch;
    for (j = i; j < s.length; j++) {
      ch = s[j];
      if (q) {
        if (esc) { esc = false; continue; }
        if (ch === '\\') { esc = true; continue; }
        if (ch === q) q = '';
        continue;
      }
      if (ch === '"' || ch === "'") { q = ch; continue; }
      if (ch === '{') depth++;
      else if (ch === '}') {
        depth--;
        if (!depth) {
          try { return JSON.parse(s.slice(i, j + 1)); } catch (e) { return null; }
        }
      }
    }
    return null;
  }
  function parsePageMeta(html) {
    html = String(html || '');
    var fm = html.match(/__CURRENT_FID\s*=\s*(-?\d+)/);
    if (fm) state.lastFid = fm[1];
    var pm = html.match(/__PAGE\s*=\s*\{([^}]{0,400})\}/);
    if (!pm) return;
    try {
      var obj = Function('return ({' + pm[1] + '})')();
      var a = parseInt(obj[1], 10) || 0;
      var cur = parseInt(obj[2], 10) || 0;
      var per = parseInt(obj[3], 10) || 0;
      var mx = 1;
      if (a > 0 && per > 1 && a > per * 2) mx = Math.ceil(a / per);
      else if (a > 0) mx = a;
      if (cur > mx) mx = cur;
      state.nav.pageMeta = { a: a, page: cur, per: per, max: mx };
    } catch (eM) {}
  }
  function mergeUserInfo(html) {
    parsePageMeta(html);
    var obj = extractJsonAfter(html, 'userInfo.setAll(') || extractJsonAfter(html, 'userInfo.setAll (');
    if (!obj) return;
    if (!state.userMap) state.userMap = {};
    Object.keys(obj).forEach(function (k) {
      var u = obj[k];
      var name = u && (u.username || u.name);
      if (name) state.userMap[String(k)] = String(name);
    });
  }
  function nameFromUid(uid) {
    if (!uid || !state.userMap) return '';
    return state.userMap[String(uid)] || '';
  }
  function pickAuthor(root) {
    var best = { name: '', uid: '', href: '' };
    if (!root || !root.querySelectorAll) return best;
    var nodes = root.querySelectorAll('a[href*="uid="], a.author, .postername a, a[id^="postauthor"], .username');
    var i, n, name, href, uid, um, inner;
    for (i = 0; i < nodes.length; i++) {
      n = nodes[i];
      name = textOf(n);
      href = n.tagName === 'A' ? linkHref(n) : '';
      if ((!name || isFloorLabel(name)) && n.querySelector) {
        inner = n.querySelector('a[href*="uid="]');
        if (inner) {
          name = textOf(inner) || name;
          href = linkHref(inner) || href;
        }
      }
      um = ((href || '') + ' ' + (n.getAttribute('onclick') || '')).match(/uid=(\d+)/i);
      uid = um ? um[1] : '';
      if (uid && !best.uid) {
        best.uid = uid;
        best.href = href || best.href;
      }
      if (!name || isFloorLabel(name)) continue;
      if (!href && uid) href = originRoot() + '/nuke.php?func=ucp&uid=' + uid;
      if (uid) best.uid = uid;
      best.name = name;
      best.href = href || best.href;
      break;
    }
    if (!best.name && best.uid) best.name = nameFromUid(best.uid);
    if (!best.href && best.uid) best.href = originRoot() + '/nuke.php?func=ucp&uid=' + best.uid;
    return best;
  }
  function scrapeFloorNo(box, idx) {
    var t = '';
    if (box && box.querySelector) {
      var el = box.querySelector('a[id^="pid"], a[href^="#pid"], a[href*="#pid"], .postinfo, [id^="postdate"]');
      t = textOf(el);
    }
    var m = String(t).match(/#(\d+)/);
    if (m) return parseInt(m[1], 10);
    return idx;
  }
  function scrapeAuthorInfo(box, idx, root) {
    root = root || document;
    var pi = null;
    if (idx != null && root.getElementById) {
      pi = root.getElementById('posterinfo' + idx) || root.getElementById('postauthor' + idx);
    }
    var picked = pickAuthor(pi || box);
    if (!picked.name && box && box !== pi) {
      var p2 = pickAuthor(box);
      if (p2.name || (!picked.uid && p2.uid)) picked = p2;
      else if (!picked.uid && p2.uid) picked.uid = p2.uid;
    }
    if (!picked.name && pi && pi.parentElement) {
      var p3 = pickAuthor(pi.parentElement);
      if (p3.name) picked = p3;
      else if (!picked.uid && p3.uid) picked.uid = p3.uid;
    }
    if (!picked.name && picked.uid) picked.name = nameFromUid(picked.uid);
    return { name: picked.name || '—', uid: picked.uid, href: picked.href };
  }
  function decodeHtmlText(s) {
    s = String(s || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '');
    try {
      var ta = document.createElement('textarea');
      ta.innerHTML = s;
      return ta.value;
    } catch (eD) { return s; }
  }
  function parseCommentBody(html) {
    var s = decodeHtmlText(html);
    var replyWho = '';
    var hm = s.match(/Post by\s*(?:\[uid=\d+\])?([^\[]+?)(?:\[\/uid\])?\s*\(([^)]*)\)/i);
    if (hm) replyWho = hm[1].replace(/^\s+|\s+$/g, '');
    s = s.replace(/\[b\]\s*Reply to[\s\S]*?\[\/b\]/gi, '');
    s = s.replace(/Reply to[\s\S]*?Post by[\s\S]*?\(\d{4}-\d{2}-\d{2}[^\)]*\)\s*/gi, '');
    s = stripUbb(s);
    s = s.replace(/^\s*Reply to\s+Topic\s*/i, '').replace(/^\s+|\s+$/g, '');
    return { text: s, replyWho: replyWho };
  }
  function scrapeComments(box, idx, root) {
    root = root || document;
    var list = [];
    var seen = {};
    function isBodyId(id) {
      return /^postcomment_+\d+$/i.test(id || '');
    }
    function addFrom(block) {
      if (!block) return;
      var bodyEl = null;
      if (block.getAttribute && isBodyId(block.id)) bodyEl = block;
      if (!bodyEl && block.querySelectorAll) {
        var nodes = block.querySelectorAll('span[id^="postcomment_"]');
        var bi;
        for (bi = 0; bi < nodes.length; bi++) {
          if (isBodyId(nodes[bi].id)) { bodyEl = nodes[bi]; break; }
        }
      }
      if (!bodyEl || seen[bodyEl.id]) return;
      seen[bodyEl.id] = 1;
      var wrap = block;
      if (block.closest) wrap = block.closest('div.comment_c') || block.closest('table.comment_c_0') || block;
      var a = wrap.querySelector && wrap.querySelector('a[id^="commentauthor"], a.author, a[href*="uid="]');
      var uid = '';
      var who = '';
      if (a) {
        who = textOf(a);
        var um = ((a.getAttribute('href') || '') + '').match(/[?&]uid=(\d+)/i);
        if (um) uid = um[1];
      }
      if (uid) {
        var mapped = nameFromUid(uid);
        if (mapped) who = mapped;
      }
      if (/^UID:\d+$/i.test(who)) who = nameFromUid(uid) || who;
      var time = '';
      if (wrap.querySelector) time = textOf(wrap.querySelector('[title="reply time"]')) || textOf(wrap.querySelector('.postInfo')).slice(0, 22);
      var parsed = parseCommentBody(bodyEl.innerHTML || bodyEl.textContent || '');
      if (!parsed.text || parsed.text.length < 1) return;
      list.push({ who: who, uid: uid, time: time, text: parsed.text.slice(0, 800), replyWho: parsed.replyWho });
    }
    var scopes = [];
    if (root.getElementById) {
      var a1 = root.getElementById('comment_for_' + idx);
      var a2 = root.getElementById('hightlight_for_' + idx);
      if (a1) scopes.push(a1);
      if (a2) scopes.push(a2);
    }
    if (box && scopes.length === 0) scopes.push(box);
    var si, blocks, bi2;
    for (si = 0; si < scopes.length; si++) {
      blocks = scopes[si].querySelectorAll ? scopes[si].querySelectorAll('div.comment_c, span[id^="postcomment_"]') : [];
      if (blocks.length) {
        for (bi2 = 0; bi2 < blocks.length; bi2++) addFrom(blocks[bi2]);
      } else addFrom(scopes[si]);
    }
    return list.slice(0, 40);
  }
  function scrapePoll(box) {
    var root = (box && box.querySelector && (box.querySelector('.vote, #votec, table.vote, .poll, [id^="vote"]'))) || document.querySelector('.vote, table.vote, #votec');
    if (!root) return null;
    var opts = [];
    Array.prototype.forEach.call(root.querySelectorAll('input[type="radio"], input[type="checkbox"]'), function (inp) {
      var lab = '';
      if (inp.id) {
        var l = root.querySelector('label[for="' + inp.id + '"]');
        if (l) lab = textOf(l);
      }
      if (!lab && inp.parentElement) lab = textOf(inp.parentElement);
      lab = lab.replace(/\s+/g, ' ').trim();
      if (!lab) return;
      opts.push({ id: inp.value || String(opts.length), text: lab.slice(0, 80), name: inp.name || 'voteid' });
    });
    if (!opts.length) {
      Array.prototype.forEach.call(root.querySelectorAll('tr, li, .voteitem'), function (row) {
        var t = textOf(row);
        if (!t || t.length < 2 || /投票|提交/.test(t) && t.length < 6) return;
        var idm = (row.getAttribute('data-id') || row.innerHTML || '').match(/vote(?:id)?[=:]?\s*(\d+)/i);
        opts.push({ id: idm ? idm[1] : String(opts.length), text: t.slice(0, 80), name: 'voteid' });
      });
    }
    if (!opts.length) return null;
    return { opts: opts.slice(0, 20) };
  }
  function scrapeCollections() {
    var rows = [];
    var seen = {};
    Array.prototype.forEach.call(document.querySelectorAll('a[href*="stid="], a[href*="colid="]'), function (a) {
      var href = linkHref(a);
      if (/post\.php/i.test(href)) return;
      var m = href.match(/[?&](stid|colid)=(-?\d+)/i);
      if (!m) return;
      var title = textOf(a).replace(/\s+/g, ' ').trim();
      if (/发表新帖/.test(title)) return;
      if (!title || title.length > 40) title = title.slice(0, 40);
      if (!title || seen[m[0]]) return;
      if (a.closest && a.closest('#pagebbtm, .pager, #footer')) return;
      seen[m[0]] = 1;
      rows.push({ title: title, href: href.replace(/#.*$/, ''), kind: m[1], id: m[2] });
    });
    return rows;
  }
  function scrapeVcode(doc) {
    doc = doc || document;
    var inp = doc.querySelector('input[name="vcode"], input[name="captcha"], #vcode, input[name="check_code"]');
    var img = doc.querySelector('#vcodeimg, img[src*="check_code"], img[src*="vcode"], img[src*="captcha"], .vcode img');
    if (!inp && !img) return null;
    var src = '';
    if (img) src = img.getAttribute('src') || img.src || '';
    if (src && src.indexOf('http') !== 0) {
      try { src = new URL(src, location.href).href; } catch (e) {}
    }
    return { name: (inp && inp.name) || 'vcode', img: src || (originRoot() + '/login/check_code.php?' + Date.now()) };
  }
  function findLoginForm() {
    var forms = document.querySelectorAll('form');
    for (var i = 0; i < forms.length; i++) {
      if (forms[i].querySelector('input[type="password"]')) return forms[i];
    }
    return null;
  }
  function findRuleBtn() {
    var nodes = document.querySelectorAll('a, button, input[type="submit"], input[type="button"]');
    for (var i = 0; i < nodes.length; i++) {
      var t = String(nodes[i].value || nodes[i].textContent || '').replace(/\s+/g, '');
      if (/同意并进入|我已阅读|同意版规|接受并继续|进入版面/.test(t)) return nodes[i];
      if (t === '同意' || t === '接受') return nodes[i];
    }
    return null;
  }
  function scrapeNgaNotice() {
    var html = (document.body && document.body.innerHTML) || '';
    var re = /<!--msginfostart-->([\s\S]*?)<!--msginfoend-->/g;
    var m;
    var first = '';
    while ((m = re.exec(html))) {
      var tmp = document.createElement('div');
      tmp.innerHTML = m[1];
      var s = (tmp.textContent || '').replace(/\s+/g, ' ').trim();
      s = s.replace(/\[恢复无法查看的主题点此\]/g, '').replace(/\[查看所需的权限\/条件\]/g, '').trim();
      if (!s) continue;
      if (!first) first = s;
    }
    if (first) return first.slice(0, 80);
    var title = (document.title || '').replace(/\s*NGA.*$/i, '').trim();
    if (/找不到主题|审核未通过|权限不足|已被删除|被删除|被隐藏|超过限制/.test(title)) return title.slice(0, 80);
    return '';
  }
  function isTopicBlocked(msg) {
    return /审核未通过|审核中|权限不足|已被删除|被删除|找不到主题|主题不存在|被隐藏|设为隐藏|超过限制|被屏蔽|被锁|nuke/i.test(msg || '');
  }
  function isLoginWall() {
    if (!document.body) return false;
    if (document.querySelector('a.topic, [id^="postcontent"]')) return false;
    var msg = scrapeNgaNotice();
    if (isTopicBlocked(msg)) return false;
    if (findLoginForm()) return true;
    var t = document.body.innerText || '';
    if (isTopicBlocked(t)) return false;
    return /未登录|请先登录|你可能需要登录/.test(t);
  }
  function needsRule() {
    if (!document.body) return false;
    if (document.querySelector('a.topic, [id^="postcontent"]')) return false;
    var t = document.body.innerText || '';
    return /版规|我已阅读|同意.*进入/.test(t) && !!findRuleBtn();
  }
  function userPage(uid, href) {
    if (href && /nuke\.php|ucp|uid=/.test(href)) return href;
    if (uid) return originRoot() + '/nuke.php?func=ucp&uid=' + encodeURIComponent(uid);
    return '';
  }
  function hideDlg() {
    var d = shadow && shadow.querySelector('#xl-dlg');
    if (d) d.classList.add('off');
  }
  function showDlg(title, bodyHtml, footerHtml) {
    var d = shadow && shadow.querySelector('#xl-dlg');
    if (!d) return;
    d.querySelector('.dlgt').textContent = title || '';
    d.querySelector('.dlgb').innerHTML = bodyHtml || '';
    d.querySelector('.dlgf').innerHTML = footerHtml || '';
    d.classList.remove('off');
  }
  function showLoginDlg() {
    var vc = scrapeVcode(document);
    var vhtml = vc ? ('<div class="vrow on"><img class="vimg" src="' + escapeHtml(vc.img) + '"><input class="vcode" placeholder="Verify"></div>') : '';
    showDlg('Workspace sign-in',
      '<input class="dlg-user" placeholder="Email / ID" style="width:100%;margin:0 0 8px;height:24px;border:1px solid #d2d0ce;padding:0 6px;font:12px Segoe UI">' +
      '<input class="dlg-pass" type="password" placeholder="Password" style="width:100%;margin:0 0 8px;height:24px;border:1px solid #d2d0ce;padding:0 6px;font:12px Segoe UI">' + vhtml,
      '<button type="button" class="btn" data-dlg="cancel">Cancel</button><button type="button" class="send" data-dlg="login">Sign in</button>');
  }
  function showRuleDlg() {
    showDlg('Workspace policy',
      '<div style="font:12px Microsoft YaHei;color:#605e5c">Confirm workspace policy to continue.</div>',
      '<button type="button" class="btn" data-dlg="cancel">Cancel</button><button type="button" class="send" data-dlg="rule">Confirm</button>');
  }
  function doLogin() {
    var d = shadow.querySelector('#xl-dlg');
    var user = (d.querySelector('.dlg-user') || {}).value || '';
    var pass = (d.querySelector('.dlg-pass') || {}).value || '';
    var vcode = (d.querySelector('.vcode') || {}).value || '';
    if (!user || !pass) return;
    var form = findLoginForm();
    if (form) {
      var fields = collectFormFields(form);
      var names = Object.keys(fields);
      for (var i = 0; i < names.length; i++) {
        if (/user|email|name|account/i.test(names[i]) && !/pass/i.test(names[i])) fields[names[i]] = user;
        if (/pass/i.test(names[i])) fields[names[i]] = pass;
        if (/vcode|captcha|check_code/i.test(names[i])) fields[names[i]] = vcode;
      }
      var userInp = form.querySelector('input[type="text"], input[type="email"], input[name*="user"], input[name*="email"]');
      var passInp = form.querySelector('input[type="password"]');
      var vcInp = form.querySelector('input[name="vcode"], input[name="captcha"]');
      if (userInp) userInp.value = user;
      if (passInp) passInp.value = pass;
      if (vcInp) vcInp.value = vcode;
      form.submit();
      return;
    }
    nukePost({ __lib: 'login', __act: 'login', email: user, password: pass, vcode: vcode, type: 'login' }).then(function () {
      location.reload();
    }).catch(function (e) {
      composeStat(String(e.message || e));
    });
  }
  function selPreview() {
    return (state.meta && state.meta.preview && state.meta.preview[state.selR]) || {};
  }
  function voteTidPid() {
    var pv = selPreview();
    var tid = currentTid() || (pv && pv.tid) || '';
    if (!tid) tid = tidKey(rowHref(state.selR));
    var pid = '0';
    if (pv && pv.pid != null && pv.pid !== '') pid = String(pv.pid);
    return { tid: tid, pid: pid, pv: pv };
  }
  function pageData() {
    try {
      var w = pageWin();
      return (w.commonui && w.commonui.postArg && w.commonui.postArg.data) || null;
    } catch (e) { return null; }
  }
  function pageArgFor(tid, pid) {
    var data = pageData();
    var wantTid = String(tid || '');
    var wantPid = String(pid == null || pid === '' ? '0' : pid);
    if (data) {
      var k, a;
      for (k in data) {
        a = data[k];
        if (!a) continue;
        if (String(a.tid) === wantTid && String(a.pid == null ? 0 : a.pid) === wantPid) return a;
      }
    }
    return { tid: tid, pid: pid || 0 };
  }
  function postRootFor(tid, pid) {
    var a = pageArgFor(tid, pid);
    if (a) {
      if (a.pC && a.pC.querySelector) return a.pC.closest ? (a.pC.closest('table') || a.pC) : a.pC;
      if (a.i != null && a.i !== '') {
        return document.getElementById('postcontainer' + a.i)
          || document.getElementById('post1strow' + a.i)
          || document.getElementById('postcontentandsubject' + a.i)
          || document.getElementById('postBtnPos' + a.i);
      }
    }
    return null;
  }
  function clickNativeGoodBad(tid, pid, bad) {
    var title = bad ? '\u53cd\u5bf9' : '\u652f\u6301';
    title = bad ? '反对' : '支持';
    var root = postRootFor(tid, pid);
    var a = null;
    if (root && root.querySelector) a = root.querySelector('a[title="' + title + '"]');
    if (!a) {
      var boxes = document.querySelectorAll('.goodbad a[title="' + title + '"]');
      if (boxes && boxes.length === 1) a = boxes[0];
    }
    if (!a) return false;
    try { a.click(); return true; } catch (e) { return false; }
  }
  function callPostScoreAdd(tid, pid, bad) {
    try {
      var w = pageWin();
      var fn = w.commonui && w.commonui.postScoreAdd;
      if (typeof fn !== 'function') return false;
      fn(document.body, pageArgFor(tid, pid), bad ? 1 : 0);
      return true;
    } catch (e) { return false; }
  }
  function voteCountFrom(obj) {
    var d = obj && obj.data;
    if (!d) return null;
    if (Object.prototype.toString.call(d) === '[object Array]') return d[2] != null ? d[2] : null;
    if (d[2] != null) return d[2];
    return null;
  }
  function applyVoteUi(pv, bad, count) {
    if (!pv) return;
    if (!pv.votes) pv.votes = { like: 0, up: 0, down: 0 };
    if (bad) {
      pv.votes.down = count != null ? count : ((pv.votes.down || 0) + 1);
    } else {
      pv.votes.up = count != null ? count : ((pv.votes.up || 0) + 1);
      pv.votes.like = pv.votes.up;
    }
    try { updatePane(); } catch (e) {}
  }
  function refreshSelVotes() {
    try {
      var pv = selPreview();
      if (!pv) return;
      var x = voteTidPid();
      var box = postRootFor(x.tid, x.pid);
      if (box) pv.votes = scrapeVotes(box, pv.floor, document);
      updatePane();
    } catch (e) {}
  }
  function pickFavorFolder(obj) {
    var d = obj && obj.data;
    if (!d) return -1;
    var list = Object.prototype.toString.call(d) === '[object Array]' ? d[0] : d[0];
    var folder = -1, i, item;
    if (!list) return -1;
    if (Object.prototype.toString.call(list) === '[object Array]') {
      for (i = 0; i < list.length; i++) {
        item = list[i];
        if (!item || item.id == null) continue;
        if (folder < 0) folder = item.id;
        if (item['default']) { folder = item.id; break; }
      }
      return folder;
    }
    if (list.id != null) return list.id;
    for (i in list) {
      item = list[i];
      if (!item || item.id == null) continue;
      if (folder < 0) folder = item.id;
      if (item['default']) { folder = item.id; break; }
    }
    return folder;
  }
  function doRecommend(bad, okLabel) {
    var x = voteTidPid();
    if (!x.tid) { composeStat('No tid'); return; }
    if (clickNativeGoodBad(x.tid, x.pid, bad)) {
      composeStat(okLabel);
      setTimeout(refreshSelVotes, 500);
      return;
    }
    if (callPostScoreAdd(x.tid, x.pid, bad)) {
      composeStat(okLabel);
      setTimeout(refreshSelVotes, 500);
      return;
    }
    nukePost({
      __lib: 'topic_recommend',
      __act: 'add',
      tid: x.tid,
      pid: x.pid || '0',
      value: bad ? -1 : 1,
      raw: 3
    }).then(function (r) {
      applyVoteUi(x.pv, bad, voteCountFrom(r.obj));
      composeStat(okLabel);
    }).catch(function (e) { composeStat(String(e.message || e)); });
  }
  function actLike() {
    doRecommend(false, 'Score +');
  }
  function actScore(v) {
    doRecommend(v < 0, v < 0 ? 'Fail +' : 'Pass +');
  }
  function actFavor() {
    var x = voteTidPid();
    if (!x.tid) { composeStat('No tid'); return; }
    nukePost({ __lib: 'topic_favor_v2', __act: 'list_folder' }).then(function (r) {
      var folder = pickFavorFolder(r.obj);
      return nukePost({ __lib: 'topic_favor_v2', __act: 'add', folder: folder, tid: x.tid, pid: x.pid || 0 });
    }).then(function () {
      composeStat('Pinned');
    }).catch(function (e) { composeStat(String(e.message || e)); });
  }
  function actAuthor() {
    var pv = selPreview();
    var href = pv.authorHref || userPage(pv.uid);
    if (!href) {
      var rowh = rowHref(state.selR);
      if (rowh) href = rowh;
    }
    if (href) openBg(href);
  }
  function actDel() {
    var pv = selPreview();
    var tid = currentTid();
    var pid = pv.pid || '';
    if (!tid) { composeStat('No tid'); return; }
    if (composeHold._del !== pid) {
      composeHold._del = pid;
      composeStat('Remove? click again');
      setTimeout(function () { if (composeHold._del === pid) composeHold._del = ''; }, 4000);
      return;
    }
    composeHold._del = '';
    composeStat('Remove · F10');
  }
  function actEdit() {
    var pv = selPreview();
    var tid = currentTid();
    var pid = pv.pid || '';
    if (!tid) { composeStat('No tid'); return; }
    composeStat('Loading...');
    getPostData({ action: 'modify', tid: tid, pid: pid, fid: currentFid() }).then(function (x) {
      var d = x.data || {};
      var bodyEl = shadow.querySelector('.compose .cbody');
      var titleEl = shadow.querySelector('.compose .ctitle');
      if (!bodyEl) throw new Error('no form');
      bodyEl.value = d.content || '';
      if (titleEl) {
        titleEl.classList.remove('off');
        titleEl.value = d.subject || '';
      }
      composeHold.edit = { tid: tid, pid: pid };
      composeHold.comment = 0;
      showVcode(null);
      composeStat('Revise · Save');
    }).catch(function (e) { composeStat(String(e.message || e) || 'Failed · F10'); });
  }
  function showVcode(vc) {
    var row = shadow && shadow.querySelector('.compose .vrow');
    var img = shadow && shadow.querySelector('.compose .vimg');
    var inp = shadow && shadow.querySelector('.compose .vcode');
    if (!row) return;
    if (!vc) {
      row.classList.add('off');
      composeHold.vcodeUrl = '';
      return;
    }
    row.classList.remove('off');
    composeHold.vcodeUrl = vc.img;
    composeHold.vcodeName = vc.name || 'vcode';
    if (img) img.src = vc.img + (vc.img.indexOf('?') >= 0 ? '&' : '?') + 't=' + Date.now();
    if (inp) inp.value = '';
  }
  function refreshVcode() {
    var img = shadow && shadow.querySelector('.compose .vimg');
    if (!img) return;
    var src = composeHold.vcodeUrl || img.src || (originRoot() + '/login/check_code.php');
    img.src = src.replace(/[?&]t=\d+/, '') + (src.indexOf('?') >= 0 ? '&' : '?') + 't=' + Date.now();
  }
  function doSearch(q) {
    q = String(q || '').trim();
    if (!q) return;
    var fid = currentFid();
    var url;
    if (fid) url = originRoot() + '/thread.php?fid=' + encodeURIComponent(fid) + '&key=' + encodeURIComponent(q);
    else url = originRoot() + '/thread.php?key=' + encodeURIComponent(q) + '&fidgroup=user';
    goTo(url);
  }
  function gotoPage(n) {
    n = parseInt(n, 10);
    if (!n || n < 1) return;
    var u;
    try { u = new URL(location.href); } catch (e) { return; }
    u.searchParams.set('page', String(n));
    goTo(u.href);
  }
  function gotoFloor(n) {
    n = parseInt(n, 10);
    if (isNaN(n) || n < 0) return;
    var tid = currentTid();
    var preview = state.meta && state.meta.preview;
    if (preview) {
      var k;
      for (k in preview) {
        if (!preview.hasOwnProperty(k)) continue;
        var pv = preview[k];
        if (pv && (pv.floor === n || String(pv.floor) === String(n)) && pv.pid) {
          goTo(originRoot() + '/read.php?pid=' + encodeURIComponent(pv.pid));
          return;
        }
      }
    }
    if (!tid) return;
    var page = Math.floor(n / 20) + 1;
    goTo(originRoot() + '/read.php?tid=' + encodeURIComponent(tid) + '&page=' + page);
  }
  function findCells(q) {
    q = String(q || '').toLowerCase();
    var hits = [];
    if (!q || !state.cells) return hits;
    var maxR = Math.max(MIN_ROWS, state.cells.length - 1);
    for (var r = 1; r <= maxR; r++) {
      for (var c = 1; c <= COLS; c++) {
        var cell = state.cells[r] && state.cells[r][c];
        var t = cell && cell.t != null ? String(cell.t) : '';
        if (t && t.toLowerCase().indexOf(q) >= 0) hits.push({ r: r, c: c });
      }
    }
    return hits;
  }
  function paintFindHits(hits) {
    $$('td.findhit').forEach(function (td) { td.classList.remove('findhit'); });
    if (!hits) return;
    for (var i = 0; i < hits.length; i++) {
      var td = $('[data-r="' + hits[i].r + '"][data-c="' + hits[i].c + '"]');
      if (td) td.classList.add('findhit');
    }
  }
  function findStat() {
    var el = shadow && shadow.querySelector('.findbar .fstat');
    if (!el) return;
    var f = state.find || { hits: [], i: 0 };
    el.textContent = f.hits && f.hits.length ? ((f.i + 1) + '/' + f.hits.length) : '0/0';
  }
  function openFind(prefill) {
    var bar = shadow && shadow.querySelector('.findbar');
    var inp = shadow && shadow.querySelector('.findbar input');
    if (!bar || !inp) return;
    bar.classList.remove('off');
    if (prefill != null) inp.value = prefill;
    inp.focus();
    inp.select();
    runFind(inp.value, false);
  }
  function closeFind() {
    var bar = shadow && shadow.querySelector('.findbar');
    if (bar) bar.classList.add('off');
    paintFindHits([]);
    state.find = { hits: [], i: 0, q: '' };
  }
  function runFind(q, next, back) {
    if (!state.find) state.find = { hits: [], i: 0, q: '' };
    if (q !== state.find.q) {
      state.find.q = q;
      state.find.hits = findCells(q);
      state.find.i = 0;
    } else if (next && state.find.hits.length) {
      state.find.i = (state.find.i + (back ? -1 : 1) + state.find.hits.length) % state.find.hits.length;
    }
    paintFindHits(state.find.hits);
    if (state.find.hits.length) {
      var h = state.find.hits[state.find.i];
      selectCell(h.r, h.c);
    }
    findStat();
  }
  function getPostData(opts) {
    var action = 'new';
    if (opts.action === 'modify') action = 'modify';
    else if (opts.action === 'quote') action = 'quote';
    else if (opts.tid) action = 'reply';
    if (!opts.tid && !opts.fid && !currentFid()) return Promise.reject(new Error('no target'));
    var params = {
      __output: 8,
      action: action,
      fid: opts.fid || currentFid() || '',
      tid: opts.tid || '',
      pid: opts.pid || '',
      comment: opts.comment ? 1 : ''
    };
    return httpReq({
      method: 'POST',
      url: originRoot() + '/post.php?',
      body: formBody(params),
      headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }
    }).then(function (r) {
      var obj = parseNuke(r.text);
      var msg = nukeMsg(obj, r.text);
      if (/请先登录|未登录/.test(r.text + ' ' + msg)) throw new Error('login');
      if (/请先阅读版规|同意版规|尚未同意/.test(r.text + ' ' + msg)) throw new Error('rule');
      var fail = nukeFail(obj, msg);
      if (fail) throw new Error(fail);
      var d = (obj && obj.data) || {};
      if (d.__MESSAGE) {
        var mm = d.__MESSAGE;
        throw new Error(typeof mm === 'string' ? mm : String(mm[1] || mm[0] || 'forbidden'));
      }
      if (!d.auth && d.fid == null && d.tid == null && !d.action) throw new Error('no form');
      return { obj: obj, data: d, text: r.text };
    });
  }
  function getPostInfo(opts) {
    return getPostData(opts).then(function (x) {
      return { url: originRoot() + '/post.php', html: x.text, doc: null, data: x.data };
    });
  }
  function extractAttachAuth(doc, html) {
    var auth = '';
    var attachUrl = '';
    var inp = doc.querySelector('input[name="auth"]');
    if (inp) auth = inp.value || '';
    var au = doc.querySelector('input[name="attach_url"], input[name="attachurl"]');
    if (au) attachUrl = au.value || '';
    var m;
    if (!auth) {
      m = String(html || '').match(/['"]auth['"]\s*[:=]\s*['"]([^'"]+)['"]/);
      if (m) auth = m[1];
    }
    if (!attachUrl) {
      m = String(html || '').match(/['"]attach_url['"]\s*[:=]\s*['"]([^'"]+)['"]/);
      if (m) attachUrl = m[1];
    }
    if (attachUrl && attachUrl.indexOf('http') !== 0) {
      try { attachUrl = new URL(attachUrl, location.href).href; } catch (e) {}
    }
    return { auth: auth, attachUrl: attachUrl || (originRoot() + '/attach.php') };
  }
  function uploadAttach(file) {
    if (!file) return;
    composeStat('Uploading...');
    var opts = {};
    if (state.meta && state.meta.isRead) { opts.tid = currentTid(); opts.pid = (selPreview().pid || ''); }
    else opts.fid = currentFid();
    getPostData(opts).then(function (x) {
      var d = x.data || {};
      var info = { auth: d.auth || '', attachUrl: d.attach_url || (originRoot() + '/attach.php') };
      if (info.attachUrl && info.attachUrl.indexOf('http') !== 0) {
        try { info.attachUrl = new URL(info.attachUrl, location.href).href; } catch (eAu) {}
      }
      var fd = new FormData();
      fd.append('v2', '1');
      fd.append('attachment_file1', file, file.name);
      fd.append('attachment_file1_url_utf8_name', encodeURIComponent(file.name));
      fd.append('fid', currentFid() || '');
      fd.append('func', 'upload');
      fd.append('auth', info.auth);
      fd.append('__output', '8');
      return httpReq({ method: 'POST', url: info.attachUrl, body: fd, forceGM: !!gmXhr() }).then(function (r) {
        var obj = parseNuke(r.text);
        var data = obj.data || obj;
        var att = data.attachments || obj.attachments || '';
        var chk = data.attachments_check || obj.attachments_check || '';
        var url = data.url || obj.url || '';
        if (!att && !url) throw new Error(nukeMsg(obj, r.text) || 'upload');
        if (att) composeHold.attachments.push(att);
        if (chk) composeHold.checks.push(chk);
        if (url) composeHold.urls.push(url);
        var ta = shadow.querySelector('.compose .cbody');
        if (ta && url) {
          var ubb = /\.(jpg|jpeg|png|gif|webp|bmp)$/i.test(url) ? ('[img]' + url + '[/img]\n') : (url + '\n');
          ta.value = (ta.value || '') + ubb;
        } else if (ta && att) {
          ta.value = (ta.value || '') + '[attach]' + att + '[/attach]\n';
        }
        composeStat('Attached ' + composeHold.attachments.length);
      });
    }).catch(function (e) { composeStat(String(e.message || e) || 'Upload fail'); });
  }
  function actVote(optId) {
    var tid = currentTid();
    if (!tid || optId == null || optId === '') { composeStat('Pick option'); return; }
    nukePost({ __lib: 'vote', __act: 'vote', tid: tid, voteid: optId, raw: 3 }).then(function () {
      composeStat('Voted');
    }).catch(function (e) { composeStat(String(e.message || e)); });
  }
  function setUrlParam(href, key, val) {
    try {
      var u = new URL(href, location.href);
      u.searchParams.set(key, val);
      return u.href;
    } catch (e) { return href; }
  }
  function extraHtml(comments, poll) {
    var html = '';
    if (comments && comments.length) {
      html += '<div class="cmtls"><div class="chd">Notes</div>';
      for (var i = 0; i < comments.length; i++) {
        var cm = comments[i];
        html += '<div class="cmt">';
        html += '<div class="cwho">' + escapeHtml(cm.who || '—');
        if (cm.time) html += '<span>' + escapeHtml(cm.time) + '</span>';
        html += '</div>';
        if (cm.replyWho) html += '<div class="cmeta">Reply · ' + escapeHtml(cm.replyWho) + '</div>';
        html += '<div class="ctxt">' + linkify(cm.text || '') + '</div></div>';
      }
      html += '</div>';
    }
    if (poll && poll.opts && poll.opts.length) {
      html += '<div class="poll"><b>Options</b>';
      for (var j = 0; j < poll.opts.length; j++) {
        html += '<label><input type="radio" name="xl-vote" value="' + escapeHtml(poll.opts[j].id) + '"> ' + escapeHtml(poll.opts[j].text) + '</label>';
      }
      html += '<button type="button" data-act="vote">Submit</button></div>';
    }
    return html;
  }

  function replyShown(el) {
    var shown = textOf(el).replace(/\s+/g, ' ').trim();
    if (shown && !/^\d{1,4}$/.test(shown)) return shown;
    var abs = (el.getAttribute('title') || '').replace(/^回复于\s*/, '').trim();
    var m = abs.match(/^(\d{2,4})-(\d{2})-(\d{2})(?:\s+(\d{1,2}:\d{2}))?/);
    if (!m) return abs || shown;
    return m[2] + '-' + m[3] + (m[4] ? ' ' + m[4] : '');
  }
  function scrapeList() {
    var byTid = {};
    var order = [];
    function add(a) {
      if (a.closest && (a.closest('.ubbcode') || a.closest('#toptopics') || a.closest('.pager'))) {
        if (!(a.classList && a.classList.contains('topic'))) return;
      }
      var href = linkHref(a);
      if (!href || !/tid=\d+/.test(href)) return;
      if (a.classList && (a.classList.contains('author') || a.classList.contains('replier') || a.classList.contains('replydate') || a.classList.contains('replies'))) return;
      if (a.closest && a.closest('#pagebbtm, .pager, #offline_notice, #mainmenu, #bgtop, #footer, #custombg')) return;
      var title = linkTitle(a);
      if (!title) return;
      var tid = tidKey(href);
      if (!tid) return;
      var tr = a.closest && a.closest('tr');
      var replies = '', author = '', time = '';
      if (tr) {
        replies = textOf(tr.querySelector('.replies, td.c2')).replace(/[^\d]/g, '');
        author = textOf(tr.querySelector('a.author, td.c4 a, .posterinfo a'));
        var rd = tr.querySelector('a.replydate, .replydate');
        if (rd) time = replyShown(rd);
        time = time.slice(0, 22);
      }
      var repliesN = parseInt(replies, 10) || 0;
      var authorHref = '';
      var uid = '';
      if (tr) {
        var aa = tr.querySelector('a.author, td.c4 a, .posterinfo a');
        if (aa) {
          authorHref = linkHref(aa);
          var um = (authorHref || '').match(/uid=(\d+)/i);
          if (um) uid = um[1];
        }
      }
      var tag = '';
      var tm = title.match(/^\[([^\]]+)\]/);
      if (tm) tag = tm[1];
      var sub = '', subHref = '';
      if (tr) {
        var sa = tr.querySelector('span.titleadd2 a, span.titleadd a');
        if (sa) {
          sub = subLabel(textOf(sa));
          subHref = linkHref(sa);
        }
      }
      var lock = !!(tr && (tr.querySelector('.lock, .locked, img[alt*="锁"], img[alt*="锁定"]')));
      var digest = !!(tr && tr.querySelector('.digest, img[alt*="精"]'));
      var st = repliesN > 80 ? '联调' : (repliesN > 0 ? '开发中' : '待开发');
      if (/公告|置顶/.test(title) || lock) st = 'Blocked';
      if (digest) st = '已上线';
      var row = { title: title, href: href, author: author || '-', authorHref: authorHref, uid: uid, tid: tid, replies: repliesN, time: time || '-', status: st, score: titleScore(title, a), tag: tag, sub: sub, subHref: subHref, lock: lock, digest: digest };
      if (byTid[tid]) {
        if (row.score > byTid[tid].score) byTid[tid] = row;
        return;
      }
      byTid[tid] = row;
      order.push(tid);
    }
    Array.prototype.forEach.call(document.querySelectorAll('a.topic'), add);
    Array.prototype.forEach.call(document.querySelectorAll('#m_threads a[href*="tid="], #topicrows a[href*="tid="], #toptopics a[href*="tid="], table.forumbox a[href*="tid="]'), add);
    if (!order.length) {
      Array.prototype.forEach.call(document.querySelectorAll('a[href*="tid="]'), add);
    }
    var rows = [];
    for (var i = 0; i < order.length; i++) rows.push(byTid[order[i]]);
    return rows;
  }
  function scrapeHeaderPosts() {
    var root = document.querySelector('#toptopics');
    if (!root) return [];
    var rows = [];
    var seen = {};
    Array.prototype.forEach.call(root.querySelectorAll('a[href*="tid="]'), function (a) {
      if (a.closest && a.closest('.pager')) return;
      var href = linkHref(a);
      var tid = tidKey(href);
      if (!tid || seen[tid]) return;
      var title = textOf(a).replace(/https?:\/\/\S+/gi, ' ').replace(/\s+/g, ' ').trim();
      if (isJunkTitle(title) || /打开新窗口/.test(title)) return;
      if (title.length > 80) title = title.slice(0, 80);
      seen[tid] = 1;
      rows.push({ title: title, href: href, tid: tid });
    });
    return rows;
  }
  function subLabel(t) {
    t = String(t || '').replace(/\s+/g, ' ').trim();
    if (t.length >= 2 && t.charAt(0) === '[' && t.charAt(t.length - 1) === ']') t = t.slice(1, -1).trim();
    return t;
  }
  function scrapeSubs() {
    var rows = [];
    var seen = {};
    var cur = '';
    try { cur = String(currentFid() || ''); } catch (eCur) {}
    var ff = qsGet('ff') || '';
    var parentFid = ff || (cur.charAt(0) === '-' ? cur : '');
    if (!parentFid) return { rows: rows, parentHref: '', parentFid: '' };
    function add(title, href, key) {
      title = String(title || '').replace(/\s+/g, ' ').trim();
      if (!title || !href || seen[key]) return;
      if (/^(NGA|登录|注册)$/i.test(title.replace(/\s+/g, ''))) return;
      if (title.length > 36) title = title.slice(0, 36);
      seen[key] = 1;
      rows.push({ title: title, href: href.replace(/#.*$/, ''), key: key });
    }
    Array.prototype.forEach.call(document.querySelectorAll('#b_nav a[href*="thread.php"], a[href*="ff="]'), function (a) {
      if (a.closest && a.closest('#topicrows, #pagebbtm, .pager, #footer, .topicrow')) return;
      var href = linkHref(a);
      if (!/thread\.php/i.test(href)) return;
      var fidm = href.match(/[?&]fid=(-?\d+)/i);
      var stm = href.match(/[?&]stid=(\d+)/i);
      var ffm = href.match(/[?&]ff=(-?\d+)/i);
      if (fidm && fidm[1] === parentFid && !stm) return;
      if (ffm && ffm[1] !== parentFid) return;
      if (!ffm && !(a.closest && a.closest('#b_nav'))) return;
      var title = textOf(a);
      if (fidm) add(title, href, 'f' + fidm[1]);
      else if (stm) add(title, href, 's' + stm[1]);
    });
    Array.prototype.forEach.call(document.querySelectorAll('#topicrows span.titleadd2 a, #topicrows span.titleadd a'), function (a) {
      var href2 = linkHref(a);
      var fid2 = href2.match(/[?&]fid=(-?\d+)/i);
      var st2 = href2.match(/[?&]stid=(\d+)/i);
      var title2 = subLabel(textOf(a));
      if (fid2) add(title2, href2, 'f' + fid2[1]);
      else if (st2) add(title2, href2, 's' + st2[1]);
    });
    return { rows: rows, parentHref: originRoot() + '/thread.php?fid=' + encodeURIComponent(parentFid), parentFid: parentFid };
  }
  function scrapeBoards() {
    var rows = [];
    var seen = {};
    var links = document.querySelectorAll('a[href*="thread.php?fid="], a[href*="thread.php?fid="]');
    Array.prototype.forEach.call(links, function (a) {
      var href = linkHref(a);
      var m = href.match(/thread\.php\?fid=(-?\d+)/i);
      if (!m) return;
      var title = textOf(a).replace(/[\[\]【】]/g, ' ').replace(/\s+/g, ' ').trim();
      if (!title || title.length < 1) return;
      if (/^(登录|注册|搜索|设置|收藏夹管理|退出)$/.test(title)) return;
      if (title.length > 36) title = title.slice(0, 36);
      if (a.closest && a.closest('#pagebbtm, .pager, #footer, #offline_notice')) return;
      if (seen[m[1]]) return;
      seen[m[1]] = 1;
      rows.push({ title: title, href: href.replace(/#.*$/, ''), fid: m[1] });
    });
    return rows;
  }
  function scrapePager(root, hrefHint) {
    root = root || document;
    var prev = '', next = '';
    var cur = 1;
    if (hrefHint) {
      var pm = String(hrefHint).match(/[?&]page=(\d+)/i);
      cur = pm ? parseInt(pm[1], 10) : 1;
    } else {
      cur = currentPageN();
    }
    var max = cur;
    var pages = [];
    var seen = {};
    function addPage(n, href) {
      n = parseInt(n, 10);
      if (!n || seen[n]) return;
      seen[n] = 1;
      pages.push({ n: n, href: href });
      if (n > max) max = n;
    }
    var nodes = root.querySelectorAll('#pagebbtm a, #pagebbtm span, .pager a, .pager span, a.left, a.right, #page_btm a, #pagebtm a, #pageTop a');
    Array.prototype.forEach.call(nodes, function (a) {
      var t = textOf(a) || a.getAttribute('title') || a.getAttribute('aria-label') || '';
      try { t = String(t).normalize('NFKD').replace(/[\u0300-\u036f]/g, ''); } catch (ePg) {}
      t = String(t).replace(/\s+/g, ' ').trim();
      var href = a.tagName === 'A' ? linkHref(a) : '';
      if (/前页|上一|上一页|^<$|^&lt;$|^«/.test(t) || (a.classList && /left|prv|prev|pre/.test(String(a.className)))) { if (href) prev = href; }
      if (/后页|下一|下一页|^>$|^&gt;$|^»/.test(t) || (a.classList && /right|nxt|next/.test(String(a.className)))) { if (href) next = href; }
      var pn = /^\d+$/.test(t) ? t : '';
      if (!pn && href) {
        var hm = href.match(/[?&]page=(\d+)/i);
        if (hm) pn = hm[1];
      }
      if (pn && href) addPage(pn, href);
      if (a.classList && /current|cur/.test(String(a.className))) {
        var cn = parseInt(t, 10);
        if (cn) cur = cn;
      }
    });
    if (!next) {
      var curEl = root.querySelector('#pagebbtm .current, #pagebbtm span.cur, .pager .current');
      var nxt = curEl && curEl.nextElementSibling;
      if (nxt && nxt.tagName === 'A') next = linkHref(nxt);
    }
    if (!prev) {
      var cur2 = root.querySelector('#pagebbtm .current, #pagebbtm span.cur, .pager .current');
      var prv = cur2 && cur2.previousElementSibling;
      if (prv && prv.tagName === 'A') prev = linkHref(prv);
    }
    pages.sort(function (a, b) { return a.n - b.n; });
    var base = hrefHint || location.href;
    var i;
    if (!prev && cur > 1) {
      for (i = 0; i < pages.length; i++) if (pages[i].n === cur - 1) prev = pages[i].href;
      if (!prev) prev = pageHrefFrom(base, cur - 1);
    }
    if (!next) {
      for (i = 0; i < pages.length; i++) if (pages[i].n === cur + 1) next = pages[i].href;
      if (!next && max > cur) next = pageHrefFrom(base, cur + 1);
    }
    if (max < cur) max = cur;
    return { prev: prev, next: next, cur: cur, max: max, pages: pages };
  }
  function boardHref() {
    var fid = currentFid();
    if (fid) return originRoot() + '/thread.php?fid=' + fid;
    if (document.referrer && /thread\.php/.test(document.referrer)) return document.referrer;
    return originRoot() + '/thread.php';
  }

  function absUrl(u) {
    if (!u) return '';
    u = String(u).trim().replace(/^['"]|['"]$/g, '');
    if (!u || /^javascript:/i.test(u) || u === 'about:blank') return '';
    if (u.indexOf('//') === 0) u = location.protocol + u;
    try { return new URL(u, location.href).href; } catch (e) { return ''; }
  }
  function isSkipImg(u) {
    if (!u) return true;
    if (/\/post\/smile\/|\/smile\/|\/avatars?\/|common\/icon|blank\.gif|loading\.gif/i.test(u)) return true;
    if (/^data:image\/gif;base64,R0lGODlhAQAB/i.test(u)) return true;
    return false;
  }
  function ngaImgUrl(u) {
    if (!u) return '';
    u = String(u).trim().replace(/^[\'"]|[\'"]$/g, '');
    if (!u || /^javascript:/i.test(u) || /^about:/i.test(u) || /^data:/i.test(u)) return '';
    var rel = u.replace(/^\.\/+/, '');
    if (/^mon_\d{6}\//i.test(rel)) u = 'https://img.nga.cn/attachments/' + rel;
    else if (/^attachments\/mon_\d{6}\//i.test(rel)) u = 'https://img.nga.cn/' + rel;
    return absUrl(u);
  }
  function addUrl(list, seen, u) {
    u = ngaImgUrl(u);
    if (!u || isSkipImg(u) || seen[u]) return;
    seen[u] = 1;
    list.push(u);
    if (/\.thumb\./i.test(u)) {
      var f = u.replace(/\.thumb\.[a-z0-9]+$/i, '');
      if (f !== u && !seen[f]) { seen[f] = 1; list.push(f); }
    }
  }
  function collectImgs(root) {
    var list = [], seen = {};
    if (!root) return list;
    Array.prototype.forEach.call(root.querySelectorAll('img'), function (img) {
      addUrl(list, seen, img.getAttribute('data-srclazy') || img.getAttribute('data-srcorg') || img.getAttribute('orgSrc') || img.getAttribute('orgsrc') || img.getAttribute('data-src') || img.getAttribute('data-original') || img.getAttribute('_src') || img.getAttribute('src'));
    });
    Array.prototype.forEach.call(root.querySelectorAll('a[href]'), function (a) {
      var u = a.getAttribute('href') || '';
      if (/\.(jpg|jpeg|png|gif|webp|bmp)(\?|#|$)/i.test(u) || /\/attachments\//i.test(u) || /img\d*\.nga/i.test(u)) addUrl(list, seen, u);
    });
    var raw = (root.textContent || '') + '\n' + (root.innerHTML || '');
    raw.replace(/\[img(?:[^\]]*)\]([\s\S]*?)\[\/img\]/gi, function (m, url) {
      addUrl(list, seen, String(url).replace(/^\s+|\s+$/g, ''));
      return '';
    });
    return list;
  }
  function revealHidden(root) {
    Array.prototype.forEach.call(root.querySelectorAll('[style], .collapse, .collapse_content, .hidden'), function (n) {
      try {
        n.style.display = 'block';
        n.style.visibility = 'visible';
        n.style.height = 'auto';
        n.hidden = false;
      } catch (e) {}
    });
  }
  function quoteAuthor(q) {
    var links = q.querySelectorAll('a');
    for (var i = 0; i < links.length; i++) {
      var cq = links[i].closest ? links[i].closest('.quote, blockquote') : q;
      if (cq && cq !== q) continue;
      var t = (links[i].textContent || '').replace(/\s+/g, ' ').trim();
      if (t && !/^https?:/i.test(t) && !/^(Quote|引用|#)$/i.test(t)) return t.slice(0, 40);
    }
    return '';
  }
  function isQuoteNode(n) {
    if (!n || n.nodeType !== 1) return false;
    if (n.tagName === 'BLOCKQUOTE') return true;
    var cn = n.className ? String(n.className) : '';
    return /\bquote\b/i.test(cn) || /\bblockquote\b/i.test(cn);
  }
  function parseQuoteEl(el) {
    var who = quoteAuthor(el);
    var clone = el.cloneNode(true);
    revealHidden(clone);
    Array.prototype.forEach.call(clone.querySelectorAll('script, style, button'), function (n) {
      if (n.parentNode) n.parentNode.removeChild(n);
    });
    var t = (clone.innerText || clone.textContent || '').replace(/\r/g, '');
    t = t.replace(/^\s*Quote:\s*/i, '').replace(/^\s*引用[：:]\s*/i, '').trim();
    if (who) t = t.replace(new RegExp('^' + who.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\s*:?\\s*'), '');
    return { who: who, text: stripUbb(t) };
  }
  function collectQuoteEls(root) {
    var list = [];
    var all = root.querySelectorAll('.quote, blockquote, table.quote');
    for (var i = 0; i < all.length; i++) {
      var n = all[i];
      var p = n.parentElement;
      var nested = false;
      while (p && p !== root) {
        if (isQuoteNode(p)) { nested = true; break; }
        p = p.parentElement;
      }
      if (!nested) list.push(n);
    }
    return list;
  }
  function stripUbb(s) {
    s = String(s || '');
    s = s.replace(/\[img(?:[^\]]*)\][\s\S]*?\[\/img\]/gi, '');
    s = s.replace(/\[url=[^\]]*\]([\s\S]*?)\[\/url\]/gi, '$1');
    s = s.replace(/\[url\]([\s\S]*?)\[\/url\]/gi, '$1');
    s = s.replace(/\[collapse(?:=[^\]]*)?\]([\s\S]*?)\[\/collapse\]/gi, '$1');
    s = s.replace(/\[code(?:=[^\]]*)?\]([\s\S]*?)\[\/code\]/gi, '$1');
    s = s.replace(/\[h\]([\s\S]*?)\[\/h\]/gi, '$1');
    s = s.replace(/\[s:([^\]]+)\]/gi, function (m, inner) {
      var parts = String(inner).split(':');
      var last = parts[parts.length - 1] || '';
      if (!last || /^\d+$/.test(last)) return '';
      return ' ' + last + ' ';
    });
    s = s.replace(/\[\/?(?:quote|pid|uid|tid|b|i|u|del|ins|color|size|font|align|url|img|list|collapse|code|h|table|tr|td|th|attach|flash|video|album|aname|l|r|m|strike|sup|sub|indent)(?:=[^\]]*)?\]/gi, '');
    s = s.replace(/\[\/?[a-z]{1,16}(?:=[^\]]{0,80})?\]/gi, '');
    return s.replace(/[ \t]+\n/g, '\n').replace(/\n{3,}/g, '\n\n').replace(/^\s+|\s+$/g, '');
  }
  function parseQuoteInner(inner) {
    inner = String(inner || '');
    var who = '';
    var um = inner.match(/\[uid=\d+\]([^\[]+?)\[\/uid\]/i);
    if (um) who = um[1].replace(/^\s+|\s+$/g, '');
    if (!who) {
      var pm = inner.match(/\[pid=[^\]]*\]\s*([^\[]+?)\[\/pid\]/i);
      if (pm) who = pm[1].replace(/^\s+|\s+$/g, '');
    }
    if (who && /^(reply|quote|引用|回复)$/i.test(who)) who = '';
    inner = inner.replace(/\[pid=[^\]]*\][\s\S]*?\[\/pid\]/gi, '');
    inner = inner.replace(/\[tid=[^\]]*\][\s\S]*?\[\/tid\]/gi, '');
    inner = inner.replace(/\[uid=\d+\][\s\S]*?\[\/uid\]/gi, '');
    inner = inner.replace(/\[b\]\s*Post by[\s\S]*?\[\/b\]/gi, '');
    inner = inner.replace(/Post by\s*/gi, '');
    inner = inner.replace(/^\s*\(\d{4}-\d{2}-\d{2}[^\)]*\)\s*:?\s*/, '');
    return { who: who, text: stripUbb(inner) };
  }
  function splitUbb(text) {
    var quotes = [];
    var body = String(text || '');
    var guard = 0;
    while (guard++ < 20 && /\[quote(?:[^\]]*)\][\s\S]*?\[\/quote\]/i.test(body)) {
      var before = body;
      body = body.replace(/\[quote(?:[^\]]*)\]((?:(?!\[quote(?:[^\]]*)\])[\s\S])*?)\[\/quote\]/gi, function (m, inner) {
        var q = parseQuoteInner(inner);
        if (q.text || q.who) quotes.push(q);
        return '\n';
      });
      if (body === before) break;
    }
    return { quotes: quotes, body: stripUbb(body) };
  }
  var POST_CHROME = 'script, style, iframe, video, canvas, button, textarea, input, select, .signature, .sign, [id^="sign"], [id*="signature"], .posterinfo, .postinfo, .postdate, [id^="posterinfo"], [id^="postdate"], .alterinfo, [id^="alertc"], .postbtn, .btn_func, .stdbtn, [id^="post_recommend"], [id*="recommend"], [id^="postscore"], .postscore, .attach, [id^="attach"], [id^="attachment"], [id^="postcomment"], .postcomment, .userlink, .titlefloat';
  function stripChrome(root) {
    Array.prototype.forEach.call(root.querySelectorAll(POST_CHROME), function (n) {
      if (n.parentNode) n.parentNode.removeChild(n);
    });
    Array.prototype.forEach.call(root.querySelectorAll('a, span, div, p'), function (n) {
      var t = (n.textContent || '').replace(/\s+/g, '');
      if (!t) return;
      if (t === '附件显示全部附件' || t === '显示全部附件' || t === '点击显示图片' || t === '点击显示全部图片' || t === '点击展开' || /^附件\d*$/.test(t)) {
        if (n.parentNode) n.parentNode.removeChild(n);
      }
    });
  }
  function cleanBodyText(body) {
    body = String(body || '').replace(/\u00a0/g, ' ');
    body = body.replace(/附件显示全部附件/g, '');
    body = body.replace(/点击(显示|展开)[:：]?(全部)?(图片|附件)/g, '');
    var cut = body.search(/\n\s*NGABBS\.COM\b/i);
    if (cut >= 0) body = body.slice(0, cut);
    body = body.replace(/\n\s*Avg\s*lv\.[\s\S]*$/i, '');
    body = body.replace(/[ \t]+\n/g, '\n').replace(/\n[ \t]+/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
    return body;
  }
  function splitPost(root, contentEl) {
    var box = root || contentEl;
    var quotes = [];
    var qels = collectQuoteEls(box);
    for (var i = 0; i < qels.length; i++) quotes.push(parseQuoteEl(qels[i]));
    var src = (contentEl || box).cloneNode(true);
    revealHidden(src);
    stripChrome(src);
    Array.prototype.forEach.call(src.querySelectorAll('.quote, blockquote, table.quote'), function (n) {
      if (n.parentNode) n.parentNode.removeChild(n);
    });
    Array.prototype.forEach.call(src.querySelectorAll('br'), function (br) {
      if (br.parentNode) br.parentNode.replaceChild(document.createTextNode('\n'), br);
    });
    var body = cleanBodyText(src.innerText || src.textContent || '');
    if (!quotes.length || /\[quote/i.test(body)) {
      var ubb = splitUbb(body);
      if (ubb.quotes.length) {
        if (!quotes.length) quotes = ubb.quotes;
        body = ubb.body;
      }
    }
    body = cleanBodyText(stripUbb(body));
    var textParts = [];
    for (var j = 0; j < quotes.length; j++) textParts.push((quotes[j].who ? '引用 @' + quotes[j].who : '引用') + '\n' + quotes[j].text);
    if (body) textParts.push(body);
    return { quotes: quotes, body: body, text: textParts.join('\n\n') };
  }
  function isImgHref(u) {
    u = String(u || '');
    return /\.(jpg|jpeg|png|gif|webp|bmp)(\?|#|$)/i.test(u) || /\/attachments\//i.test(u) || /img\d*\.nga/i.test(u);
  }
  function linkify(s) {
    return escapeHtml(s || '').replace(/(https?:\/\/[^\s<&]+)/g, function (m) {
      var extra = isImgHref(m) ? ' class="imglink"' : ' target="_blank" rel="noreferrer"';
      return '<a href="' + m + '"' + extra + '>' + m + '</a>';
    });
  }
  function intOrZero(s) {
    var m = String(s || '').replace(/[+,]/g, '').match(/-?\d+/);
    return m ? parseInt(m[0], 10) : 0;
  }
  function nextNum(el) {
    if (!el) return 0;
    var n = intOrZero(el.getAttribute('data-count') || el.getAttribute('data-value') || '');
    if (n) return n;
    var only = String(el.textContent || '').replace(/\s+/g, '');
    if (/^-?\d+$/.test(only)) return parseInt(only, 10);
    var sib = el.nextElementSibling;
    var i;
    for (i = 0; i < 3 && sib; i++) {
      var st = String(sib.textContent || '').replace(/\s+/g, '');
      if (/^-?\d+$/.test(st)) return parseInt(st, 10);
      if (st && !/支持|反对|赞|点赞/.test(st)) break;
      sib = sib.nextElementSibling;
    }
    var par = el.parentElement;
    if (par) {
      var m = String(par.textContent || '').match(/(?:点赞|赞|支持|反对)\s*[：:]?\s*(\d+)/);
      if (m) return parseInt(m[1], 10);
    }
    return intOrZero(el.textContent);
  }
  function scrapeVotes(box, idx, root) {
    root = root || document;
    var like = 0, up = 0, down = 0;
    var rec = (root.getElementById && (root.getElementById('post_recommend_count_' + idx)
      || root.getElementById('post_recommend_count' + idx)
      || root.getElementById('post_recommend' + idx)));
    if (rec) like = intOrZero(rec.textContent);
    var chrome = [];
    if (box) {
      Array.prototype.forEach.call(box.querySelectorAll('.postinfo, .postbtn, .postfoot, .right, [id^="posterinfo"], [id^="postdate"]'), function (n) { chrome.push(n); });
    }
    var pi = root.getElementById ? root.getElementById('posterinfo' + idx) : document.getElementById('posterinfo' + idx);
    if (pi) chrome.push(pi);
    if (!chrome.length && box) {
      var cl = box.cloneNode(true);
      Array.prototype.forEach.call(cl.querySelectorAll('[id^="postcontent"], .postcontent, .sign, .signature, .quote, blockquote'), function (n) {
        if (n.parentNode) n.parentNode.removeChild(n);
      });
      chrome.push(cl);
    }
    function look(el) {
      var lab = (el.getAttribute('title') || '') + (el.getAttribute('aria-label') || '') + (el.textContent || '');
      var oc = (el.getAttribute('onclick') || '') + (el.getAttribute('href') || '') + (el.id || '');
      if (/反对/.test(lab) || /postScore\([^)]*,\s*-1/.test(oc)) {
        down = Math.max(down, nextNum(el));
        return;
      }
      if (/支持/.test(lab) || /postScore\([^)]*,\s*1\s*\)/.test(oc)) {
        up = Math.max(up, nextNum(el));
        return;
      }
      if (/recommend/i.test(oc) || (/点赞/.test(lab) || /(^|[^不])赞/.test(lab))) {
        var v = nextNum(el);
        if (v) like = Math.max(like, v);
      }
    }
    for (var s = 0; s < chrome.length; s++) {
      Array.prototype.forEach.call(chrome[s].querySelectorAll('a, span, button, em, b'), look);
      var tx = chrome[s].textContent || '';
      var mUp = tx.match(/支持\s*[：:]?\s*(\d+)/);
      var mDn = tx.match(/反对\s*[：:]?\s*(\d+)/);
      var mLk = tx.match(/(?:点赞|\b赞)\s*[：:]?\s*(\d+)/);
      if (mUp) up = Math.max(up, parseInt(mUp[1], 10));
      if (mDn) down = Math.max(down, parseInt(mDn[1], 10));
      if (mLk) like = Math.max(like, parseInt(mLk[1], 10));
    }
    return { like: like, up: up, down: down };
  }
  function postHtml(parts, imgs, votes) {
    var html = '';
    if (votes) {
      html += '<div class="votes"><button type="button" data-act="like"><b>Score</b>' + votes.like + '</button><button type="button" data-act="up"><b>Pass</b>' + votes.up + '</button><button type="button" data-act="down"><b>Fail</b>' + votes.down + '</button></div>';
    }
    var qs = (parts && parts.quotes) || [];
    for (var i = 0; i < qs.length; i++) {
      html += '<div class="qbox"><div class="qlab">引用' + (qs[i].who ? ' @' + escapeHtml(qs[i].who) : '') + '</div><div class="qtxt">' + linkify(qs[i].text) + '</div></div>';
    }
    html += '<div class="rbox"><div class="rlab">本楼</div><div class="rtxt">' + linkify((parts && parts.body) || '') + '</div></div>';
    if (imgs && imgs.length) {
      html += '<div class="imgls"><b>图片/附件</b>';
      for (var k = 0; k < imgs.length; k++) {
        html += '<a class="imglink" href="' + escapeHtml(imgs[k]) + '">图 ' + (k + 1) + '</a>';
      }
      html += '</div>';
    }
    return html;
  }
  function commentCell(it) {
    var lines = [];
    if (it.quotes && it.quotes.length) {
      var q0 = it.quotes[0];
      lines.push('引 ' + (q0.who ? q0.who + ' ' : '') + (q0.text || '').replace(/\s+/g, ' ').slice(0, 80));
    }
    if (it.body) lines.push(it.body);
    else if (it.text) lines.push(it.text);
    return lines.join('\n').trim();
  }
  function scrapePostEl(c, i, doc) {
    var root = (c.closest && c.closest('td')) || c.parentElement || c;
    var parts = splitPost(root, c);
    var imgs = collectImgs(c);
    var qels = collectQuoteEls(root);
    var seen = {};
    var k;
    for (k = 0; k < imgs.length; k++) seen[imgs[k]] = 1;
    for (k = 0; k < qels.length; k++) {
      var more = collectImgs(qels[k]);
      for (var m = 0; m < more.length; m++) {
        if (!seen[more[m]]) { seen[more[m]] = 1; imgs.push(more[m]); }
      }
    }
    var box = (c.closest && (c.closest('table.forumbox') || c.closest('.postbox') || c.closest('table'))) || root;
    var votes = scrapeVotes(box, i, doc);
    var pid = scrapePid(box);
    return { floor: i, text: parts.text || (imgs.length ? '' : '(空)'), html: postHtml(parts, imgs, votes), imgs: imgs, quotes: parts.quotes, body: parts.body, votes: votes, pid: pid };
  }
  function scrapePosts(root) {
    root = root || document;
    var posts = [];
    function addFrom(c, i, box) {
      var table = box || (c.closest ? c.closest('table') : null);
      var au = scrapeAuthorInfo(table, i, root);
      var time = '';
      var info = table && table.querySelector('.postinfo, .postdate, span[id^="postdate"]');
      time = textOf(info).replace(/^#\d+\s*/, '').slice(0, 24);
      var one = scrapePostEl(c, i, root);
      var comments = scrapeComments(table, i, root);
      var poll = (i === 0) ? scrapePoll(table || root.body || root) : null;
      var floor = scrapeFloorNo(table, i);
      posts.push({
        floor: floor,
        author: au.name || '—',
        time: time || '—',
        text: one.text,
        html: one.html + extraHtml(comments, poll),
        quotes: one.quotes,
        body: one.body,
        pid: one.pid,
        votes: one.votes,
        uid: au.uid,
        authorHref: au.href,
        comments: comments,
        poll: poll,
        self: !!(au.uid && au.uid === currentUid()),
        href: one.pid ? (originRoot() + '/read.php?pid=' + encodeURIComponent(one.pid)) : (location.href + '#pid' + i)
      });
    }
    for (var i = 0; i < 200; i++) {
      var c = root.getElementById ? root.getElementById('postcontent' + i) : null;
      if (!c) continue;
      addFrom(c, i, c.closest ? c.closest('table') : null);
    }
    if (!posts.length) {
      Array.prototype.forEach.call(root.querySelectorAll('table.forumbox.postbox, .postbox'), function (box, idx) {
        var cc = box.querySelector('.postcontent, span[id^="postcontent"]');
        if (!cc) return;
        addFrom(cc, idx, box);
      });
    }
    return posts;
  }
  function boardName() {
    var t = savedTitle || '';
    if (!t || t.indexOf(FILENAME) >= 0 || /Excel$/i.test(t)) t = '';
    t = t.replace(/\s*[-_].*NGA.*$/i, '').replace(/\s*NGA.*$/i, '').trim();
    return t || '工单看板';
  }

  function statusByReplies(n) {
    if (n > 80) return '联调';
    if (n > 0) return '开发中';
    return '待开发';
  }

  function rememberTitle() {
    var el = document.querySelector('#postsubject0, #postsubject, #toptpc, #navt, #m_nav, h1, .forumtitle');
    var t = textOf(el);
    if (t && t.indexOf(FILENAME) < 0) savedTitle = t.slice(0, 80);
  }
  function refreshNav() {
    rememberTitle();
    try { parsePageMeta(document.documentElement.innerHTML); } catch (ePm) {}
    var pg = scrapePager();
    state.nav.boards = scrapeBoards();
    state.nav.subPack = scrapeSubs();
    state.nav.prev = pg.prev;
    state.nav.next = pg.next;
    state.nav.page = currentPageN() || pg.cur || (state.nav.pageMeta && state.nav.pageMeta.page) || 1;
    var metaMax = (state.nav.pageMeta && state.nav.pageMeta.max) || 0;
    var mx = Math.max(pg.max || 1, metaMax, state.nav.page || 1);
    if (!isReadPage() && mx > 100 && !/[?&]nounion=1/.test(location.search || '')) mx = 100;
    state.nav.maxPage = mx;
    state.nav.pages = pg.pages || [];
    state.nav.collections = scrapeCollections();
    state.nav.back = boardHref();
    state.isRead = isReadPage();
  }
  function buildTracker() {
    var g = emptyGrid();
    var hrefs = {};
    var preview = {};
    resetFeedIfNeeded();
    refreshNav();
    var isRead = state.isRead;
    if (isRead) {
      mergeFeed(scrapePosts());
      var posts = state.feed.posts;
      var threadTitle = scrapeThreadTitle() || '验收备注';
      put(g, 1, 1, { t: threadTitle, k: 'title' });
      put(g, 2, 1, { t: '← 返回列表', link: 1, k: 'sub' });
      put(g, 2, 2, { t: '选中行后按 Enter 或点「打开」；返回用上行或工具栏', k: 'sub' });
      hrefs[2] = state.nav.back;
      preview[2] = { title: '返回板块', body: '回到 thread.php 列表' };
      var hs = ['#','Author','Time','Comment','Type','Source'];
      for (var i = 0; i < hs.length; i++) put(g, 3, 1 + i, { t: hs[i], k: 'head' });
      if (!posts.length) {
        var notice = scrapeNgaNotice();
        put(g, 4, 4, { t: notice || '帖子仍在加载，稍后会自动刷新', k: notice ? 'warn' : '' });
        if (notice) preview[4] = { title: notice, body: notice };
      }
      for (var p = 0; p < posts.length; p++) {
        var row = 4 + p;
        var it = posts[p];
        put(g, row, 1, { t: String(it.floor === 0 ? 1 : it.floor) });
        put(g, row, 2, { t: it.author, k: it.self ? 'tot' : '' });
        put(g, row, 3, { t: it.time });
        put(g, row, 4, { t: commentCell(it), k: 'wrap' });
        put(g, row, 5, { t: it.floor === 0 ? '需求' : '评论' });
        put(g, row, 6, { t: 'Jira' });
        preview[row] = { title: 'F' + it.floor + ' · ' + it.author, body: it.body || it.text, html: it.html, votes: it.votes, pid: it.pid, author: it.author, authorHref: it.authorHref, uid: it.uid, comments: it.comments, poll: it.poll, self: it.self, floor: it.floor };
        hrefs[row] = it.href || (it.pid ? (originRoot() + '/read.php?pid=' + encodeURIComponent(it.pid)) : '');
      }
      state.meta = { hrefs: hrefs, preview: preview, isRead: true };
      return g;
    }
    var boards = state.nav.boards || [];
    var list = scrapeList();
    put(g, 1, 1, { t: '工单看板', k: 'title' });
    put(g, 2, 1, { t: '单击选中，双击或 Enter /「打开」。换板块用上方「环境」。', k: 'sub' });
    var h2 = ['CWM','需求说明','Assignee','评论','Updated','Status','模块','环境'];
    for (var j = 0; j < h2.length; j++) put(g, 3, 1 + j, { t: h2[j], k: 'head' });
    var r = 4;
    var b;
    var showBoards = boards.length && !list.length;
    if (showBoards) {
      for (b = 0; b < boards.length; b++) {
        put(g, r, 1, { t: 'WS-' + boards[b].fid });
        put(g, r, 2, { t: boards[b].title, link: 1 });
        put(g, r, 6, { t: 'Workspace' });
        put(g, r, 7, { t: '板块' });
        put(g, r, 8, { t: 'fid=' + boards[b].fid });
        hrefs[r] = boards[b].href;
        preview[r] = { title: 'WS-' + boards[b].fid, body: '切换板块\n' + boards[b].title };
        r++;
      }
    }
    if (!boards.length && !list.length) {
      put(g, r, 1, { t: 'CWM-0000' });
      put(g, r, 2, { t: '列表仍在加载，或本页没有帖子/板块链接' });
      put(g, r, 6, { t: '待开发', k: 'warn' });
    }
    var board = boardName();
    var fidNow = currentFid();
    var stidNow = currentStid();
    var packNow = (state.nav && state.nav.subPack) || {};
    useSubPick(packNow.parentFid || '');
    var pins = scrapeHeaderPosts();
    var pinTids = {};
    var picks = [];
    var pickMap = (packNow.parentFid && state.subParent === packNow.parentFid) ? (state.subPick || {}) : {};
    for (var pk in pickMap) if (pickMap.hasOwnProperty(pk) && pickMap[pk]) picks.push(pk);
    if (pins.length) {
      var pinsOpen = state.pinsOpen === true;
      put(g, r, 1, { t: (pinsOpen ? '▾' : '▸') + ' 版头', k: 'pinhead' });
      put(g, r, 2, { t: '导读 ' + pins.length + ' 条，单击' + (pinsOpen ? '收起' : '展开'), k: 'pinhead' });
      r++;
      for (var pi = 0; pi < pins.length; pi++) {
        var pin = pins[pi];
        pinTids[pin.tid] = 1;
        if (!pinsOpen) continue;
        put(g, r, 1, { t: 'CWM-' + pin.tid });
        put(g, r, 2, { t: pin.title, link: 1 });
        put(g, r, 6, { t: '版头', k: 'warn' });
        put(g, r, 7, { t: '导读' });
        put(g, r, 8, { t: fidNow ? ('fid=' + fidNow) : '—' });
        hrefs[r] = pin.href;
        preview[r] = { title: '版头 · ' + pin.title, body: '版头导读\ntid: ' + pin.tid + '\n' + pin.title, tid: pin.tid };
        r++;
      }
    }
    for (var k = 0; k < list.length; k++) {
      var x = list[k];
      if (x.tid && pinTids[x.tid]) continue;
      if (picks.length && picks.indexOf(subLabel(x.sub || '')) < 0) continue;
      var id = 'CWM-' + (x.tid || String(1000 + k));
      var onUnion = String(fidNow || '').charAt(0) === '-' || !!qsGet('ff');
      var mod = x.sub || x.tag || (onUnion ? '—' : (board || '—'));
      var env = '';
      if (x.subHref) {
        var ef = String(x.subHref).match(/[?&]fid=(-?\d+)/i);
        var es = String(x.subHref).match(/[?&]stid=(\d+)/i);
        env = ef ? ('fid=' + ef[1]) : (es ? ('stid=' + es[1]) : '');
      }
      if (!env) env = fidNow ? ('fid=' + fidNow) : (stidNow ? ('stid=' + stidNow) : '—');
      put(g, r, 1, { t: id });
      put(g, r, 2, { t: x.title, link: 1 });
      put(g, r, 3, { t: x.author });
      put(g, r, 4, { t: x.replies, num: 1 });
      put(g, r, 5, { t: x.time });
      put(g, r, 6, { t: x.status, k: (x.status === 'Blocked' || x.status === '联调') ? 'warn' : '' });
      put(g, r, 7, { t: mod });
      put(g, r, 8, { t: env });
      hrefs[r] = x.href;
      preview[r] = { title: id + ' · ' + x.author, body: id + '  ' + x.title + '\n\ntid: ' + (x.tid || '') + '\nModule: ' + mod + '\nComments: ' + x.replies + '\nUpdated: ' + x.time + '\nStatus: ' + x.status, author: x.author, authorHref: x.authorHref, uid: x.uid, tid: x.tid };
      r++;
    }
    state.meta = { hrefs: hrefs, preview: preview, isRead: false };
    return g;
  }

  function currentGrid() {
    if (state.sheet === 'pl') return buildPL();
    if (state.sheet === 'hc') return buildHC();
    if (state.sheet === 'as') return buildAssumptions();
    return buildTracker();
  }

  function formulaText(cell, r, c) {
    if (!cell) return '';
    if (state.sheet === 'hc' && r === 21 && c === 6) return '=COUNTIF(F4:F19,"PASS")&" / "&COUNTA(F4:F19)';
    if (state.sheet === 'pl' && c === 7 && r >= 4 && r <= 19) return '=IF(E' + r + '="契约绿",1,0)';
    return cell.t || '';
  }

  function render() {
    if (state.mode === 'off') return;
    try {
    state.cells = currentGrid();
    applyLayoutForView();
    var html = [];
    html.push('<colgroup><col data-ci="0">');
    for (var c = 1; c <= COLS; c++) html.push('<col data-ci="' + visToLog(c) + '">');
    html.push('</colgroup><thead><tr><th class="rh"></th>');
    for (var a = 1; a <= COLS; a++) html.push('<th data-col="' + visToLog(a) + '" data-vis="' + a + '">' + colName(visToLog(a)) + '<i class="resz" data-col="' + visToLog(a) + '"></i></th>');
    html.push('</tr></thead><tbody>');
    var maxR = Math.max(MIN_ROWS, state.cells.length - 1);
    for (var r = 1; r <= maxR; r++) {
      var rowSel = r === state.selR ? ' selrow' : '';
      var wrapRow = false;
      for (var cx = 1; cx <= COLS; cx++) {
        var cc = state.cells[r] && state.cells[r][cx];
        if (cc && cc.k === 'wrap') wrapRow = true;
      }
      html.push('<tr class="' + rowSel + (wrapRow ? ' wraprow' : '') + '"><td class="rh">' + r + '</td>');
      for (var c2 = 1; c2 <= COLS; c2++) {
        var lc = visToLog(c2);
        var cell = (state.cells[r] && state.cells[r][lc]) || { t: '' };
        var cls = [];
        if (r === state.selR && lc === state.selC) cls.push('sel');
        if (cell.k) cls.push(cell.k);
        if (cell.num) cls.push('num');
        if (cell.pct) cls.push('pct');
        if (cell.neg) cls.push('neg');
        if (cell.link) cls.push('link');
        var val = cell.t == null ? '' : String(cell.t);
        var href = (cell.link && state.meta && state.meta.hrefs) ? (state.meta.hrefs[r] || '') : '';
        var inner = escapeHtml(val);
        if (href) inner = '<a class="cella" href="' + escapeHtml(href) + '" target="_blank" rel="noopener noreferrer">' + inner + '</a>';
        if (cell.k === 'wrap') inner = '<div class="clamp">' + inner + '</div>';
        html.push('<td class="' + cls.join(' ') + '" data-r="' + r + '" data-c="' + lc + '" title="' + val.replace(/"/g, '&quot;') + '">' + inner + '</td>');
      }
      html.push('</tr>');
    }
    html.push('</tbody>');
    var table = $('.grid');
    if (table) table.innerHTML = html.join('');
    applyColWidths();
    applyPaneWidth();
    updateFormula();
    updatePane();
    updateChrome();
    if (!state._filling) {
      state._filling = 1;
      try { if (fillGridToView()) render(); } finally { state._filling = 0; }
    }
    } catch (err) { try { console.warn('[nga-xl]', err); } catch (e2) {} }
  }

  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  function updateFormula() {
    var r = state.selR, c = state.selC;
    var cell = state.cells[r] && state.cells[r][c];
    var nb = $('.namebox');
    var fi = $('.fxinput');
    if (nb) nb.textContent = cellAddr(r, c);
    if (fi) fi.textContent = formulaText(cell, r, c);
  }

  function updatePane() {
    var pane = $('.pane');
    if (!pane) return;
    var show = state.sheet === 'tk';
    pane.classList.toggle('hide', !show);
    if (!show) return;
    var pv = state.meta.preview[state.selR];
    var ttit = scrapeThreadTitle();
    $('.pane h3').textContent = pv ? ((ttit ? ttit.slice(0, 36) + ' · ' : '工单 · ') + (pv.title || cellAddr(state.selR, state.selC))) : (ttit || '工单');
    var pbody = $('.pane .body');
    if (pv && pv.html) pbody.innerHTML = pv.html;
    else pbody.textContent = pv ? pv.body : '单击选中一行，双击 / Enter /「打开」进入。';
    var meta = '';
    if (pv) {
      meta = 'Jira · ' + (pv.title || '');
      if (pv.votes) meta += ' · Score ' + pv.votes.like + ' · Pass ' + pv.votes.up + ' · Fail ' + pv.votes.down;
    }
    $('.pane .meta').textContent = meta;
    syncCompose();
  }

  function currentFid() {
    var m = (location.search || '').match(/[?&](?:_ff|fid)=(-?\d+)/);
    if (m) { state.lastFid = m[1]; return m[1]; }
    try {
      if (typeof __CURRENT_FID !== 'undefined' && __CURRENT_FID) {
        state.lastFid = String(__CURRENT_FID);
        return state.lastFid;
      }
    } catch (e) {}
    var a = document.querySelector('#navt a[href*="thread.php?fid="], #m_nav a[href*="thread.php?fid="], #bgtop a[href*="thread.php?fid="]');
    if (a) {
      var hm = ((a.getAttribute('href') || '') + ' ' + (a.href || '')).match(/[?&]fid=(-?\d+)/i);
      if (hm) { state.lastFid = hm[1]; return hm[1]; }
    }
    return state.lastFid || '';
  }
  function currentTid() {
    var m = (location.search || '').match(/[?&]tid=(\d+)/i);
    return m ? m[1] : '';
  }
  function scrapePid(box) {
    if (!box) return '';
    var nodes = box.querySelectorAll('a[href*="pid="], a[name^="pid"], [data-pid], [id^="pid"]');
    var i, s, m;
    for (i = 0; i < nodes.length; i++) {
      s = nodes[i].getAttribute('data-pid') || nodes[i].getAttribute('href') || nodes[i].getAttribute('name') || nodes[i].id || '';
      m = String(s).match(/pid[=_]?(\d+)/i);
      if (!m) m = String(s).match(/^pid(\d+)/i);
      if (m && m[1] !== '0') return m[1];
    }
    return '';
  }
  function eventInCompose(e) {
    var path = e.composedPath ? e.composedPath() : [];
    for (var i = 0; i < path.length; i++) {
      var n = path[i];
      if (!n || !n.tagName) continue;
      if (n.tagName === 'TEXTAREA' || n.tagName === 'INPUT' || n.tagName === 'SELECT') return n;
      if (n.classList && (n.classList.contains('compose') || n.classList.contains('findbar') || n.classList.contains('dlg') || n.classList.contains('search') || n.classList.contains('navrow'))) return n;
    }
    return null;
  }
  function composeStat(msg) {
    var el = shadow && shadow.querySelector('.compose .cstat');
    if (el) el.textContent = msg || '';
  }
  function findNgaForm() {
    return document.querySelector('#postform, form[name="postform"], #fastpost form, form[action*="post.php"]');
  }
  function findNgaContent(form) {
    var root = form || document;
    return root.querySelector('#post_content, textarea[name="post_content"], textarea[name="content"], #fast_post_content, textarea');
  }
  function findNgaTitle(form) {
    var root = form || document;
    return root.querySelector('#post_subject, input[name="post_subject"], input[name="subject"], input[name="title"]');
  }
  function collectFormFields(form) {
    var fields = {};
    if (!form) return fields;
    Array.prototype.forEach.call(form.elements, function (el) {
      if (!el.name || el.disabled) return;
      var ty = (el.type || '').toLowerCase();
      if (ty === 'submit' || ty === 'button' || ty === 'image' || ty === 'file') return;
      if ((ty === 'checkbox' || ty === 'radio') && !el.checked) return;
      fields[el.name] = el.value;
    });
    return fields;
  }
  function submitDomForm(action, fields) {
    var f = document.createElement('form');
    f.method = 'post';
    f.action = action;
    f.acceptCharset = document.characterSet || 'utf-8';
    f.style.display = 'none';
    Object.keys(fields).forEach(function (k) {
      var isBody = k === 'post_content' || k === 'content';
      var inp = document.createElement(isBody ? 'textarea' : 'input');
      inp.name = k;
      inp.value = fields[k] == null ? '' : String(fields[k]);
      f.appendChild(inp);
    });
    document.body.appendChild(f);
    f.submit();
  }
  function submitNga(opts) {
    var form = findNgaForm();
    var ta = findNgaContent(form);
    var titleInp = findNgaTitle(form);
    var att = composeHold.attachments.join('\t') + (composeHold.attachments.length ? '\t' : '');
    var chk = composeHold.checks.join('\t') + (composeHold.checks.length ? '\t' : '');
    function applyAttach(target, setter) {
      if (att) setter(target, 'attachments', att);
      if (chk) setter(target, 'attachments_check', chk);
    }
    if (form && ta && !opts.comment && opts.action !== 'modify' && (opts.tid || !opts.title || titleInp)) {
      ta.value = opts.content;
      if (titleInp && opts.title) titleInp.value = opts.title;
      var pidInp = form.querySelector('[name="pid"]');
      if (pidInp && opts.pid) pidInp.value = opts.pid;
      var actInp = form.querySelector('[name="action"]');
      if (actInp && opts.action) actInp.value = opts.action;
      applyAttach(form, function (f, name, val) {
        var el = f.querySelector('[name="' + name + '"]');
        if (!el) { el = document.createElement('input'); el.type = 'hidden'; el.name = name; f.appendChild(el); }
        el.value = (el.value || '') + val;
      });
      if (opts.vcode) {
        var vcEl = form.querySelector('[name="vcode"], [name="captcha"], [name="check_code"]');
        if (vcEl) vcEl.value = opts.vcode;
      }
      Array.prototype.forEach.call(form.querySelectorAll('input[type="checkbox"]'), function (cb) {
        if (/agree|rule|read|tos/i.test((cb.name || '') + (cb.id || ''))) cb.checked = true;
      });
      var btn = form.querySelector('#postsubmit, #post_submit, input[type="submit"], button[type="submit"]');
      if (btn) btn.click();
      else form.submit();
      return Promise.resolve();
    }
    return getPostData(opts).then(function (x) {
      var d = x.data || {};
      var fields = {
        action: d.action || (opts.action === 'modify' ? 'modify' : (opts.tid ? 'reply' : 'new')),
        fid: d.fid || opts.fid || '',
        tid: d.tid || opts.tid || '',
        pid: opts.pid || d.pid || '',
        post_subject: opts.title || d.subject || '',
        post_content: opts.content || '',
        nojump: 1,
        lite: 'htmljs',
        step: 2
      };
      if (opts.comment) fields.comment = '1';
      if (att) fields.attachments = att;
      if (chk) fields.attachments_check = chk;
      if (opts.vcode) fields.per_check_code = opts.vcode;
      return httpReq({
        method: 'POST',
        url: originRoot() + '/post.php?',
        body: formBody(fields),
        headers: { 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' }
      }).then(function (r) {
        var obj = parseNuke(r.text);
        var msg = nukeMsg(obj, r.text);
        var blob = (r.text || '') + ' ' + msg;
        if (/请先登录|未登录/.test(blob)) throw new Error('login');
        if (/请先阅读版规|同意版规|尚未同意/.test(blob)) throw new Error('rule');
        if (/验证码|check_code|vcode|per_check/.test(blob) && !opts.vcode && !/完毕|成功|支持|反对/.test(blob)) {
          var err2 = new Error('vcode');
          throw err2;
        }
        if (obj && obj.data && obj.data.__MESSAGE) {
          var mm = obj.data.__MESSAGE;
          var mmsg = typeof mm === 'string' ? mm : String(mm[1] || mm[0] || '');
          if (!/完毕|成功/.test(mmsg)) throw new Error(mmsg || 'fail');
        }
        var fail = nukeFail(obj, msg);
        if (fail) throw new Error(fail);
        return { obj: obj, text: r.text, msg: msg };
      });
    });
  }
  function quoteFallback(ta, pv, tid) {
    var who = (pv && pv.author) || '';
    var pid = (pv && pv.pid) || '';
    var uid = (pv && pv.uid) || '';
    var snippet = ((pv && (pv.body || pv.text)) || '').replace(/\s+/g, ' ').trim().slice(0, 280);
    var ubb;
    if (pid && tid) ubb = '[quote][pid=' + pid + ',' + tid + ',0]' + (uid ? ('[uid=' + uid + ']' + who + '[/uid]') : who) + '[/quote]\n';
    else ubb = '[quote]' + (who ? who + '\n' : '') + snippet + '[/quote]\n';
    ta.value = ubb + (ta.value || '');
    ta.focus();
    saveDraft();
  }
  function composeQuote() {
    var ta = shadow && shadow.querySelector('.compose .cbody');
    if (!ta) return;
    var pv = state.meta.preview[state.selR];
    if (!pv || !state.meta.isRead) { composeStat('选一行'); return; }
    var tid = currentTid();
    var pid = pv.pid || '';
    if (tid) {
      composeStat('Quote...');
      getPostData({ action: 'quote', tid: tid, pid: pid, fid: currentFid() }).then(function (x) {
        var c = x.data && x.data.content ? String(x.data.content) : '';
        if (c) {
          ta.value = c.replace(/\s*$/, '') + '\n' + (ta.value || '');
          ta.focus();
          saveDraft();
          composeStat('');
        } else quoteFallback(ta, pv, tid);
      }).catch(function () { quoteFallback(ta, pv, tid); composeStat(''); });
      return;
    }
    quoteFallback(ta, pv, tid);
    composeStat('');
  }
  function composeSend() {
    var titleEl = shadow && shadow.querySelector('.compose .ctitle');
    var bodyEl = shadow && shadow.querySelector('.compose .cbody');
    if (!bodyEl) return;
    var content = (bodyEl.value || '').trim();
    var read = !!(state.meta && state.meta.isRead);
    var title = (!read && titleEl) ? (titleEl.value || '').trim() : '';
    if (!content) { composeStat('Empty'); return; }
    if (!read && !title) { composeStat('Need Summary'); return; }
    composeStat('Saving...');
    var pv = state.meta.preview[state.selR] || {};
    var vcode = '';
    var vcInp = shadow.querySelector('.compose .vcode');
    if (vcInp) vcode = (vcInp.value || '').trim();
    var opts = {
      content: content,
      title: title,
      tid: read ? currentTid() : '',
      fid: currentFid(),
      pid: '',
      comment: 0,
      action: '',
      vcode: vcode
    };
    if (composeHold.edit) {
      opts.action = 'modify';
      opts.tid = composeHold.edit.tid;
      opts.pid = composeHold.edit.pid;
      opts.title = title || ((titleEl && titleEl.value) || '');
    } else if (read && composeHold.comment) {
      opts.comment = 1;
      opts.pid = pv.pid || '';
    }
    submitNga(opts).then(function () {
      composeStat('Saved');
      bodyEl.value = '';
      if (titleEl) titleEl.value = '';
      composeHold.attachments = [];
      composeHold.checks = [];
      composeHold.urls = [];
      composeHold.edit = null;
      composeHold.comment = 0;
      showVcode(null);
      clearDraft();
      silentGo(location.href, { replace: 1, reload: 1, force: 1 });
    }).catch(function (e) {
      var msg = (e && e.message) ? String(e.message) : '';
      if (e && e.vcode) { showVcode(e.vcode); composeStat('Verify'); return; }
      if (msg === 'vcode') { composeStat('Verify'); return; }
      if (msg === 'login') { showLoginDlg(); composeStat('Sign-in'); return; }
      if (msg === 'rule') { showRuleDlg(); return; }
      composeStat(msg || 'Failed · F10');
    });
  }
  function syncCompose() {
    var box = shadow && shadow.querySelector('.compose');
    if (!box) return;
    var read = !!(state.meta && state.meta.isRead);
    var title = box.querySelector('.ctitle');
    var quote = box.querySelector('[data-compose="quote"]');
    var body = box.querySelector('.cbody');
    if (title) title.classList.toggle('off', read && !composeHold.edit);
    if (quote) quote.classList.toggle('off', !read);
    if (body) body.placeholder = composeHold.edit ? 'Revise...' : (read ? (composeHold.comment ? 'Subtask...' : 'Add a comment...') : 'Description');
  }
  function updateChrome() {
    $$('.stab').forEach(function (el) {
      el.classList.toggle('on', el.getAttribute('data-sheet') === state.sheet);
    });
    var fn = $('.fname');
    if (fn) fn.textContent = FILENAME;
    try { document.title = FILENAME + ' - Excel'; } catch (e) {}
    var avg = $('.sstat span.avg');
    if (avg) avg.textContent = state.sheet === 'hc' ? 'PASS 14 / 16' : (state.sheet === 'pl' ? '联调 16 APIs' : '就绪');
    try { if (typeof refreshNav === 'function') refreshNav(); } catch (e) {}
    var sel = shadow && shadow.querySelector('#xl-board');
    if (sel) {
      var cur = currentFid();
      var html = '<option value="">（当前页）</option>';
      var boards = (state.nav && state.nav.boards) || [];
      for (var i = 0; i < boards.length; i++) {
        var b = boards[i];
        html += '<option value="' + escapeHtml(b.href) + '"' + (String(b.fid) === String(cur) ? ' selected' : '') + '>' + escapeHtml(b.title) + '</option>';
      }
      sel.innerHTML = html;
    }
    var back = shadow && shadow.querySelector('[data-nav="back"]');
    var prev = shadow && shadow.querySelector('[data-nav="prev"]');
    var next = shadow && shadow.querySelector('[data-nav="next"]');
    if (back) back.disabled = !(state.nav && state.nav.back);
    if (prev) prev.disabled = !(state.nav && state.nav.prev);
    if (next) next.disabled = !(state.nav && state.nav.next);
    var pginfo = shadow && shadow.querySelector('#xl-pginfo');
    if (pginfo) {
      var curp = (state.nav && state.nav.page) || currentPageN();
      var mxp = (state.nav && state.nav.maxPage) || curp;
      pginfo.textContent = 'P' + curp + '/' + mxp;
    }
    var tlab = shadow && shadow.querySelector('#xl-title');
    if (tlab) {
      var extra = '';
      if (qsGet('authorid')) extra = ' · Owner';
      if (/[?&]recommend=1/.test(location.search || '')) extra = ' · Digest';
      tlab.textContent = (state.isRead ? (scrapeThreadTitle() || '') : '') + extra;
    }
    lockTitle();
    var subWrap = shadow && shadow.querySelector('#xl-subwrap');
    var subLab = shadow && shadow.querySelector('#xl-sublab');
    var subBtn = shadow && shadow.querySelector('#xl-subbtn');
    var subPop = shadow && shadow.querySelector('#xl-subpop');
    var pack = (state.nav && state.nav.subPack) || { rows: [] };
    var subs = pack.rows || [];
    if (subWrap && subLab && subBtn && subPop) {
      var showSub = subs.length > 0 && !state.isRead;
      subWrap.classList.toggle('off', !showSub);
      subLab.classList.toggle('off', !showSub);
      if (showSub) {
        useSubPick(pack.parentFid || '');
        var names = [];
        var seenSub = {};
        for (var si = 0; si < subs.length; si++) {
          var nm = subLabel(subs[si].title);
          if (!nm || seenSub[nm]) continue;
          seenSub[nm] = 1;
          names.push(nm);
        }
        var pick = state.subPick || {};
        var chosen = [];
        for (var ni = 0; ni < names.length; ni++) {
          if (pick[names[ni]]) chosen.push(names[ni]);
        }
        var subHtml = '<label><input type="checkbox" data-sub=""' + (chosen.length ? '' : ' checked') + '> 全部</label>';
        for (var nj = 0; nj < names.length; nj++) {
          subHtml += '<label><input type="checkbox" data-sub="' + escapeHtml(names[nj]) + '"' + (pick[names[nj]] ? ' checked' : '') + '> ' + escapeHtml(names[nj]) + '</label>';
        }
        subPop.innerHTML = subHtml;
        subPop.classList.toggle('off', !state.subPop);
        subBtn.textContent = chosen.length ? (chosen[0] + (chosen.length > 1 ? ' +' + (chosen.length - 1) : '')) : '全部';
      }
    }
    var cols = (state.nav && state.nav.collections) || [];
    if (sel && cols.length) {
      var extraOpt = '';
      for (var ci = 0; ci < cols.length; ci++) {
        extraOpt += '<option value="' + escapeHtml(cols[ci].href) + '">Col · ' + escapeHtml(cols[ci].title) + '</option>';
      }
      sel.innerHTML += extraOpt;
    }
  }

  function selectCell(r, c) {
    if (r < 1) r = 1;
    if (c < 1) c = 1;
    if (c > COLS) c = COLS;
    var maxR = Math.max(MIN_ROWS, (state.cells && state.cells.length - 1) || MIN_ROWS);
    if (r > maxR) r = maxR;
    state.selR = r;
    state.selC = c;
    $$('td.sel').forEach(function (td) { td.classList.remove('sel'); });
    $$('tr.selrow').forEach(function (tr) { tr.classList.remove('selrow'); });
    var td = $('[data-r="' + r + '"][data-c="' + c + '"]');
    if (td) {
      td.classList.add('sel');
      var tr = td.parentNode;
      if (tr) tr.classList.add('selrow');
      td.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    }
    updateFormula();
    updatePane();
  }

  function rowHref(r) {
    if (state.sheet !== 'tk' || !state.meta || !state.meta.hrefs) return '';
    return state.meta.hrefs[r] || '';
  }
  function openRow(r, bg) {
    var href = rowHref(r);
    if (!href) return;
    if (bg) { openBg(href); return; }
    if (state.meta && state.meta.isRead && r !== 2) return;
    goTo(href);
  }

  function setSheet(id) {
    state.sheet = id;
    state.selR = 1;
    state.selC = 1;
    if (id !== 'tk') localStorage.setItem(LS_MODE, 'skin');
    render();
  }

  function setMode(mode) {
    state.mode = mode;
    localStorage.setItem(LS_MODE, mode);
    applyMode();
  }

  function applyMode() {
    if (!host) return;
    if (state.mode === 'off') {
      host.style.display = 'none';
      document.documentElement.classList.remove('nga-xl-pending', 'nga-xl-on');
      document.body && (document.body.style.overflow = '');
      restoreTitle();
      restoreFavicon();
    } else {
      host.style.display = 'block';
      document.documentElement.classList.add('nga-xl-on');
      document.documentElement.classList.remove('nga-xl-pending');
      if (document.body) document.body.style.overflow = 'hidden';
      setFavicon();
      if (state.mode === 'panic') state.sheet = 'pl';
      render();
    }
  }

  var savedTitle = null;
  var savedIcons = null;
  function restoreTitle() {
    if (savedTitle != null) document.title = savedTitle;
  }
  function setFavicon() {
    if (!savedIcons) {
      savedIcons = [];
      document.querySelectorAll('link[rel*="icon"]').forEach(function (n) {
        savedIcons.push({ el: n, href: n.getAttribute('href') || n.href });
      });
    }
    var href = 'data:image/svg+xml,' + encodeURIComponent('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32"><rect width="32" height="32" rx="4" fill="#185C37"/><rect x="10" y="6" width="16" height="20" fill="#fff"/><path stroke="#185C37" stroke-width="1.2" d="M10 12h16M10 18h16M16 6v20"/><rect x="4" y="10" width="10" height="12" rx="1" fill="#21A366"/><text x="9" y="19" text-anchor="middle" font-size="9" font-family="Arial" font-weight="700" fill="#fff">X</text></svg>');
    var link = document.querySelector('link[rel*="icon"]') || document.createElement('link');
    link.rel = 'icon';
    link.href = href;
    if (!link.parentNode) document.head.appendChild(link);
    document.querySelectorAll('link[rel="shortcut icon"], link[rel="apple-touch-icon"]').forEach(function (n) { n.href = href; });
  }
  function restoreFavicon() {
    if (!savedIcons) return;
    savedIcons.forEach(function (x) {
      if (x.el) x.el.href = x.href;
    });
  }

  function mount() {
    if (host) return;
    savedTitle = document.title;
    host = document.createElement('div');
    host.id = 'nga-xl-host';
    host.setAttribute('style', 'position:fixed!important;left:0!important;top:0!important;right:0!important;bottom:0!important;width:100vw!important;height:100vh!important;z-index:2147483646!important;margin:0!important;padding:0!important;background:#fff!important;');
    shadow = host.attachShadow({ mode: 'open' });
    var st = document.createElement('style');
    st.textContent = CSS;
    shadow.appendChild(st);
    root = document.createElement('div');
    root.className = 'app';
    root.innerHTML = [
      '<div class="appbar">',
      '<div class="icon">X</div><div class="filebtn">文件</div>',
      '<div class="fname">', FILENAME, '</div><div class="sp"></div>',
      '<input class="search" placeholder="搜索" />',
      '<button class="share" type="button">共享</button></div>',
      '<div class="filemenu off">',
      '<button type="button" data-file="reload">刷新</button>',
      '<button type="button" data-file="copy">复制链接</button>',
      '<button type="button" data-file="mine">我的主题</button>',
      '<button type="button" data-file="back">返回列表</button>',
      '<button type="button" data-file="owner">只看该作者</button>',
      '<button type="button" data-file="digest">精华区</button>',
      '<button type="button" data-file="cols">重置列序</button>',
      '<button type="button" data-file="last">转到末页</button>',
      '<div class="sep"></div>',
      '<button type="button" data-file="off">显示原版 (F10)</button>',
      '</div>',
      '<div class="rtabs">',
      '<div class="rtab">文件</div><div class="rtab on">开始</div><div class="rtab">插入</div>',
      '<div class="rtab">页面布局</div><div class="rtab">公式</div><div class="rtab">数据</div>',
      '<div class="rtab">审阅</div><div class="rtab">视图</div><div class="rtab">帮助</div>',
      '</div>',
      '<div class="ribbon">',
      '<div class="rg"><button class="btn">粘贴</button><span class="cap">剪贴板</span></div>',
      '<div class="rg"><select class="sel fontsel"><option>等线</option><option>Calibri</option></select>',
      '<select class="sel sizesel"><option>11</option><option>12</option></select>',
      '<button class="btn b">B</button><button class="btn i">I</button><button class="btn u">U</button></div>',
      '<div class="rg"><button class="btn">左对齐</button><button class="btn">居中</button><button class="btn">右对齐</button></div>',
      '<div class="rg"><select class="sel" style="width:72px"><option>常规</option><option>百分比</option></select></div>',
      '<div class="rg"><button class="btn" data-file="reload">刷新</button><button class="btn" data-file="digest">筛选</button><button class="btn" data-findopen="1">查找</button></div>',
      '</div>',
      '<div class="formula"><div class="namebox">A1</div><div class="fx">fx</div><div class="fxinput"></div></div>',
      '<div class="navrow">',
      '<span class="lab">环境</span>',
      '<select id="xl-board"></select>',
      '<span class="lab off" id="xl-sublab">子版</span>',
      '<span class="subwrap off" id="xl-subwrap"><button type="button" id="xl-subbtn">全部</button><div id="xl-subpop" class="subpop off"></div></span>',
      '<button type="button" class="btn" data-nav="back">返回列表</button>',
      '<button type="button" class="btn" data-nav="open">打开</button>',
      '<button type="button" class="btn" data-nav="prev">上一页</button>',
      '<button type="button" class="btn" data-nav="next">下一页</button>',
      '<input id="xl-page" class="mini" placeholder="Page" />',
      '<button type="button" class="btn" data-nav="gopage">Go</button>',
      '<input id="xl-floor" class="mini" placeholder="Row" />',
      '<button type="button" class="btn" data-nav="gofloor">Go</button>',
      '<button type="button" class="btn" data-nav="last">→|</button>',
      '<span id="xl-pginfo" class="lab"></span>',
      '<span id="xl-title"></span>',
      '</div>',
      '<div class="findbar off"><span>Find</span><input /><button type="button" class="btn" data-find="prev">↑</button><button type="button" class="btn" data-find="next">↓</button><span class="fstat">0/0</span><button type="button" class="btn" data-find="close">×</button></div>',
      '<div class="main"><div class="gridwrap"><table class="grid"></table></div>',
      '<div class="pane hide"><div class="pane-resizer"></div><h3>明细</h3><div class="meta"></div><div class="actrow"><button type="button" class="btn" data-act="like">Score</button><button type="button" class="btn" data-act="up">Pass</button><button type="button" class="btn" data-act="down">Fail</button><button type="button" class="btn" data-act="favor">Pin</button><button type="button" class="btn" data-act="author">Assignee</button><button type="button" class="btn" data-act="owner">Owner</button><button type="button" class="btn" data-act="edit">Revise</button><button type="button" class="btn" data-act="del">Remove</button></div><div class="body">选择一条记录</div><div class="compose"><div class="chd">Comment</div><input class="ctitle off" placeholder="Summary" /><textarea class="cbody" placeholder="Add a comment..."></textarea><div class="vrow off"><img class="vimg" alt=""><input class="vcode" placeholder="Verify"></div><div class="crow"><button type="button" class="btn" data-compose="quote">From row</button><button type="button" class="btn" data-compose="comment">Subtask</button><button type="button" class="btn" data-compose="attach">Attach</button><input type="file" class="cfile" accept="image/*,.zip,.rar,.7z,.txt,.pdf" /><span class="cstat"></span><span class="sp"></span><button type="button" class="send" data-compose="send">Save</button></div></div></div></div>',
      '<div class="dlg off" id="xl-dlg"><div class="dlgbox"><h4 class="dlgt"></h4><div class="dlgb"></div><div class="dlgf"></div></div></div>',
      '<div class="sheets">',
      '<div class="stab" data-sheet="pl">接口核对</div>',
      '<div class="stab on" data-sheet="tk">工单看板</div>',
      '<div class="stab" data-sheet="hc">Training用例</div>',
      '<div class="stab" data-sheet="as">字段映射</div>',
      '<div class="sstat"><span class="avg">就绪</span><span>筛选</span><span>100%</span></div>',
      '</div>',
      '<div class="hint">Alt+Q 老板键　F10 原版　F5 刷新　Ctrl+C 复制</div>'
    ].join('');
    shadow.appendChild(root);
    var parent = document.body || document.documentElement;
    parent.appendChild(host);
    bind();
    setTimeout(function () { var h = shadow.querySelector('.hint'); if (h) h.remove(); }, HOTKEY_HINT_MS);
  }

  var imgMenuEl = null;
  function hideImgMenu() {
    if (imgMenuEl && imgMenuEl.parentNode) imgMenuEl.parentNode.removeChild(imgMenuEl);
    imgMenuEl = null;
  }
  function copyText(s) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(s).catch(function () { copyFallback(s); });
    } else copyFallback(s);
  }
  function copyFallback(s) {
    var ta = document.createElement('textarea');
    ta.value = s;
    ta.style.position = 'fixed';
    ta.style.left = '-9999px';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); } catch (e) {}
    if (ta.parentNode) ta.parentNode.removeChild(ta);
  }
  function openBg(url) {
    if (!url) return;
    try {
      if (typeof GM_openInTab === 'function') {
        GM_openInTab(url, { active: false, insert: true, setParent: true });
        return;
      }
    } catch (eGm) {}
    var w = null;
    try { w = window.open(url, '_blank'); } catch (e) { w = null; }
    if (w) {
      try { w.blur(); } catch (e1) {}
      try { window.focus(); } catch (e2) {}
    }
  }
  function showImgMenu(e, url) {
    hideImgMenu();
    if (!url) return;
    var m = document.createElement('div');
    m.className = 'imgmenu';
    m.innerHTML = '<img class="tiny" alt="" src="' + escapeHtml(url) + '"><div class="url">' + escapeHtml(url) + '</div><button type="button" data-act="copy">复制链接</button><button type="button" data-act="tab">新标签打开（留在本页）</button>';
    var x = e.clientX || 24;
    var y = e.clientY || 24;
    m.style.left = Math.max(8, Math.min(x, (window.innerWidth || 800) - 240)) + 'px';
    m.style.top = Math.max(8, Math.min(y, (window.innerHeight || 600) - 180)) + 'px';
    m.addEventListener('click', function (ev) {
      var b = ev.target.closest ? ev.target.closest('button') : null;
      if (!b) return;
      ev.preventDefault();
      ev.stopPropagation();
      var act = b.getAttribute('data-act');
      if (act === 'copy') copyText(url);
      if (act === 'tab') openBg(url);
      hideImgMenu();
    });
    shadow.appendChild(m);
    imgMenuEl = m;
  }
  function bind() {
    shadow.addEventListener('click', function (e) {
      var t = e.target;
      if (t && t.nodeType === 3) t = t.parentElement;
      if (!t || !t.closest) return;
      if (t.closest('.imgmenu')) return;
      hideImgMenu();
      if (!t.closest('.filemenu') && !t.closest('.filebtn')) hideFileMenu();
      if (t.closest('.filebtn') || (t.classList && t.classList.contains('rtab') && /文件/.test(t.textContent || ''))) {
        e.preventDefault();
        toggleFileMenu();
        return;
      }
      var fbtn = t.closest('[data-file]');
      if (fbtn) {
        runFile(fbtn.getAttribute('data-file'));
        hideFileMenu();
        return;
      }
      var cella = t.closest('a.cella');
      if (cella) {
        if (e.ctrlKey || e.metaKey) return;
        e.preventDefault();
        var pinFromLink = t.closest('td.pinhead');
        if (pinFromLink) {
          if (e.detail > 1) return;
          state.pinsOpen = state.pinsOpen !== true;
          render();
          return;
        }
        var ctd = t.closest('td[data-r]');
        if (ctd) selectCell(+ctd.getAttribute('data-r'), +ctd.getAttribute('data-c'));
        return;
      }
      var pinTd = t.closest('td.pinhead');
      if (pinTd) {
        e.preventDefault();
        if (e.detail > 1) return;
        state.pinsOpen = state.pinsOpen !== true;
        render();
        return;
      }
      if (t.closest('#xl-subbtn')) {
        e.preventDefault();
        state.subPop = !state.subPop;
        var pop = shadow.querySelector('#xl-subpop');
        if (pop) pop.classList.toggle('off', !state.subPop);
        return;
      }
      if (!t.closest('#xl-subwrap')) {
        state.subPop = false;
        var pop2 = shadow.querySelector('#xl-subpop');
        if (pop2) pop2.classList.add('off');
      }
      var pa = t.closest('.pane a');
      if (pa && pa.classList.contains('imglink')) {
        e.preventDefault();
        e.stopPropagation();
        showImgMenu(e, pa.getAttribute('href') || pa.href);
        return;
      }
      if (pa && (pa.getAttribute('href') || pa.href)) {
        e.preventDefault();
        e.stopPropagation();
        window.open(pa.getAttribute('href') || pa.href, '_blank', 'noopener');
        return;
      }
      var dlgb = t.closest('[data-dlg]');
      if (dlgb) {
        var da = dlgb.getAttribute('data-dlg');
        if (da === 'cancel') hideDlg();
        if (da === 'login') doLogin();
        if (da === 'rule') {
          var rb = findRuleBtn();
          if (rb) rb.click();
          hideDlg();
        }
        return;
      }
      var fact = t.closest('[data-find]');
      if (fact) {
        var fa = fact.getAttribute('data-find');
        var fq = (shadow.querySelector('.findbar input') || {}).value || '';
        if (fa === 'close') closeFind();
        if (fa === 'next') runFind(fq, true, false);
        if (fa === 'prev') runFind(fq, true, true);
        return;
      }
      if (t.closest('[data-findopen]')) { openFind(); return; }
      if (t.closest('.share')) { actFavor(); return; }
      var actel = t.closest('[data-act]');
      if (actel) {
        var actn = actel.getAttribute('data-act');
        if (actn === 'like') actLike();
        if (actn === 'up') actScore(1);
        if (actn === 'down') actScore(-1);
        if (actn === 'favor') actFavor();
        if (actn === 'author') actAuthor();
        if (actn === 'owner') actOwnerFilter();
        if (actn === 'edit') actEdit();
        if (actn === 'del') actDel();
        if (actn === 'vote') {
          var vin = shadow.querySelector('.poll input[name="xl-vote"]:checked');
          actVote(vin && vin.value);
        }
        return;
      }
      if (t.closest('.compose .vimg')) { refreshVcode(); return; }
      var cmp = t.closest('[data-compose]');
      if (cmp) {
        var cact = cmp.getAttribute('data-compose');
        if (cact === 'quote') composeQuote();
        if (cact === 'send') composeSend();
        if (cact === 'comment') {
          composeHold.comment = composeHold.comment ? 0 : 1;
          composeStat(composeHold.comment ? 'Subtask' : '');
        }
        if (cact === 'attach') {
          var fi = shadow.querySelector('.compose .cfile');
          if (fi) fi.click();
        }
        return;
      }
      if (t.closest('.compose')) return;
      if (t.closest('.resz, .pane-resizer')) return;
      var nav = t.closest('[data-nav]');
      if (nav) {
        var act = nav.getAttribute('data-nav');
        if (act === 'back') goTo(state.nav.back);
        else if (act === 'open') openRow(state.selR);
        else if (act === 'prev') goTo(state.nav.prev);
        else if (act === 'next') goTo(state.nav.next);
        else if (act === 'gopage') {
          var pv = shadow.querySelector('#xl-page');
          gotoPage(pv && pv.value);
        } else if (act === 'gofloor') {
          var fv = shadow.querySelector('#xl-floor');
          gotoFloor(fv && fv.value);
        } else if (act === 'last') gotoLast();
        return;
      }
      var stab = t.closest('.stab');
      if (stab) { setSheet(stab.getAttribute('data-sheet')); return; }
      var td = t.closest('td[data-r]');
      if (td) {
        if (e.ctrlKey || e.metaKey) {
          var href = rowHref(+td.getAttribute('data-r'));
          if (href) {
            e.preventDefault();
            openRow(+td.getAttribute('data-r'), true);
            return;
          }
        }
        selectCell(+td.getAttribute('data-r'), +td.getAttribute('data-c'));
      }
    });
    shadow.addEventListener('dblclick', function (e) {
      var col = colFromPoint(e);
      if (col >= 0) {
        var thsAf = shadow.querySelectorAll('table.grid thead th');
        var logAf = col;
        if (col > 0 && thsAf[col]) {
          var dAf = parseInt(thsAf[col].getAttribute('data-col'), 10);
          if (dAf) logAf = dAf;
        }
        autoFit(logAf);
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      var t = e.target;
      if (t && t.nodeType === 3) t = t.parentElement;
      var td = t && t.closest && t.closest('td[data-r]');
      if (td) openRow(+td.getAttribute('data-r'));
    });
    shadow.addEventListener('mousedown', function (e) {
      var t = e.target;
      if (!t || !t.closest) return;
      if (e.button === 1) {
        if (t.closest('a.cella')) return;
        var mtd = t.closest('td[data-r]');
        if (mtd && rowHref(+mtd.getAttribute('data-r'))) {
          e.preventDefault();
          e.stopPropagation();
          openRow(+mtd.getAttribute('data-r'), true);
          return;
        }
      }
      if (e.button !== 0) return;
      if (t.closest('.pane-resizer')) {
        layoutDrag = { type: 'pane', x: e.clientX, w: paneWidth, moved: 0 };
        e.preventDefault();
        e.stopPropagation();
        return;
      }
      if (t.closest('table.grid')) {
        var col = colFromPoint(e);
        if (col >= 0) {
          var ths0 = shadow.querySelectorAll('table.grid thead th');
          var logc = col;
          if (col > 0 && ths0[col]) {
            var dcl = parseInt(ths0[col].getAttribute('data-col'), 10);
            if (dcl) logc = dcl;
          }
          layoutDrag = { type: 'col', col: logc, x: e.clientX, w: colWidths[logc] || 64, moved: 0 };
          e.preventDefault();
          e.stopPropagation();
        } else {
          var mth = t.closest('thead th[data-vis]');
          if (mth && !t.closest('.resz')) {
            layoutDrag = { type: 'colmove', from: +mth.getAttribute('data-vis'), x: e.clientX, moved: 0, to: +mth.getAttribute('data-vis'), after: 0 };
            e.preventDefault();
            e.stopPropagation();
          }
        }
      }
    }, true);
    shadow.addEventListener('mousemove', function (e) {
      if (layoutDrag) return;
      var wrap = shadow.querySelector('.gridwrap');
      if (wrap) wrap.classList.toggle('colresize', colFromPoint(e) >= 0);
    });
    window.addEventListener('mousemove', function (e) {
      if (!layoutDrag) return;
      if (layoutDrag.type === 'col') {
        layoutDrag.moved = 1;
        colWidths[layoutDrag.col] = Math.max(36, layoutDrag.w + (e.clientX - layoutDrag.x));
        applyColWidths();
      } else if (layoutDrag.type === 'colmove') {
        if (Math.abs(e.clientX - layoutDrag.x) > 4) layoutDrag.moved = 1;
        if (!layoutDrag.moved || !shadow) return;
        var ths2 = shadow.querySelectorAll('table.grid thead th[data-vis]');
        var i2, r2, hit = null, after = 0;
        for (i2 = 0; i2 < ths2.length; i2++) {
          ths2[i2].classList.remove('dragging', 'drop-before', 'drop-after');
          r2 = ths2[i2].getBoundingClientRect();
          if (e.clientX >= r2.left && e.clientX <= r2.right) {
            hit = ths2[i2];
            after = e.clientX > (r2.left + r2.width / 2) ? 1 : 0;
          }
        }
        var fromTh = shadow.querySelector('th[data-vis="' + layoutDrag.from + '"]');
        if (fromTh) fromTh.classList.add('dragging');
        var wrapM = shadow.querySelector('.gridwrap');
        if (wrapM) wrapM.classList.add('colmove');
        if (hit) {
          hit.classList.add(after ? 'drop-after' : 'drop-before');
          layoutDrag.to = +hit.getAttribute('data-vis');
          layoutDrag.after = after;
        }
      } else {
        layoutDrag.moved = 1;
        paneWidth = Math.max(180, Math.min(900, layoutDrag.w - (e.clientX - layoutDrag.x)));
        applyPaneWidth();
      }
    });
    window.addEventListener('mouseup', function () {
      if (!layoutDrag) return;
      if (layoutDrag.type === 'col') saveCols();
      else if (layoutDrag.type === 'colmove') {
        var fromV = layoutDrag.from, toV = layoutDrag.to, afterV = layoutDrag.after, did = layoutDrag.moved;
        clearColDropUi();
        if (did && toV) {
          var dest = toV + (afterV ? 1 : 0);
          if (fromV < dest) dest--;
          if (moveColOrder(fromV, dest)) render();
        }
      } else savePane();
      layoutDrag = null;
    });
    var board = shadow.querySelector('#xl-board');
    if (board) {
      board.addEventListener('change', function () {
        if (board.value) goTo(board.value);
      });
    }
    var subPop = shadow.querySelector('#xl-subpop');
    if (subPop) {
      subPop.addEventListener('change', function (ev) {
        var inp = ev.target;
        if (!inp || inp.type !== 'checkbox') return;
        var name = inp.getAttribute('data-sub') || '';
        state.subPick = state.subPick || {};
        if (!name) state.subPick = {};
        else if (inp.checked) state.subPick[name] = 1;
        else delete state.subPick[name];
        state.subPop = true;
        saveSubPick();
        render();
      });
    }
    var search = shadow.querySelector('.search');
    if (search) {
      search.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') {
          ev.preventDefault();
          doSearch(search.value);
        }
      });
    }
    var findInp = shadow.querySelector('.findbar input');
    if (findInp) {
      findInp.addEventListener('keydown', function (ev) {
        if (ev.key === 'Enter') {
          ev.preventDefault();
          runFind(findInp.value, true, ev.shiftKey);
        }
        if (ev.key === 'Escape') closeFind();
      });
      findInp.addEventListener('input', function () { runFind(findInp.value, false); });
    }
    var cbody = shadow.querySelector('.compose .cbody');
    var ctitle = shadow.querySelector('.compose .ctitle');
    function onDraft() { saveDraft(); }
    if (cbody) cbody.addEventListener('input', onDraft);
    if (ctitle) ctitle.addEventListener('input', onDraft);
    var fileInp = shadow.querySelector('.compose .cfile');
    if (fileInp) {
      fileInp.addEventListener('change', function () {
        var f = fileInp.files && fileInp.files[0];
        if (f) uploadAttach(f);
        fileInp.value = '';
      });
    }
    ['#xl-page', '#xl-floor'].forEach(function (seln) {
      var el = shadow.querySelector(seln);
      if (!el) return;
      el.addEventListener('keydown', function (ev) {
        if (ev.key !== 'Enter') return;
        ev.preventDefault();
        if (seln === '#xl-page') gotoPage(el.value);
        else gotoFloor(el.value);
      });
    });
    shadow.addEventListener('auxclick', function (e) {
      if (e.button !== 1) return;
      var t = e.target;
      if (t && t.nodeType === 3) t = t.parentElement;
      if (!t || !t.closest) return;
      if (t.closest('a.cella')) return;
      var imgA = t.closest('a.imglink');
      if (imgA) {
        e.preventDefault();
        openBg(imgA.getAttribute('href') || imgA.href);
        return;
      }
      var td = t.closest('td[data-r]');
      if (td && rowHref(+td.getAttribute('data-r'))) {
        e.preventDefault();
        openRow(+td.getAttribute('data-r'), true);
      }
    });
    var wrap = shadow.querySelector('.gridwrap');
    if (wrap && !wrap.getAttribute('data-xl-scroll')) {
      wrap.setAttribute('data-xl-scroll', '1');
      wrap.addEventListener('scroll', function () {
        if (wrap.scrollTop + wrap.clientHeight >= wrap.scrollHeight - 140) loadMorePosts();
      });
    }
    window.addEventListener('resize', function () {
      if (state.mode === 'off') return;
      if (fillGridToView()) render();
    });
  }

  function isSheetBoss(e) {
    if (e.repeat || e.metaKey) return false;
    var q = e.code === 'KeyQ' || e.key === 'q' || e.key === 'Q';
    if (e.altKey && !e.ctrlKey && !e.shiftKey && q) return true;
    if (e.altKey) return false;
    if (e.key === 'F9' || e.code === 'F9') return !e.ctrlKey;
    if (e.ctrlKey || e.shiftKey) return false;
    return e.code === 'Backquote' || e.key === '`';
  }
  function isHideBoss(e) {
    if (e.repeat) return false;
    if (e.key === 'F10' || e.code === 'F10') return true;
    return e.ctrlKey && e.shiftKey && !e.altKey && (e.code === 'KeyQ' || e.key === 'Q' || e.key === 'q' || e.code === 'Backquote');
  }
  function consume(e) {
    e.preventDefault();
    e.stopPropagation();
    if (e.stopImmediatePropagation) e.stopImmediatePropagation();
  }
  function onBossKey(e) {
    if (isHideBoss(e)) {
      consume(e);
      if (!host) return;
      setMode(state.mode === 'off' ? 'skin' : 'off');
      return true;
    }
    if (isSheetBoss(e)) {
      consume(e);
      if (!host) return;
      if (state.mode === 'off') {
        setMode('skin');
        setSheet('pl');
        return true;
      }
      setSheet(state.sheet === 'pl' ? 'tk' : 'pl');
      return true;
    }
    return false;
  }
  window.addEventListener('keydown', function (e) {
    if (onBossKey(e)) return;
    if (state.mode === 'off') return;
    var k = e.key;
    if (k === 'F5' || ((e.ctrlKey || e.metaKey) && !e.altKey && (k === 'r' || k === 'R') && !e.shiftKey)) {
      consume(e);
      hardReload();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && !e.altKey && (k === 'f' || k === 'F')) {
      consume(e);
      openFind();
      return;
    }
    if ((e.ctrlKey || e.metaKey) && !e.altKey && (k === 'c' || k === 'C')) {
      if (eventInCompose(e) || eventInPaneText(e)) return;
      consume(e);
      if (e.shiftKey) copyRowLink();
      else copySelCell();
      return;
    }
    if (k === 'F3' || ((e.ctrlKey || e.metaKey) && !e.altKey && (k === 'g' || k === 'G'))) {
      consume(e);
      if (!state.find || !state.find.hits || !state.find.q) openFind();
      else runFind(state.find.q, true, e.shiftKey);
      return;
    }
    if (k === 'Escape') {
      closeFind();
      hideDlg();
      hideFileMenu();
    }
    var inC = eventInCompose(e);
    if (inC) {
      if ((e.ctrlKey || e.metaKey) && (k === 'Enter' || k === 'NumpadEnter')) {
        e.preventDefault();
        composeSend();
      }
      return;
    }
    if (k === 'ArrowDown') { e.preventDefault(); selectCell(state.selR + 1, state.selC); }
    else if (k === 'ArrowUp') { e.preventDefault(); selectCell(state.selR - 1, state.selC); }
    else if (k === 'ArrowLeft') { e.preventDefault(); selectCell(state.selR, stepCol(state.selC, -1)); }
    else if (k === 'ArrowRight') { e.preventDefault(); selectCell(state.selR, stepCol(state.selC, 1)); }
    else if (k === 'Enter') { e.preventDefault(); openRow(state.selR); }
    else if (k === 'Tab') { e.preventDefault(); selectCell(state.selR, stepCol(state.selC, e.shiftKey ? -1 : 1)); }
  }, true);
  window.addEventListener('keyup', function (e) {
    if (e.key === 'F10' || e.code === 'F10') {
      e.preventDefault();
      e.stopPropagation();
    }
  }, true);

  var refreshTimer = null;
  var lastSig = '';
  function trackerSig() {
    var topics = document.querySelectorAll('a.topic, a[href*="tid="]');
    var posts = document.querySelectorAll('[id^="postcontent"]');
    var boards = document.querySelectorAll('a[href*="thread.php?fid="]');
    var htmlLen = 0;
    Array.prototype.forEach.call(posts, function (n) { htmlLen += (n.innerHTML || '').length; });
    var qn = document.querySelectorAll('.quote, blockquote').length;
    var im = document.querySelectorAll('img[orgSrc], img[data-src], [id^="postcontent"] img').length;
    return location.href + '|' + document.title + '|' + topics.length + '|' + posts.length + '|' + boards.length + '|' + htmlLen + '|' + qn + '|' + im;
  }
  function scheduleRefresh() {
    if (state.mode === 'off' || state.sheet !== 'tk') return;
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(function () {
      var sig = trackerSig();
      if (sig === lastSig) return;
      lastSig = sig;
      var wrap = $('.gridwrap');
      var sl = wrap ? wrap.scrollLeft : 0;
      var st2 = wrap ? wrap.scrollTop : 0;
      render();
      wrap = $('.gridwrap');
      if (wrap) { wrap.scrollLeft = sl; wrap.scrollTop = st2; }
    }, 400);
  }

  function bootHide() {
    if ((localStorage.getItem(LS_MODE) || 'skin') === 'off') return;
    document.documentElement.classList.add('nga-xl-pending');
    bootStyle = document.createElement('style');
    bootStyle.textContent = 'html.nga-xl-pending,html.nga-xl-pending body{background:#fff!important;opacity:0!important;}html.nga-xl-on,html.nga-xl-on body{opacity:1!important;}html.nga-xl-on body{overflow:hidden!important;}';
    (document.head || document.documentElement).appendChild(bootStyle);
  }

  function releaseIfLogin() {
    if (!host || state.mode === 'off') return;
    if (isLoginWall()) { showLoginDlg(); return; }
    if (needsRule()) { showRuleDlg(); }
  }
  function start() {
    if (!document.documentElement) return;
    mount();
    if (document.body && host.parentNode !== document.body) document.body.appendChild(host);
    hydrateLayout();
    persistLayout();
    if (state.mode === 'off') {
      applyMode();
      return;
    }
    state.sheet = 'tk';
    applyMode();
    setFavicon();
    try { mergeUserInfo((document.documentElement && document.documentElement.innerHTML) || ''); } catch (eU) {}
    lastSig = '';
    scheduleRefresh();
    setTimeout(scheduleRefresh, 800);
    setTimeout(scheduleRefresh, 2000);
    setTimeout(scheduleRefresh, 4500);
    setTimeout(releaseIfLogin, 900);
    setTimeout(releaseIfLogin, 2500);
    setTimeout(restoreDraft, 600);
    try {
      var mo = new MutationObserver(scheduleRefresh);
      mo.observe(document.documentElement, { childList: true, subtree: true });
    } catch (e) {}
    if (!window.__xlTitleLock) {
      window.__xlTitleLock = setInterval(lockTitle, 800);
    }
    if (!window.__xlPop) {
      window.__xlPop = 1;
      window.addEventListener('popstate', function () {
        if (state.mode === 'off') return;
        silentGo(location.href, { history: 'none', force: 1, reload: 1 });
      });
    }
  }

  function persistLayout() {
    try { saveCols(); savePane(); saveColOrders(); } catch (ePl) {}
  }
  window.addEventListener('pagehide', persistLayout);
  document.addEventListener('visibilitychange', function () {
    if (document.visibilityState === 'hidden') persistLayout();
  });
  bootHide();
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', start);
  } else {
    start();
  }
  setTimeout(function () { if (!host) start(); }, 1200);
})();
