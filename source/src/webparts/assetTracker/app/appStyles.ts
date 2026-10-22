// Scoped app styles (every selector is under .itat)
export const APP_CSS: string = `.itat{
  --bg:#EDF0F3; --surface:#FFFFFF; --surface-2:#F5F7F9; --ink:#17222D; --ink-2:#4A5866; --ink-3:#7A8794;
  --line:#D6DCE2; --line-2:#E6EAEE;
  --accent:#0E7480; --accent-ink:#FFFFFF; --accent-soft:#D9EEF0;
  --ok:#2E7A4C; --ok-soft:#DDEFE3; --warn:#B07A06; --warn-soft:#F7ECD2; --bad:#B23A2C; --bad-soft:#F6DDD9;
  --info:#4B5CA8; --info-soft:#E1E5F5; --reserve:#7FB8BE;
  --radius:10px; --radius-s:6px;
  --font:"Archivo",system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif;color-scheme:light;color:var(--ink);font-family:var(--font);font-size:15px;line-height:1.45;font-variant-numeric:tabular-nums;-webkit-font-smoothing:antialiased;text-align:left}
.itat[data-theme="dark"]{
  --bg:#10171E; --surface:#18222C; --surface-2:#1E2A35; --ink:#E7EDF2; --ink-2:#AEBAC5; --ink-3:#7D8B98;
  --line:#2C3A47; --line-2:#24313C; --accent:#3FB3BF; --accent-ink:#0B1A1D; --accent-soft:#16363A;
  --ok:#5CC08A; --ok-soft:#173326; --warn:#E2AE3A; --warn-soft:#3A2F14; --bad:#EE7D6E; --bad-soft:#3D1F1B;
  --info:#93A2EA; --info-soft:#232A47; --reserve:#2E6A70; ;color-scheme:dark}
.itat.itat-page{background:var(--bg);border-radius:var(--radius);min-height:560px;overflow:hidden}
.itat .topbar{display:flex;align-items:center;gap:14px 22px;flex-wrap:wrap;padding:14px clamp(16px,3vw,32px);background:var(--surface);border-bottom:1px solid var(--line)}
.itat .brand{padding:0}
.itat .tabs{display:flex;gap:4px;overflow-x:auto;flex:1 1 420px;min-width:0}
.itat .tabs button{all:unset;cursor:pointer;padding:7px 12px;border-radius:999px;white-space:nowrap;font-weight:600;color:var(--ink-2);display:inline-flex;gap:6px;align-items:center}
.itat .tabs button:hover{background:var(--surface-2);color:var(--ink)}
.itat .tabs button:focus-visible{outline:2px solid var(--accent)}
.itat .tabs button[aria-current="page"]{background:var(--ink);color:var(--surface)}
.itat .tabs .count{font-size:12px;opacity:.65;font-weight:500}
.itat .sync{font-size:12px;color:var(--ink-3);display:flex;gap:8px;align-items:center;white-space:nowrap}
.itat main{display:block;padding:24px clamp(16px,3vw,32px) 48px;max-width:1320px;min-width:0}
.itat *,.itat *::before,.itat *::after{box-sizing:border-box}
.itat button,.itat input,.itat select,.itat textarea{font:inherit;color:inherit}
.itat a{color:var(--accent)}
.itat :focus-visible{outline:2px solid var(--accent);outline-offset:2px}
.itat .brand{font-stretch:75%;font-weight:800;font-size:22px;letter-spacing:-.01em;line-height:1;padding:0 10px 4px}
.itat .brand small{display:block;font-stretch:100%;font-weight:500;font-size:12.5px;color:var(--ink-3);margin-top:6px;letter-spacing:0}
.itat h1{font-stretch:75%;font-weight:800;font-size:clamp(30px,4vw,40px);line-height:1.02;margin:0 0 6px;letter-spacing:-.015em}
.itat h2{font-stretch:85%;font-weight:750;font-size:20px;margin:0 0 12px;letter-spacing:-.005em}
.itat h3{font-weight:650;font-size:15px;margin:0 0 8px}
.itat .lede{color:var(--ink-2);margin:0 0 24px;max-width:70ch}
.itat .row{display:flex;gap:10px;align-items:center;flex-wrap:wrap}
.itat .spread{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap}
.itat .panel{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius);padding:18px 20px}
.itat .panel + .panel{margin-top:16px}
.itat .grid2>.panel+.panel,.itat .grid3>.panel+.panel{margin-top:0}
.itat a.btn{text-decoration:none;color:var(--ink)}
.itat .grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px}
.itat .grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px}
.itat .stack>*+*{margin-top:16px}
.itat .muted{color:var(--ink-3)}
.itat .small{font-size:13px}
.itat .tiny{font-size:12px}
.itat .btn{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);background:var(--surface);border-radius:var(--radius-s);padding:7px 12px;cursor:pointer;font-weight:550;white-space:nowrap}
.itat .btn:hover{border-color:var(--ink-3)}
.itat .btn.primary{background:var(--accent);border-color:var(--accent);color:var(--accent-ink)}
.itat .btn.primary:hover{filter:brightness(1.07)}
.itat .btn.ghost{border-color:transparent;background:transparent}
.itat .btn.danger{color:var(--bad)}
.itat .btn:disabled{opacity:.55;cursor:default}
.itat .btn.sm{padding:4px 9px;font-size:13px}
.itat input[type=text],.itat input[type=search],.itat input[type=number],.itat input[type=date],.itat select,.itat textarea{background:var(--surface);border:1px solid var(--line);border-radius:var(--radius-s);padding:7px 10px;min-width:0}
.itat textarea{width:100%;min-height:70px;resize:vertical}
.itat input:focus,.itat select:focus,.itat textarea:focus{outline:2px solid var(--accent);outline-offset:-1px;border-color:transparent}
.itat .search{flex:1 1 240px;max-width:420px}
.itat .chips{display:flex;gap:6px;flex-wrap:wrap}
.itat .chip{border:1px solid var(--line);background:var(--surface);border-radius:999px;padding:5px 12px;cursor:pointer;font-weight:550;font-size:14px}
.itat .chip[aria-pressed="true"]{background:var(--ink);color:var(--surface);border-color:var(--ink)}
.itat .chip .n{opacity:.6;margin-left:4px;font-weight:500}
.itat .tag{display:inline-block;font-size:12px;font-weight:600;padding:1px 7px;border-radius:4px;background:var(--surface-2);color:var(--ink-2);white-space:nowrap}
.itat .tag.ok{background:var(--ok-soft);color:var(--ok)}
.itat .tag.warn{background:var(--warn-soft);color:var(--warn)}
.itat .tag.bad{background:var(--bad-soft);color:var(--bad)}
.itat .tag.info{background:var(--info-soft);color:var(--info)}
.itat .tag.acc{background:var(--accent-soft);color:var(--accent)}
.itat .tablewrap{overflow-x:auto;border:1px solid var(--line);border-radius:var(--radius);background:var(--surface)}
.itat table{border-collapse:collapse;width:100%;font-size:14px}
.itat th .sorth{all:unset;cursor:pointer;display:inline-flex;align-items:center;gap:4px;color:inherit;font:inherit;white-space:nowrap}
.itat th .sorth:hover{color:var(--ink)}
.itat th .sorth:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:2px}
.itat th .sorth .ar{display:inline-block;width:10px;font-size:11px;opacity:.35}
.itat th[aria-sort] .sorth{color:var(--ink)}
.itat th[aria-sort] .sorth .ar{opacity:1}
.itat th.num .sorth{flex-direction:row-reverse}
.itat th{font-weight:600;color:var(--ink-3);font-size:12.5px;text-align:left;padding:9px 12px;border-bottom:1px solid var(--line);white-space:nowrap;background:var(--surface);position:sticky;top:0}
.itat td{padding:9px 12px;border-bottom:1px solid var(--line-2);vertical-align:middle}
.itat tr:last-child td{border-bottom:0}
.itat tbody tr.click{cursor:pointer}
.itat tbody tr.click:hover{background:var(--surface-2)}
.itat td.num,.itat th.num{text-align:right}
.itat .nm{font-weight:600}
.itat .sub{display:block;font-size:12.5px;color:var(--ink-3);font-weight:400}
.itat .life{display:inline-block;width:64px;height:6px;border-radius:3px;background:var(--line-2);vertical-align:middle;overflow:hidden;margin-right:8px}
.itat .life i{display:block;height:100%;background:var(--ok)}
.itat .life i.warn{background:var(--warn)}
.itat .life i.bad{background:var(--bad)}
.itat .life i.none{background:transparent}
.itat .agecell{white-space:nowrap}
.itat /* horizon */
.horizon{display:grid;grid-auto-flow:column;grid-auto-columns:minmax(0,1fr);gap:10px;align-items:end;min-height:220px;padding-top:8px}
.itat .hcol{display:flex;flex-direction:column;justify-content:flex-end;cursor:pointer;border-radius:var(--radius-s);padding:0 2px 0}
.itat .hcol:hover .hbar{filter:brightness(1.08)}
.itat .hcol[aria-pressed="true"] .hyear{color:var(--ink);font-weight:750}
.itat .hcol[aria-pressed="true"] .hbar{outline:2px solid var(--ink);outline-offset:2px}
.itat .hamt{font-weight:700;font-size:14px;margin-bottom:6px;font-stretch:85%}
.itat .hamt small{display:block;font-weight:500;color:var(--ink-3);font-size:12px;font-stretch:100%}
.itat .hbar{display:flex;flex-direction:column;justify-content:flex-end;border-radius:4px 4px 0 0;overflow:hidden;min-height:3px}
.itat .hbar .b{background:var(--accent)}
.itat .hbar .r{background:repeating-linear-gradient(135deg,var(--reserve) 0 5px,transparent 5px 9px);border:1px solid var(--reserve);border-bottom:0}
.itat .hyear{border-top:2px solid var(--ink);padding-top:6px;font-weight:600;color:var(--ink-2);font-size:14px}
.itat .legend{display:flex;gap:16px;font-size:13px;color:var(--ink-2)}
.itat .legend i{display:inline-block;width:12px;height:12px;border-radius:2px;margin-right:6px;vertical-align:-1px}
.itat .legend .lb{background:var(--accent)}
.itat .legend .lr{background:repeating-linear-gradient(135deg,var(--reserve) 0 3px,transparent 3px 6px);border:1px solid var(--reserve)}
.itat .kv{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;font-size:14px}
.itat .kv dt{color:var(--ink-3)}
.itat .kv dd{margin:0}
.itat .metric{font-stretch:80%;font-weight:750;font-size:28px;line-height:1.05}
.itat .metric-l{color:var(--ink-3);font-size:13px}
.itat .metrics{display:flex;gap:28px;flex-wrap:wrap}
.itat .sitelist{display:grid;gap:0}
.itat .siterow{display:grid;grid-template-columns:minmax(120px,1.4fr) repeat(4,minmax(70px,1fr));gap:12px;align-items:center;padding:11px 4px;border-bottom:1px solid var(--line-2);cursor:pointer}
.itat .siterow:last-child{border-bottom:0}
.itat .siterow:hover{background:var(--surface-2)}
.itat .siterow.head{cursor:default;color:var(--ink-3);font-size:12.5px;font-weight:600}
.itat .siterow.head:hover{background:none}
.itat .siterow .r{text-align:right}
.itat .attn{list-style:none;margin:0;padding:0}
.itat .attn li{display:flex;justify-content:space-between;gap:12px;padding:10px 0;border-bottom:1px solid var(--line-2);align-items:center}
.itat .attn li:last-child{border-bottom:0}
.itat .attn button{all:unset;cursor:pointer;color:var(--accent);font-weight:600;font-size:14px}
.itat .people{display:grid;grid-template-columns:minmax(260px,340px) 1fr;gap:16px;align-items:start}
.itat .plist{max-height:70vh;overflow:auto;padding:4px}
.itat .pitem{display:flex;justify-content:space-between;gap:8px;width:100%;text-align:left;border:0;background:none;padding:9px 10px;border-radius:var(--radius-s);cursor:pointer}
.itat .pitem:hover{background:var(--surface-2)}
.itat .pitem[aria-current="true"]{background:var(--accent-soft)}
.itat .badges{display:flex;gap:6px;flex-wrap:wrap}
.itat /* drawer */
.scrim{position:fixed;inset:0;background:rgba(10,16,22,.42);z-index:10000;display:none}
.itat .drawer{position:fixed;top:0;right:0;bottom:0;width:min(560px,100%);background:var(--surface);z-index:10001;transform:translateX(100%);transition:transform .22s ease;display:flex;flex-direction:column;border-left:1px solid var(--line);
  padding-top:env(safe-area-inset-top,0px);padding-bottom:env(safe-area-inset-bottom,0px)}
.itat .drawer.open{transform:none}
.itat .scrim.open{display:block}
@media (prefers-reduced-motion:reduce){.itat .drawer{transition:none}}
.itat .dhead{padding:18px 20px 12px;border-bottom:1px solid var(--line);display:flex;justify-content:space-between;gap:10px;align-items:flex-start}
.itat .dbody{flex:1;overflow:auto;padding:16px 20px 24px}
.itat .dfoot{padding:12px 20px;border-top:1px solid var(--line);display:flex;gap:8px;justify-content:space-between;flex-wrap:wrap}
.itat .form{display:grid;grid-template-columns:1fr 1fr;gap:10px 12px}
.itat .form label{display:flex;flex-direction:column;gap:4px;font-size:12.5px;color:var(--ink-3);font-weight:600}
.itat .form label.full{grid-column:1/-1}
.itat .form label > *{font-size:14.5px;color:var(--ink);font-weight:400}
.itat .sect{margin-top:22px}
.itat .sect:first-child{margin-top:0}
.itat .sect > h3{display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--line);padding-bottom:6px;margin-bottom:12px}
.itat .ai{background:var(--surface-2);border:1px solid var(--line);border-radius:var(--radius);padding:14px}
.itat .ai .model{font-weight:700;font-size:17px;font-stretch:85%}
.itat .vendors{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px}
.itat .prose{white-space:normal;max-width:75ch}
.itat .prose p{margin:.4em 0}
.itat .prose ul{margin:.3em 0;padding-left:1.2em}
.itat .prose h4{margin:.9em 0 .2em;font-size:15px}
.itat details.grp{border-bottom:1px solid var(--line-2)}
.itat details.grp:last-child{border-bottom:0}
.itat details.grp summary{list-style:none;cursor:pointer;display:grid;grid-template-columns:minmax(160px,2fr) minmax(70px,.6fr) minmax(90px,1fr) minmax(90px,1fr);gap:12px;padding:12px 6px;align-items:center}
.itat details.grp summary::-webkit-details-marker{display:none}
.itat details.grp summary:hover{background:var(--surface-2)}
.itat details.grp summary .tw{display:inline-block;width:14px;color:var(--ink-3);transition:transform .15s}
.itat details.grp[open] summary .tw{transform:rotate(90deg)}
.itat .ghead{display:grid;grid-template-columns:minmax(160px,2fr) minmax(70px,.6fr) minmax(90px,1fr) minmax(90px,1fr);gap:12px;padding:0 6px 8px;color:var(--ink-3);font-size:12.5px;font-weight:600;border-bottom:1px solid var(--line)}
.itat .r{text-align:right}
.itat .gbody{padding:0 0 12px 26px}
.itat .empty{padding:36px 20px;text-align:center;color:var(--ink-2)}
.itat .toast{position:fixed;left:50%;bottom:calc(20px + env(safe-area-inset-bottom,0px));transform:translateX(-50%);background:var(--ink);color:var(--surface);padding:10px 16px;border-radius:8px;z-index:10002;font-weight:550;display:none;max-width:90vw}
.itat .toast.show{display:block}
.itat .bulkbar{position:fixed;left:50%;bottom:calc(18px + env(safe-area-inset-bottom,0px));transform:translateX(-50%);z-index:9999;background:var(--ink);color:var(--surface);border-radius:12px;padding:8px 8px 8px 16px;display:flex;align-items:center;gap:10px;box-shadow:0 8px 24px rgba(10,16,22,.28);max-width:calc(100vw - 32px);flex-wrap:wrap}
.itat .bulkbar[hidden]{display:none}
.itat .bulkbar b{font-weight:700}
.itat .bulkbar .btn{background:transparent;color:var(--surface);border-color:rgba(255,255,255,.28)}
.itat .bulkbar .btn:hover{border-color:var(--surface)}
.itat .bulkbar .btn.primary{background:var(--surface);color:var(--ink);border-color:var(--surface)}
.itat .bulkbar .btn.danger{color:#ffb4a9}
.itat .bulkbar:not([hidden]) ~ .toast{bottom:calc(84px + env(safe-area-inset-bottom,0px))}
.itat td.selc,.itat th.selc{width:34px;padding-right:0}
.itat td.selc input,.itat th.selc input{width:16px;height:16px;cursor:pointer;accent-color:var(--accent);margin:0;vertical-align:middle}
.itat tr.sel td{background:var(--accent-soft)}
.itat .tl{list-style:none;margin:0;padding:0 0 0 14px;border-left:2px solid var(--line)}
.itat .tl li{position:relative;padding:0 0 14px 14px}
.itat .tl li::before{content:"";position:absolute;left:-21px;top:5px;width:10px;height:10px;border-radius:50%;background:var(--surface);border:2px solid var(--ink-3)}
.itat .tl li.note::before{border-color:var(--accent);background:var(--accent)}
.itat .tl li.key::before{border-color:var(--info);background:var(--info)}
.itat .tl .when{font-size:12.5px;color:var(--ink-3)}
.itat .tl .what{font-size:14px;margin-top:2px}
.itat .tl .what ul{margin:2px 0 0;padding-left:18px}
.itat .tl .what .arrow{color:var(--ink-3)}
.itat .notebox{background:var(--surface-2);border:1px solid var(--line);border-radius:var(--radius);padding:12px;margin-bottom:14px}
.itat .plink{all:unset;cursor:pointer;color:var(--ink);text-decoration:underline;text-decoration-color:var(--line);text-underline-offset:3px}
.itat .plink:hover{color:var(--accent);text-decoration-color:var(--accent)}
.itat .plink:focus-visible{outline:2px solid var(--accent);outline-offset:2px;border-radius:2px}
.itat .fieldlinks{display:flex;gap:12px;margin-top:4px;font-size:12.5px;font-weight:500}
.itat .fieldlinks button{all:unset;cursor:pointer;color:var(--accent)}
.itat .fieldlinks button:hover{text-decoration:underline}
.itat .bulkhint{font-size:13px;color:var(--ink-2);margin:0 0 12px}
.itat .banner{background:var(--warn-soft);color:var(--ink);border:1px solid var(--warn);border-radius:var(--radius);padding:12px 16px;margin-bottom:18px}
@media (max-width:900px){.itat .grid2,.itat .grid3,.itat .people{grid-template-columns:1fr}
.itat .plist{max-height:40vh}
.itat .siterow{grid-template-columns:1.3fr repeat(2,1fr)}
.itat .siterow .hm{display:none}
.itat .form{grid-template-columns:1fr}
.itat details.grp summary,.itat .ghead{grid-template-columns:1.6fr .6fr 1fr}
.itat .hide-m{display:none}
.itat .pd{display:none}
.itat .horizon{gap:6px}
.itat .hamt small{font-size:11px}}`;
