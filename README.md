<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Scholastic Testing Platform</title>
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/animate.css/4.1.1/animate.min.css"/>
<style>
@import url('https://fonts.googleapis.com/css2?family=Tajawal:wght@300;400;500;700;900&family=Montserrat:wght@400;600;700;800;900&family=Amiri:wght@400;700&family=EB+Garamond:wght@700&family=Playfair+Display:wght@700;800&family=Jost:wght@500;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
.hidden{display:none!important}
body{font-family:'Tajawal',sans-serif;background:linear-gradient(160deg,#001a3a 0%,#002147 40%,#00306b 70%,#001a3a 100%);min-height:100vh;display:flex;flex-direction:column;align-items:center;justify-content:space-between;color:white;padding:15px}
.font-en{font-family:'Montserrat',sans-serif}
.hidden-section{display:none!important}
.glass-card{background:rgba(255,255,255,.07);backdrop-filter:blur(15px);border:1px solid rgba(255,255,255,.15);border-radius:2.5rem;padding:2.5vh 20px;text-align:center;cursor:pointer;transition:box-shadow .3s,background .3s,border-color .3s;display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%}
.glass-card:hover{background:rgba(255,255,255,.13);box-shadow:0 6px 24px rgba(250,204,21,.18);border-color:rgba(250,204,21,.35)}
.logout-btn{display:none}
#loginModal{position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,.7);backdrop-filter:blur(8px);display:none;align-items:center;justify-content:center;z-index:1000}
.auth-form{background:#ffffff;border:1px solid #e2e8f0;border-radius:2rem;padding:40px;width:90%;max-width:400px;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.auth-input{width:100%;background:#f8fafc;border:1.5px solid #d1d5db;border-radius:12px;padding:12px 15px;color:#1a1a2e;margin-bottom:15px;outline:none;font-family:'Tajawal',sans-serif;font-size:14px}
.auth-input::placeholder{color:#94a3b8}
.auth-input:focus{border-color:#FACC15;box-shadow:0 0 0 3px rgba(250,204,21,.2)}
.stat-container{background:rgba(0,0,0,.3);backdrop-filter:blur(12px);border:1px solid rgba(255,255,255,.2);display:flex;flex-direction:column;align-items:center;padding:15px 25px;width:100%;max-width:400px;border-radius:1.5rem}
.update-link{font-size:10px;color:rgba(255,255,255,.4);margin-top:8px;cursor:pointer;text-transform:uppercase;letter-spacing:1px}
.update-link:hover{color:#FACC15}
@keyframes marquee{0%{transform:translateX(0)}100%{transform:translateX(50%)}}
.marquee-wrapper{width:100%;background:rgba(0,0,0,.15);padding:12px 0;border-top:1px solid rgba(255,255,255,.1);overflow:hidden}
.marquee-content{display:flex;gap:5rem;animation:marquee 35s linear infinite;width:max-content}
.flag-only{font-size:2.5rem;filter:drop-shadow(0 4px 6px rgba(0,0,0,.2))}
.admin-table{width:100%;border-collapse:collapse;font-size:13px}
.admin-table th{background:rgba(250,204,21,.2);color:#FACC15;padding:10px;border:1px solid rgba(255,255,255,.1);text-align:center}
.admin-table td{padding:9px;border:1px solid rgba(255,255,255,.1);background:rgba(0,0,0,.2);text-align:center;vertical-align:middle}
.wizard-input{width:100%;background:#ffffff;border:1.5px solid #d1d5db;border-radius:12px;padding:11px 14px;color:#1a1a2e;outline:none;font-family:'Tajawal',sans-serif;font-size:14px;transition:.2s}
.wizard-input::placeholder{color:#94a3b8}
.wizard-input:focus{border-color:#FACC15;box-shadow:0 0 0 3px rgba(250,204,21,.25);background:#ffffff}
.wizard-input option{background:#ffffff;color:#1a1a2e}
.wizard-label{font-size:11px;color:#FACC15;text-transform:uppercase;letter-spacing:1px;margin-bottom:5px;display:block}
.wizard-label-en{font-size:9px;color:rgba(255,255,255,.45);font-family:'Montserrat',sans-serif;margin-right:6px}
.step-indicator{display:flex;gap:8px;justify-content:center;margin-bottom:24px}
.step-dot{width:10px;height:10px;border-radius:50%;background:rgba(255,255,255,.2);transition:.3s}
.step-dot.active{background:#FACC15;transform:scale(1.4)}
.step-dot.done{background:#4ade80}
select.wizard-input{appearance:none;-webkit-appearance:none}
.domain-row{display:flex;gap:10px;align-items:center;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.15);border-radius:12px;padding:10px 14px;margin-bottom:8px}
.domain-badge{background:linear-gradient(135deg,#FACC15,#f59e0b);color:#1e3a8a;font-weight:800;border-radius:8px;min-width:28px;height:28px;display:flex;align-items:center;justify-content:center;font-size:13px;font-family:'Montserrat',sans-serif;flex-shrink:0}
.reviewer-box{background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.2);border-radius:16px;padding:16px 18px}
.reviewer-box h6{font-size:12px;color:#93c5fd;margin-bottom:8px}
.link-badge{background:rgba(250,204,21,.12);border:1px dashed #FACC15;border-radius:10px;padding:8px 14px;font-family:'Montserrat',sans-serif;font-size:12px;color:#FACC15;word-break:break-all;margin-top:6px}
.logo-upload-area{border:2px dashed #d1d5db;border-radius:14px;padding:18px;text-align:center;cursor:pointer;transition:.2s;min-height:80px;display:flex;align-items:center;justify-content:center;flex-direction:column;background:#ffffff}
.logo-upload-area:hover{border-color:#FACC15}
.logo-upload-area img{background:#ffffff}
.schools-multi{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.25);border-radius:12px;padding:8px;min-height:100px;max-height:160px;overflow-y:auto}
.school-option{display:flex;align-items:center;gap:8px;padding:7px 10px;border-radius:8px;cursor:pointer;transition:.2s;font-size:14px}
.school-option:hover{background:rgba(255,255,255,.1)}
.school-option input[type=checkbox]{accent-color:#FACC15;width:16px;height:16px;flex-shrink:0}
.school-option.selected{background:rgba(250,204,21,.15);border:1px solid rgba(250,204,21,.3)}
.selected-schools-tags{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;min-height:28px}
.school-tag{background:rgba(250,204,21,.2);border:1px solid #FACC15;border-radius:20px;padding:3px 10px;font-size:12px;color:#FACC15;display:flex;align-items:center;gap:5px}
.school-tag .remove-tag{cursor:pointer;font-size:14px;color:rgba(255,255,255,.6)}
.school-tag .remove-tag:hover{color:#f87171}
.rtoolbar{background:rgba(0,0,0,.4);border-radius:12px;padding:8px 10px;display:flex;flex-wrap:wrap;gap:4px;align-items:center;border:1px solid rgba(255,255,255,.1);margin-bottom:8px}
.rtoolbar button,.rtoolbar select{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:7px;padding:4px 9px;font-size:12px;cursor:pointer;transition:.15s;font-family:'Tajawal',sans-serif}
.rtoolbar button:hover,.rtoolbar select:hover{background:rgba(250,204,21,.2);border-color:#FACC15}
.rtoolbar input[type=color]{width:28px;height:28px;border-radius:6px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:1px}
.rtoolbar .sep{width:1px;height:20px;background:rgba(255,255,255,.15);margin:0 3px}
.qh-box{background:#ffffff;border:1.5px solid #e2e8f0;border-radius:12px;padding:12px 16px;margin-bottom:12px;box-shadow:0 1px 4px rgba(0,0,0,.04)}
/* داخل شريط السؤال في نافذة الطالب: بدون خلفية بيضاء، يندمج مع الشريط */
.sw-q-stem .qh-box,.wp-q-stem .qh-box,.stream-q-card .qh-box{background:transparent;border:none;box-shadow:none;padding:0;margin-bottom:0}
.qh-ar{font-family:'Tajawal',sans-serif;font-size:18px;font-weight:700;color:rgb(19,36,224);direction:rtl;text-align:right;line-height:1.7}
.qh-en{font-family:'Tajawal',sans-serif;font-size:16px;font-weight:700;color:rgb(245,20,20);direction:ltr;text-align:left;line-height:1.7;margin-top:4px}
.qh-ar[contenteditable="true"],.qh-en[contenteditable="true"]{background:#fffbeb;border:1px dashed #FACC15;border-radius:6px;padding:4px 8px;outline:none}
.rich-editor{width:100%;min-height:80px;background:#ffffff;border:1.5px solid #d1d5db;border-radius:12px;padding:12px 14px;color:#1a1a2e;outline:none;font-family:'Tajawal',sans-serif;font-size:14px;line-height:1.8}
.rich-editor:focus{border-color:#FACC15;box-shadow:0 0 0 3px rgba(250,204,21,.25)}
.rich-editor[placeholder]:empty:before{content:attr(placeholder);color:#94a3b8}
.rich-editor:focus{border-color:#FACC15}
.rich-editor[placeholder]:empty:before{content:attr(placeholder);color:rgba(255,255,255,.3);pointer-events:none}
::-webkit-scrollbar{width:6px}::-webkit-scrollbar-track{background:rgba(255,255,255,.05)}::-webkit-scrollbar-thumb{background:#FACC15;border-radius:10px}
#studentWindow{position:fixed;inset:0;z-index:9999;background:#f0f4ff;display:none;flex-direction:column;font-family:'Tajawal',sans-serif;color:#1a1a2e;overflow:hidden}
.sw-header{background:linear-gradient(135deg,#1e3a8a,#7e22ce);color:white;padding:0 20px;height:70px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;border-bottom:3px solid #FACC15;box-shadow:0 4px 20px rgba(0,0,0,.3);width:100%}
.sw-logo-area{display:flex;align-items:center;gap:10px}
.sw-logo{height:44px;width:44px;object-fit:contain;border-radius:10px;background:white;padding:4px}
.sw-logo-text{font-size:16px;font-weight:900;font-family:'Montserrat',sans-serif}
.sw-timer-wrap{display:flex;flex-direction:column;align-items:center;background:rgba(0,0,0,.3);border:2px solid rgba(250,204,21,.5);border-radius:16px;padding:6px 18px;min-width:100px}
.sw-timer-label{font-size:9px;color:rgba(255,255,255,.6);text-transform:uppercase;letter-spacing:1px;font-family:'Montserrat',sans-serif;margin-bottom:2px}
.sw-timer{font-family:'Montserrat',sans-serif;font-size:28px;font-weight:900;letter-spacing:3px;color:#FACC15;line-height:1}
.sw-timer.warn{color:#f87171;animation:timerPulse 1s infinite}
@keyframes timerPulse{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.7;transform:scale(1.05)}}
.sw-sticky-bar{position:sticky;top:0;z-index:100;background:white;border-bottom:2px solid #e2e8f0;box-shadow:0 2px 12px rgba(0,0,0,.08);flex-shrink:0}
.sw-domain-bar{padding:10px 20px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px}
.sw-domain-name{font-weight:800;font-size:15px;color:#1e3a8a}
.sw-progress-dots{display:flex;gap:6px;align-items:center;flex-wrap:wrap}
.sw-content{flex:1;overflow-y:auto;padding:24px;width:100%}
.sw-frame{border-radius:22px;padding:0;background:white;box-shadow:0 0 0 3px #1e3a8a,0 0 0 6px rgba(250,204,21,.5),0 12px 40px rgba(30,58,138,.2);margin-bottom:16px;width:100%;overflow:hidden}
.sw-q-header{display:flex;align-items:flex-start;gap:14px;margin-bottom:0;padding:18px 22px;background:rgba(186,230,253,.35);border-bottom:2px solid rgba(14,165,233,.25)}
.sw-q-badge{background:#0ea5e9;border:2px solid #0284c7;color:white;font-weight:800;border-radius:12px;min-width:42px;height:42px;display:flex;align-items:center;justify-content:center;font-size:14px;font-family:'Montserrat',sans-serif;flex-shrink:0}
.sw-q-score{background:#e0f2fe;color:#0369a1;border:1px solid #bae6fd;border-radius:8px;padding:3px 10px;font-size:12px;font-weight:700;font-family:'Montserrat',sans-serif;white-space:nowrap}
.sw-q-stem{font-size:17px;font-weight:600;line-height:1.7;color:#0c4a6e;flex:1}
.sw-mcq-options{display:flex;flex-direction:column;gap:10px;margin-top:16px}
.sw-mcq-options.horizontal{flex-direction:row;flex-wrap:wrap}
.sw-mcq-opt{border:2px solid #dbeafe;border-radius:14px;padding:14px 18px;cursor:pointer;transition:all .2s;display:flex;align-items:center;gap:12px;background:#f0f7ff}
.sw-mcq-opt:hover{border-color:#7e22ce;background:#f5f3ff}
.sw-mcq-opt.selected{border-color:#1e3a8a;background:#eff6ff}
.sw-mcq-opt.horizontal{flex:1;min-width:140px}
.sw-opt-label{width:30px;height:30px;border-radius:50%;border:2px solid #cbd5e1;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;color:#64748b;flex-shrink:0;font-family:'Montserrat',sans-serif}
.sw-mcq-opt.selected .sw-opt-label{background:#1e3a8a;border-color:#1e3a8a;color:white}
.ordering-item{border:2px solid #e2e8f0;border-radius:12px;padding:12px 16px;background:#fafafa;font-size:15px;cursor:grab;display:flex;align-items:center;gap:10px;transition:.2s;margin-bottom:8px}
.ordering-item:hover{border-color:#7e22ce;background:#f5f3ff}
.ordering-item.dragging{opacity:.5;border-color:#7e22ce}
.ordering-item.drag-over{border-color:#22c55e;background:#f0fdf4}
.drag-handle{color:#94a3b8;font-size:16px;cursor:grab;flex-shrink:0}
.match-dot{width:14px;height:14px;border-radius:50%;background:#cbd5e1;border:2px solid #94a3b8;cursor:pointer;flex-shrink:0;transition:.2s}
.match-dot:hover{background:#1e3a8a;border-color:#1e40af}
.sw-footer{background:linear-gradient(135deg,#1e3a8a,#1e40af);border-top:3px solid #FACC15;padding:6px 14px;display:flex;align-items:center;justify-content:space-between;flex-shrink:0;box-shadow:0 -3px 10px rgba(30,58,138,.25);width:100%}
.sw-nav-btn{border:none;border-radius:10px;padding:7px 18px;font-size:14px;font-weight:800;cursor:pointer;transition:.2s;font-family:'Tajawal',sans-serif;display:flex;flex-direction:column;align-items:center;gap:1px}
.sw-nav-btn .btn-sub{font-size:10px;font-weight:600;opacity:.8;font-family:'Montserrat',sans-serif}
.sw-nav-btn.prev{background:linear-gradient(135deg,#FACC15,#f59e0b);color:#1e3a8a;border:3px solid white;box-shadow:0 4px 14px rgba(250,204,21,.5)}
.sw-nav-btn.prev:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 20px rgba(250,204,21,.6)}
.sw-nav-btn.submit{background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:3px solid #86efac;box-shadow:0 4px 14px rgba(34,197,94,.4)}
.sw-nav-btn.submit:hover:not(:disabled){transform:translateY(-2px);box-shadow:0 8px 20px rgba(34,197,94,.5)}
.sw-nav-btn.submit.danger{background:linear-gradient(135deg,#ef4444,#b91c1c);border-color:#fca5a5;box-shadow:0 4px 14px rgba(239,68,68,.4)}
.sw-nav-btn:disabled{opacity:.35;cursor:not-allowed;transform:none!important;box-shadow:none!important}
.sw-branch-card{border:3px solid rgba(255,255,255,.15);border-radius:24px;padding:28px 20px;text-align:center;cursor:pointer;transition:.25s;background:white;box-shadow:0 8px 28px rgba(15,23,42,.08);position:relative;overflow:hidden}
.sw-branch-card:hover{transform:translateY(-4px);box-shadow:0 16px 40px rgba(15,23,42,.14);border-color:#3b82f6}
.sw-branch-card.done{border-color:#22c55e;background:linear-gradient(135deg,#f0fdf4,#dcfce7)}
.sw-branch-card.done::after{content:'✅ مكتمل';position:absolute;top:10px;left:10px;background:#22c55e;color:white;font-size:10px;font-weight:800;padding:3px 10px;border-radius:20px;font-family:'Tajawal',sans-serif}
.sw-textarea{width:100%;border:2px solid #e2e8f0;border-radius:14px;padding:14px;font-family:'Tajawal',sans-serif;font-size:15px;resize:vertical;min-height:100px;outline:none;color:#1a1a2e;transition:.2s}
.sw-textarea:focus{border-color:#1e3a8a}
.sw-audio-player{background:#f0f4ff;border:2px solid #c7d2fe;border-radius:14px;padding:16px;margin:12px 0;display:flex;flex-direction:column;align-items:center;gap:10px}
.sw-record-btn{background:linear-gradient(135deg,#db2777,#7e22ce);color:white;border:none;border-radius:50px;padding:12px 28px;font-size:15px;font-weight:700;cursor:pointer;display:flex;align-items:center;gap:8px;transition:.2s;font-family:'Tajawal',sans-serif}
.sw-record-btn:hover{transform:scale(1.05)}
.sw-record-btn.recording{background:linear-gradient(135deg,#ef4444,#dc2626);animation:timerPulse 1s infinite}
.sw-passage{background:#f8fafc;border:1px solid #e2e8f0;border-radius:14px;padding:20px;margin:12px 0;font-size:15px;line-height:1.9;color:#1a1a2e;max-height:250px;overflow-y:auto}
.sw-canvas-wrap{border:2px solid #e2e8f0;border-radius:14px;overflow:hidden;background:white}
.sw-canvas-tools{display:flex;gap:6px;padding:8px 12px;background:#f8fafc;border-bottom:1px solid #e2e8f0;flex-wrap:wrap;align-items:center}
.sw-canvas-tools button{border:1px solid #e2e8f0;background:white;border-radius:8px;padding:5px 10px;font-size:12px;cursor:pointer;transition:.15s}
.sw-canvas-tools button:hover,.sw-canvas-tools button.active{background:#1e3a8a;color:white;border-color:#1e3a8a}
#cheatWarning{position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.92);display:none;align-items:center;justify-content:center}
.cheat-box{background:linear-gradient(135deg,#7f1d1d,#991b1b);border:3px solid #ef4444;border-radius:24px;padding:48px;text-align:center;max-width:480px}
.status-badge{padding:4px 12px;border-radius:20px;font-size:11px;font-weight:700;font-family:'Montserrat',sans-serif;white-space:nowrap}
.status-pending{background:rgba(250,204,21,.2);color:#FACC15;border:1px solid rgba(250,204,21,.4)}
.status-reviewed{background:rgba(74,222,128,.2);color:#4ade80;border:1px solid rgba(74,222,128,.4)}
.status-returned{background:rgba(248,113,113,.2);color:#f87171;border:1px solid rgba(248,113,113,.4)}
.status-approved{background:rgba(74,222,128,.3);color:#4ade80;border:1px solid rgba(74,222,128,.5)}
.tab-btn{padding:8px 20px;border-radius:30px;font-weight:700;font-size:13px;cursor:pointer;transition:.2s;border:2px solid transparent;font-family:'Tajawal',sans-serif}
.tab-btn.active{background:rgba(250,204,21,.2);color:#FACC15;border-color:rgba(250,204,21,.4)}
.tab-btn:not(.active){background:rgba(255,255,255,.05);color:rgba(255,255,255,.6);border-color:rgba(255,255,255,.1)}
.domain-card-wrap{position:relative}
.domain-eye-btn{position:absolute;bottom:10px;left:50%;transform:translateX(-50%);background:rgba(250,204,21,.2);border:1px solid #FACC15;border-radius:20px;padding:4px 14px;font-size:12px;color:#FACC15;cursor:pointer;transition:.2s;white-space:nowrap;z-index:10}
.domain-eye-btn:hover{background:rgba(250,204,21,.4)}
/* Domain & Branch card hovers - CSS only, no JS transform that blocks clicks */
.dom-card{transition:box-shadow .2s,border-color .2s;cursor:pointer}
.dom-card:hover{box-shadow:0 0 0 1px #000, 0 10px 32px rgba(249,115,22,.6)!important}
.br-card{transition:box-shadow .2s,border-color .2s,background .2s;cursor:pointer}
.br-card:hover{background:rgba(56,189,248,.08)!important;border-color:#7dd3fc!important;box-shadow:0 4px 16px rgba(56,189,248,.2)!important}
.coord-nav-btn{width:100%;text-align:right;background:transparent;border:none;color:rgba(255,255,255,.7);padding:10px 14px;border-radius:12px;font-family:'Tajawal',sans-serif;font-size:14px;font-weight:600;cursor:pointer;transition:.2s;display:flex;align-items:center;gap:10px;line-height:1.3}
.coord-nav-btn:hover{background:rgba(255,255,255,.08);color:white}
.coord-nav-btn.active-nav{background:rgba(250,204,21,.15);color:#FACC15;border:1px solid rgba(250,204,21,.3)}
.ordering-editor-item{display:flex;align-items:center;gap:8px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.15);border-radius:12px;padding:10px 14px;margin-bottom:8px}
.ordering-editor-num{background:linear-gradient(135deg,#FACC15,#f59e0b);color:#1e3a8a;font-weight:800;border-radius:6px;min-width:26px;height:26px;display:flex;align-items:center;justify-content:center;font-size:12px;font-family:'Montserrat',sans-serif;flex-shrink:0}
/* ── DISPLAY MODE 2: Focus Card ── */
#studentWindow.dm-mode-2{background:linear-gradient(140deg,#0f172a 0%,#4c1d95 50%,#1e3a8a 100%)!important;}
#studentWindow.dm-mode-2 .sw-header{background:rgba(0,0,0,.45);border-bottom:3px solid rgba(250,204,21,.6);}
#studentWindow.dm-mode-2 .sw-sticky-bar{background:rgba(10,18,40,.97);border-bottom:1px solid rgba(255,255,255,.12);}
#studentWindow.dm-mode-2 .sw-domain-name{color:white;}
#studentWindow.dm-mode-2 #studentWindow.dm-mode-2 #studentWindow.dm-mode-2 #studentWindow.dm-mode-2 .sw-content{display:flex;align-items:flex-start;justify-content:center;background:transparent;padding:24px 16px;}
#studentWindow.dm-mode-2 .sw-frame{max-width:780px;width:100%;border-radius:28px;border:none;box-shadow:0 28px 70px rgba(0,0,0,.5),0 0 0 1px rgba(255,255,255,.08);background:white;}
#studentWindow.dm-mode-2 .sw-footer{background:rgba(10,18,40,.97);border-top:1px solid rgba(255,255,255,.1);}
@keyframes fcEnterRight{from{opacity:0;transform:translateX(70px) scale(.96)}to{opacity:1;transform:translateX(0) scale(1)}}
@keyframes fcEnterLeft{from{opacity:0;transform:translateX(-70px) scale(.96)}to{opacity:1;transform:translateX(0) scale(1)}}
.fc-enter{animation:fcEnterRight .4s cubic-bezier(.16,1,.3,1)}
.fc-enter-back{animation:fcEnterLeft .4s cubic-bezier(.16,1,.3,1)}
/* ── DISPLAY MODE 3: Modern Stream ── */
#studentWindow.dm-mode-3 .sw-content{background:#f5f5f0;padding:20px 16px;}
#studentWindow.dm-mode-3 .sw-footer{display:none!important;}
.stream-q-card{background:white;border-radius:4px;padding:32px 36px;margin-bottom:0;box-shadow:0 1px 4px rgba(0,0,0,.08),0 8px 24px rgba(30,58,138,.06);border:1px solid #e8e8e8;position:relative;}
.stream-q-separator{display:flex;align-items:center;gap:0;margin:0;height:28px;position:relative;}
.stream-q-separator::before{content:'';display:block;flex:1;height:3px;background:linear-gradient(90deg,transparent,#FACC15 20%,#F59E0B 50%,#FACC15 80%,transparent);border-radius:2px;box-shadow:0 1px 6px rgba(245,158,11,.4);}
.stream-q-answered{border-top:3px solid #22c55e;}
.stream-q-answered .stream-q-answered-badge{display:flex!important;}

/* ══ Scholastic Alert Modal ══ */
#scAlertOverlay{position:fixed;inset:0;z-index:999999;background:rgba(15,23,42,.6);backdrop-filter:blur(8px);display:none;align-items:center;justify-content:center;padding:20px}
#scAlertBox{background:linear-gradient(145deg,#f0f9ff,#e0f2fe);border:2px solid #7dd3fc;border-radius:28px;max-width:480px;width:100%;box-shadow:0 20px 60px rgba(14,165,233,.2),0 0 0 4px rgba(125,211,252,.3);overflow:hidden;animation:scPop .35s cubic-bezier(.16,1,.3,1)}
@keyframes scPop{from{opacity:0;transform:scale(.85) translateY(24px)}to{opacity:1;transform:scale(1) translateY(0)}}
.sc-hdr{background:linear-gradient(135deg,#bae6fd,#e0f2fe);border-bottom:2px solid #7dd3fc;padding:22px 26px 18px;display:flex;flex-direction:column;align-items:center;gap:8px;text-align:center}
.sc-icon{font-size:44px;filter:drop-shadow(0 4px 8px rgba(14,165,233,.3))}
.sc-t-ar{font-size:16px;font-weight:800;color:#0c4a6e;font-family:'Tajawal',sans-serif;line-height:1.3;text-align:center}
.sc-t-en{font-size:13px;font-weight:700;color:#0369a1;font-family:'Montserrat',sans-serif;letter-spacing:.3px;text-align:center}
.sc-body{padding:22px 26px 18px}
.sc-en{font-size:14px;font-weight:700;color:#0369a1;font-family:'Montserrat',sans-serif;line-height:1.7;direction:ltr;text-align:left;margin-bottom:10px}
.sc-ar{font-size:14px;font-weight:600;color:#0c4a6e;font-family:'Tajawal',sans-serif;line-height:1.75;direction:rtl;text-align:right;border-top:1px solid #bae6fd;padding-top:10px}
.sc-foot{padding:4px 26px 24px;display:flex;gap:10px;justify-content:flex-end;flex-wrap:wrap}
.sc-btn{border:none;border-radius:14px;padding:12px 28px;font-size:14px;font-weight:800;cursor:pointer;font-family:'Tajawal',sans-serif;transition:all .2s}
.sc-ok{background:linear-gradient(135deg,#0ea5e9,#0284c7);color:white;box-shadow:0 4px 16px rgba(14,165,233,.35)}
.sc-ok:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(14,165,233,.5)}
.sc-cancel{background:white;color:#64748b;border:1px solid #e2e8f0}
.sc-cancel:hover{background:#f1f5f9}
.sc-danger-btn{background:linear-gradient(135deg,#ef4444,#b91c1c);color:white;box-shadow:0 4px 16px rgba(239,68,68,.35)}
.sc-danger-btn:hover{transform:translateY(-2px);box-shadow:0 8px 24px rgba(239,68,68,.5)}
</style>
<body>

<!-- SCHOLASTIC ALERT -->
<div id="scAlertOverlay">
  <div id="scAlertBox">
    <div class="sc-hdr">
      <div class="sc-icon" id="scIcon">💬</div>
      <div><div class="sc-t-ar" id="scTAr">تنبيه</div><div class="sc-t-en" id="scTEn">Notice</div></div>
    </div>
    <div class="sc-body">
      <div class="sc-en" id="scMEn" style="display:none"></div>
      <div class="sc-ar" id="scMAr"></div>
      <textarea id="scPromptInput" style="display:none;width:100%;margin-top:12px;min-height:80px;border-radius:10px;border:1px solid #e2e8f0;padding:10px;font-family:'Tajawal',sans-serif;font-size:14px;resize:vertical" dir="auto"></textarea>
    </div>
    <div class="sc-foot" id="scFoot"></div>
  </div>
</div>
<!-- ANTI-CHEAT -->
<div id="cheatWarning">
  <div class="cheat-box animate__animated animate__shakeX">
    <div style="font-size:48px;margin-bottom:12px">⚠️</div>
    <h2 style="font-size:24px;font-weight:900;color:#fca5a5;margin-bottom:8px">تحذير / Warning</h2>
    <p id="cheatMsg" style="color:rgba(255,255,255,.8);margin-bottom:16px;font-size:14px;line-height:1.6">تم اكتشاف محاولة غش.<br><span style="font-family:'Montserrat',sans-serif;font-size:12px">A cheating attempt was detected.</span></p>
    <div style="font-size:48px;font-weight:900;color:#ef4444;font-family:'Montserrat',sans-serif" id="cheatCountDisplay">0</div>
    <p style="font-size:11px;color:#fca5a5;margin-top:4px">تحذيرات مسجلة / Recorded Warnings</p>
    <button onclick="dismissCheatWarning()" style="margin-top:20px;background:#dc2626;color:white;border:none;border-radius:12px;padding:12px 32px;font-size:15px;font-weight:700;cursor:pointer;font-family:'Tajawal',sans-serif">فهمت — العودة للاختبار / Return to Test</button>
  </div>
</div>

<!-- LOGIN MODAL -->
<div id="loginModal" onclick="closeModal(event)">
  <div class="auth-form" onclick="event.stopPropagation()">
    <div id="loginFields">
      <h3 style="font-size:22px;font-weight:800;margin-bottom:6px;text-align:center;color:#1a1a2e" id="modalTitle">تسجيل دخول</h3>
      <p style="font-size:11px;color:#94a3b8;text-align:center;font-family:'Montserrat',sans-serif;margin-bottom:20px" id="modalTitleEn">Sign In</p>
      <div id="userInputRow">
        <input type="text" id="userInput" placeholder="اسم المستخدم / Username" class="auth-input font-en">
      </div>
      <input type="password" id="passInput" placeholder="كلمة السر / Password" class="auth-input font-en">
      <button onclick="checkAuth()" style="width:100%;background:linear-gradient(135deg,#FACC15,#f59e0b);color:#1e3a8a;font-weight:800;padding:14px;border:none;border-radius:12px;font-size:16px;cursor:pointer;font-family:'Tajawal',sans-serif;margin-bottom:12px;box-shadow:0 4px 14px rgba(250,204,21,.4)">دخول / Sign In</button>
      <p style="font-size:12px;text-align:center;color:#64748b;cursor:pointer" onclick="toggleForgot(true)">نسيت بيانات الدخول؟ / Forgot credentials?</p>
    </div>
    <div id="forgotFields" class="hidden">
      <h3 style="font-size:20px;font-weight:800;margin-bottom:8px;text-align:center;color:#1a1a2e">استعادة البيانات / Recovery</h3>
      <p style="font-size:12px;text-align:center;color:#64748b;margin-bottom:16px;line-height:1.6;font-family:'Montserrat',sans-serif">أدخل بريدك الإلكتروني المسجل وسيتم إرسال بيانات الدخول إليه<br>Enter your registered email to receive your credentials</p>
      <input type="email" id="forgotEmail" placeholder="البريد الإلكتروني / Email" class="auth-input font-en">
      <button onclick="sendForgotPassword()" style="width:100%;background:linear-gradient(135deg,#2563EB,#1d4ed8);color:white;font-weight:800;padding:13px;border:none;border-radius:12px;cursor:pointer;font-family:'Tajawal',sans-serif;margin-bottom:12px;font-size:15px">📧 إرسال / Send</button>
      <p style="font-size:12px;text-align:center;color:#64748b;cursor:pointer" onclick="toggleForgot(false)">← العودة / Back</p>
    </div>
  </div>
</div>

<!-- LOGOUT -->
<div onclick="logout()" class="logout-btn" title="تسجيل خروج / Logout">
  <svg width="22" height="22" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/></svg>
</div>

<!-- TOP NAV: الإدارة + المشرف + المدارس — يمين الشريط العلوي -->
<div id="topLeftNav" style="position:fixed;top:0;right:0;left:0;z-index:200;height:60px;background:#1e3a8a;display:flex;align-items:center;justify-content:space-between;padding:0 24px;box-shadow:0 2px 16px rgba(30,58,138,.25)">

  <!-- يمين: الإدارة + المشرف + المدارس -->
  <div style="display:flex;gap:8px;align-items:center">
    <button onclick="initRoute('الإدارة')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:10px;padding:7px 16px;color:white;font-family:'Tajawal',sans-serif;cursor:pointer;display:flex;align-items:center;gap:6px;font-size:14px;font-weight:800;transition:.2s" onmouseover="this.style.background='rgba(250,204,21,.25)';this.style.borderColor='#FACC15'" onmouseout="this.style.background='rgba(255,255,255,.1)';this.style.borderColor='rgba(255,255,255,.2)'">
      <span>⚙️</span><span>الإدارة</span><span style="font-family:'Montserrat',sans-serif;font-size:9px;opacity:.7;font-weight:600">Admin</span>
    </button>
    <button onclick="initRoute('المشرف')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:10px;padding:7px 16px;color:white;font-family:'Tajawal',sans-serif;cursor:pointer;display:flex;align-items:center;gap:6px;font-size:14px;font-weight:800;transition:.2s" onmouseover="this.style.background='rgba(255,255,255,.18)';this.style.borderColor='rgba(255,255,255,.5)'" onmouseout="this.style.background='rgba(255,255,255,.1)';this.style.borderColor='rgba(255,255,255,.2)'">
      <span>🛡️</span><span>المشرف</span><span style="font-family:'Montserrat',sans-serif;font-size:9px;opacity:.7;font-weight:600">Supervisor</span>
    </button>
    <button onclick="initRoute('المدارس')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:10px;padding:7px 16px;color:white;font-family:'Tajawal',sans-serif;cursor:pointer;display:flex;align-items:center;gap:6px;font-size:14px;font-weight:800;transition:.2s" onmouseover="this.style.background='rgba(255,255,255,.18)';this.style.borderColor='rgba(255,255,255,.5)'" onmouseout="this.style.background='rgba(255,255,255,.1)';this.style.borderColor='rgba(255,255,255,.2)'">
      <span>🏫</span><span>المدارس</span><span style="font-family:'Montserrat',sans-serif;font-size:9px;opacity:.7;font-weight:600">Schools</span>
    </button>
  </div>

  <!-- يسار: لوجو الشركة مصغّر في الشريط -->
  <div style="display:flex;align-items:center;gap:10px;cursor:pointer" onclick="toggleSection('mainDashboard')">
    <div style="text-align:left">
      <div style="font-family:'EB Garamond',Georgia,serif;font-size:20px;font-weight:700;color:#fff;letter-spacing:3px;line-height:1">SCHOLASTIC</div>
      <div style="font-family:'Montserrat',sans-serif;font-size:8px;font-weight:700;color:rgba(255,255,255,.6);letter-spacing:1px;text-transform:uppercase">International Standardised Testing Platform</div>
    </div>
  </div>

</div>

<!-- UPDATE COUNTER MODAL -->
<div id="updateModal" style="display:none;position:fixed;inset:0;z-index:2000;background:rgba(0,0,0,.75);backdrop-filter:blur(8px);align-items:center;justify-content:center">
  <div style="background:linear-gradient(135deg,#001a3a,#002d5c);border:1px solid rgba(250,204,21,.3);border-radius:24px;padding:36px;width:90%;max-width:360px;text-align:center">
    <h3 style="font-size:18px;font-weight:800;color:#FACC15;margin-bottom:16px">🔐 تحديث العداد / Update Counter</h3>
    <input type="password" id="upPass" placeholder="كلمة السر / Password" class="auth-input font-en">
    <input type="number" id="upValue" placeholder="القيمة الجديدة / New Value" class="auth-input font-en">
    <input type="number" id="upDuration" placeholder="مدة الأنيميشن (ثانية)" value="2" class="auth-input font-en">
    <div style="display:flex;gap:10px;justify-content:center;margin-top:8px">
      <button onclick="applyUpdateCounter()" style="background:#FACC15;color:#1e3a8a;font-weight:800;border:none;border-radius:12px;padding:12px 28px;cursor:pointer;font-family:'Tajawal',sans-serif">✅ تحديث</button>
      <button onclick="document.getElementById('updateModal').style.display='none'" style="background:rgba(255,255,255,.1);color:white;border:1px solid rgba(255,255,255,.2);border-radius:12px;padding:12px 20px;cursor:pointer;font-family:'Tajawal',sans-serif">إلغاء</button>
    </div>
    <div id="upError" style="color:#f87171;font-size:13px;margin-top:10px;display:none">كلمة السر غير صحيحة!</div>
  </div>
</div>

<!-- MAIN DASHBOARD -->
<div id="mainDashboard" style="width:100%;min-height:100vh;display:flex;flex-direction:column;align-items:center;background:#eff6ff;padding-top:60px;">

  <!-- اللوجو الرئيسي -->
  <div style="width:100%;max-width:700px;padding:40px 20px 0;text-align:center">
    <svg width="100%" viewBox="0 0 560 115" xmlns="http://www.w3.org/2000/svg" style="display:block;max-width:520px;margin:0 auto">
      <line x1="80"  y1="88" x2="160" y2="60" stroke="#3B5BA5" stroke-width="2.2"/>
      <line x1="160" y1="60" x2="245" y2="85" stroke="#C8A020" stroke-width="2.2"/>
      <line x1="245" y1="85" x2="330" y2="60" stroke="#7B4FA0" stroke-width="2.2"/>
      <line x1="330" y1="60" x2="415" y2="85" stroke="#2E8B6A" stroke-width="2.2"/>
      <line x1="415" y1="85" x2="490" y2="65" stroke="#C0392B" stroke-width="2.2"/>
      <text x="80"  y="14" text-anchor="middle" font-family="Arial,sans-serif" font-size="9.5" font-weight="700" fill="#3B5BA5" letter-spacing="1">ARABIC</text>
      <line x1="80"  y1="17" x2="80"  y2="24" stroke="#3B5BA5" stroke-width="1" stroke-dasharray="2,2"/>
      <text x="160" y="14" text-anchor="middle" font-family="Arial,sans-serif" font-size="9.5" font-weight="700" fill="#C8A020" letter-spacing="1">ISLAMIC</text>
      <line x1="160" y1="17" x2="160" y2="36" stroke="#C8A020" stroke-width="1" stroke-dasharray="2,2"/>
      <text x="245" y="14" text-anchor="middle" font-family="Arial,sans-serif" font-size="9.5" font-weight="700" fill="#7B4FA0" letter-spacing="1">SOCIAL</text>
      <line x1="245" y1="17" x2="245" y2="65" stroke="#7B4FA0" stroke-width="1" stroke-dasharray="2,2"/>
      <text x="330" y="14" text-anchor="middle" font-family="Arial,sans-serif" font-size="9.5" font-weight="700" fill="#2E8B6A" letter-spacing="1">SCIENCE</text>
      <line x1="330" y1="17" x2="330" y2="36" stroke="#2E8B6A" stroke-width="1" stroke-dasharray="2,2"/>
      <text x="415" y="14" text-anchor="middle" font-family="Arial,sans-serif" font-size="9.5" font-weight="700" fill="#C0392B" letter-spacing="1">ENGLISH</text>
      <line x1="415" y1="17" x2="415" y2="55" stroke="#C0392B" stroke-width="1" stroke-dasharray="2,2"/>
      <text x="490" y="14" text-anchor="middle" font-family="Arial,sans-serif" font-size="9" font-weight="700" fill="#4A90C8" letter-spacing="1">MORE</text>
      <line x1="490" y1="17" x2="490" y2="38" stroke="#4A90C8" stroke-width="1" stroke-dasharray="2,2"/>
      <circle cx="80"  cy="88" r="34" fill="#3B5BA5"/>
      <circle cx="160" cy="60" r="24" fill="#C8A020"/>
      <circle cx="245" cy="85" r="20" fill="#7B4FA0"/>
      <circle cx="330" cy="60" r="24" fill="#2E8B6A"/>
      <circle cx="415" cy="85" r="30" fill="#C0392B"/>
      <circle cx="490" cy="65" r="14" fill="#4A90C8"/>
    </svg>
    <div style="font-family:'EB Garamond',Georgia,serif;font-size:clamp(38px,7vw,60px);font-weight:700;color:#1E2D6B;letter-spacing:4px;line-height:1;margin-top:10px">SCHOLASTIC</div>
    <div style="height:3px;background:#C8A020;border-radius:2px;margin:10px auto;width:80%;max-width:420px"></div>
    <div style="font-family:'Montserrat',sans-serif;font-size:11px;font-weight:800;color:#7B1010;letter-spacing:2px;text-transform:uppercase;margin-bottom:6px">INTERNATIONAL STANDARDISED TESTING PLATFORM</div>
    <div style="height:1px;background:#e0e4ef;margin:10px auto;width:65%"></div>
    <div style="font-family:'Tajawal',sans-serif;font-size:16px;font-weight:800;color:#1E2D6B;margin-bottom:0">منصة الاختبارات المعيارية الدولية للمواد المدرسية</div>
  </div>

  <!-- بطاقة دخول الطالب الكبيرة -->
  <div style="width:100%;max-width:560px;padding:32px 20px 0;">
    <button type="button" onclick="initRoute('الطالب')"
      style="width:100%;background:linear-gradient(160deg,#1e3a8a,#2563EB);border:none;border-radius:24px;padding:40px 28px;text-align:center;cursor:pointer;display:flex;flex-direction:column;align-items:center;box-shadow:0 8px 32px rgba(37,99,235,.35);transition:.25s;position:relative;overflow:hidden;"
      onmouseover="this.style.transform='translateY(-3px)';this.style.boxShadow='0 14px 40px rgba(37,99,235,.45)'"
      onmouseout="this.style.transform='';this.style.boxShadow='0 8px 32px rgba(37,99,235,.35)'">
      <div style="position:absolute;top:-30px;right:-30px;width:140px;height:140px;background:rgba(255,255,255,.05);border-radius:50%"></div>
      <div style="position:absolute;bottom:-20px;left:-20px;width:100px;height:100px;background:rgba(255,255,255,.05);border-radius:50%"></div>
      <div style="font-size:64px;margin-bottom:14px;position:relative">🎓</div>
      <div style="font-family:'Tajawal',sans-serif;font-size:32px;font-weight:900;color:#fff;margin-bottom:4px;position:relative">دخول الطالب</div>
      <div style="font-family:'Montserrat',sans-serif;font-size:13px;font-weight:800;color:rgba(255,255,255,.75);letter-spacing:2px;text-transform:uppercase;margin-bottom:18px;position:relative">Student Sign In</div>
      <div style="background:#FACC15;color:#1e3a8a;font-family:'Tajawal',sans-serif;font-size:15px;font-weight:900;padding:10px 32px;border-radius:12px;position:relative">← ادخل الآن / Enter Now</div>
    </button>
  </div>

  <!-- العدادات تحت بوابة الطالب مباشرة -->
  <div style="width:100%;max-width:560px;padding:20px 20px 0;display:grid;grid-template-columns:1fr 1fr;gap:14px;">
    <div class="stat-container" style="background:#fff;border:1.5px solid #bfdbfe;border-radius:16px;padding:18px 20px;box-shadow:0 2px 8px rgba(37,99,235,.08)">
      <div style="display:flex;justify-content:space-between;align-items:center;width:100%">
        <div>
          <div style="font-size:15px;font-weight:900;color:#1e3a8a">المدارس المشتركة</div>
          <div style="font-size:10px;font-family:'Montserrat',sans-serif;color:#93c5fd;text-transform:uppercase;letter-spacing:1px;font-weight:700">Global Partners</div>
        </div>
        <div style="font-size:36px;font-weight:900;color:#2563EB;font-family:'Montserrat',sans-serif" id="schoolsCount">0</div>
      </div>
      <span class="update-link" onclick="showUpdateModal('schoolsCount')" style="color:#93c5fd">Update</span>
    </div>
    <div class="stat-container" style="background:#fff;border:1.5px solid #bfdbfe;border-radius:16px;padding:18px 20px;box-shadow:0 2px 8px rgba(37,99,235,.08)">
      <div style="display:flex;justify-content:space-between;align-items:center;width:100%">
        <div>
          <div style="font-size:15px;font-weight:900;color:#1e3a8a">الاختبارات المكتملة</div>
          <div style="font-size:10px;font-family:'Montserrat',sans-serif;color:#93c5fd;text-transform:uppercase;letter-spacing:1px;font-weight:700">Assessments Completed</div>
        </div>
        <div style="font-size:36px;font-weight:900;color:#2563EB;font-family:'Montserrat',sans-serif" id="examsCount">0</div>
      </div>
      <span class="update-link" onclick="showUpdateModal('examsCount')" style="color:#93c5fd">Update</span>
    </div>
  </div>

  <!-- أيقونات صغيرة: الإدارة + المشرف + المدارس -->
  <div style="width:100%;max-width:560px;padding:16px 20px 0;display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px">
    <button onclick="initRoute('الإدارة')" style="background:#fff;border:1.5px solid #bfdbfe;border-radius:14px;padding:16px 10px;cursor:pointer;text-align:center;transition:.2s;box-shadow:0 2px 8px rgba(37,99,235,.06)"
      onmouseover="this.style.borderColor='#2563EB';this.style.boxShadow='0 4px 16px rgba(37,99,235,.15)'"
      onmouseout="this.style.borderColor='#bfdbfe';this.style.boxShadow='0 2px 8px rgba(37,99,235,.06)'">
      <div style="font-size:26px;margin-bottom:6px">⚙️</div>
      <div style="font-family:'Montserrat',sans-serif;font-size:9px;font-weight:800;color:#2563EB;letter-spacing:.8px;text-transform:uppercase">ADMIN</div>
      <div style="font-family:'Tajawal',sans-serif;font-size:13px;font-weight:800;color:#1e3a8a">الإدارة</div>
    </button>
    <button onclick="initRoute('المشرف')" style="background:#fff;border:1.5px solid #bfdbfe;border-radius:14px;padding:16px 10px;cursor:pointer;text-align:center;transition:.2s;box-shadow:0 2px 8px rgba(37,99,235,.06)"
      onmouseover="this.style.borderColor='#2563EB';this.style.boxShadow='0 4px 16px rgba(37,99,235,.15)'"
      onmouseout="this.style.borderColor='#bfdbfe';this.style.boxShadow='0 2px 8px rgba(37,99,235,.06)'">
      <div style="font-size:26px;margin-bottom:6px">🛡️</div>
      <div style="font-family:'Montserrat',sans-serif;font-size:9px;font-weight:800;color:#2563EB;letter-spacing:.8px;text-transform:uppercase">SUPERVISOR</div>
      <div style="font-family:'Tajawal',sans-serif;font-size:13px;font-weight:800;color:#1e3a8a">المشرف</div>
    </button>
    <button onclick="initRoute('المدارس')" style="background:#fff;border:1.5px solid #bfdbfe;border-radius:14px;padding:16px 10px;cursor:pointer;text-align:center;transition:.2s;box-shadow:0 2px 8px rgba(37,99,235,.06)"
      onmouseover="this.style.borderColor='#2563EB';this.style.boxShadow='0 4px 16px rgba(37,99,235,.15)'"
      onmouseout="this.style.borderColor='#bfdbfe';this.style.boxShadow='0 2px 8px rgba(37,99,235,.06)'">
      <div style="font-size:26px;margin-bottom:6px">🏫</div>
      <div style="font-family:'Montserrat',sans-serif;font-size:9px;font-weight:800;color:#2563EB;letter-spacing:.8px;text-transform:uppercase">SCHOOLS</div>
      <div style="font-family:'Tajawal',sans-serif;font-size:13px;font-weight:800;color:#1e3a8a">المدارس</div>
    </button>
  </div>

  <!-- أعلام الدول -->
  <div style="width:100%;max-width:560px;padding:18px 20px;text-align:center;background:#fff;border-radius:16px;margin:16px 20px;border:1px solid #dbeafe">
    <div class="marquee-wrapper">
      <div class="marquee-content" id="marqueeContent">
        <span class="flag-only">🇸🇦</span><span class="flag-only">🇦🇪</span><span class="flag-only">🇶🇦</span><span class="flag-only">🇪🇬</span><span class="flag-only">🇧🇭</span><span class="flag-only">🇴🇲</span><span class="flag-only">🇦🇺</span><span class="flag-only">🇨🇦</span><span class="flag-only">🇺🇸</span><span class="flag-only">🇯🇵</span>
      </div>
    </div>
  </div>

</div>

<!-- ADMIN PANEL -->
<div id="adminPanel" class="w-full flex flex-col items-center justify-start hidden p-6 overflow-y-auto min-h-screen">
  <h2 style="font-size:clamp(24px,4vw,40px);font-weight:900;margin-bottom:8px;text-align:center" class="animate__animated animate__fadeInDown">لوحة تحكم الإدارة<br><span style="font-family:'Montserrat',sans-serif;color:#FACC15;font-size:16px;font-weight:700">Administration Control Panel</span></h2>
  <div id="adminOptions" class="grid grid-cols-2 md:grid-cols-4 gap-6 w-full max-w-6xl mt-8">
    <div class="glass-card" style="min-height:150px" onclick="showAdminProfile()"><div style="font-size:44px;margin-bottom:10px">👤</div><h4 style="font-size:16px;font-weight:800;margin-bottom:4px">معلوماتي</h4><p style="font-size:10px;font-family:'Montserrat',sans-serif;color:#FACC15;opacity:.7;text-transform:uppercase">My Profile</p></div>
    <div onclick="showSupManager()" class="glass-card" style="min-height:150px"><div style="font-size:44px;margin-bottom:10px">🔑</div><h4 style="font-size:16px;font-weight:800;margin-bottom:4px">إدارة المشرفين</h4><p style="font-size:10px;font-family:'Montserrat',sans-serif;color:#93c5fd;opacity:.7;text-transform:uppercase">Supervisors</p></div>
    <div onclick="showSchoolManager()" class="glass-card" style="min-height:150px"><div style="font-size:44px;margin-bottom:10px">🏫</div><h4 style="font-size:16px;font-weight:800;margin-bottom:4px">إدارة المدارس</h4><p style="font-size:10px;font-family:'Montserrat',sans-serif;color:#86efac;opacity:.7;text-transform:uppercase">Schools</p></div>
    <div onclick="showGeneralReviewer()" class="glass-card" style="min-height:150px;border:1px solid rgba(250,204,21,.4)"><div style="font-size:44px;margin-bottom:10px">🔍</div><h4 style="font-size:16px;font-weight:800;margin-bottom:4px">مراجع الاختبارات</h4><p style="font-size:10px;font-family:'Montserrat',sans-serif;color:#FACC15;opacity:.7;text-transform:uppercase">General Reviewer</p></div>
    <div onclick="openGradingCommittee('admin')" class="glass-card" style="min-height:150px;border:1px solid rgba(251,146,60,.4)"><div style="font-size:44px;margin-bottom:10px">📋</div><h4 style="font-size:16px;font-weight:800;margin-bottom:4px">اعتمادية التصحيح</h4><p style="font-size:10px;font-family:'Montserrat',sans-serif;color:#fb923c;opacity:.7;text-transform:uppercase">Grading Approval</p></div>
    <div onclick="openArchivePanel()" class="glass-card" style="min-height:150px;border:1px solid rgba(250,204,21,.3)"><div style="font-size:44px;margin-bottom:10px">🗄️</div><h4 style="font-size:16px;font-weight:800;margin-bottom:4px">أرشيف الاختبارات</h4><p style="font-size:10px;font-family:'Montserrat',sans-serif;color:#FACC15;opacity:.7;text-transform:uppercase">Test Archive</p></div>
  </div>
  <!-- Sup Manager -->
  <div id="supManager" class="hidden w-full max-w-5xl p-6 rounded-3xl backdrop-blur-xl border border-white/10 mt-6" style="max-height:75vh;overflow-y:auto">
    <h3 style="font-size:22px;font-weight:800;color:#FACC15;margin-bottom:16px">🔑 إدارة المشرفين / Supervisors</h3>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
      <div><label class="wizard-label">الاسم / Name</label><input type="text" id="supName" class="wizard-input" placeholder="اسم المشرف"></div>
      <div><label class="wizard-label">الإيميل / Email</label><input type="email" id="supEmail" class="wizard-input font-en" placeholder="email@school.com"></div>
      <div class="md:col-span-2"><button onclick="createSup()" style="width:100%;background:#FACC15;color:#1e3a8a;font-weight:800;border:none;border-radius:12px;padding:14px;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif">➕ إنشاء حساب مشرف / Create Account</button></div>
    </div>
    <div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>الاسم</th><th>Username</th><th>Password</th><th>الحالة</th><th>حذف</th></tr></thead><tbody id="supList"></tbody></table></div>
    <button onclick="hideSubPanel('supManager')" style="margin-top:16px;color:rgba(255,255,255,.6);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">← العودة / Back</button>
  </div>
  <!-- School Manager -->
  <div id="schoolManager" class="hidden w-full max-w-6xl p-6 rounded-3xl backdrop-blur-xl border border-white/10 mt-6" style="max-height:80vh;overflow-y:auto">
    <h3 style="font-size:22px;font-weight:800;color:#FACC15;margin-bottom:20px">🏫 إدارة المدارس / Schools</h3>
    <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:20px;padding:24px;margin-bottom:20px">
      <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
        <div><label class="wizard-label">اسم المدرسة <span class="wizard-label-en">School Name</span></label><input type="text" id="newSchoolName" class="wizard-input" placeholder="الاسم الكامل" oninput="previewNewSchoolCode()"></div>
        <div><label class="wizard-label">الدولة <span class="wizard-label-en">Country</span></label><select id="newSchoolCountry" class="wizard-input" onchange="previewNewSchoolCode()"></select></div>
        <div><label class="wizard-label">المنهاج <span class="wizard-label-en">Curriculum</span></label>
          <select id="newSchoolCurriculum" class="wizard-input" onchange="previewNewSchoolCode()"><option value="">-- اختر / Select --</option><option value="IGCSE">🇬🇧 البريطاني / IGCSE</option><option value="American">🇺🇸 الأمريكي / American</option><option value="IB">🌐 IB</option><option value="Canadian">🇨🇦 الكندي / Canadian</option><option value="CBSE">🇮🇳 CBSE</option><option value="Local">📍 المحلي</option></select></div>
        <div><label class="wizard-label">كود المدرسة <span class="wizard-label-en">School Code</span></label>
          <div style="background:rgba(250,204,21,.1);border:1.5px dashed #FACC15;border-radius:12px;padding:11px 14px;display:flex;align-items:center;justify-content:center;height:44px">
            <span id="newSchoolCodePreview" style="font-family:Montserrat,sans-serif;font-weight:900;font-size:15px;color:#FACC15;letter-spacing:1px">—</span>
          </div>
        </div>
        <div>
          <label class="wizard-label">شعار المدرسة <span class="wizard-label-en">Logo</span></label>
          <div style="display:flex;gap:8px;margin-bottom:8px">
            <button onclick="document.getElementById('newSchoolLogo').click()" style="flex:1;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px;font-size:12px;cursor:pointer;font-family:'Tajawal',sans-serif">📁 من الجهاز</button>
            <button onclick="addLogoFromURL('newSchoolLogoPreview','newSchoolImgState')" style="flex:1;background:rgba(96,165,250,.2);border:1px solid rgba(96,165,250,.4);color:white;border-radius:10px;padding:8px;font-size:12px;cursor:pointer;font-family:'Tajawal',sans-serif">🌐 URL</button>
          </div>
          <div class="logo-upload-area" onclick="document.getElementById('newSchoolLogo').click()">
            <img id="newSchoolLogoPreview" src="" alt="" style="display:none;max-height:60px;object-fit:contain;border-radius:6px;margin-bottom:4px">
            <div id="newSchoolLogoPlaceholder"><div style="font-size:28px;margin-bottom:4px">📤</div><div style="font-size:12px;color:#94a3b8">اضغط لرفع الشعار</div></div>
          </div>
          <input type="file" id="newSchoolLogo" accept="image/*" style="display:none" onchange="previewNewSchoolLogo(event)">
          <div id="newSchoolImgControls" style="display:none;gap:6px;flex-wrap:wrap;align-items:center;margin-top:8px">
            <button onclick="imgCtrlFor('newSchoolLogoPreview','newSchoolImgState','zi')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">🔍+</button>
            <button onclick="imgCtrlFor('newSchoolLogoPreview','newSchoolImgState','zo')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">🔍−</button>
            <button onclick="imgCtrlFor('newSchoolLogoPreview','newSchoolImgState','rr')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↻</button>
            <button onclick="imgCtrlFor('newSchoolLogoPreview','newSchoolImgState','rl')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↺</button>
            <button onclick="imgCtrlFor('newSchoolLogoPreview','newSchoolImgState','rs')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↩</button>
            <button onclick="confirmLogoFinal('newSchoolLogoPreview','newSchoolImgControls')" style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);color:#4ade80;border-radius:7px;padding:3px 10px;font-size:12px;cursor:pointer">✅ تثبيت</button>
          </div>
        </div>
      </div>
      <button onclick="addSchool()" style="width:100%;background:#22c55e;color:white;font-weight:800;border:none;border-radius:12px;padding:14px;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif">➕ إضافة المدرسة / Add School</button>
    </div>
    <div style="overflow-x:auto"><table class="admin-table"><thead><tr><th>#</th><th>الشعار</th><th>اسم المدرسة</th><th>الدولة</th><th>المنهاج</th><th>كود المدرسة</th><th>Username</th><th>Password</th><th>التاريخ</th><th>حذف</th></tr></thead><tbody id="schoolList"></tbody></table></div>
    <button onclick="hideSubPanel('schoolManager')" style="margin-top:16px;color:rgba(255,255,255,.6);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">← العودة / Back</button>
  </div>
  <!-- General Reviewer -->
  <div id="generalReviewerPanel" class="hidden w-full max-w-6xl p-6 rounded-3xl backdrop-blur-xl border border-white/10 mt-6" style="max-height:80vh;overflow-y:auto">
    <h3 style="font-size:22px;font-weight:800;color:#FACC15;margin-bottom:16px">🔍 مراجع الاختبارات العام / General Reviewer</h3>
    <div id="generalReviewerContent"></div>
    <button onclick="hideSubPanel('generalReviewerPanel')" style="margin-top:16px;color:rgba(255,255,255,.6);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">← العودة / Back</button>
  </div>
  <!-- Admin Grading Committee Panel -->
  <div id="adminGradingCommitteePanel" class="hidden w-full max-w-6xl bg-black/40 backdrop-blur-xl p-6 rounded-3xl border border-white/10 mt-6" style="max-height:90vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px">
      <div>
        <h3 style="font-size:22px;font-weight:800;color:#fb923c">📋 اعتمادية التصحيح</h3>
        <p style="font-size:11px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Grading Approval Panel</p>
      </div>
      <button onclick="hideSubPanel('adminGradingCommitteePanel')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;cursor:pointer;font-size:13px">✕ إغلاق</button>
    </div>
    <div id="adminCommitteeContent"></div>
  </div>
  <button onclick="backToMain()" style="margin-top:40px;color:rgba(255,255,255,.6);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">خروج من الإدارة / Exit</button>
</div>

<!-- SUPERVISOR PANEL -->
<div id="supervisorPanel" class="w-full flex flex-col items-center justify-start hidden p-4 overflow-y-auto min-h-screen">
  <h2 id="supTitle" style="font-size:clamp(22px,4vw,36px);font-weight:900;margin:24px 0 8px;text-align:center" class="animate__animated animate__fadeInDown">لوحة المشرف الأكاديمي<br><span style="font-family:'Montserrat',sans-serif;color:#FACC15;font-size:16px;font-weight:700">Academic Supervisor Panel</span></h2>
  <div id="supMainOptions" class="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-5 w-full max-w-7xl">
    <div class="glass-card" onclick="startTestWizard()" style="min-height:150px"><div style="font-size:40px;margin-bottom:10px">📝</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">إنشاء اختبار</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:rgba(255,255,255,.6);text-transform:uppercase">Create Test</p></div>
    <div class="glass-card" onclick="openReviewerPanel()" style="min-height:150px;border:1px solid rgba(250,204,21,.3)"><div style="font-size:40px;margin-bottom:10px">🔍</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">مراجع الاختبارات</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:#FACC15;text-transform:uppercase">Reviewer</p></div>
    <div class="glass-card" onclick="openSupLiveMonitor()" style="min-height:150px;border:1px solid rgba(56,189,248,.3)"><div style="font-size:40px;margin-bottom:10px">📡</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">المراقبة المباشرة</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:#7dd3fc;text-transform:uppercase">Live Monitor</p></div>
    <div class="glass-card" onclick="openGradingPanel('sup')" style="min-height:150px;border:1px solid rgba(251,146,60,.3)"><div style="font-size:40px;margin-bottom:10px">✍️</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">التصحيح المباشر</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:#fb923c;text-transform:uppercase">Live Marking</p></div>
    <div class="glass-card" onclick="openGradingCommittee('sup')" style="min-height:150px;border:1px solid rgba(167,139,250,.3)"><div style="font-size:40px;margin-bottom:10px">👥</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">لجنة التصحيح</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:#c4b5fd;text-transform:uppercase">Grading Committee</p></div>
    <div onclick="openStandardsBank()" class="glass-card" style="min-height:150px;border:1px solid rgba(250,204,21,.3)"><div style="font-size:40px;margin-bottom:10px">🎯</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">بنك المعايير</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:#FACC15;text-transform:uppercase">Standards</p></div>
    <div onclick="openMyTests()" class="glass-card" style="min-height:150px;border:1px solid rgba(74,222,128,.3)"><div style="font-size:40px;margin-bottom:10px">📂</div><h4 style="font-size:15px;font-weight:800;margin-bottom:4px">اختباراتي</h4><p style="font-size:9px;font-family:'Montserrat',sans-serif;color:#86efac;text-transform:uppercase">My Tests</p></div>
  </div>
  <!-- WIZARD -->
  <div id="testCreationWizard" class="hidden w-full max-w-5xl bg-black/40 backdrop-blur-xl p-8 rounded-3xl border border-white/20 shadow-2xl mt-6 text-right">
    <div class="step-indicator"><div class="step-dot active" id="dot1"></div><div class="step-dot" id="dot2"></div><div class="step-dot" id="dot3"></div></div>
    <!-- Step 1 -->
    <div id="wizardStep1" class="wizard-step">
      <h3 style="font-size:20px;font-weight:800;color:#FACC15;border-bottom:1px solid rgba(255,255,255,.1);padding-bottom:12px;margin-bottom:20px">١. المعلومات الأساسية <span style="font-family:'Montserrat',sans-serif;font-size:12px;color:rgba(255,255,255,.4)">Basic Information</span></h3>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-5">
        <div><label class="wizard-label">١- الدولة <span class="wizard-label-en">Country</span></label><select id="country" class="wizard-input" onchange="filterSchoolsByCountryAndCurriculum()"></select></div>
        <div><label class="wizard-label">٢- المنهاج <span class="wizard-label-en">Curriculum</span></label>
          <select id="curriculum" class="wizard-input" onchange="checkLocalCurriculum();filterSchoolsByCountryAndCurriculum()">
            <option value="">-- اختر / Select --</option><option value="IGCSE">🇬🇧 البريطاني / IGCSE</option><option value="American">🇺🇸 الأمريكي / American</option><option value="IB">🌐 البكالوريا الدولية / IB</option><option value="Canadian">🇨🇦 الكندي / Canadian</option><option value="CBSE">🇮🇳 CBSE / الهندي</option><option value="Local">📍 المحلي / Local</option>
          </select>
          <div id="localCurriculumRow" class="hidden mt-2"><label class="wizard-label" style="color:#fb923c">دولة المنهاج المحلي / Local Curriculum Country</label><select id="localCountry" class="wizard-input"></select></div>
        </div>
        <div class="md:col-span-2"><label class="wizard-label">٣- المدارس <span class="wizard-label-en">Schools</span></label><div class="schools-multi" id="schoolsMultiList"><div style="color:rgba(255,255,255,.3);font-size:13px;padding:8px;text-align:center">اختر الدولة والمنهاج أولاً</div></div><div class="selected-schools-tags" id="selectedSchoolsTags"></div></div>
        <div><label class="wizard-label">٤- شعار المدرسة <span class="wizard-label-en">Logo</span></label>
          <div style="display:flex;gap:8px;margin-bottom:8px">
            <button onclick="document.getElementById('logoFile').click()" style="flex:1;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px;font-size:12px;cursor:pointer;font-family:'Tajawal',sans-serif">📁 من الجهاز</button>
            <button onclick="addLogoFromURL('logoPreview','wizardLogoState')" style="flex:1;background:rgba(96,165,250,.2);border:1px solid rgba(96,165,250,.4);color:white;border-radius:10px;padding:8px;font-size:12px;cursor:pointer;font-family:'Tajawal',sans-serif">🌐 URL</button>
          </div>
          <div class="logo-upload-area" onclick="document.getElementById('logoFile').click()">
            <img id="logoPreview" src="" alt="" style="display:none;max-height:60px;object-fit:contain;margin-bottom:8px">
            <div id="logoPlaceholder"><div style="font-size:24px;margin-bottom:4px">📤</div><div style="font-size:11px;color:#94a3b8">اضغط لرفع الشعار</div></div>
          </div>
          <input type="file" id="logoFile" accept="image/*" style="display:none" onchange="previewLogo(event)">
          <div id="wizardLogoControls" style="display:none;gap:6px;flex-wrap:wrap;align-items:center;margin-top:8px">
            <button onclick="imgCtrlFor('logoPreview','wizardLogoState','zi')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">🔍+</button>
            <button onclick="imgCtrlFor('logoPreview','wizardLogoState','zo')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">🔍−</button>
            <button onclick="imgCtrlFor('logoPreview','wizardLogoState','rr')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↻</button>
            <button onclick="imgCtrlFor('logoPreview','wizardLogoState','rl')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↺</button>
            <button onclick="imgCtrlFor('logoPreview','wizardLogoState','rs')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↩</button>
            <button onclick="confirmLogoFinal('logoPreview','wizardLogoControls')" style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);color:#4ade80;border-radius:7px;padding:3px 10px;font-size:12px;cursor:pointer">✅ تثبيت</button>
          </div>
        </div>
        <div><label class="wizard-label">٥- الصف <span class="wizard-label-en">Grade</span></label>
          <select id="grade" class="wizard-input"><option value="">-- اختر --</option><option>KS1/FS1</option><option>KS1/FS2</option><option>Grade 1/Year 2</option><option>Grade 2/Year 3</option><option>Grade 3/Year 4</option><option>Grade 4/Year 5</option><option>Grade 5/Year 6</option><option>Grade 6/Year 7</option><option>Grade 7/Year 8</option><option>Grade 8/Year 9</option><option>Grade 9/Year 10</option><option>Grade 10/Year 11</option><option>Grade 11/Year 12</option><option>Grade 12/Year 13</option></select></div>
        <div><label class="wizard-label">٦- السنة الدراسية</label><select id="academicYear" class="wizard-input"></select></div>
        <div><label class="wizard-label">٧- المادة <span class="wizard-label-en">Subject</span></label>
          <select id="subject" class="wizard-input" onchange="handleSubjectChange()">
            <option value="">-- اختر --</option>
            <option value="arabic_arabs">العربية للناطقين / Arabic Native</option>
            <option value="arabic_non">العربية لغير الناطقين / Arabic Non-Native</option>
            <option value="islamic_arabs">التربية الإسلامية للناطقين / Islamic (Arabic)</option>
            <option value="islamic_non">التربية الإسلامية لغير الناطقين / Islamic (English)</option>
            <option value="social_arabs">الدراسات الاجتماعية للناطقين / Social (Arabic)</option>
            <option value="social_non">الدراسات الاجتماعية لغير الناطقين / Social (English)</option>
            <option value="english_1st">الإنجليزية لغة أولى / English 1st Lang</option>
            <option value="english_2nd">الإنجليزية لغة ثانية / English 2nd Lang</option>
            <option value="french_2nd">الفرنسية لغة ثانية / French 2nd Lang</option>
            <option value="math_arabs">الرياضيات للناطقين / Math (Arabic)</option>
            <option value="math_non">الرياضيات لغير الناطقين / Math (English)</option>
            <option value="science_arabs">العلوم للناطقين / Science (Arabic)</option>
            <option value="science_non">العلوم لغير الناطقين / Science (English)</option>
            <option value="history_national">التاريخ الوطني / History National</option>
            <option value="history_international">التاريخ الدولي / History International</option>
            <option value="other">أخرى / Other</option>
          </select>
          <div id="historyNationalRow" class="hidden mt-2"><label class="wizard-label" style="color:#fb923c">دولة المنهاج الوطني</label><select id="historyCountry" class="wizard-input"><option value="">-- اختر --</option><option value="ae">🇦🇪 UAE</option><option value="sa">🇸🇦 Saudi</option><option value="qa">🇶🇦 Qatar</option><option value="eg">🇪🇬 Egypt</option></select></div>
          <div id="subjectOtherRow" class="hidden mt-2"><label class="wizard-label" style="color:#fb923c">اسم المادة</label><input type="text" id="subjectOther" class="wizard-input" placeholder="اكتب اسم المادة"></div>
        </div>
        <div><label class="wizard-label">٨- الفصل <span class="wizard-label-en">Term</span></label><select id="term" class="wizard-input"><option value="">-- اختر --</option><option value="1">الفصل الأول / Term 1</option><option value="2">الفصل الثاني / Term 2</option><option value="3">الفصل الثالث / Term 3</option></select></div>
        <div class="md:col-span-2"><label class="wizard-label">٩- اسم الاختبار</label><input type="text" id="testName" class="wizard-input" placeholder="مثال: اختبار نهاية الفصل الأول" oninput="generateTestLink()"><div id="testLinkBox" class="link-badge hidden mt-2"><span style="color:rgba(255,255,255,.5);font-size:11px;display:block;margin-bottom:4px">🔗 رابط الاختبار:</span><span id="testLinkText" class="font-en"></span></div></div>
        <div class="md:col-span-2"><label class="wizard-label">١٠- عدد المجالات <span class="wizard-label-en">Domains</span></label><input type="number" id="domQty" min="1" max="10" class="wizard-input" placeholder="أدخل عدد المجالات" oninput="generateDomains()"></div>
      </div>
      <!-- ١١. اختيار نمط عرض الاختبار -->
      <div style="margin-top:20px;background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:16px">
        <label class="wizard-label" style="color:#FACC15;display:block;margin-bottom:12px">١١- نمط عرض الاختبار للطالب <span style="font-family:Montserrat,sans-serif;font-size:10px;color:rgba(255,255,255,.4)">Student Display Mode</span></label>
        <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px">
          <div id="dm1" onclick="selectDisplayModeStep1(1)" style="border:2px solid rgba(255,255,255,.12);border-radius:12px;overflow:hidden;cursor:pointer;transition:.2s">
            <div style="background:linear-gradient(135deg,#1e3a8a,#1d4ed8);padding:10px;text-align:center"><div style="font-size:20px">📋</div><div style="font-size:11px;font-weight:800;color:white;margin-top:3px">الكلاسيكي</div><div style="font-size:8px;color:rgba(255,255,255,.55);font-family:Montserrat,sans-serif">Classic</div></div>
            <div id="dm1-check" style="display:none;text-align:center;padding:4px;background:rgba(34,197,94,.2);border-top:1px solid rgba(34,197,94,.3)"><span style="color:#4ade80;font-size:10px;font-weight:800">✓ محدد</span></div>
          </div>
          <div id="dm4" onclick="selectDisplayModeStep1(4)" style="border:2px solid rgba(255,255,255,.12);border-radius:12px;overflow:hidden;cursor:pointer;transition:.2s">
            <div style="background:linear-gradient(135deg,#374151,#111827);padding:10px;text-align:center"><div style="font-size:20px">📄</div><div style="font-size:11px;font-weight:800;color:white;margin-top:3px">الورقة البيضاء</div><div style="font-size:8px;color:rgba(255,255,255,.55);font-family:Montserrat,sans-serif">White Paper</div></div>
            <div id="dm4-check" style="display:none;text-align:center;padding:4px;background:rgba(34,197,94,.2);border-top:1px solid rgba(34,197,94,.3)"><span style="color:#4ade80;font-size:10px;font-weight:800">✓ محدد</span></div>
          </div>
          <div style="border:2px solid rgba(255,255,255,.05);border-radius:12px;overflow:hidden;opacity:.35;cursor:not-allowed">
            <div style="background:linear-gradient(135deg,#4b2d8a,#3b1f6e);padding:10px;text-align:center"><div style="font-size:20px">🎯</div><div style="font-size:11px;font-weight:800;color:white;margin-top:3px">التركيز</div><div style="font-size:8px;color:rgba(255,255,255,.55);font-family:Montserrat,sans-serif">Soon</div></div>
          </div>
          <div style="border:2px solid rgba(255,255,255,.05);border-radius:12px;overflow:hidden;opacity:.35;cursor:not-allowed">
            <div style="background:linear-gradient(135deg,#065f75,#0a3d4a);padding:10px;text-align:center"><div style="font-size:20px">🌊</div><div style="font-size:11px;font-weight:800;color:white;margin-top:3px">الانسيابي</div><div style="font-size:8px;color:rgba(255,255,255,.55);font-family:Montserrat,sans-serif">Soon</div></div>
          </div>
        </div>
      </div>
      <div id="domContainer" class="mt-4"></div>
      <!-- Reviewers -->
      <div class="mt-6">
        <h4 style="font-size:15px;font-weight:800;color:#FACC15;border-top:1px solid rgba(255,255,255,.1);padding-top:16px;margin-bottom:12px">١٢- نظام التوثيق <span style="font-family:'Montserrat',sans-serif;font-size:11px;color:rgba(255,255,255,.4)">Audit Trail</span></h4>
        <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
          <div class="reviewer-box"><h6>✍️ محرر الاختبار / Author</h6><input type="text" id="reviewer0" class="wizard-input" readonly style="opacity:.65;cursor:not-allowed"></div>
          <div class="reviewer-box"><h6>🔍 المراجع المختار / Reviewer</h6><select id="reviewer1" class="wizard-input" onchange="previewReviewerNotice()"><option value="">-- اختر --</option></select></div>
          <div class="reviewer-box"><h6>🏛️ المراجع العام / General</h6><input type="text" value="مراجع الاختبارات العام" class="wizard-input" readonly style="opacity:.65;color:#FACC15"></div>
        </div>
        <div id="reviewerNoticeBox" class="hidden mt-3" style="background:rgba(30,58,138,.3);border:1px solid rgba(96,165,250,.3);border-radius:14px;padding:12px;font-size:13px;color:#93c5fd">📨 سيتم إرسال الاختبار للمراجع والمراجع العام / Will be sent to reviewer and general reviewer</div>
      </div>
      <div style="display:flex;gap:12px;margin-top:24px;justify-content:space-between;align-items:center">
        <button onclick="cancelWizard()" style="display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.8);border-radius:14px;padding:10px 20px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px;font-weight:700">🏠 الرئيسية</button>
        <button onclick="goToStep(2)" style="background:#FACC15;color:#1e3a8a;padding:12px 32px;border:none;border-radius:30px;font-weight:800;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif">التالي ←</button>
      </div>
    </div>
    <!-- Step 2 -->
    <div id="wizardStep2" class="wizard-step hidden">
      <h3 style="font-size:20px;font-weight:800;color:#FACC15;border-bottom:1px solid rgba(255,255,255,.1);padding-bottom:12px;margin-bottom:20px">٢. تعليمات الاختبار <span style="font-family:'Montserrat',sans-serif;font-size:12px;color:rgba(255,255,255,.4)">Instructions</span></h3>
      <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
        <div>
          <label class="wizard-label" style="font-size:13px;margin-bottom:8px">📝 التعليمات بالعربية</label>
          <div class="rtoolbar"><button onclick="execRich('insAr','bold')"><b>B</b></button><button onclick="execRich('insAr','italic')"><i>I</i></button><button onclick="execRich('insAr','underline')"><u>U</u></button><div class="sep"></div><select onchange="execRich('insAr','fontSize',this.value);this.value=''"><option value="">حجم</option><option value="1">10</option><option value="2">12</option><option value="3">14</option><option value="4">16</option><option value="5">18</option><option value="6">22</option><option value="7">26</option></select><select onchange="execRich('insAr','fontName',this.value);this.value=''"><option value="">خط</option><option value="Tajawal">Tajawal</option><option value="Amiri">Amiri</option></select><div class="sep"></div><input type="color" onchange="execRich('insAr','foreColor',this.value)"><div class="sep"></div><button onclick="execRich('insAr','justifyRight')">⇥</button><button onclick="execRich('insAr','justifyCenter')">⊟</button><button onclick="execRich('insAr','justifyLeft')">⇤</button><button onclick="execRich('insAr','insertUnorderedList')">•</button></div>
          <div id="insAr" contenteditable="true" dir="rtl" class="rich-editor" style="min-height:160px" placeholder="اكتب التعليمات بالعربية..."></div>
        </div>
        <div>
          <label class="wizard-label" style="font-size:13px;margin-bottom:8px">📝 Instructions in English</label>
          <div class="rtoolbar"><button onclick="execRich('insEn','bold')"><b>B</b></button><button onclick="execRich('insEn','italic')"><i>I</i></button><button onclick="execRich('insEn','underline')"><u>U</u></button><div class="sep"></div><select onchange="execRich('insEn','fontSize',this.value);this.value=''"><option value="">Size</option><option value="1">10</option><option value="2">12</option><option value="3">14</option><option value="4">16</option><option value="5">18</option><option value="6">22</option><option value="7">26</option></select><select onchange="execRich('insEn','fontName',this.value);this.value=''"><option value="">Font</option><option value="Montserrat">Montserrat</option><option value="Arial">Arial</option></select><div class="sep"></div><input type="color" onchange="execRich('insEn','foreColor',this.value)"><div class="sep"></div><button onclick="execRich('insEn','justifyLeft')">⇤</button><button onclick="execRich('insEn','justifyCenter')">⊟</button><button onclick="execRich('insEn','justifyRight')">⇥</button><button onclick="execRich('insEn','insertUnorderedList')">•</button></div>
          <div id="insEn" contenteditable="true" dir="ltr" class="rich-editor font-en" style="min-height:160px" placeholder="Write instructions..."></div>
        </div>
      </div>
      <div style="display:flex;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-top:24px">
        <div style="display:flex;gap:10px">
          <button onclick="cancelWizard()" style="display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.8);border-radius:14px;padding:10px 20px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px;font-weight:700">🏠 الرئيسية</button>
          <button onclick="goToStep(1)" style="display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.15);color:rgba(255,255,255,.7);border-radius:14px;padding:10px 20px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">← رجوع</button>
        </div>
        <div style="display:flex;gap:10px;flex-wrap:wrap">
          <button onclick="previewCurrentInstructions()" style="background:rgba(96,165,250,.2);color:#93c5fd;padding:12px 24px;border:none;border-radius:30px;font-weight:800;font-size:14px;cursor:pointer;font-family:'Tajawal',sans-serif">👁 معاينة التعليمات</button>
          <button onclick="goToStep(3)" style="background:#22c55e;color:white;padding:12px 32px;border:none;border-radius:30px;font-weight:800;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif">إنشاء المجالات ✓</button>
        </div>
      </div>
    </div>
    <!-- Step 3 -->
    <div id="wizardStep3" class="wizard-step hidden">
      <h3 style="font-size:20px;font-weight:800;color:#FACC15;border-bottom:1px solid rgba(255,255,255,.1);padding-bottom:12px;margin-bottom:8px;text-align:center">٣. مجالات الاختبار <span style="font-family:'Montserrat',sans-serif;font-size:12px;color:rgba(255,255,255,.4)">Domains</span></h3>
      <p style="text-align:center;color:rgba(255,255,255,.5);font-size:13px;margin-bottom:24px">اضغط على كل مجال لإعداده / Click each domain to configure</p>
      <div id="domainsGrid" class="grid grid-cols-2 md:grid-cols-4 gap-6"></div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:32px;flex-wrap:wrap;gap:12px">
        <div style="display:flex;gap:10px">
          <button onclick="cancelWizard()" style="display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.8);border-radius:14px;padding:10px 20px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px;font-weight:700">🏠 الرئيسية</button>
          <button onclick="goToStep(2)" style="display:flex;align-items:center;gap:6px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.15);color:rgba(255,255,255,.7);border-radius:14px;padding:10px 20px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">← رجوع</button>
        </div>
        <button id="approveBtn" onclick="approveFinalTest()" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;padding:14px 40px;border:3px solid #86efac;border-radius:30px;font-weight:800;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif;box-shadow:0 4px 16px rgba(34,197,94,.4);display:flex;align-items:center;gap:8px"><span style="font-size:20px">✅</span><span>اعتماد وإرسال / Submit for Review</span></button>
      </div>
    </div>
  </div>
  <!-- Archive Panel -->
  <div id="archivePanel" class="hidden w-full max-w-6xl bg-black/40 backdrop-blur-xl p-6 rounded-3xl border border-white/10 mt-6" style="max-height:90vh;overflow-y:auto">
    <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;flex-wrap:wrap;gap:10px">
      <div><h3 style="font-size:22px;font-weight:800;color:#FACC15">🗄️ أرشيف الاختبارات</h3><p style="font-size:11px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Test Archive</p></div>
      <button onclick="hideSubPanel('archivePanel')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;cursor:pointer;font-size:13px">✕ إغلاق</button>
    </div>
    <!-- محرك البحث -->
    <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:16px;margin-bottom:20px">
      <div style="font-size:13px;font-weight:800;color:#FACC15;margin-bottom:12px">🔍 محرك البحث / Search</div>
      <div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:10px">
        <div><label class="wizard-label" style="font-size:10px">الدولة / Country</label><select id="arch-country" class="wizard-input" style="font-size:12px"><option value="">الكل</option></select></div>
        <div><label class="wizard-label" style="font-size:10px">المنهاج / Curriculum</label><select id="arch-curriculum" class="wizard-input" style="font-size:12px"><option value="">الكل</option><option value="igcse">🇬🇧 IGCSE</option><option value="american">🇺🇸 American</option><option value="ib">🌐 IB</option><option value="canadian">🇨🇦 Canadian</option><option value="cbse">🇮🇳 CBSE</option><option value="local">📍 Local</option></select></div>
        <div><label class="wizard-label" style="font-size:10px">كود المدرسة / School Code</label><input type="text" id="arch-schoolcode" class="wizard-input" style="font-size:12px" placeholder="مثال: UAE-25-001"></div>
        <div><label class="wizard-label" style="font-size:10px">المدرسة / School</label><select id="arch-school" class="wizard-input" style="font-size:12px"><option value="">الكل</option></select></div>
        <div><label class="wizard-label" style="font-size:10px">الصف / Grade</label><select id="arch-grade" class="wizard-input" style="font-size:12px"><option value="">الكل</option><option value="KS1/FS1">KS1/FS1</option><option value="Grade 1/Year 2">Grade 1/Year 2</option><option value="Grade 2/Year 3">Grade 2/Year 3</option><option value="Grade 3/Year 4">Grade 3/Year 4</option><option value="Grade 4/Year 5">Grade 4/Year 5</option><option value="Grade 5/Year 6">Grade 5/Year 6</option><option value="Grade 6/Year 7">Grade 6/Year 7</option><option value="Grade 7/Year 8">Grade 7/Year 8</option><option value="Grade 8/Year 9">Grade 8/Year 9</option><option value="Grade 9/Year 10">Grade 9/Year 10</option><option value="Grade 10/Year 11">Grade 10/Year 11</option><option value="Grade 11/Year 12">Grade 11/Year 12</option><option value="Grade 12/Year 13">Grade 12/Year 13</option></select></div>
        <div><label class="wizard-label" style="font-size:10px">المادة / Subject</label><input type="text" id="arch-subject" class="wizard-input" style="font-size:12px" placeholder="مثال: Arabic"></div>
        <div><label class="wizard-label" style="font-size:10px">الفصل / Term</label><select id="arch-term" class="wizard-input" style="font-size:12px"><option value="">الكل</option><option value="1">Term 1</option><option value="2">Term 2</option><option value="3">Term 3</option></select></div>
        <div><label class="wizard-label" style="font-size:10px">السنة / Year</label><input type="text" id="arch-year" class="wizard-input font-en" style="font-size:12px" placeholder="مثال: 2025"></div>
        <div><label class="wizard-label" style="font-size:10px">اسم الاختبار / Name</label><input type="text" id="arch-name" class="wizard-input" style="font-size:12px" placeholder="ابحث بالاسم..."></div>
      </div>
      <div style="display:flex;gap:10px">
        <button onclick="searchArchive()" style="background:linear-gradient(135deg,#FACC15,#f59e0b);color:#1e3a8a;border:none;border-radius:12px;padding:9px 24px;font-weight:800;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🔍 بحث / Search</button>
        <button onclick="clearArchiveSearch()" style="background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:rgba(255,255,255,.7);border-radius:12px;padding:9px 18px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🗑 مسح / Clear</button>
      </div>
    </div>
    <!-- النتائج -->
    <div id="archiveResults" style="min-height:120px">
      <div style="text-align:center;color:rgba(255,255,255,.3);padding:40px;font-family:Tajawal,sans-serif">ابحث للعرض / Search to view results</div>
    </div>
  </div>
  <div id="supReviewerPanel" class="hidden w-full max-w-6xl bg-black/40 backdrop-blur-xl p-8 rounded-3xl border border-white/20 mt-6" style="max-height:75vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px">
      <h3 style="font-size:22px;font-weight:800;color:#FACC15">🔍 مراجع الاختبارات / Reviewer</h3>
      <button onclick="closeReviewerPanel()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:12px;padding:8px 16px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:13px">✕ إغلاق</button>
    </div>
    <div id="supReviewerContent"></div>
  </div>
  <button onclick="backToMain()" style="margin-top:40px;color:rgba(255,255,255,.6);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">خروج / Exit</button>

  <!-- SUP LIVE MONITOR -->
  <div id="supLiveMonitorPanel" class="hidden w-full max-w-6xl bg-black/40 backdrop-blur-xl p-6 rounded-3xl border border-white/20 mt-6" style="max-height:80vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;flex-wrap:wrap;gap:10px">
      <div>
        <h3 style="font-size:22px;font-weight:800;color:#7dd3fc">📡 المراقبة المباشرة</h3>
        <p style="font-size:11px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Live Test Monitoring</p>
      </div>
      <div style="display:flex;gap:10px">
        <button onclick="refreshSupLiveMonitor()" style="background:rgba(56,189,248,.2);border:1px solid rgba(56,189,248,.4);color:#7dd3fc;border-radius:10px;padding:8px 16px;font-family:'Tajawal',sans-serif;font-size:13px;font-weight:700;cursor:pointer">🔄 تحديث</button>
        <button onclick="closeSupPanel('supLiveMonitorPanel')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:13px">✕ إغلاق</button>
      </div>
    </div>
    <!-- Test selector -->
    <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:14px;margin-bottom:16px">
      <label class="wizard-label" style="margin-bottom:6px">اختر الاختبار / Select Test</label>
      <select id="supLiveTestSelect" onchange="loadSupLiveMonitor()" class="wizard-input">
        <option value="">-- اختر اختباراً معتمداً --</option>
      </select>
    </div>
    <!-- Info bar -->
    <div id="supLiveInfoBar" class="hidden" style="background:rgba(56,189,248,.08);border:1px solid rgba(56,189,248,.2);border-radius:12px;padding:12px 16px;margin-bottom:14px;display:flex;gap:20px;flex-wrap:wrap">
      <div><span style="font-size:10px;color:rgba(255,255,255,.4)">الاختبار:</span> <span id="supLiveTestName" style="font-weight:700;color:#7dd3fc"></span></div>
      <div><span style="font-size:10px;color:rgba(255,255,255,.4)">المادة:</span> <span id="supLiveSubject" style="font-weight:700"></span></div>
      <div><span style="font-size:10px;color:rgba(255,255,255,.4)">الصف:</span> <span id="supLiveGrade" style="font-weight:700"></span></div>
      <div><span style="font-size:10px;color:rgba(255,255,255,.4)">الطلاب:</span> <span id="supLiveTotal" style="font-weight:700;color:#FACC15"></span></div>
      <div><span style="font-size:10px;color:rgba(255,255,255,.4)">بدأوا:</span> <span id="supLiveStarted" style="font-weight:700;color:#4ade80"></span></div>
      <div><span style="font-size:10px;color:rgba(255,255,255,.4)">أكملوا:</span> <span id="supLiveCompleted" style="font-weight:700;color:#22c55e"></span></div>
    </div>
    <!-- Stats cards -->
    <div id="supLiveStats" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:10px;margin-bottom:16px"></div>
    <!-- Table -->
    <div style="overflow-x:auto;border-radius:14px;border:1px solid rgba(255,255,255,.1)">
      <table class="admin-table" id="supLiveTable">
        <thead id="supLiveTableHead">
          <tr><th colspan="6" style="color:rgba(255,255,255,.3);font-weight:400">اختر اختباراً لعرض البيانات</th></tr>
        </thead>
        <tbody id="supLiveTableBody"></tbody>
      </table>
    </div>
    <!-- Send to grading button -->
    <div id="supSendGradingWrap" class="hidden" style="margin-top:16px;text-align:center">
      <button onclick="sendToGrading()" style="background:linear-gradient(135deg,#f97316,#ea580c);color:white;border:3px solid #fb923c;border-radius:16px;padding:14px 36px;font-size:15px;font-weight:800;cursor:pointer;font-family:'Tajawal',sans-serif;box-shadow:0 4px 18px rgba(249,115,22,.4)">✍️ إرسال للتصحيح المباشر / Send to Marking</button>
    </div>
  </div>

  <!-- GRADING PANEL (التصحيح المباشر) -->
  <div id="gradingPanel" class="hidden w-full max-w-6xl bg-black/40 backdrop-blur-xl p-6 rounded-3xl border border-white/20 mt-6" style="max-height:90vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px">
      <div>
        <h3 style="font-size:22px;font-weight:800;color:#fb923c">✍️ التصحيح المباشر</h3>
        <p id="gradingTestLabel" style="font-size:12px;color:rgba(255,255,255,.5);font-family:'Montserrat',sans-serif"></p>
      </div>
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <button onclick="saveGrading()" style="background:rgba(74,222,128,.2);border:2px solid rgba(74,222,128,.4);color:#4ade80;border-radius:10px;padding:9px 18px;font-family:'Tajawal',sans-serif;font-size:13px;font-weight:800;cursor:pointer">💾 حفظ</button>
        <button onclick="shareGrading()" style="background:rgba(167,139,250,.2);border:2px solid rgba(167,139,250,.4);color:#c4b5fd;border-radius:10px;padding:9px 18px;font-family:'Tajawal',sans-serif;font-size:13px;font-weight:800;cursor:pointer">👥 مشاركة التصحيح</button>
        <button onclick="approveGrading()" style="background:linear-gradient(135deg,#f97316,#ea580c);color:white;border:2px solid #fb923c;border-radius:10px;padding:9px 18px;font-family:'Tajawal',sans-serif;font-size:13px;font-weight:800;cursor:pointer">✅ اعتماد التصحيح</button>
        <button onclick="closeSupPanel('gradingPanel')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:9px 14px;cursor:pointer;font-size:13px">✕</button>
      </div>
    </div>
    <!-- Status bar -->
    <div id="gradingStatusBar" style="background:rgba(251,146,60,.08);border:1px solid rgba(251,146,60,.2);border-radius:12px;padding:10px 16px;margin-bottom:14px;display:flex;gap:16px;flex-wrap:wrap;font-size:12px">
      <span>الحالة: <span id="gradingStatus" style="font-weight:800;color:#fb923c">مسودة</span></span>
      <span>المصحح: <span id="gradingMarker" style="font-weight:700;color:#FACC15"></span></span>
      <span>آخر حفظ: <span id="gradingLastSave" style="font-family:'Montserrat',sans-serif;color:rgba(255,255,255,.5)">—</span></span>
    </div>
    <!-- Student selector -->
    <div style="display:flex;gap:10px;align-items:center;margin-bottom:14px;flex-wrap:wrap">
      <label class="wizard-label" style="margin:0">الطالب:</label>
      <select id="gradingStudentSelect" onchange="loadStudentGrading()" class="wizard-input" style="flex:1;min-width:200px">
        <option value="">-- اختر طالباً --</option>
      </select>
      <div style="font-size:12px;color:rgba(255,255,255,.4)" id="gradingStudentProgress"></div>
    </div>
    <!-- Student grading content -->
    <div id="gradingContent">
      <div style="text-align:center;padding:60px;color:rgba(255,255,255,.3)">اختر طالباً لبدء التصحيح</div>
    </div>
  </div>

  <!-- GRADING COMMITTEE (لجنة التصحيح) -->
  <div id="gradingCommitteePanel" class="hidden w-full max-w-6xl bg-black/40 backdrop-blur-xl p-6 rounded-3xl border border-white/20 mt-6" style="max-height:90vh;overflow-y:auto">
    <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;flex-wrap:wrap;gap:10px">
      <div>
        <h3 id="gradingCommitteeTitle" style="font-size:22px;font-weight:800;color:#c4b5fd">👥 لجنة التصحيح</h3>
        <p style="font-size:11px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Grading Committee</p>
      </div>
      <button onclick="closeGradingCommittee()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;cursor:pointer;font-size:13px">✕ إغلاق</button>
    </div>
    <!-- Tests list awaiting grading -->
    <div id="committeeContent"></div>
  </div>
</div>

<!-- DOMAIN MODAL -->
<div id="domainModal" class="hidden fixed inset-0 z-50 bg-black/80 backdrop-blur-sm overflow-y-auto">
  <div style="min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:24px 16px">
    <div style="width:100%;max-width:900px;background:linear-gradient(135deg,#0f1f5c,#2d0a5e);border-radius:28px;border:1px solid rgba(255,255,255,.2);padding:32px;text-align:right">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:20px;border-bottom:1px solid rgba(255,255,255,.1);padding-bottom:16px">
        <div style="display:flex;gap:12px;align-items:center;flex:1">
          <div class="domain-badge text-lg" id="modalDomainBadge">1</div>
          <div style="flex:1">
            <input type="text" id="modalDomainNameAr" class="wizard-input font-bold mb-2" placeholder="اسم المجال بالعربي" style="margin-bottom:6px">
            <input type="text" id="modalDomainNameEn" class="wizard-input text-sm font-en" placeholder="Domain Name in English" dir="ltr">
          </div>
        </div>
        <div style="margin-right:16px">
          <div class="logo-upload-area" onclick="document.getElementById('domainIconFile').click()" style="width:80px;height:80px;padding:8px">
            <img id="domainIconPreview" src="" alt="" style="display:none;max-height:60px;object-fit:contain;border-radius:6px">
            <div id="domainIconPlaceholder" style="text-align:center"><div style="font-size:24px">🖼️</div><div style="font-size:10px;color:#94a3b8;margin-top:4px">Icon</div></div>
          </div>
          <input type="file" id="domainIconFile" accept="image/*" style="display:none" onchange="previewDomainIcon(event)">
        </div>
      </div>
      <!-- Branch Question -->
      <div id="branchQuestionBox" style="background:rgba(250,204,21,.1);border:1px solid rgba(250,204,21,.3);border-radius:16px;padding:16px;margin-bottom:20px">
        <p style="font-weight:800;color:#FACC15;margin-bottom:12px;font-size:15px">هل هذا المجال يحتوي على فروع؟<span style="font-family:'Montserrat',sans-serif;font-size:12px;color:rgba(255,255,255,.5);display:block;margin-top:2px">Does this domain have sub-sections?</span></p>
        <div id="branchToggleBox" style="display:flex;gap:12px">
          <button onclick="setBranchMode(false)" id="btnNoBranch" style="flex:1;padding:10px;border-radius:12px;border:2px solid rgba(255,255,255,.2);background:rgba(255,255,255,.05);color:white;font-weight:700;cursor:pointer;font-family:'Tajawal',sans-serif">لا / No</button>
          <button onclick="setBranchMode(true)" id="btnYesBranch" style="flex:1;padding:10px;border-radius:12px;border:2px solid rgba(255,255,255,.2);background:rgba(255,255,255,.05);color:white;font-weight:700;cursor:pointer;font-family:'Tajawal',sans-serif">نعم / Yes — فروع</button>
        </div>
      </div>
      <!-- Branch Setup -->
      <div id="branchSetupBox" class="hidden mb-4">
        <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:16px">
          <label class="wizard-label">عدد الفروع / Number of Branches</label>
          <input type="number" id="branchCount" min="2" max="10" class="wizard-input mb-3" placeholder="مثال: 3" oninput="generateBranchInputs()" style="margin-bottom:12px">
          <div id="branchInputs"></div>
        </div>
      </div>
      <!-- Domain Settings -->
      <div id="domainSettingsBox" class="hidden">
        <div style="background:rgba(255,255,255,.05);border-radius:16px;padding:16px;border:1px solid rgba(255,255,255,.1);margin-bottom:16px">
          <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px">
            <div><label class="wizard-label">⏱ الوقت (دقيقة) / Time</label><input type="number" id="modalDomainTime" class="wizard-input" placeholder="30" min="1"></div>
            <div><label class="wizard-label">📊 وزن المجال / Weight</label><input type="text" id="modalDomainWeight" class="wizard-input" readonly style="opacity:.65"></div>
          </div>
          <!-- عدد الأسئلة هنا مع التوليد التلقائي للأوزان -->
          <div style="margin-top:14px;background:rgba(250,204,21,.08);border:1px solid rgba(250,204,21,.2);border-radius:12px;padding:12px">
            <label class="wizard-label" style="color:#FACC15;margin-bottom:8px;display:block">🔢 عدد الأسئلة / Questions</label>
            <div style="display:flex;align-items:center;gap:10px">
              <input type="number" id="modalQCount" class="wizard-input" placeholder="مثال: 5" min="1" style="flex:1;max-width:120px" oninput="autoGenQuestions()">
              <button onclick="autoGenQuestions()" style="background:linear-gradient(135deg,#FACC15,#f59e0b);color:#1e3a8a;border:none;border-radius:10px;padding:8px 18px;font-weight:800;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">⚡ توليد الأوزان</button>
              <span style="font-size:11px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">وزن كل سؤال = وزن المجال ÷ العدد</span>
            </div>
          </div>
        </div>
        <div id="questionsList" class="space-y-3 mb-4"></div>
        <button onclick="addQuestion()" style="width:100%;border:2px dashed rgba(255,255,255,.2);border-radius:20px;padding:16px;color:rgba(255,255,255,.5);background:transparent;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px" onmouseover="this.style.borderColor='#FACC15';this.style.color='#FACC15'" onmouseout="this.style.borderColor='rgba(255,255,255,.2)';this.style.color='rgba(255,255,255,.5)'">➕ إضافة سؤال / Add Question</button>
      </div>
      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:24px;padding-top:16px;border-top:1px solid rgba(255,255,255,.1)">
        <button onclick="closeDomainModal()" style="color:rgba(255,255,255,.5);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">إلغاء / Cancel</button>
        <div style="display:flex;gap:10px;align-items:center">
          <!-- Save Branch btn: hidden permanently -->
          <button id="saveBranchQBtn" onclick="saveBranchQuestions()" style="display:none!important;visibility:hidden;position:absolute">
          <!-- زر معاينة كلاسيكي -->
          <button id="previewDomainBtn" onclick="previewDomainClassic()" style="background:rgba(59,130,246,.2);border:1.5px solid rgba(59,130,246,.4);color:#93c5fd;padding:10px 22px;border-radius:20px;font-weight:700;font-size:13px;cursor:pointer;font-family:'Tajawal',sans-serif;display:none">👁 معاينة كلاسيكي</button>
          <!-- Save Domain btn -->
          <button id="saveDomainBtn" onclick="saveDomain()" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;padding:12px 40px;border:3px solid #86efac;border-radius:30px;font-weight:800;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif;box-shadow:0 4px 14px rgba(34,197,94,.4)">💾 حفظ المجال / Save</button>
        </div>
      </div>
    </div>
  </div>
</div>

<!-- QUESTION MODAL -->
<div id="questionModal" class="hidden fixed inset-0 z-[60] bg-black/90 backdrop-blur-sm overflow-y-auto">
  <div style="min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:24px 16px">
    <div style="width:100%;max-width:800px;background:linear-gradient(135deg,#0a1540,#1a0535);border-radius:28px;border:1px solid rgba(255,255,255,.2);padding:32px;text-align:right">
      <div style="display:flex;flex-wrap:wrap;gap:12px;margin-bottom:20px;background:rgba(255,255,255,.05);border-radius:18px;padding:16px;border:1px solid rgba(255,255,255,.1)">
        <div style="flex:1;min-width:120px"><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Domain</div><div style="font-size:14px;font-weight:800;color:#FACC15" id="qInfoDomain">—</div></div>
        <div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Q#</div><div style="font-size:14px;font-weight:800;font-family:'Montserrat',sans-serif" id="qInfoNum">Q.1</div></div>
        <div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif">Score</div><div style="font-size:14px;font-weight:800;color:#4ade80;font-family:'Montserrat',sans-serif" id="qInfoScore">—</div></div>
        <div style="flex:1;min-width:200px"><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif;margin-bottom:4px">Question Type / نمط السؤال</div>
          <select id="qType" class="wizard-input text-sm" onchange="renderQuestionBody()">
            <option value="">-- اختر النمط --</option>
            <option value="mcq">اختيار من متعدد / MCQ</option>
            <option value="matching">توصيل / Matching</option>
            <option value="ordering">رتب المفردات / Word Order</option>
            <option value="speaking">تحدث / Speaking</option>
            <option value="oral">قراءة جهرية / Oral Reading</option>
            <option value="listening">استماع / Listening</option>
            <option value="reading">فهم المقروء / Reading</option>
            <option value="truefalse">صواب أم خطأ / True or False</option>
            <option value="classify">صنف المفردات / Classify</option>
            <option value="writingskill">مهارة الكتابة / Writing Skill</option>
          </select>
        </div>
      </div>
      <div class="grid grid-cols-2 gap-4 mb-4">
        <div><label class="wizard-label">تصنيف بلوم / Bloom's</label><select id="qBloom" class="wizard-input text-sm"><option value="">-- اختر --</option><option value="remember">تذكر / Remember</option><option value="understand">فهم / Understand</option><option value="apply">تطبيق / Apply</option><option value="analyze">تحليل / Analyze</option><option value="evaluate">تقييم / Evaluate</option><option value="create">إبداع / Create</option></select></div>
        <div><label class="wizard-label">التصحيح / Marking</label><select id="qMarking" class="wizard-input text-sm"><option value="auto">إلكتروني / Auto ✓</option><option value="manual">يدوي / Manual ✍</option></select></div>
        <div><label class="wizard-label">نواتج التعلم / GLO</label><select id="qLO" class="wizard-input text-sm"><option value="">-- اختر --</option></select><p id="qLOEmpty" style="color:#f87171;font-size:11px;margin-top:4px;display:none">⚠️ لا توجد نواتج — أضفها من بنك المعايير</p></div>
        <div><label class="wizard-label">المعيار / Standard</label><input type="text" id="qStandard" class="wizard-input text-sm" placeholder="CCSS.ELA.RI.3.1"></div>
      </div>
      <!-- WEIGHT FIELD -->
      <div style="background:rgba(250,204,21,.08);border:2px solid rgba(250,204,21,.3);border-radius:16px;padding:14px 18px;margin-bottom:16px">
        <div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px">
          <div style="flex:1;min-width:180px">
            <label style="font-size:11px;color:#FACC15;font-weight:700;display:block;margin-bottom:6px;font-family:'Montserrat',sans-serif">📊 نسبة السؤال % / Question Weight</label>
            <input type="number" id="qScore" class="wizard-input font-en" placeholder="0.00" min="0" step="0.01" oninput="updateQWeightMeter()" style="max-width:120px">
          </div>
          <div style="flex:1;min-width:200px">
            <div style="font-size:11px;color:rgba(255,255,255,.5);margin-bottom:6px;font-family:'Montserrat',sans-serif">مجموع نسب الأسئلة / Total Used</div>
            <div style="display:flex;align-items:center;gap:10px">
              <div style="flex:1;height:10px;border-radius:10px;background:rgba(255,255,255,.1);overflow:hidden">
                <div id="qWeightBar" style="height:10px;border-radius:10px;background:#4ade80;width:0%;transition:.3s"></div>
              </div>
              <span id="qWeightLabel" style="font-family:'Montserrat',sans-serif;font-size:13px;font-weight:800;color:#4ade80;white-space:nowrap">0 / 0%</span>
            </div>
          </div>
        </div>
      </div>
      <div id="questionBodyArea" style="background:rgba(255,255,255,.05);border-radius:18px;padding:20px;border:1px solid rgba(255,255,255,.1);min-height:200px">
        <p style="color:rgba(255,255,255,.3);text-align:center;font-size:14px;padding:32px 0">اختر نمط السؤال أولاً / Select type first</p>
      </div>


      <div style="display:flex;justify-content:space-between;align-items:center;margin-top:24px;padding-top:16px;border-top:1px solid rgba(255,255,255,.1)">
        <button onclick="closeQuestionModal()" style="color:rgba(255,255,255,.5);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">إلغاء / Cancel</button>
        <button onclick="saveQuestion()" style="background:#FACC15;color:#1e3a8a;padding:12px 40px;border:none;border-radius:30px;font-weight:800;font-size:15px;cursor:pointer;font-family:'Tajawal',sans-serif">💾 حفظ السؤال / Save</button>
      </div>
    </div>
  </div>
</div>

<!-- GITHUB LOADER -->
<div id="ghLoader" style="display:none;position:fixed;inset:0;z-index:99998;background:rgba(0,0,0,.85);backdrop-filter:blur(8px);align-items:center;justify-content:center;flex-direction:column;gap:16px">
  <div style="width:56px;height:56px;border:4px solid rgba(250,204,21,.2);border-top-color:#FACC15;border-radius:50%;animation:spin 1s linear infinite"></div>
  <p style="color:#FACC15;font-family:'Montserrat',sans-serif;font-size:14px;font-weight:700">جاري التحميل / Loading...</p>
  <style>@keyframes spin{to{transform:rotate(360deg)}}

/* ══ Media Source Toggle ══ */
.media-toggle{display:inline-flex;align-items:center;gap:5px;border-radius:20px;padding:6px 12px;cursor:pointer;font-size:12px;font-weight:700;font-family:'Montserrat',sans-serif;border:2px solid;transition:all .25s;user-select:none}
.media-toggle.on{background:rgba(34,197,94,.15);border-color:#22c55e;color:#4ade80}
.media-toggle.off{background:rgba(59,130,246,.12);border-color:#3b82f6;color:#93c5fd}
/* ── DISPLAY MODE 4: White Paper ── */
#studentWindow.dm-mode-4{background:#c8c8c0;}
#studentWindow.dm-mode-4 .sw-content{padding:20px 12px;display:flex;flex-direction:column;align-items:center;}
#studentWindow.dm-mode-4 .sw-footer{display:none!important;}
.wp-page{background:white;width:100%;max-width:760px;padding:28px 32px 24px;box-shadow:0 4px 24px rgba(0,0,0,.18),0 1px 3px rgba(0,0,0,.1);margin-bottom:0;font-family:'Tajawal','Times New Roman',serif;color:#111;position:relative;}
.wp-header-bar{display:flex;align-items:center;justify-content:space-between;border-bottom:3px double #111;padding-bottom:10px;margin-bottom:16px;gap:12px;}
.wp-test-title{font-size:17px;font-weight:900;text-align:center;flex:1;font-family:'Tajawal',serif;}
.wp-school-logo{max-height:56px;max-width:70px;object-fit:contain;}
.wp-meta-grid{font-size:11px;color:#333;display:grid;grid-template-columns:1fr 1fr;gap:3px 16px;margin-bottom:16px;padding:8px;border:1px solid #ccc;border-radius:2px;}
.wp-domain-bar{font-size:14px;font-weight:900;background:#1e3a8a;color:white;padding:6px 16px;margin:16px 0 12px;display:inline-block;border-radius:2px;}
.wp-q{margin-bottom:22px;}
.wp-q-stem{font-size:14px;font-weight:700;line-height:1.8;display:flex;gap:8px;margin-bottom:8px;}
.wp-q-num{font-weight:900;min-width:26px;color:#1e3a8a;flex-shrink:0;}
.wp-indent{margin-right:26px;}
.wp-answer-line{border-bottom:1px solid #ccc;height:28px;margin-bottom:5px;}
.wp-mcq-grid{display:grid;grid-template-columns:1fr 1fr;gap:5px;}
.wp-mcq-opt{display:flex;align-items:center;gap:6px;font-size:13px;cursor:pointer;padding:5px 8px;border:1px solid #e2e8f0;border-radius:4px;transition:.15s;line-height:1.5;}
.wp-mcq-opt.sel{background:#dbeafe;border-color:#3b82f6;font-weight:700;}
.wp-opt-circle{width:16px;height:16px;border-radius:50%;border:2px solid #94a3b8;flex-shrink:0;display:flex;align-items:center;justify-content:center;font-size:9px;}
.wp-mcq-opt.sel .wp-opt-circle{border-color:#3b82f6;background:#3b82f6;color:white;}
.wp-tf-row{display:flex;align-items:center;gap:8px;margin-bottom:7px;font-size:13px;}
.wp-tf-btn{border:1.5px solid #94a3b8;border-radius:4px;padding:3px 10px;cursor:pointer;font-weight:700;font-size:12px;transition:.15s;}
.wp-tf-t.sel{background:#dcfce7;border-color:#22c55e;color:#15803d;}
.wp-tf-f.sel{background:#fee2e2;border-color:#ef4444;color:#b91c1c;}
.wp-sep{width:100%;max-width:760px;height:18px;background:repeating-linear-gradient(90deg,#c8c8c0 0,#c8c8c0 8px,#aaa 8px,#aaa 10px);}
.wp-page-footer{border-top:1px solid #bbb;margin-top:16px;padding-top:6px;display:flex;justify-content:space-between;font-size:10px;color:#777;}
.wp-q-sep{height:2px;background:linear-gradient(90deg,transparent,#FACC15 20%,#F59E0B 50%,#FACC15 80%,transparent);margin:14px 0;box-shadow:0 1px 4px rgba(245,158,11,.3);}
/* ══ Progress Dots — أبيض→ذهبي→أخضر ══ */
.sw-dot{width:26px;height:26px;border-radius:50%;background:rgba(255,255,255,.15);border:2px solid rgba(255,255,255,.3);display:flex;align-items:center;justify-content:center;font-size:10px;font-weight:700;color:rgba(255,255,255,.6);transition:all .3s;cursor:pointer;font-family:'Montserrat',sans-serif}
.sw-dot.answered{background:linear-gradient(135deg,#22c55e,#15803d);border-color:#22c55e;color:white;box-shadow:0 0 8px rgba(34,197,94,.5);transform:scale(1.08)}
.sw-dot.current{background:linear-gradient(135deg,#FACC15,#F59E0B);border-color:#FACC15;color:#1a0832;font-weight:900;box-shadow:0 0 10px rgba(250,204,21,.55);transform:scale(1.15)}

/* ══ Ordering drag indicator ══ */
.sw-word-tile.drag-over-indicator{border-left:3px solid #FACC15 !important;transform:scale(1.05) !important;}
#orderDropZone.drag-active{border-color:#FACC15 !important;background:#fffbeb !important;box-shadow:inset 0 0 0 2px rgba(250,204,21,.3) !important;}

/* ══ Student Window Frame ══ */
#studentWindow{
  border:6px solid #1e3a8a;
  box-sizing:border-box;
}
.sw-frame{
  border-radius:18px;
  overflow:hidden;
  margin-bottom:16px;
  width:100%;
  border:3px solid #FACC15;
  box-shadow:0 0 0 6px #14532d55,0 8px 32px rgba(30,58,138,.25);
}
.sw-q-header{
  display:flex;align-items:flex-start;gap:14px;
  padding:16px 20px;
  background:#00bfff;
  border-bottom:3px solid #FACC15;
}
.sw-q-stem{font-size:17px;font-weight:700;line-height:1.7;color:#00172e;flex:1}
.sw-q-badge{
  background:#00172e;border:2px solid rgba(255,255,255,.4);
  color:#00bfff;font-weight:900;border-radius:12px;
  min-width:42px;height:42px;display:flex;align-items:center;justify-content:center;
  font-size:14px;font-family:'Montserrat',sans-serif;flex-shrink:0;
}
.sw-q-score{
  background:#00172e;color:#FACC15;border:1px solid #FACC15;
  border-radius:8px;padding:3px 10px;font-size:12px;font-weight:700;
  font-family:'Montserrat',sans-serif;white-space:nowrap;
}
.sw-q-body-wrap{
  background:white;
  padding:18px 20px;
  border-bottom:3px solid #14532d;
}
.sw-q-ans-wrap{
  background:#f0fdf4;
  padding:14px 20px 18px;
  border-top:3px solid #14532d;
}
.sw-mcq-opt{
  border:2px solid #d1fae5;border-radius:14px;padding:12px 16px;
  cursor:pointer;transition:all .2s;display:flex;align-items:center;gap:12px;
  background:#f0fdf4;margin-bottom:8px;
}
.sw-mcq-opt:hover{border-color:#059669;background:#dcfce7;}
.sw-mcq-opt.selected{border-color:#059669;background:#dcfce7;}
.sw-opt-label{
  width:30px;height:30px;border-radius:50%;border:2px solid #6ee7b7;
  display:flex;align-items:center;justify-content:center;
  font-weight:700;font-size:13px;color:#065f46;flex-shrink:0;
  font-family:'Montserrat',sans-serif;background:white;
}
.sw-mcq-opt.selected .sw-opt-label{background:#059669;border-color:#059669;color:white;}
</style>
</div>

<!-- SCHOOL COORDINATOR PANEL -->
<div id="schoolCoordPanel" class="hidden w-full min-h-screen flex" dir="rtl">
  <!-- Sidebar -->
  <div style="width:260px;min-height:100vh;background:rgba(0,0,0,.5);backdrop-filter:blur(20px);border-left:1px solid rgba(255,255,255,.1);padding:24px 0;display:flex;flex-direction:column;flex-shrink:0">
    <div style="padding:0 20px 24px;border-bottom:1px solid rgba(255,255,255,.1)">
      <div style="font-size:11px;font-family:'Montserrat',sans-serif;color:rgba(255,255,255,.4);text-transform:uppercase;letter-spacing:2px;margin-bottom:4px">School Coordinator</div>
      <div id="coordSchoolName" style="font-size:17px;font-weight:800;color:#FACC15">منسق المدرسة</div>
    </div>
    <nav style="flex:1;padding:16px 12px;display:flex;flex-direction:column;gap:4px">
      <button onclick="coordShowSection('coordCodes')" id="navBtn_coordCodes" class="coord-nav-btn active-nav">
        <span>🎫</span> أكواد الطلاب <span style="font-family:'Montserrat',sans-serif;font-size:9px;opacity:.5;display:block">Student Codes</span>
      </button>
      <button onclick="coordShowSection('coordLive')" id="navBtn_coordLive" class="coord-nav-btn">
        <span>📡</span> المراقبة المباشرة <span style="font-family:'Montserrat',sans-serif;font-size:9px;opacity:.5;display:block">Live Monitoring</span>
      </button>
      <div style="margin:8px 0;border-top:1px solid rgba(255,255,255,.08)"></div>
      <div style="font-size:10px;color:rgba(255,255,255,.3);padding:4px 12px;font-family:'Montserrat',sans-serif;text-transform:uppercase;letter-spacing:1px">التقارير / Reports</div>
      <button onclick="coordShowSection('coordRep1')" id="navBtn_coordRep1" class="coord-nav-btn">
        <span>📊</span> تقارير التحصيل
      </button>
      <button onclick="coordShowSection('coordRep2')" id="navBtn_coordRep2" class="coord-nav-btn">
        <span>📈</span> تقارير التطور
      </button>
      <button onclick="coordShowSection('coordRep3')" id="navBtn_coordRep3" class="coord-nav-btn">
        <span>🧠</span> مستويات بلوم
      </button>
      <button onclick="coordShowSection('coordRep4')" id="navBtn_coordRep4" class="coord-nav-btn">
        <span>🎯</span> مدى تحقيق المعايير
      </button>
      <button onclick="coordShowSection('coordRep5')" id="navBtn_coordRep5" class="coord-nav-btn">
        <span>🏙</span> المقارنات المحلية
      </button>
      <button onclick="coordShowSection('coordRep6')" id="navBtn_coordRep6" class="coord-nav-btn">
        <span>🌍</span> المقارنات العالمية
      </button>
      <button onclick="coordShowSection('coordRep7')" id="navBtn_coordRep7" class="coord-nav-btn">
        <span>🇪🇺</span> المرجع الأوربي للغات
      </button>
    </nav>
    <div style="padding:16px 20px;border-top:1px solid rgba(255,255,255,.1)">
      <button onclick="logout()" style="width:100%;background:rgba(248,113,113,.15);border:1px solid rgba(248,113,113,.3);color:#f87171;border-radius:12px;padding:10px;font-family:'Tajawal',sans-serif;font-size:14px;font-weight:700;cursor:pointer">تسجيل خروج / Logout</button>
    </div>
  </div>

  <!-- Main Content -->
  <div style="flex:1;overflow-y:auto;padding:32px 28px">

    <!-- CODES SECTION -->
    <div id="coordCodes">
      <h2 style="font-size:24px;font-weight:900;margin-bottom:6px">🎫 أكواد الطلاب <span style="font-family:'Montserrat',sans-serif;font-size:13px;color:rgba(255,255,255,.4);font-weight:400">Student Codes</span></h2>
      <!-- Filters -->
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:16px">
        <select id="coordFilterYear" onchange="renderCoordCodes()" style="flex:1;min-width:150px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px">
          <option value="">كل السنوات الأكاديمية</option>
        </select>
        <select id="coordFilterGrade" onchange="renderCoordCodes()" style="flex:1;min-width:150px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px">
          <option value="">كل الصفوف</option>
        </select>
        <select id="coordFilterTest" onchange="renderCoordCodes()" style="flex:1;min-width:180px;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px">
          <option value="">كل الاختبارات</option>
        </select>
        <button onclick="exportCoordCodesToExcel()" style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);color:#4ade80;border-radius:10px;padding:9px 18px;font-family:'Tajawal',sans-serif;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap">⬇️ تصدير Excel</button>
      </div>
      <div style="overflow-x:auto;border-radius:16px;border:1px solid rgba(255,255,255,.1)">
        <table class="admin-table" id="coordCodesTable">
          <thead><tr>
            <th>#</th>
            <th>السنة الأكاديمية</th>
            <th>الصف</th>
            <th>اسم الاختبار</th>
            <th>الاسم الأول</th>
            <th>الاسم الثاني</th>
            <th style="font-family:'Montserrat',sans-serif">Username</th>
            <th style="font-family:'Montserrat',sans-serif">Password</th>
          </tr></thead>
          <tbody id="coordCodesBody"><tr><td colspan="8" style="color:rgba(255,255,255,.3);padding:32px;text-align:center">جاري التحميل...</td></tr></tbody>
        </table>
      </div>
    </div>

    <!-- LIVE MONITORING -->
    <div id="coordLive" class="hidden">
      <h2 style="font-size:24px;font-weight:900;margin-bottom:16px">📡 المراقبة المباشرة <span style="font-family:'Montserrat',sans-serif;font-size:13px;color:rgba(255,255,255,.4);font-weight:400">Live Monitoring</span></h2>
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:16px;align-items:flex-end">
        <div style="flex:1;min-width:140px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">السنة الأكاديمية</div>
          <select id="liveFilterYear" onchange="filterLiveTests()" style="width:100%;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px"><option value="">الكل</option></select>
        </div>
        <div style="flex:1;min-width:120px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">المادة</div>
          <select id="liveFilterSubject" onchange="filterLiveTests()" style="width:100%;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px"><option value="">الكل</option></select>
        </div>
        <div style="flex:2;min-width:180px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">اسم الاختبار</div>
          <select id="liveFilterTest" onchange="loadLiveMonitor()" style="width:100%;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px"><option value="">-- اختر اختباراً --</option></select>
        </div>
        <button onclick="refreshLiveMonitor()" style="background:rgba(96,165,250,.2);border:1px solid rgba(96,165,250,.4);color:#93c5fd;border-radius:10px;padding:9px 16px;font-family:'Tajawal',sans-serif;font-size:13px;font-weight:700;cursor:pointer;white-space:nowrap">🔄 تحديث</button>
      </div>
      <!-- Test Info Bar -->
      <div id="liveTestInfo" class="hidden" style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:14px 20px;margin-bottom:16px;display:flex;gap:24px;flex-wrap:wrap">
        <div><span style="font-size:11px;color:rgba(255,255,255,.4)">الحالة:</span> <span id="liveTestStatus" style="font-weight:700"></span></div>
        <div><span style="font-size:11px;color:rgba(255,255,255,.4)">من:</span> <span id="liveTestFrom" style="font-family:'Montserrat',sans-serif;font-size:13px;font-weight:700"></span></div>
        <div><span style="font-size:11px;color:rgba(255,255,255,.4)">إلى:</span> <span id="liveTestTo" style="font-family:'Montserrat',sans-serif;font-size:13px;font-weight:700"></span></div>
        <div><span style="font-size:11px;color:rgba(255,255,255,.4)">المدة:</span> <span id="liveTestDuration" style="font-family:'Montserrat',sans-serif;font-size:13px;font-weight:700"></span></div>
      </div>
      <!-- Live Table -->
      <div style="overflow-x:auto;border-radius:16px;border:1px solid rgba(255,255,255,.1)">
        <table class="admin-table" id="liveTable">
          <thead id="liveTableHead"><tr><th colspan="4" style="color:rgba(255,255,255,.3);font-weight:400">اختر اختباراً لعرض بيانات المراقبة</th></tr></thead>
          <tbody id="liveTableBody"></tbody>
        </table>
      </div>
    </div>

    <!-- PLACEHOLDER REPORTS -->
    <div id="coordRep1" class="hidden">
      <h2 style="font-size:24px;font-weight:900;margin-bottom:16px">📊 تقارير التحصيل <span style="font-family:'Montserrat',sans-serif;font-size:13px;color:rgba(255,255,255,.4);font-weight:400">Attainment Reports</span></h2>
      <div style="display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:16px;align-items:flex-end">
        <div style="flex:1;min-width:150px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">السنة الأكاديمية / Academic Year</div>
          <select id="attYearFilter" style="width:100%;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px"><option value="">الكل / All</option></select>
        </div>
        <div style="flex:1;min-width:150px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">المادة / Subject</div>
          <select id="attSubjectFilter" style="width:100%;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px"><option value="">الكل / All</option></select>
        </div>
        <div style="flex:1;min-width:150px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">الفصل الدراسي / Term</div>
          <select id="attTermFilter" style="width:100%;background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);color:white;border-radius:10px;padding:9px 12px;font-family:'Tajawal',sans-serif;font-size:13px"><option value="">الكل / All</option><option value="1">الفصل الأول / Term 1</option><option value="2">الفصل الثاني / Term 2</option><option value="3">الفصل الثالث / Term 3</option></select>
        </div>
        <div style="flex:1;min-width:150px">
          <div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">كود المدرسة / School Code</div>
          <input type="text" id="attSchoolCodeFilter" readonly style="width:100%;background:rgba(250,204,21,.1);border:1.5px solid rgba(250,204,21,.4);color:#FACC15;font-weight:800;border-radius:10px;padding:9px 12px;font-family:'Montserrat',sans-serif;font-size:13px;cursor:not-allowed;box-sizing:border-box">
        </div>
        <button onclick="runAttainmentReport()" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:none;border-radius:12px;padding:11px 28px;font-weight:800;font-size:14px;cursor:pointer;font-family:'Tajawal',sans-serif;white-space:nowrap;box-shadow:0 4px 14px rgba(34,197,94,.35)">▶ تشغيل / Run</button>
      </div>
      <div id="attainmentReportOutput"></div>
    </div>
    <div id="coordRep2" class="hidden"><div style="text-align:center;padding:80px 20px"><div style="font-size:64px;margin-bottom:16px">📈</div><h3 style="font-size:20px;font-weight:800;color:#FACC15;margin-bottom:8px">تقارير التطور</h3><p style="color:rgba(255,255,255,.4)">قيد التطوير / Under Development</p></div></div>
    <div id="coordRep3" class="hidden"><div style="text-align:center;padding:80px 20px"><div style="font-size:64px;margin-bottom:16px">🧠</div><h3 style="font-size:20px;font-weight:800;color:#FACC15;margin-bottom:8px">تقارير مستويات بلوم</h3><p style="color:rgba(255,255,255,.4)">قيد التطوير / Under Development</p></div></div>
    <div id="coordRep4" class="hidden"><div style="text-align:center;padding:80px 20px"><div style="font-size:64px;margin-bottom:16px">🎯</div><h3 style="font-size:20px;font-weight:800;color:#FACC15;margin-bottom:8px">تقارير مدى تحقيق المعايير</h3><p style="color:rgba(255,255,255,.4)">قيد التطوير / Under Development</p></div></div>
    <div id="coordRep5" class="hidden"><div style="text-align:center;padding:80px 20px"><div style="font-size:64px;margin-bottom:16px">🏙</div><h3 style="font-size:20px;font-weight:800;color:#FACC15;margin-bottom:8px">تقارير المقارنات المحلية</h3><p style="color:rgba(255,255,255,.4)">قيد التطوير / Under Development</p></div></div>
    <div id="coordRep6" class="hidden"><div style="text-align:center;padding:80px 20px"><div style="font-size:64px;margin-bottom:16px">🌍</div><h3 style="font-size:20px;font-weight:800;color:#FACC15;margin-bottom:8px">تقارير المقارنات العالمية</h3><p style="color:rgba(255,255,255,.4)">قيد التطوير / Under Development</p></div></div>
    <div id="coordRep7" class="hidden"><div style="text-align:center;padding:80px 20px"><div style="font-size:64px;margin-bottom:16px">🇪🇺</div><h3 style="font-size:20px;font-weight:800;color:#FACC15;margin-bottom:8px">تقارير القياس مع المرجع الأوربي للغات</h3><p style="color:rgba(255,255,255,.4)">قيد التطوير / Under Development</p></div></div>
  </div>
</div>

<!-- STUDENT WINDOW -->
<div id="studentWindow">
  <div class="sw-header" style="background:linear-gradient(135deg,#1e1245 0%,#2a1860 45%,#1e1245 100%);display:flex;align-items:stretch;min-height:64px;border-bottom:1px solid rgba(250,204,21,.12);position:relative;overflow:hidden">
    <div style="position:absolute;inset:0;background:radial-gradient(ellipse 60% 100% at 50% 50%,rgba(124,58,237,.12) 0%,transparent 70%);pointer-events:none"></div>

    <!-- ١. يمين: لوجو Scholastic -->
    <div style="display:flex;align-items:center;gap:8px;padding:0 14px;border-right:1px solid rgba(255,255,255,.07);flex-shrink:0;position:relative;z-index:1">
      <div style="width:40px;height:40px;background:linear-gradient(135deg,#f59e0b,#d97706);border-radius:9px;display:flex;align-items:center;justify-content:center;flex-shrink:0;box-shadow:0 2px 10px rgba(245,158,11,.45)">
        <svg width="26" height="26" viewBox="0 0 26 26" fill="none">
          <polygon points="13,3 4,8 4,9.5 22,9.5 22,8" fill="white" opacity=".95"/>
          <rect x="9.5" y="9.5" width="7" height="1" fill="white" opacity=".8"/>
          <rect x="7" y="10.5" width="12" height="6.5" rx="1.5" fill="white" opacity=".9"/>
          <rect x="8.5" y="11.5" width="9" height="4.5" rx="1" fill="#d97706"/>
          <line x1="11" y1="11.5" x2="11" y2="16" stroke="white" stroke-width=".7" opacity=".6"/>
          <line x1="13" y1="11.5" x2="13" y2="16" stroke="white" stroke-width=".7" opacity=".6"/>
          <line x1="15" y1="11.5" x2="15" y2="16" stroke="white" stroke-width=".7" opacity=".6"/>
          <line x1="22" y1="8" x2="22" y2="18" stroke="white" stroke-width="1.8" stroke-linecap="round"/>
          <ellipse cx="22" cy="19.5" rx="2.5" ry="1.5" fill="white"/>
        </svg>
      </div>
      <div>
        <div style="font-size:14px;font-weight:900;color:white;font-family:'Montserrat',sans-serif;line-height:1.15;letter-spacing:.3px">Scholastic</div>
        <div style="font-size:9px;color:rgba(255,255,255,.5);font-family:'Montserrat',sans-serif;letter-spacing:.2px;white-space:nowrap">Testing Platform</div>
      </div>
    </div>

    <!-- ٢. اسم الاختبار والمادة -->
    <div style="display:flex;flex-direction:column;justify-content:center;padding:0 16px;border-right:1px solid rgba(255,255,255,.07);flex:1.2;min-width:0;text-align:center;position:relative;z-index:1">
      <div id="sw-test-name-header" style="font-size:13px;font-weight:800;color:white;font-family:'Montserrat',sans-serif;letter-spacing:.2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">Test Name</div>
      <div id="sw-test-meta-header" style="font-size:11px;font-weight:700;color:#FACC15;font-family:'Montserrat',sans-serif;margin-top:3px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis"></div>
    </div>

    <!-- ٣. وسط: عداد -->
    <div style="display:flex;align-items:center;justify-content:center;padding:0 18px;border-right:1px solid rgba(255,255,255,.07);flex-shrink:0;position:relative;z-index:1">
      <div style="background:linear-gradient(135deg,#1e1c3a,#252344);border:1.5px solid rgba(255,255,255,.12);border-radius:10px;padding:5px 14px;box-shadow:inset 0 2px 8px rgba(0,0,0,.6),0 2px 12px rgba(0,0,0,.4);text-align:center;min-width:95px">
        <div class="sw-timer" id="sw-timer-display" style="font-size:30px;font-weight:900;color:white;font-family:'Montserrat',sans-serif;letter-spacing:3px;line-height:1;text-shadow:0 0 12px rgba(255,255,255,.15)">30:00</div>
        <div style="font-size:7px;color:rgba(255,255,255,.3);font-family:'Montserrat',sans-serif;letter-spacing:.5px;margin-top:2px;white-space:nowrap">الوقت المتبقي / Time Left</div>
      </div>
    </div>

    <!-- ٤. المجال والفرع -->
    <div style="display:flex;flex-direction:column;justify-content:center;padding:0 16px;border-right:1px solid rgba(255,255,255,.07);flex:1;min-width:0;position:relative;z-index:1">
      <div id="sw-domain-label-header" style="font-size:13px;font-weight:900;color:#FACC15;font-family:'Montserrat',sans-serif;letter-spacing:.3px;line-height:1.4;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">Domain : Reading</div>
      <div id="sw-branch-label-header" style="font-size:12px;font-weight:700;color:#c4b5fd;font-family:'Montserrat',sans-serif;margin-top:1px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis"></div>
    </div>

    <!-- ٥. يسار: دوائر الأسئلة -->
    <div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding:6px 12px;border-right:1px solid rgba(255,255,255,.07);flex-shrink:0;min-width:100px;position:relative;z-index:1">
      <div id="sw-progress-dots" class="sw-progress-dots" style="display:flex;flex-wrap:wrap;gap:4px;justify-content:center;max-width:170px"></div>
      <div id="sw-q-counter-top" style="font-size:9px;color:rgba(255,255,255,.45);font-family:'Montserrat',sans-serif;margin-top:3px;letter-spacing:.3px;white-space:nowrap">Domain 1 Q1/5</div>
    </div>

    <!-- زر الخروج -->
    <div style="display:flex;align-items:center;padding:0 10px;flex-shrink:0;position:relative;z-index:1">
      <div id="sw-watch-eye" style="display:none"></div>
      <button onclick="swRequestClose()" id="sw-close-btn" style="background:rgba(239,68,68,.18);border:1px solid rgba(239,68,68,.35);color:#fca5a5;border-radius:6px;padding:4px 9px;cursor:pointer;font-family:'Montserrat',sans-serif;font-size:10px;font-weight:700;white-space:nowrap;letter-spacing:.2px">✕ EXIT</button>
    </div>
  </div>
  <!-- hidden elements for JS compatibility -->
  <span id="sw-sticky-school" style="display:none"></span>
  <span id="sw-sticky-logo" style="display:none"></span>
  <span id="sw-sticky-student" style="display:none">Student</span>
  <span id="sw-sticky-info" style="display:none">Subject · Grade</span>
  <div id="sw-domain-name-bar" style="display:none">Domain</div>
  <div class="sw-content" id="sw-questions-content"></div>
  <div class="sw-footer" id="sw-footer">
    <button class="sw-nav-btn prev" id="sw-prev-btn" onclick="swNavigate(-1)" disabled>
      <span style="display:flex;align-items:center;gap:5px;font-size:14px;font-weight:800">← <span>السابق</span></span>
      <span class="btn-sub">Previous</span>
    </button>
    <div id="sw-q-counter" style="display:none"></div>
    <button class="sw-nav-btn submit" id="sw-next-btn" onclick="swNavigate(1)">
      <span style="display:flex;align-items:center;gap:5px;font-size:14px;font-weight:800"><span>التالي</span> →</span>
      <span class="btn-sub">Next</span>
    </button>
  </div>
</div>

<!-- MY TESTS MODAL -->
<div id="myTestsModal" class="hidden fixed inset-0 z-50 bg-black/80 backdrop-blur-sm overflow-y-auto">
  <div style="min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:24px 16px">
    <div style="width:100%;max-width:1000px;background:linear-gradient(135deg,#0f1f5c,#2d0a5e);border-radius:28px;border:1px solid rgba(255,255,255,.2);padding:32px;text-align:right">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;border-bottom:1px solid rgba(255,255,255,.1);padding-bottom:16px">
        <div><h2 style="font-size:28px;font-weight:800;color:#4ade80">📂 اختباراتي / My Tests</h2><p style="color:rgba(255,255,255,.5);font-size:12px;font-family:'Montserrat',sans-serif;margin-top:4px">Review & Approval Workflow</p></div>
        <button onclick="closeMyTests()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:12px;width:40px;height:40px;cursor:pointer;font-size:18px">✕</button>
      </div>
      <div style="display:flex;gap:10px;margin-bottom:20px;flex-wrap:wrap">
        <button id="tab_underReview" onclick="renderMyTestsTabs('underReview')" class="tab-btn active">🔄 تحت المراجعة / Under Review</button>
        <button id="tab_reviewDone" onclick="renderMyTestsTabs('reviewDone')" class="tab-btn">✅ تمت المراجعة / Done</button>
        <button id="tab_approved" onclick="renderMyTestsTabs('approved')" class="tab-btn">🏆 معتمد / Approved</button>
      </div>
      <div id="myTestsContent"></div>
      <!-- Student Codes Panel -->
      <div id="studentCodesPanel" class="hidden mt-6" style="background:rgba(0,0,0,.3);border:1px solid rgba(74,222,128,.3);border-radius:20px;padding:24px">
        <input type="hidden" id="codesTestId">
        <h4 style="color:#4ade80;font-weight:800;font-size:18px;margin-bottom:16px">👥 أكواد الطلاب / Student Codes</h4>
        <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:16px;margin-bottom:16px">
          <h5 style="color:#FACC15;font-weight:800;font-size:14px;margin-bottom:12px">✍️ إضافة يدوي / Manual Add</h5>
          <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-3">
            <div><label class="wizard-label" style="font-size:9px">School ID</label><input type="text" id="manSchoolID" class="wizard-input text-sm" placeholder="STU001"></div>
            <div><label class="wizard-label" style="font-size:9px">First Name / الاسم الأول</label><input type="text" id="manFirstName" class="wizard-input text-sm"></div>
            <div><label class="wizard-label" style="font-size:9px">Last Name / الاسم الثاني</label><input type="text" id="manLastName" class="wizard-input text-sm"></div>
            <div><label class="wizard-label" style="font-size:9px">Gender / الجنس</label><select id="manGender" class="wizard-input text-sm"><option value="Male">Male / ذكر</option><option value="Female">Female / أنثى</option></select></div>
            <div><label class="wizard-label" style="font-size:9px">School Name</label><input type="text" id="manSchoolName" class="wizard-input text-sm"></div>
            <div><label class="wizard-label" style="font-size:9px">Nationality</label><input type="text" id="manNationality" class="wizard-input text-sm"></div>
            <div><label class="wizard-label" style="font-size:9px">Local / مواطن</label><select id="manNational" class="wizard-input text-sm"><option value="No">No</option><option value="Yes">Yes</option></select></div>
            <div><label class="wizard-label" style="font-size:9px">Gifted / موهوب</label><select id="manGifted" class="wizard-input text-sm"><option value="No">No</option><option value="Yes">Yes</option></select></div>
            <div><label class="wizard-label" style="font-size:9px">SOD</label><select id="manSOD" class="wizard-input text-sm"><option value="No">No</option><option value="Yes">Yes</option></select></div>
            <div style="display:flex;align-items:flex-end"><button onclick="addStudentManually()" style="width:100%;background:#22c55e;color:white;font-weight:800;border:none;border-radius:12px;padding:10px;font-size:13px;cursor:pointer;font-family:'Tajawal',sans-serif">➕ إضافة</button></div>
          </div>
        </div>
        <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:16px;margin-bottom:16px">
          <h5 style="color:#FACC15;font-weight:800;font-size:14px;margin-bottom:8px">📊 رفع Excel / Upload</h5>
          <p style="font-size:11px;color:rgba(255,255,255,.4);font-family:'Montserrat',sans-serif;margin-bottom:10px">Columns: Student School ID | First Name | Last Name | Gender | School Name | Grade | Subject | Term | Local(Y/N) | Gifted(Y/N) | SOD(Y/N)<br>الأعمدة: الكود | الاسم الأول | الاسم الثاني | الجنس | المدرسة | الصف | المادة | الفصل | مواطن | موهوب | SOD</p>
          <div style="display:flex;gap:10px;flex-wrap:wrap">
            <label style="display:flex;align-items:center;gap:8px;background:rgba(34,197,94,.2);border:1px solid rgba(34,197,94,.4);border-radius:12px;padding:10px 16px;cursor:pointer;font-size:13px;font-family:'Tajawal',sans-serif">📊 رفع Excel<input type="file" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadStudentsExcelCodes(event)"></label>
            <a href="#" onclick="downloadStudentTemplate(event)" style="display:flex;align-items:center;gap:8px;background:rgba(96,165,250,.2);border:1px solid rgba(96,165,250,.4);border-radius:12px;padding:10px 16px;cursor:pointer;font-size:13px;color:#93c5fd;text-decoration:none;font-family:'Tajawal',sans-serif">⬇️ تحميل النموذج / Template</a>
          </div>
          <div id="uploadCodesResult" style="font-size:12px;color:#4ade80;font-family:'Montserrat',sans-serif;margin-top:8px"></div>
        </div>
        <div id="studentsCodesTableWrap"></div>
        <div id="exportStudentBtns" style="display:none;gap:10px;margin-top:16px;flex-wrap:wrap">
          <button onclick="exportStudentCodesExcel()" style="background:#16a34a;color:white;font-weight:800;border:none;border-radius:12px;padding:10px 20px;font-size:13px;cursor:pointer;font-family:'Tajawal',sans-serif">📊 تصدير Excel / Export</button>
          <button onclick="showStudentIDCards()" style="background:linear-gradient(135deg,#0ea5e9,#0284c7);color:white;font-weight:800;border:none;border-radius:12px;padding:10px 20px;font-size:13px;cursor:pointer;font-family:'Tajawal',sans-serif">🎫 عرض كبطاقات / View as Cards</button>
        </div>
        <div style="background:rgba(255,255,255,.05);border:1px solid rgba(250,204,21,.3);border-radius:14px;padding:16px;margin-top:16px">
          <h5 style="color:#FACC15;font-weight:800;font-size:14px;margin-bottom:12px">📅 نافذة الاختبار / Test Window</h5>
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            <div><label class="wizard-label">بداية الاختبار / Start</label><input type="datetime-local" id="testWindowFrom" class="wizard-input font-en"></div>
            <div><label class="wizard-label">نهاية الاختبار / End</label><input type="datetime-local" id="testWindowTo" class="wizard-input font-en"></div>
          </div>
          <div style="display:flex;gap:10px;flex-wrap:wrap">
            <button onclick="activateTestWindow()" style="background:#FACC15;color:#1e3a8a;font-weight:800;border:none;border-radius:12px;padding:10px 20px;font-size:14px;cursor:pointer;font-family:'Tajawal',sans-serif">✅ تفعيل / Activate</button>
            <button onclick="deactivateTest()" style="background:rgba(239,68,68,.2);color:#f87171;font-weight:800;border:1px solid rgba(239,68,68,.4);border-radius:12px;padding:10px 20px;font-size:14px;cursor:pointer;font-family:'Tajawal',sans-serif">⛔ تعطيل / Deactivate</button>
          </div>
        </div>
        <div style="display:flex;justify-content:flex-end;margin-top:16px"><button onclick="document.getElementById('studentCodesPanel').classList.add('hidden')" style="color:rgba(255,255,255,.5);text-decoration:underline;background:none;border:none;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">إغلاق / Close</button></div>
      </div>
    </div>
  </div>
</div>

<!-- STANDARDS BANK -->
<div id="standardsBankModal" class="hidden fixed inset-0 z-50 bg-black/80 backdrop-blur-sm overflow-y-auto">
  <div style="min-height:100vh;display:flex;align-items:flex-start;justify-content:center;padding:24px 16px">
    <div style="width:100%;max-width:1100px;background:linear-gradient(135deg,#0f1f5c,#2d0a5e);border-radius:28px;border:1px solid rgba(255,255,255,.2);padding:32px;text-align:right">
      <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:24px;border-bottom:1px solid rgba(255,255,255,.1);padding-bottom:16px">
        <div><h2 style="font-size:28px;font-weight:800;color:#FACC15">🎯 بنك المعايير / Standards Bank</h2><div id="sbBreadcrumb" style="display:flex;align-items:center;gap:8px;margin-top:10px;font-size:13px;flex-wrap:wrap"></div></div>
        <button onclick="closeStandardsBank()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:12px;width:40px;height:40px;cursor:pointer;font-size:18px">✕</button>
      </div>
      <div id="sbGradesLevel"><p style="color:rgba(255,255,255,.5);font-size:13px;margin-bottom:20px;text-align:center">اختر الصف / Select Grade</p><div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 lg:grid-cols-7 gap-4" id="sbGradesGrid"></div></div>
      <div id="sbSubjectsLevel" class="hidden"><p style="color:rgba(255,255,255,.5);font-size:13px;margin-bottom:20px;text-align:center">اختر المادة / Select Subject</p><div class="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-4" id="sbSubjectsGrid"></div></div>
      <div id="sbGlosLevel" class="hidden">
        <div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:18px;padding:20px;margin-bottom:20px">
          <h4 style="color:#FACC15;font-weight:800;margin-bottom:16px;font-size:14px">➕ إضافة ناتج تعلم / Add GLO</h4>
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mb-4">
            <div><label class="wizard-label">الناتج بالعربية</label><textarea id="gloAr" class="wizard-input" style="height:80px;resize:none" placeholder="اكتب ناتج التعلم..."></textarea></div>
            <div><label class="wizard-label">English GLO</label><textarea id="gloEn" class="wizard-input font-en" dir="ltr" style="height:80px;resize:none" placeholder="Write the learning outcome..."></textarea></div>
          </div>
          <div style="display:flex;gap:10px;flex-wrap:wrap">
            <button onclick="addGLOManual()" style="background:#FACC15;color:#1e3a8a;font-weight:800;border:none;border-radius:12px;padding:10px 20px;cursor:pointer;font-family:'Tajawal',sans-serif;font-size:14px">💾 إضافة / Add</button>
            <label style="display:flex;align-items:center;gap:8px;background:rgba(34,197,94,.2);border:1px solid rgba(34,197,94,.4);border-radius:12px;padding:10px 16px;cursor:pointer;font-size:13px;font-family:'Tajawal',sans-serif">📊 Excel<input type="file" accept=".xlsx,.xls,.csv" style="display:none" onchange="uploadGLOFile(event)"></label>
            <div id="uploadStatus" style="font-size:12px;color:rgba(255,255,255,.5);display:flex;align-items:center"></div>
          </div>
        </div>
        <div style="overflow-x:auto"><table class="admin-table" style="font-size:13px"><thead><tr><th class="font-en">GLO Code</th><th>الناتج بالعربية</th><th>English GLO</th><th>حذف</th></tr></thead><tbody id="gloTableBody"></tbody></table><div id="gloEmpty" style="text-align:center;color:rgba(255,255,255,.3);padding:32px;font-size:14px">لا توجد نواتج / No GLOs yet</div></div>
      </div>
    </div>
  </div>
</div>

<!-- STUDENT PORTAL -->
<div id="studentPortal" class="hidden w-full flex flex-col min-h-screen" style="padding:0">
  <!-- Header Bar -->
  <div style="background:linear-gradient(135deg,#1e3a8a,#7e22ce);padding:14px 20px;display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid rgba(250,204,21,.5);box-shadow:0 4px 20px rgba(0,0,0,.3)">
    <!-- لوجو يسار -->
    <div style="display:flex;align-items:center;gap:10px">
      <div style="width:44px;height:44px;background:white;border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:24px;font-weight:900;color:#1e3a8a;font-family:Montserrat,sans-serif">S</div>
      <div><div style="font-size:16px;font-weight:900;color:white;font-family:Montserrat,sans-serif;font-style:italic">Scholastic</div><div style="font-size:9px;color:rgba(255,255,255,.5);font-family:Montserrat,sans-serif;letter-spacing:1px">TESTING PLATFORM</div></div>
    </div>
    <!-- اسم الطالب وسط -->
    <div style="text-align:center">
      <div style="font-size:18px;font-weight:900;color:#FACC15" id="studentPortalName">Student Name</div>
      <div style="font-size:11px;color:rgba(255,255,255,.6);font-family:Montserrat,sans-serif" id="studentPortalSchool">School</div>
    </div>
    <!-- شهاداتي يمين -->
    <button onclick="showStudentCertificates()" style="background:rgba(250,204,21,.15);border:2px solid rgba(250,204,21,.4);border-radius:14px;padding:10px 16px;color:#FACC15;font-family:Tajawal,sans-serif;font-size:13px;font-weight:800;cursor:pointer;display:flex;flex-direction:column;align-items:center;gap:2px;transition:.2s" onmouseover="this.style.background='rgba(250,204,21,.25)'" onmouseout="this.style.background='rgba(250,204,21,.15)'">
      <span style="font-size:20px">🏅</span>
      <span>شهاداتي</span>
      <span style="font-size:9px;font-family:Montserrat,sans-serif;opacity:.7">My Certificates</span>
    </button>
  </div>
  <!-- Test Cards -->
  <div style="flex:1;padding:28px 20px">
    <div style="text-align:center;margin-bottom:24px">
      <div style="font-size:14px;color:rgba(255,255,255,.5);font-family:Montserrat,sans-serif">Select your test below / اختر اختبارك أدناه</div>
    </div>
    <div id="studentTestCards" class="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 gap-6 w-full max-w-5xl mx-auto"></div>
  </div>
  <!-- Footer -->
  <div style="padding:16px;text-align:center;border-top:1px solid rgba(255,255,255,.1)">
    <button onclick="backToMain()" style="color:rgba(255,255,255,.4);text-decoration:none;background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:8px 20px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px">تسجيل خروج / Logout</button>
  </div>
</div>

<script>
// ============================================================
// GITHUB STORAGE CONFIG
// ============================================================
var GH_TOKEN  = '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';
var GH_OWNER  = 'MohdGems2050';
var GH_REPO   = 'school-platform-data';
var GH_BRANCH = 'main';
var _ghSHAs   = {};

async function ghRead(filename){
  try{
    var controller=new AbortController();
    var tid=setTimeout(function(){controller.abort();},4000);
    var r=await fetch('https://api.github.com/repos/'+GH_OWNER+'/'+GH_REPO+'/contents/'+filename+'?ref='+GH_BRANCH+'&t='+Date.now(),{headers:{'Authorization':'token '+GH_TOKEN,'Accept':'application/vnd.github.v3+json'},signal:controller.signal});
    clearTimeout(tid);
    if(!r.ok) return null;
    var j=await r.json();
    _ghSHAs[filename]=j.sha;
    return JSON.parse(atob(j.content.replace(/\n/g,'')));
  }catch(e){return null;}
}

async function ghWrite(filename,data){
  try{
    var content=btoa(unescape(encodeURIComponent(JSON.stringify(data,null,2))));
    var body={message:'update '+filename,content:content,branch:GH_BRANCH};
    if(_ghSHAs[filename]) body.sha=_ghSHAs[filename];
    var r=await fetch('https://api.github.com/repos/'+GH_OWNER+'/'+GH_REPO+'/contents/'+filename,{method:'PUT',headers:{'Authorization':'token '+GH_TOKEN,'Accept':'application/vnd.github.v3+json','Content-Type':'application/json'},body:JSON.stringify(body)});
    var j=await r.json();
    if(j.content) _ghSHAs[filename]=j.content.sha;
    return r.ok;
  }catch(e){return false;}
}

// ============================================================
// DATA
// ============================================================
document.getElementById('marqueeContent').innerHTML += document.getElementById('marqueeContent').innerHTML;

var supervisors = [];
var schools     = [];
var myTests     = [];
var sbGLOs      = {};

// Load all data from GitHub on startup
async function initDataFromGitHub(){
  showGHLoader(true);
  try{
    // تشغيل كل الطلبات بالتوازي — أسرع بكثير
    var results = await Promise.all([
      ghRead('supervisors.json'),
      ghRead('schools.json'),
      ghRead('tests.json'),
      ghRead('glos.json'),
      ghRead('archive.json')
    ]);
    if(results[0]) supervisors=results[0];
    if(results[1]) schools=results[1];
    if(results[2]) myTests=results[2];
    if(results[3]) sbGLOs=results[3];
    if(results[4]) archivedTests=results[4]; else archivedTests=[];
  }catch(e){ console.warn('GH load error:',e); }
  showGHLoader(false);
  try{ if(typeof filterSchoolsByCountryAndCurriculum==='function') filterSchoolsByCountryAndCurriculum(); }catch(e){}
  try{ if(typeof renderSchools==='function') renderSchools(); }catch(e){}
}

function showGHLoader(show){
  var el=document.getElementById('ghLoader');
  if(!el) return;
  el.style.display=show?'flex':'none';
  if(show){
    clearTimeout(window._ghT);
    window._ghT=setTimeout(function(){
      var e=document.getElementById('ghLoader');
      if(e) e.style.display='none';
    },5000);
  }
}

function _saveDraft(){try{localStorage.setItem('sc_draft',JSON.stringify(testData));}catch(e){}}
function _loadDraft(){try{var s=localStorage.getItem('sc_draft');if(s)return JSON.parse(s);}catch(e){}return{domains:[],selectedSchools:[],logoSrc:'',displayMode:1};}

// Async save wrappers
async function saveSupervisors(){await ghWrite('supervisors.json',supervisors);}
async function saveSchoolsGH(){await ghWrite('schools.json',schools);}
async function saveMyTests(){await ghWrite('tests.json',myTests);}
async function saveGLOs(){await ghWrite('glos.json',sbGLOs);}
async function saveStudentProgress(){await ghWrite('tests.json',myTests);}
var testData = _loadDraft();
if(!testData.displayMode) testData.displayMode=1;
var sw_navDir=1; // 1=forward, -1=back (for animation)
var currentDomainIndex=-1, currentQuestionIndex=-1, currentEditingQ=null, editingTestId=null;


// ── Display Mode ──
// ══ Media Source Visibility Toggle ══
function toggleMediaVisibility(btn,key){
  if(!currentEditingQ) return;
  if(!currentEditingQ.mediaVisible) currentEditingQ.mediaVisible={};
  var isOn=currentEditingQ.mediaVisible[key]!==false;
  currentEditingQ.mediaVisible[key]=!isOn;
  btn.className='media-toggle '+(isOn?'off':'on');
  btn.querySelector('.mt-dot').textContent=isOn?'🔵':'🟢';
  btn.querySelector('.mt-label').textContent=isOn?'مخفي':'ظاهر';
}
function makeMediaToggle(key,label){
  var isOn=!(currentEditingQ&&currentEditingQ.mediaVisible&&currentEditingQ.mediaVisible[key]===false);
  return '<button class="media-toggle '+(isOn?'on':'off')+'" onclick="toggleMediaVisibility(this,\''+key+'\')">'
    +'<span class="mt-dot">'+(isOn?'🟢':'🔵')+'</span>'
    +'<span class="mt-label">'+(isOn?'ظاهر':'مخفي')+'</span>'
    +' '+label+'</button>';
}
// ══ White Paper Preview — معاينة الورقة البيضاء ══
function openWhitePaperPreview(domIdx){
  var d=testData.domains[domIdx];
  if(!d) return;
  var hasQ=d.questions&&d.questions.length||(d.hasBranches&&d.branches&&d.branches.some(function(b){return b.questions&&b.questions.length;}));
  if(!hasQ){scWarn('لا توجد أسئلة — أضف أسئلة أولاً','Add questions first');return;}
  // بناء HTML الورقة البيضاء
  var html=buildWhitePaperHTML(d,domIdx);
  // عرض في overlay
  var overlay=document.getElementById('wpPreviewOverlay');
  if(!overlay){
    overlay=document.createElement('div');
    overlay.id='wpPreviewOverlay';
    overlay.style.cssText='position:fixed;inset:0;z-index:9998;background:#c8c8c0;overflow-y:auto;font-family:Tajawal,"Times New Roman",serif';
    document.body.appendChild(overlay);
  }
  overlay.innerHTML=
    // شريط تحكم علوي ثابت
    '<div style="position:sticky;top:0;z-index:10;background:linear-gradient(135deg,#1e0d42,#2a1860);border-bottom:2px solid rgba(250,204,21,.3);padding:10px 20px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">'
    +'<div style="display:flex;align-items:center;gap:10px">'
    +'<div style="width:36px;height:36px;background:linear-gradient(135deg,#f59e0b,#d97706);border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px">📄</div>'
    +'<div><div style="font-size:14px;font-weight:900;color:#FACC15;font-family:Montserrat,sans-serif">معاينة الورقة البيضاء</div><div style="font-size:10px;color:rgba(255,255,255,.5);font-family:Montserrat,sans-serif">White Paper Preview</div></div>'
    +'</div>'
    +'<div style="display:flex;gap:10px;align-items:center">'
    +'<button onclick="printWhitePaper()" style="background:rgba(59,130,246,.25);border:1px solid rgba(59,130,246,.5);color:#93c5fd;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🖨️ طباعة / Print</button>'
    +'<button onclick="document.getElementById(\'wpPreviewOverlay\').style.display=\'none\';_previewReturnStep=3;" style="background:rgba(239,68,68,.2);border:1px solid rgba(239,68,68,.4);color:#fca5a5;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Montserrat,sans-serif;font-size:12px;font-weight:700">✕ إغلاق</button>'
    +'</div></div>'
    // الورقة
    +'<div style="padding:24px 16px;display:flex;flex-direction:column;align-items:center">'+html+'</div>';
  overlay.style.display='block';
}

function buildWhitePaperHTML(d,domIdx){
  var instrAr=testData.instructionsAr||'';
  var instrEn=testData.instructionsEn||'';
  var logoHtml=testData.logoSrc
    ?'<img src="'+testData.logoSrc+'" style="max-height:60px;max-width:80px;object-fit:contain;background:white">'
    :'<div style="width:60px;height:60px;background:linear-gradient(135deg,#f59e0b,#d97706);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:28px">🎓</div>';
  var html='<div class="wp-page" style="max-width:760px;width:100%;background:white;padding:32px 36px;box-shadow:0 4px 24px rgba(0,0,0,.18);border:1px solid #e8e8e8;border-radius:2px;margin-bottom:0">';
  // هيدر الورقة
  html+='<div style="display:flex;align-items:center;justify-content:space-between;border-bottom:3px double #111;padding-bottom:12px;margin-bottom:16px;gap:12px">'
    +logoHtml
    +'<div style="flex:1;text-align:center">'
    +'<div style="font-size:18px;font-weight:900;color:#1a1a2e;margin-bottom:4px">'+(testData.testName||'اختبار')+'</div>'
    +'<div style="font-size:12px;color:#555;font-family:Montserrat,sans-serif">'+(testData.subject||'')+(testData.grade?' | '+testData.grade:'')+(testData.term?' | Term '+testData.term:'')+'</div>'
    +'</div>'
    +logoHtml
    +'</div>';
  // التعليمات
  if(instrAr||instrEn){
    html+='<table style="width:100%;border-collapse:collapse;margin-bottom:16px;font-size:12px"><tr>'
      +(instrAr?'<td style="border:1px solid #ccc;padding:10px;vertical-align:top;width:50%" dir="rtl">'+instrAr+'</td>':'')
      +(instrEn?'<td style="border:1px solid #ccc;padding:10px;vertical-align:top;width:50%" dir="ltr">'+instrEn+'</td>':'')
      +'</tr></table>';
  }
  // عنوان المجال
  html+='<div style="background:#1e3a8a;color:white;padding:8px 16px;border-radius:3px;margin-bottom:16px;display:flex;justify-content:space-between;align-items:center">'
    +'<span style="font-size:14px;font-weight:900">'+(d.nameAr||'المجال')+'</span>'
    +'<span style="font-size:12px;font-family:Montserrat,sans-serif;opacity:.8">'+(d.nameEn||'Domain')+'</span>'
    +'<span style="font-size:11px;font-family:Montserrat,sans-serif;opacity:.7">'+d.weight+'%</span>'
    +'</div>';
  // الأسئلة: مباشرة أو بفروع
  if(d.hasBranches&&d.branches&&d.branches.length){
    var brLabels=['أولاً','ثانياً','ثالثاً','رابعاً','خامساً','سادساً','سابعاً','ثامناً'];
    var brLabelsEn=['First','Second','Third','Fourth','Fifth','Sixth','Seventh','Eighth'];
    d.branches.forEach(function(br,bi){
      if(bi>0){
        // فاصل أزرق بين الفروع
        html+='<div style="border-top:3px solid #1e3a8a;margin:20px 0;position:relative">'
          +'<div style="position:absolute;top:-12px;right:16px;background:white;padding:0 8px;font-size:12px;font-weight:800;color:#1e3a8a">'+(brLabels[bi]||'فرع '+(bi+1))+' / '+(brLabelsEn[bi]||'Branch '+(bi+1))+'</div>'
          +'</div>';
      } else {
        html+='<div style="font-size:13px;font-weight:900;color:#1e3a8a;margin-bottom:12px;padding:6px 12px;background:#eff6ff;border-right:4px solid #1e3a8a;border-radius:3px">'
          +(brLabels[bi]||'أولاً')+': '+(br.nameAr||br.nameEn||'الفرع الأول')
          +'<span style="float:left;font-family:Montserrat,sans-serif;font-size:11px">'+(brLabelsEn[bi]||'First')+': '+(br.nameEn||br.nameAr||'First Branch')+'</span>'
          +'<span style="float:left;margin-left:8px;color:#64748b;font-size:11px">'+br.weight+'%</span>'
          +'</div>';
      }
      html+=renderWPQuestions(br.questions||[],br.weight||0);
    });
  } else {
    html+=renderWPQuestions(d.questions||[],d.weight||0);
  }
  // تذييل الورقة
  html+='<div style="border-top:1px solid #bbb;margin-top:20px;padding-top:8px;display:flex;justify-content:space-between;font-size:10px;color:#777;font-family:Montserrat,sans-serif">'
    +'<span>Scholastic Testing Platform</span>'
    +'<span>Page 1 of 1</span>'
    +'</div>';
  html+='</div>';
  return html;
}

function renderWPQuestions(questions,totalWeight){
  var html='';
  var labels=['A','B','C','D','E','F'];
  var usedWeight=0;
  questions.forEach(function(q,qi){
    var qWeight=parseFloat(q.score||0);
    usedWeight+=qWeight;
    var remaining=parseFloat((totalWeight-usedWeight).toFixed(2));
    // فاصل ذهبي بين الأسئلة
    if(qi>0){
      html+='<div style="height:2px;background:linear-gradient(90deg,transparent,#FACC15 20%,#F59E0B 50%,#FACC15 80%,transparent);margin:14px 0;box-shadow:0 1px 4px rgba(245,158,11,.3)"></div>';
    }
    html+='<div style="margin-bottom:4px">';
    // رأس السؤال + الدرجة
    html+='<div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px;margin-bottom:8px">'
      +'<div style="display:flex;gap:8px;flex:1">'
      +'<span style="font-weight:900;color:#1e3a8a;font-size:14px;flex-shrink:0;min-width:22px">'+(qi+1)+'.</span>'
      +'<div style="font-size:14px;font-weight:600;line-height:1.7;color:#1a1a2e;flex:1" dir="auto">'+(q.stemHtml||q.stemText||'')+'</div>'
      +'</div>'
      +'<div style="text-align:center;flex-shrink:0">'
      +'<div style="background:#1e3a8a;color:white;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:700;white-space:nowrap;font-family:Montserrat,sans-serif">'+qWeight+'%</div>'
      +'<div style="font-size:9px;color:#94a3b8;margin-top:2px;font-family:Montserrat,sans-serif">متبقي: '+remaining+'%</div>'
      +'</div>'
      +'</div>';
    // صورة/ميديا
    if(q.mediaHtml&&!(q.mediaVisible&&q.mediaVisible.img===false)){
      html+='<div style="margin:8px 0 8px 22px">'+q.mediaHtml+'</div>';
    }
    // جسم السؤال (النص المكتوب من المحرر)
    if(q.bodyHtml){
      html+='<div style="margin-right:22px;margin-bottom:8px;font-size:13px;line-height:1.7;color:#1a1a2e" dir="auto">'+q.bodyHtml+'</div>';
    }
    if(q.questionText){
      html+='<div style="margin-right:22px;margin-bottom:8px;font-size:13px;font-weight:700;color:#1e3a8a" dir="auto">'+q.questionText+'</div>';
    }
    // جسم الإجابة
    html+='<div style="margin-right:22px">';
    if(q.type==='mcq'){
      html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-top:6px">';
      (q.options||[]).forEach(function(o,i){
        html+='<div style="display:flex;align-items:center;gap:6px;font-size:12px;border:1px solid #e2e8f0;border-radius:4px;padding:5px 8px">'
          +'<div style="width:14px;height:14px;border-radius:50%;border:2px solid #94a3b8;flex-shrink:0"></div>'
          +'<span style="font-weight:700;color:#1e3a8a;font-size:10px">'+(labels[i]||i+1)+'.</span>'
          +'<span dir="auto">'+(o||'')+'</span></div>';
      });
      html+='</div>';
    } else if(q.type==='truefalse'){
      (q.statements||[]).forEach(function(s,si){
        html+='<div style="display:flex;align-items:center;gap:10px;margin-bottom:7px;font-size:12px">'
          +'<span style="flex:1" dir="auto">'+(si+1)+'. '+(s.text||'')+'</span>'
          +'<div style="display:flex;gap:6px;flex-shrink:0">'
          +'<div style="border:1.5px solid #22c55e;border-radius:4px;padding:3px 10px;font-size:11px;font-weight:700;color:#15803d">✓ صواب</div>'
          +'<div style="border:1.5px solid #ef4444;border-radius:4px;padding:3px 10px;font-size:11px;font-weight:700;color:#b91c1c">✗ خطأ</div>'
          +'</div></div>';
      });
    } else if(q.type==='reading'||q.type==='writingskill'){
      if(q.passageHtml) html+='<div style="background:#f9fafb;border:1px solid #e2e8f0;border-radius:4px;padding:10px;margin-bottom:8px;font-size:12px;line-height:1.8" dir="auto">'+q.passageHtml+'</div>';
      if(q.answerStemHtml) html+='<div style="font-size:13px;font-weight:700;margin-bottom:8px" dir="auto">'+q.answerStemHtml+'</div>';
      if(q.ansType==='mcq'){
        html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:5px">';
        (q.options||[]).forEach(function(o,i){html+='<div style="display:flex;align-items:center;gap:5px;font-size:12px;border:1px solid #e2e8f0;border-radius:4px;padding:4px 8px"><div style="width:13px;height:13px;border-radius:50%;border:2px solid #94a3b8;flex-shrink:0"></div><span style="font-weight:700;color:#1e3a8a;font-size:10px">'+(labels[i]||i+1)+'.</span><span dir="auto">'+(o||'')+'</span></div>';});
        html+='</div>';
      } else {
        for(var l=0;l<4;l++) html+='<div style="border-bottom:1px solid #ccc;height:28px;margin-bottom:5px"></div>';
      }
    } else if(q.type==='ordering'){
      var pvGroups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words}]:[]);
      pvGroups.forEach(function(g,gi){
        if(pvGroups.length>1) html+='<div style="font-size:11px;font-weight:800;color:#1e3a8a;margin-top:8px">الجملة '+(gi+1)+' / Sentence '+(gi+1)+'</div>';
        html+='<div style="display:flex;flex-wrap:wrap;gap:8px;margin-top:6px">';
        (g.words||[]).forEach(function(w){html+='<div style="border:1.5px dashed #1e3a8a;border-radius:6px;padding:5px 14px;font-size:13px;font-weight:700;color:#1e3a8a;min-width:60px;text-align:center">'+w+'</div>';});
        html+='</div><div style="margin-top:8px">';
        for(var wl=0;wl<(g.words||[]).length;wl++) html+='<div style="display:inline-block;border-bottom:1px solid #ccc;height:28px;min-width:80px;margin-left:8px;margin-bottom:6px"></div>';
        html+='</div>';
      });
    } else if(q.type==='speaking'||q.type==='oral'){
      if(q.type==='oral'&&q.oralText) html+='<div style="border:2px solid #fde68a;border-radius:8px;padding:14px;background:#fffbeb;font-size:'+(q.speakingSize||16)+'px;line-height:2;font-family:'+(q.speakingFont||'Tajawal')+',serif" dir="auto">'+q.oralText+'</div>';
      html+='<div style="background:#f0f9ff;border:1px dashed #0891b2;border-radius:6px;padding:10px;text-align:center;color:#0891b2;font-size:12px;margin-top:8px;font-family:Montserrat,sans-serif">🎙 يُسجَّل صوت الطالب / Student records voice</div>';
    } else if(q.type==='classify'){
      var cols=q.columns||[];
      html+='<table style="width:100%;border-collapse:collapse;margin-top:8px"><thead><tr>';
      cols.forEach(function(c){html+='<th style="background:#1e3a8a;color:white;padding:6px;font-size:12px;border:1px solid #fff;text-align:center">'+c+'</th>';});
      html+='</tr></thead><tbody>';
      var nR=Math.ceil((q.items||[]).length/Math.max(1,cols.length));
      for(var r=0;r<nR;r++){html+='<tr>';cols.forEach(function(){html+='<td style="border:1px solid #e2e8f0;height:36px;padding:4px"></td>';});html+='</tr>';}
      html+='</tbody></table>';
    } else {
      // writing / other
      for(var al=0;al<4;al++) html+='<div style="border-bottom:1px solid #ccc;height:28px;margin-bottom:5px"></div>';
    }
    html+='</div></div>'; // end margin-right + question div
  });
  return html;
}
function printWhitePaper(){
  var el=document.querySelector('#wpPreviewOverlay .wp-page');
  if(!el){window.print();return;}
  var win=window.open('','_blank');
  win.document.write('<html><head><meta charset="UTF-8"><style>body{font-family:Tajawal,"Times New Roman",serif;margin:20px}@media print{body{margin:0}}</style></head><body>'+el.outerHTML+'</body></html>');
  win.document.close();
  win.print();
}
var SCHOLASTIC_LOGO_SVG='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 200" width="56" height="56"><rect width="200" height="200" rx="16" fill="white"/><rect x="4" y="4" width="192" height="192" rx="13" fill="none" stroke="#B8860B" stroke-width="3"/><rect x="9" y="9" width="182" height="182" rx="10" fill="none" stroke="#DAA520" stroke-width="1.5"/><rect x="12" y="12" width="176" height="176" rx="8" fill="none" stroke="#B8860B" stroke-width=".5" stroke-dasharray="4 3"/>'
  +'<polygon points="100,42 68,62 68,68 132,68 132,62" fill="#111"/>'
  +'<rect x="85" y="66" width="30" height="5" rx="2" fill="#222"/>'
  +'<rect x="78" y="70" width="44" height="28" rx="4" fill="#1a1a1a"/>'
  +'<rect x="82" y="74" width="36" height="20" rx="3" fill="#333"/>'
  +'<line x1="91" y1="74" x2="91" y2="94" stroke="#555" stroke-width="1"/>'
  +'<line x1="100" y1="74" x2="100" y2="94" stroke="#555" stroke-width="1"/>'
  +'<line x1="109" y1="74" x2="109" y2="94" stroke="#555" stroke-width="1"/>'
  +'<rect x="96" y="68" width="8" height="3" rx="1" fill="#DAA520"/>'
  +'<circle cx="100" cy="67" r="3" fill="#DAA520"/>'
  +'<polygon points="100,42 58,60 100,70 142,60" fill="#111"/>'
  +'<polygon points="100,42 58,60 62,62 100,46 138,62 142,60" fill="#1a1a1a"/>'
  +'<line x1="142" y1="60" x2="142" y2="82" stroke="#111" stroke-width="3" stroke-linecap="round"/>'
  +'<ellipse cx="142" cy="84" rx="5" ry="3" fill="#111"/>'
  +'<path d="M137 84 Q142 90 147 84" fill="#111"/>'
  +'<text x="100" y="124" text-anchor="middle" font-family="Georgia,serif" font-size="11" font-weight="bold" fill="#B8860B" letter-spacing="1">SCHOLASTIC</text>'
  +'<line x1="30" y1="128" x2="170" y2="128" stroke="#DAA520" stroke-width=".8"/>'
  +'<text x="100" y="143" text-anchor="middle" font-family="Georgia,serif" font-size="7.5" fill="#555" letter-spacing=".5">TESTING PLATFORM</text>'
  +'<text x="100" y="158" text-anchor="middle" font-family="Arial,sans-serif" font-size="7" fill="#888">منصة الاختبارات الأكاديمية</text>'
  +'<path d="M25 108 Q100 100 175 108" fill="none" stroke="#DAA520" stroke-width=".6" opacity=".5"/>'
  +'<path d="M25 162 Q100 170 175 162" fill="none" stroke="#DAA520" stroke-width=".6" opacity=".5"/>'
  +'</svg>';

var SCHOLASTIC_SEAL_SVG='<svg width="64" height="64" viewBox="0 0 460 460" xmlns="http://www.w3.org/2000/svg" role="img" aria-label="Scholastic seal"><defs><path id="wpTopArc" d="M 52,230 A 178,178 0 0,1 408,230"/><path id="wpBotArc" d="M 52,230 A 178,178 0 0,0 408,230"/><linearGradient id="wpNavyGrad" x1="0%" y1="0%" x2="100%" y2="100%"><stop offset="0%" stop-color="#233A73"/><stop offset="100%" stop-color="#141F44"/></linearGradient></defs><circle cx="230" cy="230" r="218" fill="none" stroke="#AD8628" stroke-width="1"/><circle cx="230" cy="230" r="213" fill="url(#wpNavyGrad)"/><circle cx="230" cy="230" r="206" fill="none" stroke="#D9B872" stroke-width="0.8" opacity="0.6"/><circle cx="230" cy="230" r="201" fill="#FFFFFF"/><g stroke="#141F44" stroke-width="0.6" opacity="0.3"><line x1="423.00" y1="230.00" x2="426.00" y2="230.00"/><line x1="422.74" y1="240.10" x2="425.73" y2="240.26"/><line x1="421.94" y1="250.17" x2="424.93" y2="250.49"/><line x1="420.62" y1="260.19" x2="423.59" y2="260.66"/><line x1="418.78" y1="270.13" x2="421.72" y2="270.75"/><line x1="416.42" y1="279.95" x2="419.32" y2="280.73"/><line x1="413.55" y1="289.64" x2="416.41" y2="290.57"/><line x1="410.18" y1="299.17" x2="412.98" y2="300.24"/><line x1="406.31" y1="308.50" x2="409.05" y2="309.72"/><line x1="401.96" y1="317.62" x2="404.64" y2="318.98"/><line x1="397.14" y1="326.50" x2="399.74" y2="328.00"/><line x1="391.86" y1="335.12" x2="394.38" y2="336.75"/><line x1="386.14" y1="343.44" x2="388.57" y2="345.21"/><line x1="379.99" y1="351.46" x2="382.32" y2="353.35"/><line x1="373.43" y1="359.14" x2="375.66" y2="361.15"/><line x1="366.47" y1="366.47" x2="368.59" y2="368.59"/><line x1="359.14" y1="373.43" x2="361.15" y2="375.66"/><line x1="351.46" y1="379.99" x2="353.35" y2="382.32"/><line x1="343.44" y1="386.14" x2="345.21" y2="388.57"/><line x1="335.12" y1="391.86" x2="336.75" y2="394.38"/><line x1="326.50" y1="397.14" x2="328.00" y2="399.74"/><line x1="317.62" y1="401.96" x2="318.98" y2="404.64"/><line x1="308.50" y1="406.31" x2="309.72" y2="409.05"/><line x1="299.17" y1="410.18" x2="300.24" y2="412.98"/><line x1="289.64" y1="413.55" x2="290.57" y2="416.41"/><line x1="279.95" y1="416.42" x2="280.73" y2="419.32"/><line x1="270.13" y1="418.78" x2="270.75" y2="421.72"/><line x1="260.19" y1="420.62" x2="260.66" y2="423.59"/><line x1="250.17" y1="421.94" x2="250.49" y2="424.93"/><line x1="240.10" y1="422.74" x2="240.26" y2="425.73"/><line x1="230.00" y1="423.00" x2="230.00" y2="426.00"/><line x1="219.90" y1="422.74" x2="219.74" y2="425.73"/><line x1="209.83" y1="421.94" x2="209.51" y2="424.93"/><line x1="199.81" y1="420.62" x2="199.34" y2="423.59"/><line x1="189.87" y1="418.78" x2="189.25" y2="421.72"/><line x1="180.05" y1="416.42" x2="179.27" y2="419.32"/><line x1="170.36" y1="413.55" x2="169.43" y2="416.41"/><line x1="160.83" y1="410.18" x2="159.76" y2="412.98"/><line x1="151.50" y1="406.31" x2="150.28" y2="409.05"/><line x1="142.38" y1="401.96" x2="141.02" y2="404.64"/><line x1="133.50" y1="397.14" x2="132.00" y2="399.74"/><line x1="124.88" y1="391.86" x2="123.25" y2="394.38"/><line x1="116.56" y1="386.14" x2="114.79" y2="388.57"/><line x1="108.54" y1="379.99" x2="106.65" y2="382.32"/><line x1="100.86" y1="373.43" x2="98.85" y2="375.66"/><line x1="93.53" y1="366.47" x2="91.41" y2="368.59"/><line x1="86.57" y1="359.14" x2="84.34" y2="361.15"/><line x1="80.01" y1="351.46" x2="77.68" y2="353.35"/><line x1="73.86" y1="343.44" x2="71.43" y2="345.21"/><line x1="68.14" y1="335.12" x2="65.62" y2="336.75"/><line x1="62.86" y1="326.50" x2="60.26" y2="328.00"/><line x1="58.04" y1="317.62" x2="55.36" y2="318.98"/><line x1="53.69" y1="308.50" x2="50.95" y2="309.72"/><line x1="49.82" y1="299.17" x2="47.02" y2="300.24"/><line x1="46.45" y1="289.64" x2="43.59" y2="290.57"/><line x1="43.58" y1="279.95" x2="40.68" y2="280.73"/><line x1="41.22" y1="270.13" x2="38.28" y2="270.75"/><line x1="39.38" y1="260.19" x2="36.41" y2="260.66"/><line x1="38.06" y1="250.17" x2="35.07" y2="250.49"/><line x1="37.26" y1="240.10" x2="34.27" y2="240.26"/><line x1="37.00" y1="230.00" x2="34.00" y2="230.00"/><line x1="37.26" y1="219.90" x2="34.27" y2="219.74"/><line x1="38.06" y1="209.83" x2="35.07" y2="209.51"/><line x1="39.38" y1="199.81" x2="36.41" y2="199.34"/><line x1="41.22" y1="189.87" x2="38.28" y2="189.25"/><line x1="43.58" y1="180.05" x2="40.68" y2="179.27"/><line x1="46.45" y1="170.36" x2="43.59" y2="169.43"/><line x1="49.82" y1="160.83" x2="47.02" y2="159.76"/><line x1="53.69" y1="151.50" x2="50.95" y2="150.28"/><line x1="58.04" y1="142.38" x2="55.36" y2="141.02"/><line x1="62.86" y1="133.50" x2="60.26" y2="132.00"/><line x1="68.14" y1="124.88" x2="65.62" y2="123.25"/><line x1="73.86" y1="116.56" x2="71.43" y2="114.79"/><line x1="80.01" y1="108.54" x2="77.68" y2="106.65"/><line x1="86.57" y1="100.86" x2="84.34" y2="98.85"/><line x1="93.53" y1="93.53" x2="91.41" y2="91.41"/><line x1="100.86" y1="86.57" x2="98.85" y2="84.34"/><line x1="108.54" y1="80.01" x2="106.65" y2="77.68"/><line x1="116.56" y1="73.86" x2="114.79" y2="71.43"/><line x1="124.88" y1="68.14" x2="123.25" y2="65.62"/><line x1="133.50" y1="62.86" x2="132.00" y2="60.26"/><line x1="142.38" y1="58.04" x2="141.02" y2="55.36"/><line x1="151.50" y1="53.69" x2="150.28" y2="50.95"/><line x1="160.83" y1="49.82" x2="159.76" y2="47.02"/><line x1="170.36" y1="46.45" x2="169.43" y2="43.59"/><line x1="180.05" y1="43.58" x2="179.27" y2="40.68"/><line x1="189.87" y1="41.22" x2="189.25" y2="38.28"/><line x1="199.81" y1="39.38" x2="199.34" y2="36.41"/><line x1="209.83" y1="38.06" x2="209.51" y2="35.07"/><line x1="219.90" y1="37.26" x2="219.74" y2="34.27"/><line x1="230.00" y1="37.00" x2="230.00" y2="34.00"/><line x1="240.10" y1="37.26" x2="240.26" y2="34.27"/><line x1="250.17" y1="38.06" x2="250.49" y2="35.07"/><line x1="260.19" y1="39.38" x2="260.66" y2="36.41"/><line x1="270.13" y1="41.22" x2="270.75" y2="38.28"/><line x1="279.95" y1="43.58" x2="280.73" y2="40.68"/><line x1="289.64" y1="46.45" x2="290.57" y2="43.59"/><line x1="299.17" y1="49.82" x2="300.24" y2="47.02"/><line x1="308.50" y1="53.69" x2="309.72" y2="50.95"/><line x1="317.62" y1="58.04" x2="318.98" y2="55.36"/><line x1="326.50" y1="62.86" x2="328.00" y2="60.26"/><line x1="335.12" y1="68.14" x2="336.75" y2="65.62"/><line x1="343.44" y1="73.86" x2="345.21" y2="71.43"/><line x1="351.46" y1="80.01" x2="353.35" y2="77.68"/><line x1="359.14" y1="86.57" x2="361.15" y2="84.34"/><line x1="366.47" y1="93.53" x2="368.59" y2="91.41"/><line x1="373.43" y1="100.86" x2="375.66" y2="98.85"/><line x1="379.99" y1="108.54" x2="382.32" y2="106.65"/><line x1="386.14" y1="116.56" x2="388.57" y2="114.79"/><line x1="391.86" y1="124.88" x2="394.38" y2="123.25"/><line x1="397.14" y1="133.50" x2="399.74" y2="132.00"/><line x1="401.96" y1="142.38" x2="404.64" y2="141.02"/><line x1="406.31" y1="151.50" x2="409.05" y2="150.28"/><line x1="410.18" y1="160.83" x2="412.98" y2="159.76"/><line x1="413.55" y1="170.36" x2="416.41" y2="169.43"/><line x1="416.42" y1="180.05" x2="419.32" y2="179.27"/><line x1="418.78" y1="189.87" x2="421.72" y2="189.25"/><line x1="420.62" y1="199.81" x2="423.59" y2="199.34"/><line x1="421.94" y1="209.83" x2="424.93" y2="209.51"/><line x1="422.74" y1="219.90" x2="425.73" y2="219.74"/></g><circle cx="230" cy="230" r="196" fill="none" stroke="#141F44" stroke-width="1.4"/><circle cx="230" cy="230" r="168" fill="none" stroke="#AD8628" stroke-width="1"/><text font-family="Jost,Arial,sans-serif" font-size="11" font-weight="600" fill="#141F44" letter-spacing="1.6"><textPath href="#wpTopArc" startOffset="50%" text-anchor="middle">INTERNATIONAL STANDARDIZED TESTING PLATFORM</textPath></text><text font-family="Tajawal,Tahoma,Arial,sans-serif" font-size="14.5" font-weight="800" fill="#AD8628" direction="rtl"><textPath href="#wpBotArc" startOffset="50%" text-anchor="middle" direction="rtl">منصة الاختبارات المعيارية الدولية للمواد المدرسية</textPath></text><g fill="#AD8628"><rect x="226" y="14" width="8" height="8" transform="rotate(45 230 18)"/><rect x="226" y="438" width="8" height="8" transform="rotate(45 230 442)"/><rect x="14" y="226" width="8" height="8" transform="rotate(45 18 230)"/><rect x="438" y="226" width="8" height="8" transform="rotate(45 442 230)"/></g><g stroke-width="1.8"><line x1="230.0" y1="98.0" x2="325.3" y2="143.9" stroke="#3B5BA5"/><line x1="325.3" y1="143.9" x2="357.8" y2="239.5" stroke="#C8A020"/><line x1="357.8" y1="239.5" x2="291.5" y2="320.1" stroke="#7B4FA0"/><line x1="291.5" y1="320.1" x2="168.5" y2="320.1" stroke="#2E8B6A"/><line x1="168.5" y1="320.1" x2="102.2" y2="239.5" stroke="#C0392B"/><line x1="102.2" y1="239.5" x2="134.7" y2="143.9" stroke="#4A90C8"/><line x1="134.7" y1="143.9" x2="230.0" y2="98.0" stroke="#E67E22"/></g><g font-family="Jost,Arial,sans-serif" font-size="8.8" font-weight="600" fill="#FFFFFF" letter-spacing="0.5"><circle cx="230.0" cy="98.0" r="27" fill="#3B5BA5" stroke="#AD8628" stroke-width="1.4"/><text x="230.0" y="98.0" text-anchor="middle" dominant-baseline="middle">ARABIC</text><circle cx="325.3" cy="143.9" r="27" fill="#C8A020" stroke="#AD8628" stroke-width="1.4"/><text x="325.3" y="143.9" text-anchor="middle" dominant-baseline="middle">ISLAMIC</text><circle cx="357.8" cy="239.5" r="27" fill="#7B4FA0" stroke="#AD8628" stroke-width="1.4"/><text x="357.8" y="239.5" text-anchor="middle" dominant-baseline="middle">SOCIAL</text><circle cx="291.5" cy="320.1" r="27" fill="#2E8B6A" stroke="#AD8628" stroke-width="1.4"/><text x="291.5" y="320.1" text-anchor="middle" dominant-baseline="middle">SCIENCE</text><circle cx="168.5" cy="320.1" r="27" fill="#C0392B" stroke="#AD8628" stroke-width="1.4"/><text x="168.5" y="320.1" text-anchor="middle" dominant-baseline="middle">ENGLISH</text><circle cx="102.2" cy="239.5" r="27" fill="#4A90C8" stroke="#AD8628" stroke-width="1.4"/><text x="102.2" y="239.5" text-anchor="middle" dominant-baseline="middle">MATH</text><circle cx="134.7" cy="143.9" r="27" fill="#E67E22" stroke="#AD8628" stroke-width="1.4"/><text x="134.7" y="143.9" text-anchor="middle" dominant-baseline="middle">ICT</text></g><text x="230" y="222" text-anchor="middle" font-family="\'Playfair Display\',Georgia,serif" font-size="25" font-weight="800" fill="#141F44" letter-spacing="3.5">SCHOLASTIC</text><line x1="150" y1="231" x2="310" y2="231" stroke="#AD8628" stroke-width="1.2"/><circle cx="230" cy="231" r="2.6" fill="#AD8628"/><text x="230" y="247" text-anchor="middle" font-family="Jost,Arial,sans-serif" font-size="9" font-weight="600" fill="#8A8A8A" letter-spacing="3.5">EST. 2025</text></svg>';
var SCHOLASTIC_EAGLE_LOGO_PNG='data:image/png;base64,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';

// ══ Preview In Box ══
var _previewBoxMode=0;
var _previewBoxData=null;
function previewInBox(m){
  if(!testData.domains||!testData.domains.length){scWarn('أضف مجالات وأسئلة أولاً','Add domains first');return;}
  // حدد البطاقة المضغوطة
  [1,2,3,4].forEach(function(i){
    var c=document.getElementById('modeCard'+i);
    if(c){c.classList.remove('mode-active');c.style.borderColor='rgba(255,255,255,.12)';}
  });
  var ac=document.getElementById('modeCard'+m);
  if(ac){ac.classList.add('mode-active');ac.style.borderColor=['','#1d4ed8','#7c3aed','#0891b2','#6b7280'][m]||'#22c55e';}
  _previewBoxMode=m;
  var mNames={1:'📋 الكلاسيكي',2:'🎯 التركيز الكامل',3:'🌊 الانسيابي',4:'📄 الورقة البيضاء'};
  var el=document.getElementById('previewBoxModeName');
  if(el) el.textContent='النمط: '+(mNames[m]||m);
  // ابنِ محتوى المعاينة داخل الصندوق
  buildPreviewBoxContent(m);
}
function buildPreviewBoxContent(m){
  var cont=document.getElementById('previewBoxContent');
  if(!cont) return;
  if(!_previewBoxData) _previewBoxData=JSON.parse(JSON.stringify(testData));
  testData.displayMode=m;
  var d=testData.domains[0];
  if(!d){cont.innerHTML='<div style="color:#94a3b8;padding:40px;text-align:center;font-family:Tajawal,sans-serif">لا توجد مجالات</div>';return;}
  var qs=d.questions||[];
  if(d.hasBranches&&d.branches&&d.branches[0]) qs=d.branches[0].questions||[];
  if(!qs.length){cont.innerHTML='<div style="color:#94a3b8;padding:40px;text-align:center;font-family:Tajawal,sans-serif">لا توجد أسئلة</div>';return;}
  // عرض التعليمات أولاً كطالب حقيقي
  var instrAr=testData.instructionsAr||'';
  var instrEn=testData.instructionsEn||'';
  var instrHtml='<div style="font-family:Tajawal,sans-serif;padding:20px;color:#1a1a2e;height:420px;overflow:auto;background:linear-gradient(135deg,#f8fafc,#eff6ff)">'
    +'<div style="max-width:600px;margin:0 auto">'
    +'<div style="text-align:center;margin-bottom:20px"><div style="font-size:32px;margin-bottom:8px">📋</div>'
    +'<div style="font-size:18px;font-weight:900;color:#1e3a8a">'+(testData.testName||'الاختبار')+'</div>'
    +(testData.grade||testData.subject?'<div style="font-size:13px;color:#64748b;margin-top:4px">'+(testData.subject||'')+(testData.grade?' | '+testData.grade:'')+'</div>':'')
    +'</div>'
    +(instrAr?'<div style="background:white;border-radius:14px;padding:16px;margin-bottom:12px;border-right:4px solid #1e3a8a;box-shadow:0 2px 8px rgba(0,0,0,.06)" dir="rtl">'+instrAr+'</div>':'')
    +(instrEn?'<div style="background:white;border-radius:14px;padding:16px;margin-bottom:16px;border-left:4px solid #3b82f6;box-shadow:0 2px 8px rgba(0,0,0,.06)" dir="ltr">'+instrEn+'</div>':'')
    +(!instrAr&&!instrEn?'<div style="background:white;border-radius:14px;padding:16px;margin-bottom:16px;text-align:center;color:#94a3b8">لا توجد تعليمات / No instructions</div>':'')
    +'<div style="text-align:center">'
    +'<button onclick="startPreviewTest(\''+m+'\')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:none;border-radius:20px;padding:14px 36px;font-size:15px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;box-shadow:0 4px 14px rgba(34,197,94,.4)">▶ ابدأ الاختبار / Start Test</button>'
    +'</div></div></div>';
  cont.innerHTML=instrHtml;
  selectDisplayModeStep4(m);
}
function startPreviewTest(m){
  var cont=document.getElementById('previewBoxContent');
  if(!cont) return;
  var d=testData.domains[0];
  if(!d) return;
  var qs=d.questions||[];
  if(d.hasBranches&&d.branches&&d.branches[0]) qs=d.branches[0].questions||[];
  var mm=parseInt(m)||1;
  if(mm===4){
    cont.innerHTML='';cont.style.background='#c8c8c0';cont.style.padding='16px';cont.style.overflowY='auto';
    cont.innerHTML=_buildWPPreview(d,qs);
  } else if(mm===3){
    cont.innerHTML='';cont.style.background='#f5f5f0';cont.style.overflowY='auto';cont.style.padding='16px';
    cont.innerHTML=_buildStreamPreview(d,qs);
  } else {
    cont.innerHTML='';cont.style.background='linear-gradient(135deg,#1e3a8a,#1e1b4b)';cont.style.overflowY='auto';cont.style.padding='0';
    cont.innerHTML=_buildClassicPreview(d,qs,mm);
  }
}
function _buildClassicPreview(d,qs,m){
  var q=qs[0];if(!q) return '';
  var labels=['A','B','C','D','E'];
  var html='<div style="padding:16px;font-family:Tajawal,sans-serif">';
  html+='<div style="background:rgba(255,255,255,.08);border-radius:10px;padding:12px;margin-bottom:12px;display:flex;align-items:center;gap:10px">';
  html+='<div style="flex:1;font-size:14px;font-weight:700;color:white" dir="auto">'+(q.stemHtml||q.stemText||'رأس السؤال')+'</div>';
  html+='<div style="background:#FACC15;color:#1e3a8a;border-radius:8px;padding:4px 10px;font-size:11px;font-weight:800;white-space:nowrap">Q 1/'+qs.length+'</div></div>';
  if(q.mediaHtml&&!(q.mediaVisible&&q.mediaVisible.img===false)) html+='<div style="margin-bottom:10px;border-radius:8px;overflow:hidden">'+q.mediaHtml+'</div>';
  if(q.type==='mcq'){
    (q.options||[]).forEach(function(o,i){html+='<div style="background:rgba(255,255,255,.08);border:1.5px solid rgba(255,255,255,.15);border-radius:10px;padding:10px 14px;margin-bottom:8px;display:flex;align-items:center;gap:10px;cursor:pointer"><div style="width:28px;height:28px;border-radius:50%;background:rgba(250,204,21,.2);border:2px solid #FACC15;display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:800;color:#FACC15;flex-shrink:0">'+(labels[i]||i+1)+'</div><div style="color:white;font-size:13px" dir="auto">'+(o||'')+'</div></div>';});
  } else if(q.type==='truefalse'){
    (q.statements||[]).forEach(function(s,si){html+='<div style="background:rgba(255,255,255,.06);border-radius:10px;padding:10px;margin-bottom:8px"><div style="color:white;font-size:13px;margin-bottom:8px" dir="auto">'+(si+1)+'. '+(s.text||'')+'</div><div style="display:flex;gap:8px"><button style="flex:1;padding:8px;border-radius:8px;border:none;background:rgba(34,197,94,.2);color:#4ade80;font-weight:700;font-size:12px">✅ صواب</button><button style="flex:1;padding:8px;border-radius:8px;border:none;background:rgba(239,68,68,.2);color:#f87171;font-weight:700;font-size:12px">❌ خطأ</button></div></div>';});
  } else {
    html+='<textarea style="width:100%;min-height:100px;background:rgba(255,255,255,.08);border:1.5px solid rgba(255,255,255,.15);border-radius:10px;padding:10px;color:white;font-family:Tajawal,sans-serif;font-size:14px;resize:none" placeholder="اكتب إجابتك هنا..." dir="auto"></textarea>';
  }
  html+='<div style="display:flex;justify-content:space-between;margin-top:12px"><button style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 18px;font-size:13px;font-family:Tajawal,sans-serif;cursor:default">← السابق</button><button style="background:#FACC15;color:#1e3a8a;border:none;border-radius:10px;padding:8px 18px;font-size:13px;font-weight:800;font-family:Tajawal,sans-serif;cursor:default">التالي →</button></div>';
  html+='</div>';return html;
}
function _buildStreamPreview(d,qs){
  var labels=['A','B','C','D'];
  var html='';
  qs.slice(0,3).forEach(function(q,qi){
    if(qi>0) html+='<div style="height:3px;background:linear-gradient(90deg,transparent,#FACC15,transparent);margin:0 0 0 0;box-shadow:0 1px 6px rgba(245,158,11,.4)"></div>';
    html+='<div style="background:white;padding:20px 22px;font-family:Tajawal,sans-serif">';
    html+='<div style="display:flex;gap:10px;margin-bottom:12px"><div style="min-width:36px;height:36px;background:linear-gradient(135deg,#1e3a8a,#7e22ce);border-radius:10px;display:flex;align-items:center;justify-content:center;color:white;font-weight:800;font-size:12px;font-family:Montserrat,sans-serif">Q.'+(qi+1)+'</div><div style="font-size:14px;font-weight:600;line-height:1.7;color:#1a1a2e" dir="auto">'+(q.stemHtml||q.stemText||'السؤال')+'</div></div>';
    if(q.type==='mcq'){html+='<div>';(q.options||[]).forEach(function(o,i){html+='<div style="display:flex;align-items:center;gap:8px;padding:8px 10px;border:1.5px solid #e2e8f0;border-radius:10px;margin-bottom:6px;cursor:pointer"><div style="width:24px;height:24px;border-radius:50%;border:2px solid #1e3a8a;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#1e3a8a;flex-shrink:0">'+(labels[i]||i+1)+'</div><span style="font-size:13px;color:#1a1a2e" dir="auto">'+(o||'')+'</span></div>';});html+='</div>';}
    else{html+='<textarea style="width:100%;min-height:80px;border:1.5px solid #e2e8f0;border-radius:10px;padding:10px;font-family:Tajawal,sans-serif;font-size:14px;color:#1a1a2e;resize:none" dir="auto" placeholder="اكتب إجابتك..."></textarea>';}
    html+='</div>';
  });
  if(qs.length>3) html+='<div style="background:white;padding:14px;text-align:center;color:#94a3b8;font-size:12px;font-family:Montserrat,sans-serif">+ '+(qs.length-3)+' more questions...</div>';
  return html;
}
function _buildWPPreview(d,qs){
  var labels=['A','B','C','D'];
  var html='<div style="background:white;max-width:680px;margin:0 auto;padding:28px 30px;box-shadow:0 4px 20px rgba(0,0,0,.15);font-family:Tajawal,\'Times New Roman\',serif">';
  html+='<div style="border-bottom:3px double #111;padding-bottom:10px;margin-bottom:14px;text-align:center"><div style="font-size:16px;font-weight:900">'+(testData.testName||'اختبار')+'</div></div>';
  html+='<div style="font-size:11px;border:1px solid #ccc;padding:8px;margin-bottom:14px;display:grid;grid-template-columns:1fr 1fr;gap:4px"><span><b>الطالب:</b> _______________</span><span><b>الصف:</b> '+(testData.grade||'___')+'</span></div>';
  qs.slice(0,3).forEach(function(q,qi){
    html+='<div style="margin-bottom:18px"><div style="font-size:14px;font-weight:700;display:flex;gap:6px;margin-bottom:8px"><span style="color:#1e3a8a;font-weight:900">'+(qi+1)+'.</span><span dir="auto">'+(q.stemHtml||q.stemText||'السؤال')+'</span></div>';
    if(q.type==='mcq'){html+='<div style="display:grid;grid-template-columns:1fr 1fr;gap:5px;margin-right:20px">';(q.options||[]).forEach(function(o,i){html+='<div style="display:flex;align-items:center;gap:5px;font-size:12px;border:1px solid #e2e8f0;border-radius:4px;padding:4px 8px"><div style="width:14px;height:14px;border-radius:50%;border:2px solid #94a3b8;flex-shrink:0"></div><span dir="auto">'+(o||'')+'</span></div>';});html+='</div>';}
    else{html+='<div style="margin-right:20px"><div style="border-bottom:1px solid #ccc;height:26px;margin-bottom:5px"></div><div style="border-bottom:1px solid #ccc;height:26px"></div></div>';}
    html+='</div>';
  });
  if(qs.length>3) html+='<div style="text-align:center;color:#94a3b8;font-size:11px">+ '+(qs.length-3)+' more...</div>';
  html+='</div>';return html;
}
function resetPreviewBox(){
  _previewBoxMode=0;_previewBoxData=null;
  [1,2,3,4].forEach(function(i){var c=document.getElementById('modeCard'+i);if(c){c.classList.remove('mode-active');c.style.borderColor='rgba(255,255,255,.12)';}});
  var el=document.getElementById('previewBoxModeName');if(el) el.textContent='— اختر نمطاً من الأزرار أدناه —';
  var cont=document.getElementById('previewBoxContent');
  if(cont) cont.innerHTML='<div style="text-align:center;color:rgba(255,255,255,.3);padding:60px 20px"><div style="font-size:48px;margin-bottom:12px">🎯</div><div style="font-size:14px;font-family:Tajawal,sans-serif;font-weight:700">اضغط على أحد الأنماط أدناه لمعاينة الاختبار</div></div>';
  var bar=document.getElementById('selectedModeBar');if(bar) bar.style.display='none';
  var btn=document.getElementById('approveBtn');if(btn){btn.disabled=true;btn.style.opacity='.4';btn.style.cursor='not-allowed';}
}
function selectDisplayModeStep1(m){
  testData.displayMode=m;
  [1,2,3,4].forEach(function(i){
    var card=document.getElementById('dm'+i);
    var check=document.getElementById('dm'+i+'-check');
    if(card) card.style.borderColor=i===m?'#22c55e':'rgba(255,255,255,.12)';
    if(check) check.style.display=i===m?'block':'none';
  });
}

// ══ كود المدرسة التخليقي ══
var _curriculumCodeMap={IGCSE:'IG',American:'AM',IB:'IB',Canadian:'CA',CBSE:'CB',Local:'LO'};
function _extractCountryCode(countryStr){
  if(!countryStr) return 'XX';
  var parts=countryStr.split('/');
  var en=(parts[1]||parts[0]||'').trim();
  var upperOnly=en.replace(/[^A-Z]/g,'');
  if(upperOnly.length>=2) return upperOnly.slice(0,4);
  return en.replace(/\s+/g,'').substring(0,3).toUpperCase()||'XX';
}
function generateSchoolCode(schoolName,country,curriculum){
  var cc=_extractCountryCode(country);
  var curCode=_curriculumCodeMap[curriculum]||(curriculum?curriculum.replace(/\s+/g,'').substring(0,2).toUpperCase():'GN');
  var sameCount=(schools||[]).filter(function(s){return s.country===country&&s.curriculum===curriculum;}).length;
  var seq=String(sameCount+1).padStart(2,'0');
  return 'S'+cc+'_'+curCode+seq;
}
function previewNewSchoolCode(){
  var nameEl=document.getElementById('newSchoolName');
  var countryEl=document.getElementById('newSchoolCountry');
  var curEl=document.getElementById('newSchoolCurriculum');
  var out=document.getElementById('newSchoolCodePreview');
  if(!out) return;
  var country=countryEl?countryEl.value:'';
  var curriculum=curEl?curEl.value:'';
  if(!country||!curriculum){ out.textContent='—'; return; }
  out.textContent=generateSchoolCode(nameEl?nameEl.value:'',country,curriculum);
}
function addSchoolWithCode(schoolData){
  if(!schoolData.code) schoolData.code=generateSchoolCode(schoolData.name,schoolData.country,schoolData.curriculum);
  return schoolData;
}

function selectDisplayModeStep4(m){selectDisplayModeStep1(m);}
function selectDisplayMode(m){selectDisplayModeStep1(m);}
function tryMode(m){
  if(!testData.domains||!testData.domains.length){scWarn('أضف مجالات وأسئلة أولاً','Add domains first');return;}
  var hasQ=testData.domains.some(function(d){return (d.questions&&d.questions.length)||(d.branches&&d.branches.some(function(b){return b.questions&&b.questions.length;}));});
  if(!hasQ){scWarn('أضف أسئلة للمجالات أولاً','Add questions first');return;}
  _savedTestDataForPreview=JSON.parse(JSON.stringify(testData));
  testData.displayMode=m;
  sw_domainIdx=0;sw_branchIdx=-1;sw_qIdx=0;sw_answers={};
  _tryModeActive=true;_previewReturnStep=4;
  applyDisplayMode();
  document.getElementById('studentWindow').style.display='flex';
  var tnEl=document.getElementById('sw-test-name-header');if(tnEl)tnEl.textContent=testData.testName||'معاينة';
  var tmEl=document.getElementById('sw-test-meta-header');
  if(tmEl){var mt=[];if(testData.subject)mt.push(testData.subject);if(testData.grade)mt.push(testData.grade);tmEl.textContent=mt.join(' — ');}
  var cb=document.getElementById('sw-close-btn');if(cb)cb.textContent='← رجوع للأنماط';
  sw_instructionsConfirmed=false;
  showStudentInstructions();
}
var _tryModeActive=false;
function applyDisplayMode(){
  var sw=document.getElementById('studentWindow');
  var m=testData&&testData.displayMode?testData.displayMode:1;
  sw.classList.remove('dm-mode-1','dm-mode-2','dm-mode-3','dm-mode-4');
  sw.classList.add('dm-mode-'+m);
  var cb=document.getElementById('sw-close-btn');
  if(cb&&_tryModeActive) cb.textContent='← رجوع للأنماط';
}
var currentBranchIndex=-1; // -1 = domain questions, >=0 = branch questions

function _getCurrentQuestions(){
  var d=testData.domains[currentDomainIndex];
  if(currentBranchIndex>=0&&d.branches&&d.branches[currentBranchIndex]) return d.branches[currentBranchIndex].questions=d.branches[currentBranchIndex].questions||[];
  return d.questions=d.questions||[];
}

function openBranchQuestions(branchIdx){
  var d=testData.domains[currentDomainIndex];
  var count=parseInt(document.getElementById('branchCount').value)||0;
  var existingBranches=d.branches||[];
  d.branches=[];
  for(var j=0;j<count;j++){
    var existing=existingBranches[j]||{};
    d.branches.push({
      nameAr:document.getElementById('brName'+j)?document.getElementById('brName'+j).value||('فرع '+(j+1)):'فرع '+(j+1),
      nameEn:document.getElementById('brNameEn'+j)?document.getElementById('brNameEn'+j).value||('Branch '+(j+1)):'Branch '+(j+1),
      weight:Number(document.getElementById('brW'+j)?document.getElementById('brW'+j).value||0:0),
      time:Number(document.getElementById('brT'+j)?document.getElementById('brT'+j).value||0:0),
      qCount:Number(document.getElementById('brQ'+j)?document.getElementById('brQ'+j).value||0:0),
      questions:existing.questions||[]
    });
  }
  currentBranchIndex=branchIdx;
  var br=d.branches[branchIdx];
  // Show domainSettingsBox as questions panel for this branch
  var settingsBox=document.getElementById('domainSettingsBox');
  settingsBox.classList.remove('hidden');
  document.getElementById('modalDomainWeight').value=(br.weight||0)+'%';
  document.getElementById('modalQCount').value=br.qCount||'';
  // Inject back-button header above questionsList
  var oldHeader=document.getElementById('branchQHeader');
  if(oldHeader) oldHeader.remove();
  var header=document.createElement('div');
  header.id='branchQHeader';
  header.style.cssText='background:rgba(59,130,246,.15);border:2px solid rgba(59,130,246,.4);border-radius:14px;padding:12px 16px;margin-bottom:16px;display:flex;align-items:center;justify-content:space-between;gap:12px';
  header.innerHTML='<div><div style="font-size:13px;font-weight:800;color:#93c5fd">📝 أسئلة الفرع / Branch Questions</div><div style="font-size:15px;font-weight:900;color:white;margin-top:2px">'+(br.nameAr||'فرع '+(branchIdx+1))+(br.nameEn?' / '+br.nameEn:'')+'</div></div>'
    +'<button type="button" onclick="closeBranchQuestions()" style="background:rgba(255,255,255,.1);border:2px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 16px;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700;cursor:pointer">← رجوع للفروع</button>';
  var qList=document.getElementById('questionsList');
  qList.parentNode.insertBefore(header,qList);
  // Switch buttons - show Save Domain (it saves everything including branch)
  var sbq=document.getElementById('saveBranchQBtn');
  var sdb=document.getElementById('saveDomainBtn');
  if(sbq){sbq.style.cssText='display:none!important';}
  if(sdb){sdb.style.removeProperty('display');sdb.style.display='inline-flex';sdb.style.visibility='visible';}
  document.getElementById('branchSetupBox').classList.add('hidden');
  document.getElementById('branchQuestionBox').classList.add('hidden');
  renderQuestionsList();
}
function closeBranchQuestions(){
  currentBranchIndex=-1;
  var oldHeader=document.getElementById('branchQHeader');
  if(oldHeader) oldHeader.remove();
  document.getElementById('domainSettingsBox').classList.add('hidden');
  document.getElementById('branchSetupBox').classList.remove('hidden');
  document.getElementById('branchQuestionBox').classList.remove('hidden');
  // Restore buttons
  var sbq2=document.getElementById('saveBranchQBtn');
  var sdb2=document.getElementById('saveDomainBtn');
  if(sbq2){sbq2.style.display='none';}
  if(sdb2){sdb2.style.display='inline-flex';sdb2.style.visibility='visible';}
  var d=testData.domains[currentDomainIndex];
  if(d&&d.branches) generateBranchInputs(d.branches);
}
function saveBranchQuestions(){
  var d=testData.domains[currentDomainIndex];
  var br=d.branches&&d.branches[currentBranchIndex]?d.branches[currentBranchIndex]:null;
  if(!br) return;
  var qs=_getCurrentQuestions();
  if(qs.length===0){
    scWarn('أضف أسئلة للفرع أولاً','Add questions to branch first');
    return;
  }
  d.branches[currentBranchIndex].questions=qs.slice();
  _saveDraft();
  closeBranchQuestions();
  scOk('تم حفظ الفرع ✅','Branch Saved',
    'تم حفظ أسئلة الفرع بنجاح / Branch questions saved successfully','✅');
}
var selectedSchools=[], currentSupName='', currentStudentData=null;
var sw_domainIdx=0, sw_qIdx=0, sw_currentPage='instructions', sw_answers={}, sw_timerInterval=null, sw_timeLeft=0, sw_instructionsConfirmed=false;
var sw_branchIdx=-1; // -1 = not in branch mode, >=0 = current branch index
var sw_branchAnswers={}; // {domainIdx: {branchIdx: {qIdx: answer}}}
var sw_completedBranches={}; // {domainIdx: [branchIdx,...]}
var sw_completedDomains=[]; // [domainIdx,...]
var cheatWarnings=0, cheatLog=[];
var imgStates={}, newSchoolLogoData='';
var pendingStudentCodes=[], activeCodesTestId=null;
var currentDomainHasBranches=false;
var matchConnections=[], matchSelected=null;
var smSelected=null, smConnections={};
var orderingCorrect=[], orderingCurrent=[];
var dragSrcIdx=null, touchSrcEl=null, touchClone=null;
var canvasCtx=null, isDrawingCanvas=false, drawTool='pen', brushSize=5, penColor='#1a1a2e';
var recordingInterval=null, mediaRecorder=null, recordingChunks=[];
var _imgS={sc:1,rot:0,x:0,y:0,drag:false,sx:0,sy:0};
var _updateTargetId='';
var sb_currentGrade=null, sb_currentSubject=null;

// ============================================================
// COUNTRIES & YEARS
// ============================================================
var worldCountries=[
  {n:'الإمارات / UAE',f:'🇦🇪'},{n:'المملكة العربية السعودية / Saudi Arabia',f:'🇸🇦'},
  {n:'قطر / Qatar',f:'🇶🇦'},{n:'الكويت / Kuwait',f:'🇰🇼'},{n:'البحرين / Bahrain',f:'🇧🇭'},
  {n:'عُمان / Oman',f:'🇴🇲'},{n:'مصر / Egypt',f:'🇪🇬'},{n:'الأردن / Jordan',f:'🇯🇴'},
  {n:'لبنان / Lebanon',f:'🇱🇧'},{n:'المغرب / Morocco',f:'🇲🇦'},{n:'تونس / Tunisia',f:'🇹🇳'},
  {n:'فلسطين / Palestine',f:'🇵🇸'},{n:'العراق / Iraq',f:'🇮🇶'},{n:'اليمن / Yemen',f:'🇾🇪'},
  {n:'المملكة المتحدة / UK',f:'🇬🇧'},{n:'الولايات المتحدة / USA',f:'🇺🇸'},
  {n:'كندا / Canada',f:'🇨🇦'},{n:'أستراليا / Australia',f:'🇦🇺'},
  {n:'الهند / India',f:'🇮🇳'},{n:'باكستان / Pakistan',f:'🇵🇰'},
  {n:'فرنسا / France',f:'🇫🇷'},{n:'تركيا / Turkey',f:'🇹🇷'},
  {n:'سنغافورة / Singapore',f:'🇸🇬'},{n:'اليابان / Japan',f:'🇯🇵'}
];

function populateCountrySelects(){
  ['country','localCountry','newSchoolCountry'].forEach(function(id){
    var el=document.getElementById(id); if(!el) return;
    var prev=el.value;
    el.innerHTML='<option value="">-- اختر الدولة / Country --</option>';
    worldCountries.forEach(function(c){var o=document.createElement('option');o.value=c.n;o.textContent=c.f+' '+c.n;el.appendChild(o);});
    if(prev) el.value=prev;
  });
}

function populateAcademicYears(){
  var sel=document.getElementById('academicYear'); if(!sel) return;
  sel.innerHTML='<option value="">-- اختر السنة --</option>';
  for(var y=2020;y<=2199;y++){var o=document.createElement('option');o.value=y+'-'+(y+1);o.textContent=y+' - '+(y+1);if(y===2025)o.selected=true;sel.appendChild(o);}
}

function checkLocalCurriculum(){
  document.getElementById('localCurriculumRow').classList.toggle('hidden',document.getElementById('curriculum').value!=='Local');
}
function handleSubjectChange(){
  var v=document.getElementById('subject').value;
  document.getElementById('historyNationalRow').classList.toggle('hidden',v!=='history_national');
  document.getElementById('subjectOtherRow').classList.toggle('hidden',v!=='other');
}

// ============================================================
// FILTER SCHOOLS
// ============================================================
function filterSchoolsByCountryAndCurriculum(){
  var cv=document.getElementById('country')?document.getElementById('country').value:'';
  var cur=document.getElementById('curriculum')?document.getElementById('curriculum').value:'';
  var container=document.getElementById('schoolsMultiList'); if(!container) return;
  var filtered=schools.slice();
  if(cv) filtered=filtered.filter(function(s){return s.country===cv;});
  if(cur) filtered=filtered.filter(function(s){return s.curriculum===cur;});
  if(!filtered.length){container.innerHTML='<div style="color:rgba(255,255,255,.3);font-size:13px;padding:8px;text-align:center">لا توجد مدارس مطابقة / No matching schools</div>';return;}
  container.innerHTML=filtered.map(function(s){return '<label class="school-option'+(selectedSchools.indexOf(s.id)>=0?' selected':'')+'" id="sopt_'+s.id+'"><input type="checkbox" value="'+s.id+'" onchange="toggleSchool('+s.id+',\''+s.name+'\')" '+(selectedSchools.indexOf(s.id)>=0?'checked':'')+'>'+(s.logoSrc?'<img src="'+s.logoSrc+'" style="width:22px;height:22px;object-fit:contain;border-radius:4px;background:white">':'🏫')+' '+s.name+'</label>';}).join('');
}
function toggleSchool(id,name){
  var idx=selectedSchools.indexOf(id);
  if(idx>=0) selectedSchools.splice(idx,1); else selectedSchools.push(id);
  var el=document.getElementById('sopt_'+id); if(el) el.classList.toggle('selected',selectedSchools.indexOf(id)>=0);
  renderSchoolTags();
}
function renderSchoolTags(){
  document.getElementById('selectedSchoolsTags').innerHTML=selectedSchools.map(function(id){
    var s=schools.find(function(x){return x.id===id;});
    return s?'<div class="school-tag">'+(s.logoSrc?'<img src="'+s.logoSrc+'" style="width:16px;height:16px;object-fit:contain;border-radius:2px;background:white">':'🏫')+' '+s.name+'<span class="remove-tag" onclick="toggleSchool('+s.id+',\''+s.name+'\')">×</span></div>':'';
  }).join('');
}

// ============================================================
// IMAGE CONTROLS
// ============================================================
function addLogoFromURL(previewId,stateKey){
  scPromptText('رابط الصورة','Image URL','https://example.com/logo.png','🌐').then(function(url){
    if(!url||!url.trim()) return;
    var img=document.getElementById(previewId); if(!img) return;
    img.src=url.trim(); img.style.display='block';
    var ph=document.getElementById(previewId.replace('Preview','Placeholder'));
    if(ph) ph.style.display='none';
    var ctrl=document.getElementById(stateKey==='wizardLogoState'?'wizardLogoControls':'newSchoolImgControls');
    _finishLogoFromURL(stateKey,ctrl);
  });
}
function _finishLogoFromURL(stateKey,ctrl){
  if(ctrl) ctrl.style.display='flex';
  if(!imgStates[stateKey]) imgStates[stateKey]={sc:1,rot:0,x:0,y:0};
}
function previewNewSchoolLogo(e){
  var file=e.target.files[0]; if(!file) return;
  var reader=new FileReader();
  reader.onload=function(ev){
    newSchoolLogoData=ev.target.result;
    var img=document.getElementById('newSchoolLogoPreview');
    img.src=ev.target.result; img.style.display='block';
    document.getElementById('newSchoolLogoPlaceholder').style.display='none';
    document.getElementById('newSchoolImgControls').style.display='flex';
    imgStates['newSchoolImgState']={sc:1,rot:0,x:0,y:0};
  };
  reader.readAsDataURL(file);
}
function previewLogo(e){
  var file=e.target.files[0]; if(!file) return;
  var reader=new FileReader();
  reader.onload=function(ev){
    document.getElementById('logoPreview').src=ev.target.result;
    document.getElementById('logoPreview').style.display='block';
    document.getElementById('logoPlaceholder').style.display='none';
    document.getElementById('wizardLogoControls').style.display='flex';
    testData.logoSrc=ev.target.result;
    imgStates['wizardLogoState']={sc:1,rot:0,x:0,y:0};
  };
  reader.readAsDataURL(file);
}
function imgCtrlFor(imgId,stateKey,action){
  if(!imgStates[stateKey]) imgStates[stateKey]={sc:1,rot:0,x:0,y:0};
  var s=imgStates[stateKey];
  if(action==='zi') s.sc=Math.min(4,parseFloat((s.sc+.2).toFixed(2)));
  else if(action==='zo') s.sc=Math.max(.2,parseFloat((s.sc-.2).toFixed(2)));
  else if(action==='rr') s.rot=(s.rot+90)%360;
  else if(action==='rl') s.rot=(s.rot-90+360)%360;
  else if(action==='rs'){s.sc=1;s.rot=0;s.x=0;s.y=0;}
  var img=document.getElementById(imgId);
  if(img) img.style.transform='translate('+s.x+'px,'+s.y+'px) scale('+s.sc+') rotate('+s.rot+'deg)';
}
function confirmLogoFinal(imgId,controlsId){
  scConfirm('تثبيت الشعار','Confirm Logo','هل هذا هو الشكل النهائي للشعار؟','Is this the final logo?','🖼️').then(function(ok){
    if(!ok)return;
    var img=document.getElementById(imgId);
    if(imgId==='logoPreview'&&img) testData.logoSrc=img.src;
    if(imgId==='newSchoolLogoPreview'&&img) newSchoolLogoData=img.src;
    var ctrl=document.getElementById(controlsId);
    if(ctrl) ctrl.style.display='none';
    scOk('تم التثبيت ✅','Confirmed','تم تثبيت الشعار بنجاح','Logo confirmed successfully','✅');
  });
}

// ============================================================
// SCHOOL MANAGER
// ============================================================
function addSchool(){
  var name=document.getElementById('newSchoolName').value.trim();
  var country=document.getElementById('newSchoolCountry').value;
  var curriculum=document.getElementById('newSchoolCurriculum').value;
  if(!name){alert('يرجى كتابة اسم المدرسة / Enter school name');return;}
  if(schools.find(function(s){return s.name===name;})){alert('المدرسة موجودة / Already exists!');return;}
  var slug=name.replace(/\s+/g,'').substring(0,20);
  var username=slug+'@ist', password='123456';
  var schoolCode=generateSchoolCode(name,country,curriculum);
  schools.push({id:Date.now(),name:name,country:country,curriculum:curriculum,logoSrc:newSchoolLogoData,username:username,password:password,code:schoolCode,added:new Date().toLocaleDateString('ar')});
  saveSchoolsGH().then(function(){
    document.getElementById('newSchoolName').value='';
    document.getElementById('newSchoolCountry').value='';
    document.getElementById('newSchoolCurriculum').value='';
    var cpEl=document.getElementById('newSchoolCodePreview');if(cpEl)cpEl.textContent='—';
    document.getElementById('newSchoolLogoPreview').style.display='none';
    document.getElementById('newSchoolLogoPlaceholder').style.display='flex';
    document.getElementById('newSchoolImgControls').style.display='none';
    newSchoolLogoData='';
    renderSchools();
    alert('✅ تم إنشاء حساب المدرسة:\nUsername: '+username+'\nPassword: '+password);
  });
}
function getCurriculumLabel(code){
  switch(code){
    case 'IGCSE': return '🇬🇧 البريطاني / IGCSE';
    case 'American': return '🇺🇸 الأمريكي / American';
    case 'IB': return '🌐 البكالوريا الدولية / IB';
    case 'Canadian': return '🇨🇦 الكندي / Canadian';
    case 'CBSE': return '🇮🇳 CBSE / الهندي';
    case 'Local': return '📍 المحلي / Local';
    default: return code||'—';
  }
}
function renderSchools(){
  var list=document.getElementById('schoolList'); if(!list) return;
  if(!schools.length){list.innerHTML='<tr><td colspan="10" style="color:rgba(255,255,255,.3);padding:24px">لا توجد مدارس / No schools</td></tr>';return;}
  list.innerHTML=schools.map(function(s,i){return '<tr><td class="font-en">'+(i+1)+'</td><td>'+(s.logoSrc?'<img src="'+s.logoSrc+'" style="width:32px;height:32px;object-fit:contain;border-radius:6px;margin:auto;background:white;padding:2px">':'—')+'</td><td>🏫 '+s.name+'</td><td style="font-size:11px">'+(s.country||'—')+'</td><td style="font-size:11px;font-family:Montserrat,sans-serif">'+getCurriculumLabel(s.curriculum)+'</td><td style="font-family:Montserrat,sans-serif;color:#FACC15;font-size:11px;font-weight:700">'+(s.code||'—')+'</td><td style="font-family:Montserrat,sans-serif;color:#93c5fd;font-size:11px">'+(s.username||'—')+'</td><td style="font-family:Montserrat,sans-serif;font-size:11px">'+(s.password||'—')+'</td><td style="font-size:11px">'+(s.added||'—')+'</td><td><button onclick="deleteSchool('+s.id+')" style="color:#f87171;background:none;border:none;cursor:pointer;font-size:14px">🗑</button></td></tr>';}).join('');
}
function deleteSchool(id){scConfirm('حذف المدرسة','Delete School','هل تريد حذف هذه المدرسة؟','Delete this school?','🗑').then(function(ok){if(ok){schools=schools.filter(function(s){return s.id!==id;});saveSchoolsGH().then(renderSchools);}});}
function showSchoolManager(){document.getElementById('adminOptions').classList.add('hidden');document.getElementById('schoolManager').classList.remove('hidden');populateCountrySelects();renderSchools();}
function showSupManager(){document.getElementById('adminOptions').classList.add('hidden');document.getElementById('supManager').classList.remove('hidden');renderSups();}
function showGeneralReviewer(){document.getElementById('adminOptions').classList.add('hidden');document.getElementById('generalReviewerPanel').classList.remove('hidden');renderGeneralReviewerContent();}
function hideSubPanel(id){
  document.getElementById(id).classList.add('hidden');
  document.getElementById('adminOptions').classList.remove('hidden');
  // Also close grading committee if admin
  if(id === 'gradingCommitteePanel'){
    document.getElementById('adminOptions').classList.remove('hidden');
  }
}

// ============================================================
// GENERAL REVIEWER
// ============================================================
function renderGeneralReviewerContent(){
  var cont=document.getElementById('generalReviewerContent');
  var pending=myTests.filter(function(t){return t.status==='underReview'||t.status==='reviewDone';});
  if(!pending.length){cont.innerHTML='<div style="text-align:center;color:rgba(255,255,255,.3);padding:48px"><div style="font-size:48px;margin-bottom:12px">🔍</div><p>لا توجد اختبارات / No tests pending</p></div>';return;}
  cont.innerHTML=pending.map(function(t){
    var notesHtml='';
    if(t.returnNotes&&t.returnNotes.length){notesHtml='<div style="margin-top:10px">'+t.returnNotes.map(function(n){return '<div style="background:rgba(248,113,113,.1);border:1px solid rgba(248,113,113,.3);border-radius:10px;padding:8px 12px;margin-bottom:5px;font-size:12px"><div style="color:#fca5a5;font-weight:700;margin-bottom:3px">'+n.name+' — <span style="font-size:10px;color:rgba(255,255,255,.4)">'+n.at+'</span></div><div style="color:#fecaca">'+n.note+'</div></div>';}).join('')+'</div>';}
    return '<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:18px;padding:20px;margin-bottom:16px"><div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px"><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;flex:1"><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المدرسة</div><div style="font-weight:700;font-size:13px">'+(t.school||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المادة</div><div style="font-weight:700;font-size:13px">'+(t.subject||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المحرر</div><div style="font-weight:700;font-size:13px">'+(t.author||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المراجع المختار</div><div style="font-weight:700;font-size:13px">'+(t.reviewer1||'—')+' <span class="status-badge '+(t.r1done?'status-reviewed':'status-pending')+'" style="font-size:10px">'+(t.r1done?'✅ راجع':'⏳ انتظار')+'</span></div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">الحالة</div><div><span class="status-badge '+(t.generalStatus==='approved'?'status-approved':t.generalStatus==='returned'?'status-returned':'status-pending')+'">'+(t.generalStatus==='approved'?'✅ معتمد':t.generalStatus==='returned'?'⚠️ مُرجع':'⏳ انتظار')+'</span></div></div></div><div style="display:flex;flex-direction:column;gap:8px"><button onclick="grPreviewTest('+t.id+')" style="background:rgba(96,165,250,.2);color:#93c5fd;border:none;border-radius:10px;padding:7px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">👁 معاينة</button><button onclick="grApproveTest('+t.id+')" style="background:rgba(74,222,128,.2);color:#4ade80;border:none;border-radius:10px;padding:7px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">✅ اعتماد</button><button onclick="grReturnTest('+t.id+')" style="background:rgba(248,113,113,.2);color:#f87171;border:none;border-radius:10px;padding:7px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">⚠️ إرجاع</button><button onclick="grSendToArchive('+t.id+')" style="background:rgba(250,204,21,.15);color:#FACC15;border:1px solid rgba(250,204,21,.3);border-radius:10px;padding:7px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">🗄️ أرشيف</button><button onclick="grDeleteTest('+t.id+')" style="background:rgba(239,68,68,.15);color:#f87171;border:1px solid rgba(239,68,68,.3);border-radius:10px;padding:7px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">🗑 حذف</button></div></div>'+notesHtml+'</div>';
  }).join('');
}
function grPreviewTest(id){
  var t=myTests.find(function(x){return x.id===id;});
  if(!t||!t.domains||!t.domains.length){scWarn('لا توجد بيانات لمعاينتها','No data to preview');return;}
  _savedTestDataForPreview=JSON.parse(JSON.stringify(testData));
  testData={
    domains:JSON.parse(JSON.stringify(t.domains)),
    selectedSchools:[{name:t.school||'معاينة',logo:''}],
    logoSrc:t.logoSrc||'',
    instructionsAr:t.instructionsAr||'',
    instructionsEn:t.instructionsEn||'',
    displayMode:t.displayMode||1,
    testName:t.testName||t.name||'معاينة',
    subject:t.subject||'',
    grade:t.grade||'',
    term:t.term||'',
    year:t.year||''
  };
  sw_domainIdx=0;sw_branchIdx=-1;sw_qIdx=0;sw_answers={};
  _tryModeActive=true;_previewReturnStep=-1; // -1 = لا يرجع لخطوة wizard (لوحة المراجع)
  applyDisplayMode();
  document.getElementById('studentWindow').style.display='flex';
  var tnEl=document.getElementById('sw-test-name-header');if(tnEl)tnEl.textContent=testData.testName||'معاينة';
  var tmEl=document.getElementById('sw-test-meta-header');
  if(tmEl){var mt=[];if(testData.subject)mt.push(testData.subject);if(testData.grade)mt.push(testData.grade);tmEl.textContent=mt.join(' — ');}
  var cb=document.getElementById('sw-close-btn');if(cb)cb.textContent='✕ إغلاق المعاينة';
  sw_instructionsConfirmed=false;
  showStudentInstructions();
}
function grApproveTest(id){
  var t=myTests.find(function(x){return x.id===id;});
  if(!t)return;
  scConfirm('اعتماد الاختبار','Approve Test','هل تريد اعتماد هذا الاختبار؟<br><b>'+(t.testName||t.name||'')+'</b>','Approve this test?','✅').then(function(ok){
    if(!ok)return;
    t.generalStatus='approved';t.status='reviewDone';t.approvedAt=new Date().toLocaleDateString('ar');
    saveMyTests();renderGeneralReviewerContent();
    scOk('تم الاعتماد ✅','Approved','تم اعتماد الاختبار من المراجع العام بنجاح','Test approved successfully','✅');
  });
}
function grReturnTest(id){
  var t=myTests.find(function(x){return x.id===id;});
  if(!t)return;
  scPromptText('ملاحظات الإرجاع','Return Notes','اكتب سبب إرجاع الاختبار / Write the reason for returning the test','⚠️').then(function(reason){
    if(!reason)return;
    t.generalStatus='returned';t.status='underReview';
    if(!t.returnNotes)t.returnNotes=[];
    t.returnNotes.push({from:'general',name:'المراجع العام',note:reason,at:new Date().toLocaleString('ar')});
    saveMyTests();renderGeneralReviewerContent();
    scOk('تم الإرجاع ⚠️','Returned','تم إرجاع الاختبار مع الملاحظة','Test returned with notes','⚠️');
  });
}
function grSendToArchive(id){
  var t=myTests.find(function(x){return x.id===id;});
  if(!t)return;
  scConfirm('إرسال للأرشيف','Send to Archive','إرسال هذا الاختبار للأرشيف؟<br><b>'+(t.testName||t.name||'')+'</b>','Send this test to archive?','🗄️').then(function(ok){
    if(!ok)return;
    if(!archivedTests) archivedTests=[];
    var exists=archivedTests.some(function(a){return a.id==='arch-'+id;});
    if(!exists){
      archivedTests.unshift(Object.assign({},t,{id:'arch-'+id,archivedAt:Date.now()}));
      ghWrite('archive.json',archivedTests).catch(function(){});
    }
    myTests=myTests.filter(function(x){return x.id!==id;});
    saveMyTests();renderGeneralReviewerContent();
    scOk('تم الإرسال للأرشيف 🗄️','Archived','تم حفظ الاختبار في الأرشيف بنجاح','Test archived successfully','🗄️');
  });
}
function grDeleteTest(id){
  var t=myTests.find(function(x){return x.id===id;});
  scConfirm('حذف نهائي','Permanent Delete','⚠️ حذف هذا الاختبار نهائياً من الجميع؟<br><b>'+(t?(t.testName||t.name||''):'')+'</b>','Delete permanently for everyone?','🗑').then(function(ok){
    if(!ok)return;
    myTests=myTests.filter(function(x){return x.id!==id;});
    saveMyTests();renderGeneralReviewerContent();
    scOk('تم الحذف 🗑','Deleted','تم حذف الاختبار نهائياً','Test deleted permanently','🗑');
  });
}

// ============================================================
// SUPERVISORS
// ============================================================
var DEFAULT_SUP={id:0,name:'المشرف الرئيسي',email:'admin@scholastic.app',username:'Supervisor001',password:'Pass2050',active:true};
function createSup(){
  var name=document.getElementById('supName').value.trim(), email=document.getElementById('supEmail').value.trim();
  if(!name||!email){alert('يرجى إكمال البيانات');return;}
  var username='Super'+Math.floor(100+Math.random()*900), password=Math.random().toString(36).slice(-6);
  supervisors.push({id:Date.now(),name:name,email:email,username:username,password:password,active:true});
  saveSupervisors().then(function(){
    document.getElementById('supName').value=''; document.getElementById('supEmail').value='';
    renderSups(); alert('تم:\nUser: '+username+'\nPass: '+password);
  });
}
function renderSups(){
  var list=document.getElementById('supList');
  if(!supervisors.length){list.innerHTML='<tr><td colspan="5" style="color:rgba(255,255,255,.3);padding:24px">لا يوجد مشرفون</td></tr>';return;}
  list.innerHTML=supervisors.map(function(s){return '<tr><td>'+s.name+'</td><td class="font-en">'+s.username+'</td><td class="font-en">'+s.password+'</td><td><button onclick="toggleSup('+s.id+')" style="color:'+(s.active?'#4ade80':'#f87171')+';background:none;border:none;cursor:pointer;font-size:13px;font-weight:700">'+(s.active?'✅ نشط':'⛔ معلق')+'</button></td><td><button onclick="deleteSup('+s.id+')" style="color:#f87171;background:none;border:none;cursor:pointer">🗑</button></td></tr>';}).join('');
}
function toggleSup(id){supervisors=supervisors.map(function(s){return s.id===id?Object.assign({},s,{active:!s.active}):s;});saveSupervisors().then(renderSups);}
function deleteSup(id){scConfirm('حذف المشرف','Delete Supervisor','هل تريد حذف هذا المشرف؟','Delete this supervisor?','🗑').then(function(ok){if(ok){supervisors=supervisors.filter(function(s){return s.id!==id;});saveSupervisors().then(renderSups);}});}

// ============================================================
// AUTH
// ============================================================
var currentRole='';
function initRoute(role){
  currentRole=role;
  document.getElementById('modalTitle').innerText=
    role==='الإدارة'?'الإدارة':
    role==='المشرف'?'دخول بوابة المشرف':
    role==='الطالب'?'دخول الطالب':
    role==='المدارس'?'دخول بوابة المدارس':'دخول';
  document.getElementById('modalTitleEn').innerText=
    role==='الإدارة'?'Administration Access':
    role==='المشرف'?'Sign in to Supervisor Portal':
    role==='الطالب'?'Student Sign In':
    role==='المدارس'?'Sign in to Schools Portal':'Sign In';
  // إدارة: كلمة مرور فقط، بقية الأدوار: اسم مستخدم + كلمة مرور
  var userRow=document.getElementById('userInputRow');
  if(userRow) userRow.style.display=role==='الإدارة'?'none':'block';
  document.getElementById('userInput').value='';
  document.getElementById('passInput').value='';
  document.getElementById('loginModal').style.display='flex';
}
function closeModal(e){document.getElementById('loginModal').style.display='none';}
function toggleForgot(show){document.getElementById('loginFields').classList.toggle('hidden',show);document.getElementById('forgotFields').classList.toggle('hidden',!show);}
function sendForgotPassword(){
  var email=document.getElementById('forgotEmail').value.trim();
  if(!email||!email.includes('@')){scWarn('يرجى إدخال بريد إلكتروني صحيح','Please enter a valid email address');return;}
  var found=false;
  if(supervisors&&supervisors.length){var sup=supervisors.find(function(s){return s.email&&s.email.toLowerCase()===email.toLowerCase();});if(sup)found=true;}
  if(!found&&schools&&schools.length){var sch=schools.find(function(s){return s.email&&s.email.toLowerCase()===email.toLowerCase();});if(sch)found=true;}
  scOk(
    found?'تم إرسال البيانات 📧':'طلب الاسترداد 📧',
    found?'Recovery Email Sent':'Recovery Request Received',
    found
      ?'تم إرسال بيانات الدخول إلى بريدك الإلكتروني\nيرجى مراجعة صندوق الوارد أو البريد المزعج'
      :'إذا كان البريد الإلكتروني مسجلاً في النظام سيتم إرسال بيانات الدخول إليه\nتواصل مع مسؤول المنصة إذا لم تصلك الرسالة',
    found
      ?'Login credentials have been sent to your email\nPlease check your inbox or spam folder'
      :'If this email is registered, login credentials will be sent\nContact platform admin if you do not receive it',
    '📧'
  );
  toggleForgot(false);
}
function checkAuth(){
  var user=document.getElementById('userInput').value.trim();
  var pass=document.getElementById('passInput').value.trim();
  // قراءة كلمة مرور الإدارة المخزنة (يمكن تغييرها من صفحة معلوماتي)
  var adminPass=localStorage.getItem('scholastic_admin_pass')||'Gems@2050';
  if(currentRole==='الإدارة'){
    if(pass===adminPass){toggleSection('adminPanel');closeModal();}
    else scWarn('كلمة المرور غير صحيحة','Incorrect Password');
  } else if(currentRole==='المشرف'){
    if(user===DEFAULT_SUP.username&&pass===DEFAULT_SUP.password){currentSupName=DEFAULT_SUP.name;document.getElementById('reviewer0').value=DEFAULT_SUP.name;toggleSection('supervisorPanel');closeModal();return;}
    var sup=supervisors.find(function(s){return s.username===user&&s.password===pass;});
    if(sup){if(sup.active){currentSupName=sup.name;document.getElementById('reviewer0').value=sup.name;toggleSection('supervisorPanel');closeModal();}
    else scWarn('الحساب معلق / Account Suspended','Account Suspended');}
    else scWarn('بيانات غير صحيحة / Incorrect credentials','Incorrect credentials');
  } else if(currentRole==='الطالب'){
    var foundStudent=null, foundTest=null;
    for(var i=0;i<myTests.length;i++){var t=myTests[i];if(t.status==='approved'&&t.students){var s=t.students.find(function(st){return st.username===user&&st.password===pass&&st.active!==false;});if(s){foundStudent=s;foundTest=t;break;}}}
    if(foundStudent&&foundTest){currentStudentData={student:foundStudent,test:foundTest};showStudentPortal(foundStudent,foundTest);closeModal();}
    else scWarn('بيانات الطالب غير صحيحة / Student credentials incorrect','Incorrect student credentials');
  } else if(currentRole==='المدارس'){
    var sch=schools.find(function(s){return s.username===user&&s.password===pass;});
    if(sch){
      currentCoordSchool=sch;
      document.getElementById('coordSchoolName').textContent=sch.name;
      toggleSection('schoolCoordPanel');
      closeModal();
      initCoordPanel();
    } else scWarn('بيانات غير صحيحة / Incorrect credentials','Incorrect credentials');
  }
}

// ============================================================
// SCHOOL COORDINATOR
// ============================================================
var currentCoordSchool=null;
var liveRefreshTimer=null;

function initCoordPanel(){
  populateCoordFilters();
  renderCoordCodes();
  populateLiveFilters();
}

function coordShowSection(id){
  var sections=['coordCodes','coordLive','coordRep1','coordRep2','coordRep3','coordRep4','coordRep5','coordRep6','coordRep7'];
  sections.forEach(function(s){
    var el=document.getElementById(s);
    if(el) el.classList.toggle('hidden', s!==id);
  });
  document.querySelectorAll('.coord-nav-btn').forEach(function(btn){
    btn.classList.remove('active-nav');
    if(btn.id==='navBtn_'+id) btn.classList.add('active-nav');
  });
  if(id==='coordCodes') renderCoordCodes();
  if(id==='coordLive'){ populateLiveFilters(); loadLiveMonitor(); }
  if(id==='coordRep1') populateAttainmentFilters();
  if(liveRefreshTimer){ clearInterval(liveRefreshTimer); liveRefreshTimer=null; }
  if(id==='coordLive') liveRefreshTimer=setInterval(refreshLiveMonitor,30000);
}

function getSchoolTests(){
  if(!currentCoordSchool) return [];
  return myTests.filter(function(t){
    if(!t.students||!t.students.length) return false;
    var schoolName=currentCoordSchool.name;
    return t.school&&t.school.indexOf(schoolName)>=0;
  });
}

function populateCoordFilters(){
  var tests=getSchoolTests();
  var years={}, grades={};
  tests.forEach(function(t){
    if(t.year) years[t.year]=1;
    if(t.grade) grades[t.grade]=1;
  });
  var yearSel=document.getElementById('coordFilterYear');
  var gradeSel=document.getElementById('coordFilterGrade');
  var testSel=document.getElementById('coordFilterTest');
  yearSel.innerHTML='<option value="">كل السنوات الأكاديمية</option>'+Object.keys(years).sort().reverse().map(function(y){return '<option>'+y+'</option>';}).join('');
  gradeSel.innerHTML='<option value="">كل الصفوف</option>'+Object.keys(grades).sort().map(function(g){return '<option>'+g+'</option>';}).join('');
  testSel.innerHTML='<option value="">كل الاختبارات</option>'+tests.map(function(t){return '<option value="'+t.id+'">'+t.name+'</option>';}).join('');
}

// ══ تقارير التحصيل — Attainment Reports ══
function populateAttainmentFilters(){
  var tests=getSchoolTests();
  var years={}, subjects={};
  tests.forEach(function(t){
    if(t.year) years[t.year]=1;
    if(t.subject) subjects[t.subject]=1;
  });
  var yearSel=document.getElementById('attYearFilter');
  var subSel=document.getElementById('attSubjectFilter');
  var codeEl=document.getElementById('attSchoolCodeFilter');
  if(yearSel) yearSel.innerHTML='<option value="">الكل / All</option>'+Object.keys(years).sort().reverse().map(function(y){return '<option>'+y+'</option>';}).join('');
  if(subSel) subSel.innerHTML='<option value="">الكل / All</option>'+Object.keys(subjects).sort().map(function(s){return '<option>'+s+'</option>';}).join('');
  if(codeEl) codeEl.value=(currentCoordSchool&&currentCoordSchool.code)?currentCoordSchool.code:'—';
  var outEl=document.getElementById('attainmentReportOutput');
  if(outEl) outEl.innerHTML='';
}

function computeStudentAttainmentPct(st){
  if(!st||!st.grades) return null;
  var total=0, hasAny=false;
  Object.keys(st.grades).forEach(function(di){
    Object.keys(st.grades[di]).forEach(function(qk){
      total+=parseFloat(st.grades[di][qk])||0; hasAny=true;
    });
  });
  return hasAny?total:null;
}

// ══ نظام تصنيف مستويات التحصيل الثلاثة (Above / Met Expectation / Below) ══
var ATT_LEVELS=[
  {key:'above',en:'Above',statementAr:'الطلاب الذين يحققون مستوى أعلى من معايير المنهاج',statementEn:'Students Attaining Above Curriculum Standards',color:'#22c55e'},
  {key:'met',en:'Met Expectation',statementAr:'الطلاب الذين يحققون مستوى متوافقاً مع توقعات المنهاج',statementEn:'Students Meeting Curriculum Expectations',color:'#f59e0b'},
  {key:'below',en:'Below',statementAr:'الطلاب الذين يحققون مستوى أقل من معايير المنهاج',statementEn:'Students Attaining Below Curriculum Standards',color:'#ef4444'}
];
function _attBandOf(pct){
  if(pct>=80) return 'above';
  if(pct>=50) return 'met';
  return 'below';
}

// ══ مقياس المقادير الكمية المعتمد ══
var ATT_QUANT_TERMS=[
  {en:'Almost all',label:'Greater than 90%'},
  {en:'Most',label:'75% - 90%'},
  {en:'Large majority',label:'61% - 74%'},
  {en:'Majority',label:'50% - 60%'},
  {en:'Large minority',label:'31% - 49%'},
  {en:'Minority',label:'16% - 30%'},
  {en:'Few',label:'Up to 15%'}
];

// ══ مقياس الأحكام المعتمد (وفق 1.1.1) — ستة مستويات ══
var ATT_JUDGEMENTS=[
  {en:'Outstanding',ar:'متميز',color:'#7c3aed'},
  {en:'Very Good',ar:'جيد جداً',color:'#2563eb'},
  {en:'Good',ar:'جيد',color:'#0891b2'},
  {en:'Acceptable',ar:'مقبول',color:'#f59e0b'},
  {en:'Weak',ar:'ضعيف',color:'#f97316'},
  {en:'Very Weak',ar:'ضعيف جداً',color:'#ef4444'}
];
function computeAttainmentJudgement(pctAbove,pctAtLeastInLine){
  if(pctAbove>75) return ATT_JUDGEMENTS[0];
  if(pctAbove>=61) return ATT_JUDGEMENTS[1];
  if(pctAbove>=50) return ATT_JUDGEMENTS[2];
  if(pctAtLeastInLine>=75) return ATT_JUDGEMENTS[3];
  if(pctAtLeastInLine>=16) return ATT_JUDGEMENTS[4];
  return ATT_JUDGEMENTS[5];
}

function runAttainmentReport(){
  var yearF=document.getElementById('attYearFilter').value;
  var subF=document.getElementById('attSubjectFilter').value;
  var termF=document.getElementById('attTermFilter').value;
  var tests=getSchoolTests().filter(function(t){
    if(yearF&&t.year!==yearF) return false;
    if(subF&&t.subject!==subF) return false;
    if(termF&&t.term!==termF) return false;
    return true;
  });
  var out=document.getElementById('attainmentReportOutput');
  if(!out) return;
  out.innerHTML='<div id="attReportPrintArea">'+buildAttainmentPreambleHtml()+buildAttainmentVariableHtml(tests,{year:yearF,subject:subF,term:termF})+'</div>'
    +'<div style="display:flex;gap:10px;justify-content:center;margin-top:16px;flex-wrap:wrap">'
    +'<button onclick="printAttainmentReport()" style="background:#FACC15;color:#1e3a8a;border:none;border-radius:10px;padding:10px 22px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px">🖨️ طباعة / حفظ PDF</button>'
    +'<button onclick="exportAttainmentExcel()" style="background:#16a34a;color:white;border:none;border-radius:10px;padding:10px 22px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px">📊 تصدير Excel</button>'
    +'</div>';
  window._lastAttainmentTests=tests;
  window._lastAttainmentFilters={year:yearF,subject:subF,term:termF};
}

function buildAttainmentPreambleHtml(){
  var rows=[
    ['نوع التقرير | Report Type','تقارير التحصيل | Attainment Reports'],
    ['طبيعة الصفحة | Page Type','ديباجة ثابتة | Fixed Preamble'],
    ['الاستخدام | Use','لوحة التحكم + PDF + Excel | Dashboard + PDF + Excel'],
    ['المتغيرات | Variables','اسم المدرسة، السنة، المادة، الصف، الاختبار | School, year, subject, grade, assessment']
  ];
  var metaHtml=rows.map(function(r){return '<tr><td style="background:#f0f4ff;font-weight:800;color:#1e3a8a;padding:8px 12px;border:1px solid #dbeafe;font-size:12px;width:35%">'+r[0]+'</td><td style="padding:8px 12px;border:1px solid #dbeafe;font-size:12px;color:#1a1a2e">'+r[1]+'</td></tr>';}).join('');

  var guide=[
    {ar:'الغرض من التقرير',en:'Purpose of the Report',
     arT:'يوضح تقرير التحصيل مستوى أداء الطلاب في الاختبار المختار، ويعرض توزيعهم حسب نطاقات الأداء المتوقعة. يساعد التقرير المدرسة على فهم الوضع الحالي قبل اتخاذ قرارات الدعم والتحسين.',
     enT:"The attainment report presents students' performance in the selected assessment and shows their distribution across expected performance bands. It helps the school understand current attainment before planning support and improvement actions."},
    {ar:'نطاقات الأداء',en:'Performance Bands',
     arT:'أقل من التوقعات: أداء أقل من المستوى المستهدف. متوافق مع التوقعات: أداء داخل المستوى المتوقع. أعلى من التوقعات: أداء يتجاوز المستوى المستهدف للمنهاج أو المعيار المستخدم.',
     enT:'Below Expectations: performance below the target level. In Line with Expectations: performance within the expected level. Above Expectations: performance exceeding the target curriculum or benchmark level.'},
    {ar:'قراءة الجداول',en:'Reading the Tables',
     arT:'تعرض الجداول عدد الطلاب ونسبتهم المئوية في كل نطاق أداء. وقد تظهر النتائج على مستوى المدرسة، الصف، الشعبة، المادة، المهارة، أو الفئة الطلابية.',
     enT:'Tables show the number and percentage of students in each performance band. Results may be displayed by school, grade, class, subject, skill, or student category.'},
    {ar:'قراءة الرسوم البيانية',en:'Reading the Charts',
     arT:'تساعد الرسوم البيانية على مقارنة نسب الطلاب بين نطاقات الأداء، أو بين اختبارين، أو بين فئات مختلفة مثل النوع، الجنسية، الإماراتيين، أو ذوي الاحتياجات.',
     enT:'Charts help compare student percentages across performance bands, between assessments, or across groups such as gender, nationality, Emirati students, or students of determination.'},
    {ar:'الحكم العام',en:'Overall Judgement',
     arT:'يصدر الحكم العام بناء على توزيع الطلاب داخل نطاقات الأداء. يستخدم هذا الحكم كمؤشر موجز لجودة التحصيل، ولا يغني عن قراءة تفاصيل المهارات والفئات.',
     enT:'The overall judgement is generated from the distribution of students across performance bands. It provides a concise indicator of attainment quality, but should be read alongside skill and group-level details.'}
  ];
  var guideHtml=guide.map(function(g){
    return '<div style="margin-bottom:14px;padding:12px 16px;background:#f8fafc;border-radius:10px;border-right:3px solid #1e3a8a">'
      +'<div style="font-weight:800;color:#1e3a8a;font-size:13px;margin-bottom:4px" dir="rtl">'+g.ar+' <span style="font-weight:700;color:#64748b;font-size:11px;font-family:Montserrat,sans-serif">| '+g.en+'</span></div>'
      +'<div style="font-size:12.5px;color:#334155;line-height:1.8;margin-bottom:4px" dir="rtl">'+g.arT+'</div>'
      +'<div style="font-size:11.5px;color:#64748b;line-height:1.7;direction:ltr;text-align:left;font-family:Montserrat,sans-serif">'+g.enT+'</div>'
      +'</div>';
  }).join('');

  var terms=[
    {ar:'التحصيل',en:'Attainment',arT:'درجة أو مستوى أداء الطلاب في اختبار محدد مقارنة بنطاقات الأداء المتوقعة.',enT:'The score or performance level achieved by students in a selected assessment compared with expected performance bands.'},
    {ar:'المعيار',en:'Benchmark',arT:'الإطار أو الحدود المستخدمة لتفسير النتائج، مثل المعيار المحلي أو العالمي أو مرجع لغوي معتمد.',enT:'The framework or cut scores used to interpret results, such as a local benchmark, international benchmark, or approved language reference.'},
    {ar:'تحليل المهارات',en:'Skill Analysis',arT:'تحليل أداء الطلاب في عناصر المادة أو المهارات، مثل القراءة، الاستماع، الكتابة، التحدث، المفردات، أو القواعد.',enT:'Analysis of student performance by subject components or skills, such as reading, listening, writing, speaking, vocabulary, or grammar.'},
    {ar:'الفئات الطلابية',en:'Student Groups',arT:'تقسيم النتائج حسب خصائص محددة مثل النوع، الجنسية، الطلبة الإماراتيين، ذوي الاحتياجات، الصف، أو الشعبة.',enT:'Grouping results by selected attributes such as gender, nationality, Emirati status, students of determination, grade, or class.'},
    {ar:'ملاحظات القراءة',en:'Reading Notes',arT:'تعتمد النتائج على البيانات المتاحة وقت إصدار التقرير. وقد تختلف الأحكام إذا تغير المعيار المختار أو نطاق الدرجات أو عينة الطلاب.',enT:'Results are based on the available data at the time of report generation. Judgements may change if the selected benchmark, score ranges, or student sample changes.'}
  ];
  var termsHtml=terms.map(function(t){
    return '<div style="margin-bottom:12px;padding:10px 14px;background:#fffbeb;border-radius:10px;border-right:3px solid #AD8628">'
      +'<div style="font-weight:800;color:#8A6D1F;font-size:12.5px;margin-bottom:3px" dir="rtl">'+t.ar+' <span style="font-weight:700;color:#a68a4a;font-size:10.5px;font-family:Montserrat,sans-serif">| '+t.en+'</span></div>'
      +'<div style="font-size:12px;color:#334155;line-height:1.75" dir="rtl">'+t.arT+'</div>'
      +'<div style="font-size:11px;color:#78716c;line-height:1.6;direction:ltr;text-align:left;font-family:Montserrat,sans-serif">'+t.enT+'</div>'
      +'</div>';
  }).join('');

  return '<div style="background:white;border-radius:18px;padding:28px 30px;box-shadow:0 2px 12px rgba(0,0,0,.08);margin-bottom:24px;border:1.5px solid #e2e8f0" dir="rtl">'
    +'<div style="text-align:center;margin-bottom:18px;padding-bottom:16px;border-bottom:3px double #141F44">'
    +'<div style="font-family:Tajawal,sans-serif;font-size:19px;font-weight:900;color:#141F44">إرشادات قراءة تقرير التحصيل</div>'
    +'<div style="font-family:Montserrat,sans-serif;font-size:13px;font-weight:700;color:#AD8628;margin-top:2px">How to Read the Attainment Report</div>'
    +'</div>'
    +'<table style="width:100%;border-collapse:collapse;margin-bottom:18px">'+metaHtml+'</table>'
    +'<div style="background:#eff6ff;border:1px solid #bfdbfe;border-radius:10px;padding:12px 16px;margin-bottom:20px;font-size:12px;color:#1e40af;line-height:1.8">'
    +'<div dir="rtl">هذه الصفحات ثابتة وتظهر في بداية تقرير التحصيل لتوحيد طريقة قراءة النتائج قبل عرض الجداول والرسوم المتغيرة.</div>'
    +'<div style="direction:ltr;text-align:left;font-family:Montserrat,sans-serif;margin-top:4px">These pages are fixed front matter that standardise how attainment results are read before the dynamic tables and charts are displayed.</div>'
    +'</div>'
    +guideHtml
    +'<div style="text-align:center;margin:20px 0 14px;padding-bottom:10px;border-bottom:2px solid #e2e8f0">'
    +'<div style="font-family:Tajawal,sans-serif;font-size:15px;font-weight:900;color:#141F44">المصطلحات والمعايير</div>'
    +'<div style="font-family:Montserrat,sans-serif;font-size:11px;font-weight:700;color:#AD8628">Terms and Benchmarks</div>'
    +'</div>'
    +termsHtml
    +'<div style="text-align:center;margin-top:18px;padding-top:12px;border-top:1px solid #e2e8f0;font-size:10px;color:#94a3b8;font-family:Montserrat,sans-serif">Scholastic Platform | Fixed preamble template for attainment reports</div>'
    +'</div>';
}

function buildAttainmentDataCountsHtml(allStudents,graded){
  var registered=allStudents.length;
  var boys=graded.filter(function(x){return x.st.gender==='Male';}).length;
  var girls=graded.filter(function(x){return x.st.gender==='Female';}).length;
  var gifted=graded.filter(function(x){return x.st.gifted==='Yes';}).length;
  var sod=graded.filter(function(x){return x.st.sod==='Yes';}).length;
  var national=graded.filter(function(x){return x.st.national==='Yes';}).length;
  var items=[
    {label:'الطلاب المسجلون',labelEn:'Registered Students',val:registered,color:'#64748b'},
    {label:'أكملوا الاختبار',labelEn:'Completed the Test',val:graded.length,color:'#1e3a8a',highlight:true},
    {label:'البنين',labelEn:'Boys Completed',val:boys,color:'#2563eb'},
    {label:'البنات',labelEn:'Girls Completed',val:girls,color:'#db2777'},
    {label:'الموهوبون',labelEn:'Gifted',val:gifted,color:'#a855f7'},
    {label:'ذوو الهمم',labelEn:'Students of Determination',val:sod,color:'#ec4899'},
    {label:'المواطنون',labelEn:'National Students',val:national,color:'#16a34a'}
  ];
  var cards=items.map(function(it){
    return '<div style="flex:1;min-width:120px;background:white;border-radius:12px;padding:14px;text-align:center;border-top:3px solid '+it.color+(it.highlight?';box-shadow:0 4px 14px rgba(30,58,138,.18)':';box-shadow:0 2px 8px rgba(0,0,0,.05)')+'">'
      +'<div style="font-size:26px;font-weight:900;color:'+it.color+';font-family:Montserrat,sans-serif">'+it.val+'</div>'
      +'<div style="font-size:11.5px;font-weight:800;color:#1e293b" dir="rtl">'+it.label+'</div>'
      +'<div style="font-size:9.5px;color:#94a3b8;font-family:Montserrat,sans-serif">'+it.labelEn+'</div>'
      +'</div>';
  }).join('');
  return '<div style="margin-bottom:22px">'
    +'<div style="text-align:center;margin-bottom:12px"><div style="font-family:Tajawal,sans-serif;font-size:15px;font-weight:900;color:#141F44">البيانات</div><div style="font-family:Montserrat,sans-serif;font-size:11px;font-weight:700;color:#AD8628">Data</div></div>'
    +'<div style="display:flex;gap:10px;flex-wrap:wrap">'+cards+'</div></div>';
}

function buildAttainmentScalesReferenceHtml(){
  var descRows=[
    {level:'Outstanding',color:'#7c3aed',text:'Most students attain levels that are above curriculum standards.',textAr:'يحقق معظم الطلاب مستويات أعلى من معايير المنهاج.'},
    {level:'Very Good',color:'#2563eb',text:'The large majority of students attain levels that are above curriculum standards.',textAr:'تحقق الأغلبية الكبيرة من الطلاب مستويات أعلى من معايير المنهاج.'},
    {level:'Good',color:'#0891b2',text:'The majority of students attain levels that are above curriculum standards.',textAr:'تحقق أغلبية الطلاب مستويات أعلى من معايير المنهاج.'},
    {level:'Acceptable',color:'#f59e0b',text:'Most students attain levels that are in line with curriculum standards and a few are above.',textAr:'يحقق معظم الطلاب مستويات متوافقة مع معايير المنهاج، ويحقق قلة منهم مستوى أعلى.'},
    {level:'Weak',color:'#f97316',text:'Less than three-quarters of students attain levels that are at least in line with curriculum standards.',textAr:'يحقق أقل من ثلاثة أرباع الطلاب مستوى متوافقاً على الأقل مع معايير المنهاج.'},
    {level:'Very Weak',color:'#ef4444',text:'Few students attain levels that are in line with curriculum standards.',textAr:'يحقق قلة من الطلاب مستوى متوافقاً مع معايير المنهاج.'}
  ];
  var descHtml=descRows.map(function(r){
    return '<div style="display:flex;gap:10px;align-items:flex-start;padding:9px 12px;border-bottom:1px solid #f1f5f9">'
      +'<span style="background:'+r.color+';color:white;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:800;font-family:Montserrat,sans-serif;white-space:nowrap;min-width:88px;text-align:center">'+r.level+'</span>'
      +'<div style="flex:1"><div style="font-size:11.5px;color:#334155;line-height:1.6;direction:ltr;text-align:left;font-family:Montserrat,sans-serif">'+r.text+'</div>'
      +'<div style="font-size:11.5px;color:#64748b;line-height:1.7;margin-top:2px" dir="rtl">'+r.textAr+'</div></div>'
      +'</div>';
  }).join('');
  var termsHtml=ATT_QUANT_TERMS.map(function(t){
    return '<tr><td style="padding:7px 12px;border:1px solid #e2e8f0;background:#7c2d4a10;font-weight:800;color:#7c2d4a;font-size:12px">'+t.en+'</td><td style="padding:7px 12px;border:1px solid #e2e8f0;text-align:center;font-size:12px;color:#334155;font-family:Montserrat,sans-serif">'+t.label+'</td></tr>';
  }).join('');
  return '<div style="background:white;border-radius:16px;padding:22px 24px;margin-bottom:20px;box-shadow:0 2px 10px rgba(0,0,0,.06);border:1px solid #e2e8f0;page-break-inside:avoid">'
    +'<div style="text-align:center;margin-bottom:14px">'
    +'<div style="font-family:Tajawal,sans-serif;font-size:15px;font-weight:900;color:#141F44">معايير الحكم المعتمدة</div>'
    +'<div style="font-family:Montserrat,sans-serif;font-size:11.5px;font-weight:700;color:#AD8628">Approved Attainment Scales — 1.1.1</div>'
    +'</div>'
    +'<div style="border:1px solid #f1f5f9;border-radius:10px;overflow:hidden;margin-bottom:16px">'+descHtml+'</div>'
    +'<table style="width:100%;border-collapse:collapse"><thead><tr><th style="padding:7px 12px;border:1px solid #e2e8f0;background:#7c2d4a;color:white;font-size:11px">Term</th><th style="padding:7px 12px;border:1px solid #e2e8f0;background:#7c2d4a;color:white;font-size:11px">Percentage Range</th></tr></thead><tbody>'+termsHtml+'</tbody></table>'
    +'</div>';
}

function buildGroupAttainmentReportHtml(groupStudents,groupLabelAr,groupLabelEn){
  var n=groupStudents.length;
  if(!n) return '';
  var counts={above:0,met:0,below:0};
  groupStudents.forEach(function(x){ counts[_attBandOf(x.pct)]++; });
  var pctAbove=Math.round(counts.above/n*100);
  var pctMet=Math.round(counts.met/n*100);
  var pctAtLeastInLine=pctAbove+pctMet;
  var judgement=computeAttainmentJudgement(pctAbove,pctAtLeastInLine);

  var tableRows=ATT_LEVELS.map(function(lv){
    var cnt=counts[lv.key];
    var pct=Math.round(cnt/n*100);
    return '<tr>'
      +'<td style="padding:9px 12px;border:1px solid #e2e8f0;text-align:center"><span style="background:'+lv.color+';color:white;border-radius:6px;padding:3px 12px;font-size:11px;font-weight:800;font-family:Montserrat,sans-serif">'+lv.en+'</span></td>'
      +'<td style="padding:9px 12px;border:1px solid #e2e8f0;font-size:11.5px;direction:ltr;text-align:left;font-family:Montserrat,sans-serif;color:#334155">'+lv.statementEn+'</td>'
      +'<td style="padding:9px 12px;border:1px solid #e2e8f0;font-size:12px;color:#334155" dir="rtl">'+lv.statementAr+'</td>'
      +'<td style="padding:9px 12px;border:1px solid #e2e8f0;text-align:center;font-weight:800;color:#1e293b;font-family:Montserrat,sans-serif">'+cnt+'</td>'
      +'<td style="padding:9px 12px;border:1px solid #e2e8f0;text-align:center;font-weight:900;color:'+lv.color+';font-family:Montserrat,sans-serif">'+pct+'%</td>'
      +'</tr>';
  }).join('');

  var chartHtml=ATT_LEVELS.map(function(lv){
    var cnt=counts[lv.key];
    var pct=Math.round(cnt/n*100);
    return '<div style="display:flex;align-items:center;gap:10px;margin-bottom:9px">'
      +'<div style="width:110px;font-size:11px;color:#475569;text-align:right;flex-shrink:0;font-family:Montserrat,sans-serif">'+lv.en+'</div>'
      +'<div style="flex:1;height:20px;background:#f1f5f9;border-radius:6px;overflow:hidden"><div style="height:100%;width:'+pct+'%;background:'+lv.color+';border-radius:6px"></div></div>'
      +'<div style="width:40px;font-size:12px;font-weight:800;color:'+lv.color+';font-family:Montserrat,sans-serif">'+pct+'%</div>'
      +'</div>';
  }).join('');

  return '<div style="background:white;border-radius:16px;padding:22px 24px;margin-bottom:20px;box-shadow:0 2px 10px rgba(0,0,0,.06);border:1px solid #e2e8f0;page-break-inside:avoid">'
    +'<div style="display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:16px;padding-bottom:12px;border-bottom:2px solid #f1f5f9">'
    +'<div><div style="font-family:Tajawal,sans-serif;font-size:15px;font-weight:900;color:#141F44">'+groupLabelAr+' <span style="font-size:11px;color:#94a3b8;font-weight:700">('+n+' طالب)</span></div>'
    +'<div style="font-family:Montserrat,sans-serif;font-size:11px;font-weight:700;color:#AD8628">'+groupLabelEn+'</div></div>'
    +'<div style="background:'+judgement.color+';color:white;border-radius:10px;padding:8px 18px;text-align:center">'
    +'<div style="font-size:13px;font-weight:900;font-family:Montserrat,sans-serif">'+judgement.en+'</div>'
    +'<div style="font-size:11px;font-weight:700" dir="rtl">'+judgement.ar+'</div>'
    +'</div></div>'
    +'<div style="font-size:12px;font-weight:800;color:#1e293b;margin-bottom:8px" dir="rtl">جدول التقديرات <span style="font-weight:600;color:#94a3b8;font-family:Montserrat,sans-serif;font-size:10.5px">| Ratings Table</span></div>'
    +'<table style="width:100%;border-collapse:collapse;margin-bottom:18px"><thead><tr>'
    +'<th style="padding:8px 12px;border:1px solid #e2e8f0;background:#f8fafc;font-size:11px">Level</th>'
    +'<th style="padding:8px 12px;border:1px solid #e2e8f0;background:#f8fafc;font-size:11px">English Statement</th>'
    +'<th style="padding:8px 12px;border:1px solid #e2e8f0;background:#f8fafc;font-size:11px">العبارة بالعربي</th>'
    +'<th style="padding:8px 12px;border:1px solid #e2e8f0;background:#f8fafc;font-size:11px">العدد</th>'
    +'<th style="padding:8px 12px;border:1px solid #e2e8f0;background:#f8fafc;font-size:11px">النسبة</th>'
    +'</tr></thead><tbody>'+tableRows+'</tbody></table>'
    +'<div style="font-size:12px;font-weight:800;color:#1e293b;margin-bottom:10px" dir="rtl">الرسم البياني <span style="font-weight:600;color:#94a3b8;font-family:Montserrat,sans-serif;font-size:10.5px">| Chart</span></div>'
    +'<div>'+chartHtml+'</div>'
    +'</div>';
}

function buildAttainmentVariableHtml(tests,filters){
  var allStudents=[];
  tests.forEach(function(t){
    (t.students||[]).forEach(function(st){
      var pct=computeStudentAttainmentPct(st);
      allStudents.push({st:st,test:t,pct:pct});
    });
  });
  var header='<div style="text-align:center;margin-bottom:18px">'
    +'<div style="font-family:Tajawal,sans-serif;font-size:17px;font-weight:900;color:#141F44">نتائج التحصيل</div>'
    +'<div style="font-family:Montserrat,sans-serif;font-size:12px;font-weight:700;color:#AD8628">Attainment Results</div>'
    +'<div style="font-size:11px;color:rgba(255,255,255,.4);margin-top:6px;font-family:Montserrat,sans-serif">'+(filters.year||'All Years')+' — '+(filters.subject||'All Subjects')+' — '+(filters.term?'Term '+filters.term:'All Terms')+'</div>'
    +'</div>';

  if(!tests.length){
    return header+'<div style="text-align:center;padding:50px 20px;color:rgba(255,255,255,.35);background:rgba(255,255,255,.03);border-radius:16px">لا توجد اختبارات مطابقة للمرشحات المحددة<br><span style="font-family:Montserrat,sans-serif;font-size:12px">No tests match the selected filters</span></div>';
  }

  var graded=allStudents.filter(function(x){return x.pct!==null;});
  var dataCountsHtml=buildAttainmentDataCountsHtml(allStudents,graded);

  if(!graded.length){
    return header+dataCountsHtml+'<div style="text-align:center;padding:50px 20px;color:rgba(255,255,255,.35);background:rgba(255,255,255,.03);border-radius:16px">لا توجد نتائج تصحيح معتمدة بعد لهذه الاختبارات<br><span style="font-family:Montserrat,sans-serif;font-size:12px">No approved grading data yet — group reports will appear once grading is completed</span></div>';
  }

  var scalesHtml=buildAttainmentScalesReferenceHtml();

  var boys=graded.filter(function(x){return x.st.gender==='Male';});
  var girls=graded.filter(function(x){return x.st.gender==='Female';});
  var gifted=graded.filter(function(x){return x.st.gifted==='Yes';});
  var sod=graded.filter(function(x){return x.st.sod==='Yes';});
  var national=graded.filter(function(x){return x.st.national==='Yes';});

  var reportsHtml=buildGroupAttainmentReportHtml(graded,'جميع الطلاب','All Students')
    +buildGroupAttainmentReportHtml(boys,'البنين','Boys')
    +buildGroupAttainmentReportHtml(girls,'البنات','Girls');
  if(gifted.length>1) reportsHtml+=buildGroupAttainmentReportHtml(gifted,'الموهوبون','Gifted Students');
  if(sod.length>1) reportsHtml+=buildGroupAttainmentReportHtml(sod,'ذوو الهمم','Students of Determination');
  if(national.length>1) reportsHtml+=buildGroupAttainmentReportHtml(national,'المواطنون','National Students');

  return header+dataCountsHtml+scalesHtml+reportsHtml
    +(allStudents.length>graded.length?'<div style="text-align:center;margin-top:14px;font-size:11px;color:rgba(255,255,255,.35);font-family:Montserrat,sans-serif">'+(allStudents.length-graded.length)+' student(s) not yet graded are excluded from this analysis</div>':'');
}

function printAttainmentReport(){
  var el=document.getElementById('attReportPrintArea');
  if(!el){window.print();return;}
  var win=window.open('','_blank');
  win.document.write('<html><head><meta charset="UTF-8"><title>Attainment Report</title><style>body{font-family:Tajawal,Arial,sans-serif;margin:24px;background:white}@media print{body{margin:0}}</style></head><body>'+el.innerHTML+'</body></html>');
  win.document.close();
  setTimeout(function(){win.print();},400);
}

function exportAttainmentExcel(){
  var tests=window._lastAttainmentTests||[];
  var data=[['Test','Subject','Grade','Term','Year','Student First Name','Student Last Name','Gender','Gifted','SOD','National','Attainment %','Level']];
  tests.forEach(function(t){
    (t.students||[]).forEach(function(st){
      var pct=computeStudentAttainmentPct(st);
      if(pct===null) return;
      var lv=_attBandOf(pct);
      var levelLabel=lv==='below'?'Below':lv==='met'?'Met Expectation':'Above';
      data.push([t.name||'',t.subject||'',t.grade||'',t.term||'',t.year||'',st.firstName||'',st.lastName||'',st.gender||'',st.gifted||'No',st.sod||'No',st.national||'No',pct.toFixed(1),levelLabel]);
    });
  });
  if(data.length===1){scWarn('لا توجد بيانات تصحيح معتمدة للتصدير','No graded data available to export');return;}
  var ws=XLSX.utils.aoa_to_sheet(data);
  var wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Attainment');
  XLSX.writeFile(wb,'Attainment_Report_'+Date.now()+'.xlsx');
}

function renderCoordCodes(){
  var tests=getSchoolTests();
  var yearF=document.getElementById('coordFilterYear').value;
  var gradeF=document.getElementById('coordFilterGrade').value;
  var testF=document.getElementById('coordFilterTest').value;
  var rows=[];
  var idx=1;
  tests.forEach(function(t){
    if(yearF&&t.year!==yearF) return;
    if(gradeF&&t.grade!==gradeF) return;
    if(testF&&String(t.id)!==testF) return;
    (t.students||[]).forEach(function(st){
      rows.push('<tr>'
        +'<td class="font-en">'+idx+'</td>'
        +'<td>'+(t.year||'—')+'</td>'
        +'<td>'+(t.grade||'—')+'</td>'
        +'<td style="font-weight:700">'+(t.name||'—')+'</td>'
        +'<td>'+(st.firstName||'—')+'</td>'
        +'<td>'+(st.lastName||'—')+'</td>'
        +'<td class="font-en" style="color:#FACC15">'+(st.username||'—')+'</td>'
        +'<td class="font-en">'+(st.password||'—')+'</td>'
        +'</tr>');
      idx++;
    });
  });
  var body=document.getElementById('coordCodesBody');
  body.innerHTML=rows.length?rows.join(''):'<tr><td colspan="8" style="color:rgba(255,255,255,.3);padding:32px;text-align:center">لا توجد أكواد / No codes found</td></tr>';
}

function exportCoordCodesToExcel(){
  var tests=getSchoolTests();
  var yearF=document.getElementById('coordFilterYear').value;
  var gradeF=document.getElementById('coordFilterGrade').value;
  var testF=document.getElementById('coordFilterTest').value;
  var csv='#,السنة الأكاديمية,الصف,اسم الاختبار,الاسم الأول,الاسم الثاني,Username,Password\n';
  var idx=1;
  tests.forEach(function(t){
    if(yearF&&t.year!==yearF) return;
    if(gradeF&&t.grade!==gradeF) return;
    if(testF&&String(t.id)!==testF) return;
    (t.students||[]).forEach(function(st){
      csv+=idx+','+[t.year,t.grade,t.name,st.firstName,st.lastName,st.username,st.password].map(function(v){return '"'+(v||'')+'\"';}).join(',')+'\n';
      idx++;
    });
  });
  var BOM='\uFEFF';
  var blob=new Blob([BOM+csv],{type:'text/csv;charset=utf-8'});
  var a=document.createElement('a');
  a.href=URL.createObjectURL(blob);
  a.download='student_codes_'+( currentCoordSchool?currentCoordSchool.name:'')+'.csv';
  a.click();
}

function populateLiveFilters(){
  var tests=getSchoolTests();
  var years={}, subjects={};
  tests.forEach(function(t){
    if(t.year) years[t.year]=1;
    if(t.subject) subjects[t.subject]=1;
  });
  document.getElementById('liveFilterYear').innerHTML='<option value="">الكل</option>'+Object.keys(years).sort().reverse().map(function(y){return '<option>'+y+'</option>';}).join('');
  document.getElementById('liveFilterSubject').innerHTML='<option value="">الكل</option>'+Object.keys(subjects).sort().map(function(s){return '<option>'+s+'</option>';}).join('');
  filterLiveTests();
}

function filterLiveTests(){
  var tests=getSchoolTests();
  var yearF=document.getElementById('liveFilterYear').value;
  var subF=document.getElementById('liveFilterSubject').value;
  var filtered=tests.filter(function(t){
    if(yearF&&t.year!==yearF) return false;
    if(subF&&t.subject!==subF) return false;
    return true;
  });
  var sel=document.getElementById('liveFilterTest');
  var cur=sel.value;
  sel.innerHTML='<option value="">-- اختر اختباراً --</option>'+filtered.map(function(t){return '<option value="'+t.id+'">'+t.name+'</option>';}).join('');
  if(cur) sel.value=cur;
  loadLiveMonitor();
}

function loadLiveMonitor(){
  var testId=document.getElementById('liveFilterTest').value;
  if(!testId){
    document.getElementById('liveTestInfo').classList.add('hidden');
    document.getElementById('liveTableHead').innerHTML='<tr><th colspan="10" style="color:rgba(255,255,255,.3);font-weight:400">اختر اختباراً لعرض بيانات المراقبة</th></tr>';
    document.getElementById('liveTableBody').innerHTML='';
    return;
  }
  var t=myTests.find(function(x){return String(x.id)===testId;});
  if(!t) return;

  // Test info bar
  var statusLabel=t.isActive?'<span style="background:rgba(74,222,128,.2);color:#4ade80;border:1px solid rgba(74,222,128,.4);border-radius:20px;padding:3px 12px;font-size:12px;font-weight:700">نشط ✅</span>':'<span style="background:rgba(248,113,113,.2);color:#f87171;border:1px solid rgba(248,113,113,.4);border-radius:20px;padding:3px 12px;font-size:12px;font-weight:700">معطل ⛔</span>';
  document.getElementById('liveTestStatus').innerHTML=statusLabel;
  document.getElementById('liveTestFrom').textContent=t.activeFrom||'—';
  document.getElementById('liveTestTo').textContent=t.activeTo||'—';
  document.getElementById('liveTestDuration').textContent=(t.durationMin||t.duration||'—')+' دقيقة';
  document.getElementById('liveTestInfo').classList.remove('hidden');

  // Build domain headers
  var domains=t.domains||[];
  var domHeaders=domains.map(function(d,i){return '<th style="white-space:nowrap">المجال '+(i+1)+'<br><span style="font-size:10px;opacity:.6;font-weight:400">'+(d.name||'')+'</span></th>';}).join('');

  document.getElementById('liveTableHead').innerHTML='<tr style="font-size:12px">'
    +'<th style="font-family:Montserrat,sans-serif">Student ID</th>'
    +'<th>الاسم الأول</th>'
    +'<th>الاسم الثاني</th>'
    +'<th>المادة</th>'
    +'<th>اسم الاختبار</th>'
    +'<th>الحالة</th>'
    +domHeaders
    +'</tr>';

  var students=t.students||[];
  if(!students.length){
    document.getElementById('liveTableBody').innerHTML='<tr><td colspan="'+(6+domains.length)+'" style="color:rgba(255,255,255,.3);padding:32px;text-align:center">لا يوجد طلاب</td></tr>';
    return;
  }

  var rows=students.map(function(st){
    var statusCell=st.started
      ?'<span style="background:rgba(74,222,128,.2);color:#4ade80;border:1px solid rgba(74,222,128,.4);border-radius:20px;padding:3px 10px;font-size:11px;font-weight:700">تم البدء ✅</span>'
      :'<span style="background:rgba(250,204,21,.15);color:#FACC15;border:1px solid rgba(250,204,21,.3);border-radius:20px;padding:3px 10px;font-size:11px;font-weight:700">لم يبدأ بعد ⏳</span>';

    var domCells=domains.map(function(d,i){
      var done=st.domainsCompleted&&st.domainsCompleted[i];
      return '<td style="text-align:center">'+(done
        ?'<span style="color:#4ade80;font-size:13px;font-weight:700">أكمل ✅</span>'
        :'<span style="color:rgba(255,255,255,.3);font-size:12px">لم يكمل</span>')+'</td>';
    }).join('');

    return '<tr style="font-size:13px">'
      +'<td class="font-en" style="color:#FACC15;font-weight:700">'+(st.username||'—')+'</td>'
      +'<td>'+(st.firstName||'—')+'</td>'
      +'<td>'+(st.lastName||'—')+'</td>'
      +'<td style="font-size:12px">'+(t.subject||'—')+'</td>'
      +'<td style="font-size:12px;font-weight:700">'+(t.name||'—')+'</td>'
      +'<td>'+statusCell+'</td>'
      +domCells
      +'</tr>';
  }).join('');

  document.getElementById('liveTableBody').innerHTML=rows;
}

async function refreshLiveMonitor(){
  showGHLoader(true);
  var fresh=await ghRead('tests.json');
  if(fresh) myTests=fresh;
  showGHLoader(false);
  loadLiveMonitor();
}
function toggleSection(id){
  document.getElementById('mainDashboard').classList.add('hidden-section');
  document.getElementById('studentPortal').classList.add('hidden');
  ['adminPanel','supervisorPanel','schoolCoordPanel'].forEach(function(x){document.getElementById(x).classList.add('hidden');});
  document.getElementById(id).classList.remove('hidden');
  var nav=document.getElementById('topLeftNav');if(nav)nav.style.display='none';
  if(id==='supervisorPanel'){populateCountrySelects();populateAcademicYears();filterSchoolsByCountryAndCurriculum();}
}

// ============================================================
// STUDENT PORTAL
// ============================================================
function showStudentPortal(student,test){
  document.getElementById('mainDashboard').classList.add('hidden-section');
  document.getElementById('studentPortalName').textContent=(student.firstName||'')+' '+(student.lastName||'');
  document.getElementById('studentPortalSchool').textContent=(student.schoolName||'')+' | '+(test.grade||'')+' | '+(test.subject||'');
  var cards=document.getElementById('studentTestCards');
  var isComplete=student.completed||false;
  var now=new Date();
  var isActive=true;
  if(test.activeFrom&&test.activeTo){isActive=now>=new Date(test.activeFrom)&&now<=new Date(test.activeTo);}
  var statusColor=isComplete?'#22c55e':isActive?'#f59e0b':'#94a3b8';
  var statusBg=isComplete?'linear-gradient(135deg,#dcfce7,#bbf7d0)':isActive?'linear-gradient(135deg,#fef9c3,#fde68a)':'linear-gradient(135deg,#f1f5f9,#e2e8f0)';
  var statusText=isComplete?'✅ Completed / مكتمل':isActive?'📝 Available / متاح':'🔒 Unavailable / غير متاح';
  var statusTextColor=isComplete?'#15803d':isActive?'#92400e':'#64748b';
  var clickAction=isComplete?'_spAlreadyDone()':isActive?'startStudentTest()':'_spNotAvail()';

  cards.innerHTML='<div onclick="'+clickAction+'" style="background:white;border-radius:24px;overflow:hidden;box-shadow:0 0 0 3px '+statusColor+',0 8px 32px rgba(0,0,0,.12);cursor:pointer;transition:.3s;width:100%;max-width:380px;margin:0 auto">'
    +'<div style="background:linear-gradient(135deg,#1e3a8a,#7e22ce);padding:24px;text-align:center">'
    +(test.logoSrc?'<img src="'+test.logoSrc+'" style="width:56px;height:56px;object-fit:contain;border-radius:12px;background:white;padding:4px;margin-bottom:12px">':'<div style="font-size:52px;margin-bottom:12px">'+(isComplete?'✅':isActive?'📝':'🔒')+'</div>')
    +'<h3 style="font-size:19px;font-weight:900;color:white;margin-bottom:6px">'+(test.name||'الاختبار')+'</h3>'
    +'<p style="font-size:12px;color:rgba(255,255,255,.7);font-family:Montserrat,sans-serif">'+(test.subject||'')+' | '+(test.grade||'')+'</p>'
    +'</div>'
    +'<div style="padding:14px 20px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #f1f5f9">'
    +'<div style="font-size:12px;color:#64748b"><span style="font-family:Montserrat,sans-serif">Term </span>'+(test.term||'—')+'</div>'
    +'<div style="font-size:12px;color:#64748b;font-family:Montserrat,sans-serif">'+(test.year||'')+'</div>'
    +'</div>'
    +'<div style="padding:16px 20px;text-align:center;background:'+statusBg+'">'
    +'<div style="font-size:14px;font-weight:800;color:'+statusTextColor+';font-family:Montserrat,sans-serif">'+statusText+'</div>'
    +'</div>'
    +'</div>';
  document.getElementById('studentPortal').classList.remove('hidden');
}

function startStudentTest(){
  if(!currentStudentData)return;
  testData.domains=currentStudentData.test.domains||[];
  testData.logoSrc=currentStudentData.test.logoSrc||'';
  testData.instructionsAr=currentStudentData.test.instructionsAr||'';
  testData.instructionsEn=currentStudentData.test.instructionsEn||'';
  testData.displayMode=currentStudentData.test.displayMode||1;
  sw_domainIdx=0;sw_qIdx=0;sw_answers={};
  openStudentWindow(0,true);
}

// ============================================================
// WIZARD
// ============================================================
function startTestWizard(){
  // تصفير كامل — لا بيانات سابقة
  testData={domains:[],selectedSchools:[],logoSrc:'',displayMode:1,instructionsAr:'',instructionsEn:''};
  selectedSchools=[];editingTestId=null;
  // تصفير حقول الخطوة الأولى
  setTimeout(function(){
    var fields=['testCountry','testCurriculum','testGrade','testYear','testSubject','testTerm','testName','domainCount','testLocalCountry','testNationalCountry','testSubjectOther'];
    fields.forEach(function(id){var el=document.getElementById(id);if(el)el.value='';});
    var logo=document.getElementById('testLogoPreview');if(logo){logo.src='';logo.style.display='none';}
    var ph=document.getElementById('testLogoPlaceholder');if(ph)ph.style.display='block';
    var instrAr=document.getElementById('instrEditorAr');if(instrAr)instrAr.innerHTML='';
    var instrEn=document.getElementById('instrEditorEn');if(instrEn)instrEn.innerHTML='';
    var link=document.getElementById('testLinkDisplay');if(link)link.textContent='';
    _saveDraft();
  },100);
  document.getElementById('supMainOptions').classList.add('hidden');
  document.getElementById('supTitle').classList.add('hidden');
  document.getElementById('testCreationWizard').classList.remove('hidden');
  populateCountrySelects();populateAcademicYears();filterSchoolsByCountryAndCurriculum();populateReviewerDropdowns();
  testData.displayMode=1;
  setTimeout(function(){selectDisplayMode(1);},80);
  goToStep(1);
}
function syncStep1BasicInfoToTestData(){
  var testNameEl=document.getElementById('testName');
  testData.testName=testNameEl?testNameEl.value||'':'';
  var gradeEl=document.getElementById('grade');
  testData.grade=gradeEl?gradeEl.value||'':'';
  var subjectEl=document.getElementById('subject');
  testData.subject=subjectEl&&subjectEl.selectedIndex>=0?subjectEl.options[subjectEl.selectedIndex].text:'';
  var termEl=document.getElementById('term');
  testData.term=termEl?termEl.value||'':'';
  var yearEl=document.getElementById('academicYear');
  testData.year=yearEl?yearEl.value||'':'';
  var countryEl=document.getElementById('country');
  testData.country=countryEl?countryEl.value||'':'';
  var curriculumEl=document.getElementById('curriculum');
  testData.curriculum=curriculumEl?curriculumEl.value||'':'';
}
function goToStep(s){
  try{
    if(s===2){
      var errors=validateStep1();
      if(errors.length){
        alert('⚠️ يرجى إكمال البيانات التالية قبل المتابعة:\n\n'+errors.join('\n'));
        return;
      }
      try{ saveDomainDataFromStep1(); }catch(e){ console.warn('saveDomainDataFromStep1 error:',e); }
      try{ syncStep1BasicInfoToTestData(); }catch(e){}
    }
    document.querySelectorAll('.wizard-step').forEach(function(x){x.classList.add('hidden');});
    var stepEl=document.getElementById('wizardStep'+s);
    if(stepEl) stepEl.classList.remove('hidden');
    for(var i=1;i<=3;i++){var d=document.getElementById('dot'+i);if(d) d.className='step-dot'+(i<s?' done':i===s?' active':'');}
    if(s===3){
      try{ syncStep1BasicInfoToTestData(); }catch(e){}
      try{
        var insArEl=document.getElementById('insAr');
        var insEnEl=document.getElementById('insEn');
        if(insArEl) testData.instructionsAr=insArEl.innerHTML.trim();
        if(insEnEl) testData.instructionsEn=insEnEl.innerHTML.trim();
      }catch(e){}
      buildDomainsGrid();
    }
  }catch(e){
    console.error('goToStep error:',e);
    alert('حدث خطأ غير متوقع / Unexpected error: '+e.message);
  }
}

function validateStep1(){
  var errors=[];
  // Only validate essential basic info — domain details are configured in Step 3
  if(!(document.getElementById('testName')&&document.getElementById('testName').value.trim()))
    errors.push('• اسم الاختبار / Test Name');
  if(!(document.getElementById('grade')&&document.getElementById('grade').value))
    errors.push('• الصف / Grade');
  if(!(document.getElementById('subject')&&document.getElementById('subject').value))
    errors.push('• المادة / Subject');
  if(!(document.getElementById('term')&&document.getElementById('term').value))
    errors.push('• الفصل الدراسي / Term');
  if(!(document.getElementById('academicYear')&&document.getElementById('academicYear').value))
    errors.push('• السنة الدراسية / Academic Year');
  var qty=parseInt(document.getElementById('domQty')?document.getElementById('domQty').value:0)||0;
  if(!qty) errors.push('• عدد المجالات / Number of Domains');
  return errors;
}

function saveDomainDataFromStep1(){
  var qty=parseInt(document.getElementById('domQty').value)||0;
  if(!qty) return;
  // Grow or shrink testData.domains to match qty
  while(testData.domains.length<qty) testData.domains.push({nameAr:'',nameEn:'',weight:Math.floor(100/qty),time:0,qCount:0,questions:[],hasBranches:false,branches:[],complete:false});
  testData.domains.length=qty;
  // Save whatever domain row data exists (may be empty if user hasn't filled yet)
  for(var i=1;i<=qty;i++){
    var d=testData.domains[i-1];
    if(document.getElementById('dn'+i)) d.nameAr=document.getElementById('dn'+i).value||d.nameAr||'';
    if(document.getElementById('dne'+i)) d.nameEn=document.getElementById('dne'+i).value||d.nameEn||'';
    if(document.getElementById('dw'+i)) d.weight=Number(document.getElementById('dw'+i).value)||d.weight||Math.floor(100/qty);
    if(document.getElementById('dt'+i)) d.time=Number(document.getElementById('dt'+i).value)||d.time||0;
    var hasBr=document.getElementById('brYes'+i)&&document.getElementById('brYes'+i).getAttribute('data-active')==='1';
    d.hasBranches=hasBr;
    if(hasBr){
      var brCount=parseInt((document.getElementById('brc'+i)||{}).value)||0;
      var oldBranches=d.branches||[];
      d.branches=[];
      for(var b=0;b<brCount;b++){
        var prev=oldBranches[b]||{};
        d.branches.push({
          nameAr:(document.getElementById('brn_'+i+'_'+b)||{}).value||prev.nameAr||('فرع '+(b+1)),
          nameEn:(document.getElementById('brne_'+i+'_'+b)||{}).value||prev.nameEn||('Branch '+(b+1)),
          weight:Number((document.getElementById('brw_'+i+'_'+b)||{}).value)||prev.weight||0,
          time:Number((document.getElementById('brt_'+i+'_'+b)||{}).value)||prev.time||0,
          qCount:Number((document.getElementById('brq_'+i+'_'+b)||{}).value)||prev.qCount||0,
          questions:prev.questions||[]
        });
      }
      d.qCount=0;
    } else {
      if(document.getElementById('dq'+i)) d.qCount=Number(document.getElementById('dq'+i).value)||d.qCount||0;
      d.branches=d.hasBranches?d.branches:[];
    }
  }
  _saveDraft();
}
function cancelWizard(){document.getElementById('testCreationWizard').classList.add('hidden');document.getElementById('supMainOptions').classList.remove('hidden');document.getElementById('supTitle').classList.remove('hidden');}
function previewReviewerNotice(){var r1=document.getElementById('reviewer1').value;document.getElementById('reviewerNoticeBox').classList.toggle('hidden',!r1);}
function populateReviewerDropdowns(){
  var sel=document.getElementById('reviewer1'); if(!sel) return;
  sel.innerHTML='<option value="">-- اختر --</option>';
  var o0=document.createElement('option');o0.value=DEFAULT_SUP.name;o0.textContent=DEFAULT_SUP.name+' ('+DEFAULT_SUP.username+')';sel.appendChild(o0);
  supervisors.filter(function(s){return s.active;}).forEach(function(s){var o=document.createElement('option');o.value=s.name;o.textContent=s.name+' ('+s.username+')';sel.appendChild(o);});
}
function generateTestLink(){
  var name=document.getElementById('testName').value.trim(),box=document.getElementById('testLinkBox');
  if(name.length>2){document.getElementById('testLinkText').textContent='scholastic.app/test/'+name.replace(/\s+/g,'-').toLowerCase()+'-'+Math.random().toString(36).slice(-4);box.classList.remove('hidden');}
  else box.classList.add('hidden');
}
function generateDomains(){
  var qty=parseInt(document.getElementById('domQty').value)||0,cont=document.getElementById('domContainer');
  if(qty<1||qty>10){cont.innerHTML='';return;}
  var base=Math.floor(100/qty),rem=100-base*qty;
  var html='<div style="border-top:1px solid rgba(255,255,255,.1);padding-top:16px;margin-top:8px">'
    +'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:14px">'
    +'<p style="color:#FACC15;font-size:14px;font-weight:800;margin:0">📊 مجالات الاختبار / Domains</p>'
    +'<div style="display:flex;align-items:center;gap:8px"><span style="font-size:12px;color:rgba(255,255,255,.4)">المجموع:</span><span id="domainTotal" style="font-family:Montserrat,sans-serif;font-weight:900;color:#4ade80;font-size:15px">0%</span></div>'
    +'</div><div id="domainsList"></div></div>';
  cont.innerHTML=html;
  var rows='';
  for(var i=1;i<=qty;i++){
    var existing=testData.domains[i-1]||{};
    var w=existing.weight||(i===1?base+rem:base);
    rows+=buildDomainRow(i,w,existing.time||'',existing.qCount||'',existing.hasBranches||false,existing.branches?existing.branches.length:2,existing.branches||[],existing);
  }
  document.getElementById('domainsList').innerHTML=rows;
  recalcTotal();
}

function buildDomainRow(i,w,t,q,hasBr,brCount,branches,existing){
  var brRows='';
  if(hasBr){
    for(var b=0;b<brCount;b++){
      var br=branches[b]||{};
      var bw=br.weight||(b===0?Math.floor(w/brCount)+(w%brCount):Math.floor(w/brCount));
      var bt=br.time||'';var bq=br.qCount||'';
      brRows+='<div style="display:flex;align-items:flex-start;gap:8px;background:rgba(34,197,94,.05);border:2px solid rgba(34,197,94,.35);border-radius:14px;padding:12px;margin-bottom:8px">'
        +'<div style="width:26px;height:26px;border-radius:50%;background:#22c55e;color:white;font-weight:800;font-size:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0">'+(b+1)+'</div>'
        +'<div style="flex:1;display:grid;grid-template-columns:1fr 1fr;gap:6px">'
        +'<input type="text" id="brn_'+i+'_'+b+'" class="wizard-input" placeholder="اسم الفرع بالعربي" value="'+(br.nameAr||'')+'" style="font-size:12px">'
        +'<input type="text" id="brne_'+i+'_'+b+'" class="wizard-input font-en" placeholder="Branch name" dir="ltr" value="'+(br.nameEn||'')+'" style="font-size:12px">'
        +'<div style="display:flex;align-items:center;gap:4px"><span style="font-size:10px;color:rgba(255,255,255,.4)">وزن%</span><input type="number" id="brw_'+i+'_'+b+'" class="wizard-input font-en text-center" value="'+bw+'" min="1" oninput="recalcBranchW('+i+')" style="width:58px;font-size:12px"></div>'
        +'<div style="display:flex;align-items:center;gap:4px"><span style="font-size:10px;color:rgba(255,255,255,.4)">زمن(د)</span><input type="number" id="brt_'+i+'_'+b+'" class="wizard-input font-en text-center" value="'+bt+'" min="1" placeholder="20" oninput="recalcBranchTime('+i+')" style="width:58px;font-size:12px"></div>'
        +'<div style="display:flex;align-items:center;gap:4px"><span style="font-size:10px;color:rgba(255,255,255,.4)">أسئلة</span><input type="number" id="brq_'+i+'_'+b+'" class="wizard-input font-en text-center" value="'+bq+'" min="1" placeholder="5" style="width:58px;font-size:12px"></div>'
        +'</div></div>';
    }
  }
  var brYStyle='border:2px solid '+(hasBr?'#22c55e':'rgba(255,255,255,.2)')+';background:'+(hasBr?'rgba(34,197,94,.2)':'transparent')+';color:'+(hasBr?'#4ade80':'rgba(255,255,255,.6)');
  var brNStyle='border:2px solid '+(!hasBr?'#f87171':'rgba(255,255,255,.2)')+';background:'+(!hasBr?'rgba(248,113,113,.2)':'transparent')+';color:'+(!hasBr?'#f87171':'rgba(255,255,255,.6)');
  return '<div id="domRow_'+i+'" style="border:2px solid rgba(251,146,60,.6);border-radius:18px;padding:16px;margin-bottom:14px;background:rgba(251,146,60,.04)">'
    +'<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:12px">'
    +'<div style="width:32px;height:32px;border-radius:50%;background:linear-gradient(135deg,#f97316,#ea580c);color:white;font-weight:900;font-size:14px;display:flex;align-items:center;justify-content:center;flex-shrink:0">'+i+'</div>'
    +'<div style="flex:1;display:grid;grid-template-columns:1fr 1fr;gap:8px">'
    +'<input type="text" id="dn'+i+'" class="wizard-input" placeholder="اسم المجال بالعربي" value="'+(existing.nameAr||'')+'" style="font-size:13px">'
    +'<input type="text" id="dne'+i+'" class="wizard-input font-en" placeholder="Domain name in English" dir="ltr" value="'+(existing.nameEn||'')+'" style="font-size:13px">'
    +'</div></div>'
    +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(110px,1fr));gap:8px;margin-bottom:12px">'
    +'<div><label style="font-size:10px;color:#fb923c;font-weight:700;display:block;margin-bottom:4px">📊 الوزن % / Weight</label><div style="display:flex;align-items:center;gap:4px"><input type="number" id="dw'+i+'" class="wizard-input weight-val font-en text-center" value="'+w+'" min="1" max="100" oninput="recalcTotal()" style="font-size:13px"><span style="color:#FACC15;font-weight:800">%</span></div></div>'
    +'<div><label style="font-size:10px;color:#fb923c;font-weight:700;display:block;margin-bottom:4px">⏱ الزمن (د) / Time</label><input type="number" id="dt'+i+'" class="wizard-input font-en text-center" value="'+t+'" min="1" placeholder="30" oninput="recalcBranchTime('+i+')" style="font-size:13px"></div>'
    +'<div id="dqBox'+i+'" style="display:'+(hasBr?'none':'block')+'"><label style="font-size:10px;color:#fb923c;font-weight:700;display:block;margin-bottom:4px">❓ عدد الأسئلة</label><input type="number" id="dq'+i+'" class="wizard-input font-en text-center" value="'+q+'" min="1" placeholder="10" style="font-size:13px"></div>'
    +'</div>'
    +'<div style="background:rgba(255,255,255,.05);border-radius:12px;padding:10px 14px;margin-bottom:10px;display:flex;align-items:center;gap:10px;flex-wrap:wrap">'
    +'<span style="font-size:13px;font-weight:700">هل للمجال فروع؟ / Has Branches?</span>'
    +'<button type="button" id="brYes'+i+'" data-active="'+(hasBr?'1':'0')+'" onclick="toggleBranches('+i+',true)" style="border-radius:20px;padding:5px 14px;font-size:12px;font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;'+brYStyle+'">✅ نعم</button>'
    +'<button type="button" id="brNo'+i+'" onclick="toggleBranches('+i+',false)" style="border-radius:20px;padding:5px 14px;font-size:12px;font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;'+brNStyle+'">❌ لا</button>'
    +(hasBr?'<div style="display:flex;align-items:center;gap:6px"><span style="font-size:12px;color:rgba(255,255,255,.5)">عدد الفروع:</span><input type="number" id="brc'+i+'" class="wizard-input font-en text-center" value="'+brCount+'" min="2" max="8" oninput="regenerateBranches('+i+')" style="width:58px;font-size:12px"></div>':'<input type="number" id="brc'+i+'" style="display:none" value="'+brCount+'">')
    +'</div>'
    +'<div id="brContainer_'+i+'">'+(hasBr?'<p style="font-size:11px;color:#4ade80;margin-bottom:8px;font-weight:700">الفروع — مجموع الأوزان = وزن المجال | مجموع الأزمنة = زمن المجال</p>'+brRows:'')+'</div>'
    +(hasBr?'<div style="display:flex;gap:16px;background:rgba(0,0,0,.25);border-radius:10px;padding:8px 14px;font-size:12px;font-family:Montserrat,sans-serif;margin-top:6px"><span>وزن الفروع: <span id="brWTotal_'+i+'" style="font-weight:800;color:#4ade80">0%</span></span><span>زمن الفروع: <span id="brTTotal_'+i+'" style="font-weight:800;color:#4ade80">0د</span></span></div>':'')
    +'</div>';
}

function toggleBranches(i,yes){
  var yBtn=document.getElementById('brYes'+i);var nBtn=document.getElementById('brNo'+i);
  var dqBox=document.getElementById('dqBox'+i);var brcInput=document.getElementById('brc'+i);
  var cont=document.getElementById('brContainer_'+i);
  if(yBtn) yBtn.setAttribute('data-active',yes?'1':'0');
  if(yBtn) yBtn.style.cssText='border-radius:20px;padding:5px 14px;font-size:12px;font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;border:2px solid '+(yes?'#22c55e':'rgba(255,255,255,.2)')+';background:'+(yes?'rgba(34,197,94,.2)':'transparent')+';color:'+(yes?'#4ade80':'rgba(255,255,255,.6)');
  if(nBtn) nBtn.style.cssText='border-radius:20px;padding:5px 14px;font-size:12px;font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;border:2px solid '+(!yes?'#f87171':'rgba(255,255,255,.2)')+';background:'+(!yes?'rgba(248,113,113,.2)':'transparent')+';color:'+(!yes?'#f87171':'rgba(255,255,255,.6)');
  if(dqBox) dqBox.style.display=yes?'none':'block';
  if(brcInput) brcInput.style.display=yes?'':'none';
  if(yes){if(brcInput) brcInput.style.display='';regenerateBranches(i);}
  else if(cont){cont.innerHTML='';var tb=cont.nextElementSibling;if(tb&&tb.querySelector&&tb.querySelector('[id^="brWTotal_"]')) tb.remove();}
}

function regenerateBranches(i){
  var count=parseInt(document.getElementById('brc'+i).value)||2;
  var domW=Number(document.getElementById('dw'+i).value)||100;
  var domT=Number(document.getElementById('dt'+i).value)||0;
  var bW=Math.floor(domW/count),rW=domW-bW*count;
  var bT=Math.floor(domT/count),rT=domT-bT*count;
  var cont=document.getElementById('brContainer_'+i);
  var existing=[];
  for(var b=0;b<20;b++){
    var bwEl=document.getElementById('brw_'+i+'_'+b);if(!bwEl) break;
    existing.push({nameAr:(document.getElementById('brn_'+i+'_'+b)||{}).value||'',nameEn:(document.getElementById('brne_'+i+'_'+b)||{}).value||'',weight:Number(bwEl.value)||0,time:Number((document.getElementById('brt_'+i+'_'+b)||{}).value)||0,qCount:Number((document.getElementById('brq_'+i+'_'+b)||{}).value)||0});
  }
  var rows='<p style="font-size:11px;color:#4ade80;margin-bottom:8px;font-weight:700">الفروع — مجموع الأوزان = وزن المجال | مجموع الأزمنة = زمن المجال</p>';
  for(var b=0;b<count;b++){
    var prev=existing[b]||{};
    var bw=prev.weight||(b===0?bW+rW:bW);var bt=prev.time||(b===0?bT+rT:bT);var bq=prev.qCount||'';
    rows+='<div style="display:flex;align-items:flex-start;gap:8px;background:rgba(34,197,94,.05);border:2px solid rgba(34,197,94,.35);border-radius:14px;padding:12px;margin-bottom:8px">'
      +'<div style="width:26px;height:26px;border-radius:50%;background:#22c55e;color:white;font-weight:800;font-size:12px;display:flex;align-items:center;justify-content:center;flex-shrink:0">'+(b+1)+'</div>'
      +'<div style="flex:1;display:grid;grid-template-columns:1fr 1fr;gap:6px">'
      +'<input type="text" id="brn_'+i+'_'+b+'" class="wizard-input" placeholder="اسم الفرع بالعربي" value="'+(prev.nameAr||'')+'" style="font-size:12px">'
      +'<input type="text" id="brne_'+i+'_'+b+'" class="wizard-input font-en" placeholder="Branch name" dir="ltr" value="'+(prev.nameEn||'')+'" style="font-size:12px">'
      +'<div style="display:flex;align-items:center;gap:4px"><span style="font-size:10px;color:rgba(255,255,255,.4)">وزن%</span><input type="number" id="brw_'+i+'_'+b+'" class="wizard-input font-en text-center" value="'+bw+'" min="1" oninput="recalcBranchW('+i+')" style="width:58px;font-size:12px"></div>'
      +'<div style="display:flex;align-items:center;gap:4px"><span style="font-size:10px;color:rgba(255,255,255,.4)">زمن(د)</span><input type="number" id="brt_'+i+'_'+b+'" class="wizard-input font-en text-center" value="'+bt+'" min="1" placeholder="20" oninput="recalcBranchTime('+i+')" style="width:58px;font-size:12px"></div>'
      +'<div style="display:flex;align-items:center;gap:4px"><span style="font-size:10px;color:rgba(255,255,255,.4)">أسئلة</span><input type="number" id="brq_'+i+'_'+b+'" class="wizard-input font-en text-center" value="'+bq+'" min="1" placeholder="5" style="width:58px;font-size:12px"></div>'
      +'</div></div>';
  }
  cont.innerHTML=rows;
  var existing2=cont.nextElementSibling;
  if(!existing2||!existing2.querySelector||!existing2.querySelector('[id^="brWTotal_"]')){
    var bar=document.createElement('div');
    bar.style.cssText='display:flex;gap:16px;background:rgba(0,0,0,.25);border-radius:10px;padding:8px 14px;font-size:12px;font-family:Montserrat,sans-serif;margin-top:6px';
    bar.innerHTML='<span>وزن الفروع: <span id="brWTotal_'+i+'" style="font-weight:800;color:#4ade80">0%</span></span><span>زمن الفروع: <span id="brTTotal_'+i+'" style="font-weight:800;color:#4ade80">0د</span></span>';
    cont.parentNode.insertBefore(bar,cont.nextSibling);
  }
  recalcBranchW(i);recalcBranchTime(i);
}

function recalcBranchW(i){
  var domW=Number(document.getElementById('dw'+i)?document.getElementById('dw'+i).value:0)||0;
  var total=0;for(var b=0;b<20;b++){var el=document.getElementById('brw_'+i+'_'+b);if(!el) break;total+=Number(el.value)||0;}
  var el=document.getElementById('brWTotal_'+i);if(el){el.textContent=total+'%';el.style.color=Math.round(total)===Math.round(domW)?'#4ade80':'#f87171';}
}

function recalcBranchTime(i){
  var domT=Number(document.getElementById('dt'+i)?document.getElementById('dt'+i).value:0)||0;
  var total=0;for(var b=0;b<20;b++){var el=document.getElementById('brt_'+i+'_'+b);if(!el) break;total+=Number(el.value)||0;}
  var el=document.getElementById('brTTotal_'+i);if(el){el.textContent=total+'د';el.style.color=domT>0&&Math.round(total)===Math.round(domT)?'#4ade80':'#f87171';}
}

function syncDomainName(idx,lang,val){
  // sync to testData immediately
  if(!testData.domains[idx]) testData.domains[idx]={};
  if(lang==='ar') testData.domains[idx].nameAr=val;
  else testData.domains[idx].nameEn=val;
  _saveDraft();
}
function syncBranchName(domIdx,brIdx,lang,val){
  if(!testData.domains[domIdx]) return;
  if(!testData.domains[domIdx].branches) testData.domains[domIdx].branches=[];
  if(!testData.domains[domIdx].branches[brIdx]) testData.domains[domIdx].branches[brIdx]={};
  if(lang==='ar') testData.domains[domIdx].branches[brIdx].nameAr=val;
  else testData.domains[domIdx].branches[brIdx].nameEn=val;
  _saveDraft();
}
function recalcTotal(){
  var total=0;document.querySelectorAll('.weight-val').forEach(function(i){total+=Number(i.value)||0;});
  var el=document.getElementById('domainTotal');if(el){el.textContent=total+'%';el.style.color=total===100?'#4ade80':'#f87171';}
}

function buildDomainsGrid(){
  var qty=parseInt(document.getElementById('domQty').value)||0;if(!qty)return;
  // Auto-fix weights to sum to 100% if needed
  var totalW=0;
  for(var k=1;k<=qty;k++) totalW+=Number(document.getElementById('dw'+k)?document.getElementById('dw'+k).value||0:0);
  if(Math.round(totalW)!==100){
    var base=Math.floor(100/qty),rem=100-base*qty;
    for(var k2=1;k2<=qty;k2++){var we=document.getElementById('dw'+k2);if(we)we.value=(k2===1?base+rem:base);}
  }
  var existing=testData.domains.slice();
  testData.domains=[];
  for(var i=0;i<qty;i++){
    var nameAr=document.getElementById('dn'+(i+1))?document.getElementById('dn'+(i+1)).value.trim()||(existing[i]&&existing[i].nameAr)||('مجال '+(i+1)):'مجال '+(i+1);
    var nameEn=document.getElementById('dne'+(i+1))?document.getElementById('dne'+(i+1)).value.trim()||(existing[i]&&existing[i].nameEn)||('Domain '+(i+1)):'Domain '+(i+1);
    var wEl=document.getElementById('dw'+(i+1));
    var weight=wEl?Number(wEl.value)||0:(existing[i]&&existing[i].weight)||Math.floor(100/qty);
    var existingDomain=existing[i]||{};
    testData.domains.push({
      index:i,nameAr:nameAr,nameEn:nameEn,weight:weight,
      time:existingDomain.time||0,qCount:existingDomain.qCount||0,
      questions:existingDomain.questions||[],complete:existingDomain.complete||false,
      iconSrc:existingDomain.iconSrc||'',hasBranches:existingDomain.hasBranches||false,
      branches:existingDomain.branches||[]
    });
  }
  renderDomainsGrid();
}
function getDomainMissing(d){
  var msgs=[];
  if(d.hasBranches){
    if(!d.branches||!d.branches.length){msgs.push('لا توجد فروع محددة');}
    else{d.branches.forEach(function(b,bi){if(!b.time||b.time<=0)msgs.push('فرع '+(bi+1)+': الوقت غير محدد');if(!b.qCount||b.qCount<=0)msgs.push('فرع '+(bi+1)+': عدد الأسئلة غير محدد');else if(!b.questions||b.questions.length<b.qCount)msgs.push('فرع '+(bi+1)+': الأسئلة المضافة '+(b.questions?b.questions.length:0)+' من '+b.qCount);});}
  } else {
    if(!d.time||d.time<=0)msgs.push('وقت المجال غير محدد');
    if(!d.qCount||d.qCount<=0)msgs.push('عدد الأسئلة غير محدد');
    else if(!d.questions||d.questions.length<d.qCount)msgs.push('الأسئلة المضافة '+(d.questions?d.questions.length:0)+' من '+d.qCount);
  }
  return msgs;
}
function renderDomainsGrid(){
  var grid=document.getElementById('domainsGrid');
  grid.className='';
  grid.style.cssText='display:flex;flex-direction:column;gap:24px';
  grid.innerHTML=testData.domains.map(function(d,i){
    var missing=getDomainMissing(d);
    var missingHtml=(!d.complete&&missing.length)?'<div style="margin-top:8px;background:rgba(248,113,113,.1);border:1px solid rgba(248,113,113,.3);border-radius:8px;padding:6px 10px">'+missing.map(function(m){return '<div style="font-size:10px;color:#fca5a5;line-height:1.6">• '+m+'</div>';}).join('')+'</div>':'';

    // ── BRANCH CARDS (ghost/outlined) ──
    var branchCardsHtml='';
    if(d.hasBranches&&d.branches&&d.branches.length){
      var brCards=d.branches.map(function(br,bi){
        var brDone=br.questions&&br.questions.length>=(br.qCount||0)&&br.qCount>0;
        var brWSum=br.questions?br.questions.reduce(function(s,q){return s+(Number(q.score)||0);},0):0;
        var brComplete=brDone&&Math.round(brWSum*100)===Math.round(br.weight*100);
        return '<div onclick="openBranchFromGrid('+i+','+bi+')" class="br-card" style="flex:1;min-width:160px;background:transparent;border:1px dashed #38bdf8;border-radius:18px;padding:16px 14px;text-align:center;box-shadow:none">'
          +'<div style="width:32px;height:32px;border:2px solid #38bdf8;border-radius:50%;display:flex;align-items:center;justify-content:center;margin:0 auto 10px;color:#7dd3fc;font-weight:900;font-size:13px;font-family:Montserrat,sans-serif">'+(bi+1)+'</div>'
          +(br.iconSrc?'<img src="'+br.iconSrc+'" style="width:40px;height:40px;object-fit:contain;border-radius:10px;margin-bottom:8px">':'<div style="font-size:28px;margin-bottom:8px;opacity:.7">📁</div>')
          +'<div style="font-size:12px;font-weight:800;color:#7dd3fc;margin-bottom:3px">'+(br.nameAr||('فرع '+(bi+1)))+'</div>'
          +(br.nameEn?'<div style="font-size:10px;color:rgba(125,211,252,.5);font-family:Montserrat,sans-serif;margin-bottom:6px">'+br.nameEn+'</div>':'')
          +'<div style="display:flex;justify-content:center;gap:4px;flex-wrap:wrap;margin-bottom:6px">'
          +'<span style="border:1px solid rgba(56,189,248,.4);border-radius:8px;padding:2px 7px;font-size:9px;font-family:Montserrat,sans-serif;color:#7dd3fc">'+br.weight+'%</span>'
          +'<span style="border:1px solid rgba(56,189,248,.4);border-radius:8px;padding:2px 7px;font-size:9px;font-family:Montserrat,sans-serif;color:#7dd3fc">⏱'+(br.time||0)+'د</span>'
          +'</div>'
          +'<div style="font-size:10px;font-weight:800;color:'+(brComplete?'#4ade80':'#f87171')+'">'+(brComplete?'✅ مكتمل':'◉ '+(br.questions?br.questions.length:0)+'/'+(br.qCount||0))+'</div>'
          +'</div>';
      }).join('');
      // Connector line + branch cards
      branchCardsHtml='<div style="padding-right:28px;margin-top:0">'
        +'<div style="display:flex;align-items:flex-start;gap:0">'
        // Vertical connector
        +'<div style="width:28px;flex-shrink:0;display:flex;flex-direction:column;align-items:center;padding-top:8px">'
        +'<div style="width:2px;flex:1;background:rgba(56,189,248,.3);min-height:20px"></div>'
        +'</div>'
        // Branch cards row
        +'<div style="flex:1;display:flex;flex-wrap:wrap;gap:10px;padding-bottom:8px">'+brCards+'</div>'
        +'</div></div>';
    }

    // ── DOMAIN CARD (filled/solid) ──
    var domCard='<div onclick="openDomainModal('+i+')" class="dom-card" style="background:linear-gradient(135deg,#f97316,#c2410c);border:3px solid #fb923c;box-shadow:0 0 0 1px #000, 0 6px 24px rgba(249,115,22,.45);border-radius:18px;padding:18px 20px">'
      +'<div style="display:flex;align-items:center;gap:14px">'
      // Badge number - solid filled
      +'<div style="width:38px;height:38px;background:#000;border:2px solid #fb923c;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:900;color:#fb923c;font-size:16px;font-family:Montserrat,sans-serif;flex-shrink:0">'+( i+1)+'</div>'
      // Icon
      +(d.iconSrc?'<img src="'+d.iconSrc+'" style="width:42px;height:42px;object-fit:contain;border-radius:10px;flex-shrink:0">':'')
      // Names
      +'<div style="flex:1;min-width:0">'
      +'<div style="font-size:16px;font-weight:900;color:#9a3412;margin-bottom:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+d.nameAr+'</div>'
      +(d.nameEn?'<div style="font-size:10px;color:rgba(255,255,255,.6);font-family:Montserrat,sans-serif;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">'+d.nameEn+'</div>':'')
      +'</div>'
      // Stats
      +'<div style="display:flex;flex-direction:column;align-items:flex-end;gap:4px;flex-shrink:0">'
      +'<div style="font-size:22px;font-weight:900;color:#c2410c;font-family:Montserrat,sans-serif;line-height:1">'+d.weight+'%</div>'
      +(d.time?'<div style="font-size:10px;color:rgba(255,255,255,.7);font-family:Montserrat,sans-serif">⏱ '+d.time+' د</div>':'')
      +'</div>'
      // Status
      +'<div style="font-size:12px;font-weight:800;color:'+(d.complete?'#fff':'#fed7aa')+';flex-shrink:0;text-align:center;background:'+(d.complete?'rgba(0,0,0,.3)':'rgba(0,0,0,.2)')+';border-radius:20px;padding:4px 10px;border:1px solid '+(d.complete?'rgba(74,222,128,.5)':'rgba(255,255,255,.2)')+';">'
      +(d.complete?'✅ مكتمل':'⚠ ناقص')
      +'</div>'
      +'</div>'
      +missingHtml
      +(d.complete?'<div onclick="event.stopPropagation();previewDomainByIndex('+i+')" class="domain-eye-btn" style="margin-top:10px;display:flex;align-items:center;justify-content:center;gap:6px;background:linear-gradient(135deg,rgba(59,130,246,.25),rgba(29,78,216,.2));border:1.5px solid rgba(59,130,246,.4);border-radius:12px;padding:8px 16px;cursor:pointer;font-size:12px;font-weight:800;color:#93c5fd;font-family:Tajawal,sans-serif;transition:.2s" onmouseover="this.style.background=\'linear-gradient(135deg,rgba(59,130,246,.4),rgba(29,78,216,.3))\'" onmouseout="this.style.background=\'linear-gradient(135deg,rgba(59,130,246,.25),rgba(29,78,216,.2))\'">'+(testData.displayMode===4?'👁 معاينة الورقة البيضاء':'👁 معاينة كلاسيكي')+'</div>':'')
      +'</div>';

    return '<div style="position:relative">'+domCard+branchCardsHtml+'</div>';
  }).join('');

  var allDone=testData.domains.every(function(d){return d.complete;});
  var btn=document.getElementById('approveBtn');
  if(btn){btn.disabled=!allDone;btn.style.opacity=allDone?'1':'.5';btn.style.cursor=allDone?'pointer':'not-allowed';}
}

function openBranchFromGrid(domIdx,branchIdx){
  currentDomainIndex=domIdx;
  currentBranchIndex=-1;
  openDomainModal(domIdx);
  setTimeout(function(){openBranchQuestions(branchIdx);},80);
}

// ============================================================
// DOMAIN MODAL
// ============================================================
function openDomainModal(idx){
  currentDomainIndex=idx; currentBranchIndex=-1; var d=testData.domains[idx];
  document.getElementById('modalDomainBadge').textContent=idx+1;
  document.getElementById('modalDomainNameAr').value=d.nameAr||'';
  document.getElementById('modalDomainNameEn').value=d.nameEn||'';
  document.getElementById('modalDomainWeight').value=d.weight+'%';
  if(d.iconSrc){document.getElementById('domainIconPreview').src=d.iconSrc;document.getElementById('domainIconPreview').style.display='block';document.getElementById('domainIconPlaceholder').style.display='none';}
  else{document.getElementById('domainIconPreview').style.display='none';document.getElementById('domainIconPlaceholder').style.display='flex';}
  // Reset all panels first
  document.getElementById('branchSetupBox').classList.add('hidden');
  document.getElementById('domainSettingsBox').classList.add('hidden');
  document.getElementById('branchQuestionBox').classList.add('hidden');
  // Show/hide branch toggle
  var branchToggleBox=document.getElementById('branchToggleBox');
  if(branchToggleBox) branchToggleBox.style.display='flex';
  document.getElementById('btnNoBranch').style.background='rgba(255,255,255,.05)';
  document.getElementById('btnYesBranch').style.background='rgba(255,255,255,.05)';
  // setBranchMode handles showing the correct panel
  if(d.hasBranches) setBranchMode(true,true,d);
  else setBranchMode(false,true,d);
  document.getElementById('domainModal').classList.remove('hidden');
}
function closeDomainModal(){
  var d=currentDomainIndex>=0?testData.domains[currentDomainIndex]:null;
  if(d&&!d.complete){
    var missing=getDomainMissing(d);
    scConfirm(
      'المجال غير مكتمل','Incomplete Domain',
      'المجال غير مكتمل. النواقص:<br>'+( missing.length?missing.map(function(m){return '• '+m;}).join('<br>'):'معلومات ناقصة')+'<br><br>هل تريد الخروج بدون حفظ؟',
      'Domain incomplete. Exit without saving?','⚠️'
    ).then(function(ok){
      if(ok) document.getElementById('domainModal').classList.add('hidden');
    });
    return;
  }
  document.getElementById('domainModal').classList.add('hidden');
}
function setBranchMode(hasBranches,skip,existingDomain){
  currentDomainHasBranches=hasBranches;
  document.getElementById('btnNoBranch').style.background=!hasBranches?'rgba(74,222,128,.2)':'rgba(255,255,255,.05)';
  document.getElementById('btnYesBranch').style.background=hasBranches?'rgba(250,204,21,.2)':'rgba(255,255,255,.05)';
  if(hasBranches){
    document.getElementById('branchSetupBox').classList.remove('hidden');
    document.getElementById('domainSettingsBox').classList.add('hidden');
    document.getElementById('branchQuestionBox').classList.add('hidden');
    var brCount=existingDomain&&existingDomain.branches?existingDomain.branches.length:2;
    var bcEl=document.getElementById('branchCount');
    if(bcEl) bcEl.value=brCount;
    generateBranchInputs(existingDomain?existingDomain.branches:null);
  } else {
    document.getElementById('branchSetupBox').classList.add('hidden');
    document.getElementById('branchQuestionBox').classList.remove('hidden');
    document.getElementById('domainSettingsBox').classList.remove('hidden');
    var d=testData.domains[currentDomainIndex];
    if(d){
      var timeEl=document.getElementById('modalDomainTime');
      var qEl=document.getElementById('modalQCount');
      if(timeEl) timeEl.value=d.time||'';
      if(qEl) qEl.value=d.qCount||'';
      renderQuestionsList();
    }
  }
}
function generateBranchInputs(existingBranches){
  var bcEl=document.getElementById('branchCount');
  var d=testData.domains[currentDomainIndex];
  var count=bcEl?parseInt(bcEl.value)||0:(existingBranches?existingBranches.length:(d&&d.branches?d.branches.length:2));
  if(!existingBranches&&d&&d.branches) existingBranches=d.branches;
  var cont=document.getElementById('branchInputs');if(!count||!cont)return;
  var totalWeight=d?d.weight:100;
  var baseW=Math.floor(totalWeight/count),remW=totalWeight-baseW*count;
  var html='<p style="font-size:12px;color:rgba(255,255,255,.5);margin-bottom:10px;font-family:Montserrat,sans-serif">Total: '+totalWeight+'% — distribute among branches</p>';
  for(var i=0;i<count;i++){
    var b=existingBranches?existingBranches[i]:{};
    var w=b&&b.weight?b.weight:(i===0?baseW+remW:baseW);
    var qAdded=b&&b.questions?b.questions.length:0;
    var qNeeded=b&&b.qCount?b.qCount:0;
    var qStatusColor=qAdded>=qNeeded&&qNeeded>0?'#4ade80':'#f87171';
    html+='<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.15);border-radius:14px;padding:16px;margin-bottom:12px">'
      +'<div style="display:flex;align-items:center;gap:8px;margin-bottom:12px"><div class="domain-badge">'+(i+1)+'</div><span style="font-size:13px;font-weight:700;color:#FACC15">فرع '+(i+1)+' / Branch '+(i+1)+'</span></div>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:8px">'
      +'<input type="text" id="brName'+i+'" class="wizard-input" placeholder="اسم الفرع" value="'+(b&&b.nameAr?b.nameAr:'')+'">'
      +'<input type="text" id="brNameEn'+i+'" class="wizard-input font-en" placeholder="Branch name" dir="ltr" value="'+(b&&b.nameEn?b.nameEn:'')+'">'
      +'</div>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:8px;margin-bottom:12px">'
      +'<div><label class="wizard-label" style="font-size:9px">الوزن %</label><input type="number" id="brW'+i+'" class="wizard-input font-en" value="'+w+'" min="1" max="'+totalWeight+'" oninput="recalcBranchTotal()"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">الوقت (د)</label><input type="number" id="brT'+i+'" class="wizard-input font-en" placeholder="15" min="1" value="'+(b&&b.time?b.time:'')+'" oninput="updateBranchQStatus('+i+')"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">عدد الأسئلة</label><input type="number" id="brQ'+i+'" class="wizard-input font-en" placeholder="5" min="1" value="'+(b&&b.qCount?b.qCount:'')+'" oninput="updateBranchQStatus('+i+')"></div>'
      +'</div>'
      +'<div style="display:flex;align-items:center;justify-content:space-between;background:rgba(0,0,0,.2);border-radius:10px;padding:10px 14px">'
      +'<div><span style="font-size:12px;color:rgba(255,255,255,.5)">الأسئلة المضافة: </span><span id="brQStatus'+i+'" style="font-size:13px;font-weight:800;color:'+qStatusColor+'">'+qAdded+' / <span id=\"brQNeed'+i+'\">'+qNeeded+'</span></span></div>'
      +'<button type="button" onclick="openBranchQuestions('+i+')" style="background:linear-gradient(135deg,#3b82f6,#1d4ed8);color:white;border:2px solid #93c5fd;border-radius:10px;padding:8px 16px;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700;cursor:pointer;box-shadow:0 4px 12px rgba(59,130,246,.3)">📝 إضافة أسئلة / Add Questions</button>'
      +'</div>'
      +'</div>';
  }
  html+='<div style="display:flex;justify-content:space-between;background:rgba(255,255,255,.05);border-radius:10px;padding:10px 16px;margin-top:4px"><span style="font-size:13px;color:rgba(255,255,255,.5)">المجموع</span><span id="branchTotal" style="font-family:Montserrat,sans-serif;font-weight:800;color:#4ade80">'+totalWeight+'%</span></div>';
  cont.innerHTML=html;
}
function updateBranchQStatus(i){
  var qEl=document.getElementById('brQ'+i);
  var statusEl=document.getElementById('brQStatus'+i);
  var needEl=document.getElementById('brQNeed'+i);
  if(!qEl||!needEl) return;
  var needed=parseInt(qEl.value)||0;
  if(needEl) needEl.textContent=needed;
  var d=testData.domains[currentDomainIndex];
  var added=d&&d.branches&&d.branches[i]&&d.branches[i].questions?d.branches[i].questions.length:0;
  if(statusEl){
    statusEl.style.color=added>=needed&&needed>0?'#4ade80':'#f87171';
    statusEl.innerHTML=added+' / <span id="brQNeed'+i+'">'+needed+'</span>';
  }
}
function recalcBranchTotal(){
  var bcEl=document.getElementById('branchCount');
  var d=testData.domains[currentDomainIndex];
  var count=bcEl?parseInt(bcEl.value)||0:(d&&d.branches?d.branches.length:0);
  var total=0;
  for(var i=0;i<count;i++) total+=Number(document.getElementById('brW'+i)?document.getElementById('brW'+i).value||0:0);
  var el=document.getElementById('branchTotal');
  var tw=d?d.weight:100;
  if(el){el.textContent=total+'%';el.style.color=total===tw?'#4ade80':'#f87171';}
}
function previewDomainIcon(e){
  var file=e.target.files[0];if(!file)return;
  var reader=new FileReader();
  reader.onload=function(ev){document.getElementById('domainIconPreview').src=ev.target.result;document.getElementById('domainIconPreview').style.display='block';document.getElementById('domainIconPlaceholder').style.display='none';if(currentDomainIndex>=0)testData.domains[currentDomainIndex].iconSrc=ev.target.result;};
  reader.readAsDataURL(file);
}
// ══ Auto-generate question slots with weights ══
function autoGenQuestions(){
  var countEl=document.getElementById('modalQCount');
  var count=parseInt(countEl?countEl.value:0)||0;
  if(count<1) return;
  // احسب الوزن: وزن المجال أو الفرع
  var domainWeight=0;
  var d=testData.domains[currentDomainIndex];
  if(currentBranchIndex>=0&&d&&d.branches&&d.branches[currentBranchIndex]){
    domainWeight=parseFloat(d.branches[currentBranchIndex].weight)||0;
  } else if(d){
    domainWeight=parseFloat(d.weight)||0;
  }
  var perQ=domainWeight>0?parseFloat((domainWeight/count).toFixed(2)):0;
  // أنشئ أو حافظ على الأسئلة الموجودة
  var targetList=currentBranchIndex>=0&&d&&d.branches&&d.branches[currentBranchIndex]
    ?d.branches[currentBranchIndex].questions:d?d.questions:[];
  // أضف أسئلة جديدة إذا العدد أكبر
  while(targetList.length<count){
    targetList.push({type:'',stemHtml:'',score:perQ,bloom:'',marking:'auto',lo:'',standard:'',mediaHtml:'',mediaVisible:{}});
  }
  // احذف الزائدة إذا العدد أصغر (مع تأكيد)
  if(targetList.length>count){
    scConfirm('حذف أسئلة','Delete Questions','سيتم حذف '+(targetList.length-count)+' أسئلة إضافية. متأكد؟','This will delete '+(targetList.length-count)+' extra questions. Are you sure?','⚠️').then(function(ok){
      if(!ok)return;
      targetList.splice(count);
      _finishAutoGenQuestions(targetList,count,perQ,domainWeight,d);
    });
    return;
  }
  _finishAutoGenQuestions(targetList,count,perQ,domainWeight,d);
}
function _finishAutoGenQuestions(targetList,count,perQ,domainWeight,d){
  // وزّع الأوزان تلقائياً
  targetList.forEach(function(q,i){
    // آخر سؤال يأخذ الباقي لضمان المجموع الصحيح
    if(i===count-1){
      var used=targetList.slice(0,count-1).reduce(function(s,qq){return s+parseFloat(qq.score||0);},0);
      q.score=parseFloat((domainWeight-used).toFixed(2));
    } else {
      q.score=perQ;
    }
  });
  // حدّث عدد الأسئلة في المجال
  if(currentBranchIndex>=0&&d&&d.branches&&d.branches[currentBranchIndex]){
    d.branches[currentBranchIndex].qCount=count;
  } else if(d){
    d.qCount=count;
  }
  renderQuestionsList();
  // أظهر زر المعاينة
  var pvBtn=document.getElementById('previewDomainBtn');
  if(pvBtn&&count>0) pvBtn.style.display='inline-flex';
}

function previewDomainByIndex(domIdx){
  var d=testData.domains[domIdx];
  if(!d){scWarn('المجال غير موجود','Domain not found');return;}
  try{
    var insArEl2=document.getElementById('insAr');
    var insEnEl2=document.getElementById('insEn');
    if(insArEl2) testData.instructionsAr=insArEl2.innerHTML.trim();
    if(insEnEl2) testData.instructionsEn=insEnEl2.innerHTML.trim();
  }catch(e){}
  var hasBr=d.hasBranches&&d.branches&&d.branches.length;
  var allQuestions=[];
  if(hasBr){
    d.branches.forEach(function(br){(br.questions||[]).forEach(function(q){allQuestions.push(q);});});
  } else {
    allQuestions=(d.questions||[]).slice();
  }
  if(!allQuestions.length){scWarn('لا توجد أسئلة في هذا المجال — أضف أسئلة أولاً','No questions yet — add questions first');return;}
  var modeForPreview=testData.displayMode===4?4:1;
  _savedTestDataForPreview=JSON.parse(JSON.stringify(testData));
  // حافظ على بنية الفروع كما هي بدل دمج كل الأسئلة في قائمة واحدة —
  // بذلك تظهر أيقونة المجال وبداخلها أيقونات الفروع، وكل فرع صفحة منفصلة، وتكتمل الفروع فيكتمل المجال
  var domainForPreview=hasBr
    ?{nameAr:d.nameAr||'المجال',nameEn:d.nameEn||'Domain',weight:d.weight||100,time:parseInt(d.time||30),iconSrc:d.iconSrc||'',hasBranches:true,branches:JSON.parse(JSON.stringify(d.branches)),questions:[]}
    :{nameAr:d.nameAr||'المجال',nameEn:d.nameEn||'Domain',weight:d.weight||100,time:parseInt(d.time||30),iconSrc:d.iconSrc||'',hasBranches:false,questions:JSON.parse(JSON.stringify(allQuestions))};
  testData={
    domains:[domainForPreview],
    selectedSchools:[{name:'معاينة',logo:''}],
    logoSrc:_savedTestDataForPreview.logoSrc||'',
    displayMode:modeForPreview,
    testName:(_savedTestDataForPreview.testName||'معاينة')+' — '+(d.nameAr||'Domain'),
    subject:_savedTestDataForPreview.subject||'',
    grade:_savedTestDataForPreview.grade||'',
    term:_savedTestDataForPreview.term||'',
    year:_savedTestDataForPreview.year||'',
    instructionsAr:_savedTestDataForPreview.instructionsAr||'',
    instructionsEn:_savedTestDataForPreview.instructionsEn||''
  };
  sw_domainIdx=0;sw_branchIdx=-1;sw_qIdx=0;sw_answers={};
  _tryModeActive=true;_previewReturnStep=3;
  applyDisplayMode();
  document.getElementById('studentWindow').style.display='flex';
  var testNameEl=document.getElementById('sw-test-name-header');if(testNameEl)testNameEl.textContent=testData.testName||'معاينة';
  var metaEl=document.getElementById('sw-test-meta-header');if(metaEl){var m2=[];if(testData.subject)m2.push(testData.subject);if(testData.grade)m2.push(testData.grade);metaEl.textContent=m2.join(' — ');}
  var closeBtn=document.getElementById('sw-close-btn');if(closeBtn)closeBtn.textContent='← رجوع للمجالات';
  sw_instructionsConfirmed=false;showStudentInstructions();
}
function saveDomain(){
  // Auto-save branch questions if currently editing a branch
  if(currentBranchIndex>=0){
    var d0=testData.domains[currentDomainIndex];
    if(d0&&d0.branches&&d0.branches[currentBranchIndex]){
      var qs0=_getCurrentQuestions();
      d0.branches[currentBranchIndex].questions=qs0.slice();
    }
  }
  var d=testData.domains[currentDomainIndex];
  d.nameAr=document.getElementById('modalDomainNameAr').value.trim()||d.nameAr;
  d.nameEn=document.getElementById('modalDomainNameEn').value.trim();
  d.hasBranches=currentDomainHasBranches;
  if(currentDomainHasBranches){
    // If branchSetupBox visible, read from DOM inputs first
    var branchSetupVisible=!document.getElementById('branchSetupBox').classList.contains('hidden');
    if(branchSetupVisible){
      var bcEl=document.getElementById('branchCount');
    var count=bcEl?parseInt(bcEl.value)||0:(d.branches?d.branches.length:0);
      if(count<1){scWarn('حدد عدد الفروع أولاً','Enter number of branches first');return;}
      var totalW=0;
      for(var i=0;i<count;i++) totalW+=Number(document.getElementById('brW'+i)?document.getElementById('brW'+i).value||0:0);
      if(Math.round(totalW*100)!==Math.round(d.weight*100)){
        scWarn('مجموع نسب الفروع ('+totalW.toFixed(2)+'%) يجب أن يساوي نسبة المجال ('+d.weight+'%)','Branch weights must equal domain weight ('+d.weight+'%)');return;
        return;
      }
      var oldBranches=d.branches||[];
      d.branches=[];
      for(var j=0;j<count;j++){
        var prev=oldBranches[j]||{};
        d.branches.push({
          nameAr:(document.getElementById('brName'+j)||{}).value||('فرع '+(j+1)),
          nameEn:(document.getElementById('brNameEn'+j)||{}).value||('Branch '+(j+1)),
          weight:Number((document.getElementById('brW'+j)||{}).value||0),
          time:Number((document.getElementById('brT'+j)||{}).value||0),
          qCount:Number((document.getElementById('brQ'+j)||{}).value||0),
          questions:prev.questions||[]
        });
      }
    }
    // Validate all branches have their questions saved
    var branches=d.branches||[];
    var incomplete=branches.filter(function(b){
      return !b.questions||b.questions.length===0;
    }).map(function(b,bi){return b.nameAr||('فرع '+(bi+1));});
    if(incomplete.length>0){
      scWarn('الفروع التالية ليس فيها أسئلة — أضف أسئلتها أولاً:<br>'+incomplete.join('<br>'),'Branches with no questions:<br>'+incomplete.join('<br>'));
      return;
    }
    // Mark complete
    d.complete=branches.length>0&&branches.every(function(b){
      return b.nameAr&&b.weight>0&&b.questions&&b.questions.length>0;
    });
  } else {
    d.time=parseInt(document.getElementById('modalDomainTime').value)||0;
    d.qCount=parseInt(document.getElementById('modalQCount').value)||d.questions?d.questions.length:0;
    var qWTotal=d.questions?d.questions.reduce(function(s,q){return s+(Number(q.score)||0);},0):0;
    d.complete=d.questions&&d.questions.length>0;
  }
  _saveDraft();closeDomainModal();renderDomainsGrid();
}

// ============================================================
// QUESTIONS
// ============================================================
function renderQuestionsList(){
  var qs=_getCurrentQuestions();
  var list=document.getElementById('questionsList');
  var d=testData.domains[currentDomainIndex];
  var parentWeight=currentBranchIndex>=0&&d&&d.branches&&d.branches[currentBranchIndex]?d.branches[currentBranchIndex].weight:d?d.weight:100;
  var totalUsed=qs.reduce(function(s,q){return s+(Number(q.score)||0);},0);
  var remaining=(parentWeight-totalUsed).toFixed(2);
  var neededQ=currentBranchIndex>=0&&d&&d.branches&&d.branches[currentBranchIndex]?d.branches[currentBranchIndex].qCount:d?d.qCount:0;
  var isQComplete=qs.length>0&&qs.length>=neededQ&&Math.round(totalUsed*100)===Math.round(parentWeight*100);
  var meterPct=parentWeight>0?Math.min(100,(totalUsed/parentWeight)*100):0;
  var meterColor=totalUsed>parentWeight?'#f87171':Math.round(totalUsed*100)===Math.round(parentWeight*100)?'#4ade80':'#FACC15';
  // Weight meter bar above list
  var completeBadge=isQComplete
    ?'<div style="background:rgba(74,222,128,.2);border:1px solid #4ade80;border-radius:10px;padding:5px 12px;font-size:11px;font-weight:800;color:#4ade80;font-family:Montserrat,sans-serif;white-space:nowrap">✅ Questions Completed</div>'
    :'<div style="background:rgba(248,113,113,.15);border:1px solid rgba(248,113,113,.4);border-radius:10px;padding:5px 12px;font-size:11px;font-weight:800;color:#f87171;font-family:Montserrat,sans-serif;white-space:nowrap">⚠️ Questions Not Completed</div>';
  var meterHtml='<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:12px;padding:10px 14px;margin-bottom:12px;display:flex;align-items:center;gap:12px">'
    +'<div style="flex:1"><div style="display:flex;justify-content:space-between;margin-bottom:4px"><span style="font-size:11px;color:rgba(255,255,255,.5)">مجموع النسب المستخدمة</span><span style="font-family:Montserrat,sans-serif;font-size:12px;font-weight:800;color:'+meterColor+'">'+totalUsed.toFixed(2)+' / '+parentWeight+'%</span></div><div style="height:8px;border-radius:8px;background:rgba(255,255,255,.1);overflow:hidden"><div style="height:8px;border-radius:8px;background:'+meterColor+';width:'+meterPct+'%;transition:.3s"></div></div></div>'
    +'<div style="font-size:11px;color:'+(Number(remaining)<0?'#f87171':'#4ade80')+';font-weight:700;white-space:nowrap">متبقي: '+remaining+'%</div>'
    +completeBadge+'</div>';
  if(!qs||!qs.length){
    list.innerHTML=meterHtml+'<p style="color:rgba(255,255,255,.3);font-size:13px;text-align:center;padding:16px">لا توجد أسئلة / No questions yet</p>';
    if(currentBranchIndex>=0) updateBranchQStatus(currentBranchIndex);
    return;
  }
  var types={mcq:'MCQ',matching:'Matching / توصيل',ordering:'Word Order / ترتيب',speaking:'Speaking / تحدث',oral:'Oral Reading / قراءة جهرية',listening:'Listening / استماع',reading:'Reading / قراءة',writingskill:'Writing Skill / مهارة الكتابة',truefalse:'True/False / صواب وخطأ',classify:'Classify / تصنيف'};
  list.innerHTML=meterHtml+qs.map(function(q,i){
    return '<div style="display:flex;align-items:center;gap:12px;background:rgba(255,255,255,.05);border-radius:14px;padding:12px;border:1px solid rgba(255,255,255,.1);margin-bottom:8px">'
      +'<div class="domain-badge" style="font-size:12px">Q.'+(i+1)+'</div>'
      +'<div style="flex:1"><div style="font-size:14px;font-weight:700">'+(q.stemText||'Question')+'</div>'
      +'<div style="font-size:11px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">'+(types[q.type]||q.type||'—')+' | '+(q.bloom||'—')+' | '+(q.marking==='auto'?'Auto ✓':'Manual ✍')+'</div></div>'
      +'<div style="background:rgba(250,204,21,.15);border:1px solid rgba(250,204,21,.3);border-radius:10px;padding:4px 10px;font-family:Montserrat,sans-serif;font-size:13px;font-weight:800;color:#FACC15;white-space:nowrap">'+(q.score||0)+'%</div>'
      +'<button onclick="previewSingleQuestion('+i+')" style="color:#93c5fd;background:none;border:none;cursor:pointer;font-size:14px" title="معاينة / Preview">👁</button>'
      +'<button onclick="editQuestion('+i+')" style="color:#FACC15;background:none;border:none;cursor:pointer;font-size:16px">✏️</button>'
      +'<button onclick="deleteQuestion('+i+')" style="color:#f87171;background:none;border:none;cursor:pointer;font-size:16px">🗑</button>'
      +'</div>';
  }).join('');
  if(currentBranchIndex>=0) updateBranchQStatus(currentBranchIndex);
  // أظهر زر المعاينة إذا فيه أسئلة
  var pvBtn=document.getElementById('previewDomainBtn');
  if(pvBtn) pvBtn.style.display=qs.length>0?'inline-flex':'none';
}
function addQuestion(){currentQuestionIndex=-1;currentEditingQ={};openQuestionModal();}
function editQuestion(idx){currentQuestionIndex=idx;currentEditingQ=Object.assign({},_getCurrentQuestions()[idx]);openQuestionModal();}
function deleteQuestion(idx){
  scDanger('حذف السؤال','Delete Question',
    'هل تريد حذف هذا السؤال نهائياً؟',
    'Are you sure you want to permanently delete this question?'
  ).then(function(ok){
    if(!ok) return;
    try{
      var qs=_getCurrentQuestions();
      if(idx>=0&&idx<qs.length){ qs.splice(idx,1); _saveDraft(); renderQuestionsList(); }
    }catch(e){ console.error('deleteQuestion:',e); }
  });
}
function openQuestionModal(){
  var d=testData.domains[currentDomainIndex];
  var qs=_getCurrentQuestions();
  var qNum=currentQuestionIndex>=0?currentQuestionIndex+1:qs.length+1;
  var weight=currentBranchIndex>=0&&d.branches&&d.branches[currentBranchIndex]?d.branches[currentBranchIndex].weight:d.weight;
  var qCount=currentBranchIndex>=0&&d.branches&&d.branches[currentBranchIndex]?d.branches[currentBranchIndex].qCount:d.qCount;
  var suggestedScore=(weight/Math.max(qCount||1,1)).toFixed(2);
  var label=currentBranchIndex>=0&&d.branches&&d.branches[currentBranchIndex]?d.nameAr+' ← '+(d.branches[currentBranchIndex].nameAr||'فرع '+(currentBranchIndex+1)):d.nameAr;
  document.getElementById('qInfoDomain').textContent=label;
  document.getElementById('qInfoNum').textContent='Q.'+qNum;
  document.getElementById('qInfoScore').textContent=suggestedScore+'%';
  document.getElementById('qType').value=currentEditingQ?currentEditingQ.type||'':'';
  document.getElementById('qBloom').value=currentEditingQ?currentEditingQ.bloom||'':'';
  document.getElementById('qMarking').value=currentEditingQ?currentEditingQ.marking||'auto':'auto';
  document.getElementById('qStandard').value=currentEditingQ?currentEditingQ.standard||'':'';
  // Set question score field
  var scoreEl=document.getElementById('qScore');
  if(scoreEl) scoreEl.value=currentEditingQ&&currentEditingQ.score!=null?currentEditingQ.score:suggestedScore;
  populateGLODropdown(currentEditingQ?currentEditingQ.lo||'':'');
  renderQuestionBody();
  updateQWeightMeter();
  document.getElementById('questionModal').classList.remove('hidden');
}
function updateQWeightMeter(){
  var d=testData.domains[currentDomainIndex];
  if(!d) return;
  var parentWeight=currentBranchIndex>=0&&d.branches&&d.branches[currentBranchIndex]?d.branches[currentBranchIndex].weight:d.weight;
  var qs=_getCurrentQuestions();
  // Sum all existing questions except current editing one
  var usedSum=qs.reduce(function(s,q,i){return i===currentQuestionIndex?s:s+(Number(q.score)||0);},0);
  var thisScore=Number(document.getElementById('qScore')?document.getElementById('qScore').value||0:0);
  var totalUsed=usedSum+thisScore;
  var pct=parentWeight>0?Math.min(100,(totalUsed/parentWeight)*100):0;
  var remaining=(parentWeight-usedSum).toFixed(2);
  var bar=document.getElementById('qWeightBar');
  var lbl=document.getElementById('qWeightLabel');
  var infoScore=document.getElementById('qInfoScore');
  if(bar){bar.style.width=pct+'%';bar.style.background=totalUsed>parentWeight?'#f87171':totalUsed===parentWeight?'#4ade80':'#FACC15';}
  if(lbl){lbl.textContent=totalUsed.toFixed(2)+' / '+parentWeight+'%';lbl.style.color=totalUsed>parentWeight?'#f87171':totalUsed===parentWeight?'#4ade80':'#FACC15';}
  if(infoScore){infoScore.textContent='متبقي: '+remaining+'%';infoScore.style.color=remaining<0?'#f87171':'#4ade80';}
}
function closeQuestionModal(){document.getElementById('questionModal').classList.add('hidden');currentEditingQ=null;}
function populateGLODropdown(selectedVal){
  var gradeMap={'KS1/FS1':'FS1','KS1/FS2':'FS2','Grade 1/Year 2':'G1','Grade 2/Year 3':'G2','Grade 3/Year 4':'G3','Grade 4/Year 5':'G4','Grade 5/Year 6':'G5','Grade 6/Year 7':'G6','Grade 7/Year 8':'G7','Grade 8/Year 9':'G8','Grade 9/Year 10':'G9','Grade 10/Year 11':'G10','Grade 11/Year 12':'G11','Grade 12/Year 13':'G12'};
  var gradeId=gradeMap[document.getElementById('grade')?document.getElementById('grade').value||'':'']||'';
  var subjectVal=document.getElementById('subject')?document.getElementById('subject').value||'':'';
  var key=gradeId&&subjectVal?gradeId+'_'+subjectVal:'';
  var glos=key?sbGLOs[key]||[]:[];
  var sel=document.getElementById('qLO'),emptyMsg=document.getElementById('qLOEmpty');
  sel.innerHTML='<option value="">-- اختر --</option>';
  if(!glos.length){if(emptyMsg)emptyMsg.style.display='block';}
  else{if(emptyMsg)emptyMsg.style.display='none';glos.forEach(function(g,i){var code='GLO:'+subjectVal+'/'+gradeId+'/'+String(i+1).padStart(2,'0');var o=document.createElement('option');o.value=code;o.textContent=code+' — '+(g.ar||g.en||'');if(selectedVal===code)o.selected=true;sel.appendChild(o);});}
}

// ============================================================
// RICH TEXT & TOOLBAR
// ============================================================
function execRich(id,cmd,val){document.getElementById(id)?document.getElementById(id).focus():null;document.execCommand(cmd,false,val||null);}
// ١٣. تنظيف HTML من الرموز الغريبة ومشاكل الترميز
function cleanHTML(html){
  if(!html)return'';
  // إزالة أحرف Unicode الغير مرئية والمشكلة
  return html
    .replace(/\u200B/g,'') // zero-width space
    .replace(/\u200C/g,'') // zero-width non-joiner
    .replace(/\u200D/g,'') // zero-width joiner
    .replace(/\uFEFF/g,'') // BOM
    .replace(/[\uE000-\uF8FF]/g,'') // private use area
    .replace(/<font face="[^"]*">/gi,function(m){return m;}) // keep font tags
    .replace(/\s+>/g,'>') // clean whitespace before >
    .trim();
}
function makeToolbar(targetId,dir){
  if(!dir)dir='rtl';
  return '<div class="rtoolbar"><button onclick="execRich(\''+targetId+'\',\'bold\')"><b>B</b></button><button onclick="execRich(\''+targetId+'\',\'italic\')"><i>I</i></button><button onclick="execRich(\''+targetId+'\',\'underline\')"><u>U</u></button><div class="sep"></div><select onchange="execRich(\''+targetId+'\',\'fontSize\',this.value);this.value=\'\'"><option value="">حجم</option><option value="1">10</option><option value="2">12</option><option value="3">14</option><option value="4">16</option><option value="5">18</option><option value="6">22</option><option value="7">26</option></select><select onchange="execRich(\''+targetId+'\',\'fontName\',this.value);this.value=\'\'"><option value="">خط</option><option value="Tajawal">Tajawal</option><option value="Amiri">Amiri</option><option value="Montserrat">Montserrat</option></select><div class="sep"></div><input type="color" onchange="execRich(\''+targetId+'\',\'foreColor\',this.value)"><div class="sep"></div><button onclick="execRich(\''+targetId+'\',\'justifyRight\')">⇥</button><button onclick="execRich(\''+targetId+'\',\'justifyCenter\')">⊟</button><button onclick="execRich(\''+targetId+'\',\'justifyLeft\')">⇤</button><button onclick="execRich(\''+targetId+'\',\'insertUnorderedList\')">•</button></div><div id="'+targetId+'" contenteditable="true" dir="'+dir+'" class="rich-editor" style="min-height:80px" placeholder="اكتب هنا..."></div>';
}
// ══ أدوات تنسيق جماعية — لتوفير الوقت بتطبيق نفس التنسيق على كل الاختيارات/الجمل دفعة واحدة ══
function _groupSelector(prefixes){
  return prefixes.split(',').map(function(p){return '[id^="'+p+'"]';}).join(',');
}
function applyGroupCommand(prefixes,cmd,val){
  document.querySelectorAll(_groupSelector(prefixes)).forEach(function(el){
    var range=document.createRange();range.selectNodeContents(el);
    var sel=window.getSelection();sel.removeAllRanges();sel.addRange(range);
    document.execCommand(cmd,false,val||null);
  });
  window.getSelection().removeAllRanges();
}
function applyGroupStyle(prefixes,prop,val){
  if(!val) return;
  document.querySelectorAll(_groupSelector(prefixes)).forEach(function(el){el.style[prop]=val;});
}
function buildGroupToolbar(prefixes,dir){
  if(!dir) dir='auto';
  return '<div class="rtoolbar" style="background:rgba(96,165,250,.1);border-color:rgba(96,165,250,.35);margin-bottom:10px">'
    +'<span style="font-size:11px;color:#93c5fd;font-family:Tajawal,sans-serif;font-weight:800;margin-left:8px;white-space:nowrap">🎨 تنسيق الكل دفعة واحدة</span>'
    +'<button type="button" onclick="applyGroupCommand(\''+prefixes+'\',\'bold\')"><b>B</b></button>'
    +'<button type="button" onclick="applyGroupCommand(\''+prefixes+'\',\'italic\')"><i>I</i></button>'
    +'<button type="button" onclick="applyGroupCommand(\''+prefixes+'\',\'underline\')"><u>U</u></button>'
    +'<div class="sep"></div>'
    +'<select onchange="applyGroupStyle(\''+prefixes+'\',\'fontSize\',this.value?this.value+\'px\':\'\');this.value=\'\'"><option value="">حجم</option><option value="12">12</option><option value="13">13</option><option value="14">14</option><option value="15">15</option><option value="16">16</option><option value="18">18</option><option value="20">20</option><option value="24">24</option></select>'
    +'<select onchange="applyGroupStyle(\''+prefixes+'\',\'fontFamily\',this.value);this.value=\'\'"><option value="">خط</option><option value="Tajawal">Tajawal</option><option value="Amiri">Amiri</option><option value="Montserrat">Montserrat</option></select>'
    +'<div class="sep"></div>'
    +'<input type="color" title="لون النص للكل" onchange="applyGroupStyle(\''+prefixes+'\',\'color\',this.value)">'
    +'<div class="sep"></div>'
    +'<button type="button" title="محاذاة يمين" onclick="applyGroupStyle(\''+prefixes+'\',\'textAlign\',\'right\')">⇥</button>'
    +'<button type="button" title="توسيط" onclick="applyGroupStyle(\''+prefixes+'\',\'textAlign\',\'center\')">⊟</button>'
    +'<button type="button" title="محاذاة يسار" onclick="applyGroupStyle(\''+prefixes+'\',\'textAlign\',\'left\')">⇤</button>'
    +'</div>';
}

// ============================================================
// MEDIA
// ============================================================
function showMediaMenu(){var area=document.getElementById('mediaPreviewArea');if(!area)return;if(area.querySelector('._mmenu')){area.querySelector('._mmenu').remove();return;}var menu=document.createElement('div');menu.className='_mmenu';menu.style.cssText='display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;padding:10px;background:rgba(0,0,0,.4);border-radius:12px;border:1px solid rgba(255,255,255,.15)';menu.innerHTML='<span style="font-size:11px;color:rgba(255,255,255,.5);width:100%;margin-bottom:4px">اختر مصدر الصورة:</span><label style="display:inline-flex;align-items:center;gap:6px;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);border-radius:10px;padding:7px 14px;cursor:pointer;font-size:13px;color:white;font-family:Tajawal,sans-serif">📁 من الجهاز<input type="file" accept="image/*" style="display:none" onchange="handleMediaFile(event);this.closest(\'._mmenu\').remove()"></label><button onclick="handleMediaURL();this.closest(\'._mmenu\').remove()" style="background:rgba(96,165,250,.2);border:1px solid rgba(96,165,250,.4);border-radius:10px;padding:7px 14px;font-size:13px;color:white;cursor:pointer;font-family:Tajawal,sans-serif">🌐 URL</button><button onclick="this.closest(\'._mmenu\').remove()" style="background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);border-radius:10px;padding:7px 12px;font-size:13px;color:#f87171;cursor:pointer">✕</button>';area.insertBefore(menu,area.firstChild);}
function handleMediaFile(event){var file=event.target.files?event.target.files[0]:null;if(!file)return;var reader=new FileReader();reader.onload=function(ev){insertMediaImg(ev.target.result);};reader.readAsDataURL(file);}
function handleMediaURL(){scPromptText('رابط الصورة','Image URL','https://example.com/image.png','🌐').then(function(url){if(!url||!url.trim())return;insertMediaImg(url.trim());});}
// يستخرج عنصر الميديا النظيف فقط (audio/video/iframe/img) من منطقة المعاينة، بدون أزرار وأدوات المحرر
function extractCleanMediaHtml(){
  var mp=document.getElementById('mediaPreviewArea');
  if(!mp) return '';
  // فيديو (ملف مباشر محفوظ — بدون يوتيوب/iframe لتفادي أخطاء التضمين ومنع رابط خروج من الاختبار)
  var video=mp.querySelector('video');
  if(video){
    return '<video controls src="'+video.getAttribute('src')+'" style="width:100%;border-radius:10px;max-height:280px"></video>';
  }
  // صوت
  var audio=mp.querySelector('audio');
  if(audio){
    return '<audio controls src="'+audio.getAttribute('src')+'" style="width:100%;border-radius:10px;direction:ltr" preload="metadata"></audio>';
  }
  // صورة (مع التحويل/التكبير المُثبَّت)
  var img=mp.querySelector('#mediaImg,img');
  if(img){
    var transform=img.style.transform||'';
    return '<div style="overflow:hidden;border-radius:10px;display:flex;align-items:center;justify-content:center;background:#fafafa"><img src="'+img.getAttribute('src')+'" style="max-width:100%;max-height:280px;display:block;transform:'+transform+'"></div>';
  }
  return '';
}
function attachAudio(){
  // Show audio options
  var mp=document.getElementById('mediaPreviewArea');if(!mp)return;
  var div=document.createElement('div');
  div.style.cssText='margin-top:8px;background:rgba(59,130,246,.1);border:1px solid rgba(59,130,246,.3);border-radius:12px;padding:12px';
  div.innerHTML='<div style="font-size:12px;color:#93c5fd;margin-bottom:8px;font-family:Montserrat,sans-serif">🎙 مصدر الصوت / Audio Source</div>'
    +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
    +'<label style="display:inline-flex;align-items:center;gap:6px;background:rgba(59,130,246,.2);border:1px solid rgba(59,130,246,.4);color:white;border-radius:8px;padding:7px 12px;font-size:12px;cursor:pointer;font-family:Tajawal,sans-serif">📁 من الجهاز<input type="file" accept="audio/*" style="display:none" onchange="loadAudioFile(event)"></label>'
    +'<button onclick="loadAudioURL()" style="background:rgba(99,102,241,.2);border:1px solid rgba(99,102,241,.4);color:white;border-radius:8px;padding:7px 12px;font-size:12px;cursor:pointer;font-family:Tajawal,sans-serif">🔗 رابط URL</button>'
    +'</div>'
    +'<div id="audio-preview-wrap" style="margin-top:10px"></div>';
  mp.innerHTML='';mp.appendChild(div);
}
function loadAudioFile(event){
  var file=event.target.files?event.target.files[0]:null;if(!file)return;
  var reader=new FileReader();
  reader.onload=function(ev){
    _showAudioPreview(ev.target.result);
  };
  reader.readAsDataURL(file);
  event.target.value='';
}
function loadAudioURL(){
  scPromptText('رابط الصوت','Audio URL','https://example.com/audio.mp3','🔗').then(function(url){
    if(!url||!url.trim())return;
    _showAudioPreview(url.trim());
  });
}
function _showAudioPreview(src){
  var wrap=document.getElementById('audio-preview-wrap');
  var mp=document.getElementById('mediaPreviewArea');
  var target=wrap||mp;if(!target)return;
  target.innerHTML='<div style="background:rgba(0,0,0,.3);border-radius:10px;padding:10px">'
    +'<audio controls src="'+src+'" style="width:100%;border-radius:8px;direction:ltr" preload="metadata"></audio>'
    +'<div style="display:flex;gap:6px;margin-top:8px">'
    +'<button onclick="removeAudioPreview(this)" style="background:rgba(239,68,68,.2);border:1px solid rgba(239,68,68,.3);color:#f87171;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer">🗑 إزالة</button>'
    +'<span style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif;align-self:center">✅ جاهز للاستخدام</span>'
    +'</div></div>';
}
function insertMediaImg(src){
  var mp=document.getElementById('mediaPreviewArea');if(!mp)return;
  _imgS={sc:1,rot:0,x:0,y:0,drag:false};
  mp.innerHTML='<div style="background:rgba(0,0,0,.3);border:1px solid rgba(255,255,255,.15);border-radius:12px;padding:10px;margin-top:8px"><div style="display:flex;gap:5px;flex-wrap:wrap;align-items:center;margin-bottom:8px"><button onclick="imgCtrl(\'zi\')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">🔍+</button><button onclick="imgCtrl(\'zo\')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">🔍−</button><button onclick="imgCtrl(\'rr\')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↻</button><button onclick="imgCtrl(\'rl\')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↺</button><button onclick="imgCtrl(\'rs\')" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer">↩</button><input type="range" id="_zsl" min="0.2" max="4" step="0.05" value="1" oninput="_imgS.sc=parseFloat(this.value);_applyT()" style="width:80px;accent-color:#FACC15"><button onclick="confirmMediaFinal()" style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);color:#4ade80;border-radius:7px;padding:3px 10px;font-size:12px;cursor:pointer">✅ تثبيت</button><button onclick="removeMedia()" style="background:rgba(239,68,68,.2);border:1px solid rgba(239,68,68,.4);color:#f87171;border-radius:7px;padding:3px 8px;font-size:12px;cursor:pointer;margin-right:auto">🗑</button></div><div id="_imgC" style="overflow:hidden;border-radius:10px;height:200px;display:flex;align-items:center;justify-content:center;background:#0a0a1a;cursor:grab" onmousedown="_imd(event)" onmousemove="_imm(event)" onmouseup="_ime()" onmouseleave="_ime()"><img id="mediaImg" src="'+src+'" style="max-height:190px;max-width:100%;transform-origin:center;user-select:none;pointer-events:none;display:block" draggable="false" onerror="document.getElementById(\'_imgErr\').style.display=\'flex\';this.style.display=\'none\'"><div id="_imgErr" style="display:none;flex-direction:column;align-items:center;gap:6px;color:rgba(255,255,255,.4)"><div style="font-size:28px">🖼️</div><div style="font-size:11px">تعذّر تحميل الصورة</div></div></div><div style="font-size:10px;color:rgba(255,255,255,.25);margin-top:6px;text-align:center">اسحب لتحريك · السلايدر للتكبير</div></div>';
  setTimeout(function(){var c=document.getElementById('_imgC');if(c)c.addEventListener('wheel',function(e){e.preventDefault();_imgS.sc=Math.max(.2,Math.min(4,_imgS.sc-(e.deltaY>0?.1:-.1)));var s=document.getElementById('_zsl');if(s)s.value=_imgS.sc;_applyT();},{passive:false});},80);
}
function confirmMediaFinal(){
  scConfirm('تثبيت الصورة','Confirm Image','هل هذا الشكل النهائي للصورة؟','Is this the final form of the image?','🖼️').then(function(ok){
    if(!ok) return;
    var img=document.getElementById('mediaImg');if(!img)return;
    var src=img.src;var mp=document.getElementById('mediaPreviewArea');
    var div=document.createElement('div');
    div.style.cssText='margin-top:8px;border-radius:12px;overflow:hidden;background:white;border:1px solid #e2e8f0;position:relative';
    var toolbar=document.createElement('div');
    toolbar.style.cssText='display:flex;gap:6px;padding:8px 10px;background:#f8fafc;border-bottom:1px solid #e2e8f0;align-items:center';
    toolbar.innerHTML='<span style="font-size:11px;color:#64748b;font-family:Montserrat,sans-serif">عرض / Display:</span>';
    var btnC=document.createElement('button');
    btnC.textContent='توسيط / Center';
    btnC.style.cssText='background:#e0f2fe;color:#0369a1;border:none;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif';
    btnC.onclick=function(){setImgDisplay('contain');};
    var btnF=document.createElement('button');
    btnF.textContent='تغطية / Full';
    btnF.style.cssText='background:#dcfce7;color:#15803d;border:none;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif';
    btnF.onclick=function(){setImgDisplay('cover');};
    var btnX=document.createElement('button');
    btnX.textContent='✕';
    btnX.style.cssText='margin-right:auto;background:rgba(239,68,68,.1);border:1px solid rgba(239,68,68,.3);color:#ef4444;border-radius:6px;padding:3px 8px;font-size:11px;cursor:pointer';
    btnX.onclick=function(){removeMedia();};
    toolbar.appendChild(btnC);toolbar.appendChild(btnF);toolbar.appendChild(btnX);
    var wrap=document.createElement('div');
    wrap.id='media-img-wrap';
    wrap.style.cssText='background:white;min-height:100px;display:flex;align-items:center;justify-content:center;padding:8px';
    var imgEl=document.createElement('img');
    imgEl.id='confirmed-img';imgEl.src=src;
    imgEl.style.cssText='max-height:200px;max-width:100%;object-fit:contain;border-radius:6px;display:block;margin:0 auto';
    wrap.appendChild(imgEl);
    div.appendChild(toolbar);div.appendChild(wrap);
    mp.innerHTML='';mp.appendChild(div);
  });
}
function setImgDisplay(mode){
  var img=document.getElementById('confirmed-img');
  var wrap=document.getElementById('media-img-wrap');
  if(!img||!wrap)return;
  if(mode==='cover'){
    img.style.cssText='width:100%;height:200px;object-fit:cover;border-radius:0;display:block';
    wrap.style.padding='0';
  } else {
    img.style.cssText='max-height:200px;max-width:100%;object-fit:contain;border-radius:6px;display:block;margin:0 auto';
    wrap.style.padding='8px';
  }
}
// ══ Question Header Standards — رأس السؤال القياسي لكل نمط ══
var QH_STANDARDS={
  oral:{ar:'اقرأ النصَّ قراءةً جهريةً سَليمَةً .',en:'Read the following text aloud clearly.'},
  speaking:{ar:'تحدَّثْ عن الموضوعِ الآتي بلغةٍ عربيةٍ فصيحةٍ لمدةِ (___) دقائق.',en:'Speak about the following topic in standard Arabic for (___) minutes.'},
  ordering:{ar:'رتِّبِ الكلماتِ الآتيةَ لتُكوِّنَ جملةً مفيدةً صحيحةً.',en:'Arrange the following words to form a correct and meaningful sentence.'},
  listening:{ar:'استمِعْ إلى النصِّ الآتي جيِّدًا، ثم أجِبْ عن الأسئلةِ التي تليه.',en:'Listen carefully to the following audio text, then answer the questions below.'},
  writingskill:{ar:'اكتُبْ فقرةً / موضوعًا عن (___) في ما لا يقلُّ عن (___) كلمة، مستخدمًا مفرداتٍ مناسبة.',en:'Write a paragraph / essay about (___) in no fewer than (___) words, using appropriate vocabulary.'},
  reading:{ar:'اقرأِ النصَّ الآتيَ قراءةً صامتةً متأنِّيةً، ثم أجِبْ عن الأسئلةِ التي تليه.',en:'Read the following text carefully and silently, then answer the questions that follow.'},
  classify:{ar:'صنِّفِ الكلماتِ / العباراتِ الآتيةَ في الجدولِ المناسبِ وَفْقَ معيارِ التصنيفِ المُحدَّد.',en:'Classify the following words / phrases into the appropriate columns according to the given criterion.'},
  truefalse:{ar:'اقرأِ العباراتِ الآتيةَ، ثم ضَعْ علامةَ (✓) أمامَ العبارةِ الصحيحة و(✗) أمامَ الخاطئة.',en:'Read the following statements, then mark (✓) for True and (✗) for False.'},
  matching:{ar:'صِلْ كلَّ كلمةٍ / عبارةٍ في العمودِ (أ) بما يناسبُها من العمودِ (ب).',en:'Match each word / phrase in column (A) with its correct counterpart in column (B).'},
  mcq:{ar:'اختَرِ الإجابةَ الصحيحةَ مما يأتي:',en:'Choose the correct answer from the following:'}
};
function qhGetStandard(type){return QH_STANDARDS[type]||{ar:'',en:''};}
function qhGetActive(q){
  if(q&&q.headerCustom&&(q.headerAr||q.headerEn)) return {ar:q.headerAr||'',en:q.headerEn||''};
  return qhGetStandard(q?q.type:'');
}
// بناء صندوق رأس السؤال — أبيض، عربي يمين بـ "- " وإنجليزي يسار بـ "- "
function buildQHBoxHtml(q){
  var h=qhGetActive(q);
  if(!h.ar&&!h.en) return '';
  var html='<div class="qh-box">';
  if(h.ar) html+='<div class="qh-ar" dir="rtl">- '+h.ar+'</div>';
  if(h.en) html+='<div class="qh-en" dir="ltr">- '+h.en+'</div>';
  html+='</div>';
  return html;
}
// محرر رأس السؤال (صندوق أبيض + أزرار تعديل/حفظ/استعادة)
function buildQHEditorHtml(){
  return '<div style="margin-bottom:14px"><div id="qhBox" class="qh-box" data-custom="0">'
    +'<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:8px;flex-wrap:wrap;gap:8px">'
    +'<label class="wizard-label" style="margin:0;color:#1a1a2e">① رأس السؤال / Question Header</label>'
    +'<div style="display:flex;gap:6px;flex-wrap:wrap">'
    +'<button type="button" onclick="qhToggleEdit()" id="qhEditBtn" style="background:#f1f5f9;border:1px solid #d1d5db;border-radius:8px;padding:5px 11px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif;color:#1a1a2e;font-weight:700">✏️ تعديل / Edit</button>'
    +'<button type="button" onclick="qhSave()" id="qhSaveBtn" style="display:none;background:#FACC15;border:1px solid #FACC15;border-radius:8px;padding:5px 11px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif;color:#1a1a2e;font-weight:800">💾 حفظ / Save</button>'
    +'<button type="button" onclick="qhResetStandard()" id="qhResetBtn" style="display:none;background:#f1f5f9;border:1px solid #d1d5db;border-radius:8px;padding:5px 11px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif;color:#1a1a2e;font-weight:700">↺ القياسي / Standard</button>'
    +'</div></div>'
    +'<div id="qhArLine" class="qh-ar" dir="rtl" contenteditable="false"></div>'
    +'<div id="qhEnLine" class="qh-en" dir="ltr" contenteditable="false"></div>'
    +'<div id="qhNote" style="font-size:10px;color:#94a3b8;margin-top:8px;font-family:Montserrat,sans-serif"></div>'
    +'</div></div>';
}
function qhInitEditor(){
  var box=document.getElementById('qhBox');if(!box)return;
  var typeEl=document.getElementById('qType');
  var type=typeEl?typeEl.value:(currentEditingQ?currentEditingQ.type:'');
  var custom=!!(currentEditingQ&&currentEditingQ.headerCustom);
  var h=custom?{ar:currentEditingQ.headerAr||'',en:currentEditingQ.headerEn||''}:qhGetStandard(type);
  var arEl=document.getElementById('qhArLine'),enEl=document.getElementById('qhEnLine');
  if(arEl){arEl.textContent='- '+(h.ar||'');arEl.contentEditable='false';}
  if(enEl){enEl.textContent='- '+(h.en||'');enEl.contentEditable='false';}
  box.dataset.custom=custom?'1':'0';
  var editBtn=document.getElementById('qhEditBtn'),saveBtn=document.getElementById('qhSaveBtn'),resetBtn=document.getElementById('qhResetBtn');
  if(editBtn)editBtn.style.display='inline-block';
  if(saveBtn)saveBtn.style.display='none';
  if(resetBtn)resetBtn.style.display=custom?'inline-block':'none';
  var note=document.getElementById('qhNote');
  if(note)note.textContent=custom?'✓ رأس مخصص لهذا السؤال / Custom header for this question':'رأس قياسي لهذا النمط — اضغط تعديل للتخصيص / Standard header — click Edit to customize';
}
function qhToggleEdit(){
  var ar=document.getElementById('qhArLine'),en=document.getElementById('qhEnLine');
  if(!ar||!en)return;
  ar.contentEditable='true';en.contentEditable='true';
  document.getElementById('qhEditBtn').style.display='none';
  document.getElementById('qhSaveBtn').style.display='inline-block';
  ar.focus();
}
function qhSave(){
  var ar=document.getElementById('qhArLine'),en=document.getElementById('qhEnLine');
  if(!ar||!en)return;
  var arT=(ar.textContent||'').trim().replace(/^-\s*/,'');
  var enT=(en.textContent||'').trim().replace(/^-\s*/,'');
  if(currentEditingQ){currentEditingQ.headerAr=arT;currentEditingQ.headerEn=enT;currentEditingQ.headerCustom=true;}
  ar.contentEditable='false';en.contentEditable='false';
  ar.textContent='- '+arT;en.textContent='- '+enT;
  document.getElementById('qhEditBtn').style.display='inline-block';
  document.getElementById('qhSaveBtn').style.display='none';
  document.getElementById('qhResetBtn').style.display='inline-block';
  var box=document.getElementById('qhBox');if(box)box.dataset.custom='1';
  var note=document.getElementById('qhNote');if(note)note.textContent='✓ رأس مخصص لهذا السؤال / Custom header for this question';
}
function qhResetStandard(){
  if(currentEditingQ){currentEditingQ.headerCustom=false;currentEditingQ.headerAr='';currentEditingQ.headerEn='';}
  qhInitEditor();
}
function imgCtrl(a){if(a==='zi')_imgS.sc=Math.min(4,parseFloat((_imgS.sc+.2).toFixed(2)));else if(a==='zo')_imgS.sc=Math.max(.2,parseFloat((_imgS.sc-.2).toFixed(2)));else if(a==='rr')_imgS.rot=(_imgS.rot+90)%360;else if(a==='rl')_imgS.rot=(_imgS.rot-90+360)%360;else if(a==='rs'){_imgS.sc=1;_imgS.rot=0;_imgS.x=0;_imgS.y=0;}var s=document.getElementById('_zsl');if(s)s.value=_imgS.sc;_applyT();}
function _applyT(){var img=document.getElementById('mediaImg');if(img)img.style.transform='translate('+_imgS.x+'px,'+_imgS.y+'px) scale('+_imgS.sc+') rotate('+_imgS.rot+'deg)';}
function _imd(e){_imgS.drag=true;_imgS.sx=e.clientX-_imgS.x;_imgS.sy=e.clientY-_imgS.y;var c=document.getElementById('_imgC');if(c)c.style.cursor='grabbing';}
function _imm(e){if(!_imgS.drag)return;_imgS.x=e.clientX-_imgS.sx;_imgS.y=e.clientY-_imgS.sy;_applyT();}
function _ime(){_imgS.drag=false;var c=document.getElementById('_imgC');if(c)c.style.cursor='grab';}
function removeMedia(){var mp=document.getElementById('mediaPreviewArea');if(mp)mp.innerHTML='';_imgS={sc:1,rot:0,x:0,y:0};}
function printAsPDF() {
    var element = document.body; // بقوله خد محتوى الصفحة كلها
    html2pdf(element);           // بقوله حولها لـ PDF
}
// ============================================================
// QUESTION BODY RENDERER
// ============================================================
function renderQuestionBody(){
  var type=document.getElementById('qType').value;
  var area=document.getElementById('questionBodyArea');
  if(!type){area.innerHTML='<p style="color:rgba(255,255,255,.3);text-align:center;font-size:14px;padding:32px 0">اختر نمط السؤال أولاً / Select type first</p>';return;}
  if(currentEditingQ) currentEditingQ.type=type;
  var stemSection=buildQHEditorHtml()+'<div style="margin-bottom:14px"><label class="wizard-label" style="margin-bottom:6px">② المصادر / Media</label><div id="_stemMediaBtns" style="display:flex;gap:6px;flex-wrap:wrap;margin-top:6px"></div><div id="mediaPreviewArea" style="margin-top:8px">'+(currentEditingQ&&currentEditingQ.mediaHtml?currentEditingQ.mediaHtml:'')+'</div></div>';
  setTimeout(function(){qhInitEditor();},30);
  setTimeout(function(){
    var btnWrap=document.getElementById('_stemMediaBtns');if(!btnWrap)return;btnWrap.innerHTML='';
    function mkBtn(label,bg,border,fn){var b=document.createElement('button');b.textContent=label;b.style.cssText='background:'+bg+';border:1px solid '+border+';color:white;border-radius:10px;padding:7px 12px;font-size:12px;cursor:pointer;font-family:Tajawal,sans-serif';b.onclick=fn;return b;}
    btnWrap.appendChild(mkBtn('🖼️ صورة','rgba(255,255,255,.1)','rgba(255,255,255,.2)',showMediaMenu));
    btnWrap.appendChild(mkBtn('🔗 رابط صورة','rgba(255,255,255,.1)','rgba(255,255,255,.2)',handleMediaURL));
    if(['mcq','listening','truefalse'].indexOf(type)>=0){
      btnWrap.appendChild(mkBtn('🎙 صوت جهاز','rgba(59,130,246,.25)','rgba(59,130,246,.5)',attachAudio));
      btnWrap.appendChild(mkBtn('🔗 رابط صوت','rgba(59,130,246,.25)','rgba(59,130,246,.5)',loadAudioURL));
      var bVid=document.createElement('label');bVid.style.cssText='background:rgba(124,58,237,.25);border:1px solid rgba(124,58,237,.5);color:white;border-radius:10px;padding:7px 12px;font-size:12px;cursor:pointer;font-family:Tajawal,sans-serif;display:inline-flex;align-items:center;gap:4px';bVid.innerHTML='🎬 فيديو جهاز<input type="file" accept="video/*" style="display:none">';bVid.querySelector('input').onchange=attachVideoFile;btnWrap.appendChild(bVid);
      btnWrap.appendChild(mkBtn('🔗 رابط فيديو','rgba(124,58,237,.25)','rgba(124,58,237,.5)',attachVideoURL));
    }
  },40);

  if(type==='mcq'){
    var dir=currentEditingQ&&currentEditingQ.mcqDir?currentEditingQ.mcqDir:'vertical';
    var qText=currentEditingQ&&currentEditingQ.questionText?currentEditingQ.questionText:'';
    var qBodyHtml=currentEditingQ&&currentEditingQ.bodyHtml?currentEditingQ.bodyHtml:'';
    var qTextColor=currentEditingQ&&currentEditingQ.qTextColor?currentEditingQ.qTextColor:'#1e3a8a';
    var qTextBg=currentEditingQ&&currentEditingQ.qTextBg?currentEditingQ.qTextBg:'#eff6ff';
    var qTextFont=currentEditingQ&&currentEditingQ.qTextFont?currentEditingQ.qTextFont:'Tajawal';
    var qTextSize=currentEditingQ&&currentEditingQ.qTextSize?currentEditingQ.qTextSize:16;
    var qTextDir=currentEditingQ&&currentEditingQ.qTextDir?currentEditingQ.qTextDir:'rtl';
    // Build the section using DOM to avoid quote issues
    area.innerHTML=stemSection;
    // ③ جسم السؤال / Body
    var bodySection=document.createElement('div');
    bodySection.style.cssText='margin-bottom:14px';
    bodySection.innerHTML='<label class="wizard-label" style="margin-bottom:8px;display:block">③ جسم السؤال / Question Body</label>'+makeToolbar('qBodyEditorMcq','auto');
    area.appendChild(bodySection);
    if(qBodyHtml) setTimeout(function(){var el=document.getElementById('qBodyEditorMcq');if(el)el.innerHTML=qBodyHtml;},60);
    var qtSection=document.createElement('div');
    qtSection.style.cssText='margin-bottom:16px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:14px';
    qtSection.innerHTML='<label class="wizard-label" style="margin-bottom:10px;color:#93c5fd;display:block">③ صياغة السؤال / Question Wording</label>'
      +'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:10px">'
      +'<div><label class="wizard-label" style="font-size:9px">الخط</label>'
      +'<select id="qt-font" class="wizard-input" style="font-size:12px">'
      +'<option value="Tajawal">Tajawal</option>'
      +'<option value="Amiri">Amiri</option>'
      +'<option value="Montserrat">Montserrat</option>'
      +'</select></div>'
      +'<div><label class="wizard-label" style="font-size:9px">حجم / Size</label>'
      +'<input type="number" id="qt-size" class="wizard-input font-en" value="'+qTextSize+'" min="12" max="28" style="font-size:12px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">لون الخط</label>'
      +'<input type="color" id="qt-color" value="'+qTextColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">الخلفية</label>'
      +'<input type="color" id="qt-bg" value="'+qTextBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'</div>'
      +'<div style="display:flex;gap:16px;margin-bottom:8px">'
      +'<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px"><input type="radio" name="qtDir" value="rtl" '+(qTextDir==="rtl"?'checked':'')+' style="accent-color:#FACC15"> RTL ←</label>'
      +'<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:13px"><input type="radio" name="qtDir" value="ltr" '+(qTextDir==="ltr"?'checked':'')+' style="accent-color:#FACC15"> LTR →</label>'
      +'</div>'
      +'<textarea id="qt-text" class="wizard-input" style="min-height:70px;resize:none;font-size:14px" dir="'+qTextDir+'">'+qText+'</textarea>';
    area.appendChild(qtSection);
    // Direction options
    var dirSection=document.createElement('div');
    dirSection.style.cssText='margin-bottom:12px';
    dirSection.innerHTML='<label class="wizard-label">④ اتجاه الاختيارات</label>'
      +'<div style="display:flex;gap:16px;margin-top:6px">'
      +'<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:14px"><input type="radio" name="mcqDir" value="vertical" '+(dir==='vertical'?'checked':'')+' style="accent-color:#FACC15"> عمودي / Vertical</label>'
      +'<label style="display:flex;align-items:center;gap:8px;cursor:pointer;font-size:14px"><input type="radio" name="mcqDir" value="horizontal" '+(dir==='horizontal'?'checked':'')+' style="accent-color:#FACC15"> أفقي / Horizontal</label>'
      +'</div>';
    area.appendChild(dirSection);
    // Options
    var optsSection=document.createElement('div');
    optsSection.innerHTML='<label class="wizard-label">⑤ الاختيارات (✓ للصحيح)</label>'
      +'<div id="mcqOptions" style="margin-top:8px">'+buildMcqOptions(currentEditingQ&&currentEditingQ.options?currentEditingQ.options:['','','',''],currentEditingQ?currentEditingQ.correct:null)+'</div>'
      +'<button onclick="addMcqOption()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px;font-family:Tajawal,sans-serif">+ إضافة اختيار</button>';
    area.appendChild(optsSection);
    // Set font select value
    setTimeout(function(){
      var sel=document.getElementById('qt-font');if(sel)sel.value=qTextFont;
      var el=document.getElementById('qStemEditor');
      if(el&&currentEditingQ&&currentEditingQ.stemHtml)el.innerHTML=currentEditingQ.stemHtml;
    },50);

  } else if(type==='matching'){
    var pairs=currentEditingQ&&currentEditingQ.pairs?currentEditingQ.pairs:[{aHtml:'',aImg:'',bHtml:'',bImg:''},{aHtml:'',aImg:'',bHtml:'',bImg:''},{aHtml:'',aImg:'',bHtml:'',bImg:''}];
    area.innerHTML=stemSection
      +'<div><label class="wizard-label" style="margin-bottom:8px">③ أزواج التوصيل / Matching Pairs</label>'
      +'<p style="font-size:12px;color:rgba(255,255,255,.4);margin-bottom:12px;font-family:Montserrat,sans-serif">لكل زوج: نص A ← نص B (يمكن إضافة صورة لكل منهما)</p>'
      +buildGroupToolbar('mAEdit,mBEdit')
      +'<div id="matchPairsEditor"></div>'
      +'<button onclick="addMatchPairEditor()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px;font-family:Tajawal,sans-serif">➕ إضافة زوج / Add Pair</button>'
      +'<div style="margin-top:14px"><label class="wizard-label" style="margin-bottom:10px;display:block">④ نموذج الإجابة / Answer Key</label>'
      +'<p style="font-size:11px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif;margin-bottom:12px">كل سطر: الطرف الأيسر = الطرف الأيمن — المحاذاة تلقائية</p>'
      +'<div id="matchAnswerKeyRows"></div>'
      +'<button onclick="addMatchAnswerKeyRow()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px;font-family:Tajawal,sans-serif">➕ إضافة زوج إجابة</button>'
      +'<input type="hidden" id="matchAnswerKey" value="'+(currentEditingQ&&currentEditingQ.matchAnswerKey?currentEditingQ.matchAnswerKey:'')+'"></div></div>';
    setTimeout(function(){
      var cont=document.getElementById('matchPairsEditor');if(!cont)return;
      pairs.forEach(function(p,pi){ addMatchPairEditor(p,pi); });
      setTimeout(initMatchAnswerKeyRows,100);
    },50);
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},60);

    } else if(type==='ordering'){
    var orTileBg=currentEditingQ&&currentEditingQ.orTileBg?currentEditingQ.orTileBg:'#1e3a8a';
    var orTileColor=currentEditingQ&&currentEditingQ.orTileColor?currentEditingQ.orTileColor:'#ffffff';
    var orAnsBg=currentEditingQ&&currentEditingQ.orAnsBg?currentEditingQ.orAnsBg:'#f0f7ff';
    var orAnsColor=currentEditingQ&&currentEditingQ.orAnsColor?currentEditingQ.orAnsColor:'#1e3a8a';
    var orFont=currentEditingQ&&currentEditingQ.orFont?currentEditingQ.orFont:'Tajawal';
    var orDir=currentEditingQ&&currentEditingQ.orDir?currentEditingQ.orDir:'rtl';
    var orFontSize=currentEditingQ&&currentEditingQ.orFontSize?currentEditingQ.orFontSize:15;
    // تهيئة مجموعات الترتيب — تدعم تكرار السؤال لأكثر من جملة تحت نفس السؤال
    if(currentEditingQ&&currentEditingQ.orderGroups&&currentEditingQ.orderGroups.length){
      _orderGroups=JSON.parse(JSON.stringify(currentEditingQ.orderGroups));
    } else if(currentEditingQ&&currentEditingQ.words&&currentEditingQ.words.length){
      _orderGroups=[{words:currentEditingQ.words.slice(),answerOrder:currentEditingQ.answerOrder?currentEditingQ.answerOrder.slice():null}];
    } else {
      _orderGroups=[{words:['','','','',''],answerOrder:null}];
    }
    area.innerHTML=stemSection
      +'<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:14px;margin-bottom:16px">'
      +'<div style="font-size:11px;color:#FACC15;font-weight:700;margin-bottom:10px;font-family:Montserrat,sans-serif">🎨 تحكم في التصميم — مرة واحدة لكل الجمل / Design Controls — once for all sentences</div>'
      +'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:10px">'
      +'<div><label class="wizard-label" style="font-size:9px">الخط / Font</label>'
      +'<select id="or-font" class="wizard-input" style="font-size:12px">'
      +'<option value="Tajawal" '+(orFont==="Tajawal"?"selected":"")+'>Tajawal</option>'
      +'<option value="Amiri" '+(orFont==="Amiri"?"selected":"")+'>Amiri</option>'
      +'<option value="Montserrat" '+(orFont==="Montserrat"?"selected":"")+'>Montserrat</option>'
      +'<option value="Arial" '+(orFont==="Arial"?"selected":"")+'>Arial</option>'
      +'</select></div>'
      +'<div><label class="wizard-label" style="font-size:9px">حجم الخط / Size</label>'
      +'<input type="number" id="or-font-size" class="wizard-input font-en" value="'+orFontSize+'" min="10" max="32" style="font-size:12px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">الاتجاه / Dir</label>'
      +'<select id="or-dir" class="wizard-input" style="font-size:12px">'
      +'<option value="rtl" '+(orDir==="rtl"?"selected":"")+'>RTL ←</option>'
      +'<option value="ltr" '+(orDir==="ltr"?"selected":"")+'>LTR →</option>'
      +'</select></div>'
      +'<div></div>'
      +'</div>'
      +'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px">'
      +'<div><label class="wizard-label" style="font-size:9px">خلفية الكلمة / Tile BG</label>'
      +'<input type="color" id="or-tile-bg" value="'+orTileBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">خط الكلمة / Tile Text</label>'
      +'<input type="color" id="or-tile-color" value="'+orTileColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">خلفية الإجابة / Box BG</label>'
      +'<input type="color" id="or-ans-bg" value="'+orAnsBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">خط الإجابة / Box Text</label>'
      +'<input type="color" id="or-ans-color" value="'+orAnsColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:2px"></div>'
      +'</div></div>'
      +'<label class="wizard-label" style="margin-bottom:6px;display:block">③ جمل الترتيب / Ordering Sentences <span style="color:rgba(255,255,255,.4);font-weight:400;font-family:Montserrat,sans-serif;font-size:10px">— كرر السؤال لأكثر من جملة تحت نفس السؤال، والدرجة تُوزَّع تلقائياً</span></label>'
      +'<div id="orderGroupsContainer"></div>'
      +'<button type="button" onclick="addOrderGroup()" style="margin-top:4px;margin-bottom:6px;background:rgba(250,204,21,.12);border:1.5px dashed #FACC15;color:#FACC15;border-radius:12px;padding:10px 20px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;width:100%">➕ تكرار السؤال (جملة جديدة للترتيب) / Repeat Question</button>';
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},50);
    setTimeout(renderOrderGroups,60);

  } else if(type==='speaking'){
    var spInst=currentEditingQ&&currentEditingQ.speakingInst?currentEditingQ.speakingInst:'';
    var spDir=currentEditingQ&&currentEditingQ.speakingDir?currentEditingQ.speakingDir:'rtl';
    var spColor=currentEditingQ&&currentEditingQ.speakingColor?currentEditingQ.speakingColor:'#1a1a2e';
    var spBg=currentEditingQ&&currentEditingQ.speakingBg?currentEditingQ.speakingBg:'#f8fafc';
    var spFont=currentEditingQ&&currentEditingQ.speakingFont?currentEditingQ.speakingFont:'Tajawal';
    area.innerHTML=stemSection
      // ② نص التعليمات مع تحكم كامل
      +'<div style="margin-bottom:14px"><label class="wizard-label">② تعليمات التسجيل / Instructions</label>'
      +'<div style="display:grid;grid-template-columns:1fr 1fr 1fr 1fr;gap:8px;margin-bottom:10px;margin-top:8px;background:rgba(255,255,255,.05);border-radius:12px;padding:10px;border:1px solid rgba(255,255,255,.1)">'
      +'<div><label class="wizard-label" style="font-size:9px">الخط / Font</label>'
      +'<select id="sp-font" class="wizard-input" style="font-size:12px"><option value="Tajawal" '+(spFont==="Tajawal"?"selected":"")+'>Tajawal</option><option value="Amiri" '+(spFont==="Amiri"?"selected":"")+'>Amiri</option><option value="Montserrat" '+(spFont==="Montserrat"?"selected":"")+'>Montserrat</option><option value="Arial" '+(spFont==="Arial"?"selected":"")+'>Arial</option></select></div>'
      +'<div><label class="wizard-label" style="font-size:9px">لون الخط / Color</label>'
      +'<input type="color" id="sp-color" value="'+spColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">خلفية / BG</label>'
      +'<input type="color" id="sp-bg" value="'+spBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);background:transparent;padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">حجم الخط / Size</label>'
      +'<input type="number" id="sp-size" class="wizard-input font-en" value="'+(currentEditingQ&&currentEditingQ.speakingSize?currentEditingQ.speakingSize:15)+'" min="10" max="32" style="font-size:12px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">الاتجاه / Dir</label>'
      +'<select id="sp-dir" class="wizard-input" style="font-size:12px"><option value="rtl" '+(spDir==="rtl"?"selected":"")+'>RTL ←</option><option value="ltr" '+(spDir==="ltr"?"selected":"")+'>LTR →</option></select></div>'
      +'</div>'
      +'<textarea id="speakingInst" class="wizard-input" style="height:80px;resize:none;margin-top:4px" dir="'+spDir+'" placeholder="اكتب التعليمات هنا...">'+spInst+'</textarea></div>'
      // ③ أدوات التسجيل
      +'<div style="background:rgba(255,255,255,.05);border-radius:12px;padding:12px;font-size:12px;color:rgba(255,255,255,.4);border:1px solid rgba(255,255,255,.1);font-family:Montserrat,sans-serif">🎙 Record → ⏹ Stop → 🔊 Listen → ✅ Confirm</div>';
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},50);

  } else if(type==='oral'){
    var orInst=currentEditingQ&&currentEditingQ.speakingInst?currentEditingQ.speakingInst:'';
    area.innerHTML=stemSection
      +'<div style="margin-bottom:14px"><label class="wizard-label">② النص للقراءة الجهرية / Reading Text</label>'
      +makeToolbar('oralTextEditor','auto')
      +'</div>'
      +'<div style="margin-bottom:14px"><label class="wizard-label">③ تعليمات التسجيل / Recording Instructions</label>'
      +'<textarea id="speakingInst" class="wizard-input" style="height:60px;resize:none;margin-top:6px" dir="auto" placeholder="اقرأ النص التالي بصوت واضح...">'+orInst+'</textarea></div>'
      +'<div style="background:rgba(250,204,21,.08);border:1px solid rgba(250,204,21,.2);border-radius:12px;padding:12px;font-size:12px;color:rgba(255,255,255,.6);font-family:Montserrat,sans-serif">📖 الطالب يقرأ النص المعروض ويسجل صوته / Student reads displayed text aloud and records</div>';
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},50);
    if(currentEditingQ&&currentEditingQ.oralText) setTimeout(function(){var el=document.getElementById('oralTextEditor');if(el){el.innerHTML=currentEditingQ.oralText;el.style.minHeight='140px';}},50);
    else setTimeout(function(){var el=document.getElementById('oralTextEditor');if(el)el.style.minHeight='140px';},50);

  } else if(type==='listening'){
    area.innerHTML=stemSection+'<div><label class="wizard-label">③ نوع الإجابة / Answer Type</label><select id="listeningAnswerType" class="wizard-input text-sm" style="margin-top:6px" onchange="renderSubAnswer(\'listeningAnswerBody\',this.value)"><option value="">-- اختر --</option><option value="text" '+(currentEditingQ&&currentEditingQ.ansType==='text'?'selected':'')+'>كتابة / Write</option><option value="mcq" '+(currentEditingQ&&currentEditingQ.ansType==='mcq'?'selected':'')+'>MCQ</option><option value="matching" '+(currentEditingQ&&currentEditingQ.ansType==='matching'?'selected':'')+'>توصيل / Match</option></select><div id="listeningAnswerBody" style="margin-top:12px"></div></div>';
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},50);
    if(currentEditingQ&&currentEditingQ.ansType){
      setTimeout(function(){
        renderSubAnswer('listeningAnswerBody',currentEditingQ.ansType);
        if(currentEditingQ.ansType==='mcq'){
          setTimeout(function(){
            var rt=document.getElementById('rdmcq-text');if(rt&&currentEditingQ.questionText)rt.value=currentEditingQ.questionText;
            var rf=document.getElementById('rdmcq-font');if(rf&&currentEditingQ.qTextFont)rf.value=currentEditingQ.qTextFont;
            var rs=document.getElementById('rdmcq-size');if(rs&&currentEditingQ.qTextSize)rs.value=currentEditingQ.qTextSize;
            var rc=document.getElementById('rdmcq-color');if(rc&&currentEditingQ.qTextColor)rc.value=currentEditingQ.qTextColor;
            var rb=document.getElementById('rdmcq-bg');if(rb&&currentEditingQ.qTextBg)rb.value=currentEditingQ.qTextBg;
            if(currentEditingQ.options&&currentEditingQ.options.length){
              var optEls=document.querySelectorAll('[id^="mcqOptEditor"]');
              optEls.forEach(function(el,i){if(currentEditingQ.options[i]!==undefined)el.innerHTML=currentEditingQ.options[i];});
              if(currentEditingQ.correct!==null&&currentEditingQ.correct!==undefined){
                var radios=document.querySelectorAll('input[name="correctMcq"]');
                if(radios[currentEditingQ.correct])radios[currentEditingQ.correct].checked=true;
              }
            }
          },120);
        }
      },80);
    }

  } else if(type==='truefalse'){
    var tfFont=currentEditingQ&&currentEditingQ.tfFont?currentEditingQ.tfFont:'Tajawal';
    var tfSize=currentEditingQ&&currentEditingQ.tfSize?currentEditingQ.tfSize:15;
    var tfColor=currentEditingQ&&currentEditingQ.tfColor?currentEditingQ.tfColor:'#1a1a2e';
    var tfBg=currentEditingQ&&currentEditingQ.tfBg?currentEditingQ.tfBg:'#f8fafc';
    var tfTrueColor=currentEditingQ&&currentEditingQ.tfTrueColor?currentEditingQ.tfTrueColor:'#15803d';
    var tfFalseColor=currentEditingQ&&currentEditingQ.tfFalseColor?currentEditingQ.tfFalseColor:'#b91c1c';
    area.innerHTML=stemSection
      // أدوات التصميم
      +'<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:12px;margin-bottom:14px">'
      +'<div style="font-size:11px;color:#FACC15;font-weight:700;margin-bottom:10px;font-family:Montserrat,sans-serif">🎨 تصميم الأسئلة / Design</div>'
      +'<div style="display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:10px">'
      +'<div><label class="wizard-label" style="font-size:9px">الخط / Font</label><select id="tf-font" class="wizard-input" style="font-size:12px"><option value="Tajawal" '+(tfFont==="Tajawal"?"selected":"")+'>Tajawal</option><option value="Amiri" '+(tfFont==="Amiri"?"selected":"")+'>Amiri</option><option value="Montserrat" '+(tfFont==="Montserrat"?"selected":"")+'>Montserrat</option></select></div>'
      +'<div><label class="wizard-label" style="font-size:9px">الحجم / Size</label><input type="number" id="tf-size" class="wizard-input font-en" value="'+tfSize+'" min="10" max="28" style="font-size:12px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">لون الخط / Text</label><input type="color" id="tf-color" value="'+tfColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">خلفية السؤال / BG</label><input type="color" id="tf-bg" value="'+tfBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">لون صواب / True</label><input type="color" id="tf-true-color" value="'+tfTrueColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">لون خطأ / False</label><input type="color" id="tf-false-color" value="'+tfFalseColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'</div></div>'
      // الجمل
      +'<div><label class="wizard-label" style="margin-bottom:10px">③ الجمل / Statements</label>'
      +'<p style="font-size:12px;color:rgba(255,255,255,.4);margin-bottom:12px;font-family:Montserrat,sans-serif">أضف جملة واحدة أو أكثر — الطالب يختار صواب أو خطأ لكل منها</p>'
      +'<div id="tf-statements"></div>'
      +'<button onclick="addTFStatement()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px;font-family:Tajawal,sans-serif">➕ إضافة جملة / Add Statement</button></div>';
    var tfStmts=currentEditingQ&&currentEditingQ.statements?currentEditingQ.statements:[];
    setTimeout(function(){
      if(!tfStmts.length) addTFStatement();
      else tfStmts.forEach(function(s,i){ addTFStatement(s.text,s.answer); });
    },50);
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},60);

  } else if(type==='classify'){
    var clBankBg=currentEditingQ&&currentEditingQ.clBankBg?currentEditingQ.clBankBg:'#6366f1';
    var clTableBg=currentEditingQ&&currentEditingQ.clTableBg?currentEditingQ.clTableBg:'#f0f7ff';
    var clFontSize=currentEditingQ&&currentEditingQ.clFontSize?currentEditingQ.clFontSize:14;
    var clTextColor=currentEditingQ&&currentEditingQ.clTextColor?currentEditingQ.clTextColor:'#ffffff';
    area.innerHTML=stemSection
      // أدوات التصميم
      +'<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:12px;margin-bottom:14px">'
      +'<div style="font-size:11px;color:#FACC15;font-weight:700;margin-bottom:10px;font-family:Montserrat,sans-serif">🎨 تصميم / Design</div>'
      +'<div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px">'
      +'<div><label class="wizard-label" style="font-size:9px">حجم الخط / Size</label><input type="number" id="cl-font-size" class="wizard-input font-en" value="'+clFontSize+'" min="10" max="28" style="font-size:12px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">لون النص / Text</label><input type="color" id="cl-text-color" value="'+clTextColor+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">صندوق البنك / Bank</label><input type="color" id="cl-tile-bg" value="'+clBankBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'<div><label class="wizard-label" style="font-size:9px">صندوق الجدول / Table</label><input type="color" id="cl-table-bg" value="'+clTableBg+'" style="width:100%;height:36px;border-radius:8px;cursor:pointer;border:1px solid rgba(255,255,255,.2);padding:2px"></div>'
      +'</div></div>'
      // الأعمدة
      +'<div><label class="wizard-label" style="margin-bottom:10px">③ الأعمدة / Columns</label>'
      +'<div style="display:flex;gap:10px;align-items:center;margin-bottom:12px">'
      +'<input type="number" id="classify-cols" min="2" max="5" value="'+(currentEditingQ&&currentEditingQ.columns?currentEditingQ.columns.length:2)+'" class="wizard-input" style="width:80px" oninput="buildClassifyCols()">'
      +'<span style="font-size:13px;color:rgba(255,255,255,.5)">عدد الأعمدة / Number of columns</span></div>'
      +'<div id="classify-cols-container"></div>'
      +'<label class="wizard-label" style="margin-top:16px;margin-bottom:8px">④ الكلمات/الجمل للتصنيف / Items to classify</label>'
      +'<div id="classify-items-container"></div>'
      +'<button onclick="addClassifyItem()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px;font-family:Tajawal,sans-serif">➕ إضافة مفردة / Add Item</button></div>';
    setTimeout(function(){
      buildClassifyCols();
      if(currentEditingQ&&currentEditingQ.items) currentEditingQ.items.forEach(function(item){ addClassifyItem(item.text,item.correctCol); });
      else { addClassifyItem(); addClassifyItem(); }
    },50);
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},60);

  } else if(type==='writing'){
    var wMode=currentEditingQ&&currentEditingQ.writingMode?currentEditingQ.writingMode:'free';
    var wMin=currentEditingQ&&currentEditingQ.writingMin?currentEditingQ.writingMin:'';
    var wMax=currentEditingQ&&currentEditingQ.writingMax?currentEditingQ.writingMax:'';
    // Writing = مثل Reading تماماً مع passage + answer stem + type
    area.innerHTML=
      // ① رأس السؤال
      stemSection
      // ② المصادر
      +'<div style="margin-bottom:14px"><label class="wizard-label">② المصادر / Media</label>'
      +'<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:6px">'
      +'<button onclick="showMediaMenu()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🖼️ صورة</button>'
      +'<button onclick="attachAudio()" style="background:rgba(59,130,246,.25);border:1px solid rgba(59,130,246,.5);color:white;border-radius:10px;padding:8px 14px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🎙 صوت</button>'
      +'<label style="display:inline-flex;align-items:center;gap:6px;background:rgba(124,58,237,.25);border:1px solid rgba(124,58,237,.5);color:white;border-radius:10px;padding:8px 14px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🎬 فيديو<input type="file" accept="video/*" style="display:none" onchange="attachVideoFile(event)"></label>'
      +'</div><div id="mediaPreviewArea" style="margin-top:8px">'+(currentEditingQ&&currentEditingQ.mediaHtml?currentEditingQ.mediaHtml:'')+'</div></div>'
      // ③ النص / Passage
      +'<div style="margin-bottom:14px"><label class="wizard-label" style="margin-bottom:8px">③ النص / Passage</label>'
      +makeToolbar('writingPassage','auto')
      +'<div id="writingPassage" contenteditable="true" dir="auto" class="rich-editor" style="min-height:120px" placeholder="النص / Passage...">'+(currentEditingQ&&currentEditingQ.passageHtml?currentEditingQ.passageHtml:'')+'</div></div>'
      // ④ رأس سؤال الإجابة
      +'<div style="margin-bottom:14px"><label class="wizard-label" style="margin-bottom:8px">④ رأس سؤال الإجابة / Answer Stem</label>'+makeToolbar('writingAnswerStem','auto')+'</div>'
      // ⑤ نوع الكتابة
      +'<div><label class="wizard-label" style="margin-bottom:10px">⑤ نوع الكتابة / Writing Type</label>'
      +'<div style="display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px">'
      +'<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:14px"><input type="radio" name="writingMode" value="free" '+(wMode==='free'?'checked':'')+' style="accent-color:#FACC15"> حر / Free</label>'
      +'<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:14px"><input type="radio" name="writingMode" value="words" '+(wMode==='words'?'checked':'')+' style="accent-color:#FACC15"> بالكلمات / Words</label>'
      +'<label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:14px"><input type="radio" name="writingMode" value="sentences" '+(wMode==='sentences'?'checked':'')+' style="accent-color:#FACC15"> بالجمل / Sentences</label>'
      +'</div>'
      +'<div style="display:flex;gap:12px;flex-wrap:wrap">'
      +'<div><label class="wizard-label" style="font-size:10px">Min</label><input type="number" id="writing-min" class="wizard-input font-en" value="'+wMin+'" placeholder="0" style="width:80px" min="0"></div>'
      +'<div><label class="wizard-label" style="font-size:10px">Max</label><input type="number" id="writing-max" class="wizard-input font-en" value="'+wMax+'" placeholder="∞" style="width:80px" min="0"></div>'
      +'</div></div>';
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},60);
    if(currentEditingQ&&currentEditingQ.answerStemHtml) setTimeout(function(){var el=document.getElementById('writingAnswerStem');if(el)el.innerHTML=currentEditingQ.answerStemHtml;},60);

  } else if(type==='reading'||type==='writingskill'){
    area.innerHTML=buildQHEditorHtml()+'<div style="margin-bottom:16px"><label class="wizard-label">② المصادر</label><div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:6px"><button onclick="showMediaMenu()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🖼️ صورة</button><button onclick="handleMediaURL()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🔗 رابط</button><button onclick="attachAudio()" style="background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:white;border-radius:10px;padding:8px 14px;font-size:13px;cursor:pointer;font-family:Tajawal,sans-serif">🎙 صوت</button></div><div id="mediaPreviewArea" style="margin-top:8px">'+(currentEditingQ&&currentEditingQ.mediaHtml?currentEditingQ.mediaHtml:'')+'</div></div><div style="margin-bottom:16px"><label class="wizard-label" style="margin-bottom:8px">③ النص المقروء / Passage</label><div class="rtoolbar"><button onclick="execRich(\'passageEditor\',\'bold\')"><b>B</b></button><button onclick="execRich(\'passageEditor\',\'italic\')"><i>I</i></button><button onclick="execRich(\'passageEditor\',\'underline\')"><u>U</u></button><div class="sep"></div><select onchange="execRich(\'passageEditor\',\'fontSize\',this.value);this.value=\'\'"><option value="">حجم</option><option value="1">10</option><option value="2">12</option><option value="3">14</option><option value="4">16</option><option value="5">18</option><option value="6">22</option><option value="7">26</option></select><select onchange="execRich(\'passageEditor\',\'fontName\',this.value);this.value=\'\'"><option value="">خط</option><option value="Tajawal">Tajawal</option><option value="Amiri">Amiri</option><option value="Montserrat">Montserrat</option></select><div class="sep"></div><input type="color" onchange="execRich(\'passageEditor\',\'foreColor\',this.value)"><div class="sep"></div><button onclick="execRich(\'passageEditor\',\'justifyRight\')">⇥</button><button onclick="execRich(\'passageEditor\',\'justifyCenter\')">⊟</button><button onclick="execRich(\'passageEditor\',\'justifyLeft\')">⇤</button><button onclick="execRich(\'passageEditor\',\'insertUnorderedList\')">•</button><button onclick="execRich(\'passageEditor\',\'insertOrderedList\')">1.</button></div><div id="passageEditor" contenteditable="true" dir="auto" class="rich-editor" style="min-height:140px" placeholder="النص المقروء...">'+(currentEditingQ&&currentEditingQ.passageHtml?currentEditingQ.passageHtml:'')+'</div></div><div style="margin-bottom:16px"><label class="wizard-label" style="margin-bottom:8px">④ رأس سؤال الإجابة</label>'+makeToolbar('answerStemEditor','auto')+'</div><div><label class="wizard-label">⑤ نوع الإجابة</label><select id="readingAnswerType" class="wizard-input text-sm" style="margin-top:6px" onchange="renderSubAnswer(\'readingAnswerBody\',this.value)"><option value="">-- اختر --</option><option value="text" '+(currentEditingQ&&currentEditingQ.ansType==='text'?'selected':'')+'>كتابة / Draw</option><option value="mcq" '+(currentEditingQ&&currentEditingQ.ansType==='mcq'?'selected':'')+'>MCQ</option><option value="matching" '+(currentEditingQ&&currentEditingQ.ansType==='matching'?'selected':'')+'>توصيل</option></select><div id="readingAnswerBody" style="margin-top:12px"></div></div>';
    if(currentEditingQ&&currentEditingQ.stemHtml) setTimeout(function(){var el=document.getElementById('qStemEditor');if(el)el.innerHTML=currentEditingQ.stemHtml;},50);
    if(currentEditingQ&&currentEditingQ.answerStemHtml){setTimeout(function(){var el2=document.getElementById('answerStemEditor');if(el2)el2.innerHTML=currentEditingQ.answerStemHtml;},60);}
    if(currentEditingQ&&currentEditingQ.ansType){
      setTimeout(function(){
        renderSubAnswer('readingAnswerBody',currentEditingQ.ansType);
        // استرجاع MCQ options
        if(currentEditingQ.ansType==='mcq'){
          setTimeout(function(){
            // نص السؤال وخصائصه
            var rt=document.getElementById('rdmcq-text');if(rt&&currentEditingQ.questionText)rt.value=currentEditingQ.questionText;
            var rf=document.getElementById('rdmcq-font');if(rf&&currentEditingQ.qTextFont)rf.value=currentEditingQ.qTextFont;
            var rs=document.getElementById('rdmcq-size');if(rs&&currentEditingQ.qTextSize)rs.value=currentEditingQ.qTextSize;
            var rc=document.getElementById('rdmcq-color');if(rc&&currentEditingQ.qTextColor)rc.value=currentEditingQ.qTextColor;
            var rb=document.getElementById('rdmcq-bg');if(rb&&currentEditingQ.qTextBg)rb.value=currentEditingQ.qTextBg;
            var rd=document.querySelector('input[name="rdmcqDir"][value="'+(currentEditingQ.qTextDir||'rtl')+'"]');if(rd)rd.checked=true;
            // الخيارات
            if(currentEditingQ.options&&currentEditingQ.options.length){
              var optEls=document.querySelectorAll('[id^="mcqOptEditor"]');
              optEls.forEach(function(el,i){if(currentEditingQ.options[i]!==undefined)el.innerHTML=currentEditingQ.options[i];});
              if(currentEditingQ.correct!==null&&currentEditingQ.correct!==undefined){
                var radios=document.querySelectorAll('input[name="correctMcq"]');
                if(radios[currentEditingQ.correct])radios[currentEditingQ.correct].checked=true;
              }
            }
          },120);
        }
      },80);
    }
  }
}

function renderSubAnswer(containerId,type){
  var body=document.getElementById(containerId);if(!body)return;
  if(type==='text'){body.innerHTML='<div style="background:rgba(255,255,255,.05);border-radius:12px;padding:12px;border:1px solid rgba(255,255,255,.1)"><p style="font-size:12px;color:rgba(255,255,255,.5);margin-bottom:8px">الطالب يستطيع:</p><div style="display:flex;gap:16px;flex-wrap:wrap"><label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:14px"><input type="checkbox" checked style="accent-color:#FACC15"> ⌨️ كتابة</label><label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:14px"><input type="checkbox" checked style="accent-color:#FACC15"> 🎨 رسم</label><label style="display:flex;align-items:center;gap:6px;cursor:pointer;font-size:14px"><input type="checkbox" checked style="accent-color:#FACC15"> 📷 صورة</label></div></div>';}
  else if(type==='mcq'){
    body.innerHTML=
      // الاختيارات فقط — بدون صياغة سؤال مكررة
      '<div><label class="wizard-label" style="margin-bottom:6px">① الاختيارات (✓ للصحيح) / Options</label>'
      +'<div id="mcqOptions" style="margin-top:8px">'+buildMcqOptions(['','','',''],null)+'</div>'
      +'<button onclick="addMcqOption()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px;font-family:Tajawal,sans-serif">+ إضافة اختيار</button></div>';
  }
  else if(type==='matching'){var pairs=[{aHtml:'',bHtml:''},{aHtml:'',bHtml:''},{aHtml:'',bHtml:''}];body.innerHTML=buildSimpleMatchEditor(pairs);}
  else body.innerHTML='';
}

function buildSimpleMatchEditor(pairs){
  return buildGroupToolbar('mAEdit,mBEdit')+'<div class="grid grid-cols-1 md:grid-cols-2 gap-4 mt-2">'+pairs.map(function(p,i){return '<div style="background:rgba(255,255,255,.05);border-radius:12px;padding:12px;border:1px solid rgba(255,255,255,.1)"><div style="display:grid;grid-template-columns:1fr 1fr;gap:8px"><div><label class="wizard-label" style="font-size:9px">A '+( i+1)+'</label><div id="mAEdit'+i+'" contenteditable="true" dir="auto" class="rich-editor" style="min-height:50px" placeholder="نص A..."></div></div><div><label class="wizard-label" style="font-size:9px">B '+(i+1)+'</label><div id="mBEdit'+i+'" contenteditable="true" dir="auto" class="rich-editor" style="min-height:50px" placeholder="نص B..."></div></div></div></div>';}).join('')+'</div><button onclick="addMatchPair()" style="margin-top:8px;color:#FACC15;background:none;border:none;cursor:pointer;font-size:13px">+ إضافة زوج</button>';
}

function buildWordTiles(words){
  return words.map(function(w,i){return '<div style="background:rgba(250,204,21,.15);border:2px solid rgba(250,204,21,.4);border-radius:10px;padding:8px 14px;font-size:14px;font-weight:700;min-width:60px;text-align:center;cursor:text" contenteditable="true" id="wordTile'+i+'" dir="auto">'+(w||'')+'</div>';}).join('');
}
function buildWordInputs(words){
  return buildWordTiles(words);
}
function updateWordInputs(){
  var count=parseInt(document.getElementById('wordCount').value)||0;
  var cont=document.getElementById('wordInputs');if(!cont)return;
  var existing=[];
  for(var i=0;i<20;i++){var el=document.getElementById('wordTile'+i);if(el)existing.push(el.textContent||'');else break;}
  var words=[];for(var j=0;j<Math.max(2,Math.min(15,count));j++) words.push(existing[j]||'');
  cont.innerHTML=buildWordTiles(words);
  setTimeout(function(){buildOrderAnswerKeyBoxes();},50);
}
// كل صندوق = موضع في الترتيب، يختار المحرر أي كلمة من القائمة تقع هنا
function buildOrderAnswerKeyBoxes(savedOrder){
  var boxesCont=document.getElementById('orderAnswerKeyBoxes');
  if(!boxesCont) return;
  var words=[];
  for(var i=0;i<20;i++){var el=document.getElementById('wordTile'+i);if(el)words.push(el.textContent||'');else break;}
  if(!words.length){boxesCont.innerHTML='<span style="font-size:12px;color:rgba(255,255,255,.4)">أدخل الكلمات أولاً / Enter words first</span>';return;}
  var order=savedOrder&&savedOrder.length===words.length?savedOrder:words.map(function(_,i){return i;});
  boxesCont.innerHTML=words.map(function(_,posIdx){
    var selVal=order[posIdx]!==undefined?order[posIdx]:posIdx;
    return '<div style="display:flex;flex-direction:column;align-items:center;gap:4px">'
      +'<div style="font-size:10px;font-weight:800;color:#4ade80;font-family:Montserrat,sans-serif">الموضع '+(posIdx+1)+'</div>'
      +'<select id="orderKeyBox'+posIdx+'" class="wizard-input" style="font-size:12px;min-width:90px;text-align:center">'
      +words.map(function(w,wi){return '<option value="'+wi+'" '+(parseInt(selVal)===wi?'selected':'')+'>'+(w||'(كلمة '+(wi+1)+')')+'</option>';}).join('')
      +'</select></div>';
  }).join('');
}

// ══ سؤال الترتيب — دعم تكرار الجملة داخل نفس السؤال (مجموعات متعددة) ══
var _orderGroups=[{words:['','','','',''],answerOrder:null}];
function renderOrderGroups(){
  var cont=document.getElementById('orderGroupsContainer');if(!cont)return;
  cont.innerHTML=_orderGroups.map(function(g,gi){return buildOrderGroupBlockHtml(gi,g);}).join('');
  _orderGroups.forEach(function(g,gi){
    (g.words||[]).forEach(function(w,i){var el=document.getElementById('wordTile_'+gi+'_'+i);if(el&&w)el.textContent=w;});
    var hasWords=(g.words||[]).some(function(w){return w&&w.trim();});
    if(hasWords){ buildOrderAnswerKeyBoxesForGroup(gi); }
  });
}
function buildOrderGroupBlockHtml(gi,g){
  var words=g.words&&g.words.length?g.words:['','',''];
  var delBtn=_orderGroups.length>1?'<button type="button" onclick="removeOrderGroup('+gi+')" style="background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);color:#f87171;border-radius:8px;padding:4px 10px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif">🗑 حذف الجملة</button>':'';
  return '<div class="order-group-block" id="orderGroupBlock'+gi+'" style="background:rgba(255,255,255,.04);border:1.5px solid rgba(96,165,250,.25);border-radius:14px;padding:14px;margin-bottom:12px">'
    +'<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:10px">'
    +'<span style="font-size:13px;font-weight:800;color:#93c5fd">📝 الجملة '+(gi+1)+' / Sentence '+(gi+1)+'</span>'
    +delBtn
    +'</div>'
    +'<div style="display:flex;align-items:center;gap:12px;margin-bottom:10px">'
    +'<label class="wizard-label" style="font-size:10px;margin-bottom:0">عدد الكلمات</label>'
    +'<input type="number" id="wordCount_'+gi+'" class="wizard-input" min="2" max="15" value="'+words.length+'" oninput="updateOrderWordCount('+gi+')" style="width:90px">'
    +'</div>'
    +'<div id="wordInputs_'+gi+'" style="display:flex;flex-wrap:wrap;gap:8px;padding:12px;background:rgba(255,255,255,.05);border-radius:12px;border:1px solid rgba(255,255,255,.15);margin-bottom:10px">'+buildOrderWordTiles(gi,words)+'</div>'
    +'<div style="text-align:center;margin-bottom:12px">'
    +'<button type="button" onclick="confirmOrderWords('+gi+')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:2px solid #86efac;border-radius:10px;padding:8px 22px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">✅ تأكيد الكلمات / Confirm Words</button>'
    +'</div>'
    +'<label class="wizard-label" style="font-size:11px;margin-bottom:6px;display:block">نموذج الإجابة — الترتيب الصحيح / Answer Key</label>'
    +'<div id="orderAnswerKeyBoxes_'+gi+'" style="display:flex;flex-wrap:wrap;gap:8px;padding:12px;background:rgba(34,197,94,.06);border-radius:12px;border:1.5px dashed rgba(34,197,94,.3)"><span style="font-size:11px;color:rgba(255,255,255,.4)">اضغط "تأكيد الكلمات" لبناء نموذج الإجابة / Press "Confirm Words" to build the answer key</span></div>'
    +'</div>';
}
function buildOrderWordTiles(gi,words){
  return words.map(function(w,i){return '<div style="background:rgba(250,204,21,.15);border:2px solid rgba(250,204,21,.4);border-radius:10px;padding:8px 14px;font-size:14px;font-weight:700;min-width:60px;text-align:center;cursor:text" contenteditable="true" id="wordTile_'+gi+'_'+i+'" dir="auto">'+(w||'')+'</div>';}).join('');
}
function updateOrderWordCount(gi){
  var countEl=document.getElementById('wordCount_'+gi);if(!countEl)return;
  var count=parseInt(countEl.value)||0;
  var cont=document.getElementById('wordInputs_'+gi);if(!cont)return;
  var existing=[];
  for(var i=0;i<20;i++){var el=document.getElementById('wordTile_'+gi+'_'+i);if(el)existing.push(el.textContent||'');else break;}
  var words=[];for(var j=0;j<Math.max(2,Math.min(15,count));j++) words.push(existing[j]||'');
  cont.innerHTML=buildOrderWordTiles(gi,words);
  var akCont=document.getElementById('orderAnswerKeyBoxes_'+gi);
  if(akCont) akCont.innerHTML='<span style="font-size:11px;color:rgba(255,255,255,.4)">اضغط "تأكيد الكلمات" لبناء نموذج الإجابة / Press "Confirm Words" to build the answer key</span>';
}
function confirmOrderWords(gi){
  var words=[];
  for(var i=0;i<20;i++){var el=document.getElementById('wordTile_'+gi+'_'+i);if(el)words.push(el.textContent||'');else break;}
  if(!words.length||!words.some(function(w){return w.trim();})){scWarn('أدخل الكلمات أولاً','Please enter the words first');return;}
  if(!_orderGroups[gi]) _orderGroups[gi]={};
  _orderGroups[gi].words=words;
  buildOrderAnswerKeyBoxesForGroup(gi);
}
function buildOrderAnswerKeyBoxesForGroup(gi){
  var boxesCont=document.getElementById('orderAnswerKeyBoxes_'+gi);
  if(!boxesCont) return;
  var words=[];
  for(var i=0;i<20;i++){var el=document.getElementById('wordTile_'+gi+'_'+i);if(el)words.push(el.textContent||'');else break;}
  if(!words.length||!words.some(function(w){return w.trim();})){boxesCont.innerHTML='<span style="font-size:11px;color:rgba(255,255,255,.4)">اضغط "تأكيد الكلمات" لبناء نموذج الإجابة / Press "Confirm Words" to build the answer key</span>';return;}
  var saved=_orderGroups[gi]&&_orderGroups[gi].answerOrder&&_orderGroups[gi].answerOrder.length===words.length?_orderGroups[gi].answerOrder:null;
  var order=saved||words.map(function(_,i){return i;});
  boxesCont.innerHTML=words.map(function(_,posIdx){
    var selVal=order[posIdx]!==undefined?order[posIdx]:posIdx;
    return '<div style="display:flex;flex-direction:column;align-items:center;gap:4px">'
      +'<div style="font-size:10px;font-weight:800;color:#4ade80;font-family:Montserrat,sans-serif">الموضع '+(posIdx+1)+'</div>'
      +'<select id="orderKeyBox_'+gi+'_'+posIdx+'" class="wizard-input" style="font-size:12px;min-width:90px;text-align:center">'
      +words.map(function(w,wi){return '<option value="'+wi+'" '+(parseInt(selVal)===wi?'selected':'')+'>'+(w||'(كلمة '+(wi+1)+')')+'</option>';}).join('')
      +'</select></div>';
  }).join('');
}
function addOrderGroup(){
  _orderGroups.push({words:['','','','',''],answerOrder:null});
  renderOrderGroups();
}
function removeOrderGroup(gi){
  if(_orderGroups.length<=1)return;
  scConfirm('حذف الجملة','Delete Sentence','هل تريد حذف هذه الجملة من سؤال الترتيب؟','Delete this sentence from the ordering question?','🗑').then(function(ok){
    if(!ok)return;
    _orderGroups.splice(gi,1);
    renderOrderGroups();
  });
}

function buildMcqOptions(opts,correct){
  var labels=['A','B','C','D','E','F'];
  var toolbar=buildGroupToolbar('mcqOptEditor');
  var html=toolbar+opts.map(function(o,i){return '<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:12px;margin-bottom:8px"><div style="display:flex;align-items:center;gap:8px;margin-bottom:8px"><div class="domain-badge" style="font-size:11px;flex-shrink:0">'+(labels[i]||i+1)+'</div><span style="font-size:12px;color:rgba(255,255,255,.5);flex:1">Option '+(labels[i]||i+1)+'</span><label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#4ade80;cursor:pointer"><input type="radio" name="correctMcq" value="'+i+'" '+(correct==i?'checked':'')+' style="accent-color:#22c55e;width:16px;height:16px"> ✓ صحيح</label></div><div id="mcqOptEditor'+i+'" contenteditable="true" dir="auto" class="rich-editor" style="min-height:44px" placeholder="نص الاختيار..."></div></div>';}).join('');
  setTimeout(function(){opts.forEach(function(o,i){if(!o)return;var el=document.getElementById('mcqOptEditor'+i);if(el&&!el.innerHTML.trim())el.innerHTML=o;});},100);
  return html;
}
function addMcqOption(){
  var existing=document.querySelectorAll('[id^="mcqOptEditor"]').length;
  var labels=['A','B','C','D','E','F'];
  var container=document.getElementById('mcqOptions');if(!container)return;
  var div=document.createElement('div');
  div.style.cssText='background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:12px;margin-bottom:8px';
  div.innerHTML='<div style="display:flex;align-items:center;gap:8px;margin-bottom:8px"><div class="domain-badge" style="font-size:11px;flex-shrink:0">'+(labels[existing]||existing+1)+'</div><span style="font-size:12px;color:rgba(255,255,255,.5);flex:1">Option '+(labels[existing]||existing+1)+'</span><label style="display:flex;align-items:center;gap:4px;font-size:12px;color:#4ade80;cursor:pointer"><input type="radio" name="correctMcq" value="'+existing+'" style="accent-color:#22c55e;width:16px;height:16px"> ✓ صحيح</label></div><div id="mcqOptEditor'+existing+'" contenteditable="true" dir="auto" class="rich-editor" style="min-height:44px" placeholder="نص الاختيار..."></div>';
  container.appendChild(div);
}

// Matching editor
function matchDotClick(col,idx){
  if(col==='A'){matchSelected={col:'A',idx:idx};document.querySelectorAll('.match-dot').forEach(function(d){d.style.background='#cbd5e1';d.style.borderColor='#94a3b8';});var dot=document.getElementById('dotA'+idx);if(dot){dot.style.background='#1e3a8a';dot.style.borderColor='#1e40af';}}
  else if(col==='B'&&matchSelected&&matchSelected.col==='A'){matchConnections=matchConnections.filter(function(c){return c.a!==matchSelected.idx&&c.b!==idx;});matchConnections.push({a:matchSelected.idx,b:idx});matchSelected=null;document.querySelectorAll('.match-dot').forEach(function(d){d.style.background='#cbd5e1';d.style.borderColor='#94a3b8';});renderMatchConnections();}
}
function renderMatchConnections(){
  var list=document.getElementById('matchConnectionsList');if(!list)return;
  list.innerHTML=matchConnections.map(function(c,i){return '<div style="display:flex;align-items:center;gap:8px;background:rgba(74,222,128,.1);border:1px solid rgba(74,222,128,.2);border-radius:8px;padding:6px 10px;margin-bottom:4px;font-size:12px"><span style="color:#93c5fd">A'+(c.a+1)+'</span><span style="color:rgba(255,255,255,.5)">↔</span><span style="color:#86efac">B'+(c.b+1)+'</span><span style="color:rgba(255,255,255,.3);font-size:10px;font-family:Montserrat,sans-serif">= Match</span><button onclick="matchConnections.splice('+i+',1);renderMatchConnections()" style="margin-right:auto;background:none;border:none;color:#f87171;cursor:pointer;font-size:12px">✕</button></div>';}).join('');
  var count=document.querySelectorAll('[id^="dotA"]').length;
  for(var i=0;i<count;i++){var dA=document.getElementById('dotA'+i),dB=document.getElementById('dotB'+i);var conn=matchConnections.find(function(c){return c.a===i;});if(dA){dA.style.background=conn?'#22c55e':'#cbd5e1';dA.style.borderColor=conn?'#16a34a':'#94a3b8';}if(conn&&dB){var bEl=document.getElementById('dotB'+conn.b);if(bEl){bEl.style.background='#22c55e';bEl.style.borderColor='#16a34a';}}}
}
function clearMatchConnections(){matchConnections=[];renderMatchConnections();}
function addMatchPair(){
  var newIdx=document.querySelectorAll('[id^="dotA"]').length;
  var colA=document.getElementById('matchColA'),colB=document.getElementById('matchColB');
  if(!colA||!colB)return;
  var divA=document.createElement('div');divA.style.marginBottom='8px';
  divA.innerHTML='<div style="display:flex;align-items:center;gap:6px"><div style="flex:1">'+makeToolbar('mAEdit'+newIdx,'auto')+'</div><div id="dotA'+newIdx+'" class="match-dot" onclick="matchDotClick(\'A\','+newIdx+')" style="width:16px;height:16px;border-radius:50%;background:#cbd5e1;border:2px solid #94a3b8;cursor:pointer;flex-shrink:0"></div></div>';
  colA.appendChild(divA);
  var divB=document.createElement('div');divB.style.marginBottom='8px';
  divB.innerHTML='<div style="display:flex;align-items:center;gap:6px"><div id="dotB'+newIdx+'" class="match-dot" onclick="matchDotClick(\'B\','+newIdx+')" style="width:16px;height:16px;border-radius:50%;background:#cbd5e1;border:2px solid #94a3b8;cursor:pointer;flex-shrink:0"></div><div style="flex:1">'+makeToolbar('mBEdit'+newIdx,'auto')+'</div></div>';
  colB.appendChild(divB);
}

// ============================================================
// SAVE QUESTION
// ============================================================
function saveQuestion(){
  var type=document.getElementById('qType').value;if(!type){alert('اختر نمط السؤال أولاً');return;}
  // ── رأس السؤال الموحد ──
  var hBox=document.getElementById('qhBox');
  var hAr=document.getElementById('qhArLine'),hEn=document.getElementById('qhEnLine');
  var headerCustom=hBox?hBox.dataset.custom==='1':!!(currentEditingQ&&currentEditingQ.headerCustom);
  var headerAr='',headerEn='';
  if(headerCustom){
    headerAr=hAr?(hAr.textContent||'').trim().replace(/^-\s*/,''):(currentEditingQ&&currentEditingQ.headerAr)||'';
    headerEn=hEn?(hEn.textContent||'').trim().replace(/^-\s*/,''):(currentEditingQ&&currentEditingQ.headerEn)||'';
  }
  var stemText=headerCustom?(headerAr+' / '+headerEn):'';
  var q={type:type,headerCustom:headerCustom,headerAr:headerAr,headerEn:headerEn,stemText:stemText,bloom:document.getElementById('qBloom').value,marking:document.getElementById('qMarking').value,lo:document.getElementById('qLO').value,standard:document.getElementById('qStandard').value,mediaHtml:extractCleanMediaHtml(),mediaVisible:currentEditingQ&&currentEditingQ.mediaVisible?currentEditingQ.mediaVisible:{}};
  q.stemHtml=buildQHBoxHtml(q);
  if(type==='mcq'){
    var qtEl=document.getElementById('qt-text');
    if(qtEl) q.questionText=qtEl.value.trim();
    var qBodyMcq=document.getElementById('qBodyEditorMcq');
    if(qBodyMcq) q.bodyHtml=qBodyMcq.innerHTML;
    var qtFont=document.getElementById('qt-font');if(qtFont) q.qTextFont=qtFont.value;
    var qtSize=document.getElementById('qt-size');if(qtSize) q.qTextSize=parseInt(qtSize.value)||16;
    var qtColor=document.getElementById('qt-color');if(qtColor) q.qTextColor=qtColor.value;
    var qtBg=document.getElementById('qt-bg');if(qtBg) q.qTextBg=qtBg.value;
    var qtDir=document.querySelector('input[name="qtDir"]:checked');if(qtDir) q.qTextDir=qtDir.value;
    var opts=[];document.querySelectorAll('[id^="mcqOptEditor"]').forEach(function(el){opts.push(el.innerHTML||'');});var cEl=document.querySelector('input[name="correctMcq"]:checked');q.options=opts;q.correct=cEl?parseInt(cEl.value):null;q.mcqDir=document.querySelector('input[name="mcqDir"]:checked')?document.querySelector('input[name="mcqDir"]:checked').value:'vertical';}
  else if(type==='matching'){
    var pairs=[];var i=0;
    // Check new editor first
    var matchEditorCont=document.getElementById('matchPairsEditor');
    if(matchEditorCont&&matchEditorCont.children.length>0){
      for(var mi=0;mi<matchEditorCont.children.length;mi++){
        var aEl=document.getElementById('mAEdit'+mi);
        var bEl=document.getElementById('mBEdit'+mi);
        var aImgWrap=document.getElementById('mAImg'+mi);
        var bImgWrap=document.getElementById('mBImg'+mi);
        var aImg=aImgWrap&&aImgWrap.querySelector('img')?aImgWrap.querySelector('img').src:'';
        var bImg=bImgWrap&&bImgWrap.querySelector('img')?bImgWrap.querySelector('img').src:'';
        var aImgEl=aImgWrap?aImgWrap.querySelector('img'):null;
        var bImgEl=bImgWrap?bImgWrap.querySelector('img'):null;
        pairs.push({
          aHtml:aEl?aEl.innerHTML||'':'',
          bHtml:bEl?bEl.innerHTML||'':'',
          aImg:aImg||'',bImg:bImg||'',
          aImgScale:aImgEl&&aImgEl.dataset.scale?parseFloat(aImgEl.dataset.scale):1,
          bImgScale:bImgEl&&bImgEl.dataset.scale?parseFloat(bImgEl.dataset.scale):1
        });
      }
    } else {
      pairs=_collectMatchPairsFromDOM();
    }
    q.pairs=pairs;q.connections=matchConnections.slice();
    // حفظ نموذج الإجابة
    var matchKeyEl=document.getElementById('matchAnswerKey');
    if(matchKeyEl) q.matchAnswerKey=matchKeyEl.value.trim();
  }
  else if(type==='ordering'){
    var groups=[];
    _orderGroups.forEach(function(g,gi){
      var words=[];
      for(var i=0;i<20;i++){var el=document.getElementById('wordTile_'+gi+'_'+i);if(el)words.push(el.textContent||'');else break;}
      if(!words.length) words=g.words||[];
      var answerOrder=[];
      for(var ai=0;ai<words.length;ai++){
        var box=document.getElementById('orderKeyBox_'+gi+'_'+ai);
        answerOrder.push(box?parseInt(box.value):ai);
      }
      groups.push({words:words,answerOrder:answerOrder});
    });
    if(!groups.length||groups.every(function(g){return !g.words.some(function(w){return w.trim();});})){
      scWarn('أضف كلمات جملة واحدة على الأقل واضغط "تأكيد الكلمات"','Add words for at least one sentence and press "Confirm Words"');return;
    }
    q.orderGroups=groups;
    // للتوافق مع أي كود قديم يقرأ q.words مباشرة (الجملة الأولى فقط)
    q.words=groups[0].words;
    q.answerOrder=groups[0].answerOrder;
    q.correctAnswer=groups[0].answerOrder.map(function(idx){return groups[0].words[idx]||'';}).join(' ');
    q.orTileBg=document.getElementById('or-tile-bg')?document.getElementById('or-tile-bg').value:'#1e3a8a';
    q.orTileColor=document.getElementById('or-tile-color')?document.getElementById('or-tile-color').value:'#ffffff';
    q.orFontSize=document.getElementById('or-font-size')?parseInt(document.getElementById('or-font-size').value)||15:15;
    q.orAnsBg=document.getElementById('or-ans-bg')?document.getElementById('or-ans-bg').value:'#f0f7ff';
    q.orAnsColor=document.getElementById('or-ans-color')?document.getElementById('or-ans-color').value:'#1e3a8a';
    q.orFont=document.getElementById('or-font')?document.getElementById('or-font').value:'Tajawal';
    q.orDir=document.getElementById('or-dir')?document.getElementById('or-dir').value:'rtl';
  }
  else if(type==='truefalse'){
    var stmts=[];
    var tfCont=document.getElementById('tf-statements');
    if(tfCont){
      Array.prototype.forEach.call(tfCont.children,function(rowDiv){
        var tfEl=rowDiv.querySelector('input[type="text"]');
        var tfRa=rowDiv.querySelector('input[type="radio"]:checked');
        if(tfEl) stmts.push({text:tfEl.value||'',answer:tfRa?tfRa.value:'true'});
      });
    }
    q.statements=stmts.length?stmts:[{text:'',answer:'true'}];
    q.tfFont=document.getElementById('tf-font')?document.getElementById('tf-font').value:'Tajawal';
    q.tfSize=document.getElementById('tf-size')?parseInt(document.getElementById('tf-size').value)||15:15;
    q.tfColor=document.getElementById('tf-color')?document.getElementById('tf-color').value:'#1a1a2e';
    q.tfBg=document.getElementById('tf-bg')?document.getElementById('tf-bg').value:'#f8fafc';
    q.tfTrueColor=document.getElementById('tf-true-color')?document.getElementById('tf-true-color').value:'#15803d';
    q.tfFalseColor=document.getElementById('tf-false-color')?document.getElementById('tf-false-color').value:'#b91c1c';
  }
  else if(type==='classify'){
    var colsEl=document.getElementById('classify-cols');
    var nC=parseInt(colsEl?colsEl.value:2)||2;
    var cols=[];
    for(var ci=0;ci<nC;ci++){
      var cEl=document.getElementById('classify-col-'+ci);
      cols.push(cEl?cEl.value||'عمود '+(ci+1):'عمود '+(ci+1));
    }
    var items=[];
    var ic=document.getElementById('classify-items-container');
    if(ic){
      Array.prototype.forEach.call(ic.children,function(rowDiv){
        var itxt=rowDiv.querySelector('input[type="text"]');
        var icol2=rowDiv.querySelector('select');
        if(itxt) items.push({text:itxt.value||'',correctCol:icol2?parseInt(icol2.value)||0:0});
      });
    }
    q.columns=cols;
    q.items=items;
    q.orTileBg=document.getElementById('cl-tile-bg')?document.getElementById('cl-tile-bg').value:'#6366f1';
    q.clTableBg=document.getElementById('cl-table-bg')?document.getElementById('cl-table-bg').value:'#f0f7ff';
    q.clFontSize=document.getElementById('cl-font-size')?parseInt(document.getElementById('cl-font-size').value)||14:14;
    q.clTextColor=document.getElementById('cl-text-color')?document.getElementById('cl-text-color').value:'#ffffff';
  }
  else if(type==='speaking'){
    q.speakingInst=document.getElementById('speakingInst')?document.getElementById('speakingInst').value:'';
    q.speakingDir=document.getElementById('sp-dir')?document.getElementById('sp-dir').value:'rtl';
    q.speakingColor=document.getElementById('sp-color')?document.getElementById('sp-color').value:'#1a1a2e';
    q.speakingBg=document.getElementById('sp-bg')?document.getElementById('sp-bg').value:'#f8fafc';
    q.speakingFont=document.getElementById('sp-font')?document.getElementById('sp-font').value:'Tajawal';
    q.speakingSize=document.getElementById('sp-size')?parseInt(document.getElementById('sp-size').value)||15:15;
  }
  else if(type==='oral'){
    var oralEl=document.getElementById('oralTextEditor');
    q.oralText=oralEl?oralEl.innerHTML:'';
    q.speakingInst=document.getElementById('speakingInst')?document.getElementById('speakingInst').value:'';
    q.speakingDir=document.getElementById('sp-dir')?document.getElementById('sp-dir').value:'rtl';
    q.speakingColor=document.getElementById('sp-color')?document.getElementById('sp-color').value:'#1a1a2e';
    q.speakingBg=document.getElementById('sp-bg')?document.getElementById('sp-bg').value:'#fffbeb';
    q.speakingFont=document.getElementById('sp-font')?document.getElementById('sp-font').value:'Tajawal';
    q.speakingSize=document.getElementById('sp-size')?parseInt(document.getElementById('sp-size').value)||20:20;
  }
  else if(type==='listening'){
    q.ansType=document.getElementById('listeningAnswerType')?document.getElementById('listeningAnswerType').value:'';
    var lisAud=document.getElementById('listen-audio-preview');
    if(lisAud&&lisAud.src) q.audioSrc=lisAud.src;if(q.ansType==='mcq'){var opts2=[];document.querySelectorAll('[id^="mcqOptEditor"]').forEach(function(el){opts2.push(el.innerHTML||'');});var c2=document.querySelector('input[name="correctMcq"]:checked');q.options=opts2;q.correct=c2?parseInt(c2.value):null;}else if(q.ansType==='matching'){q.pairs=_collectMatchPairsFromDOM();q.connections=matchConnections.slice();}}
  else if(type==='reading'||type==='writingskill'){q.passageHtml=document.getElementById('passageEditor')?document.getElementById('passageEditor').innerHTML:'';q.answerStemHtml=document.getElementById('answerStemEditor')?document.getElementById('answerStemEditor').innerHTML:'';q.ansType=document.getElementById('readingAnswerType')?document.getElementById('readingAnswerType').value:'';if(q.ansType==='mcq'){var opts3=[];document.querySelectorAll('[id^="mcqOptEditor"]').forEach(function(el){opts3.push(el.innerHTML||'');});var c3=document.querySelector('input[name="correctMcq"]:checked');q.options=opts3;q.correct=c3?parseInt(c3.value):null;
    // خصائص خط نص السؤال
    var rdmqText=document.getElementById('rdmcq-text');if(rdmqText)q.questionText=rdmqText.value.trim();
    var rdmqFont=document.getElementById('rdmcq-font');if(rdmqFont)q.qTextFont=rdmqFont.value;
    var rdmqSize=document.getElementById('rdmcq-size');if(rdmqSize)q.qTextSize=parseInt(rdmqSize.value)||15;
    var rdmqColor=document.getElementById('rdmcq-color');if(rdmqColor)q.qTextColor=rdmqColor.value;
    var rdmqBg=document.getElementById('rdmcq-bg');if(rdmqBg)q.qTextBg=rdmqBg.value;
    var rdmqDir=document.querySelector('input[name="rdmcqDir"]:checked');if(rdmqDir)q.qTextDir=rdmqDir.value;
  }else if(q.ansType==='matching'){q.pairs=_collectMatchPairsFromDOM();q.connections=matchConnections.slice();}}
  var scoreVal=Number(document.getElementById('qScore')?document.getElementById('qScore').value||0:0);
  if(!scoreVal||scoreVal<=0){alert('⚠️ يجب تحديد نسبة السؤال\nPlease enter question weight %');return;}
  // Validate total weight
  var d2=testData.domains[currentDomainIndex];
  var parentWeight2=currentBranchIndex>=0&&d2.branches&&d2.branches[currentBranchIndex]?d2.branches[currentBranchIndex].weight:d2.weight;
  var qs2=_getCurrentQuestions();
  var usedSum2=qs2.reduce(function(s,q,i){return i===currentQuestionIndex?s:s+(Number(q.score)||0);},0);
  var newTotal=usedSum2+scoreVal;
  if(Math.round(newTotal*100)>Math.round(parentWeight2*100)){
    alert('⚠️ نسبة السؤال ('+scoreVal+'%) ستتجاوز المجموع المسموح به!\nالمتبقي: '+(parentWeight2-usedSum2).toFixed(2)+'%\n\nQuestion weight exceeds available: '+(parentWeight2-usedSum2).toFixed(2)+'% remaining');
    return;
  }
  q.score=scoreVal;
  var qs=_getCurrentQuestions();
  if(currentQuestionIndex>=0) qs[currentQuestionIndex]=q; else qs.push(q);
  _saveDraft();closeQuestionModal();renderQuestionsList();
}

// ============================================================
// STUDENT WINDOW
// ============================================================
// ══ Inject Logo SVG ══
function injectLogoSVG(){
  var wrap=document.getElementById('sw-logo-svg');
  if(wrap) wrap.innerHTML=SCHOLASTIC_LOGO_SVG;
  // If school logo exists, show it instead
  var logoImg=document.getElementById('sw-logo-img');
  if(logoImg&&testData.logoSrc){
    logoImg.src=testData.logoSrc;logoImg.style.display='block';
    if(wrap) wrap.style.display='none';
  } else {
    if(logoImg) logoImg.style.display='none';
    if(wrap) wrap.style.display='flex';
  }
}

function openStudentWindow(domainIdx,isRealTest){
  sw_domainIdx=domainIdx||0;sw_branchIdx=-1;sw_qIdx=0;sw_answers={};
  applyDisplayMode();
  document.getElementById('studentWindow').style.display='flex';
  var d=testData.domains[sw_domainIdx];
  if(!d) return;
  // اسم الاختبار والمعلومات
  var testNameEl=document.getElementById('sw-test-name-header');
  if(testNameEl) testNameEl.textContent=testData.testName||'اختبار';
  var metaEl=document.getElementById('sw-test-meta-header');
  if(metaEl){
    var meta=[];
    if(testData.subject) meta.push(testData.subject);
    if(testData.grade) meta.push(testData.grade);
    if(testData.term) meta.push('Term '+testData.term);
    metaEl.textContent=meta.join(' — ');
  }
  // اسم المجال
  var domEl=document.getElementById('sw-domain-label-header');
  if(domEl) domEl.textContent='Domain: '+(d.nameEn||d.nameAr||'Domain')+(d.nameAr&&d.nameEn?' | المجال: '+d.nameAr:'');
  var brEl=document.getElementById('sw-branch-label-header');
  if(brEl) brEl.textContent=d?(d.nameEn||d.nameAr||''):'';
  // في المعاينة: تخطى التعليمات مباشرة للأسئلة
  if(_tryModeActive||!isRealTest){
    sw_instructionsConfirmed=true;
    startStudentTimer(parseInt(d.time||30)*60);
    renderStudentQuestion();
  } else {
    sw_instructionsConfirmed=false;
    showStudentInstructions();
  }
}
function previewCurrentInstructions(){
  testData.instructionsAr=document.getElementById('insAr')?document.getElementById('insAr').innerHTML.trim():'',
  testData.instructionsEn=document.getElementById('insEn')?document.getElementById('insEn').innerHTML.trim():'';
  // معاينة صفحة التعليمات فقط — بنفس الشكل في كل الأنماط (كلاسيكي/ورقة بيضاء)
  sw_domainIdx=0;sw_branchIdx=-1;sw_qIdx=0;sw_answers={};
  _tryModeActive=true;_previewReturnStep=2;
  applyDisplayMode();
  document.getElementById('studentWindow').style.display='flex';
  var testNameEl=document.getElementById('sw-test-name-header');if(testNameEl)testNameEl.textContent=testData.testName||'اختبار';
  var metaEl=document.getElementById('sw-test-meta-header');
  if(metaEl){var meta=[];if(testData.subject)meta.push(testData.subject);if(testData.grade)meta.push(testData.grade);if(testData.term)meta.push('Term '+testData.term);metaEl.textContent=meta.join(' — ');}
  var cb=document.getElementById('sw-close-btn');if(cb)cb.textContent='✕ إغلاق المعاينة';
  sw_instructionsConfirmed=false;
  showStudentInstructions();
}
var _savedTestDataForPreview=null;
var _previewReturnStep=4; // الخطوة التي يرجع إليها بعد المعاينة
function closeStudentWindow(){
  clearInterval(sw_timerInterval);disableAntiCheat();swDisableFullLock();
  document.getElementById('studentWindow').style.display='none';
  if(_savedTestDataForPreview){
    testData=_savedTestDataForPreview;
    _savedTestDataForPreview=null;
  }
  if(_tryModeActive){
    _tryModeActive=false;
    if(_previewReturnStep>=0) goToStep(_previewReturnStep);
    _previewReturnStep=4;
  }
}
function updateTimerDisplay(){
  var m=Math.floor(sw_timeLeft/60),s=sw_timeLeft%60;
  var el=document.getElementById('sw-timer-display');
  el.textContent=String(m).padStart(2,'0')+':'+String(s).padStart(2,'0');
  el.className='sw-timer'+(sw_timeLeft<=60?' warn':'');
}
function startStudentTimer(seconds){
  stopStudentTimer();
  sw_timeLeft=seconds||30*60;
  updateTimerDisplay();
  sw_timerInterval=setInterval(function(){sw_timeLeft--;updateTimerDisplay();if(sw_timeLeft<=0){clearInterval(sw_timerInterval);swAutoAdvance();}},1000);
}
function stopStudentTimer(){
  if(sw_timerInterval){clearInterval(sw_timerInterval);sw_timerInterval=null;}
}
function selectStudentDomain(idx){
  if(idx<0||idx>=testData.domains.length) return;
  var d=testData.domains[idx];
  sw_domainIdx=idx;
  sw_branchIdx=-1;
  if(d.hasBranches&&d.branches&&d.branches.length){
    sw_currentPage='branches';
    sw_qIdx=0;
    stopStudentTimer();
  } else {
    sw_qIdx=0;sw_currentPage='questions';
    sw_answers=sw_branchAnswers[idx]&&sw_branchAnswers[idx][-1]?sw_branchAnswers[idx][-1]:{};
    if(d){startStudentTimer((d.time||30)*60);}
  }
  renderStudentWindowPage();
}
function selectStudentBranch(branchIdx){
  var d=testData.domains[sw_domainIdx];
  if(!d||!d.branches||!d.branches[branchIdx]) return;
  sw_branchIdx=branchIdx;
  sw_qIdx=0;
  sw_currentPage='questions';
  if(!sw_branchAnswers[sw_domainIdx]) sw_branchAnswers[sw_domainIdx]={};
  sw_answers=sw_branchAnswers[sw_domainIdx][branchIdx]||{};
  var br=d.branches[branchIdx];
  startStudentTimer((br.time||20)*60);
  renderStudentWindowPage();
}
function renderStudentWindowPage(){
  var footer=document.getElementById('sw-footer');
  if(footer) footer.style.display=(testData.displayMode===3)?'none':'';
  var content=document.getElementById('sw-questions-content');
  var header=document.getElementById('sw-domain-label-header');
  var barName=document.getElementById('sw-domain-name-bar');
  var counterTop=document.getElementById('sw-q-counter-top');
  var qCounter=document.getElementById('sw-q-counter');
  var prevBtn=document.getElementById('sw-prev-btn');
  var nextBtn=document.getElementById('sw-next-btn');
  if(sw_currentPage==='instructions'){
    header.textContent='تعليمات الاختبار';
    barName.textContent='Instructions';
    counterTop.textContent='هذه هي نفس تعليمات المشرف. اضغط أبدأ بعد القراءة / This is the same supervisor instructions. Press Start after reading';
    qCounter.textContent='';
    prevBtn.disabled=true;
    nextBtn.disabled=true;
    nextBtn.innerHTML='<span>التالي ←</span><span class="btn-sub">Domains</span>';
    nextBtn.className='sw-nav-btn submit';
    var ar=testData.instructionsAr||'';
    var en=testData.instructionsEn||'';
    // هل يوجد سؤال تحدث في الاختبار؟
    var hasSpeaking=testData.domains.some(function(d){
      var qs=d.questions||[];
      var brQs=d.branches?d.branches.reduce(function(a,b){return a.concat(b.questions||[]);},[]):[];
      return qs.concat(brQs).some(function(q){return q.type==='speaking'||q.type==='oral';});
    });
    var micBtn=hasSpeaking
      ?'<div id="sw-mic-check" style="margin-bottom:14px;padding:14px 18px;border:2px solid rgba(250,204,21,.4);border-radius:14px;background:rgba(250,204,21,.07);display:flex;align-items:center;gap:12px;flex-wrap:wrap">'
        +'<span style="font-size:20px">🎙</span>'
        +'<div style="flex:1"><div style="font-weight:800;color:#92400e;font-size:14px">هذا الاختبار يحتوي على أسئلة تحدث — يرجى تفعيل الميكروفون أولاً</div>'
        +'<div style="font-size:12px;color:#78350f;font-family:Montserrat,sans-serif;direction:ltr;text-align:left;margin-top:2px">This test has speaking questions — please enable your microphone first</div></div>'
        +'<button onclick="swCheckMic()" id="sw-mic-btn" style="background:linear-gradient(135deg,#f59e0b,#d97706);color:white;border:none;border-radius:10px;padding:9px 18px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;white-space:nowrap">🎙 تحقق / Check</button>'
        +'<span id="sw-mic-status" style="font-size:12px;font-weight:700"></span>'
        +'</div>'
      :'';
    var confirmText =
      '<div style="margin-bottom:10px;padding:14px 18px;border:2px solid #3b82f6;border-radius:14px;background:#eff6ff">'
      +'<div style="font-weight:800;color:#1e40af;font-size:15px;line-height:1.6;direction:rtl;text-align:right;margin-bottom:10px">أؤكد أنني قرأت معلومات الاختبار بالكامل.</div>'
      +'<label style="display:flex;align-items:center;gap:12px;cursor:pointer" dir="ltr">'
      +'<input type="checkbox" id="sw-confirm-check" onchange="swCheckboxChanged()" style="width:22px;height:22px;accent-color:#1e3a8a;flex-shrink:0;cursor:pointer">'
      +'<span style="font-weight:700;color:#1e40af;font-size:14px;line-height:1.6;font-family:Montserrat,sans-serif">I confirm I have read the test information completely.</span>'
      +'</label>'
      +'</div>';
    var startButton = micBtn+'<div style="text-align:center;margin-top:16px">'
      +'<button id="sw-start-btn" onclick="swConfirmInstructions()" disabled style="background:linear-gradient(135deg,#94a3b8,#64748b);color:white;border:3px solid #cbd5e1;border-radius:18px;padding:14px 32px;font-size:16px;font-weight:800;cursor:not-allowed;opacity:.6;font-family:Tajawal,sans-serif;transition:all .3s">✅ أبدأ الاختبار / Start Test</button>'
      +'<div id="sw-start-hint" style="font-size:12px;color:#94a3b8;margin-top:8px;font-family:Montserrat,sans-serif">'
      +'✅ أكد القراءة باللغتين'
      +(hasSpeaking?' + 🎙 تأكد من الميكروفون':'')
      +' / Confirm both checkboxes'+(hasSpeaking?' + mic':'')+'</div>'
      +'</div>';
    if(!ar && !en){
      content.innerHTML='<div class="sw-frame" style="padding:32px;text-align:center;color:#64748b"><div style="font-size:20px;font-weight:800;margin-bottom:12px">لا توجد تعليمات محددة</div><p>تابع إلى صفحة المجالات للبدء.</p>'+confirmText+startButton+'</div>';
    } else {
      var instructionsHtml = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px">'
        +'<div style="padding:20px;border:2px solid #3b82f6;border-radius:16px;background:#eff6ff;text-align:right;line-height:1.8;color:#1a1a2e" dir="rtl"><div style="font-size:18px;font-weight:800;color:#1e40af;margin-bottom:12px">📘 التعليمات بالعربية</div>' + (ar||'<span style="color:#94a3b8;font-size:13px">لا توجد تعليمات بالعربية</span>') + '</div>'
        +'<div style="padding:20px;border:2px solid #3b82f6;border-radius:16px;background:#eff6ff;text-align:left;line-height:1.8;color:#1a1a2e;font-family:Montserrat,sans-serif" dir="ltr"><div style="font-size:18px;font-weight:800;color:#1e40af;margin-bottom:12px">📘 Instructions in English</div>' + (en||'<span style="color:#94a3b8;font-size:13px">No English instructions</span>') + '</div>'
        +'</div>';
      content.innerHTML='<div class="sw-frame" style="padding:32px"><div style="font-size:22px;font-weight:900;color:#1e3a8a;margin-bottom:18px">📘 تعليمات الاختبار</div>' + instructionsHtml + confirmText + startButton + '</div>';
    }
    stopStudentTimer();
    document.getElementById('sw-progress-dots').innerHTML='';
    document.getElementById('sw-q-counter-top').textContent='';
    return;
  }
  if(sw_currentPage==='domains'){
    header.textContent='اختر المجال';
    barName.textContent='Choose Domain';
    counterTop.textContent='اضغط أيقونة المجال للبدء';
    qCounter.textContent='';
    prevBtn.disabled=false;
    prevBtn.innerHTML='<span>← تعليمات</span><span class="btn-sub">Instructions</span>';
    prevBtn.className='sw-nav-btn prev';
    nextBtn.disabled=true;
    nextBtn.innerHTML='<span>اختر مجالاً</span><span class="btn-sub">Select Domain</span>';
    nextBtn.className='sw-nav-btn submit';
    var domColors=['linear-gradient(135deg,#f97316,#dc2626)','linear-gradient(135deg,#7e22ce,#1e3a8a)','linear-gradient(135deg,#0891b2,#0d9488)','linear-gradient(135deg,#be185d,#7e22ce)','linear-gradient(135deg,#ca8a04,#b45309)','linear-gradient(135deg,#15803d,#0891b2)'];
    var html=testData.domains.map(function(d,i){
      var isDone=sw_completedDomains.indexOf(i)>=0;
      var iconHtml='<div style="width:72px;height:72px;border-radius:50%;background:'+(isDone?'linear-gradient(135deg,#22c55e,#15803d)':domColors[i%domColors.length])+';display:flex;align-items:center;justify-content:center;margin:0 auto 14px;box-shadow:0 6px 18px rgba(0,0,0,.18);border:3px solid rgba(255,255,255,.3);flex-shrink:0">'+(d.iconSrc?'<img src="'+d.iconSrc+'" style="width:46px;height:46px;object-fit:contain;border-radius:50%">':'<span style="font-size:30px">📊</span>')+'</div>';
      var doneTag=isDone?'<div style="display:inline-block;background:#22c55e;color:white;font-size:11px;font-weight:800;padding:4px 14px;border-radius:20px;margin-top:8px;font-family:Tajawal,sans-serif">✅ مكتمل / Done</div>':'<div style="display:inline-block;background:#f59e0b;color:white;font-size:11px;font-weight:800;padding:4px 14px;border-radius:20px;margin-top:8px;font-family:Tajawal,sans-serif">⏳ لم يكتمل</div>';
      return '<button onclick="selectStudentDomain('+i+')" style="width:100%;border:3px solid '+(isDone?'#22c55e':'rgba(30,58,138,.15)')+';background:'+(isDone?'linear-gradient(135deg,#f0fdf4,#dcfce7)':'white')+';color:#1a1a2e;padding:24px 18px;border-radius:24px;text-align:center;cursor:pointer;transition:.25s;min-height:210px;box-shadow:0 8px 28px rgba(15,23,42,.08)">'+iconHtml+'<div style="font-size:17px;font-weight:800;margin-bottom:6px">'+(d.nameAr||('مجال '+(i+1)))+'</div>'+(d.nameEn?'<div style="font-size:13px;color:#64748b;font-family:Montserrat,sans-serif;margin-bottom:10px">'+d.nameEn+'</div>':'')+'<div style="font-size:13px;font-weight:700;color:#1e3a8a">'+(d.weight||0)+'%</div>'+doneTag+'</button>';
    }).join('');
    content.innerHTML='<div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">'+html+'</div>';
    stopStudentTimer();
    document.getElementById('sw-progress-dots').innerHTML='';
    document.getElementById('sw-q-counter-top').textContent='';
    return;
  }
  // BRANCHES PAGE
  if(sw_currentPage==='branches'){
    var d=testData.domains[sw_domainIdx];
    header.textContent=d.nameAr||'المجال';
    barName.textContent=d.nameEn||'Domain';
    counterTop.textContent='اختر فرعاً للبدء / Choose a Branch';
    qCounter.textContent='';
    prevBtn.disabled=false;
    prevBtn.innerHTML='<span>← المجالات</span><span class="btn-sub">Domains</span>';
    prevBtn.className='sw-nav-btn prev';
    nextBtn.disabled=true;
    nextBtn.innerHTML='<span>اختر فرعاً</span><span class="btn-sub">Select Branch</span>';
    nextBtn.className='sw-nav-btn submit';
    var completedBr=sw_completedBranches[sw_domainIdx]||[];
    var allDone=d.branches&&d.branches.length>0&&completedBr.length>=d.branches.length;
    var brColors=['linear-gradient(135deg,#0ea5e9,#6366f1)','linear-gradient(135deg,#6366f1,#7e22ce)','linear-gradient(135deg,#f59e0b,#ef4444)','linear-gradient(135deg,#22c55e,#0891b2)','linear-gradient(135deg,#ec4899,#f97316)','linear-gradient(135deg,#0d9488,#3b82f6)'];
    var brHtml=(d.branches||[]).map(function(br,bi){
      var isDone=completedBr.indexOf(bi)>=0;
      var iconHtml='<div style="width:72px;height:72px;border-radius:50%;background:'+(isDone?'linear-gradient(135deg,#22c55e,#15803d)':brColors[bi%brColors.length])+';display:flex;align-items:center;justify-content:center;margin:0 auto 12px;box-shadow:0 6px 18px rgba(0,0,0,.16);border:3px solid rgba(255,255,255,.3)">'+(br.iconSrc?'<img src="'+br.iconSrc+'" style="width:46px;height:46px;object-fit:contain;border-radius:50%">':'<span style="font-size:28px">📁</span>')+'</div>';
      return '<div class="sw-branch-card'+(isDone?' done':'')+'" onclick="selectStudentBranch('+bi+')">'+iconHtml+'<div style="font-size:16px;font-weight:800;color:#1a1a2e;margin-bottom:4px">'+(br.nameAr||('فرع '+(bi+1)))+'</div>'+(br.nameEn?'<div style="font-size:12px;color:#64748b;font-family:Montserrat,sans-serif;margin-bottom:8px">'+br.nameEn+'</div>':'')+'<div style="font-size:12px;font-weight:700;color:#1e3a8a">'+(br.qCount||0)+' سؤال | '+(br.time||0)+' دقيقة</div></div>';
    }).join('');
    var doneBar=allDone?'<div style="text-align:center;margin-top:20px;padding:16px;background:linear-gradient(135deg,#f0fdf4,#dcfce7);border:2px solid #22c55e;border-radius:18px"><div style="font-size:32px;margin-bottom:6px">🎉</div><div style="font-size:16px;font-weight:800;color:#16a34a">أكملت جميع الفروع! / All Branches Done!</div><div style="margin-top:12px"><button onclick="swSubmitDomain(0)" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:3px solid #86efac;border-radius:14px;padding:12px 28px;font-size:15px;font-weight:800;cursor:pointer;box-shadow:0 4px 14px rgba(34,197,94,.4);font-family:Tajawal,sans-serif">✅ تسليم المجال / Submit Domain</button></div></div>':'';
    content.innerHTML='<div style="padding:8px"><div style="font-size:13px;color:#64748b;margin-bottom:16px;text-align:center">المجال يحتوي على فروع — أكمل كل الفروع لإتمام المجال<br><span style="font-family:Montserrat,sans-serif;font-size:11px">Domain has branches — complete all to finish the domain</span></div><div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">'+brHtml+'</div>'+doneBar+'</div>';
    document.getElementById('sw-progress-dots').innerHTML='';
    document.getElementById('sw-q-counter-top').textContent='';
    return;
  }
  renderStudentQuestion();
}
function swJumpTo(qi){
  var mode=testData.displayMode||1;
  if(mode===4){
    var elWP=document.getElementById('wp-q-'+qi);
    if(elWP) elWP.scrollIntoView({behavior:'smooth',block:'center'});
    return;
  }
  if(mode===3){
    var elStream=document.getElementById('stream-q-'+qi);
    if(elStream) elStream.scrollIntoView({behavior:'smooth',block:'center'});
    return;
  }
  sw_qIdx=qi;
  renderStudentQuestion();
  updateProgressDots();
}
function renderProgressDots(total){
  var html='';
  for(var i=0;i<total;i++){
    var cls='sw-dot';
    if(sw_answers[i]!==undefined) cls+=' answered';
    else if(i===sw_qIdx) cls+=' current';
    html+='<div class="'+cls+'" id="sw-dot-'+i+'" onclick="swJumpTo('+i+')" title="Q.'+(i+1)+'">'+(i+1)+'</div>';
  }
  document.getElementById('sw-progress-dots').innerHTML=html;
  var d=testData&&testData.domains&&testData.domains[sw_domainIdx];
  var domLabel=d?(d.nameEn||d.nameAr||'Domain'):'Domain';
  var brLabel='';
  if(sw_branchIdx>=0&&d&&d.branches&&d.branches[sw_branchIdx]){
    var br=d.branches[sw_branchIdx];
    brLabel='Branch: '+(br.nameEn||br.nameAr||('Branch '+(sw_branchIdx+1)));
  }
  var ctr=document.getElementById('sw-q-counter-top');
  if(ctr) ctr.textContent='Domain '+(sw_domainIdx+1)+' Q'+(sw_qIdx+1)+'/'+total;
  var domEl=document.getElementById('sw-domain-label-header');
  if(domEl){
    var termS=testData.term?'Term '+testData.term:'';
    var subS=testData.subject||testData.testName||'';
    domEl.textContent=(termS&&subS)?termS+' — '+subS:(termS||subS||domLabel);
  }
  var brEl=document.getElementById('sw-branch-label-header');
  if(brEl) brEl.textContent=brLabel;
}
function swQuestionIsComplete(q,qi){
  if(!q) return true;
  if(q.type==='mcq') return sw_answers[qi]!==undefined;
  if(q.type==='truefalse'){
    var stmts=q.statements||[];
    if(!stmts.length) return true;
    var ans=sw_answers[qi];
    if(!ans) return false;
    return stmts.every(function(_,si){return ans[si]!==undefined;});
  }
  if(q.type==='classify'){
    var items=q.items||[];
    if(!items.length) return true;
    var ansC=sw_answers[qi];
    if(!ansC||!ansC.placed) return false;
    var placedCount=0;
    Object.keys(ansC.placed).forEach(function(k){placedCount+=ansC.placed[k].length;});
    return placedCount>=items.length;
  }
  if(q.type==='ordering'){
    var groups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words}]:[]);
    if(!groups.length) return true;
    var ansO=sw_answers[qi];
    if(!ansO||!ansO.groups) return false;
    return groups.every(function(g,gi){
      var gg=ansO.groups[gi];
      if(!gg||!gg.placed) return false;
      return gg.placed.every(function(p){return p!==null&&p!==undefined;});
    });
  }
  return true; // باقي الأنواع (كتابة/تحدث/توصيل/قراءة..) لا تفرض اكتمالاً إلزامياً
}
function swConfirmAndNext(){
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[]):(d.questions||[]);
  var q=questions[sw_qIdx];
  if(!swQuestionIsComplete(q,sw_qIdx)){
    scWarn('لم تكتمل إجابتك على هذا السؤال بعد، يرجى إكمالها قبل المتابعة','You have not completed this question yet, please finish before continuing');
    return;
  }
  if(sw_answers[sw_qIdx]===undefined) sw_answers[sw_qIdx]=sw_answers[sw_qIdx]; // يبقى كما هو إن وُجد
  updateProgressDots();
  if(sw_qIdx<questions.length-1){
    sw_navDir=1;sw_qIdx++;
    renderStudentQuestion();updateProgressDots();
  } else {
    swNavigate(1); // آخر سؤال — يُسلّم المجال/الفرع كما في المنطق الحالي
  }
}
function updateProgressDots(){
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[]):(d.questions||[]);
  questions.forEach(function(_,i){
    var dot=document.getElementById('sw-dot-'+i);if(!dot)return;
    var cls='sw-dot';
    if(sw_answers[i]!==undefined) cls+=' answered';
    else if(i===sw_qIdx) cls+=' current';
    dot.className=cls;
  });
  var domLabel=d?(d.nameEn||d.nameAr||'Domain'):'Domain';
  var ctr=document.getElementById('sw-q-counter-top');
  if(ctr) ctr.textContent=(testData.displayMode===4)?'':('Domain '+(sw_domainIdx+1)+' Q'+(sw_qIdx+1)+'/'+questions.length);
  var domEl=document.getElementById('sw-domain-label-header');
  if(domEl){
    var termS2=testData.term?'Term '+testData.term:'';
    var subS2=testData.subject||testData.testName||'';
    domEl.textContent=(termS2&&subS2)?termS2+' — '+subS2:(termS2||subS2||domLabel);
  }
  // اسم الفرع
  var brEl=document.getElementById('sw-branch-label-header');
  if(brEl){
    if(sw_branchIdx>=0&&d.branches&&d.branches[sw_branchIdx]){
      var br=d.branches[sw_branchIdx];
      brEl.textContent=(d.nameEn||d.nameAr||'')+' — '+(br.nameEn||br.nameAr||('Branch '+(sw_branchIdx+1)));
    } else {
      brEl.textContent=d?(d.nameEn||d.nameAr||''):'';
    }
  }
}
function renderStudentQuestion(){
  var mode=testData.displayMode||1;
  if(mode===4){renderWhitePaper();return;}
  if(mode===3){renderStreamMode();return;}
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[]):(d.questions||[]);
  var q=questions[sw_qIdx];if(!q)return;
  var score=((sw_branchIdx>=0?(d.branches[sw_branchIdx].weight||0):d.weight)/Math.max(questions.length,1)).toFixed(2);
  var total=questions.length;
  var labels=['A','B','C','D','E','F'];
  // بناء دوائر الإرشاد
  renderProgressDots(total);
  document.getElementById('sw-q-counter').textContent='السؤال '+(sw_qIdx+1)+' من '+total+' / Q '+(sw_qIdx+1)+' of '+total;
  document.getElementById('sw-prev-btn').disabled=sw_qIdx===0;
  document.getElementById('sw-prev-btn').innerHTML=sw_qIdx===0?'<span>← رجوع</span><span class="btn-sub">'+(sw_branchIdx>=0?'Branches':'Domains')+'</span>':'<span>← السابق</span><span class="btn-sub">Previous</span>';
  document.getElementById('sw-prev-btn').className='sw-nav-btn prev';
  document.getElementById('sw-footer').style.display='';
  var nextBtn=document.getElementById('sw-next-btn');
  nextBtn.disabled=false;
  var isLast=sw_qIdx===total-1;
  nextBtn.innerHTML=isLast?'<span>تسليم ✓</span><span class="btn-sub">'+(sw_branchIdx>=0?'Submit Branch':'Submit Domain')+'</span>':'<span>التالي ←</span><span class="btn-sub">Next</span>';
  nextBtn.className=isLast?'sw-nav-btn submit danger':'sw-nav-btn submit';
  var bodyHtml=buildQuestionBodyHtml(q,sw_qIdx,labels,score);
  var animClass=mode===2?(sw_navDir>=0?'fc-enter':'fc-enter-back'):'';
  document.getElementById('sw-questions-content').innerHTML=
    '<div class="sw-frame '+animClass+'">'
    // رأس السؤال — Deep Sky Blue
    +'<div class="sw-q-header">'
      +'<div class="sw-q-badge">Q.'+(sw_qIdx+1)+'</div>'
      +'<div class="sw-q-stem" dir="auto">'+(q.stemHtml||'Question')+'</div>'
      +'<div class="sw-q-score">'+score+'%</div>'
    +'</div>'
    // المصادر والجسم — إطار ذهبي علوي
    +'<div class="sw-q-body-wrap">'
      +(q.mediaHtml&&q.type!=='listening'&&q.type!=='reading'&&q.type!=='speaking'&&q.type!=='oral'&&(!(q.mediaVisible&&q.mediaVisible.img===false))?'<div style="margin-bottom:14px;border-radius:10px;overflow:hidden">'+q.mediaHtml+'</div>':'')
      +(q.bodyHtml?'<div style="font-size:15px;line-height:1.8;color:#1a1a2e;margin-bottom:12px" dir="auto">'+q.bodyHtml+'</div>':'')
      +(q.questionText?'<div style="font-size:16px;font-weight:700;color:#1e3a8a;margin-bottom:12px;padding:10px 14px;background:#eff6ff;border-radius:10px;border-right:4px solid #3b82f6" dir="auto">'+q.questionText+'</div>':'')
    +'</div>'
    // الإجابات — إطار أخضر غامق سفلي
    +'<div class="sw-q-ans-wrap">'
      +bodyHtml
    +'</div>'
    // زر تأكيد موحّد لكل الأسئلة
    +'<div style="padding:14px 20px;text-align:center;border-top:1px solid rgba(0,0,0,.06)">'
      +'<button onclick="swConfirmAndNext()" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:none;border-radius:16px;padding:12px 32px;font-size:14px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;box-shadow:0 4px 14px rgba(34,197,94,.4)">✅ تأكيد الإجابة والانتقال للسؤال التالي / Confirm &amp; Next</button>'
    +'</div>'
    +'</div>';
  if((q.type==='reading'||q.type==='writingskill')&&q.ansType==='text') setTimeout(initCanvas,100);
  if(q.type==='ordering') setTimeout(initStudentOrdering,100);
  if(q.type==='matching') setTimeout(function(){initStudentMatching(q);},100);
  updateProgressDots();
}

// Shared question body builder (used by both classic/focus and stream modes)
function buildQuestionBodyHtml(q,qi,labels,score){
  var bodyHtml='';
  if(q.type==='mcq'){
    var isH=q.mcqDir==='horizontal';
    bodyHtml='<div class="sw-mcq-options'+(isH?' horizontal':'')+'">'+
      (q.options||[]).map(function(o,i){return '<div class="sw-mcq-opt'+(isH?' horizontal':'')+(sw_answers[qi]===i?' selected':'')+'" onclick="swSelectMcq('+i+')"><div class="sw-opt-label">'+(labels[i]||i+1)+'</div><div style="flex:1;font-size:15px;color:#1a1a2e;line-height:1.6" dir="auto">'+(o||'—')+'</div></div>';}).join('')+'</div>';
  } else if(q.type==='matching'){
    bodyHtml=buildStudentMatching(q);
  } else if(q.type==='ordering'){
    bodyHtml=buildStudentOrdering(q);
  } else if(q.type==='speaking'||q.type==='oral'){
    var spStyle='font-size:'+(q.speakingSize||15)+'px;color:'+(q.speakingColor||'#1a1a2e')+';background:'+(q.speakingBg||'#f8fafc')+';font-family:'+(q.speakingFont||'Tajawal')+',sans-serif;direction:'+(q.speakingDir||'rtl');
    bodyHtml='';
    if(q.mediaHtml&&!(q.mediaVisible&&q.mediaVisible.img===false)){
      bodyHtml+='<div style="margin-bottom:14px;border-radius:14px;overflow:hidden">'+q.mediaHtml+'</div>';
    }
    if(q.type==='oral'&&q.oralText){
      bodyHtml+='<div style="'+spStyle+';padding:22px 26px;border-radius:18px;border:3px solid rgba(250,204,21,.5);margin-bottom:18px;line-height:2.2;box-shadow:0 4px 20px rgba(0,0,0,.08);font-size:'+(q.speakingSize||20)+'px" dir="auto">'+q.oralText+'</div>';
    }
    bodyHtml+='<div style="color:#64748b;font-size:14px;margin-bottom:14px;line-height:1.6;padding:10px 14px;background:#f8fafc;border-radius:10px;border-right:3px solid #e2e8f0" dir="auto">'+(q.speakingInst||(q.type==='oral'?'📖 اقرأ النص أعلاه بصوت واضح ثم سجّل قراءتك':'🎙 سجّل إجابتك الشفهية'))+'</div>'
    +'<div class="sw-audio-player" style="background:linear-gradient(135deg,#f0f9ff,#e0f2fe);border:2px solid #bae6fd;border-radius:18px;padding:20px">'
    +'<div style="display:flex;justify-content:center;gap:10px;flex-wrap:wrap;margin-bottom:14px">'
    +'<button id="sw-record-btn" onclick="swToggleRecord()" style="display:flex;align-items:center;gap:8px;background:linear-gradient(135deg,#ef4444,#dc2626);color:white;border:none;border-radius:14px;padding:12px 24px;font-size:14px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;box-shadow:0 4px 12px rgba(239,68,68,.4)">🎙 ابدأ التسجيل / Record</button>'
    +'</div>'
    +'<div id="sw-record-status" style="font-size:12px;color:#64748b;text-align:center;margin-bottom:10px;font-family:Montserrat,sans-serif;min-height:18px"></div>'
    +'<audio id="sw-playback-audio" controls style="display:none;width:100%;border-radius:12px;direction:ltr;margin-bottom:12px"></audio>'
    +'<div id="sw-record-controls" style="display:none;justify-content:center;gap:10px;flex-wrap:wrap">'
    +'<button onclick="swListenPlayback()" style="display:flex;align-items:center;gap:6px;background:#3b82f6;color:white;border:none;border-radius:12px;padding:10px 20px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700;box-shadow:0 2px 8px rgba(59,130,246,.4)">🔊 استماع / Listen</button>'
    +'<button onclick="swDeleteRecording()" style="display:flex;align-items:center;gap:6px;background:#f59e0b;color:white;border:none;border-radius:12px;padding:10px 20px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🔄 إعادة المحاولة / Retry</button>'
    +'<button onclick="swConfirmRecording()" style="display:flex;align-items:center;gap:6px;background:#22c55e;color:white;border:none;border-radius:12px;padding:10px 20px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:800;box-shadow:0 2px 8px rgba(34,197,94,.4)">✅ تأكيد / Confirm</button>'
    +'</div></div>';
  } else if(q.type==='listening'){
    bodyHtml='';if(q.mediaHtml) bodyHtml+='<div class="sw-audio-player">'+q.mediaHtml+'</div>';
    if(q.ansType==='text') bodyHtml+='<textarea class="sw-textarea" style="margin-top:12px" placeholder="اكتب إجابتك / Write..." onchange="sw_answers['+qi+']=this.value;updateProgressDots()">'+(sw_answers[qi]||'')+'</textarea>';
    else if(q.ansType==='mcq') bodyHtml+='<div class="sw-mcq-options">'+(q.options||[]).map(function(o,i){return '<div class="sw-mcq-opt '+(sw_answers[qi]===i?'selected':'')+'" onclick="swSelectMcq('+i+')"><div class="sw-opt-label">'+(labels[i]||i+1)+'</div><div dir="auto" style="color:#1a1a2e;flex:1">'+(o||'—')+'</div></div>';}).join('')+'</div>';
    else if(q.ansType==='matching') bodyHtml+=buildStudentMatching(q);
  } else if(q.type==='reading'){
    if(q.mediaHtml) bodyHtml+='<div style="margin-bottom:12px">'+q.mediaHtml+'</div>';
    bodyHtml+='<div class="sw-passage" dir="auto">'+(q.passageHtml||'النص المقروء...')+'</div>';
    if(q.answerStemHtml) bodyHtml+='<div style="font-size:16px;font-weight:700;margin:14px 0;color:#1a1a2e" dir="auto">'+q.answerStemHtml+'</div>';
    if(q.ansType==='text'){
      bodyHtml+='<div><div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap"><button onclick="swToggleWriteMode(\'write\')" style="border:2px solid #1e3a8a;background:#1e3a8a;color:white;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">⌨️ كتابة / Write</button><button onclick="swToggleWriteMode(\'draw\')" style="border:2px solid #e2e8f0;background:white;color:#475569;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🎨 رسم / Draw</button><label style="border:2px solid #e2e8f0;background:white;color:#475569;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px">📷 صورة<input type="file" accept="image/*" style="display:none" onchange="swUploadImage(event)"></label></div><div id="sw-write-area"><textarea class="sw-textarea" placeholder="اكتب إجابتك..." onchange="sw_answers['+qi+']=this.value;updateProgressDots()">'+(sw_answers[qi]||'')+'</textarea></div><div id="sw-draw-area" style="display:none"><div class="sw-canvas-wrap"><div class="sw-canvas-tools"><button id="sw-pen" class="active" onclick="setDrawTool(\'pen\')">✏️ قلم / Pen</button><button id="sw-eraser" onclick="setDrawTool(\'eraser\')">🧹 ممحاة / Eraser</button><select id="sw-brush-size" onchange="setBrushSize(this.value)" style="border:1px solid #e2e8f0;border-radius:6px;padding:4px 8px;font-size:12px"><option value="2">رفيع / Thin</option><option value="5" selected>عادي</option><option value="10">سميك / Thick</option></select><input type="color" id="sw-pen-color" value="#1a1a2e" onchange="setPenColor(this.value)"><button onclick="clearCanvas()" style="color:#ef4444;font-size:12px;border:1px solid #fca5a5;border-radius:6px;padding:4px 10px">🗑 مسح</button></div><canvas id="drawingCanvas" height="240" style="width:100%;display:block;cursor:crosshair;touch-action:none"></canvas></div></div><div id="sw-upload-area" style="display:none"><div id="sw-uploaded-img-wrap" style="border:2px dashed #e2e8f0;border-radius:14px;padding:24px;text-align:center;color:#94a3b8;font-size:13px">اضغط زر الصورة أعلاه</div></div></div>';
    } else if(q.ansType==='mcq'){
      // نص السؤال بخصائص الخط المحددة
      if(q.questionText){
        var qtFont=q.qTextFont||'Tajawal';
        var qtSize=q.qTextSize||15;
        var qtColor=q.qTextColor||'#1e3a8a';
        var qtBg=q.qTextBg||'#eff6ff';
        var qtDir=q.qTextDir||'rtl';
        bodyHtml+='<div style="font-size:'+qtSize+'px;font-weight:700;color:'+qtColor+';margin-bottom:12px;padding:12px 16px;background:'+qtBg+';border-radius:12px;border-right:4px solid '+qtColor+';font-family:'+qtFont+',sans-serif" dir="'+qtDir+'">'+q.questionText+'</div>';
      }
      bodyHtml+='<div class="sw-mcq-options">'+(q.options||[]).map(function(o,i){return '<div class="sw-mcq-opt '+(sw_answers[qi]===i?'selected':'')+'" onclick="swSelectMcq('+i+')"><div class="sw-opt-label">'+(labels[i]||i+1)+'</div><div dir="auto" style="color:#1a1a2e;flex:1">'+(o||'—')+'</div></div>';}).join('')+'</div>';
    }
    else if(q.ansType==='matching') bodyHtml+=buildStudentMatching(q);
  // True/False student
  } else if(q.type==='truefalse'){
    var tfFont=q.tfFont||'Tajawal';
    var tfSize=q.tfSize||15;
    var tfColor=q.tfColor||'#1a1a2e';
    var tfBg=q.tfBg||'#f8fafc';
    var tfTrueColor=q.tfTrueColor||'#15803d';
    var tfFalseColor=q.tfFalseColor||'#b91c1c';
    bodyHtml='<div style="display:flex;flex-direction:column;gap:12px">'
      +(q.statements||[]).map(function(s,si){
        var ans=sw_answers&&sw_answers[qi]?sw_answers[qi][si]:null;
        var trueSel=ans==='true';var falseSel=ans==='false';
        return '<div style="background:'+tfBg+';border:2px solid '+(trueSel?tfTrueColor:falseSel?tfFalseColor:'#e2e8f0')+';border-radius:14px;padding:14px 18px;transition:.2s">'
          +'<div style="font-size:'+tfSize+'px;font-weight:600;color:'+tfColor+';margin-bottom:12px;line-height:1.6;font-family:'+tfFont+',sans-serif" dir="auto">'+(si+1)+'. '+(s.text||'')+'</div>'
          +'<div style="display:flex;gap:12px">'
          +'<button onclick="swTFSelect('+qi+','+si+',\'true\')" style="flex:1;padding:10px;border-radius:12px;border:2px solid '+(trueSel?tfTrueColor:'#e2e8f0')+';background:'+(trueSel?'rgba('+hexToRgb(tfTrueColor)+',.15)':'white')+';color:'+(trueSel?tfTrueColor:'#64748b')+';font-weight:800;font-size:14px;cursor:pointer;transition:.2s;font-family:Tajawal,sans-serif">✅ صواب / True</button>'
          +'<button onclick="swTFSelect('+qi+','+si+',\'false\')" style="flex:1;padding:10px;border-radius:12px;border:2px solid '+(falseSel?tfFalseColor:'#e2e8f0')+';background:'+(falseSel?'rgba('+hexToRgb(tfFalseColor)+',.15)':'white')+';color:'+(falseSel?tfFalseColor:'#64748b')+';font-weight:800;font-size:14px;cursor:pointer;transition:.2s;font-family:Tajawal,sans-serif">❌ خطأ / False</button>'
          +'</div></div>';
      }).join('')
      +'</div>';
  // Classify student — Full Drag & Drop with grid table
  } else if(q.type==='classify'){
    var cols=q.columns||[];
    var items=q.items||[];
    var numCols=Math.max(1,cols.length);
    var numRows=Math.ceil(items.length/numCols);
    if(!sw_answers[qi]){
      var initBank=items.map(function(it){return it.text;}).sort(function(){return Math.random()-.5;});
      sw_answers[qi]={placed:{},bank:initBank,confirmed:false};
      cols.forEach(function(_,ci){sw_answers[qi].placed[ci]=[];});
    }
    var bankItems=sw_answers[qi].bank||[];
    var chipBg=q.orTileBg||'#6366f1';
    var tableBg=q.clTableBg||'#f0f7ff';
    var clFs=(q.clFontSize||14)+'px';
    var clTxt=q.clTextColor||'#ffffff';

    // بناء الجدول: أعمدة × صفوف
    var tableHtml='<table style="width:100%;border-collapse:collapse;margin-bottom:14px;background:'+tableBg+';table-layout:fixed">';
    // رأس الجدول — أسماء الأعمدة
    tableHtml+='<thead><tr>';
    cols.forEach(function(col){
      tableHtml+='<th style="background:'+chipBg+';color:'+clTxt+';font-weight:900;font-size:'+clFs+';font-family:Tajawal,sans-serif;padding:10px 8px;text-align:center;border:2px solid rgba(255,255,255,.3)">'+(col||'—')+'</th>';
    });
    tableHtml+='</tr></thead><tbody>';

    // صفوف الجدول
    for(var ri=0;ri<numRows;ri++){
      tableHtml+='<tr>';
      for(var ci=0;ci<numCols;ci++){
        var placed=sw_answers[qi].placed[ci]||[];
        var cellWord=placed[ri]||null;
        tableHtml+='<td id="cl-cell-'+qi+'-'+ci+'-'+ri+'" data-qi="'+qi+'" data-ci="'+ci+'" data-ri="'+ri+'"'
          +' ondragover="clCellDragOver(event)" ondrop="clCellDrop(event,'+qi+','+ci+','+ri+')" ondragleave="clCellDragLeave(event)"'
          +' style="border:2px dashed '+(cellWord?chipBg+'88':'#cbd5e1')+';border-radius:10px;min-height:48px;height:52px;background:'+(cellWord?chipBg+'11':'#f8fafc')+';text-align:center;vertical-align:middle;padding:6px;transition:.2s;cursor:default;background:'+(cellWord?chipBg+'18':tableBg)+'">';
        if(cellWord){
          tableHtml+='<div draggable="true" id="cl-placed-'+qi+'-'+ci+'-'+ri+'" data-word="'+cellWord+'" data-from-col="'+ci+'" data-from-row="'+ri+'"'
            +' ondragstart="clPlacedDragStart(event,'+qi+','+ci+','+ri+')" onclick="clPlacedClick('+qi+','+ci+','+ri+')"'
            +' style="background:'+chipBg+';color:'+clTxt+';border-radius:8px;padding:6px 12px;font-size:'+clFs+';font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;display:inline-block;user-select:none;max-width:100%;word-break:break-word">'+cellWord+'</div>';
        }
        tableHtml+='</td>';
      }
      tableHtml+='</tr>';
    }
    tableHtml+='</tbody></table>';

    bodyHtml=
      // رأس البنك
      '<div style="background:linear-gradient(135deg,'+chipBg+','+chipBg+'cc);border-radius:16px;padding:14px;margin-bottom:14px;box-shadow:0 4px 16px '+chipBg+'44">'
      +'<div style="font-size:10px;font-weight:800;color:rgba(255,255,255,.8);font-family:Montserrat,sans-serif;letter-spacing:1px;margin-bottom:10px;text-align:center">WORD BANK / مصدر المفردات</div>'
      +'<div id="cl-sw-bank-'+qi+'" ondragover="clBankDragOver(event)" ondrop="clBankDrop(event,'+qi+')" ondragleave="clBankDragLeave(event)" style="display:flex;flex-wrap:wrap;gap:8px;min-height:44px;justify-content:center;border-radius:10px;padding:6px;border:2px dashed rgba(255,255,255,.3);transition:.2s">'
      +bankItems.map(function(w,wi){
        return '<div draggable="true" class="cl-bank-chip" id="cl-chip-'+qi+'-'+wi+'" data-word="'+w+'" data-wi="'+wi+'"'
          +' ondragstart="clBankDragStart(event,'+qi+','+wi+')"'
          +' style="background:rgba(255,255,255,.22);color:'+clTxt+';border:2px solid rgba(255,255,255,.4);border-radius:10px;padding:7px 16px;font-size:'+clFs+';font-weight:700;cursor:grab;user-select:none;font-family:Tajawal,sans-serif;transition:.2s">'+w+'</div>';
      }).join('')
      +(bankItems.length===0?'<span style="color:rgba(255,255,255,.6);font-size:13px;font-family:Montserrat,sans-serif">✅ All words placed</span>':'')
      +'</div></div>'
      // الجدول
      +tableHtml
      // أزرار
      +'<div style="display:flex;gap:10px;justify-content:center">'
      +'<button onclick="clSwClear('+qi+')" style="background:rgba(239,68,68,.1);color:#ef4444;border:2px solid rgba(239,68,68,.3);border-radius:14px;padding:10px 22px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🗑 مسح الكل / Clear</button>'
      +'</div>'
      +'<div id="cl-sw-result-'+qi+'" style="display:none;margin-top:10px;text-align:center;padding:12px;border-radius:14px;font-size:15px;font-weight:800"></div>';
  // Writing Skill = نفس Reading
  } else if(q.type==='writing'||q.type==='writingskill'){
    bodyHtml='';
    if(q.passageHtml) bodyHtml+='<div class="sw-passage" dir="auto" style="margin-bottom:14px">'+q.passageHtml+'</div>';
    if(q.answerStemHtml) bodyHtml+='<div style="font-size:16px;font-weight:700;margin:12px 0;color:#1a1a2e" dir="auto">'+q.answerStemHtml+'</div>';
    bodyHtml+='<div>'
      +'<div style="display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap">'
      +'<button onclick="swToggleWriteMode(\'write\')" style="border:2px solid #1e3a8a;background:#1e3a8a;color:white;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">⌨️ كتابة / Write</button>'
      +'<button onclick="swToggleWriteMode(\'draw\')" style="border:2px solid #e2e8f0;background:white;color:#475569;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🎨 رسم / Draw</button>'
      +'<label style="border:2px solid #e2e8f0;background:white;color:#475569;border-radius:10px;padding:8px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700;display:inline-flex;align-items:center;gap:6px">📷 صورة<input type="file" accept="image/*" style="display:none" onchange="swUploadImage(event)"></label>'
      +'</div>'
      +'<div id="sw-write-area"><textarea class="sw-textarea" style="min-height:160px" placeholder="اكتب إجابتك هنا..." dir="auto" onchange="sw_answers['+qi+']=this.value;updateProgressDots()">'+(typeof sw_answers[qi]==='string'?sw_answers[qi]:'')+'</textarea></div>'
      +'<div id="sw-draw-area" style="display:none"><div class="sw-canvas-wrap"><div class="sw-canvas-tools"><button id="sw-pen" class="active" onclick="setDrawTool(\'pen\')">✏️ قلم</button><button id="sw-eraser" onclick="setDrawTool(\'eraser\')">🧹 ممحاة</button><select id="sw-brush-size" onchange="setBrushSize(this.value)" style="border:1px solid #e2e8f0;border-radius:6px;padding:4px 8px;font-size:12px"><option value="2">رفيع</option><option value="5" selected>عادي</option><option value="10">سميك</option></select><input type="color" id="sw-pen-color" value="#1a1a2e" onchange="setPenColor(this.value)"><button onclick="clearCanvas()" style="color:#ef4444;font-size:12px;border:1px solid #fca5a5;border-radius:6px;padding:4px 10px">🗑</button></div><canvas id="drawingCanvas" height="240" style="width:100%;display:block;cursor:crosshair;touch-action:none"></canvas></div></div>'
      +'<div id="sw-upload-area" style="display:none"><div id="sw-uploaded-img-wrap" style="border:2px dashed #e2e8f0;border-radius:14px;padding:24px;text-align:center;color:#94a3b8;font-size:13px">اضغط زر الصورة أعلاه</div></div>'
      +'</div>';
  }
  return bodyHtml;
}

// ── MODE 3: Modern Stream ──
function renderStreamMode(){
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[]):(d.questions||[]);
  var labels=['A','B','C','D','E','F'];
  var total=questions.length;
  document.getElementById('sw-q-counter').textContent='';
  document.getElementById('sw-q-counter-top').textContent=total+' سؤال / questions';
  var footer=document.getElementById('sw-footer');if(footer)footer.style.display='none';
  var html='<div style="padding-bottom:20px">';
  html+='<div style="text-align:center;padding:24px 20px;background:linear-gradient(135deg,#1e3a8a,#7e22ce);border-radius:22px;margin-bottom:28px;box-shadow:0 8px 28px rgba(30,58,138,.25)">';
  html+='<div style="width:64px;height:64px;border-radius:50%;background:rgba(255,255,255,.15);display:flex;align-items:center;justify-content:center;margin:0 auto 12px;border:2px solid rgba(255,255,255,.3)">'+(d.iconSrc?'<img src="'+d.iconSrc+'" style="width:44px;height:44px;object-fit:contain;border-radius:50%">':'<span style="font-size:28px;color:white">📊</span>')+'</div>';
  html+='<div style="font-size:20px;font-weight:900;color:white;margin-bottom:4px">'+(d.nameAr||('المجال '+(sw_domainIdx+1)))+'</div>';
  if(d.nameEn) html+='<div style="font-size:13px;color:rgba(255,255,255,.7);font-family:Montserrat,sans-serif;margin-bottom:8px">'+d.nameEn+'</div>';
  html+='<div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap">'
    +'<span style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 14px;font-size:12px;color:white;font-family:Montserrat,sans-serif">'+d.weight+'%</span>'
    +'<span style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 14px;font-size:12px;color:white;font-family:Montserrat,sans-serif">⏱ '+(d.time||0)+' min</span>'
    +'<span style="background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);border-radius:20px;padding:4px 14px;font-size:12px;color:white;font-family:Montserrat,sans-serif">'+total+' questions</span>'
    +'</div>';
  html+='</div>';
  questions.forEach(function(q,qi){
    var score=Number(q.score||(d.weight/Math.max(total,1))).toFixed(2);
    var isAnswered=sw_answers[qi]!==undefined&&sw_answers[qi]!==null;
    // فاصل ذهبي بين الأسئلة
    if(qi>0) html+='<div class="stream-q-separator"></div>';
    html+='<div class="stream-q-card'+(isAnswered?' stream-q-answered':'')+'" id="stream-q-'+qi+'">';
    // شارة الإجابة
    html+='<div class="stream-q-answered-badge" style="display:none;position:absolute;top:12px;left:12px;background:#22c55e;color:white;border-radius:8px;padding:3px 10px;font-size:11px;font-weight:800;font-family:Montserrat,sans-serif">✓ تمت الإجابة</div>';
    html+='<div style="display:flex;align-items:flex-start;gap:14px;margin-bottom:18px;padding-bottom:14px;border-bottom:2px solid #f1f5f9">';
    html+='<div style="min-width:42px;height:42px;border-radius:14px;background:'+(isAnswered?'linear-gradient(135deg,#22c55e,#15803d)':'linear-gradient(135deg,#1e3a8a,#7e22ce)')+';display:flex;align-items:center;justify-content:center;color:white;font-weight:800;font-size:14px;font-family:Montserrat,sans-serif;flex-shrink:0;transition:.3s" id="stream-badge-'+qi+'">'+(isAnswered?'✓':('Q.'+(qi+1)))+'</div>';
    html+='<div style="flex:1;font-size:16px;font-weight:600;line-height:1.7;color:#1a1a2e" dir="auto">'+(q.stemHtml||'Question')+'</div>';
    html+='<div style="background:#fef9c3;color:#92400e;border:1px solid #fde68a;border-radius:8px;padding:3px 10px;font-size:11px;font-weight:700;font-family:Montserrat,sans-serif;white-space:nowrap;flex-shrink:0">'+score+'%</div>';
    html+='</div>';
    if(q.mediaHtml&&q.type!=='listening'&&q.type!=='reading'&&q.type!=='speaking'&&q.type!=='oral'&&(!(q.mediaVisible&&q.mediaVisible.img===false))) html+='<div style="margin-bottom:16px">'+q.mediaHtml+'</div>';
    // MCQ
    if(q.type==='mcq'){
      var isH=q.mcqDir==='horizontal';
      if(q.bodyHtml) html+='<div style="font-size:15px;line-height:1.8;color:#1a1a2e;margin-bottom:14px" dir="auto">'+q.bodyHtml+'</div>';
      if(q.questionText) html+='<div style="font-size:16px;font-weight:700;color:#1e3a8a;margin-bottom:14px;padding:10px 14px;background:#eff6ff;border-radius:10px;border-right:4px solid #3b82f6" dir="auto">'+q.questionText+'</div>';
      html+='<div class="sw-mcq-options'+(isH?' horizontal':'')+'">';
      (q.options||[]).forEach(function(o,i){
        html+='<div class="sw-mcq-opt'+(isH?' horizontal':'')+(sw_answers[qi]===i?' selected':'')+'" id="sopt-'+qi+'-'+i+'" onclick="swSelectMcqStream('+qi+','+i+')">'
          +'<div class="sw-opt-label">'+(labels[i]||i+1)+'</div>'
          +'<div style="flex:1;font-size:15px;color:#1a1a2e;line-height:1.6" dir="auto">'+(o||'—')+'</div>'
          +'</div>';
      });
      html+='</div>';
    // True/False
    } else if(q.type==='truefalse'){
      html+='<div style="display:flex;flex-direction:column;gap:10px">';
      (q.statements||[]).forEach(function(s,si){
        var ans=sw_answers[qi]?sw_answers[qi][si]:null;
        html+='<div style="background:#f8fafc;border:2px solid #e2e8f0;border-radius:14px;padding:14px 18px">'
          +'<div style="font-size:15px;font-weight:600;color:#1a1a2e;margin-bottom:12px;line-height:1.6" dir="auto">'+(si+1)+'. '+(s.text||'')+'</div>'
          +'<div style="display:flex;gap:12px">'
          +'<button onclick="swTFStreamSelect('+qi+','+si+',\'true\')" style="flex:1;padding:10px;border-radius:12px;border:2px solid '+(ans==='true'?'#22c55e':'#e2e8f0')+';background:'+(ans==='true'?'#dcfce7':'white')+';color:'+(ans==='true'?'#15803d':'#64748b')+';font-weight:800;font-size:14px;cursor:pointer;font-family:Tajawal,sans-serif">✅ صواب</button>'
          +'<button onclick="swTFStreamSelect('+qi+','+si+',\'false\')" style="flex:1;padding:10px;border-radius:12px;border:2px solid '+(ans==='false'?'#ef4444':'#e2e8f0')+';background:'+(ans==='false'?'#fee2e2':'white')+';color:'+(ans==='false'?'#b91c1c':'#64748b')+';font-weight:800;font-size:14px;cursor:pointer;font-family:Tajawal,sans-serif">❌ خطأ</button>'
          +'</div></div>';
      });
      html+='</div>';
    // Classify
    } else if(q.type==='classify'){
      var cols=q.columns||[];var items2=q.items||[];
      var numC=Math.max(1,cols.length);var numR=Math.ceil(items2.length/numC);
      if(!sw_answers[qi]){var initBk=items2.map(function(it){return it.text;}).sort(function(){return Math.random()-.5;});sw_answers[qi]={placed:{},bank:initBk,confirmed:false};cols.forEach(function(_,ci){sw_answers[qi].placed[ci]=[];});}
      var clBg=q.orTileBg||'#6366f1';
      html+='<div style="background:linear-gradient(135deg,'+clBg+','+clBg+'cc);border-radius:12px;padding:12px;margin-bottom:12px"><div style="display:flex;flex-wrap:wrap;gap:6px;justify-content:center" id="cl-sw-bank-'+qi+'">';
      (sw_answers[qi].bank||[]).forEach(function(w,wi){html+='<div draggable="true" id="cl-chip-stream-'+qi+'-'+wi+'" ondragstart="clBankDragStart(event,'+qi+','+wi+')" style="background:rgba(255,255,255,.2);color:white;border:2px solid rgba(255,255,255,.4);border-radius:8px;padding:6px 14px;font-size:13px;font-weight:700;cursor:grab;font-family:Tajawal,sans-serif">'+w+'</div>';});
      html+='</div></div>';
      html+='<table style="width:100%;border-collapse:collapse;table-layout:fixed"><thead><tr>';
      cols.forEach(function(c){html+='<th style="background:'+clBg+';color:white;padding:8px;font-family:Tajawal,sans-serif;border:2px solid white">'+c+'</th>';});
      html+='</tr></thead><tbody>';
      for(var rr=0;rr<numR;rr++){html+='<tr>';for(var cc=0;cc<numC;cc++){var pw2=(sw_answers[qi].placed[cc]||[])[rr]||null;html+='<td id="cl-cell-'+qi+'-'+cc+'-'+rr+'" ondragover="clCellDragOver(event)" ondrop="clCellDrop(event,'+qi+','+cc+','+rr+')" ondragleave="clCellDragLeave(event)" style="border:2px dashed '+(pw2?clBg+'88':'#cbd5e1')+';border-radius:8px;min-height:44px;text-align:center;padding:4px;background:'+(pw2?clBg+'11':'#f8fafc')+'">'+(pw2?'<div draggable="true" ondragstart="clPlacedDragStart(event,'+qi+','+cc+','+rr+')" onclick="clSwCellReturn('+qi+','+cc+','+rr+')" style="background:'+clBg+';color:white;border-radius:6px;padding:5px 10px;font-size:12px;font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;display:inline-block">'+pw2+' ✕</div>':'<span style="color:#cbd5e1;font-size:11px">'+(rr+1)+'</span>')+'</td>';}html+='</tr>';}
      html+='</tbody></table>';
    // Ordering
    } else if(q.type==='ordering'){
      var orGroups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words,answerOrder:q.answerOrder}]:[]);
      if(!sw_answers[qi]||!sw_answers[qi].groups){
        sw_answers[qi]={groups:orGroups.map(function(g){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});return {placed:new Array((g.words||[]).length).fill(null),bank:bank};})};
      }
      orGroups.forEach(function(g,gi){
        if(!sw_answers[qi].groups[gi]){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});sw_answers[qi].groups[gi]={placed:new Array((g.words||[]).length).fill(null),bank:bank};}
      });
      var sFz=(q.orFontSize||15)+'px';var sFont=q.orFont||'Tajawal';var sTBg=q.orTileBg||'#1e3a8a';var sTClr=q.orTileColor||'#ffffff';var sABg=q.orAnsBg||'#f0f7ff';
      var sPerGroupScore=orGroups.length>1?(Number(q.score||0)/orGroups.length).toFixed(2):null;
      orGroups.forEach(function(g,gi){
        var gAns=sw_answers[qi].groups[gi];
        if(orGroups.length>1){
          html+='<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px"><span style="font-size:13px;font-weight:800;color:'+sTBg+';font-family:Tajawal,sans-serif">📝 الجملة '+(gi+1)+' / Sentence '+(gi+1)+'</span>'+(sPerGroupScore?'<span style="font-size:11px;font-weight:700;color:'+sTBg+';background:'+sTBg+'18;border-radius:20px;padding:3px 12px;font-family:Montserrat,sans-serif">'+sPerGroupScore+'%</span>':'')+'</div>';
        }
        html+='<div style="display:flex;flex-wrap:wrap;gap:8px;padding:14px;background:linear-gradient(135deg,'+sTBg+'22,'+sTBg+'11);border:2px solid '+sTBg+'44;border-radius:14px;min-height:50px;justify-content:center;margin-bottom:12px" id="sor-bank-'+qi+'-'+gi+'">';
        (gAns.bank||[]).forEach(function(w,wi){html+='<div draggable="true" ondragstart="sorBankDragStart(event,'+qi+','+gi+','+wi+')" style="background:'+sTBg+';color:'+sTClr+';border-radius:10px;padding:8px 18px;font-size:'+sFz+';font-weight:800;cursor:grab;font-family:'+sFont+',sans-serif;user-select:none">'+w+'</div>';});
        html+='</div>';
        html+='<div style="display:flex;flex-wrap:wrap;gap:8px;padding:14px;background:'+sABg+';border:2px dashed '+sTBg+'66;border-radius:14px;min-height:50px;justify-content:center;margin-bottom:'+(gi<orGroups.length-1?'18px':'0')+'" id="sor-slots-'+qi+'-'+gi+'">';
        (gAns.placed||[]).forEach(function(pw,si){html+='<div id="sor-slot-stream-'+qi+'-'+gi+'-'+si+'" data-slot="'+si+'" ondragover="sorSlotDragOver(event)" ondrop="sorSlotDrop(event,'+qi+','+gi+','+si+')" ondragleave="sorSlotDragLeave(event,'+qi+','+gi+','+si+')" style="min-width:65px;height:44px;border:2px dashed '+(pw?sTBg:'#cbd5e1')+';border-radius:10px;display:flex;align-items:center;justify-content:center;background:'+(pw?sTBg+'18':'white')+';padding:0 6px">'+(pw?'<div onclick="sorReturn('+qi+','+gi+','+si+')" style="background:'+sTBg+';color:'+sTClr+';border-radius:8px;padding:6px 12px;font-size:'+sFz+';font-weight:800;cursor:pointer;font-family:'+sFont+',sans-serif">'+pw+' ✕</div>':'<span style="color:#cbd5e1;font-size:11px">'+(si+1)+'</span>')+'</div>';});
        html+='</div>';
      });
      html+='<div style="display:flex;gap:10px;justify-content:center;margin-top:14px;flex-wrap:wrap">'
        +'<button onclick="sorClearAll('+qi+')" style="background:rgba(239,68,68,.1);color:#ef4444;border:2px solid rgba(239,68,68,.3);border-radius:14px;padding:9px 20px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🔄 إعادة المحاولة / Try Again</button>'
        +'<button onclick="swOrderConfirm('+qi+')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:2px solid #22c55e;border-radius:14px;padding:9px 24px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">✅ تأكيد / Confirm</button>'
        +'</div>';
    } else if(q.type==='speaking'||q.type==='oral'){
      if(q.mediaHtml&&!(q.mediaVisible&&q.mediaVisible.img===false)){
        html+='<div style="margin-bottom:14px;border-radius:14px;overflow:hidden">'+q.mediaHtml+'</div>';
      }
      if(q.type==='oral'&&q.oralText){
        var oFs=(q.speakingSize||18)+'px';var oFont=q.speakingFont||'Tajawal';var oClr=q.speakingColor||'#1a1a2e';var oBg=q.speakingBg||'#fffbeb';
        html+='<div style="font-size:'+oFs+';color:'+oClr+';background:'+oBg+';padding:20px 24px;border-radius:16px;border:3px solid rgba(250,204,21,.4);margin-bottom:16px;line-height:2.2;font-family:'+oFont+',sans-serif" dir="auto">'+q.oralText+'</div>';
      }
      html+='<div style="color:#64748b;font-size:14px;margin-bottom:12px;padding:10px 14px;background:#f8fafc;border-radius:10px" dir="auto">'+(q.speakingInst||(q.type==='oral'?'📖 اقرأ النص بصوت واضح':'🎙 سجّل إجابتك'))+'</div>';
      html+='<div style="background:#f0f9ff;border:2px solid #bae6fd;border-radius:14px;padding:16px;text-align:center">';
      html+='<button id="sRecBtn-'+qi+'" onclick="sRecToggle('+qi+')" style="background:linear-gradient(135deg,#ef4444,#dc2626);color:white;border:none;border-radius:14px;padding:10px 24px;font-size:14px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🎙 ابدأ التسجيل / Record</button>';
      html+='<div id="sRecSt-'+qi+'" style="font-size:12px;color:#64748b;margin-top:6px;font-family:Montserrat,sans-serif;min-height:16px"></div>';
      html+='<audio id="sRecAud-'+qi+'" controls style="display:none;width:100%;margin-top:8px;border-radius:12px;direction:ltr"></audio>';
      html+='<div id="sRecCtrl-'+qi+'" style="display:none;gap:10px;margin-top:12px;flex-wrap:wrap;justify-content:center">';
      html+='<button onclick="sRecListen('+qi+')" style="background:#3b82f6;color:white;border:none;border-radius:10px;padding:8px 18px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🔊 استماع / Listen</button>';
      html+='<button onclick="sRecDelete('+qi+')" style="background:#f59e0b;color:white;border:none;border-radius:10px;padding:8px 18px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🔄 إعادة / Retry</button>';
      html+='<button onclick="sRecConfirm('+qi+')" style="background:#22c55e;color:white;border:none;border-radius:10px;padding:8px 18px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:800">✅ تأكيد / Confirm</button>';
      html+='</div></div>';
    // Listening
    } else if(q.type==='listening'){
      if(q.mediaHtml) html+='<div class="sw-audio-player" style="margin-bottom:12px">'+q.mediaHtml+'</div>';
      if(q.ansType==='mcq'){
        html+='<div class="sw-mcq-options">';
        (q.options||[]).forEach(function(o,i){
          html+='<div class="sw-mcq-opt'+(sw_answers[qi]===i?' selected':'')+'" id="sopt-'+qi+'-'+i+'" onclick="swSelectMcqStream('+qi+','+i+')">'
            +'<div class="sw-opt-label">'+(labels[i]||i+1)+'</div>'
            +'<div dir="auto" style="color:#1a1a2e;flex:1">'+(o||'—')+'</div></div>';
        });
        html+='</div>';
      } else {
        html+='<textarea class="sw-textarea" style="margin-top:10px" placeholder="اكتب إجابتك / Write..." onchange="sw_answers['+qi+']=this.value;streamMarkAnswered('+qi+')">'+(typeof sw_answers[qi]==='string'?sw_answers[qi]:'')+'</textarea>';
      }
    // Reading
    } else if(q.type==='reading'){
      if(q.mediaHtml) html+='<div style="margin-bottom:12px">'+q.mediaHtml+'</div>';
      html+='<div class="sw-passage" dir="auto" style="margin-bottom:14px">'+(q.passageHtml||'')+'</div>';
      if(q.answerStemHtml) html+='<div style="font-size:15px;font-weight:700;margin-bottom:12px;color:#1a1a2e" dir="auto">'+q.answerStemHtml+'</div>';
      if(q.ansType==='mcq'){
        // نص السؤال بخصائص الخط
        if(q.questionText){
          var rqtFont=q.qTextFont||'Tajawal';var rqtSize=q.qTextSize||15;var rqtColor=q.qTextColor||'#1e3a8a';var rqtBg=q.qTextBg||'#eff6ff';var rqtDir=q.qTextDir||'rtl';
          html+='<div style="font-size:'+rqtSize+'px;font-weight:700;color:'+rqtColor+';margin-bottom:12px;padding:12px 16px;background:'+rqtBg+';border-radius:12px;border-right:4px solid '+rqtColor+';font-family:'+rqtFont+',sans-serif" dir="'+rqtDir+'">'+q.questionText+'</div>';
        }
        html+='<div class="sw-mcq-options">';
        (q.options||[]).forEach(function(o,i){
          html+='<div class="sw-mcq-opt'+(sw_answers[qi]===i?' selected':'')+'" id="sopt-'+qi+'-'+i+'" onclick="swSelectMcqStream('+qi+','+i+')">'
            +'<div class="sw-opt-label">'+(labels[i]||i+1)+'</div>'
            +'<div dir="auto" style="color:#1a1a2e;flex:1">'+(o||'—')+'</div></div>';
        });
        html+='</div>';
      } else {
        html+='<textarea class="sw-textarea" placeholder="اكتب إجابتك / Write..." onchange="sw_answers['+qi+']=this.value;streamMarkAnswered('+qi+')">'+(typeof sw_answers[qi]==='string'?sw_answers[qi]:'')+'</textarea>';
      }
    // Writing
    } else if(q.type==='writing'){
      if(q.passageHtml) html+='<div class="sw-passage" dir="auto" style="margin-bottom:14px">'+q.passageHtml+'</div>';
      if(q.answerStemHtml) html+='<div style="font-size:15px;font-weight:700;margin-bottom:12px;color:#1a1a2e" dir="auto">'+q.answerStemHtml+'</div>';
      var wMd=q.writingMode||'free';var wMn=parseInt(q.writingMin)||0;var wMx=parseInt(q.writingMax)||0;
      var wHnt=wMd==='words'?'كلمة / words':wMd==='sentences'?'جملة / sentences':'';
      var wLmt=wMn&&wMx?'('+wMn+' – '+wMx+' '+wHnt+')':wMn?'('+wMn+'+ '+wHnt+')':wMx?'(max '+wMx+' '+wHnt+')':'';
      if(wLmt) html+='<div style="font-size:12px;color:#64748b;margin-bottom:8px;font-family:Montserrat,sans-serif">'+wLmt+'</div>';
      html+='<textarea class="sw-textarea" style="min-height:120px" placeholder="اكتب إجابتك هنا..." dir="auto" onchange="sw_answers['+qi+']=this.value;streamMarkAnswered('+qi+')">'+(typeof sw_answers[qi]==='string'?sw_answers[qi]:'')+'</textarea>';
    }
    html+='</div>';
  });
  html+='<div style="text-align:center;padding:20px 16px 32px">';
  html+='<div style="font-size:13px;color:#64748b;margin-bottom:12px;font-family:Montserrat,sans-serif">أجب على جميع الأسئلة قبل التسليم / Answer all before submitting</div>';
  html+='<button onclick="swSubmitDomain('+total+')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:3px solid #86efac;border-radius:16px;padding:16px 52px;font-size:17px;font-weight:800;cursor:pointer;box-shadow:0 6px 20px rgba(34,197,94,.4);font-family:Tajawal,sans-serif">✅ تسليم المجال / Submit Domain</button>';
  html+='</div></div>';
  document.getElementById('sw-questions-content').innerHTML=html;
  questions.forEach(function(q,qi){ if(q.type==='matching') smStreamInit(qi); });
  updateProgressDots();
}
// Stream Matching
var smStreamSel={};var smStreamConn={};
function smStreamInit(qi){smStreamConn[qi]=sw_answers[qi]&&sw_answers[qi].connections?Object.assign({},sw_answers[qi].connections):{};smStreamDraw(qi);}
function smStreamClick(col,qi,idx){
  if(!smStreamConn[qi]) smStreamConn[qi]={};
  if(col==='A'){
    smStreamSel[qi]={col:'A',idx:idx};
    document.querySelectorAll('[id^="smA-'+qi+'-"],[id^="smB-'+qi+'-"]').forEach(function(el){el.style.borderColor='#e2e8f0';el.style.background='#fafafa';});
    var el=document.getElementById('smA-'+qi+'-'+idx);
    if(el){el.style.borderColor='#1e3a8a';el.style.background='#eff6ff';}
  } else if(col==='B'&&smStreamSel[qi]&&smStreamSel[qi].col==='A'){
    var aIdxConfirm2=smStreamSel[qi].idx;
    smStreamConn[qi][smStreamSel[qi].idx]=idx;
    smStreamSel[qi]=null;
    sw_answers[qi]={connections:Object.assign({},smStreamConn[qi])};
    streamMarkAnswered(qi);updateProgressDots();smStreamDraw(qi);
    setTimeout(function(){
      var itemA=document.getElementById('smA-'+qi+'-'+aIdxConfirm2),itemB=document.getElementById('smB-'+qi+'-'+idx);
      [itemA,itemB].forEach(function(el){
        if(!el)return;
        el.style.transition='transform .15s';
        el.style.transform='scale(1.04)';
        setTimeout(function(){el.style.transform='scale(1)';},150);
      });
    },20);
  }
}
function smStreamClear(qi){smStreamConn[qi]={};delete sw_answers[qi];smStreamDraw(qi);updateProgressDots();}
function smStreamDraw(qi){
  var lineColors=['#ef4444','#3b82f6','#22c55e','#f59e0b','#8b5cf6','#ec4899','#0891b2','#f97316'];
  var svg=document.getElementById('smSvg-'+qi);if(!svg)return;svg.innerHTML='';
  var wrap=document.getElementById('smWrap-'+qi);if(!wrap)return;
  var wr=wrap.getBoundingClientRect();
  // في نمط الورقة البيضاء يُطبَّق transform:scale() على #wp-page، ما يجعل getBoundingClientRect
  // يُرجع إحداثيات مُكبَّرة بالفعل بصريًا. نُعيد القسمة على معامل الزووم لتحويلها لوحدات SVG الداخلية
  // الحقيقية (ما قبل التكبير)، وإلا يحدث تكبير مضاعف يُسبب طفو الخط في مكان خاطئ عند أي زووم غير 100%.
  var zf=(testData.displayMode===4&&_wpZoomLevel)?_wpZoomLevel:1;
  Object.keys(smStreamConn[qi]||{}).forEach(function(ai){
    var bi=smStreamConn[qi][ai];
    var clr=lineColors[(parseInt(ai)||0)%lineColors.length];
    var dA=document.getElementById('smDA-'+qi+'-'+ai);
    var dB=document.getElementById('smDB-'+qi+'-'+bi);
    var iA=document.getElementById('smA-'+qi+'-'+ai);
    var iB=document.getElementById('smB-'+qi+'-'+bi);
    if(!dA||!dB)return;
    dA.style.background=clr;dA.style.borderColor=clr;
    dB.style.background=clr;dB.style.borderColor=clr;
    if(iA){iA.style.borderColor=clr;iA.style.background=clr+'18';}
    if(iB){iB.style.borderColor=clr;iB.style.background=clr+'18';}
    var rA=dA.getBoundingClientRect(),rB=dB.getBoundingClientRect();
    var x1=(rA.left+7-wr.left)/zf,y1=(rA.top+7-wr.top)/zf;
    var x2=(rB.left+7-wr.left)/zf,y2=(rB.top+7-wr.top)/zf;
    var p=document.createElementNS('http://www.w3.org/2000/svg','path');
    p.setAttribute('d','M'+x1+','+y1+' L'+x2+','+y2);
    p.setAttribute('stroke',clr);p.setAttribute('stroke-width','3');p.setAttribute('fill','none');
    p.setAttribute('stroke-linecap','round');
    svg.appendChild(p);
  });
}
// Stream Ordering
function sOrdPlace(qi,el){
  var w=el.textContent;
  if(!sw_answers[qi]) return;
  var bi=sw_answers[qi].bank.indexOf(w);if(bi<0)return;
  sw_answers[qi].bank.splice(bi,1);
  sw_answers[qi].order.push(w);
  streamMarkAnswered(qi);updateProgressDots();_sOrdRefresh(qi);
}
function sOrdRemove(qi,wi){
  if(!sw_answers[qi])return;
  var w=sw_answers[qi].order[wi];
  sw_answers[qi].order.splice(wi,1);
  sw_answers[qi].bank.push(w);
  updateProgressDots();_sOrdRefresh(qi);
}
function _sOrdRefresh(qi){
  var zone=document.getElementById('sOrdZone-'+qi);
  var bank=document.getElementById('sOrdBank-'+qi);
  if(!zone||!bank||!sw_answers[qi])return;
  var ord=sw_answers[qi].order,bnk=sw_answers[qi].bank;
  zone.innerHTML=ord.length?ord.map(function(w,wi){return '<div style="background:#1e3a8a;color:white;border-radius:10px;padding:8px 14px;font-size:14px;font-weight:700;cursor:pointer" onclick="sOrdRemove('+qi+','+wi+')">'+w+'</div>';}).join(''):'<span style="color:#94a3b8;font-size:13px;font-family:Montserrat,sans-serif">اضغط الكلمات أدناه...</span>';
  bank.innerHTML=bnk.length?bnk.map(function(w){return '<div style="background:white;color:#1e3a8a;border:2px solid #1e3a8a;border-radius:10px;padding:8px 14px;font-size:14px;font-weight:700;cursor:pointer" onclick="sOrdPlace('+qi+',this)">'+w+'</div>';}).join(''):'<span style="color:#94a3b8;font-size:13px;">✅ كل الكلمات وُضعت</span>';
}
// Stream Speaking
var sRecData={};
function sRecToggle(qi){
  var btn=document.getElementById('sRecBtn-'+qi);
  var st=document.getElementById('sRecSt-'+qi);
  if(!btn)return;
  if(!sRecData[qi]||!sRecData[qi].recording){
    navigator.mediaDevices&&navigator.mediaDevices.getUserMedia({audio:true}).then(function(stream){
      var chunks=[];
      var mr=new MediaRecorder(stream);
      mr.ondataavailable=function(e){chunks.push(e.data);};
      mr.onstop=function(){
        var blob=new Blob(chunks,{type:'audio/webm'});
        var url=URL.createObjectURL(blob);
        var aud=document.getElementById('sRecAud-'+qi);
        if(aud){aud.src=url;aud.style.display='block';}
        var ctrl=document.getElementById('sRecCtrl-'+qi);
        if(ctrl)ctrl.style.display='flex';
      };
      mr.start();
      sRecData[qi]={mr:mr,recording:true,sec:0};
      sRecData[qi].iv=setInterval(function(){sRecData[qi].sec++;if(st)st.textContent='⏱ '+sRecData[qi].sec+'s';},1000);
      btn.style.background='linear-gradient(135deg,#ef4444,#dc2626)';
      btn.innerHTML='⏹ إيقاف / Stop';
    }).catch(function(){if(st)st.textContent='⚠️ لا يمكن الوصول للميكروفون';});
  } else {
    clearInterval(sRecData[qi].iv);
    sRecData[qi].mr.stop();
    sRecData[qi].recording=false;
    btn.style.background='linear-gradient(135deg,#db2777,#7e22ce)';
    btn.innerHTML='🎙 ابدأ التسجيل / Start';
    if(st)st.textContent='✅ تم التسجيل';
  }
}
function sRecListen(qi){var aud=document.getElementById('sRecAud-'+qi);if(aud&&aud.src){aud.style.display='block';aud.play();}else{scWarn('لم يتم تسجيل صوت بعد، سجّل أولاً','No recording yet, please record first');}}
function sRecConfirm(qi){sw_answers[qi]='recorded';streamMarkAnswered(qi);updateProgressDots();var c=document.getElementById('sRecCtrl-'+qi);if(c)c.style.display='none';var s=document.getElementById('sRecSt-'+qi);if(s)s.textContent='✅ تم التأكيد / Confirmed';}
function sRecDelete(qi){delete sw_answers[qi];var a=document.getElementById('sRecAud-'+qi);if(a){a.src='';a.style.display='none';}var c=document.getElementById('sRecCtrl-'+qi);if(c)c.style.display='none';sRecData[qi]=null;updateProgressDots();}

function swSelectMcqStream(qi,idx){
  sw_answers[qi]=idx;
  // Update visuals for this question's options
  var opts=document.querySelectorAll('[id^="sopt-'+qi+'-"]');
  opts.forEach(function(el,i){el.classList.toggle('selected',i===idx);});
  streamMarkAnswered(qi);
  updateProgressDots();
}
function streamMarkAnswered(qi){
  var card=document.getElementById('stream-q-'+qi);
  if(!card) return;
  card.classList.add('stream-q-answered');
  var badge=card.querySelector('.sw-q-badge,[style*="border-radius:14px"]');
  var firstDiv=card.querySelector('div[style*="min-width:42px"]');
  if(firstDiv){firstDiv.style.background='linear-gradient(135deg,#22c55e,#15803d)';firstDiv.textContent='✓';}
}

// ============================================================
// STUDENT MATCHING
// ============================================================
function buildStudentMatching(q){
  var pairs=q.pairs||[];
  var lineColors=['#ef4444','#3b82f6','#22c55e','#f59e0b','#8b5cf6','#ec4899','#0891b2','#f97316'];
  var shuffledB=pairs.map(function(_,i){return i;}).sort(function(){return Math.random()-.5;});
  var rowsHtml=pairs.map(function(p,i){
    var bi=shuffledB[i];
    var pb=pairs[bi]||{};
    var boxA='<div style="display:flex;align-items:stretch;gap:8px;height:100%" id="smRowA'+i+'">'
      +'<div style="flex:1;border:2px solid #e2e8f0;border-radius:12px;overflow:hidden;background:#fafafa;cursor:pointer;transition:.2s;display:flex;flex-direction:column;justify-content:center" id="smItemA'+i+'" onclick="smDotClick(\'A\','+i+')">'
      +(p.aImg?'<div style="width:100%;max-height:70px;overflow:hidden;display:flex;align-items:center;justify-content:center"><img src="'+p.aImg+'" style="max-width:100%;max-height:70px;object-fit:contain;display:block;transform:scale('+(p.aImgScale||1)+')"></div>':'')
      +'<div style="padding:10px 12px;font-size:14px;text-align:center" dir="auto">'+(p.aHtml||'—')+'</div></div>'
      +'<div style="display:flex;align-items:center;flex-shrink:0">'
      +'<div id="smDotA'+i+'" style="width:16px;height:16px;border-radius:50%;background:'+lineColors[i%lineColors.length]+'44;border:2.5px solid '+lineColors[i%lineColors.length]+';cursor:pointer;transition:.2s" onclick="smDotClick(\'A\','+i+')"></div>'
      +'</div></div>';
    var boxB='<div style="display:flex;align-items:stretch;gap:8px;height:100%" id="smRowB'+bi+'">'
      +'<div style="display:flex;align-items:center;flex-shrink:0">'
      +'<div id="smDotB'+bi+'" style="width:16px;height:16px;border-radius:50%;background:#e2e8f044;border:2.5px solid #cbd5e1;cursor:pointer;transition:.2s" onclick="smDotClick(\'B\','+bi+')"></div>'
      +'</div>'
      +'<div style="flex:1;border:2px solid #e2e8f0;border-radius:12px;overflow:hidden;background:#fafafa;cursor:pointer;transition:.2s;display:flex;flex-direction:column;justify-content:center" id="smItemB'+bi+'" onclick="smDotClick(\'B\','+bi+')">'
      +(pb.bImg?'<div style="width:100%;max-height:70px;overflow:hidden;display:flex;align-items:center;justify-content:center"><img src="'+pb.bImg+'" style="max-width:100%;max-height:70px;object-fit:contain;display:block;transform:scale('+(pb.bImgScale||1)+')"></div>':'')
      +'<div style="padding:10px 12px;font-size:14px;text-align:center" dir="auto">'+(pb.bHtml||'—')+'</div></div></div>';
    return boxA+boxB;
  }).join('');
  return '<div style="position:relative;overflow:visible" id="studentMatchWrap">'
    +'<svg id="smSvg" style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:5;overflow:visible" overflow="visible"></svg>'
    +'<div style="display:grid;grid-template-columns:1fr 1fr;grid-auto-rows:min-content;align-items:stretch;gap:10px 16px">'
    +rowsHtml
    +'</div>'
    +'<div style="margin-top:14px;display:flex;justify-content:center">'
    +'<button onclick="smClearAll()" style="background:#fee2e2;border:1.5px solid #fca5a5;color:#ef4444;border-radius:10px;padding:9px 22px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">✕ مسح الكل / Clear All</button>'
    +'</div>'
    +'<div id="matchConnectionsList" style="margin-top:8px"></div>'
    +'</div>';
}
function initStudentMatching(q){smSelected=null;smConnections=sw_answers[sw_qIdx]&&sw_answers[sw_qIdx].connections?sw_answers[sw_qIdx].connections:{};drawSmLines();}
function smDotClick(col,idx){
  if(col==='A'){
    smSelected={col:'A',idx:idx};
    document.querySelectorAll('[id^="smItemA"],[id^="smItemB"]').forEach(function(d){d.style.borderColor='#e2e8f0';d.style.background='#fafafa';});
    var el=document.getElementById('smItemA'+idx);if(el){el.style.borderColor='#1e3a8a';el.style.background='#eff6ff';}
  } else if(col==='B'&&smSelected&&smSelected.col==='A'){
    smConnections[smSelected.idx]=idx;
    var aIdxConfirm=smSelected.idx;
    smSelected=null;
    document.querySelectorAll('[id^="smItemA"],[id^="smItemB"]').forEach(function(d){d.style.borderColor='#e2e8f0';d.style.background='#fafafa';});
    sw_answers[sw_qIdx]={connections:Object.assign({},smConnections)};
    updateProgressDots();
    drawSmLines();
    // تأثير بصري تأكيدي قصير عند نجاح التوصيل
    setTimeout(function(){
      var itemA=document.getElementById('smItemA'+aIdxConfirm),itemB=document.getElementById('smItemB'+idx);
      [itemA,itemB].forEach(function(el){
        if(!el)return;
        el.style.transition='transform .15s';
        el.style.transform='scale(1.04)';
        setTimeout(function(){el.style.transform='scale(1)';},150);
      });
    },20);
  }
}
function drawSmLines(){
  var lineColors=['#ef4444','#3b82f6','#22c55e','#f59e0b','#8b5cf6','#ec4899','#0891b2','#f97316'];
  var svg=document.getElementById('smSvg');if(!svg)return;svg.innerHTML='';
  document.querySelectorAll('[id^="smDotA"],[id^="smDotB"]').forEach(function(d){d.style.background='';d.style.borderColor='';});
  document.querySelectorAll('[id^="smItemA"],[id^="smItemB"]').forEach(function(d){d.style.borderColor='#e2e8f0';d.style.background='#fafafa';});
  var wrap=document.getElementById('studentMatchWrap');if(!wrap)return;
  var wrapRect=wrap.getBoundingClientRect();
  var colorIdx=0;
  Object.keys(smConnections).forEach(function(aIdx){
    var bIdx=smConnections[aIdx];
    var clr=lineColors[(parseInt(aIdx)||colorIdx)%lineColors.length];
    var dotA=document.getElementById('smDotA'+aIdx),dotB=document.getElementById('smDotB'+bIdx);
    var itemA=document.getElementById('smItemA'+aIdx),itemB=document.getElementById('smItemB'+bIdx);
    if(!dotA||!dotB)return;
    dotA.style.background=clr;dotA.style.borderColor=clr;
    dotB.style.background=clr;dotB.style.borderColor=clr;
    if(itemA){itemA.style.borderColor=clr;itemA.style.background=clr+'18';}
    if(itemB){itemB.style.borderColor=clr;itemB.style.background=clr+'18';}
    var rA=dotA.getBoundingClientRect(),rB=dotB.getBoundingClientRect();
    var x1=rA.left+rA.width/2-wrapRect.left, y1=rA.top+rA.height/2-wrapRect.top;
    var x2=rB.left+rB.width/2-wrapRect.left, y2=rB.top+rB.height/2-wrapRect.top;
    var path=document.createElementNS('http://www.w3.org/2000/svg','path');
    path.setAttribute('d','M'+x1+','+y1+' L'+x2+','+y2);
    path.setAttribute('stroke',clr);path.setAttribute('stroke-width','3');path.setAttribute('fill','none');
    path.setAttribute('stroke-linecap','round');
    // إضافة shadow filter
    var defs=document.createElementNS('http://www.w3.org/2000/svg','defs');
    var filter=document.createElementNS('http://www.w3.org/2000/svg','filter');
    filter.setAttribute('id','shadow'+aIdx);
    var fe=document.createElementNS('http://www.w3.org/2000/svg','feDropShadow');
    fe.setAttribute('dx','0');fe.setAttribute('dy','1');fe.setAttribute('stdDeviation','2');fe.setAttribute('flood-color',clr);fe.setAttribute('flood-opacity','0.4');
    filter.appendChild(fe);defs.appendChild(filter);svg.appendChild(defs);
    path.setAttribute('filter','url(#shadow'+aIdx+')');
    svg.appendChild(path);
    colorIdx++;
  });
}
function smClearAll(){smConnections={};delete sw_answers[sw_qIdx];drawSmLines();updateProgressDots();}

// ============================================================
// STUDENT ORDERING
// ============================================================
function buildStudentOrdering(q){
  var groups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words,answerOrder:q.answerOrder}]:[]);
  if(!sw_answers[sw_qIdx]||!sw_answers[sw_qIdx].groups){
    sw_answers[sw_qIdx]={groups:groups.map(function(g){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});return {placed:new Array((g.words||[]).length).fill(null),bank:bank};})};
  }
  var ansGroups=sw_answers[sw_qIdx].groups;
  // تأكد أن عدد المجموعات المخزنة يطابق السؤال
  groups.forEach(function(g,gi){
    if(!ansGroups[gi]){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});ansGroups[gi]={placed:new Array((g.words||[]).length).fill(null),bank:bank};}
    var total=(g.words||[]).length;
    if(!ansGroups[gi].placed) ansGroups[gi].placed=new Array(total).fill(null);
    while(ansGroups[gi].placed.length<total) ansGroups[gi].placed.push(null);
  });
  var tileBg=q.orTileBg||'#1e3a8a';
  var tileClr=q.orTileColor||'#ffffff';
  var ansBg=q.orAnsBg||'#f0f7ff';
  var font=q.orFont||'Tajawal';
  var fontSize=(q.orFontSize||15)+'px';
  var perGroupScore=groups.length>1?(Number(q.score||0)/groups.length).toFixed(2):null;

  var html='';
  groups.forEach(function(g,gi){
    var words=g.words||[];
    var total=words.length;
    var bankWords=ansGroups[gi].bank||[];
    var placed=ansGroups[gi].placed||[];
    var groupLabel=groups.length>1?('<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:8px"><span style="font-size:13px;font-weight:800;color:'+tileBg+';font-family:Tajawal,sans-serif">📝 الجملة '+(gi+1)+' / Sentence '+(gi+1)+'</span>'+(perGroupScore?'<span style="font-size:11px;font-weight:700;color:'+tileBg+';background:'+tileBg+'18;border-radius:20px;padding:3px 12px;font-family:Montserrat,sans-serif">'+perGroupScore+'%</span>':'')+'</div>'):'';
    // بنك الكلمات
    var bankHtml='<div id="or-bank-'+gi+'" style="display:flex;flex-wrap:wrap;gap:10px;padding:16px;background:linear-gradient(135deg,'+tileBg+'22,'+tileBg+'11);border:2px solid '+tileBg+'44;border-radius:16px;min-height:56px;justify-content:center">';
    bankWords.forEach(function(w,wi){
      bankHtml+='<div draggable="true" id="or-bw-'+gi+'-'+wi+'" data-word="'+w+'" data-from="bank" data-idx="'+wi+'"'
        +' ondragstart="orDragStart(event,\'bank\','+gi+','+wi+')"'
        +' style="background:'+tileBg+';color:'+tileClr+';border-radius:12px;padding:9px 20px;font-size:'+fontSize+';font-weight:800;cursor:grab;font-family:'+font+',sans-serif;box-shadow:0 3px 10px '+tileBg+'55;user-select:none;border:2px solid '+tileBg+';transition:.2s">'+w+'</div>';
    });
    if(bankWords.length===0) bankHtml+='<span style="color:'+tileBg+';font-size:13px;font-weight:700;font-family:Montserrat,sans-serif">✅ كل الكلمات وُضعت / All placed</span>';
    bankHtml+='</div>';
    // الخانات
    var slotsHtml='<div id="or-slots-'+gi+'" style="display:flex;flex-wrap:wrap;gap:10px;padding:16px;background:'+ansBg+';border:2px dashed '+tileBg+'66;border-radius:16px;min-height:56px;justify-content:center">';
    for(var si=0;si<total;si++){
      var pw=placed[si]||null;
      slotsHtml+='<div id="or-slot-'+gi+'-'+si+'" data-slot="'+si+'"'
        +' ondragover="orSlotDragOver(event)" ondrop="orSlotDrop(event,'+gi+','+si+')" ondragleave="orSlotDragLeave(event,'+gi+','+si+')"'
        +' style="min-width:70px;height:46px;border:2px dashed '+(pw?tileBg:'#cbd5e1')+';border-radius:12px;display:flex;align-items:center;justify-content:center;background:'+(pw?tileBg+'18':'white')+';transition:.2s;padding:0 8px">';
      if(pw){
        slotsHtml+='<div draggable="true" id="or-pw-'+gi+'-'+si+'" data-word="'+pw+'" data-from="placed" data-idx="'+si+'"'
          +' ondragstart="orDragStart(event,\'placed\','+gi+','+si+')"'
          +' onclick="orReturnToBank('+gi+','+si+')"'
          +' style="background:'+tileBg+';color:'+tileClr+';border-radius:10px;padding:7px 16px;font-size:'+fontSize+';font-weight:800;cursor:pointer;font-family:'+font+',sans-serif;user-select:none;display:flex;align-items:center;gap:6px">'
          +pw+'<span style="font-size:10px;opacity:.7">✕</span></div>';
      } else {
        slotsHtml+='<span style="color:#cbd5e1;font-size:11px;font-family:Montserrat,sans-serif">'+(si+1)+'</span>';
      }
      slotsHtml+='</div>';
    }
    slotsHtml+='</div>';
    html+='<div style="margin-bottom:18px;'+(groups.length>1?'padding:14px;border:1.5px dashed '+tileBg+'40;border-radius:16px':'')+'">'+groupLabel+'<div style="margin-bottom:12px">'+bankHtml+'</div><div>'+slotsHtml+'</div></div>';
  });

  return html
    +'<div style="display:flex;gap:10px;justify-content:center;margin-top:8px;flex-wrap:wrap">'
    +'<button onclick="swOrderClearAll()" style="background:rgba(239,68,68,.1);color:#ef4444;border:2px solid rgba(239,68,68,.3);border-radius:14px;padding:10px 22px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🔄 إعادة المحاولة / Try Again</button>'
    +'<button onclick="swOrderConfirm('+sw_qIdx+')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:2px solid #22c55e;border-radius:14px;padding:10px 26px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;box-shadow:0 3px 10px rgba(34,197,94,.35)">✅ تأكيد / Confirm</button>'
    +'</div>';
}

// ── Drag & Drop للـ Ordering (يدعم عدة مجموعات/جمل داخل نفس السؤال) ──
var _orDrag={from:null,gi:-1,idx:-1,word:null};
function orDragStart(e,from,gi,idx){
  var g=sw_answers[sw_qIdx].groups[gi];
  var word=from==='bank'?(g.bank[idx]||''):(g.placed[idx]||'');
  _orDrag={from:from,gi:gi,idx:idx,word:word};
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain',word);
}
function orSlotDragOver(e){
  e.preventDefault();
  e.currentTarget.style.background='#dbeafe';
  e.currentTarget.style.borderColor='#3b82f6';
}
function orSlotDragLeave(e,gi,si){
  var g=sw_answers[sw_qIdx].groups[gi];
  var pw=g&&g.placed&&g.placed[si];
  e.currentTarget.style.background=pw?'':'white';
  e.currentTarget.style.borderColor=pw?'':'#cbd5e1';
}
function orSlotDrop(e,gi,si){
  e.preventDefault();
  if(!_orDrag.word||_orDrag.gi!==gi) return;
  var g=sw_answers[sw_qIdx].groups[gi];
  var displaced=g.placed[si]||null;
  g.placed[si]=_orDrag.word;
  if(_orDrag.from==='bank'){
    g.bank=g.bank.filter(function(_,i){return i!==_orDrag.idx;});
    if(displaced) g.bank.push(displaced);
  } else if(_orDrag.from==='placed'){
    g.placed[_orDrag.idx]=displaced||null;
  }
  _orDrag={from:null,gi:-1,idx:-1,word:null};
  updateProgressDots();
  renderStudentQuestion();
}
function orReturnToBank(gi,si){
  var g=sw_answers[sw_qIdx].groups[gi];
  if(!g||!g.placed[si]) return;
  g.bank.push(g.placed[si]);
  g.placed[si]=null;
  updateProgressDots();
  renderStudentQuestion();
}
function swOrderConfirm(qi){
  if(sw_answers[qi]) sw_answers[qi].confirmed=true;
  updateProgressDots();
  scOk('تم التأكيد ✅','Confirmed','تم تأكيد ترتيب الكلمات','Word order confirmed','✅');
}
function swOrderClearAll(){
  var d=testData.domains[sw_domainIdx];
  var q=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions[sw_qIdx]:null):(d.questions?d.questions[sw_qIdx]:null);
  if(!q) return;
  var groups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words}]:[]);
  sw_answers[sw_qIdx]={groups:groups.map(function(g){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});return {placed:new Array((g.words||[]).length).fill(null),bank:bank};})};
  updateProgressDots();renderStudentQuestion();
}
function swOrderRemovePlaced(idx){
  if(!sw_answers[sw_qIdx])return;
  var w=sw_answers[sw_qIdx].order[idx];
  sw_answers[sw_qIdx].order.splice(idx,1);
  sw_answers[sw_qIdx].bank.push(w);
  updateProgressDots();
  swOrderDropToZone({preventDefault:function(){}});
}
function orderDragStart(e){dragSrcIdx=parseInt(e.currentTarget.dataset.idx);e.currentTarget.classList.add('dragging');e.dataTransfer.effectAllowed='move';}
function orderDragOver(e){e.preventDefault();e.currentTarget.classList.add('drag-over');}
function orderDrop(e){
  e.preventDefault();
  var targetIdx=parseInt(e.currentTarget.dataset.idx);
  if(dragSrcIdx===null||dragSrcIdx===targetIdx){e.currentTarget.classList.remove('drag-over');return;}
  var tmp=orderingCurrent[dragSrcIdx];orderingCurrent[dragSrcIdx]=orderingCurrent[targetIdx];orderingCurrent[targetIdx]=tmp;
  sw_answers[sw_qIdx]={order:orderingCurrent.slice()};updateProgressDots();
  var cont=document.getElementById('orderingContainer');
  if(cont) cont.innerHTML=orderingCurrent.map(function(w,i){return '<div class="ordering-item" draggable="true" data-idx="'+i+'" id="orderItem'+i+'" ondragstart="orderDragStart(event)" ondragover="orderDragOver(event)" ondrop="orderDrop(event)" ondragend="orderDragEnd(event)"><div class="drag-handle">⠿</div><div style="flex:1;font-size:15px;color:#1a1a2e;line-height:1.6" dir="auto">'+(w||'—')+'</div></div>';}).join('');
}
function orderDragEnd(e){e.currentTarget.classList.remove('dragging');document.querySelectorAll('.ordering-item').forEach(function(el){el.classList.remove('drag-over');});}
function orderTouchStart(e){touchSrcEl=e.currentTarget;touchClone=touchSrcEl.cloneNode(true);touchClone.style.cssText='position:fixed;opacity:.8;pointer-events:none;z-index:9999;width:'+touchSrcEl.offsetWidth+'px;background:white;border:2px solid #7e22ce;border-radius:12px;padding:12px 16px;';document.body.appendChild(touchClone);}
function orderTouchMove(e){e.preventDefault();var t=e.touches[0];if(touchClone){touchClone.style.left=(t.clientX-touchClone.offsetWidth/2)+'px';touchClone.style.top=(t.clientY-20)+'px';}}
function orderTouchEnd(e){
  if(touchClone){document.body.removeChild(touchClone);touchClone=null;}
  var t=e.changedTouches[0];
  var target=document.elementFromPoint(t.clientX,t.clientY);if(target) target=target.closest('.ordering-item');
  if(target&&touchSrcEl&&target!==touchSrcEl){var sIdx=parseInt(touchSrcEl.dataset.idx),tIdx=parseInt(target.dataset.idx);var tmp=orderingCurrent[sIdx];orderingCurrent[sIdx]=orderingCurrent[tIdx];orderingCurrent[tIdx]=tmp;sw_answers[sw_qIdx]={order:orderingCurrent.slice()};updateProgressDots();var cont=document.getElementById('orderingContainer');if(cont) cont.innerHTML=orderingCurrent.map(function(w,i){return '<div class="ordering-item" draggable="true" data-idx="'+i+'" id="orderItem'+i+'" ondragstart="orderDragStart(event)" ondragover="orderDragOver(event)" ondrop="orderDrop(event)" ondragend="orderDragEnd(event)" ontouchstart="orderTouchStart(event)" ontouchmove="orderTouchMove(event)" ontouchend="orderTouchEnd(event)"><div class="drag-handle">⠿</div><div style="flex:1;font-size:15px;color:#1a1a2e;line-height:1.6" dir="auto">'+(w||'—')+'</div></div>';}).join('');}
  touchSrcEl=null;
}
function swSelectMcq(idx){sw_answers[sw_qIdx]=idx;renderStudentQuestion();}
function swConfirmInstructions(){
  var cb1=document.getElementById('sw-confirm-check');
  if(!cb1||!cb1.checked){
    scWarn('يرجى تأكيد قراءة التعليمات باللغتين أولاً','Please confirm both checkboxes before starting');
    return;
  }
  var hasMic=document.getElementById('sw-mic-check');
  if(hasMic&&!window._swMicGranted){
    scWarn('يرجى التأكد من عمل الميكروفون أولاً','Please verify your microphone is working first');
    return;
  }
  sw_instructionsConfirmed=true;
  sw_currentPage='domains';
  if(currentStudentData) swEnableFullLock();
  renderStudentWindowPage();
}
function swNavigate(dir){
  sw_navDir=dir;
  if(sw_currentPage==='instructions'){
    if(dir===1){
      if(sw_instructionsConfirmed){sw_currentPage='domains';renderStudentWindowPage();}
      else alert('اضغط زر أبدأ بعد قراءة التعليمات قبل المتابعة.\nPress Start after reading the instructions before continuing.');
    }
    return;
  }
  if(sw_currentPage==='domains'){
    if(dir===-1){sw_currentPage='instructions';renderStudentWindowPage();return;}
    return;
  }
  if(sw_currentPage==='branches'){
    if(dir===-1){sw_currentPage='domains';renderStudentWindowPage();return;}
    return;
  }
  // Mode 3: stream — footer is hidden, submit is in-page
  var mode=testData.displayMode||1;
  if(mode===3){
    if(dir===-1){
      if(sw_branchIdx>=0){sw_currentPage='branches';}else{sw_currentPage='domains';}
      stopStudentTimer();renderStudentWindowPage();return;
    }
    return;
  }
  // Modes 1 & 2: navigate individual questions
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches[sw_branchIdx].questions||[]):(d.questions||[]);
  var total=questions.length;
  if(dir===-1&&sw_qIdx===0){
    if(sw_branchIdx>=0){
      if(!sw_branchAnswers[sw_domainIdx]) sw_branchAnswers[sw_domainIdx]={};
      sw_branchAnswers[sw_domainIdx][sw_branchIdx]=Object.assign({},sw_answers);
      sw_currentPage='branches';
    } else {
      sw_branchAnswers[sw_domainIdx]=sw_branchAnswers[sw_domainIdx]||{};
      sw_branchAnswers[sw_domainIdx][-1]=Object.assign({},sw_answers);
      sw_currentPage='domains';
    }
    stopStudentTimer();renderStudentWindowPage();return;
  }
  if(dir===1&&sw_qIdx===total-1){swSubmitDomain(total);return;}
  sw_qIdx=Math.max(0,Math.min(total-1,sw_qIdx+dir));
  renderStudentQuestion();updateProgressDots();
}
function swAutoAdvance(){
  // ٥ دقائق إضافية
  var extraSec=5*60;
  sw_timeLeft=extraSec;
  // أظهر Extra Time في العداد
  var timerEl=document.getElementById('sw-timer-display');
  if(timerEl){
    timerEl.style.color='#f97316';
    timerEl.parentElement&&(timerEl.parentElement.style.borderColor='rgba(249,115,22,.6)');
  }
  // أظهر badge
  var badge=document.getElementById('sw-extra-time-badge');
  if(!badge){
    badge=document.createElement('div');
    badge.id='sw-extra-time-badge';
    badge.style.cssText='position:fixed;top:80px;left:50%;transform:translateX(-50%);z-index:9999;background:linear-gradient(135deg,#f97316,#ea580c);color:white;font-family:Montserrat,sans-serif;font-weight:900;font-size:14px;padding:8px 24px;border-radius:30px;box-shadow:0 4px 20px rgba(249,115,22,.5);animation:timerPulse 1s infinite;letter-spacing:1px';
    badge.textContent='⏱ EXTRA TIME — 5 min';
    document.body.appendChild(badge);
  }
  updateTimerDisplay();
  sw_timerInterval=setInterval(function(){
    sw_timeLeft--;
    updateTimerDisplay();
    if(sw_timeLeft<=0){
      clearInterval(sw_timerInterval);
      var b=document.getElementById('sw-extra-time-badge');
      if(b) b.remove();
      // تسليم تلقائي
      _swAutoFinalSubmit();
    }
  },1000);
}

function _swAutoFinalSubmit(){
  // أكمل أي أسئلة غير مجابة بقيمة فارغة
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[]):(d.questions||[]);
  questions.forEach(function(_,i){if(sw_answers[i]===undefined) sw_answers[i]=null;});
  // إظهار رسالة
  scOk('انتهى الوقت ⏰','Time Up',
    'انتهى وقت الاختبار. سيتم تسليم إجاباتك تلقائياً.',
    'Time is up. Your answers will be submitted automatically.','⏰'
  ).then(function(){
    // تسليم بدون شرط 80%
    clearInterval(sw_timerInterval);
    if(sw_branchIdx>=0){
      if(!sw_branchAnswers[sw_domainIdx]) sw_branchAnswers[sw_domainIdx]={};
      sw_branchAnswers[sw_domainIdx][sw_branchIdx]=Object.assign({},sw_answers);
      if(!sw_completedBranches[sw_domainIdx]) sw_completedBranches[sw_domainIdx]=[];
      if(sw_completedBranches[sw_domainIdx].indexOf(sw_branchIdx)<0) sw_completedBranches[sw_domainIdx].push(sw_branchIdx);
      sw_branchIdx=-1; sw_answers={};
      sw_currentPage='branches'; renderStudentWindowPage();
    } else {
      if(sw_completedDomains.indexOf(sw_domainIdx)<0) sw_completedDomains.push(sw_domainIdx);
      if(currentStudentData){
        if(!currentStudentData.student.domainsCompleted) currentStudentData.student.domainsCompleted={};
        currentStudentData.student.domainsCompleted[sw_domainIdx]=true;
        saveMyTests();
      }
      sw_answers={};
      sw_currentPage='domains'; renderStudentWindowPage();
    }
  });
}
function swSubmitDomain(total){
  var d=testData.domains[sw_domainIdx];
  if(sw_branchIdx>=0){
    var br=d.branches[sw_branchIdx];
    if(!swCheck80Percent()) return;
    scDanger('تسليم الفرع','Submit Branch',
      'هل أنت متأكد من تسليم الفرع: <strong style="color:#FACC15">'+(br.nameAr||'الفرع')+'</strong>؟<br>لن تتمكن من العودة إليه.',
      'Submit branch: <strong>'+(br.nameEn||br.nameAr||'Branch')+'</strong>?<br>You cannot return to it.'
    ).then(function(ok){
      if(!ok)return;
      clearInterval(sw_timerInterval);
      if(!sw_branchAnswers[sw_domainIdx]) sw_branchAnswers[sw_domainIdx]={};
      sw_branchAnswers[sw_domainIdx][sw_branchIdx]=Object.assign({},sw_answers);
      if(!sw_completedBranches[sw_domainIdx]) sw_completedBranches[sw_domainIdx]=[];
      if(sw_completedBranches[sw_domainIdx].indexOf(sw_branchIdx)<0) sw_completedBranches[sw_domainIdx].push(sw_branchIdx);
      if(currentStudentData){
        if(!currentStudentData.student.domainsCompleted) currentStudentData.student.domainsCompleted={};
        currentStudentData.student.domainsCompleted[sw_domainIdx]=sw_completedBranches[sw_domainIdx];
        saveMyTests();
      }
      sw_branchIdx=-1; sw_answers={};
      var allBrDone=d.branches&&sw_completedBranches[sw_domainIdx]&&sw_completedBranches[sw_domainIdx].length>=d.branches.length;
      if(allBrDone){
        scOk('أحسنت! 🏆','Well Done!',
          'لقد أكملت جميع فروع مجال <strong style="color:#FACC15">'+d.nameAr+'</strong><br>سيتم نقلك لصفحة المجالات.',
          'You have completed all branches of <strong>'+(d.nameEn||d.nameAr)+'</strong><br>Redirecting to domains page.','🏆'
        ).then(function(){sw_currentPage='domains';renderStudentWindowPage();});
      } else {
        sw_currentPage='branches'; renderStudentWindowPage();
      }
    });
    return;
  }
  if(!swCheck80Percent()) return;
  var isLast=sw_domainIdx===testData.domains.length-1;
  scDanger(
    isLast?'التسليم النهائي':'تسليم المجال',
    isLast?'Final Submission':'Submit Domain',
    isLast
      ?'هذا هو <strong style="color:#f87171">المجال الأخير</strong> في الاختبار.<br>بمجرد التسليم سيتم إنهاء الاختبار بالكامل.'
      :'أنت على وشك تسليم: <strong style="color:#FACC15">'+d.nameAr+'</strong><br>لن تتمكن من العودة إليه بعد التسليم.',
    isLast
      ?'This is the <strong>LAST domain</strong>. The test will be finalized and you cannot return.'
      :'Submitting: <strong>'+(d.nameEn||d.nameAr)+'</strong><br>You cannot return once submitted.'
  ).then(function(ok){
    if(!ok)return;
    clearInterval(sw_timerInterval);
    if(sw_completedDomains.indexOf(sw_domainIdx)<0) sw_completedDomains.push(sw_domainIdx);
    if(currentStudentData){
      currentStudentData.student.started=true;
      if(!currentStudentData.student.domainsCompleted) currentStudentData.student.domainsCompleted={};
      currentStudentData.student.domainsCompleted[sw_domainIdx]=true;
      if(isLast) currentStudentData.student.completed=true;
      saveMyTests();
    }
    if(isLast){
      document.getElementById('sw-questions-content').innerHTML=
        '<div class="sw-frame" style="text-align:center;padding:60px 32px">'
        +'<div style="font-size:80px;margin-bottom:16px;animation:scPop .5s">🏆</div>'
        +'<h2 style="font-size:26px;font-weight:900;color:#1e3a8a;margin-bottom:8px">تم تسليم الاختبار بنجاح!</h2>'
        +'<p style="color:#16a34a;font-weight:700;font-size:16px;margin-bottom:6px;font-family:Montserrat,sans-serif">Test Submitted Successfully!</p>'
        +'<p style="color:#64748b;font-size:14px">سيتم إعلامك بالنتائج لاحقاً<br><span style="font-family:Montserrat,sans-serif">You will be notified of results later.</span></p>'
        +'</div>';
      setTimeout(function(){disableAntiCheat();document.getElementById('studentWindow').style.display='none';backToMain();},3500);
    } else {
      sw_answers={};
      scOk('تم تسليم المجال ✅','Domain Submitted',
        'تم تسليم مجال <strong style="color:#FACC15">'+d.nameAr+'</strong> بنجاح.<br>يمكنك الانتقال للمجال التالي.',
        'Domain <strong>'+(d.nameEn||d.nameAr)+'</strong> submitted.<br>You may proceed to the next domain.','✅'
      ).then(function(){sw_currentPage='domains';renderStudentWindowPage();});
    }
  });
}
function _oldSwSubmitDomain_removed(){
}

// ============================================================
// CANVAS
// ============================================================
function initCanvas(){var canvas=document.getElementById('drawingCanvas');if(!canvas)return;canvas.width=canvas.offsetWidth||560;canvasCtx=canvas.getContext('2d');canvas.addEventListener('mousedown',startDraw);canvas.addEventListener('mousemove',draw);canvas.addEventListener('mouseup',endDraw);canvas.addEventListener('mouseleave',endDraw);canvas.addEventListener('touchstart',function(e){e.preventDefault();startDraw(e.touches[0]);},{passive:false});canvas.addEventListener('touchmove',function(e){e.preventDefault();draw(e.touches[0]);},{passive:false});canvas.addEventListener('touchend',endDraw);}
function startDraw(e){isDrawingCanvas=true;var pos=getCanvasPos(e);canvasCtx.beginPath();canvasCtx.moveTo(pos.x,pos.y);}
function draw(e){if(!isDrawingCanvas||!canvasCtx)return;var pos=getCanvasPos(e);canvasCtx.lineWidth=brushSize;canvasCtx.lineCap='round';canvasCtx.strokeStyle=drawTool==='eraser'?'#ffffff':penColor;canvasCtx.lineTo(pos.x,pos.y);canvasCtx.stroke();}
function endDraw(){isDrawingCanvas=false;if(canvasCtx)canvasCtx.closePath();sw_answers[sw_qIdx]='drawn';updateProgressDots();}
function getCanvasPos(e){var canvas=document.getElementById('drawingCanvas');var rect=canvas.getBoundingClientRect();var zf=(testData.displayMode===4&&typeof _wpZoomLevel!=='undefined'&&_wpZoomLevel)?_wpZoomLevel:1;return{x:((e.clientX||e.pageX)-rect.left)/zf,y:((e.clientY||e.pageY)-rect.top)/zf};}
function setDrawTool(t){drawTool=t;var p=document.getElementById('sw-pen'),er=document.getElementById('sw-eraser');if(p)p.classList.toggle('active',t==='pen');if(er)er.classList.toggle('active',t==='eraser');}
function setBrushSize(v){brushSize=parseInt(v);}
function setPenColor(c){penColor=c;}
function clearCanvas(){var canvas=document.getElementById('drawingCanvas');if(canvas&&canvasCtx)canvasCtx.clearRect(0,0,canvas.width,canvas.height);}
function swToggleWriteMode(mode){var w=document.getElementById('sw-write-area'),dr=document.getElementById('sw-draw-area'),up=document.getElementById('sw-upload-area');if(w)w.style.display=mode==='write'?'block':'none';if(dr)dr.style.display=mode==='draw'?'block':'none';if(up)up.style.display=mode==='upload'?'block':'none';if(mode==='draw')setTimeout(initCanvas,50);}
function swUploadImage(event){var file=event.target.files?event.target.files[0]:null;if(!file)return;var reader=new FileReader();reader.onload=function(ev){var wrap=document.getElementById('sw-uploaded-img-wrap');if(wrap)wrap.innerHTML='<img src="'+ev.target.result+'" style="max-height:200px;border-radius:12px;margin:auto;display:block;border:2px solid #e2e8f0">';var up=document.getElementById('sw-upload-area');if(up)up.style.display='block';var w=document.getElementById('sw-write-area'),dr=document.getElementById('sw-draw-area');if(w)w.style.display='none';if(dr)dr.style.display='none';sw_answers[sw_qIdx]='image_uploaded';updateProgressDots();};reader.readAsDataURL(file);}

// ============================================================
// SPEAKING
// ============================================================
function swToggleRecord(){
  var btn=document.getElementById('sw-record-btn'),status=document.getElementById('sw-record-status');if(!btn)return;
  if(!btn.classList.contains('recording')){
    btn.classList.add('recording');btn.innerHTML='⏹ إيقاف / Stop Recording';
    var sec=0;recordingInterval=setInterval(function(){sec++;if(status)status.textContent='⏱ '+sec+'s recording...';},1000);
    if(navigator.mediaDevices&&navigator.mediaDevices.getUserMedia){navigator.mediaDevices.getUserMedia({audio:true}).then(function(stream){recordingChunks=[];mediaRecorder=new MediaRecorder(stream);mediaRecorder.ondataavailable=function(e){recordingChunks.push(e.data);};mediaRecorder.onstop=function(){var blob=new Blob(recordingChunks,{type:'audio/webm'});var url=URL.createObjectURL(blob);var audio=document.getElementById('sw-playback-audio');if(audio){audio.src=url;audio.style.display='block';}};mediaRecorder.start();}).catch(function(){if(status)status.textContent='⚠️ لا يمكن الوصول للميكروفون / Mic denied';});}
  } else {
    btn.classList.remove('recording');btn.innerHTML='🎙 ابدأ التسجيل / Start Recording';
    clearInterval(recordingInterval);if(mediaRecorder&&mediaRecorder.state!=='inactive')mediaRecorder.stop();
    if(status)status.textContent='✅ تم التسجيل / Done';
    var controls=document.getElementById('sw-record-controls');if(controls)controls.style.display='flex';
  }
}
function swListenPlayback(){var audio=document.getElementById('sw-playback-audio');if(audio&&audio.src){audio.style.display='block';audio.play();}else alert('سجّل أولاً / Record first');}
function swConfirmRecording(){sw_answers[sw_qIdx]='recorded';updateProgressDots();var c=document.getElementById('sw-record-controls');if(c)c.style.display='none';var s=document.getElementById('sw-record-status');if(s)s.textContent='✅ تم التأكيد / Confirmed';}
function swDeleteRecording(){var audio=document.getElementById('sw-playback-audio');if(audio){audio.src='';audio.style.display='none';}var c=document.getElementById('sw-record-controls');if(c)c.style.display='none';var btn=document.getElementById('sw-record-btn');if(btn)btn.innerHTML='🎙 ابدأ التسجيل / Start Recording';delete sw_answers[sw_qIdx];updateProgressDots();}

// ============================================================
// ANTI-CHEAT
// ============================================================
function enableAntiCheat(){document.addEventListener('copy',handleCopy);document.addEventListener('keydown',handleKeyDown);document.addEventListener('visibilitychange',handleVisibility);window.addEventListener('blur',handleBlur);}
function disableAntiCheat(){document.removeEventListener('copy',handleCopy);document.removeEventListener('keydown',handleKeyDown);document.removeEventListener('visibilitychange',handleVisibility);window.removeEventListener('blur',handleBlur);}
function handleCopy(e){if(document.getElementById('studentWindow')&&document.getElementById('studentWindow').style.display==='flex'){e.preventDefault();triggerCheatWarning('محاولة نسخ / Copy');}}
function handleKeyDown(e){if(!document.getElementById('studentWindow')||document.getElementById('studentWindow').style.display!=='flex')return;if(e.key==='F12'||(e.ctrlKey&&e.shiftKey&&'IJ'.includes(e.key))||(e.ctrlKey&&'uU'.includes(e.key))){e.preventDefault();triggerCheatWarning('DevTools');}if(e.ctrlKey&&e.key==='p'){e.preventDefault();triggerCheatWarning('Print');}}
function handleVisibility(){var btn=document.getElementById('sw-record-btn');if(btn&&btn.classList.contains('recording'))return;if(document.getElementById('studentWindow')&&document.getElementById('studentWindow').style.display==='flex'&&document.hidden)triggerCheatWarning('Tab switch');}
function handleBlur(){var btn=document.getElementById('sw-record-btn');if(btn&&btn.classList.contains('recording'))return;if(document.getElementById('studentWindow')&&document.getElementById('studentWindow').style.display==='flex')triggerCheatWarning('Window blur');}
function triggerCheatWarning(reason){cheatWarnings++;cheatLog.push({time:new Date().toLocaleTimeString('ar'),reason:reason,count:cheatWarnings});document.getElementById('cheatCountDisplay').textContent=cheatWarnings;document.getElementById('cheatMsg').innerHTML='تم اكتشاف: <strong>'+reason+'</strong><br><span style="font-size:12px;opacity:.8">تم التسجيل / Recorded</span>';document.getElementById('cheatWarning').style.display='flex';}
function dismissCheatWarning(){document.getElementById('cheatWarning').style.display='none';}

// ============================================================
// REVIEWER PANEL
// ============================================================
function openReviewerPanel(){document.getElementById('supMainOptions').classList.add('hidden');document.getElementById('supTitle').classList.add('hidden');document.getElementById('supReviewerPanel').classList.remove('hidden');renderSupReviewerContent();}
function closeReviewerPanel(){document.getElementById('supReviewerPanel').classList.add('hidden');document.getElementById('supMainOptions').classList.remove('hidden');document.getElementById('supTitle').classList.remove('hidden');}
function renderSupReviewerContent(){
  var assigned=myTests.filter(function(t){return t.reviewer1===currentSupName&&t.status==='underReview';});
  var cont=document.getElementById('supReviewerContent');
  if(!assigned.length){cont.innerHTML='<div style="text-align:center;color:rgba(255,255,255,.3);padding:48px"><div style="font-size:48px;margin-bottom:12px">🔍</div><p>لا توجد اختبارات لمراجعتها / No tests assigned</p></div>';return;}
  cont.innerHTML=assigned.map(function(t){
    var notesHtml='';
    if(t.returnNotes&&t.returnNotes.length){notesHtml='<div style="margin-top:10px">'+t.returnNotes.map(function(n){return '<div style="background:rgba(248,113,113,.1);border:1px solid rgba(248,113,113,.3);border-radius:10px;padding:8px 12px;margin-bottom:5px;font-size:12px"><div style="color:#fca5a5;font-weight:700;margin-bottom:3px">'+n.name+' — <span style="font-size:10px;color:rgba(255,255,255,.4)">'+n.at+'</span></div><div style="color:#fecaca">'+n.note+'</div></div>';}).join('')+'</div>';}
    return '<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:18px;padding:20px;margin-bottom:16px"><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:14px"><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المدرسة</div><div style="font-weight:700;font-size:13px">'+(t.school||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المادة</div><div style="font-weight:700;font-size:13px">'+(t.subject||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المحرر</div><div style="font-weight:700;font-size:13px">'+(t.author||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">الاسم</div><div style="font-weight:700;font-size:13px">'+(t.name||'—')+'</div></div><div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">المراجع العام</div><div><span class="status-badge '+(t.generalStatus==='approved'?'status-approved':t.generalStatus==='returned'?'status-returned':'status-pending')+'" style="font-size:11px">'+(t.generalStatus==='approved'?'✅ اعتمد':t.generalStatus==='returned'?'⚠️ أرجع':'⏳ انتظار')+'</span></div></div></div>'+notesHtml+'<div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:12px"><button onclick="srPreviewTest('+t.id+')" style="background:rgba(96,165,250,.2);color:#93c5fd;border:none;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">👁 معاينة</button><button onclick="srMarkReviewed('+t.id+')" style="background:rgba(74,222,128,.2);color:#4ade80;border:none;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">✅ تم الفحص</button><button onclick="srReturnTest('+t.id+')" style="background:rgba(248,113,113,.2);color:#f87171;border:none;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">⚠️ إرجاع + ملاحظة</button></div></div>';
  }).join('');
}
function srPreviewTest(id){var t=myTests.find(function(x){return x.id===id;});if(!t||!t.domains||!t.domains.length)return;testData.domains=t.domains;testData.logoSrc=t.logoSrc||'';testData.instructionsAr=t.instructionsAr||'';testData.instructionsEn=t.instructionsEn||'';testData.testName=t.testName||t.name||testData.testName;testData.subject=t.subject||testData.subject;testData.grade=t.grade||testData.grade;testData.term=t.term||testData.term;testData.year=t.year||testData.year;openStudentWindow(0);}
function srMarkReviewed(id){var t=myTests.find(function(x){return x.id===id;});if(!t)return;t.r1done=true;if(!t.returnNotes)t.returnNotes=[];saveMyTests();renderSupReviewerContent();alert('✅ تم تسجيل مراجعتك!');}
function srReturnTest(id){
  var t=myTests.find(function(x){return x.id===id;});
  if(!t)return;
  scPromptText('ملاحظات الإرجاع','Return Notes','اكتب سبب إرجاع الاختبار / Write the reason','⚠️').then(function(r){
    if(!r)return;
    t.r1done=false;
    if(!t.returnNotes)t.returnNotes=[];
    t.returnNotes.push({from:'reviewer1',name:t.reviewer1||'المراجع',note:r,at:new Date().toLocaleString('ar')});
    saveMyTests();renderSupReviewerContent();
    scOk('تم الإرجاع ⚠️','Returned','تم إرجاع الاختبار مع الملاحظة','Test returned with notes','⚠️');
  });
}

// ============================================================
// MY TESTS
// ============================================================
function saveMyTests(){return ghWrite('tests.json',myTests);}
function openMyTests(){renderMyTestsTabs('underReview');document.getElementById('myTestsModal').classList.remove('hidden');}
function closeMyTests(){document.getElementById('myTestsModal').classList.add('hidden');}
function renderMyTestsTabs(tab){
  ['underReview','reviewDone','approved'].forEach(function(t){var btn=document.getElementById('tab_'+t);if(btn)btn.className='tab-btn'+(t===tab?' active':'');});
  document.getElementById('myTestsContent').innerHTML=renderMyTestsContent(tab);
  document.getElementById('studentCodesPanel').classList.add('hidden');
}
function renderMyTestsContent(tab){
  var curName=(currentSupName||'').trim();
  var filtered=myTests.filter(function(t){
    var tAuthor=(t.author||'').trim();
    var match=tAuthor===curName||(!curName&&!tAuthor);
    if(tab==='underReview') return t.status==='underReview'&&match;
    if(tab==='reviewDone')  return t.status==='reviewDone'&&match;
    if(tab==='approved')    return t.status==='approved'&&match;
  });
  if(!filtered.length) return '<div style="text-align:center;color:rgba(255,255,255,.3);padding:48px"><div style="font-size:48px;margin-bottom:12px">'+(tab==='underReview'?'🔄':tab==='reviewDone'?'✅':'🏆')+'</div><p>لا توجد اختبارات / No tests yet</p></div>';
  return filtered.map(function(t){
    var notesHtml='';
    if(t.returnNotes&&t.returnNotes.length){
      notesHtml='<div style="margin-top:10px"><div style="font-size:11px;color:rgba(255,255,255,.4);margin-bottom:6px">ملاحظات المراجعين:</div>'+t.returnNotes.map(function(n){return '<div style="background:rgba(248,113,113,.1);border:1px solid rgba(248,113,113,.3);border-radius:10px;padding:8px 12px;margin-bottom:6px;font-size:12px"><div style="color:#fca5a5;font-weight:700;margin-bottom:3px">'+n.name+' — <span style="font-size:10px;color:rgba(255,255,255,.4)">'+n.at+'</span></div><div style="color:#fecaca">'+n.note+'</div></div>';}).join('')+'</div>';
    }
    return '<div style="background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:18px;padding:20px;margin-bottom:16px"><div style="display:flex;align-items:flex-start;justify-content:space-between;gap:16px;flex-wrap:wrap"><div style="flex:1"><div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap"><div class="domain-badge" style="font-size:11px">'+(t.subject||'—')+'</div><h4 style="font-weight:800;font-size:16px">'+(t.name||'اختبار')+'</h4><span style="font-size:11px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">'+(t.grade||'')+' | Term '+(t.term||'')+' | '+(t.year||'')+'</span></div><div style="font-size:12px;color:rgba(255,255,255,.5);margin-bottom:10px">'+(t.school||'')+'</div><div style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px"><div style="background:rgba(255,255,255,.05);border-radius:10px;padding:6px 12px;font-size:12px">✍️ '+(t.author||'—')+' <span style="color:#4ade80">✓</span></div><div style="background:rgba(255,255,255,.05);border-radius:10px;padding:6px 12px;font-size:12px">🔍 '+(t.reviewer1||'—')+' <span class="status-badge '+(t.r1done?'status-reviewed':'status-pending')+'" style="font-size:10px">'+(t.r1done?'✅ راجع':'⏳ انتظار')+'</span></div><div style="background:rgba(255,255,255,.05);border-radius:10px;padding:6px 12px;font-size:12px">🏛️ المراجع العام <span class="status-badge '+(t.generalStatus==='approved'?'status-approved':t.generalStatus==='returned'?'status-returned':'status-pending')+'" style="font-size:10px">'+(t.generalStatus==='approved'?'✅ اعتمد':t.generalStatus==='returned'?'⚠️ أرجع':'⏳ انتظار')+'</span></div></div>'+notesHtml+'</div><div style="display:flex;flex-direction:column;gap:8px;min-width:140px"><button onclick="previewMyTest('+t.id+')" style="background:rgba(96,165,250,.2);border:none;color:#93c5fd;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">👁 معاينة</button>'+(tab==='underReview'?'<button onclick="editMyTest('+t.id+')" style="background:rgba(250,204,21,.15);border:none;color:#FACC15;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">✏️ تعديل</button>':'')+(tab==='reviewDone'&&t.generalStatus==='approved'?'<button onclick="finalApproveTest('+t.id+')" style="background:rgba(74,222,128,.2);border:none;color:#4ade80;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">✅ اعتماد نهائي</button>':'')+(tab==='approved'?'<button onclick="openStudentCodesPanel('+t.id+')" style="background:rgba(167,139,250,.2);border:none;color:#c4b5fd;border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">👥 أكواد الطلاب</button>':'')+'</div></div></div>';
  }).join('');
}
function previewMyTest(id){var t=myTests.find(function(x){return x.id===id;});if(!t||!t.domains||!t.domains.length)return;testData.domains=t.domains;testData.logoSrc=t.logoSrc||'';testData.instructionsAr=t.instructionsAr||'';testData.instructionsEn=t.instructionsEn||'';testData.testName=t.testName||t.name||testData.testName;testData.subject=t.subject||testData.subject;testData.grade=t.grade||testData.grade;testData.term=t.term||testData.term;testData.year=t.year||testData.year;openStudentWindow(0);}
function editMyTest(id){var t=myTests.find(function(x){return x.id===id;});if(!t){alert('لم يتم العثور على الاختبار');return;}editingTestId=t.id;startTestWizard();document.getElementById('testName').value=t.name||'';document.getElementById('grade').value=t.grade||'';document.getElementById('academicYear').value=t.year||'';document.getElementById('term').value=t.term||'';var subjectEl=document.getElementById('subject');if(subjectEl){for(var i=0;i<subjectEl.options.length;i++){if(subjectEl.options[i].text===t.subject||subjectEl.options[i].value===t.subject){subjectEl.selectedIndex=i;break;}}}
  var rev1El=document.getElementById('reviewer1');if(rev1El){for(var j=0;j<rev1El.options.length;j++){var text=rev1El.options[j].text||'';if(text.indexOf(t.reviewer1)===0||rev1El.options[j].value===t.reviewer1){rev1El.selectedIndex=j;break;}}}
  document.getElementById('reviewer0').value=currentSupName||t.author||'';
  document.getElementById('insAr').innerHTML=t.instructionsAr||'';
  document.getElementById('insEn').innerHTML=t.instructionsEn||'';
  document.getElementById('domQty').value=(t.domains? t.domains.length:0)||0;
  testData.domains=JSON.parse(JSON.stringify(t.domains||[]));
  testData.logoSrc=t.logoSrc||'';
  testData.displayMode=t.displayMode||1;
  setTimeout(function(){selectDisplayMode(testData.displayMode);},100);
  if(t.school){selectedSchools=schools.filter(function(s){return t.school.split(',').map(function(name){return name.trim();}).includes(s.name);}).map(function(s){return s.id;});}
  generateTestLink();
  previewReviewerNotice();
  generateDomains();
  alert('🔧 تم فتح الاختبار للتحرير. عدّل البيانات ثم أرسلها للمراجعة.');}
function finalApproveTest(id){
  var t=myTests.find(function(x){return x.id===id;});
  if(!t)return;
  if(t.generalStatus!=='approved'){
    scWarn('يجب موافقة المراجع العام قبل الاعتماد النهائي','General reviewer approval required first');
    return;
  }
  scConfirm('اعتماد نهائي','Final Approval','هل تريد الاعتماد النهائي لهذا الاختبار؟<br><b>'+(t.testName||t.name||'')+'</b><br>سيتم تفعيله للطلاب وتخليق أكواد الدخول','Finalize approval and generate student codes?','🏆').then(function(ok){
    if(!ok)return;
    t.status='approved';
    t.approvedAt=new Date().toLocaleDateString('ar');
    saveMyTests();
    renderMyTestsTabs('approved');
    scOk('تم الاعتماد النهائي 🏆','Finalized','تم اعتماد الاختبار نهائياً — يمكنك الآن إنشاء أكواد الطلاب','Test finalized — you can now generate student codes','🏆');
  });
}

// ============================================================
// APPROVE & SUBMIT TEST
// ============================================================
function approveFinalTest(){
  var testName=document.getElementById('testName')?document.getElementById('testName').value||'اختبار':'اختبار';
  var grade=document.getElementById('grade')?document.getElementById('grade').value||'':'';
  var subjectEl=document.getElementById('subject');
  var subjectVal=subjectEl&&subjectEl.selectedIndex>=0?subjectEl.options[subjectEl.selectedIndex].text:'';
  var term=document.getElementById('term')?document.getElementById('term').value||'':'';
  var year=document.getElementById('academicYear')?document.getElementById('academicYear').value||'':'';
  var author=document.getElementById('reviewer0')?document.getElementById('reviewer0').value||'':'';
  var rev1El=document.getElementById('reviewer1');
  var rev1=rev1El&&rev1El.selectedIndex>=0?rev1El.options[rev1El.selectedIndex].text:'';
  var instructionsAr=document.getElementById('insAr')?cleanHTML(document.getElementById('insAr').innerHTML.trim()):'',
      instructionsEn=document.getElementById('insEn')?document.getElementById('insEn').innerHTML.trim():'';
  var selectedSchoolNames=selectedSchools.map(function(id){var s=schools.find(function(x){return x.id===id;});return s?s.name:'';}).filter(Boolean).join(', ');
  if(editingTestId!==null){
    var t=myTests.find(function(x){return x.id===editingTestId;});
    if(!t){alert('خطأ: لم يتم العثور على الاختبار للتحرير.');editingTestId=null;return;}
    t.name=testName;
    t.grade=grade;
    t.subject=subjectVal;
    t.term=term;
    t.year=year;
    t.author=author;
    t.reviewer1=rev1;
    t.school=selectedSchoolNames||t.school||'';
    t.logoSrc=testData.logoSrc;
    t.domains=JSON.parse(JSON.stringify(testData.domains));
    t.instructionsAr=instructionsAr;
    t.instructionsEn=instructionsEn;
    t.displayMode=testData.displayMode||1;
    t.status='underReview';
    t.r1done=false;
    t.generalStatus='pending';
    if(!t.returnNotes) t.returnNotes=[];
    saveMyTests();
    editingTestId=null;
    alert('✅ تم تحديث الاختبار وإرساله للمراجعة مرة أخرى.');
  } else {
    var schoolObj=selectedSchools.map(function(id){return schools.find(function(x){return x.id===id;});}).filter(Boolean)[0]||null;
    var newTest={id:Date.now(),name:testName,grade:grade,subject:subjectVal,term:term,year:year,author:author,reviewer1:rev1,status:'underReview',r1done:false,generalStatus:'pending',returnNotes:[],school:selectedSchoolNames,schoolCode:schoolObj?schoolObj.code||'':'',country:schoolObj?schoolObj.country||'':'',curriculum:testData.curriculum||'',logoSrc:testData.logoSrc,domains:JSON.parse(JSON.stringify(testData.domains)),instructionsAr:instructionsAr,instructionsEn:instructionsEn,displayMode:testData.displayMode||1,students:[],createdAt:new Date().toLocaleDateString('ar'),testName:testName};
    myTests.push(newTest);saveMyTests();
    // أضف للأرشيف
    if(!archivedTests) archivedTests=[];
    archivedTests.unshift(Object.assign({},newTest,{archivedAt:Date.now(),id:'arch-'+newTest.id}));
    ghWrite('archive.json',archivedTests).catch(function(){});
    scOk('تم الإرسال ✅','Submitted','تم إرسال الاختبار للمراجعة بنجاح<br><b>'+testName+'</b>','Test submitted successfully','✅');
  }
  testData={domains:[],selectedSchools:[],logoSrc:'',displayMode:1};selectedSchools=[];_saveDraft();
  cancelWizard();
}

// ============================================================
// STUDENT CODES
// ============================================================
function generateStudentCode(firstName,lastName){
  var fn=(firstName||'Student').trim().replace(/\s+/g,'').replace(/[^a-zA-Z\u0600-\u06FF]/g,'');
  var digits=String(Math.floor(1000+Math.random()*9000));
  var username=fn+'@'+digits;
  var password=digits+'@ist';
  return{username:username,password:password};
}
function openStudentCodesPanel(testId){activeCodesTestId=testId;var t=myTests.find(function(x){return x.id===testId;});if(!t)return;if(!t.students)t.students=[];document.getElementById('codesTestId').value=testId;pendingStudentCodes=t.students.slice();document.getElementById('studentCodesPanel').classList.remove('hidden');document.getElementById('uploadCodesResult').textContent='';renderStudentCodesTable();}
function addStudentManually(){
  var firstName=document.getElementById('manFirstName').value.trim(),lastName=document.getElementById('manLastName').value.trim();
  if(!firstName&&!lastName){alert('يرجى إدخال الاسم');return;}
  var creds=generateStudentCode(firstName,lastName);
  pendingStudentCodes.push({schoolID:document.getElementById('manSchoolID').value.trim(),firstName:firstName,lastName:lastName,gender:document.getElementById('manGender').value,schoolName:document.getElementById('manSchoolName').value.trim(),nationality:document.getElementById('manNationality').value.trim()||'—',national:document.getElementById('manNational').value,gifted:document.getElementById('manGifted').value,sod:document.getElementById('manSOD').value,username:creds.username,password:creds.password,active:true});
  ['manSchoolID','manFirstName','manLastName','manSchoolName','manNationality'].forEach(function(id){var el=document.getElementById(id);if(el)el.value='';});
  saveStudentCodes();renderStudentCodesTable();
}
function saveStudentCodes(){var t=myTests.find(function(x){return x.id===activeCodesTestId;});if(!t)return;t.students=pendingStudentCodes;saveMyTests();}
function uploadStudentsExcelCodes(event){
  var file=event.target.files?event.target.files[0]:null;if(!file)return;
  var status=document.getElementById('uploadCodesResult');status.textContent='⏳ جاري القراءة...';
  var reader=new FileReader();
  // دالة معالجة الصفوف — تدعم الفورمات القديم والجديد
  function processRows(rows){
    var added=0;
    var isNewFormat=false;
    // كشف الفورمات: إذا الـ header الأول يحتوي "First Name" في col0 → فورمات جديد
    if(rows.length>0){
      var h=rows[0];
      var h0=String(h[0]||'').toLowerCase();
      var h1=String(h[1]||'').toLowerCase();
      // الفورمات الجديد: col0=FirstName, col1=LastName, col2=School...
      // الفورمات القديم: col0=SchoolID, col1=FirstName, col2=LastName...
      if(h0.indexOf('first')>=0||h0.indexOf('اسم')>=0) isNewFormat=true;
      else if(h1.indexOf('first')>=0||h1.indexOf('اسم')>=0) isNewFormat=false;
      else {
        // كشف تلقائي: إذا col0 في row1 يبدو رقم مدرسي → قديم, وإلا جديد
        var testRow=rows[1];
        if(testRow){ var v0=String(testRow[0]||''); isNewFormat=!(v0.match(/^\d+$/)||v0.toLowerCase().startsWith('stu')); }
      }
    }
    rows.forEach(function(row,idx){
      if(idx===0)return; // تخطي الهيدر
      var fn,ln,school,grade,section,nationality,gender,sen,gifted,local;
      if(isNewFormat){
        // الفورمات الجديد من الملف المرفوع
        fn=String(row[0]||'').trim();
        ln=String(row[1]||'').trim();
        school=String(row[2]||'').trim();
        grade=String(row[3]||'').trim();
        section=String(row[4]||'').trim();
        nationality=String(row[5]||'').trim();
        gender=String(row[6]||'').trim();
        sen=String(row[7]||'-').trim();
        gifted=String(row[8]||'-').trim();
        local=String(row[9]||'No').trim();
      } else {
        // الفورمات القديم: col0=SchoolID, col1=FirstName, col2=LastName...
        fn=String(row[1]||'').trim();
        ln=String(row[2]||'').trim();
        school=String(row[4]||'').trim();
        grade=String(row[5]||'').trim();
        section='';
        nationality=String(row[12]||'').trim();
        gender=String(row[3]||'Male').trim();
        sen='No';
        gifted=String(row[9]||'No').trim();
        local=String(row[8]||'No').trim();
      }
      if(!fn&&!ln)return;
      // توحيد الجنس
      var genderNorm=gender.toLowerCase();
      if(genderNorm==='girl'||genderNorm==='female'||genderNorm==='أنثى'||genderNorm==='f') gender='Female';
      else gender='Male';
      // SEN/G&T → gifted/sod
      var isSen=(sen&&sen!=='-'&&sen.toLowerCase()!=='no'&&sen.toLowerCase()!=='false')?'Yes':'No';
      var isGifted=(gifted&&gifted!=='-'&&gifted.toLowerCase()!=='no'&&gifted.toLowerCase()!=='false'&&gifted.toLowerCase().indexOf('g')===0)?'Yes':'No';
      var isLocal=(local&&local.toLowerCase()==='yes'||local==='Y')?'Yes':'No';
      var cr=generateStudentCode(fn,ln);
      pendingStudentCodes.push({
        schoolID:String(row[0]||'').trim(),
        firstName:fn,lastName:ln,
        gender:gender,
        schoolName:school,
        grade:grade,
        section:section,
        nationality:nationality,
        national:isLocal,
        gifted:isGifted,
        sod:isSen,
        username:cr.username,
        password:cr.password,
        active:true
      });
      added++;
    });
    saveStudentCodes();renderStudentCodesTable();
    status.textContent='✅ '+added+' طالب / students uploaded';
  }
  if(file.name.match(/\.csv$/i)){
    reader.onload=function(e){
      try{
        var lines=e.target.result.split('\n').filter(function(l){return l.trim();});
        var rows=lines.map(function(l){return l.split(',').map(function(c){return c.replace(/"/g,'').trim();});});
        processRows(rows);
      }catch(err){status.textContent='❌ '+err.message;}
    };
    reader.readAsText(file,'UTF-8');
  } else {
    reader.onload=function(e){
      try{
        var wb=XLSX.read(new Uint8Array(e.target.result),{type:'array'});
        var rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:''});
        processRows(rows);
      }catch(err){status.textContent='❌ Excel: '+err.message;}
    };
    reader.readAsArrayBuffer(file);
  }
  event.target.value='';
}
function downloadStudentTemplate(e){e.preventDefault();var ws=XLSX.utils.aoa_to_sheet([['Student School ID / الكود المدرسي','First Name / الاسم الأول','Last Name / الاسم الثاني','Gender / الجنس (Male/Female)','School Name / اسم المدرسة','Grade / الصف','Subject / المادة','Term / الفصل','Local/National (Y/N) / مواطن','Gifted (Y/N) / موهوب','SOD (Y/N)'],['STU001','Ahmed','Ali','Male','IST School','Grade 5/Year 6','Math','1','Yes','No','No'],['STU002','Sara','Hassan','Female','IST School','Grade 5/Year 6','Math','1','No','Yes','No']]);ws['!cols']=[{wch:22},{wch:16},{wch:16},{wch:16},{wch:20},{wch:16},{wch:16},{wch:10},{wch:14},{wch:12},{wch:10}];var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Students');XLSX.writeFile(wb,'Students_Template.xlsx');}

function deleteAllStudentCodes(){
  if(!pendingStudentCodes.length)return;
  scConfirm('حذف الجميع','Delete All','⚠️ حذف جميع الطلاب ('+pendingStudentCodes.length+')؟','Delete all '+pendingStudentCodes.length+' students?','🗑').then(function(ok){
    if(!ok)return;
    pendingStudentCodes=[];
    saveStudentCodes();renderStudentCodesTable();
  });
}
function toggleStudentCode(idx){pendingStudentCodes[idx].active=!(pendingStudentCodes[idx].active!==false);saveStudentCodes();renderStudentCodesTable();}
function removeStudentCode(idx){scConfirm('حذف الطالب','Delete Student','هل تريد حذف هذا الطالب؟','Delete this student?','🗑').then(function(ok){if(ok){pendingStudentCodes.splice(idx,1);saveStudentCodes();renderStudentCodesTable();}});}
function renderStudentCodesTable(){
  var wrap=document.getElementById('studentsCodesTableWrap'),exportBtns=document.getElementById('exportStudentBtns');
  if(!pendingStudentCodes.length){
    wrap.innerHTML='<p style="color:rgba(255,255,255,.3);font-size:13px;text-align:center;padding:24px">لا يوجد طلاب / No students yet</p>';
    exportBtns.style.display='none';return;
  }
  exportBtns.style.display='flex';
  var delAllBtn='<div style="margin-bottom:10px;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:8px">'
    +'<span style="font-size:12px;color:rgba(255,255,255,.5);font-family:Montserrat,sans-serif">'+pendingStudentCodes.length+' طالب / students</span>'
    +'<button onclick="deleteAllStudentCodes()" style="background:rgba(239,68,68,.2);border:1.5px solid rgba(239,68,68,.4);color:#f87171;border-radius:10px;padding:7px 16px;font-size:12px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🗑 حذف الجميع / Delete All</button>'
    +'</div>';
  wrap.innerHTML=delAllBtn
    +'<div style="overflow-x:auto"><table class="admin-table" style="font-size:11px">'
    +'<thead><tr>'
    +'<th>#</th><th>الاسم الأول</th><th>الاسم الثاني</th><th>الجنس</th>'
    +'<th>المدرسة</th><th>الصف</th><th>الجنسية</th>'
    +'<th>مواطن</th><th>موهوب</th><th>SEN</th>'
    +'<th style="color:#FACC15">Username</th><th style="color:#4ade80">Password</th>'
    +'<th>الحالة</th><th>إجراء</th>'
    +'</tr></thead><tbody>'
    +pendingStudentCodes.map(function(s,i){
      return '<tr>'
        +'<td class="font-en" style="color:rgba(255,255,255,.4)">'+(i+1)+'</td>'
        +'<td style="font-weight:700">'+s.firstName+'</td>'
        +'<td>'+s.lastName+'</td>'
        +'<td>'+(s.gender==='Female'?'👧':'👦')+'</td>'
        +'<td style="font-size:10px;max-width:90px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+(s.schoolName||'')+'">'+(s.schoolName||'—')+'</td>'
        +'<td style="font-size:10px">'+(s.grade||'—')+(s.section?' / '+s.section:'')+'</td>'
        +'<td style="font-size:10px">'+(s.nationality||'—')+'</td>'
        +'<td>'+(s.national==='Yes'?'✅':'—')+'</td>'
        +'<td>'+(s.gifted==='Yes'?'⭐':'—')+'</td>'
        +'<td>'+(s.sod==='Yes'?'♿':'—')+'</td>'
        +'<td class="font-en" style="color:#FACC15;font-size:10px;font-weight:700">'+s.username+'</td>'
        +'<td class="font-en" style="color:#4ade80;font-size:10px;font-weight:700">'+s.password+'</td>'
        +'<td><span style="color:'+(s.active!==false?'#4ade80':'#f87171')+';font-weight:700">'+(s.active!==false?'✅':'⛔')+'</span></td>'
        +'<td style="white-space:nowrap">'
        +'<button onclick="toggleStudentCode('+i+')" style="background:none;border:none;cursor:pointer;font-size:13px">'+(s.active!==false?'⛔':'✅')+'</button>'
        +'<button onclick="removeStudentCode('+i+')" style="background:none;border:none;cursor:pointer;font-size:13px;color:#f87171">🗑</button>'
        +'</td></tr>';
    }).join('')
    +'</tbody></table></div>';
}
function exportStudentCodesExcel(){var t=myTests.find(function(x){return x.id===activeCodesTestId;});var testName=(t?t.name||'Test':'Test').replace(/\s+/g,'_');var data=[['Student School ID','First Name','Last Name','Gender','Local','Gifted','SOD','Username','Password','Status']];pendingStudentCodes.forEach(function(s){data.push([s.schoolID||'',s.firstName,s.lastName,s.gender,s.national,s.gifted,s.sod,s.username,s.password,s.active!==false?'Active':'Disabled']);});var ws=XLSX.utils.aoa_to_sheet(data);var wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Codes');XLSX.writeFile(wb,testName+'_StudentCodes.xlsx');}

// ══ بطاقات دخول الطلاب — قابلة للطباعة PDF والمشاركة ══
function showStudentIDCards(){
  var t=myTests.find(function(x){return x.id===activeCodesTestId;});
  if(!t||!pendingStudentCodes||!pendingStudentCodes.length){scWarn('لا يوجد طلاب لعرضهم بعد — أضف طلاباً أولاً','No students to display yet — add students first');return;}
  var sch=schools.find(function(s){return s.name===t.school||(t.school||'').indexOf(s.name)>=0;});
  var cardsHtml=pendingStudentCodes.map(function(st){return buildOneStudentCardHtml(st,t,sch);}).join('');
  var old=document.getElementById('studentCardsOverlay');if(old)old.remove();
  var overlay=document.createElement('div');
  overlay.id='studentCardsOverlay';
  overlay.style.cssText='position:fixed;inset:0;z-index:99999;background:#eef2f7;overflow-y:auto;font-family:Tajawal,Arial,sans-serif';
  overlay.innerHTML=
    '<div style="position:sticky;top:0;z-index:10;background:linear-gradient(135deg,#0f172a,#1e3a8a);padding:14px 24px;display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;box-shadow:0 2px 10px rgba(0,0,0,.25)">'
    +'<div style="color:white;font-weight:900;font-size:15px">🎫 بطاقات دخول الطلاب / Student Login Cards — <span style="color:#FACC15">'+(t.name||'')+'</span></div>'
    +'<div style="display:flex;gap:10px;flex-wrap:wrap">'
    +'<button onclick="printStudentCards()" style="background:#FACC15;color:#1e3a8a;border:none;border-radius:10px;padding:9px 20px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px">🖨️ طباعة / حفظ PDF</button>'
    +'<button onclick="shareStudentCards()" style="background:rgba(255,255,255,.15);color:white;border:1px solid rgba(255,255,255,.3);border-radius:10px;padding:9px 20px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px">📤 مشاركة / Share</button>'
    +'<button onclick="document.getElementById(\'studentCardsOverlay\').remove()" style="background:rgba(239,68,68,.2);color:#fca5a5;border:1px solid rgba(239,68,68,.4);border-radius:10px;padding:9px 16px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px">✕ إغلاق</button>'
    +'</div></div>'
    +'<div id="studentCardsPrintArea" style="padding:24px;display:grid;grid-template-columns:repeat(auto-fill,minmax(360px,1fr));gap:20px;max-width:1140px;margin:0 auto">'
    +cardsHtml
    +'</div>';
  document.body.appendChild(overlay);
}
function buildOneStudentCardHtml(st,t,sch){
  var qrText='Scholastic Student Login — Username: '+st.username+' | Password: '+st.password;
  var qrUrl='https://api.qrserver.com/v1/create-qr-code/?size=140x140&data='+encodeURIComponent(qrText);
  var logoHtml=(t&&t.logoSrc)?'<img src="'+t.logoSrc+'" style="width:38px;height:38px;object-fit:contain;border-radius:8px;background:white;flex-shrink:0">':(sch&&sch.logoSrc?'<img src="'+sch.logoSrc+'" style="width:38px;height:38px;object-fit:contain;border-radius:8px;background:white;flex-shrink:0">':'<div style="width:38px;height:38px;border-radius:8px;background:linear-gradient(135deg,#f59e0b,#d97706);display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0">🎓</div>');
  var schoolName=(t&&t.school)||(sch&&sch.name)||'Scholastic';
  var isLocal=st.national==='Yes';
  var badge=isLocal
    ?'<span style="background:#3b82f6;color:white;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:800;font-family:Montserrat,sans-serif">مواطن / Local</span>'
    :'<span style="background:#f59e0b;color:white;border-radius:6px;padding:3px 10px;font-size:11px;font-weight:800;font-family:Montserrat,sans-serif">غير مواطن / Non-Local</span>';
  var extraBadges='';
  if(st.gifted==='Yes') extraBadges+='<span style="background:#a855f7;color:white;border-radius:6px;padding:3px 8px;font-size:10px;font-weight:800;font-family:Montserrat,sans-serif;margin-right:4px">⭐ Gifted</span>';
  if(st.sod==='Yes') extraBadges+='<span style="background:#ec4899;color:white;border-radius:6px;padding:3px 8px;font-size:10px;font-weight:800;font-family:Montserrat,sans-serif">SOD</span>';
  return '<div class="id-card-outer" style="background:linear-gradient(135deg,#5eead4,#2dd4bf);border-radius:20px;padding:10px;page-break-inside:avoid">'
    +'<div style="background:white;border-radius:16px;padding:18px">'
    +'<div style="display:flex;align-items:center;gap:10px;padding-bottom:10px;border-bottom:1px solid #e2e8f0;margin-bottom:12px">'
    +logoHtml
    +'<div><div style="font-weight:900;font-size:13px;color:#1e293b">'+schoolName+'</div><div style="font-weight:700;font-size:11px;color:#64748b;font-family:Montserrat,sans-serif">Student Login</div></div>'
    +'</div>'
    +'<div style="display:flex;gap:14px;align-items:flex-start">'
    +'<div style="flex:1;font-size:13px;color:#1e293b;line-height:1.85" dir="auto">'
    +'<div>Name : '+(st.firstName||'')+' '+(st.lastName||'')+'</div>'
    +'<div>ID : '+(st.schoolID||st.username||'—')+'</div>'
    +'<div>Grade : '+((t&&t.grade)||'—')+'</div>'
    +'<div>Section : '+(st.section||'—')+'</div>'
    +'<div style="margin-top:6px;display:flex;gap:4px;flex-wrap:wrap">'+badge+extraBadges+'</div>'
    +'</div>'
    +'<img src="'+qrUrl+'" style="width:100px;height:100px;flex-shrink:0;border-radius:8px" alt="QR">'
    +'</div>'
    +'<ul style="margin-top:12px;padding-right:18px;font-size:12px;color:#334155;line-height:1.9">'
    +'<li style="font-family:Montserrat,sans-serif">scholastic.app — Student Login</li>'
    +'<li>Username: <span style="background:#fde68a;font-weight:800;padding:1px 8px;border-radius:5px;font-family:Montserrat,sans-serif">'+st.username+'</span></li>'
    +'<li>Password: <span style="background:#fde68a;font-weight:800;padding:1px 8px;border-radius:5px;font-family:Montserrat,sans-serif">'+st.password+'</span></li>'
    +'</ul>'
    +'</div></div>';
}
function printStudentCards(){
  var el=document.getElementById('studentCardsPrintArea');
  if(!el){window.print();return;}
  var win=window.open('','_blank');
  win.document.write('<html><head><meta charset="UTF-8"><title>Student ID Cards</title><style>body{font-family:Tajawal,Arial,sans-serif;margin:20px;background:white}.grid{display:grid;grid-template-columns:repeat(2,1fr);gap:20px}@media print{body{margin:0}.id-card-outer{break-inside:avoid}}</style></head><body><div class="grid">'+el.innerHTML+'</div></body></html>');
  win.document.close();
  setTimeout(function(){win.print();},400);
}
function shareStudentCards(){
  var t=myTests.find(function(x){return x.id===activeCodesTestId;});
  var text=(t?t.name+' — ':'')+'بطاقات دخول الطلاب / Student Login Credentials:\n\n'+pendingStudentCodes.map(function(st){return (st.firstName||'')+' '+(st.lastName||'')+' — Username: '+st.username+' — Password: '+st.password;}).join('\n');
  if(navigator.share){
    navigator.share({title:'Student Login Cards',text:text}).catch(function(){});
  } else if(navigator.clipboard&&navigator.clipboard.writeText){
    navigator.clipboard.writeText(text).then(function(){
      scOk('تم النسخ ✅','Copied','تم نسخ بيانات الدخول لجميع الطلاب — يمكنك لصقها ومشاركتها الآن','Student credentials copied — you can paste and share them now','📋');
    }).catch(function(){
      scWarn('تعذّر النسخ التلقائي، استخدم زر الطباعة لحفظ PDF ومشاركته','Automatic copy failed — use the print button to save a PDF and share it');
    });
  } else {
    scWarn('المشاركة غير مدعومة في هذا المتصفح — استخدم زر الطباعة لحفظ PDF ومشاركته','Sharing is not supported in this browser — use the print button to save a PDF and share it');
  }
}
function activateTestWindow(){var from=document.getElementById('testWindowFrom').value,to=document.getElementById('testWindowTo').value;if(!from||!to){alert('حدد وقت البداية والنهاية');return;}var t=myTests.find(function(x){return x.id===activeCodesTestId;});if(!t)return;t.windowFrom=from;t.windowTo=to;t.active=true;saveMyTests();alert('✅ تم تفعيل الاختبار!\nFrom: '+new Date(from).toLocaleString()+'\nTo: '+new Date(to).toLocaleString());}
function deactivateTest(){var t=myTests.find(function(x){return x.id===activeCodesTestId;});if(!t)return;scConfirm('تعطيل الاختبار','Deactivate Test','هل تريد تعطيل هذا الاختبار؟','Deactivate this test?','⛔').then(function(ok){if(!ok)return;t.active=false;saveMyTests();scOk('تم التعطيل ⛔','Deactivated','تم تعطيل الاختبار','Test deactivated','⛔');});}

// ============================================================
// STANDARDS BANK
// ============================================================
var SB_GRADES=[{id:'FS1',label:'KS1/FS1',icon:'🌱'},{id:'FS2',label:'KS1/FS2',icon:'🌿'},{id:'G1',label:'Grade 1/Y2',icon:'📗'},{id:'G2',label:'Grade 2/Y3',icon:'📘'},{id:'G3',label:'Grade 3/Y4',icon:'📙'},{id:'G4',label:'Grade 4/Y5',icon:'📒'},{id:'G5',label:'Grade 5/Y6',icon:'📕'},{id:'G6',label:'Grade 6/Y7',icon:'📓'},{id:'G7',label:'Grade 7/Y8',icon:'📔'},{id:'G8',label:'Grade 8/Y9',icon:'📃'},{id:'G9',label:'Grade 9/Y10',icon:'📑'},{id:'G10',label:'Grade 10/Y11',icon:'🎓'},{id:'G11',label:'Grade 11/Y12',icon:'🏅'},{id:'G12',label:'Grade 12/Y13',icon:'🎖️'}];
var SB_SUBJECTS=[{id:'arabic_arabs',label:'العربية للناطقين',labelEn:'Arabic Native',icon:'📖'},{id:'arabic_non',label:'العربية لغير الناطقين',labelEn:'Arabic Non-Native',icon:'📚'},{id:'islamic_arabs',label:'الإسلامية للناطقين',labelEn:'Islamic (Arabic)',icon:'🕌'},{id:'islamic_non',label:'الإسلامية لغير الناطقين',labelEn:'Islamic (English)',icon:'🕋'},{id:'social_arabs',label:'الاجتماعيات للناطقين',labelEn:'Social (Arabic)',icon:'🌍'},{id:'social_non',label:'الاجتماعيات لغير الناطقين',labelEn:'Social (English)',icon:'🗺️'},{id:'english_1st',label:'الإنجليزية لغة أولى',labelEn:'English 1st Lang',icon:'🇬🇧'},{id:'english_2nd',label:'الإنجليزية لغة ثانية',labelEn:'English 2nd Lang',icon:'🔤'},{id:'french_2nd',label:'الفرنسية لغة ثانية',labelEn:'French 2nd Lang',icon:'🇫🇷'},{id:'math_arabs',label:'الرياضيات للناطقين',labelEn:'Math (Arabic)',icon:'🔢'},{id:'math_non',label:'الرياضيات لغير الناطقين',labelEn:'Math (English)',icon:'➗'},{id:'science_arabs',label:'العلوم للناطقين',labelEn:'Science (Arabic)',icon:'🔬'},{id:'science_non',label:'العلوم لغير الناطقين',labelEn:'Science (English)',icon:'⚗️'},{id:'history_national',label:'التاريخ الوطني',labelEn:'History National',icon:'🏛️'},{id:'history_international',label:'التاريخ الدولي',labelEn:'History International',icon:'🌐'},{id:'other',label:'أخرى',labelEn:'Other',icon:'📋'}];
function saveGLOs(){return ghWrite('glos.json',sbGLOs);}
function openStandardsBank(){sb_currentGrade=null;sb_currentSubject=null;renderSBGrades();showSBLevel('grades');updateSBBreadcrumb();document.getElementById('standardsBankModal').classList.remove('hidden');}
function closeStandardsBank(){document.getElementById('standardsBankModal').classList.add('hidden');}
function showSBLevel(level){['grades','subjects','glos'].forEach(function(l){var el=document.getElementById('sb'+l.charAt(0).toUpperCase()+l.slice(1)+'Level');if(el)el.classList.toggle('hidden',l!==level);});}
function updateSBBreadcrumb(){var bc=document.getElementById('sbBreadcrumb');var html='<span style="cursor:pointer;color:#FACC15" onclick="sbGoGrades()">🎯 بنك المعايير</span>';if(sb_currentGrade){var g=SB_GRADES.find(function(x){return x.id===sb_currentGrade;});html+='<span style="color:rgba(255,255,255,.3);margin:0 6px">›</span><span style="cursor:pointer;color:#93c5fd" onclick="sbGoSubjects()">'+g.icon+' '+g.label+'</span>';}if(sb_currentSubject){var s=SB_SUBJECTS.find(function(x){return x.id===sb_currentSubject;});html+='<span style="color:rgba(255,255,255,.3);margin:0 6px">›</span><span style="color:rgba(255,255,255,.8)">'+s.icon+' '+s.labelEn+'</span>';}bc.innerHTML=html;}
function renderSBGrades(){var grid=document.getElementById('sbGradesGrid');grid.innerHTML=SB_GRADES.map(function(g){var total=Object.keys(sbGLOs).filter(function(k){return k.indexOf(g.id+'_')===0;}).reduce(function(a,k){return a+(sbGLOs[k]?sbGLOs[k].length:0);},0);return '<div onclick="sbSelectGrade(\''+g.id+'\')" class="glass-card" style="min-height:120px;padding:16px 10px;cursor:pointer"><div style="font-size:28px;margin-bottom:8px">'+g.icon+'</div><div style="font-weight:800;font-size:12px;text-align:center;line-height:1.4">'+g.label+'</div><div style="margin-top:6px;font-size:10px;font-family:Montserrat,sans-serif;'+(total>0?'color:#FACC15':'color:rgba(255,255,255,.2)')+'">'+(total>0?total+' GLOs':'Empty')+'</div></div>';}).join('');}
function sbSelectGrade(id){sb_currentGrade=id;renderSBSubjects();showSBLevel('subjects');updateSBBreadcrumb();}
function sbGoGrades(){sb_currentGrade=null;sb_currentSubject=null;renderSBGrades();showSBLevel('grades');updateSBBreadcrumb();}
function sbGoSubjects(){sb_currentSubject=null;renderSBSubjects();showSBLevel('subjects');updateSBBreadcrumb();}
function renderSBSubjects(){var grid=document.getElementById('sbSubjectsGrid');grid.innerHTML=SB_SUBJECTS.map(function(s){var key=sb_currentGrade+'_'+s.id;var count=sbGLOs[key]?sbGLOs[key].length:0;return '<div onclick="sbSelectSubject(\''+s.id+'\')" class="glass-card" style="min-height:130px;padding:16px 12px;cursor:pointer"><div style="font-size:28px;margin-bottom:8px">'+s.icon+'</div><div style="font-weight:800;font-size:12px;text-align:center;margin-bottom:4px">'+s.label+'</div><div style="font-size:10px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif;text-align:center">'+s.labelEn+'</div><div style="margin-top:6px;font-size:10px;font-family:Montserrat,sans-serif;'+(count>0?'color:#FACC15':'color:rgba(255,255,255,.2)')+'">'+(count>0?count+' GLOs':'Empty')+'</div></div>';}).join('');}
function sbSelectSubject(id){sb_currentSubject=id;renderGLOTable();showSBLevel('glos');updateSBBreadcrumb();document.getElementById('gloAr').value='';document.getElementById('gloEn').value='';document.getElementById('uploadStatus').textContent='';}
function getGLOKey(){return sb_currentGrade+'_'+sb_currentSubject;}
function getGLOCode(index){var s=SB_SUBJECTS.find(function(x){return x.id===sb_currentSubject;});var subCode=s?s.labelEn.split(' ').slice(0,2).join('-').replace(/[^a-zA-Z0-9-]/g,''):sb_currentSubject;return 'GLO:'+subCode+'/'+sb_currentGrade+'/'+String(index+1).padStart(2,'0');}
function renderGLOTable(){var key=getGLOKey();var list=sbGLOs[key]||[];var tbody=document.getElementById('gloTableBody');var empty=document.getElementById('gloEmpty');if(!list.length){tbody.innerHTML='';empty.style.display='block';return;}empty.style.display='none';tbody.innerHTML=list.map(function(glo,i){return '<tr><td class="font-en" style="color:#FACC15;font-size:11px;direction:ltr">'+getGLOCode(i)+'</td><td style="text-align:right;direction:rtl">'+(glo.ar||'—')+'</td><td style="text-align:left;direction:ltr;font-family:Montserrat,sans-serif;font-size:12px">'+(glo.en||'—')+'</td><td><button onclick="deleteGLO('+i+')" style="color:#f87171;background:none;border:none;cursor:pointer">🗑</button></td></tr>';}).join('');}
function addGLOManual(){var ar=document.getElementById('gloAr').value.trim(),en=document.getElementById('gloEn').value.trim();if(!ar&&!en){alert('أدخل ناتج التعلم');return;}var key=getGLOKey();if(!sbGLOs[key])sbGLOs[key]=[];sbGLOs[key].push({ar:ar,en:en});saveGLOs();document.getElementById('gloAr').value='';document.getElementById('gloEn').value='';renderGLOTable();}
function deleteGLO(index){scConfirm('حذف ناتج التعلم','Delete GLO','هل تريد حذف هذا الناتج؟','Delete this GLO?','🗑').then(function(ok){if(!ok)return;var key=getGLOKey();sbGLOs[key].splice(index,1);saveGLOs();renderGLOTable();});}
function uploadGLOFile(event){var file=event.target.files?event.target.files[0]:null;if(!file)return;var status=document.getElementById('uploadStatus');status.textContent='⏳';var reader=new FileReader();reader.onload=function(e){try{var lines=e.target.result.split('\n').filter(function(l){return l.trim();});var key=getGLOKey();if(!sbGLOs[key])sbGLOs[key]=[];var added=0;lines.forEach(function(line,idx){if(idx===0&&line.toLowerCase().indexOf('arabic')>=0)return;var cols=line.split(',');var ar=(cols[0]||'').replace(/"/g,'').trim();var en=(cols[1]||'').replace(/"/g,'').trim();if(ar||en){sbGLOs[key].push({ar:ar,en:en});added++;}});saveGLOs();renderGLOTable();status.textContent='✅ '+added;}catch(err){status.textContent='❌';}};reader.readAsText(file,'UTF-8');event.target.value='';}

// ============================================================
// MISC
// ============================================================
// ============================================================
// GRADING SYSTEM — التصحيح المباشر
// ============================================================
var _gradingTestId = null;
var _gradingRole = 'sup'; // 'sup' | 'admin'
var _gradingStudentIdx = -1;

// ── Helper: get approved tests for current supervisor ──
function _getApprovedTests(){
  return myTests.filter(function(t){
    if(t.status !== 'approved') return false;
    if(_gradingRole === 'admin') return true;
    // supervisor sees their own authored tests + shared grading tests
    return t.author === currentSupName || (t.gradingSharedWith && t.gradingSharedWith.indexOf(currentSupName) >= 0);
  });
}

// ── Close any sub-panel and show supMainOptions ──
function closeSupPanel(id){
  document.getElementById(id).classList.add('hidden');
  document.getElementById('supMainOptions').classList.remove('hidden');
  document.getElementById('supTitle').classList.remove('hidden');
}

function _hideSupMain(){
  document.getElementById('supMainOptions').classList.add('hidden');
  document.getElementById('supTitle').classList.add('hidden');
}

// ============================================================
// LIVE MONITOR (Supervisor)
// ============================================================
var _supLiveRefresh = null;

function openSupLiveMonitor(){
  _hideSupMain();
  document.getElementById('supLiveMonitorPanel').classList.remove('hidden');
  _populateSupLiveTestSelect();
}

function _populateSupLiveTestSelect(){
  var tests = myTests.filter(function(t){ return t.status === 'approved' && t.students && t.students.length; });
  var sel = document.getElementById('supLiveTestSelect');
  sel.innerHTML = '<option value="">-- اختر اختباراً معتمداً --</option>'
    + tests.map(function(t){
        return '<option value="'+t.id+'">'+t.name+' | '+(t.subject||'')+(t.grade?' | '+t.grade:'')+'</option>';
      }).join('');
  document.getElementById('supLiveInfoBar').classList.add('hidden');
  document.getElementById('supLiveStats').innerHTML = '';
  document.getElementById('supLiveTableHead').innerHTML = '<tr><th colspan="8" style="color:rgba(255,255,255,.3);font-weight:400">اختر اختباراً</th></tr>';
  document.getElementById('supLiveTableBody').innerHTML = '';
  document.getElementById('supSendGradingWrap').classList.add('hidden');
}

function loadSupLiveMonitor(){
  var testId = parseInt(document.getElementById('supLiveTestSelect').value);
  if(!testId){ document.getElementById('supLiveInfoBar').classList.add('hidden'); return; }
  var t = myTests.find(function(x){ return x.id === testId; });
  if(!t) return;
  _gradingTestId = testId;
  var students = t.students || [];
  var started = students.filter(function(s){ return s.started; }).length;
  var completed = students.filter(function(s){ return s.completed; }).length;
  var notStarted = students.length - started;

  // Info bar
  document.getElementById('supLiveInfoBar').classList.remove('hidden');
  document.getElementById('supLiveTestName').textContent = t.name || '—';
  document.getElementById('supLiveSubject').textContent = t.subject || '—';
  document.getElementById('supLiveGrade').textContent = t.grade || '—';
  document.getElementById('supLiveTotal').textContent = students.length;
  document.getElementById('supLiveStarted').textContent = started;
  document.getElementById('supLiveCompleted').textContent = completed;

  // Stats cards
  document.getElementById('supLiveStats').innerHTML = [
    {label:'الإجمالي', val:students.length, color:'#FACC15', icon:'👥'},
    {label:'لم يبدأ', val:notStarted, color:'#f87171', icon:'⏳'},
    {label:'بدأ', val:started - completed, color:'#7dd3fc', icon:'🔄'},
    {label:'أكمل', val:completed, color:'#4ade80', icon:'✅'}
  ].map(function(s){
    return '<div style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:14px;text-align:center">'
      +'<div style="font-size:22px;margin-bottom:4px">'+s.icon+'</div>'
      +'<div style="font-size:24px;font-weight:900;color:'+s.color+';font-family:Montserrat,sans-serif">'+s.val+'</div>'
      +'<div style="font-size:11px;color:rgba(255,255,255,.5)">'+s.label+'</div>'
      +'</div>';
  }).join('');

  // Build domain headers
  var domains = t.domains || [];
  var domHeaders = domains.map(function(d,i){
    return '<th style="font-size:11px;white-space:nowrap">'+(d.nameAr||'مجال '+(i+1))+'<br>'
      +(d.hasBranches ? '<span style="color:#7dd3fc;font-size:9px">(فروع)</span>' : '')
      +'</th>';
  }).join('');

  document.getElementById('supLiveTableHead').innerHTML =
    '<tr style="font-size:11px">'
    +'<th>الاسم الأول</th><th>الاسم الثاني</th>'
    +'<th>School ID</th><th>الحالة</th>'
    +domHeaders
    +'<th>التصحيح</th>'
    +'</tr>';

  document.getElementById('supLiveTableBody').innerHTML = students.map(function(st){
    var statusBadge = st.completed
      ? '<span style="background:rgba(34,197,94,.2);color:#4ade80;border:1px solid rgba(34,197,94,.4);border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700">✅ أكمل</span>'
      : st.started
        ? '<span style="background:rgba(56,189,248,.2);color:#7dd3fc;border:1px solid rgba(56,189,248,.4);border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700">🔄 جاري</span>'
        : '<span style="background:rgba(248,113,113,.2);color:#f87171;border:1px solid rgba(248,113,113,.4);border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700">⏳ لم يبدأ</span>';

    var domCells = domains.map(function(d,di){
      var done = st.domainsCompleted && st.domainsCompleted[di];
      return '<td style="text-align:center;font-size:12px">'
        +(done ? '✅' : '<span style="color:rgba(255,255,255,.3)">—</span>')
        +'</td>';
    }).join('');

    var gradingDone = st.grades && Object.keys(st.grades).length > 0;
    var gradingBadge = gradingDone
      ? '<span style="background:rgba(251,146,60,.2);color:#fb923c;border-radius:20px;padding:2px 10px;font-size:10px">✍️ جاري</span>'
      : '<span style="color:rgba(255,255,255,.2);font-size:11px">—</span>';

    return '<tr style="font-size:12px">'
      +'<td>'+(st.firstName||'—')+'</td>'
      +'<td>'+(st.lastName||'—')+'</td>'
      +'<td class="font-en" style="color:#FACC15;font-size:10px">'+(st.schoolID||st.username||'—')+'</td>'
      +'<td>'+statusBadge+'</td>'
      +domCells
      +'<td>'+gradingBadge+'</td>'
      +'</tr>';
  }).join('');

  // Show send-to-grading button if any student completed
  if(completed > 0){
    document.getElementById('supSendGradingWrap').classList.remove('hidden');
  }
}

function refreshSupLiveMonitor(){
  showGHLoader(true);
  ghRead('tests.json').then(function(data){
    if(data) myTests = data;
    showGHLoader(false);
    loadSupLiveMonitor();
  });
}

function sendToGrading(){
  if(!_gradingTestId){ scWarn('اختر اختباراً أولاً','Select a test first'); return; }
  var t = myTests.find(function(x){ return x.id === _gradingTestId; });
  if(!t){ return; }
  scConfirm('إرسال للتصحيح','Send to Marking',
    'سيتم إرسال اختبار <strong style="color:#fb923c">'+t.name+'</strong> للتصحيح المباشر.',
    'Send test <strong>'+t.name+'</strong> to live marking?'
  ).then(function(ok){
    if(!ok) return;
    t.gradingStatus = 'inProgress';
    t.gradingSentAt = new Date().toISOString();
    t.gradingMarker = currentSupName;
    saveMyTests().then(function(){
      scOk('تم الإرسال ✅','Sent','تم إرسال الاختبار للتصحيح المباشر.','Test sent to marking.','✍️').then(function(){
        closeSupPanel('supLiveMonitorPanel');
        openGradingPanel('sup', _gradingTestId);
      });
    });
  });
}

// ============================================================
// GRADING PANEL — التصحيح المباشر
// ============================================================
function openGradingPanel(role, preloadTestId){
  _gradingRole = role || 'sup';
  _hideSupMain();
  document.getElementById('gradingPanel').classList.remove('hidden');
  document.getElementById('gradingMarker').textContent = currentSupName || 'المصحح';

  // Populate student selector from approved/inProgress tests
  var tests = myTests.filter(function(t){
    if(t.status !== 'approved') return false;
    if(role === 'admin') return true;
    return t.author === currentSupName
      || (t.gradingSharedWith && t.gradingSharedWith.indexOf(currentSupName) >= 0)
      || t.gradingMarker === currentSupName;
  });

  // If preloaded test
  if(preloadTestId){
    _gradingTestId = preloadTestId;
  }

  // Build test+student dropdown
  var sel = document.getElementById('gradingStudentSelect');
  sel.innerHTML = '<option value="">-- اختر طالباً --</option>';
  tests.forEach(function(t){
    if(!t.students || !t.students.length) return;
    var grp = document.createElement('optgroup');
    grp.label = t.name + ' | ' + (t.subject||'') + ' | ' + (t.grade||'');
    t.students.forEach(function(st, si){
      var opt = document.createElement('option');
      opt.value = t.id + '_' + si;
      opt.textContent = (st.firstName||'') + ' ' + (st.lastName||'') + ' — ' + (st.schoolID || st.username || '');
      var gradedCount = st.grades ? Object.keys(st.grades).length : 0;
      if(gradedCount > 0) opt.textContent += ' ✍️';
      grp.appendChild(opt);
    });
    sel.appendChild(grp);
  });

  // Auto-select first student of preloaded test
  if(preloadTestId){
    var t2 = myTests.find(function(x){ return x.id === preloadTestId; });
    if(t2 && t2.students && t2.students.length){
      sel.value = preloadTestId + '_0';
      loadStudentGrading();
    }
    document.getElementById('gradingTestLabel').textContent = (t2 ? t2.name + ' | ' + (t2.subject||'') : '');
  }

  // Status
  var statuses = {draft:'🖊 مسودة', shared:'👥 مشتركة', approved:'✅ معتمد من المحرر', finalApproved:'🏆 اعتماد نهائي'};
  var t3 = _gradingTestId ? myTests.find(function(x){ return x.id === _gradingTestId; }) : null;
  document.getElementById('gradingStatus').textContent = t3 ? (statuses[t3.gradingStatus]||'مسودة') : 'مسودة';
  document.getElementById('gradingLastSave').textContent = t3 && t3.gradingSavedAt ? new Date(t3.gradingSavedAt).toLocaleString('ar') : '—';
}

function loadStudentGrading(){
  var val = document.getElementById('gradingStudentSelect').value;
  if(!val){ document.getElementById('gradingContent').innerHTML = '<div style="text-align:center;padding:60px;color:rgba(255,255,255,.3)">اختر طالباً لبدء التصحيح</div>'; return; }
  var parts = val.split('_');
  var testId = parseInt(parts[0]);
  var stIdx = parseInt(parts[1]);
  _gradingTestId = testId;
  _gradingStudentIdx = stIdx;

  var t = myTests.find(function(x){ return x.id === testId; });
  if(!t || !t.students || !t.students[stIdx]){ return; }
  var st = t.students[stIdx];
  var domains = t.domains || [];

  // Progress indicator
  var totalQ = 0, gradedQ = 0;
  domains.forEach(function(d, di){
    var qs = d.hasBranches
      ? (d.branches||[]).reduce(function(a,b){ return a.concat(b.questions||[]); },[])
      : (d.questions||[]);
    totalQ += qs.length;
    if(st.grades && st.grades[di]){
      gradedQ += Object.keys(st.grades[di]).length;
    }
  });
  document.getElementById('gradingStudentProgress').textContent = 'تم تصحيح '+gradedQ+' / '+totalQ+' سؤال';

  // Build grading UI
  var html = '<div style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:16px;margin-bottom:16px">'
    +'<div style="display:flex;gap:16px;flex-wrap:wrap;align-items:center">'
    +'<div style="font-size:18px;font-weight:900;color:white">'+(st.firstName||'')+' '+(st.lastName||'')+'</div>'
    +'<div style="font-size:12px;color:rgba(255,255,255,.5);font-family:Montserrat,sans-serif">'+(st.schoolID||st.username||'')+'</div>'
    +'<div style="font-size:12px;color:rgba(255,255,255,.5)">'+(st.gender||'')+'</div>'
    +(st.gifted==='Yes'?'<span style="background:rgba(250,204,21,.2);color:#FACC15;border:1px solid rgba(250,204,21,.4);border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700">⭐ موهوب</span>':'')
    +(st.sod==='Yes'?'<span style="background:rgba(167,139,250,.2);color:#c4b5fd;border:1px solid rgba(167,139,250,.4);border-radius:20px;padding:2px 10px;font-size:10px;font-weight:700">♿ SOD</span>':'')
    +'</div></div>';

  // Loop domains
  domains.forEach(function(d, di){
    var questions = d.hasBranches
      ? (d.branches||[]).reduce(function(acc, br, bi){
          return acc.concat((br.questions||[]).map(function(q,qi){ return Object.assign({},q,{_br:bi,_brName:br.nameAr||'فرع '+(bi+1),_qi:qi}); }));
        },[])
      : (d.questions||[]).map(function(q,qi){ return Object.assign({},q,{_qi:qi}); });

    if(!questions.length) return;

    html += '<div style="background:rgba(249,115,22,.06);border:2px solid rgba(249,115,22,.25);border-radius:18px;padding:18px;margin-bottom:16px">';
    html += '<div style="display:flex;align-items:center;gap:12px;margin-bottom:14px">';
    html += '<div style="width:36px;height:36px;border-radius:50%;background:linear-gradient(135deg,#f97316,#ea580c);display:flex;align-items:center;justify-content:center;font-weight:900;color:white;font-family:Montserrat,sans-serif">'+(di+1)+'</div>';
    html += '<div style="font-size:16px;font-weight:800;color:#fb923c">'+(d.nameAr||'مجال '+(di+1))+'</div>';
    if(d.nameEn) html += '<div style="font-size:11px;color:rgba(255,255,255,.4);font-family:Montserrat,sans-serif">'+d.nameEn+'</div>';
    html += '<div style="margin-right:auto;font-size:13px;font-weight:700;color:#FACC15">'+d.weight+'%</div>';
    html += '</div>';

    questions.forEach(function(q, qi_global){
      var brKey = q._br !== undefined ? 'br'+q._br+'_' : '';
      var gradeKey = 'g_'+di+'_'+brKey+q._qi;
      var existingGrade = st.grades && st.grades[di] && st.grades[di][brKey+q._qi] !== undefined
        ? st.grades[di][brKey+q._qi] : '';
      var isAuto = q.type === 'mcq' || q.type === 'truefalse' || q.type === 'ordering';
      var autoResult = _getAutoGrade(q, st, di, q._br, q._qi);

      html += '<div style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);border-radius:14px;padding:14px;margin-bottom:10px">';

      // Question header
      html += '<div style="display:flex;align-items:flex-start;gap:10px;margin-bottom:10px">';
      html += '<div style="min-width:36px;height:36px;border-radius:12px;background:#0c4a6e;border:2px solid #0ea5e9;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:12px;color:#7dd3fc;font-family:Montserrat,sans-serif;flex-shrink:0">Q.'+(qi_global+1)+'</div>';
      html += '<div style="flex:1">';
      if(q._brName) html += '<div style="font-size:10px;color:#7dd3fc;font-weight:700;margin-bottom:3px">'+q._brName+'</div>';
      html += '<div style="font-size:14px;font-weight:600;color:white;line-height:1.6" dir="auto">'+(q.stemHtml||q.stemText||'—')+'</div>';
      html += '<div style="display:flex;gap:8px;margin-top:6px;flex-wrap:wrap">';
      html += '<span style="background:rgba(255,255,255,.08);border-radius:8px;padding:2px 10px;font-size:10px;font-family:Montserrat,sans-serif;color:rgba(255,255,255,.5)">'+(q.type||'—')+'</span>';
      html += '<span style="background:rgba(250,204,21,.1);border-radius:8px;padding:2px 10px;font-size:10px;font-family:Montserrat,sans-serif;color:#FACC15">'+q.score+'%</span>';
      html += isAuto
        ? '<span style="background:rgba(34,197,94,.1);border-radius:8px;padding:2px 10px;font-size:10px;color:#4ade80">⚡ تلقائي</span>'
        : '<span style="background:rgba(251,146,60,.1);border-radius:8px;padding:2px 10px;font-size:10px;color:#fb923c">✍️ يدوي — إجباري</span>';
      html += '</div></div></div>';

      // Student answer
      html += '<div style="background:rgba(0,0,0,.25);border-radius:10px;padding:10px 14px;margin-bottom:10px">';
      html += '<div style="font-size:10px;color:rgba(255,255,255,.4);margin-bottom:6px;font-family:Montserrat,sans-serif">إجابة الطالب / Student Answer</div>';
      html += _renderStudentAnswer(q, st, di, q._br, q._qi);
      html += '</div>';

      // Auto suggestion
      if(isAuto && autoResult !== null){
        html += '<div style="background:rgba(34,197,94,.08);border:1px solid rgba(34,197,94,.2);border-radius:10px;padding:8px 14px;margin-bottom:10px;display:flex;align-items:center;gap:10px">';
        html += '<span style="font-size:11px;color:#4ade80;font-family:Montserrat,sans-serif">⚡ الدرجة المقترحة تلقائياً:</span>';
        html += '<strong style="color:#4ade80;font-family:Montserrat,sans-serif">'+(autoResult ? q.score : 0)+'%</strong>';
        html += '<span style="font-size:10px;color:rgba(74,222,128,.6)">'+(autoResult ? '✅ إجابة صحيحة' : '❌ إجابة خاطئة')+'</span>';
        html += '<button onclick="applyAutoGrade(\''+gradeKey+'\','+di+',\''+brKey+q._qi+'\','+q.score+','+autoResult+')" style="margin-right:auto;background:rgba(34,197,94,.2);border:1px solid rgba(34,197,94,.4);color:#4ade80;border-radius:8px;padding:4px 12px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif">اعتماد / Accept</button>';
        html += '</div>';
      }

      // Grade input
      html += '<div style="display:flex;align-items:center;gap:12px;flex-wrap:wrap">';
      html += '<label style="font-size:12px;color:#FACC15;font-weight:700">الدرجة الممنوحة %</label>';
      html += '<input type="number" id="'+gradeKey+'" class="wizard-input font-en" value="'+(existingGrade!==''?existingGrade:'')+'" min="0" max="'+q.score+'" step="0.01" placeholder="0 – '+q.score+'" style="width:100px;font-size:14px;font-weight:800;color:#FACC15" oninput="updateGradeTotal('+di+')">';
      html += '<span style="font-size:11px;color:rgba(255,255,255,.4)">من / of '+q.score+'%</span>';
      if(!isAuto) html += '<span style="font-size:11px;color:#fb923c;font-weight:700">⚠️ إجباري</span>';
      html += '</div>';
      html += '</div>'; // end question card
    });

    // Domain total row
    html += '<div id="domain-total-'+di+'" style="background:rgba(250,204,21,.08);border:1px solid rgba(250,204,21,.2);border-radius:10px;padding:8px 14px;margin-top:8px;display:flex;align-items:center;justify-content:space-between">';
    html += '<span style="font-size:12px;color:#FACC15;font-weight:700">مجموع المجال '+(di+1)+':</span>';
    html += '<span id="domain-total-val-'+di+'" style="font-size:16px;font-weight:900;color:#FACC15;font-family:Montserrat,sans-serif">0 / '+d.weight+'%</span>';
    html += '</div>';
    html += '</div>'; // end domain card
  });

  // Overall total
  html += '<div style="background:linear-gradient(135deg,rgba(249,115,22,.15),rgba(234,88,12,.1));border:2px solid rgba(249,115,22,.4);border-radius:16px;padding:16px;text-align:center">';
  html += '<div style="font-size:13px;color:rgba(255,255,255,.6);margin-bottom:6px">المجموع الكلي / Total Score</div>';
  html += '<div id="grading-grand-total" style="font-size:32px;font-weight:900;color:#fb923c;font-family:Montserrat,sans-serif">0%</div>';
  html += '</div>';

  document.getElementById('gradingContent').innerHTML = html;

  // Init totals
  var t4 = myTests.find(function(x){ return x.id === testId; });
  if(t4) t4.domains.forEach(function(_,di){ updateGradeTotal(di); });
}

function _getAutoGrade(q, st, di, brIdx, qi){
  if(!st.domainsCompleted) return null;
  var ans = null;
  if(st.answers && st.answers[di]){
    ans = brIdx !== undefined ? (st.answers[di]['br'+brIdx+'_'+qi]) : st.answers[di][qi];
  }
  if(ans === undefined || ans === null) return null;
  if(q.type === 'mcq') return ans === q.correct;
  if(q.type === 'truefalse'){
    if(!q.statements || !q.statements.length) return null;
    if(typeof ans !== 'object') return null;
    var correct = q.statements.every(function(s,i){ return ans[i] === s.answer; });
    return correct;
  }
  if(q.type === 'ordering'){
    if(!ans) return null;
    if(ans.groups){
      var oGroups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words?[{words:q.words}]:[]);
      if(!oGroups.length) return null;
      return oGroups.every(function(g,gi){
        var gg=ans.groups[gi];
        if(!gg||!gg.placed) return false;
        return JSON.stringify(gg.placed)===JSON.stringify(g.words||[]);
      });
    }
    if(!ans.order) return null;
    return JSON.stringify(ans.order) === JSON.stringify(q.words);
  }
  return null;
}

function _renderStudentAnswer(q, st, di, brIdx, qi){
  var ans = null;
  if(st.answers && st.answers[di]){
    var key = brIdx !== undefined ? 'br'+brIdx+'_'+qi : qi;
    ans = st.answers[di][key];
  }
  if(ans === undefined || ans === null) return '<span style="color:rgba(255,255,255,.3);font-size:13px">لم يجب / No answer</span>';
  var labels = ['A','B','C','D','E','F'];
  if(q.type === 'mcq'){
    var opt = q.options && q.options[ans] ? q.options[ans] : '—';
    return '<div style="background:rgba(30,58,138,.3);border:2px solid #3b82f6;border-radius:10px;padding:8px 14px;font-size:14px;font-weight:700;color:white" dir="auto">'+(labels[ans]||ans)+'. '+opt+'</div>';
  }
  if(q.type === 'truefalse'){
    if(typeof ans !== 'object') return '<span style="color:white">'+ans+'</span>';
    return (q.statements||[]).map(function(s,i){
      var a = ans[i];
      return '<div style="display:flex;gap:8px;align-items:center;margin-bottom:4px">'
        +'<span style="font-size:13px;color:white" dir="auto">'+(i+1)+'. '+s.text+'</span>'
        +'<span style="font-weight:800;color:'+(a==='true'?'#4ade80':'#f87171')+'">'+(a==='true'?'✅ صواب':'❌ خطأ')+'</span>'
        +'</div>';
    }).join('');
  }
  if(q.type === 'ordering'){
    if(ans.groups){
      return ans.groups.map(function(g){
        var order=g.placed||[];
        return '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:6px">'+order.map(function(w,i){return '<span style="background:#1e3a8a;color:white;border-radius:8px;padding:5px 12px;font-size:13px;font-weight:700">'+(i+1)+'. '+(w||'—')+'</span>';}).join('')+'</div>';
      }).join('');
    }
    var order = ans.order || ans;
    if(!Array.isArray(order)) return '<span style="color:rgba(255,255,255,.4)">—</span>';
    return '<div style="display:flex;flex-wrap:wrap;gap:6px">'+order.map(function(w,i){
      return '<span style="background:#1e3a8a;color:white;border-radius:8px;padding:5px 12px;font-size:13px;font-weight:700">'+(i+1)+'. '+w+'</span>';
    }).join('')+'</div>';
  }
  if(q.type === 'matching'){
    var conns = ans.connections || {};
    return '<div>'+Object.keys(conns).map(function(ai){
      var bi = conns[ai];
      var aText = q.pairs && q.pairs[ai] ? q.pairs[ai].aHtml||'—' : '—';
      var bText = q.pairs && q.pairs[bi] ? q.pairs[bi].bHtml||'—' : '—';
      return '<div style="display:flex;gap:8px;align-items:center;margin-bottom:4px;font-size:13px">'
        +'<span style="color:#93c5fd" dir="auto">'+aText+'</span>'
        +'<span style="color:rgba(255,255,255,.3)">↔</span>'
        +'<span style="color:#86efac" dir="auto">'+bText+'</span>'
        +'</div>';
    }).join('')+'</div>';
  }
  if(q.type === 'speaking') return '<span style="color:#c4b5fd">🎙 تسجيل صوتي محفوظ</span>';
  if(typeof ans === 'string') return '<div style="font-size:14px;color:white;line-height:1.7" dir="auto">'+ans+'</div>';
  return '<span style="color:rgba(255,255,255,.4);font-size:12px">'+JSON.stringify(ans)+'</span>';
}

function applyAutoGrade(gradeKey, di, qKey, maxScore, isCorrect){
  var el = document.getElementById(gradeKey);
  if(el){ el.value = isCorrect ? maxScore : 0; updateGradeTotal(di); }
}

function updateGradeTotal(di){
  var t = _gradingTestId ? myTests.find(function(x){ return x.id === _gradingTestId; }) : null;
  if(!t) return;
  var d = t.domains[di];
  if(!d) return;
  var questions = d.hasBranches
    ? (d.branches||[]).reduce(function(acc,br,bi){ return acc.concat((br.questions||[]).map(function(q,qi){ return {q:q,brKey:'br'+bi+'_',qi:qi}; })); },[])
    : (d.questions||[]).map(function(q,qi){ return {q:q,brKey:'',qi:qi}; });

  var domTotal = 0;
  questions.forEach(function(item){
    var key = 'g_'+di+'_'+item.brKey+item.qi;
    var el = document.getElementById(key);
    if(el) domTotal += parseFloat(el.value)||0;
  });

  var el = document.getElementById('domain-total-val-'+di);
  if(el){ el.textContent = domTotal.toFixed(2)+' / '+d.weight+'%'; el.style.color = domTotal>d.weight?'#f87171':'#FACC15'; }

  // Grand total
  var grand = 0;
  t.domains.forEach(function(dd, ddi){
    var qs2 = dd.hasBranches
      ? (dd.branches||[]).reduce(function(acc,br,bi){ return acc.concat((br.questions||[]).map(function(q,qi){ return {brKey:'br'+bi+'_',qi:qi}; })); },[])
      : (dd.questions||[]).map(function(q,qi){ return {brKey:'',qi:qi}; });
    qs2.forEach(function(item){
      var el2 = document.getElementById('g_'+ddi+'_'+item.brKey+item.qi);
      if(el2) grand += parseFloat(el2.value)||0;
    });
  });
  var grandEl = document.getElementById('grading-grand-total');
  if(grandEl){ grandEl.textContent = grand.toFixed(2)+'%'; grandEl.style.color = grand>=50?'#4ade80':'#f87171'; }
}

function _collectGradesFromUI(){
  var t = _gradingTestId ? myTests.find(function(x){ return x.id === _gradingTestId; }) : null;
  if(!t || _gradingStudentIdx < 0) return null;
  var st = t.students[_gradingStudentIdx];
  if(!st) return null;
  if(!st.grades) st.grades = {};
  t.domains.forEach(function(d, di){
    if(!st.grades[di]) st.grades[di] = {};
    var questions = d.hasBranches
      ? (d.branches||[]).reduce(function(acc,br,bi){ return acc.concat((br.questions||[]).map(function(q,qi){ return {brKey:'br'+bi+'_',qi:qi,q:q}; })); },[])
      : (d.questions||[]).map(function(q,qi){ return {brKey:'',qi:qi,q:q}; });
    questions.forEach(function(item){
      var el = document.getElementById('g_'+di+'_'+item.brKey+item.qi);
      if(el && el.value !== '') st.grades[di][item.brKey+item.qi] = parseFloat(el.value)||0;
    });
  });
  return st;
}

function saveGrading(){
  _collectGradesFromUI();
  var t = _gradingTestId ? myTests.find(function(x){ return x.id === _gradingTestId; }) : null;
  if(!t){ scWarn('اختر اختباراً أولاً','Select a test'); return; }
  t.gradingSavedAt = new Date().toISOString();
  t.gradingStatus = t.gradingStatus || 'draft';
  t.gradingMarker = currentSupName;
  saveMyTests().then(function(){
    document.getElementById('gradingLastSave').textContent = new Date().toLocaleString('ar');
    scOk('تم الحفظ 💾','Saved','تم حفظ درجات التصحيح.','Grading saved.','💾');
  });
}

function shareGrading(){
  _collectGradesFromUI();
  var t = _gradingTestId ? myTests.find(function(x){ return x.id === _gradingTestId; }) : null;
  if(!t){ scWarn('اختر اختباراً','Select a test'); return; }
  // Show supervisor selection
  var supList = supervisors.filter(function(s){ return s.active && s.name !== currentSupName; });
  supList.unshift(DEFAULT_SUP);
  var opts = supList.map(function(s,i){ return '<label style="display:flex;align-items:center;gap:8px;padding:8px;border-radius:10px;cursor:pointer;background:rgba(255,255,255,.05);margin-bottom:6px"><input type="checkbox" value="'+s.name+'" style="accent-color:#c4b5fd;width:16px;height:16px"> <span style="font-size:14px">'+s.name+'</span></label>'; }).join('');
  scAlert({
    icon:'👥',
    titleAr:'مشاركة التصحيح',
    titleEn:'Share Grading',
    msgAr:'اختر المشرفين للمشاركة:<br><div style="margin-top:10px" id="shareSupsBox">'+opts+'</div>',
    type:'confirm',
    confirmAr:'مشاركة',
    confirmEn:'Share'
  }).then(function(ok){
    if(!ok) return;
    var checked = document.querySelectorAll('#shareSupsBox input:checked');
    var names = Array.from(checked).map(function(c){ return c.value; });
    if(!names.length){ scWarn('اختر مشرفاً على الأقل','Select at least one supervisor'); return; }
    if(!t.gradingSharedWith) t.gradingSharedWith = [];
    names.forEach(function(n){ if(t.gradingSharedWith.indexOf(n)<0) t.gradingSharedWith.push(n); });
    t.gradingStatus = 'shared';
    t.gradingSavedAt = new Date().toISOString();
    saveMyTests().then(function(){
      document.getElementById('gradingStatus').textContent = '👥 مشتركة';
      scOk('تمت المشاركة ✅','Shared','تم مشاركة التصحيح مع: '+names.join('، '),'Shared with: '+names.join(', '),'👥');
    });
  });
}

function approveGrading(){
  _collectGradesFromUI();
  var t = _gradingTestId ? myTests.find(function(x){ return x.id === _gradingTestId; }) : null;
  if(!t){ scWarn('اختر اختباراً','Select a test'); return; }
  // Validate all manual questions graded
  var ungraded = [];
  var st = _gradingStudentIdx >= 0 ? t.students[_gradingStudentIdx] : null;
  if(!st){ scWarn('اختر طالباً أولاً','Select a student'); return; }
  t.domains.forEach(function(d,di){
    var questions = d.hasBranches
      ? (d.branches||[]).reduce(function(acc,br,bi){ return acc.concat((br.questions||[]).map(function(q,qi){ return {q:q,brKey:'br'+bi+'_',qi:qi,di:di}; })); },[])
      : (d.questions||[]).map(function(q,qi){ return {q:q,brKey:'',qi:qi,di:di}; });
    questions.forEach(function(item){
      var isManual = item.q.type==='writingskill'||item.q.type==='speaking'||item.q.type==='oral'||item.q.type==='reading';
      if(isManual){
        var el = document.getElementById('g_'+di+'_'+item.brKey+item.qi);
        if(!el || el.value === '') ungraded.push('Q.'+(item.qi+1)+' في '+d.nameAr);
      }
    });
  });
  if(ungraded.length){
    scWarn('يوجد أسئلة يدوية لم تُصحح بعد:<br>'+ungraded.join('<br>'),'Manual questions not graded:<br>'+ungraded.join('<br>'));
    return;
  }
  scDanger('اعتماد التصحيح','Approve Grading',
    'هل تريد اعتماد تصحيح الطالب <strong style="color:#fb923c">'+(st.firstName||'')+' '+(st.lastName||'')+'</strong>؟<br>سيُرسل تلقائياً للإدارة للاعتمادية النهائية.',
    'Approve grading for <strong>'+(st.firstName||'')+' '+(st.lastName||'')+'</strong>? Will be sent for final admin approval.'
  ).then(function(ok){
    if(!ok) return;
    st.gradingApproved = true;
    st.gradingApprovedBy = currentSupName;
    st.gradingApprovedAt = new Date().toISOString();
    t.gradingStatus = 'approved';
    t.gradingSavedAt = new Date().toISOString();
    saveMyTests().then(function(){
      document.getElementById('gradingStatus').textContent = '✅ معتمد من المحرر';
      scOk('تم الاعتماد ✅','Approved','تم اعتماد التصحيح وإرساله للإدارة.','Grading approved and sent to admin.','✅');
    });
  });
}

// ============================================================
// GRADING COMMITTEE — لجنة التصحيح
// ============================================================
function openGradingCommittee(role){
  _gradingRole = role;
  if(role === 'admin'){
    document.getElementById('adminOptions').classList.add('hidden');
    document.getElementById('adminGradingCommitteePanel').classList.remove('hidden');
    _renderAdminCommitteeContent();
  } else {
    _hideSupMain();
    document.getElementById('gradingCommitteePanel').classList.remove('hidden');
    document.getElementById('gradingCommitteeTitle').textContent = '👥 لجنة التصحيح';
    document.getElementById('gradingCommitteeTitle').style.color = '#c4b5fd';
    renderCommitteeContent('sup');
  }
}

function _renderAdminCommitteeContent(){
  var tests = myTests.filter(function(t){
    return t.status === 'approved' && (t.gradingStatus === 'approved' || t.gradingStatus === 'shared' || t.gradingStatus === 'inProgress' || t.gradingStatus === 'prelimApproved' || t.gradingStatus === 'finalApproved');
  });
  var cont = document.getElementById('adminCommitteeContent');
  if(!tests.length){
    cont.innerHTML = '<div style="text-align:center;padding:60px;color:rgba(255,255,255,.3)"><div style="font-size:48px;margin-bottom:12px">📋</div><p>لا توجد اختبارات في انتظار الاعتمادية</p></div>';
    return;
  }
  var statusLabels = {draft:'مسودة',shared:'👥 مشتركة',inProgress:'✍️ جاري',approved:'✅ معتمد المحرر',prelimApproved:'🔵 مبدئي',finalApproved:'🟢 نهائي'};
  var statusColors = {draft:'rgba(255,255,255,.3)',shared:'#c4b5fd',inProgress:'#fb923c',approved:'#4ade80',prelimApproved:'#7dd3fc',finalApproved:'#FACC15'};
  cont.innerHTML = tests.map(function(t){
    var sc = statusColors[t.gradingStatus]||'rgba(255,255,255,.3)';
    var sl = statusLabels[t.gradingStatus]||'—';
    var students = t.students||[];
    var gradedCount = students.filter(function(s){ return s.gradingApproved; }).length;
    var totalScore = 0, scoredCount = 0;
    students.forEach(function(st){
      if(st.grades){
        var s = 0;
        Object.values(st.grades).forEach(function(dg){ Object.values(dg).forEach(function(v){ s+=parseFloat(v)||0; }); });
        totalScore += s; scoredCount++;
      }
    });
    var avgScore = scoredCount ? (totalScore/scoredCount).toFixed(1) : '—';

    return '<div style="background:rgba(255,255,255,.04);border:2px solid rgba(251,146,60,.2);border-radius:18px;padding:20px;margin-bottom:14px">'
      +'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap">'
      +'<div style="flex:1">'
      +'<div style="display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap">'
      +'<h4 style="font-size:17px;font-weight:800;color:white">'+(t.name||'—')+'</h4>'
      +'<span style="color:'+sc+';background:rgba(255,255,255,.05);border:1px solid '+sc+';border-radius:20px;padding:2px 12px;font-size:11px;font-weight:700">'+sl+'</span>'
      +'</div>'
      +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:8px;margin-bottom:10px">'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">المادة</div><div style="font-size:12px;font-weight:700">'+(t.subject||'—')+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">الصف</div><div style="font-size:12px;font-weight:700">'+(t.grade||'—')+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">المحرر</div><div style="font-size:12px;font-weight:700">'+(t.author||'—')+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">الطلاب</div><div style="font-size:13px;font-weight:800;color:#FACC15">'+students.length+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">تم تصحيحهم</div><div style="font-size:13px;font-weight:800;color:#4ade80">'+gradedCount+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">متوسط الدرجات</div><div style="font-size:13px;font-weight:800;color:#fb923c">'+avgScore+'%</div></div>'
      +'</div></div>'
      +'<div style="display:flex;flex-direction:column;gap:8px;min-width:160px">'
      +'<button onclick="openAdminGrading('+t.id+')" style="background:rgba(251,146,60,.2);color:#fb923c;border:1px solid rgba(251,146,60,.4);border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">✍️ فتح التصحيح</button>'
      +'<button onclick="adminPrelimApprove('+t.id+')" style="background:rgba(56,189,248,.15);color:#7dd3fc;border:1px solid rgba(56,189,248,.3);border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🔵 اعتمادية مبدئية</button>'
      +'<button onclick="adminFinalApprove('+t.id+')" style="background:rgba(250,204,21,.15);color:#FACC15;border:1px solid rgba(250,204,21,.3);border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🟢 اعتمادية نهائية</button>'
      +'</div></div></div>';
  }).join('');
}

function openAdminGrading(testId){
  // Open grading panel from admin context
  document.getElementById('adminGradingCommitteePanel').classList.add('hidden');
  document.getElementById('gradingPanel').classList.remove('hidden');
  // Move gradingPanel into adminPanel if needed
  var adminPanel = document.getElementById('adminPanel');
  var gp = document.getElementById('gradingPanel');
  if(gp.parentElement !== adminPanel) adminPanel.appendChild(gp);
  openGradingPanel('admin', testId);
}

function renderCommitteeContent(role){
  var tests = myTests.filter(function(t){
    if(t.status !== 'approved') return false;
    if(role === 'admin') return t.gradingStatus === 'approved' || t.gradingStatus === 'shared' || t.gradingStatus === 'inProgress' || t.gradingStatus === 'finalApproved';
    return t.gradingSharedWith && t.gradingSharedWith.indexOf(currentSupName) >= 0;
  });

  var cont = document.getElementById('committeeContent');
  if(!tests.length){
    cont.innerHTML = '<div style="text-align:center;padding:60px;color:rgba(255,255,255,.3)"><div style="font-size:48px;margin-bottom:12px">'+(role==='admin'?'📋':'👥')+'</div><p>لا توجد اختبارات / No tests yet</p></div>';
    return;
  }

  cont.innerHTML = tests.map(function(t){
    var statusColors = {draft:'rgba(255,255,255,.3)',shared:'#c4b5fd',inProgress:'#fb923c',approved:'#4ade80',finalApproved:'#FACC15'};
    var statusLabels = {draft:'مسودة',shared:'👥 مشتركة',inProgress:'✍️ جاري',approved:'✅ معتمد المحرر',finalApproved:'🏆 اعتماد نهائي'};
    var sc = statusColors[t.gradingStatus] || 'rgba(255,255,255,.3)';
    var sl = statusLabels[t.gradingStatus] || '—';
    var students = t.students || [];
    var gradedCount = students.filter(function(s){ return s.gradingApproved; }).length;

    var adminBtns = role === 'admin' ? ''
      +'<button onclick="adminPrelimApprove('+t.id+')" style="background:rgba(251,146,60,.2);color:#fb923c;border:1px solid rgba(251,146,60,.4);border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🔵 اعتمادية مبدئية</button>'
      +'<button onclick="adminFinalApprove('+t.id+')" style="background:rgba(250,204,21,.2);color:#FACC15;border:1px solid rgba(250,204,21,.4);border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">🟢 اعتمادية نهائية</button>'
      : '';

    return '<div style="background:rgba(255,255,255,.04);border:1px solid rgba(255,255,255,.1);border-radius:18px;padding:20px;margin-bottom:14px">'
      +'<div style="display:flex;align-items:flex-start;justify-content:space-between;gap:12px;flex-wrap:wrap">'
      +'<div style="flex:1">'
      +'<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px;flex-wrap:wrap">'
      +'<h4 style="font-weight:800;font-size:16px;color:white">'+(t.name||'—')+'</h4>'
      +'<span style="font-size:11px;color:'+sc+';background:rgba(255,255,255,.05);border:1px solid '+sc+';border-radius:20px;padding:2px 10px;">'+sl+'</span>'
      +'</div>'
      +'<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(100px,1fr));gap:8px;margin-bottom:10px">'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">المادة</div><div style="font-size:12px;font-weight:700">'+(t.subject||'—')+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">الصف</div><div style="font-size:12px;font-weight:700">'+(t.grade||'—')+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">المحرر</div><div style="font-size:12px;font-weight:700">'+(t.author||'—')+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">الطلاب</div><div style="font-size:12px;font-weight:700;color:#FACC15">'+students.length+'</div></div>'
      +'<div><div style="font-size:9px;color:rgba(255,255,255,.4)">تم تصحيحهم</div><div style="font-size:12px;font-weight:700;color:#4ade80">'+gradedCount+'</div></div>'
      +'</div>'
      +(t.gradingSharedWith&&t.gradingSharedWith.length?'<div style="font-size:11px;color:rgba(255,255,255,.4)">مشاركة مع: '+t.gradingSharedWith.join('، ')+'</div>':'')
      +'</div>'
      +'<div style="display:flex;flex-direction:column;gap:8px">'
      +'<button onclick="openGradingPanel(\''+role+'\','+t.id+')" style="background:rgba(251,146,60,.2);color:#fb923c;border:1px solid rgba(251,146,60,.4);border-radius:10px;padding:8px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:13px;font-weight:700">✍️ فتح التصحيح</button>'
      +adminBtns
      +'</div></div></div>';
  }).join('');
}

function adminPrelimApprove(testId){
  var t = myTests.find(function(x){ return x.id === testId; });
  if(!t) return;
  scConfirm('اعتمادية مبدئية','Preliminary Approval',
    'اعتماد مبدئي لتصحيح اختبار <strong style="color:#7dd3fc">'+t.name+'</strong>؟',
    'Preliminary approval for <strong>'+t.name+'</strong>?'
  ).then(function(ok){
    if(!ok) return;
    t.gradingStatus = 'prelimApproved';
    t.prelimApprovedAt = new Date().toISOString();
    saveMyTests().then(function(){
      if(_gradingRole==='admin') _renderAdminCommitteeContent();
      else renderCommitteeContent('sup');
      scOk('🔵 اعتمادية مبدئية','Preliminary Approved','تم الاعتماد المبدئي.','Preliminary approval done.','🔵');
    });
  });
}

function adminFinalApprove(testId){
  var t = myTests.find(function(x){ return x.id === testId; });
  if(!t) return;
  scDanger('🟢 اعتمادية التصحيح النهائي','Final Grading Approval',
    'اعتماد <strong>نهائي</strong> لتصحيح اختبار <strong style="color:#FACC15">'+t.name+'</strong>.<br>هذا الاعتماد يُبني عليه التقارير ولا يمكن التراجع عنه.',
    'FINAL approval for <strong>'+t.name+'</strong>. Used for reports, cannot be reversed.'
  ).then(function(ok){
    if(!ok) return;
    t.gradingStatus = 'finalApproved';
    t.finalApprovedAt = new Date().toISOString();
    saveMyTests().then(function(){
      if(_gradingRole==='admin') _renderAdminCommitteeContent();
      else renderCommitteeContent('sup');
      scOk('🟢 اعتمادية نهائية','Final Approved','تم الاعتماد النهائي — التقارير جاهزة للبناء.','Final grading approved — reports ready.','🟢');
    });
  });
}

function closeGradingCommittee(){
  document.getElementById('gradingCommitteePanel').classList.add('hidden');
  if(_gradingRole === 'admin'){
    document.getElementById('adminOptions').classList.remove('hidden');
  } else {
    document.getElementById('supMainOptions').classList.remove('hidden');
    document.getElementById('supTitle').classList.remove('hidden');
  }
}

function backToMain(){document.getElementById('mainDashboard').classList.remove('hidden-section');document.getElementById('studentPortal').classList.add('hidden');['adminPanel','supervisorPanel','schoolCoordPanel'].forEach(function(x){document.getElementById(x).classList.add('hidden');});document.getElementById('domainModal').classList.add('hidden');document.getElementById('questionModal').classList.add('hidden');document.getElementById('myTestsModal').classList.add('hidden');document.getElementById('standardsBankModal').classList.add('hidden');closeStudentWindow();currentStudentData=null;currentCoordSchool=null;if(liveRefreshTimer){clearInterval(liveRefreshTimer);liveRefreshTimer=null;}var nav=document.getElementById('topLeftNav');if(nav)nav.style.display='flex';}
function logout(){scConfirm('تسجيل الخروج','Sign Out','هل تريد تسجيل الخروج؟','Sign out?','🚪').then(function(ok){if(ok)location.reload();});}
function showAdminProfile(){
  var currentPass=localStorage.getItem('scholastic_admin_pass')||'Gems@2050';
  var html='<div style="max-width:480px;margin:0 auto">'
    +'<h2 style="font-size:22px;font-weight:800;margin-bottom:4px">👤 معلوماتي / My Profile</h2>'
    +'<p style="font-size:12px;color:#94a3b8;margin-bottom:24px;font-family:Montserrat,sans-serif">Administration Account</p>'
    +'<div style="background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:14px;padding:22px;margin-bottom:16px">'
    +'<div style="font-size:13px;font-weight:700;color:#64748b;margin-bottom:16px;font-family:Montserrat,sans-serif;letter-spacing:.5px">🔑 تغيير كلمة مرور الإدارة</div>'
    +'<div style="margin-bottom:12px">'
    +'<label style="font-size:12px;font-weight:700;color:#475569;display:block;margin-bottom:6px">كلمة المرور الحالية / Current Password</label>'
    +'<input type="password" id="ap_current" class="wizard-input" placeholder="أدخل كلمة المرور الحالية" style="margin-bottom:0"></div>'
    +'<div style="margin-bottom:12px">'
    +'<label style="font-size:12px;font-weight:700;color:#475569;display:block;margin-bottom:6px">كلمة المرور الجديدة / New Password</label>'
    +'<input type="password" id="ap_new" class="wizard-input" placeholder="أدخل كلمة المرور الجديدة" style="margin-bottom:0"></div>'
    +'<div style="margin-bottom:16px">'
    +'<label style="font-size:12px;font-weight:700;color:#475569;display:block;margin-bottom:6px">تأكيد كلمة المرور الجديدة / Confirm New Password</label>'
    +'<input type="password" id="ap_confirm" class="wizard-input" placeholder="أعد إدخال كلمة المرور الجديدة" style="margin-bottom:0"></div>'
    +'<button onclick="_saveAdminPass()" style="width:100%;background:linear-gradient(135deg,#2563EB,#1d4ed8);color:white;border:none;border-radius:12px;padding:13px;font-family:Tajawal,sans-serif;font-size:15px;font-weight:800;cursor:pointer">💾 حفظ كلمة المرور الجديدة / Save New Password</button>'
    +'</div>'
    +'<div style="background:#fffbeb;border:1.5px solid #fde68a;border-radius:12px;padding:14px;margin-bottom:16px">'
    +'<div style="font-size:12px;font-weight:700;color:#92400e">⚠️ كلمة المرور الحالية: <span style="font-family:Montserrat,sans-serif;letter-spacing:1px">'+currentPass+'</span></div>'
    +'<div style="font-size:11px;color:#b45309;margin-top:4px;font-family:Montserrat,sans-serif">تغيير كلمة المرور يؤثر على الجلسات القادمة فقط</div>'
    +'</div>'
    +'</div>';
  var ap=document.createElement('div');
  ap.id='adminProfileOverlay';
  ap.style.cssText='position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:9999;display:flex;align-items:center;justify-content:center;padding:20px';
  ap.innerHTML='<div style="background:white;border-radius:20px;padding:32px;max-width:500px;width:100%;max-height:90vh;overflow-y:auto">'+html+'<button onclick="document.getElementById(\'adminProfileOverlay\').remove()" style="width:100%;margin-top:12px;background:#f1f5f9;border:none;border-radius:12px;padding:11px;font-family:Tajawal,sans-serif;font-size:14px;font-weight:700;cursor:pointer;color:#475569">إغلاق / Close</button></div>';
  document.body.appendChild(ap);
}
function _saveAdminPass(){
  var cur=document.getElementById('ap_current').value.trim();
  var nw=document.getElementById('ap_new').value.trim();
  var conf=document.getElementById('ap_confirm').value.trim();
  var stored=localStorage.getItem('scholastic_admin_pass')||'Gems@2050';
  if(cur!==stored){scWarn('كلمة المرور الحالية غير صحيحة','Current password is incorrect');return;}
  if(!nw||nw.length<6){scWarn('كلمة المرور الجديدة يجب أن تكون 6 أحرف على الأقل','New password must be at least 6 characters');return;}
  if(nw!==conf){scWarn('كلمة المرور الجديدة وتأكيدها غير متطابقتين','New passwords do not match');return;}
  localStorage.setItem('scholastic_admin_pass',nw);
  document.getElementById('adminProfileOverlay').remove();
  scOk('تم تغيير كلمة المرور ✅','Password Changed','تم حفظ كلمة المرور الجديدة بنجاح','New password saved successfully','✅');
}
var _updateTargetId='';
function showUpdateModal(id){_updateTargetId=id;document.getElementById('upPass').value='';document.getElementById('upValue').value='';document.getElementById('upDuration').value='2';document.getElementById('upError').style.display='none';document.getElementById('updateModal').style.display='flex';setTimeout(function(){document.getElementById('upPass').focus();},100);}
function applyUpdateCounter(){var pass=document.getElementById('upPass').value,val=parseInt(document.getElementById('upValue').value),dur=parseFloat(document.getElementById('upDuration').value)||2;if(pass!=='Gems@2050'){document.getElementById('upError').style.display='block';return;}if(!val||val<0){alert('أدخل قيمة صحيحة');return;}document.getElementById('updateModal').style.display='none';animateValue(_updateTargetId,0,val,dur);}
function animateValue(id,start,end,duration){var obj=document.getElementById(id);if(!obj)return;var t0=null;var step=function(ts){if(!t0)t0=ts;var p=Math.min((ts-t0)/(duration*1000),1);obj.innerHTML=Math.floor(p*(end-start)+start).toLocaleString();if(p<1)requestAnimationFrame(step);};requestAnimationFrame(step);}

// ============================================================
// INIT
// ============================================================
// ══ Scholastic Alert System ══
var _scRes=null;
function scAlert(o){
  return new Promise(function(resolve){
    _scRes=resolve;
    document.getElementById('scIcon').textContent=o.icon||'💬';
    document.getElementById('scTAr').textContent=o.titleAr||'تنبيه';
    document.getElementById('scTEn').textContent=o.titleEn||'Notice';
    document.getElementById('scMAr').innerHTML=o.msgAr||'';
    var en=document.getElementById('scMEn');
    if(o.msgEn){en.innerHTML=o.msgEn;en.style.display='block';}
    else{en.style.display='none';}
    // Arabic shown below English always
    var ar=document.getElementById('scMAr');
    if(ar) ar.style.display=o.msgAr?'block':'none';
    var f=document.getElementById('scFoot'); f.innerHTML='';
    if(o.type==='confirm'||o.type==='danger'){
      var cb=document.createElement('button');
      cb.className='sc-btn sc-cancel';
      cb.innerHTML=(o.cancelAr||'إلغاء')+' <span style="opacity:.5;font-size:11px">/ '+(o.cancelEn||'Cancel')+'</span>';
      cb.onclick=function(){_scDone(false);};
      f.appendChild(cb);
    }
    var ob=document.createElement('button');
    ob.className='sc-btn '+(o.type==='danger'?'sc-danger-btn':'sc-ok');
    ob.innerHTML=(o.confirmAr||'موافق')+' <span style="opacity:.7;font-size:11px">/ '+(o.confirmEn||'OK')+'</span>';
    ob.onclick=function(){_scDone(true);};
    f.appendChild(ob);
    document.getElementById('scAlertOverlay').style.display='flex';
  });
}
function scPromptText(tAr,tEn,placeholder,icon){
  return new Promise(function(resolve){
    _scRes=resolve;
    document.getElementById('scIcon').textContent=icon||'✏️';
    document.getElementById('scTAr').textContent=tAr||'إدخال';
    document.getElementById('scTEn').textContent=tEn||'Input';
    document.getElementById('scMAr').innerHTML='';
    document.getElementById('scMAr').style.display='none';
    document.getElementById('scMEn').style.display='none';
    var input=document.getElementById('scPromptInput');
    input.style.display='block';input.value='';input.placeholder=placeholder||'';
    var f=document.getElementById('scFoot');f.innerHTML='';
    var cb=document.createElement('button');
    cb.className='sc-btn sc-cancel';
    cb.innerHTML='إلغاء <span style="opacity:.5;font-size:11px">/ Cancel</span>';
    cb.onclick=function(){input.style.display='none';_scPromptDone(null);};
    f.appendChild(cb);
    var ob=document.createElement('button');
    ob.className='sc-btn sc-ok';
    ob.innerHTML='تأكيد <span style="opacity:.7;font-size:11px">/ Confirm</span>';
    ob.onclick=function(){var v=input.value.trim();input.style.display='none';_scPromptDone(v||null);};
    f.appendChild(ob);
    document.getElementById('scAlertOverlay').style.display='flex';
    setTimeout(function(){input.focus();},100);
  });
}
function _scPromptDone(v){document.getElementById('scAlertOverlay').style.display='none';if(_scRes){_scRes(v);_scRes=null;}}
function _scDone(v){document.getElementById('scAlertOverlay').style.display='none';if(_scRes){_scRes(v);_scRes=null;}}
function scOk(tAr,tEn,mAr,mEn,icon){return scAlert({icon:icon||'✅',titleAr:tAr,titleEn:tEn,msgAr:mAr,msgEn:mEn,type:'alert',confirmAr:'موافق',confirmEn:'OK'});}
function scConfirm(tAr,tEn,mAr,mEn,icon){return scAlert({icon:icon||'❓',titleAr:tAr,titleEn:tEn,msgAr:mAr,msgEn:mEn,type:'confirm'});}
function scDanger(tAr,tEn,mAr,mEn){return scAlert({icon:'⚠️',titleAr:tAr,titleEn:tEn,msgAr:mAr,msgEn:mEn,type:'danger',confirmAr:'نعم، متأكد',confirmEn:'Yes, confirm'});}
function scWarn(mAr,mEn){return scAlert({icon:'⚠️',titleAr:'تحذير',titleEn:'Warning',msgAr:mAr,msgEn:mEn,type:'alert',confirmAr:'فهمت',confirmEn:'Understood'});}

// ══ Instructions Checkbox & Mic ══
function _swUpdateStartBtn(){
  var cb=document.getElementById('sw-confirm-check');
  var btn=document.getElementById('sw-start-btn');
  if(!btn) return;
  var micOk=!document.getElementById('sw-mic-check')||window._swMicGranted;
  var checked=cb&&cb.checked;
  var allOk=checked&&micOk;
  btn.disabled=!allOk;
  btn.style.background=allOk?'linear-gradient(135deg,#22c55e,#15803d)':'linear-gradient(135deg,#94a3b8,#64748b)';
  btn.style.borderColor=allOk?'#86efac':'#cbd5e1';
  btn.style.cursor=allOk?'pointer':'not-allowed';
  btn.style.opacity=allOk?'1':'.6';
  var hint=document.getElementById('sw-start-hint');
  if(hint) hint.style.display=allOk?'none':'block';
}
function swCheckboxChanged(){ _swUpdateStartBtn(); }
window._swMicGranted=false;
var _swMicStream=null, _swMicRecorder=null, _swMicChunks=[];

function swCheckMic(){
  var btn=document.getElementById('sw-mic-btn');
  var st=document.getElementById('sw-mic-status');
  var mc=document.getElementById('sw-mic-check');
  if(btn){btn.textContent='⏳ جاري الاتصال...';btn.disabled=true;}
  if(!navigator.mediaDevices||!navigator.mediaDevices.getUserMedia){
    _swMicFailed('المتصفح لا يدعم الميكروفون','Browser does not support microphone');
    return;
  }
  navigator.mediaDevices.getUserMedia({audio:true}).then(function(stream){
    _swMicStream=stream;
    // أظهر واجهة تجريب التسجيل
    if(mc) mc.innerHTML=
      '<span style="font-size:20px">🎙</span>'
      +'<div style="flex:1">'
        +'<div style="font-weight:800;color:#065f46;font-size:14px">الميكروفون متصل — سجّل جملة قصيرة للتأكد</div>'
        +'<div style="font-size:12px;color:#047857;font-family:Montserrat,sans-serif;direction:ltr;text-align:left">Mic connected — record a short phrase to verify</div>'
      +'</div>'
      +'<div style="display:flex;flex-direction:column;gap:6px;align-items:flex-end">'
        +'<button id="sw-mic-rec-btn" onclick="swMicTestRecord()" style="background:linear-gradient(135deg,#dc2626,#b91c1c);color:white;border:none;border-radius:10px;padding:8px 16px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🔴 ابدأ التسجيل / Record</button>'
        +'<audio id="sw-mic-playback" controls style="display:none;width:200px;height:32px;border-radius:8px" dir="ltr"></audio>'
        +'<button id="sw-mic-confirm-btn" onclick="swMicConfirm()" style="display:none;background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:none;border-radius:10px;padding:8px 16px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">✅ يعمل — تأكيد / Confirm</button>'
      +'</div>';
  }).catch(function(err){
    if(err.name==='NotAllowedError'||err.name==='PermissionDeniedError'){
      _swMicFailed('تم رفض إذن الميكروفون','Microphone permission denied');
      _swShowMicSettings();
    } else {
      _swMicFailed('لا يوجد ميكروفون متصل','No microphone found');
    }
  });
}

function swMicTestRecord(){
  var recBtn=document.getElementById('sw-mic-rec-btn');
  if(!_swMicStream) return;
  if(_swMicRecorder&&_swMicRecorder.state==='recording'){
    _swMicRecorder.stop();
    if(recBtn){recBtn.textContent='🔴 سجّل مجدداً / Re-record';recBtn.style.background='linear-gradient(135deg,#dc2626,#b91c1c)';}
  } else {
    _swMicChunks=[];
    _swMicRecorder=new MediaRecorder(_swMicStream);
    _swMicRecorder.ondataavailable=function(e){_swMicChunks.push(e.data);};
    _swMicRecorder.onstop=function(){
      var blob=new Blob(_swMicChunks,{type:'audio/webm'});
      var url=URL.createObjectURL(blob);
      var aud=document.getElementById('sw-mic-playback');
      var confBtn=document.getElementById('sw-mic-confirm-btn');
      if(aud){aud.src=url;aud.style.display='block';aud.play();}
      if(confBtn) confBtn.style.display='block';
    };
    _swMicRecorder.start();
    if(recBtn){recBtn.textContent='⏹ إيقاف / Stop';recBtn.style.background='linear-gradient(135deg,#f59e0b,#d97706)';}
    setTimeout(function(){if(_swMicRecorder&&_swMicRecorder.state==='recording')_swMicRecorder.stop();},8000);
  }
}

function swMicConfirm(){
  if(_swMicStream) _swMicStream.getTracks().forEach(function(t){t.stop();});
  window._swMicGranted=true;
  var mc=document.getElementById('sw-mic-check');
  if(mc) mc.innerHTML=
    '<span style="font-size:24px">✅</span>'
    +'<div style="flex:1;font-weight:800;color:#065f46;font-size:14px">الميكروفون يعمل بشكل صحيح / Microphone is working</div>';
  mc&&(mc.style.background='rgba(34,197,94,.1)')&&(mc.style.borderColor='#22c55e');
  _swUpdateStartBtn();
}

function _swMicFailed(ar,en){
  var mc=document.getElementById('sw-mic-check');
  if(mc) mc.innerHTML=
    '<span style="font-size:20px">⚠️</span>'
    +'<div style="flex:1"><div style="font-weight:800;color:#991b1b;font-size:14px">'+ar+'</div>'
    +'<div style="font-size:12px;color:#b91c1c;font-family:Montserrat,sans-serif;direction:ltr;text-align:left">'+en+'</div></div>'
    +'<div style="display:flex;gap:8px">'
      +'<button onclick="swCheckMic()" style="background:linear-gradient(135deg,#f59e0b,#d97706);color:white;border:none;border-radius:10px;padding:8px 14px;font-size:12px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🔄 إعادة المحاولة</button>'
      +'<button onclick="_swShowMicSettings()" style="background:rgba(255,255,255,.2);color:#92400e;border:1px solid #d97706;border-radius:10px;padding:8px 14px;font-size:12px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">⚙️ الإعدادات</button>'
    +'</div>';
}

function _swShowMicSettings(){
  scAlert({
    icon:'🎙',
    titleAr:'كيفية تفعيل الميكروفون',
    titleEn:'How to Enable Microphone',
    msgAr:'<strong>Chrome/Edge:</strong> اضغط على 🔒 في شريط العنوان ← الميكروفون ← السماح<br><br>'
         +'<strong>Firefox:</strong> اضغط على 🔒 ← الأذونات ← الميكروفون ← السماح<br><br>'
         +'<strong>Safari:</strong> الإعدادات ← مواقع الويب ← الميكروفون ← السماح<br><br>'
         +'بعد السماح اضغط <strong>إعادة المحاولة</strong>',
    msgEn:'Click the 🔒 lock icon in the address bar → Microphone → Allow, then click Retry.',
    type:'alert',
    confirmAr:'فهمت',
    confirmEn:'Got it'
  });
}

// ══ Full Lock: منع الخروج بعد بدء الاختبار ══
var _swLocked=false;
var _swExitAttempts=[];
function swEnableFullLock(){
  _swLocked=true;
  window._swBeforeUnload=function(e){
    if(!_swLocked) return;
    _swLogExit('beforeunload');
    e.preventDefault();
    e.returnValue='لا يمكنك مغادرة الاختبار / You cannot leave the test.';
    return e.returnValue;
  };
  window.addEventListener('beforeunload',window._swBeforeUnload);
  // منع back button
  history.pushState(null,null,location.href);
  window._swPopState=function(){
    if(!_swLocked) return;
    history.pushState(null,null,location.href);
    _swLogExit('back button');
    scWarn('لا يمكنك الرجوع أثناء الاختبار','You cannot go back during the test');
  };
  window.addEventListener('popstate',window._swPopState);
}
function swDisableFullLock(){
  _swLocked=false;
  if(window._swBeforeUnload) window.removeEventListener('beforeunload',window._swBeforeUnload);
  if(window._swPopState) window.removeEventListener('popstate',window._swPopState);
}
function _swLogExit(reason){
  var entry={time:new Date().toLocaleTimeString('ar'),reason:reason,student:currentStudentData?((currentStudentData.student.firstName||'')+(currentStudentData.student.lastName||'')):''};
  _swExitAttempts.push(entry);
  // Save to test record
  if(currentStudentData&&currentStudentData.test){
    if(!currentStudentData.student.exitAttempts) currentStudentData.student.exitAttempts=[];
    currentStudentData.student.exitAttempts.push(entry);
    saveMyTests();
  }
  triggerCheatWarning('محاولة خروج: '+reason);
}

// ══ 100% Completion Check ══
function swCheck80Percent(){
  var d=testData.domains[sw_domainIdx];
  var questions=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[]):(d.questions||[]);
  var total=questions.length;
  if(total===0) return true;
  var answered=Object.keys(sw_answers).filter(function(k){return sw_answers[k]!==undefined&&sw_answers[k]!==null;}).length;
  var remaining=total-answered;
  if(remaining>0){
    scAlert({
      icon:'📝',
      titleAr:'أسئلة غير مكتملة',
      titleEn:'Unanswered Questions',
      msgAr:'لا يزال هناك <strong style="color:#f87171;font-size:18px">'+remaining+'</strong> سؤال لم تجب عليه.<br>يجب الإجابة على <strong>جميع الأسئلة</strong> قبل التسليم.',
      msgEn:'You still have <strong style="color:#f87171">'+remaining+'</strong> unanswered question(s).<br>You must answer <strong>all questions</strong> before submitting.',
      type:'alert',
      confirmAr:'حسناً، سأكمل',
      confirmEn:'OK, I will complete'
    });
    return false;
  }
  return true;
}

// ══ Eye animation + close guard ══
function _swStartEyeAnim(){
  clearInterval(window._swEyeInterval);
  window._swEyeInterval=setInterval(function(){
    var e=document.getElementById('sw-eye-icon');
    if(!e) return;
    e.textContent='🌕';
    setTimeout(function(){if(e)e.textContent='👁';},200);
  },2000);
}
function swRequestClose(){
  if(_tryModeActive){closeStudentWindow();return;}
  if(_swLocked){
    _swLogExit('close button attempt');
    scAlert({icon:'👁',titleAr:'⚠️ تحذير — أنت تحت المراقبة',titleEn:'⚠️ Warning — You Are Being Monitored',msgEn:'👁 <strong>MONITORED</strong> — You cannot close the test before submitting <strong>all domains</strong>.<br>This exit attempt has been recorded.',msgAr:'👁 <strong>أنت تحت المراقبة</strong> — لا يمكنك الإغلاق قبل تسليم جميع المجالات.<br>تم تسجيل محاولة الخروج هذه.',type:'alert',confirmAr:'فهمت — سأكمل الاختبار',confirmEn:'Understood — I will complete the test'});
  } else {
    closeStudentWindow();
  }
}

// ══ True/False helpers ══
function addTFStatement(text,answer){
  var cont=document.getElementById('tf-statements'); if(!cont) return;
  var idx=cont.children.length;
  var div=document.createElement('div');
  div.style.cssText='display:flex;align-items:center;gap:10px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.15);border-radius:12px;padding:12px;margin-bottom:8px';
  div.innerHTML='<div class="domain-badge" style="font-size:11px;flex-shrink:0">'+(idx+1)+'</div>'
    +'<input type="text" id="tf-text-'+idx+'" class="wizard-input" placeholder="اكتب الجملة..." value="'+(text||'')+'" dir="auto" style="flex:1;font-size:13px">'
    +'<label style="display:flex;align-items:center;gap:5px;cursor:pointer;font-size:13px;white-space:nowrap"><input type="radio" name="tf-ans-'+idx+'" value="true" '+(answer==='true'?'checked':'')+' style="accent-color:#22c55e"> ✅ صواب</label>'
    +'<label style="display:flex;align-items:center;gap:5px;cursor:pointer;font-size:13px;white-space:nowrap"><input type="radio" name="tf-ans-'+idx+'" value="false" '+(answer==='false'?'checked':'')+' style="accent-color:#ef4444"> ❌ خطأ</label>'
    +'<button onclick="this.parentElement.remove()" style="color:#f87171;background:none;border:none;cursor:pointer;font-size:16px;flex-shrink:0">🗑</button>';
  cont.appendChild(div);
}

// ══ Classify helpers ══
function buildClassifyCols(){
  var n=parseInt(document.getElementById('classify-cols').value)||2;
  var cont=document.getElementById('classify-cols-container'); if(!cont) return;
  var existing=[];
  for(var i=0;i<10;i++){var el=document.getElementById('classify-col-'+i);if(el)existing.push(el.value);else break;}
  cont.innerHTML='<div style="display:grid;grid-template-columns:repeat('+n+',1fr);gap:8px;margin-bottom:8px">'
    +Array.from({length:n},function(_,i){return '<input type="text" id="classify-col-'+i+'" class="wizard-input" placeholder="عنوان العمود '+(i+1)+'" value="'+(existing[i]||'')+'" style="font-size:12px;text-align:center;background:rgba(250,204,21,.1);border-color:rgba(250,204,21,.4);color:#FACC15;font-weight:700">';}).join('')
    +'</div>';
}
function addClassifyItem(text,col){
  var cont=document.getElementById('classify-items-container'); if(!cont) return;
  var nCols=parseInt(document.getElementById('classify-cols').value)||2;
  var idx=cont.children.length;
  var div=document.createElement('div');
  div.style.cssText='display:flex;align-items:center;gap:8px;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:10px;margin-bottom:6px';
  div.innerHTML='<div class="domain-badge" style="font-size:10px;flex-shrink:0">'+(idx+1)+'</div>'
    +'<input type="text" id="cl-item-'+idx+'" class="wizard-input" placeholder="الكلمة أو الجملة..." value="'+(text||'')+'" dir="auto" style="flex:1;font-size:12px">'
    +'<select id="cl-col-'+idx+'" class="wizard-input" style="width:120px;font-size:12px">'
    +Array.from({length:nCols},function(_,i){var el=document.getElementById('classify-col-'+i);var lbl=el?el.value:('عمود '+(i+1));return '<option value="'+i+'" '+(col==i?'selected':'')+'>'+lbl+'</option>';}).join('')
    +'</select>'
    +'<button onclick="this.parentElement.remove()" style="color:#f87171;background:none;border:none;cursor:pointer;font-size:14px">🗑</button>';
  cont.appendChild(div);
}

// ══ Save new types in saveQuestion ══

function hexToRgb(hex){
  hex=hex.replace('#','');
  if(hex.length===3) hex=hex.split('').map(function(c){return c+c;}).join('');
  var r=parseInt(hex.substring(0,2),16);
  var g=parseInt(hex.substring(2,4),16);
  var b=parseInt(hex.substring(4,6),16);
  return r+','+g+','+b;
}
function swTFSelect(qi,si,val){
  if(!sw_answers[qi]) sw_answers[qi]={};
  sw_answers[qi][si]=val;
  // تحقق إذا تمت الإجابة على كل الجمل
  var d=testData.domains[sw_domainIdx];
  var q=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions[qi]:null):(d.questions?d.questions[qi]:null);
  if(q&&q.statements){
    var allDone=q.statements.every(function(_,i){return sw_answers[qi]&&sw_answers[qi][i]!==undefined;});
    if(allDone) updateProgressDots();
  }
  renderStudentQuestion();
}
function swTFStreamSelect(qi,si,val){
  if(!sw_answers[qi])sw_answers[qi]={};
  sw_answers[qi][si]=val;
  streamMarkAnswered(qi);
  renderStudentQuestion();
}
function clSwCellReturn(qi,ci,ri){
  var ans=sw_answers[qi];if(!ans)return;
  var w=ans.placed[ci]&&ans.placed[ci][ri];if(!w)return;
  ans.placed[ci][ri]=null;
  ans.bank.push(w);
  Object.keys(ans.placed).forEach(function(k){ans.placed[k]=ans.placed[k].filter(function(x){return x!==null;});});
  renderStudentQuestion();
}
// Stream Ordering drag/drop
var _sorDrag={qi:-1,gi:-1,from:null,idx:-1,word:null};
function sorBankDragStart(e,qi,gi,wi){
  _sorDrag={qi:qi,gi:gi,from:'bank',idx:wi,word:sw_answers[qi].groups[gi].bank[wi]};
  e.dataTransfer.effectAllowed='move';
}
function sorSlotDragOver(e){e.preventDefault();e.currentTarget.style.background='#dbeafe';e.currentTarget.style.borderColor='#3b82f6';}
function sorSlotDragLeave(e,qi,gi,si){e.currentTarget.style.background='white';e.currentTarget.style.borderColor='#cbd5e1';}
function sorSlotDrop(e,qi,gi,si){
  e.preventDefault();
  if(_sorDrag.qi!==qi||_sorDrag.gi!==gi||!_sorDrag.word)return;
  var g=sw_answers[qi].groups[gi];
  var displaced=g.placed[si]||null;
  g.placed[si]=_sorDrag.word;
  if(_sorDrag.from==='bank'){
    g.bank=g.bank.filter(function(_,i){return i!==_sorDrag.idx;});
    if(displaced)g.bank.push(displaced);
  }
  _sorDrag={qi:-1,gi:-1,from:null,idx:-1,word:null};
  streamMarkAnswered(qi);renderStudentQuestion();
}
function sorReturn(qi,gi,si){
  var g=sw_answers[qi].groups[gi];if(!g||!g.placed[si])return;
  g.bank.push(g.placed[si]);g.placed[si]=null;
  renderStudentQuestion();
}
function sorClearAll(qi){
  var d=testData.domains[sw_domainIdx];
  var q=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions[qi]:null):(d.questions?d.questions[qi]:null);
  if(!q) return;
  var groups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words}]:[]);
  if(!groups.length) return;
  sw_answers[qi]={groups:groups.map(function(g){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});return {placed:new Array((g.words||[]).length).fill(null),bank:bank};})};
  updateProgressDots();renderStudentQuestion();
}

// ══ Classify student ══
var _clSelected={qi:null,word:null};
function clSelectIdx(qi,wi){
  if(!sw_answers[qi]) return;
  var word=sw_answers[qi].bank[wi];
  clSelectItem(qi,word);
}
function clConfirm(qi){
  sw_answers[qi].confirmed=true;
  streamMarkAnswered&&streamMarkAnswered(qi);
  updateProgressDots();
  scOk('تم التأكيد ✅','Confirmed','تم حفظ تصنيفاتك بنجاح.','Your classifications have been saved.','✅');
}
function clClearAll(qi){
  if(!sw_answers[qi])return;
  var allItems=[];
  Object.values(sw_answers[qi].placed||{}).forEach(function(arr){allItems=allItems.concat(arr);});
  allItems=allItems.concat(sw_answers[qi].bank||[]);
  sw_answers[qi]={placed:{},bank:allItems,confirmed:false};
  delete sw_answers[qi].confirmed;
  updateProgressDots();
  renderStudentQuestion();
}
function clSelectItem(qi,word){
  _clSelected={qi:qi,word:word};
  // highlight
  document.querySelectorAll('[id^="cl-bank-item-'+qi+'-"]').forEach(function(el){
    el.style.background=el.textContent.trim()===word?'#1e3a8a':'white';
    el.style.color=el.textContent.trim()===word?'white':'#1e3a8a';
  });
}
function clPlaceInCol(qi,ci){
  if(!_clSelected.word||_clSelected.qi!==qi) return;
  if(!sw_answers[qi]) return;
  var w=_clSelected.word;
  // remove from bank
  var bi=sw_answers[qi].bank.indexOf(w);
  if(bi>=0) sw_answers[qi].bank.splice(bi,1);
  // add to column
  if(!sw_answers[qi].placed[ci]) sw_answers[qi].placed[ci]=[];
  sw_answers[qi].placed[ci].push(w);
  _clSelected={qi:null,word:null};
  updateProgressDots();
  renderStudentQuestion();
}
function clReturn(qi,ci,wi){
  if(!sw_answers[qi]) return;
  var w=sw_answers[qi].placed[ci][wi];
  sw_answers[qi].placed[ci].splice(wi,1);
  sw_answers[qi].bank.push(w);
  updateProgressDots();
  renderStudentQuestion();
}

// ══ Listening Record helpers ══
var _listenStream=null,_listenRecorder=null,_listenChunks=[];
function listenRecordToggle(){
  var btn=document.getElementById('listen-rec-btn');
  var st=document.getElementById('listen-rec-status');
  if(!_listenRecorder||_listenRecorder.state==='inactive'){
    navigator.mediaDevices&&navigator.mediaDevices.getUserMedia({audio:true}).then(function(stream){
      _listenStream=stream;_listenChunks=[];
      _listenRecorder=new MediaRecorder(stream);
      _listenRecorder.ondataavailable=function(e){_listenChunks.push(e.data);};
      _listenRecorder.onstop=function(){
        var blob=new Blob(_listenChunks,{type:'audio/webm'});
        var url=URL.createObjectURL(blob);
        var aud=document.getElementById('listen-audio-preview');
        if(aud){aud.src=url;aud.style.display='block';}
        if(btn){btn.innerHTML='🔴 تسجيل جديد / Re-record';btn.style.background='linear-gradient(135deg,#dc2626,#b91c1c)';}
      };
      _listenRecorder.start();
      if(btn){btn.innerHTML='⏹ إيقاف / Stop';btn.style.background='linear-gradient(135deg,#f59e0b,#d97706)';}
      var sec=0;
      window._listenRecInt=setInterval(function(){sec++;if(st)st.textContent='⏱ '+sec+'s recording...';},1000);
    }).catch(function(){if(st)st.textContent='⚠️ لا يمكن الوصول للميكروفون';});
  } else {
    clearInterval(window._listenRecInt);
    _listenRecorder.stop();
    if(_listenStream) _listenStream.getTracks().forEach(function(t){t.stop();});
    if(st)st.textContent='✅ تم التسجيل';
  }
}
function listenUploadAudio(event){
  var file=event.target.files?event.target.files[0]:null;if(!file)return;
  var reader=new FileReader();
  reader.onload=function(ev){
    var aud=document.getElementById('listen-audio-preview');
    if(aud){aud.src=ev.target.result;aud.style.display='block';}
  };
  reader.readAsDataURL(file);
}

// ══ Matching Pair Editor ══
// يجمع أزواج التوصيل من الصفحة بشكل موثوق حتى لو حُذف زوج من المنتصف (بدل الاعتماد على تسلسل IDs بدون فجوات)
function _collectMatchPairsFromDOM(){
  var aEls=document.querySelectorAll('[id^="mAEdit"]');
  var pairs=[];
  Array.prototype.forEach.call(aEls,function(aEl){
    var idx=aEl.id.replace('mAEdit','');
    var bEl=document.getElementById('mBEdit'+idx);
    pairs.push({aHtml:aEl.innerHTML||'',bHtml:bEl?bEl.innerHTML||'':''});
  });
  return pairs;
}
function addMatchPairEditor(p,pi){
  var cont=document.getElementById('matchPairsEditor');if(!cont)return;
  var idx=(pi!==undefined)?pi:cont.children.length;
  p=p||{aHtml:'',aImg:'',bHtml:'',bImg:''};
  var div=document.createElement('div');
  div.id='matchPairRow'+idx;
  div.style.cssText='display:grid;grid-template-columns:1fr 48px 1fr;gap:10px;align-items:start;background:rgba(255,255,255,.05);border:1px solid rgba(255,255,255,.12);border-radius:14px;padding:14px;margin-bottom:10px;position:relative';
  var aBox='<div id="mAEdit'+idx+'" contenteditable="true" dir="auto" class="rich-editor" style="min-height:50px" placeholder="نص A..."></div>';
  var bBox='<div id="mBEdit'+idx+'" contenteditable="true" dir="auto" class="rich-editor" style="min-height:50px" placeholder="نص B..."></div>';
  div.innerHTML=
    // عمود A
    '<div>'
    +'<div style="font-size:10px;color:#93c5fd;font-family:Montserrat,sans-serif;font-weight:700;margin-bottom:6px;text-align:center">◀ A '+(idx+1)+'</div>'
    +aBox
    // صورة A
    +'<div id="mAImgWrap'+idx+'" style="margin-top:6px">'
    +'<div id="mAImg'+idx+'" style="display:none;margin-bottom:4px;border-radius:8px;overflow:hidden;border:1px solid rgba(255,255,255,.2);height:80px;display:flex;align-items:center;justify-content:center;background:#0a0a1a"><img src="" style="max-width:100%;max-height:100%;object-fit:contain;display:block;transform:scale(1)"></div>'
    +'<div id="mAImgZoom'+idx+'" style="display:none;align-items:center;gap:4px;margin-bottom:4px"><span style="font-size:9px;color:#93c5fd">🔍</span><input type="range" min="0.3" max="3" step="0.05" value="1" oninput="matchPairImgZoom(\'mAImg'+idx+'\',this.value)" style="flex:1;accent-color:#FACC15;height:14px"><button onclick="matchPairImgConfirm(\'mAImg'+idx+'\')" style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);color:#4ade80;border-radius:6px;padding:2px 7px;font-size:9px;cursor:pointer">✅</button></div>'
    +'<div style="display:flex;gap:4px;margin-top:4px">'
    +'<label style="flex:1;display:inline-flex;align-items:center;justify-content:center;gap:4px;background:rgba(59,130,246,.15);border:1px solid rgba(59,130,246,.3);border-radius:8px;padding:5px;font-size:10px;cursor:pointer;color:#93c5fd">📁<input type="file" accept="image/*" style="display:none" data-target="mAImg'+idx+'" onchange="matchPairImg(event,\'mAImg'+idx+'\')"></label>'
    +'<button onclick="matchPairImgURL(\'mAImg'+idx+'\')" style="flex:1;background:rgba(59,130,246,.15);border:1px solid rgba(59,130,246,.3);border-radius:8px;padding:5px;font-size:10px;cursor:pointer;color:#93c5fd">🔗 URL</button>'
    +'</div></div></div>'
    // الوسط =
    +'<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;padding-top:32px;gap:4px">'
    +'<div style="width:36px;height:36px;background:rgba(250,204,21,.15);border:2px solid rgba(250,204,21,.4);border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:18px;font-weight:900;color:#FACC15">=</div>'
    +'<button onclick="var d=document.getElementById(\'matchPairRow'+idx+'\');if(d)d.remove();" style="background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);color:#f87171;border-radius:6px;width:28px;height:28px;cursor:pointer;font-size:12px">✕</button>'
    +'</div>'
    // عمود B
    +'<div>'
    +'<div style="font-size:10px;color:#86efac;font-family:Montserrat,sans-serif;font-weight:700;margin-bottom:6px;text-align:center">B '+(idx+1)+' ▶</div>'
    +bBox
    // صورة B
    +'<div id="mBImgWrap'+idx+'" style="margin-top:6px">'
    +'<div id="mBImg'+idx+'" style="display:none;margin-bottom:4px;border-radius:8px;overflow:hidden;border:1px solid rgba(255,255,255,.2);height:80px;display:flex;align-items:center;justify-content:center;background:#0a0a1a"><img src="" style="max-width:100%;max-height:100%;object-fit:contain;display:block;transform:scale(1)"></div>'
    +'<div id="mBImgZoom'+idx+'" style="display:none;align-items:center;gap:4px;margin-bottom:4px"><span style="font-size:9px;color:#86efac">🔍</span><input type="range" min="0.3" max="3" step="0.05" value="1" oninput="matchPairImgZoom(\'mBImg'+idx+'\',this.value)" style="flex:1;accent-color:#FACC15;height:14px"><button onclick="matchPairImgConfirm(\'mBImg'+idx+'\')" style="background:rgba(74,222,128,.2);border:1px solid rgba(74,222,128,.4);color:#4ade80;border-radius:6px;padding:2px 7px;font-size:9px;cursor:pointer">✅</button></div>'
    +'<div style="display:flex;gap:4px;margin-top:4px">'
    +'<label style="flex:1;display:inline-flex;align-items:center;justify-content:center;gap:4px;background:rgba(134,239,172,.1);border:1px solid rgba(134,239,172,.3);border-radius:8px;padding:5px;font-size:10px;cursor:pointer;color:#86efac">📁<input type="file" accept="image/*" style="display:none" onchange="matchPairImg(event,\'mBImg'+idx+'\')"></label>'
    +'<button onclick="matchPairImgURL(\'mBImg'+idx+'\')" style="flex:1;background:rgba(134,239,172,.1);border:1px solid rgba(134,239,172,.3);border-radius:8px;padding:5px;font-size:10px;cursor:pointer;color:#86efac">🔗 URL</button>'
    +'</div></div></div>';
  cont.appendChild(div);
  // Load existing data
  if(p.aHtml) setTimeout(function(){var el=document.getElementById('mAEdit'+idx);if(el)el.innerHTML=p.aHtml;},50);
  if(p.bHtml) setTimeout(function(){var el=document.getElementById('mBEdit'+idx);if(el)el.innerHTML=p.bHtml;},50);
  if(p.aImg){var wA=document.getElementById('mAImg'+idx);if(wA){wA.style.display='flex';var imgA=wA.querySelector('img');if(imgA){imgA.src=p.aImg;imgA.style.transform='scale('+(p.aImgScale||1)+')';imgA.dataset.scale=p.aImgScale||1;}}}
  if(p.bImg){var wB=document.getElementById('mBImg'+idx);if(wB){wB.style.display='flex';var imgB=wB.querySelector('img');if(imgB){imgB.src=p.bImg;imgB.style.transform='scale('+(p.bImgScale||1)+')';imgB.dataset.scale=p.bImgScale||1;}}}
}
function matchPairImgURL(wrapId){
  scPromptText('رابط الصورة','Image URL','https://example.com/image.png','🌐').then(function(url){
    if(!url)return;
    var wrap=document.getElementById(wrapId);if(!wrap)return;
    wrap.style.display='flex';
    var img=wrap.querySelector('img');if(img){img.src=url;img.style.transform='scale(1)';img.dataset.scale='1';}
    var zoomEl=document.getElementById(wrapId+'Zoom');if(zoomEl){zoomEl.style.display='flex';var sl=zoomEl.querySelector('input[type="range"]');if(sl)sl.value=1;}
  });
}
function matchPairImg(event,wrapId){
  var file=event.target.files?event.target.files[0]:null;if(!file)return;
  var reader=new FileReader();
  reader.onload=function(ev){
    var wrap=document.getElementById(wrapId);
    if(wrap){wrap.style.display='flex';var img=wrap.querySelector('img');if(img){img.src=ev.target.result;img.style.transform='scale(1)';img.dataset.scale='1';}}
    var zoomEl=document.getElementById(wrapId+'Zoom');if(zoomEl){zoomEl.style.display='flex';var sl=zoomEl.querySelector('input[type="range"]');if(sl)sl.value=1;}
  };
  reader.readAsDataURL(file);
}
function matchPairImgZoom(wrapId,val){
  var wrap=document.getElementById(wrapId);if(!wrap)return;
  var img=wrap.querySelector('img');if(img){img.style.transform='scale('+val+')';img.dataset.scale=val;}
}
function matchPairImgConfirm(wrapId){
  var wrap=document.getElementById(wrapId);if(!wrap)return;
  var zoomEl=document.getElementById(wrapId+'Zoom');if(zoomEl)zoomEl.style.display='none';
  scOk('تم التثبيت ✅','Confirmed','تم تثبيت حجم الصورة كما سيظهر للطالب','Image size locked as it will appear to students','✅');
}

// ══ Match Answer Key Rows ══
function addMatchAnswerKeyRow(leftVal,rightVal){
  var cont=document.getElementById('matchAnswerKeyRows');if(!cont)return;
  var idx=cont.children.length;
  var row=document.createElement('div');
  row.style.cssText='display:grid;grid-template-columns:1fr 48px 1fr 32px;gap:8px;align-items:center;margin-bottom:8px';
  var leftIn=document.createElement('input');
  leftIn.type='text';leftIn.placeholder='الطرف الأيسر / Left '+( idx+1);
  leftIn.className='wizard-input match-ak-left';leftIn.style.cssText='font-size:13px;text-align:center';
  leftIn.dir='auto';if(leftVal)leftIn.value=leftVal;
  var eqDiv=document.createElement('div');
  eqDiv.style.cssText='display:flex;align-items:center;justify-content:center;background:rgba(250,204,21,.15);border:2px solid rgba(250,204,21,.4);border-radius:10px;height:40px;font-size:16px;font-weight:900;color:#FACC15';
  eqDiv.textContent='=';
  var rightIn=document.createElement('input');
  rightIn.type='text';rightIn.placeholder='الطرف الأيمن / Right '+( idx+1);
  rightIn.className='wizard-input match-ak-right';rightIn.style.cssText='font-size:13px;text-align:center';
  rightIn.dir='auto';if(rightVal)rightIn.value=rightVal;
  var delBtn=document.createElement('button');
  delBtn.style.cssText='background:rgba(239,68,68,.15);border:1px solid rgba(239,68,68,.3);color:#f87171;border-radius:8px;height:36px;cursor:pointer;font-size:14px';
  delBtn.textContent='✕';delBtn.onclick=function(){row.remove();_syncMatchAnswerKey();};
  leftIn.oninput=rightIn.oninput=_syncMatchAnswerKey;
  row.appendChild(leftIn);row.appendChild(eqDiv);row.appendChild(rightIn);row.appendChild(delBtn);
  cont.appendChild(row);
}
function _syncMatchAnswerKey(){
  var cont=document.getElementById('matchAnswerKeyRows');
  var hidden=document.getElementById('matchAnswerKey');if(!hidden)return;
  var pairs=[];
  (cont?cont.querySelectorAll('div'):[]).forEach(function(row){
    var l=row.querySelector('.match-ak-left');var r=row.querySelector('.match-ak-right');
    if(l&&r&&(l.value.trim()||r.value.trim()))pairs.push(l.value.trim()+'='+r.value.trim());
  });
  hidden.value=pairs.join(', ');
}
// تهيئة صفوف نموذج الإجابة عند الفتح
function initMatchAnswerKeyRows(){
  var cont=document.getElementById('matchAnswerKeyRows');if(!cont)return;
  cont.innerHTML='';
  var hidden=document.getElementById('matchAnswerKey');
  var existing=hidden&&hidden.value?hidden.value:'';
  if(existing){
    existing.split(',').forEach(function(pair){
      var parts=pair.split('=');
      addMatchAnswerKeyRow((parts[0]||'').trim(),(parts[1]||'').trim());
    });
  }
  // أضف صفوف فارغة حسب عدد الأزواج المدخلة
  if(cont.children.length===0){addMatchAnswerKeyRow();addMatchAnswerKeyRow();addMatchAnswerKeyRow();}
}


function showStudentCertificates(){
  scOk('شهاداتي 🏅','My Certificates',
    'ستظهر شهاداتك هنا بعد اجتياز الاختبارات.<br>سيتم ربطها برمجياً قريباً.',
    'Your certificates will appear here after completing tests.<br>This feature will be linked programmatically soon.','🏅'
  );
}

// ══ White Paper Mode (Mode 4) ══
var _wpZoomLevel=1.7;
function wpZoomChange(delta){
  _wpZoomLevel=Math.max(0.6,Math.min(1.6,parseFloat((_wpZoomLevel+delta).toFixed(2))));
  var el=document.getElementById('wp-page');
  if(el){el.style.transform='scale('+_wpZoomLevel+')';el.style.transformOrigin='top center';}
  var lbl=document.getElementById('wp-zoom-label');
  if(lbl) lbl.textContent=Math.round(_wpZoomLevel*100)+'%';
  // أعد رسم خطوط أسئلة التوصيل الظاهرة حالياً حتى تطابق موضعها الجديد بعد تغيير الزووم
  setTimeout(function(){
    document.querySelectorAll('[id^="smWrap-"]').forEach(function(w){
      var qi=w.id.replace('smWrap-','');
      if(typeof smStreamDraw==='function') smStreamDraw(qi);
    });
  },20);
}
function renderWhitePaper(){
  var d=testData.domains[sw_domainIdx];
  if(!d)return;
  var questions=sw_branchIdx>=0&&d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions||[]:d.questions||[];
  // أظهر دوائر إكمال الأسئلة بنفس أسلوب النمط الكلاسيكي، وألغِ نص "Domain X QY/Z" المكرر
  renderProgressDots(questions.length);
  var ctrTop=document.getElementById('sw-q-counter-top');
  if(ctrTop) ctrTop.textContent='';
  var labels=['A','B','C','D','E','F'];
  var logoHtml='<img src="'+SCHOLASTIC_EAGLE_LOGO_PNG+'" style="width:52px;height:66px;object-fit:contain;flex-shrink:0" alt="Scholastic">';
  var gradeMap={'KS1/FS1':{en:'KS1/FS1',ar:'المرحلة التأسيسية الأولى'},'KS1/FS2':{en:'KS1/FS2',ar:'المرحلة التأسيسية الثانية'},'Grade 1/Year 2':{en:'Grade 1/Year 2',ar:'الصف الأول/السنة الثانية'},'Grade 2/Year 3':{en:'Grade 2/Year 3',ar:'الصف الثاني/السنة الثالثة'},'Grade 3/Year 4':{en:'Grade 3/Year 4',ar:'الصف الثالث/السنة الرابعة'},'Grade 4/Year 5':{en:'Grade 4/Year 5',ar:'الصف الرابع/السنة الخامسة'},'Grade 5/Year 6':{en:'Grade 5/Year 6',ar:'الصف الخامس/السنة السادسة'},'Grade 6/Year 7':{en:'Grade 6/Year 7',ar:'الصف السادس/السنة السابعة'},'Grade 7/Year 8':{en:'Grade 7/Year 8',ar:'الصف السابع/السنة الثامنة'},'Grade 8/Year 9':{en:'Grade 8/Year 9',ar:'الصف الثامن/السنة التاسعة'},'Grade 9/Year 10':{en:'Grade 9/Year 10',ar:'الصف التاسع/السنة العاشرة'},'Grade 10/Year 11':{en:'Grade 10/Year 11',ar:'الصف العاشر/السنة الحادية عشرة'},'Grade 11/Year 12':{en:'Grade 11/Year 12',ar:'الصف الحادي عشر/السنة الثانية عشرة'},'Grade 12/Year 13':{en:'Grade 12/Year 13',ar:'الصف الثاني عشر/السنة الثالثة عشرة'}};
  var termMap={'1':{en:'Term 1',ar:'الفصل الدراسي الأول'},'2':{en:'Term 2',ar:'الفصل الدراسي الثاني'},'3':{en:'Term 3',ar:'الفصل الدراسي الثالث'}};
  var gInfo=gradeMap[testData.grade]||{en:testData.grade||'',ar:testData.grade||''};
  var tInfo=termMap[testData.term]||{en:'Term '+(testData.term||''),ar:'الفصل الدراسي '+(testData.term||'')};
  var yearStr=testData.year||'';
  var subjEn=testData.subject||'';
  var subjAr=testData.subject||''; // المادة عادة تُكتب بنفس النص العربي/الإنجليزي المختار من القائمة
  var html='<div class="wp-page" id="wp-page" style="transform:scale('+_wpZoomLevel+');transform-origin:top center;transition:transform .15s">';
  // Header: اللوجو أعلى اليسار + اسم الاختبار فقط في المنتصف
  html+='<div style="position:relative;text-align:center;padding-bottom:12px;margin-bottom:16px;border-bottom:3px double #141F44;min-height:40px">'
    +'<div style="position:absolute;top:0;left:0">'+logoHtml+'</div>'
    +'<div style="font-family:\'Playfair Display\',Georgia,serif;font-size:21px;font-weight:800;color:#141F44;letter-spacing:.5px;padding-top:6px">'+(testData.testName||'اختبار')+'</div>'
    +'</div>';
  // بطاقة معلومات احترافية (الشريط ثنائي اللغة) — جزء أصيل من الصفحة، بخلفية فاتحة تناسب الورقة البيضاء
  html+='<div style="background:linear-gradient(135deg,#FBF8F1,#F5EFDF);border:1.5px solid #AD8628;border-radius:14px;overflow:hidden;margin-bottom:20px;box-shadow:0 2px 10px rgba(173,134,40,.14)">'
    // Bilingual academic info row: direction ltr → column1(English)=left, column2(Arabic)=right
    +'<div style="display:grid;grid-template-columns:1fr 1fr;direction:ltr">'
      +'<div style="padding:16px 20px;border-right:1px solid rgba(173,134,40,.3);text-align:left;display:flex;align-items:center">'
        +'<div style="font-family:Jost,Arial,sans-serif;font-size:13px;font-weight:600;line-height:1.7;color:#141F44" dir="ltr">Academic Year ('+yearStr+') — '+subjEn+' — '+gInfo.en+' — '+tInfo.en+' Test</div>'
      +'</div>'
      +'<div style="padding:16px 20px;text-align:right;display:flex;align-items:center;justify-content:flex-end">'
        +'<div style="font-family:Tajawal,sans-serif;font-size:14px;font-weight:800;line-height:1.8;color:#141F44" dir="rtl">الاختبار المعياري الدولي في مادة '+subjAr+' للصف '+gInfo.ar+' ('+tInfo.ar+' '+yearStr+')</div>'
      +'</div>'
    +'</div>'
    +'</div>';
  // Domain title (+ branch name if applicable) — في الوسط، إنجليزي أزرق / عربي أحمر غامق
  var domOrdinalsAr=['الأول','الثاني','الثالث','الرابع','الخامس','السادس','السابع','الثامن','التاسع','العاشر','الحادي عشر','الثاني عشر','الثالث عشر','الرابع عشر','الخامس عشر','السادس عشر','السابع عشر','الثامن عشر','التاسع عشر','العشرون'];
  var domOrdinalsEn=['First','Second','Third','Fourth','Fifth','Sixth','Seventh','Eighth','Ninth','Tenth','Eleventh','Twelfth','Thirteenth','Fourteenth','Fifteenth','Sixteenth','Seventeenth','Eighteenth','Nineteenth','Twentieth'];
  var domEn=d.nameEn||d.nameAr||'Domain';
  var domAr=d.nameAr||d.nameEn||'المجال';
  var domOrdAr=domOrdinalsAr[sw_domainIdx]||(sw_domainIdx+1);
  var domOrdEn=domOrdinalsEn[sw_domainIdx]||(sw_domainIdx+1);
  var domainTitle='<div style="text-align:center;line-height:1.6">'
    +'<div style="font-size:13px;font-weight:800;color:#1e40af;font-family:Montserrat,sans-serif" dir="ltr">'+domOrdEn+' Domain: '+domEn+'</div>'
    +'<div style="font-size:13px;font-weight:800;color:#b91c1c" dir="rtl">المجال '+domOrdAr+': '+domAr+'</div>';
  if(sw_branchIdx>=0&&d.branches&&d.branches[sw_branchIdx]){
    var br=d.branches[sw_branchIdx];
    var brEn=br.nameEn||br.nameAr||'Branch';
    var brAr=br.nameAr||br.nameEn||'الفرع';
    var brOrdAr=domOrdinalsAr[sw_branchIdx]||(sw_branchIdx+1);
    var brOrdEn=domOrdinalsEn[sw_branchIdx]||(sw_branchIdx+1);
    domainTitle+='<div style="font-size:12px;font-weight:700;color:#1e40af;font-family:Montserrat,sans-serif;margin-top:3px" dir="ltr">'+brOrdEn+' Branch: '+brEn+'</div>'
      +'<div style="font-size:12px;font-weight:700;color:#b91c1c;margin-top:1px" dir="rtl">الفرع '+brOrdAr+': '+brAr+'</div>';
  }
  domainTitle+='</div>';
  html+='<div class="wp-domain-bar" style="display:block;text-align:center;background:#f8fafc;border:1px solid #e2e8f0;border-radius:8px;padding:10px 16px">'+domainTitle+'</div>';
  // Questions
  questions.forEach(function(q,qi){
    html+='<div class="wp-q" id="wp-q-'+qi+'">';
    // Stem + رقم السؤال ثنائي اللغة
    var qOrdAr=domOrdinalsAr[qi]||(qi+1);
    var qOrdEn=domOrdinalsEn[qi]||(qi+1);
    html+='<div class="wp-q-stem" style="display:flex;flex-direction:column;gap:8px;margin-bottom:8px">'
      +'<div style="display:flex;justify-content:space-between;align-items:center;gap:10px;direction:ltr">'
      +'<div style="background:linear-gradient(135deg,#fef3c7,#fde68a);border:1.5px solid #f59e0b;border-radius:8px;padding:5px 14px;font-size:12px;font-weight:900;color:#1e40af;font-family:Montserrat,sans-serif;white-space:nowrap;box-shadow:0 1px 4px rgba(245,158,11,.2)" dir="ltr">'+qOrdEn+' Question</div>'
      +'<div style="background:linear-gradient(135deg,#fef3c7,#fde68a);border:1.5px solid #f59e0b;border-radius:8px;padding:5px 14px;font-size:13px;font-weight:900;color:#991b1b;white-space:nowrap;box-shadow:0 1px 4px rgba(245,158,11,.2)" dir="rtl">السؤال '+qOrdAr+'</div>'
      +'</div>'
      +'<div style="flex:1" dir="auto">'+(q.stemHtml||q.stemText||'')+'</div>'
      +'</div>';
    // Media - only if visible (و ليس للأنماط التي تعرض الميديا داخل جسمها)
    if(q.mediaHtml&&q.type!=='listening'&&q.type!=='reading'&&q.type!=='writingskill'&&q.type!=='speaking'&&q.type!=='oral'&&!(q.mediaVisible&&q.mediaVisible.img===false)){
      html+='<div class="wp-indent" style="margin-bottom:8px">'+q.mediaHtml+'</div>';
    }
    var labels2=labels;
    if(q.type==='mcq'){
      if(q.bodyHtml) html+='<div class="wp-indent" style="font-size:14px;line-height:1.7;color:#1a1a2e;margin-bottom:8px" dir="auto">'+q.bodyHtml+'</div>';
      if(q.questionText) html+='<div class="wp-indent" style="font-size:14px;font-weight:700;color:#1e3a8a;margin-bottom:8px" dir="auto">'+q.questionText+'</div>';
      html+='<div class="wp-indent wp-mcq-grid">';
      (q.options||[]).forEach(function(o,i){
        var sel=sw_answers[qi]===i;
        html+='<div class="wp-mcq-opt'+(sel?' sel':'')+'" onclick="wpSelectMcq('+qi+','+i+')">'
          +'<div class="wp-opt-circle">'+(sel?'✓':'')+'</div>'
          +'<span style="font-size:11px;font-weight:700;color:#1e3a8a;margin-left:2px">'+(labels2[i]||i+1)+'.</span>'
          +'<span dir="auto" style="flex:1">'+(o||'')+'</span></div>';
      });
      html+='</div>';

    } else if(q.type==='truefalse'){
      var tfFont=q.tfFont||'Tajawal',tfSize=q.tfSize||15,tfColor=q.tfColor||'#1a1a2e',tfBg=q.tfBg||'#f8fafc',tfTrueColor=q.tfTrueColor||'#15803d',tfFalseColor=q.tfFalseColor||'#b91c1c';
      html+='<div class="wp-indent" style="display:flex;flex-direction:column;gap:10px">';
      (q.statements||[]).forEach(function(s,si){
        var ans=sw_answers[qi]?sw_answers[qi][si]:null;
        var trueSel=ans==='true',falseSel=ans==='false';
        html+='<div style="background:'+tfBg+';border:2px solid '+(trueSel?tfTrueColor:falseSel?tfFalseColor:'#e2e8f0')+';border-radius:10px;padding:10px 14px">'
          +'<div style="font-size:'+tfSize+'px;font-weight:600;color:'+tfColor+';margin-bottom:8px;line-height:1.6;font-family:'+tfFont+',sans-serif" dir="auto">'+(si+1)+'. '+(s.text||'')+'</div>'
          +'<div style="display:flex;gap:8px">'
          +'<button class="wp-tf-btn wp-tf-t'+(trueSel?' sel':'')+'" onclick="wpTFSelect('+qi+','+si+',\'true\')">✅ صواب / True</button>'
          +'<button class="wp-tf-btn wp-tf-f'+(falseSel?' sel':'')+'" onclick="wpTFSelect('+qi+','+si+',\'false\')">❌ خطأ / False</button>'
          +'</div></div>';
      });
      html+='</div>';

    } else if(q.type==='reading'||q.type==='writingskill'){
      if(q.mediaHtml&&!(q.mediaVisible&&q.mediaVisible.img===false)) html+='<div class="wp-indent" style="margin-bottom:8px">'+q.mediaHtml+'</div>';
      if(q.passageHtml) html+='<div class="wp-indent" style="background:#f9fafb;border:1px solid #e2e8f0;border-radius:8px;padding:10px;margin-bottom:8px;font-size:13px;line-height:1.8" dir="auto">'+q.passageHtml+'</div>';
      if(q.answerStemHtml) html+='<div class="wp-indent" style="font-size:14px;font-weight:700;margin-bottom:8px" dir="auto">'+q.answerStemHtml+'</div>';
      if(q.ansType==='mcq'){
        html+='<div class="wp-indent wp-mcq-grid">';
        (q.options||[]).forEach(function(o,i){
          var sel=sw_answers[qi]===i;
          html+='<div class="wp-mcq-opt'+(sel?' sel':'')+'" onclick="wpSelectMcq('+qi+','+i+')">'
            +'<div class="wp-opt-circle">'+(sel?'✓':'')+'</div>'
            +'<span style="font-size:11px;font-weight:700;color:#1e3a8a;margin-left:2px">'+(labels2[i]||i+1)+'.</span>'
            +'<span dir="auto" style="flex:1">'+(o||'')+'</span></div>';
        });
        html+='</div>';
      } else if(q.ansType==='matching'){
        html+='<div class="wp-indent">'+buildWPMatching(q,qi)+'</div>';
      } else {
        // كتابة / رسم / صورة
        html+='<div class="wp-indent">'+buildWPWriteArea(qi)+'</div>';
      }

    } else if(q.type==='ordering'){
      var wpOrGroups=q.orderGroups&&q.orderGroups.length?q.orderGroups:(q.words&&q.words.length?[{words:q.words,answerOrder:q.answerOrder}]:[]);
      if(!sw_answers[qi]||!sw_answers[qi].groups){
        sw_answers[qi]={groups:wpOrGroups.map(function(g){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});return {placed:new Array((g.words||[]).length).fill(null),bank:bank};})};
      }
      wpOrGroups.forEach(function(g,gi){
        if(!sw_answers[qi].groups[gi]){var bank=(g.words||[]).slice().sort(function(){return Math.random()-.5;});sw_answers[qi].groups[gi]={placed:new Array((g.words||[]).length).fill(null),bank:bank};}
      });
      var sFz=(q.orFontSize||15)+'px',sFont=q.orFont||'Tajawal',sTBg=q.orTileBg||'#1e3a8a',sTClr=q.orTileColor||'#ffffff',sABg=q.orAnsBg||'#f0f7ff';
      var wpPerGroupScore=wpOrGroups.length>1?(Number(q.score||0)/wpOrGroups.length).toFixed(2):null;
      html+='<div class="wp-indent">';
      wpOrGroups.forEach(function(g,gi){
        var gAns=sw_answers[qi].groups[gi];
        if(wpOrGroups.length>1){
          html+='<div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px"><span style="font-size:12px;font-weight:800;color:'+sTBg+'">📝 الجملة '+(gi+1)+' / Sentence '+(gi+1)+'</span>'+(wpPerGroupScore?'<span style="font-size:10px;font-weight:700;color:'+sTBg+';background:'+sTBg+'18;border-radius:20px;padding:2px 10px">'+wpPerGroupScore+'%</span>':'')+'</div>';
        }
        html+='<div style="display:flex;flex-wrap:wrap;gap:8px;padding:10px;background:linear-gradient(135deg,'+sTBg+'22,'+sTBg+'11);border:2px solid '+sTBg+'44;border-radius:10px;min-height:46px;margin-bottom:10px" id="sor-bank-'+qi+'-'+gi+'">';
        (gAns.bank||[]).forEach(function(w,wi){html+='<div draggable="true" ondragstart="sorBankDragStart(event,'+qi+','+gi+','+wi+')" style="background:'+sTBg+';color:'+sTClr+';border-radius:8px;padding:6px 14px;font-size:'+sFz+';font-weight:800;cursor:grab;font-family:'+sFont+',sans-serif;user-select:none">'+w+'</div>';});
        html+='</div>';
        html+='<div style="display:flex;flex-wrap:wrap;gap:8px;padding:10px;background:'+sABg+';border:2px dashed '+sTBg+'66;border-radius:10px;min-height:46px;margin-bottom:'+(gi<wpOrGroups.length-1?'14px':'0')+'" id="sor-slots-'+qi+'-'+gi+'">';
        (gAns.placed||[]).forEach(function(pw,si){html+='<div id="sor-slot-stream-'+qi+'-'+gi+'-'+si+'" data-slot="'+si+'" ondragover="sorSlotDragOver(event)" ondrop="sorSlotDrop(event,'+qi+','+gi+','+si+')" ondragleave="sorSlotDragLeave(event,'+qi+','+gi+','+si+')" style="min-width:60px;height:40px;border:2px dashed '+(pw?sTBg:'#cbd5e1')+';border-radius:8px;display:flex;align-items:center;justify-content:center;background:'+(pw?sTBg+'18':'white')+'">'+(pw?'<div onclick="sorReturn('+qi+','+gi+','+si+')" style="background:'+sTBg+';color:'+sTClr+';border-radius:6px;padding:5px 10px;font-size:'+sFz+';font-weight:800;cursor:pointer;font-family:'+sFont+',sans-serif">'+pw+' ✕</div>':'<span style="color:#cbd5e1;font-size:11px">'+(si+1)+'</span>')+'</div>';});
        html+='</div>';
      });
      html+='</div>';
      html+='<div class="wp-indent" style="display:flex;gap:8px;justify-content:center;margin-top:10px;flex-wrap:wrap">'
        +'<button onclick="sorClearAll('+qi+')" style="background:rgba(239,68,68,.1);color:#ef4444;border:1.5px solid rgba(239,68,68,.3);border-radius:10px;padding:6px 16px;font-size:11px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🔄 إعادة المحاولة / Try Again</button>'
        +'<button onclick="swOrderConfirm('+qi+')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:1.5px solid #22c55e;border-radius:10px;padding:6px 18px;font-size:11px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">✅ تأكيد / Confirm</button>'
        +'</div>';

    } else if(q.type==='speaking'||q.type==='oral'){
      html+='<div class="wp-indent">';
      if(q.mediaHtml&&!(q.mediaVisible&&q.mediaVisible.img===false)) html+='<div style="margin-bottom:10px;border-radius:10px;overflow:hidden">'+q.mediaHtml+'</div>';
      if(q.type==='oral'&&q.oralText){
        var oFs=(q.speakingSize||16)+'px',oFont=q.speakingFont||'Tajawal',oClr=q.speakingColor||'#1a1a2e',oBg=q.speakingBg||'#fffbeb';
        html+='<div style="font-size:'+oFs+';color:'+oClr+';background:'+oBg+';padding:14px 18px;border-radius:10px;border:2px solid rgba(250,204,21,.4);margin-bottom:10px;line-height:2;font-family:'+oFont+',sans-serif" dir="auto">'+q.oralText+'</div>';
      }
      html+='<div style="color:#64748b;font-size:12px;margin-bottom:8px;padding:8px 12px;background:#f8fafc;border-radius:8px" dir="auto">'+(q.speakingInst||(q.type==='oral'?'📖 اقرأ النص أعلاه بصوت واضح ثم سجّل قراءتك':'🎙 سجّل إجابتك الشفهية'))+'</div>';
      html+='<div style="background:#f0f9ff;border:2px solid #bae6fd;border-radius:10px;padding:12px;text-align:center">';
      html+='<button id="sRecBtn-'+qi+'" onclick="sRecToggle('+qi+')" style="background:linear-gradient(135deg,#ef4444,#dc2626);color:white;border:none;border-radius:10px;padding:8px 20px;font-size:13px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🎙 ابدأ التسجيل / Record</button>';
      html+='<div id="sRecSt-'+qi+'" style="font-size:11px;color:#64748b;margin-top:6px;font-family:Montserrat,sans-serif;min-height:14px"></div>';
      html+='<audio id="sRecAud-'+qi+'" controls style="display:none;width:100%;margin-top:8px;border-radius:8px;direction:ltr"></audio>';
      html+='<div id="sRecCtrl-'+qi+'" style="display:none;gap:8px;margin-top:10px;flex-wrap:wrap;justify-content:center">';
      html+='<button onclick="sRecListen('+qi+')" style="background:#3b82f6;color:white;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">🔊 استماع / Listen</button>';
      html+='<button onclick="sRecDelete('+qi+')" style="background:#f59e0b;color:white;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:700">🔄 إعادة / Retry</button>';
      html+='<button onclick="sRecConfirm('+qi+')" style="background:#22c55e;color:white;border:none;border-radius:8px;padding:6px 14px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:12px;font-weight:800">✅ تأكيد / Confirm</button>';
      html+='</div></div></div>';

    } else if(q.type==='listening'){
      html+='<div class="wp-indent">';
      if(q.mediaHtml) html+='<div style="margin-bottom:10px">'+q.mediaHtml+'</div>';
      if(q.ansType==='mcq'){
        html+='<div class="wp-mcq-grid">';
        (q.options||[]).forEach(function(o,i){
          var sel=sw_answers[qi]===i;
          html+='<div class="wp-mcq-opt'+(sel?' sel':'')+'" onclick="wpSelectMcq('+qi+','+i+')">'
            +'<div class="wp-opt-circle">'+(sel?'✓':'')+'</div>'
            +'<span style="font-size:11px;font-weight:700;color:#1e3a8a;margin-left:2px">'+(labels2[i]||i+1)+'.</span>'
            +'<span dir="auto" style="flex:1">'+(o||'')+'</span></div>';
        });
        html+='</div>';
      } else if(q.ansType==='matching'){
        html+=buildWPMatching(q,qi);
      } else {
        html+=buildWPWriteArea(qi,true);
      }
      html+='</div>';

    } else if(q.type==='matching'){
      html+='<div class="wp-indent">'+buildWPMatching(q,qi)+'</div>';

    } else if(q.type==='classify'){
      html+='<div class="wp-indent">'+buildWPClassify(q,qi)+'</div>';

    } else {
      // writing / other
      html+='<div class="wp-indent">'+buildWPWriteArea(qi)+'</div>';
    }
    html+='</div>'; // wp-q
    if(qi<questions.length-1) html+='<div class="wp-q-sep"></div>';
  });
  // زر تسليم المجال/الفرع — بنفس منطق وتحذير النمط الكلاسيكي
  html+='<div style="margin-top:24px;padding:20px;text-align:center;background:linear-gradient(135deg,#fef2f2,#fee2e2);border:2px dashed #ef4444;border-radius:14px">'
    +'<div style="font-size:12px;color:#991b1b;margin-bottom:12px;font-weight:700;line-height:1.6">'
    +'⚠️ بعد الضغط على زر التسليم لن تتمكن من العودة لهذا '+(sw_branchIdx>=0?'الفرع':'المجال')+' مرة أخرى<br>'
    +'<span style="font-family:Montserrat,sans-serif;font-weight:600">After submitting, you will not be able to return to this '+(sw_branchIdx>=0?'branch':'domain')+' again</span>'
    +'</div>'
    +'<button onclick="swSubmitDomain('+questions.length+')" style="background:linear-gradient(135deg,#22c55e,#15803d);color:white;border:none;border-radius:16px;padding:14px 36px;font-size:15px;font-weight:900;cursor:pointer;font-family:Tajawal,sans-serif;box-shadow:0 4px 16px rgba(34,197,94,.4)">'
    +'✅ تسليم '+(sw_branchIdx>=0?'الفرع':'المجال')+' / Submit '+(sw_branchIdx>=0?'Branch':'Domain')
    +'</button>'
    +'</div>';
  // Footer
  html+='<div class="wp-page-footer"><span>'+('Scholastic Testing Platform')+'</span><span style="font-family:Montserrat,sans-serif">Page 1</span></div>';
  html+='</div>'; // wp-page
  var cont=document.getElementById('sw-questions-content');
  if(cont) cont.innerHTML=html;
  // Update close btn label
  var cb=document.getElementById('sw-close-btn');
  if(cb&&_tryModeActive) cb.textContent='← رجوع للأنماط';
}
function wpSelectMcq(qi,i){sw_answers[qi]=i;streamMarkAnswered?streamMarkAnswered(qi):null;updateProgressDots();renderWhitePaper();}
function wpTFSelect(qi,si,val){if(!sw_answers[qi])sw_answers[qi]={};sw_answers[qi][si]=val;updateProgressDots();renderWhitePaper();}

// ══ Matching — ورقة بيضاء (يعيد استخدام نفس آلية smStreamClick) ══
function buildWPMatching(q,qi){
  var pairs=q.pairs||[];
  var shuffledB=pairs.map(function(_,bi){return bi;}).sort(function(){return Math.random()-.5;});
  var html='<div style="position:relative;overflow:visible" id="smWrap-'+qi+'">';
  html+='<svg id="smSvg-'+qi+'" style="position:absolute;top:0;left:0;width:100%;height:100%;pointer-events:none;z-index:5;overflow:visible" overflow="visible"></svg>';
  html+='<div style="display:grid;grid-template-columns:1fr 1fr;grid-auto-rows:min-content;align-items:stretch;gap:8px 10px">';
  pairs.forEach(function(p,i){
    var bi=shuffledB[i];
    var pb=pairs[bi]||{};
    html+='<div style="display:flex;align-items:stretch;gap:6px;height:100%">'
      +'<div style="flex:1;border:1.5px solid #e2e8f0;border-radius:8px;padding:8px 10px;background:#fafafa;font-size:12px;cursor:pointer;display:flex;align-items:center;justify-content:center" id="smA-'+qi+'-'+i+'" dir="auto" onclick="smStreamClick(\'A\','+qi+','+i+')">'+(p.aHtml||'—')+'</div>'
      +'<div style="display:flex;align-items:center;flex-shrink:0"><div id="smDA-'+qi+'-'+i+'" style="width:11px;height:11px;border-radius:50%;background:#cbd5e1;border:2px solid #94a3b8"></div></div>'
      +'</div>';
    html+='<div style="display:flex;align-items:stretch;gap:6px;height:100%">'
      +'<div style="display:flex;align-items:center;flex-shrink:0"><div id="smDB-'+qi+'-'+bi+'" style="width:11px;height:11px;border-radius:50%;background:#cbd5e1;border:2px solid #94a3b8"></div></div>'
      +'<div style="flex:1;border:1.5px solid #e2e8f0;border-radius:8px;padding:8px 10px;background:#fafafa;font-size:12px;cursor:pointer;display:flex;align-items:center;justify-content:center" id="smB-'+qi+'-'+bi+'" dir="auto" onclick="smStreamClick(\'B\','+qi+','+bi+')">'+(pb.bHtml||'—')+'</div>'
      +'</div>';
  });
  html+='</div>';
  html+='<div style="margin-top:8px;display:flex;justify-content:center">'
    +'<button onclick="smStreamClear('+qi+')" style="background:#fee2e2;border:1.5px solid #fca5a5;color:#ef4444;border-radius:8px;padding:6px 16px;font-size:12px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">✕ مسح الكل / Clear All</button>'
    +'</div></div>';
  // تهيئة الاتصالات بعد الرسم
  setTimeout(function(){ smStreamInit(qi); },60);
  return html;
}

// ══ كتابة / رسم / صورة — ورقة بيضاء (نفس أدوات الكلاسيكي) ══
function buildWPWriteArea(qi,compact){
  var val=typeof sw_answers[qi]==='string'?sw_answers[qi]:'';
  var minH=compact?'90px':'120px';
  var html='<div>'
    +'<div style="display:flex;gap:6px;margin-bottom:8px;flex-wrap:wrap">'
    +'<button onclick="swToggleWriteMode(\'write\')" style="border:1.5px solid #1e3a8a;background:#1e3a8a;color:white;border-radius:8px;padding:5px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:11px;font-weight:700">⌨️ كتابة / Write</button>'
    +'<button onclick="swToggleWriteMode(\'draw\')" style="border:1.5px solid #e2e8f0;background:white;color:#475569;border-radius:8px;padding:5px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:11px;font-weight:700">🎨 رسم / Draw</button>'
    +'<label style="border:1.5px solid #e2e8f0;background:white;color:#475569;border-radius:8px;padding:5px 12px;cursor:pointer;font-family:Tajawal,sans-serif;font-size:11px;font-weight:700;display:inline-flex;align-items:center;gap:5px">📷 صورة / Image<input type="file" accept="image/*" style="display:none" onchange="swUploadImage(event)"></label>'
    +'</div>'
    +'<div id="sw-write-area"><textarea class="sw-textarea" style="min-height:'+minH+'" placeholder="اكتب إجابتك هنا... / Write your answer here..." dir="auto" onchange="sw_answers['+qi+']=this.value;updateProgressDots()">'+val+'</textarea></div>'
    +'<div id="sw-draw-area" style="display:none"><div class="sw-canvas-wrap"><div class="sw-canvas-tools">'
    +'<button id="sw-pen" class="active" onclick="setDrawTool(\'pen\')">✏️ قلم / Pen</button>'
    +'<button id="sw-eraser" onclick="setDrawTool(\'eraser\')">🧹 ممحاة / Eraser</button>'
    +'<select id="sw-brush-size" onchange="setBrushSize(this.value)" style="border:1px solid #e2e8f0;border-radius:6px;padding:4px 8px;font-size:12px"><option value="2">رفيع / Thin</option><option value="5" selected>عادي</option><option value="10">سميك / Thick</option></select>'
    +'<input type="color" id="sw-pen-color" value="#1a1a2e" onchange="setPenColor(this.value)">'
    +'<button onclick="clearCanvas()" style="color:#ef4444;font-size:12px;border:1px solid #fca5a5;border-radius:6px;padding:4px 10px">🗑 مسح</button>'
    +'</div><canvas id="drawingCanvas" height="220" style="width:100%;display:block;cursor:crosshair;touch-action:none"></canvas></div></div>'
    +'<div id="sw-upload-area" style="display:none"><div id="sw-uploaded-img-wrap" style="border:2px dashed #e2e8f0;border-radius:12px;padding:20px;text-align:center;color:#94a3b8;font-size:12px">اضغط زر الصورة أعلاه / Click image button above</div></div>'
    +'</div>';
  return html;
}

// ══ التصنيف — ورقة بيضاء (يعيد استخدام نفس آلية السحب والإفلات) ══
function buildWPClassify(q,qi){
  var cols=q.columns||[];
  var items=q.items||[];
  var numCols=Math.max(1,cols.length);
  var numRows=Math.ceil(items.length/numCols);
  if(!sw_answers[qi]){
    var initBank=items.map(function(it){return it.text;}).sort(function(){return Math.random()-.5;});
    sw_answers[qi]={placed:{},bank:initBank,confirmed:false};
    cols.forEach(function(_,ci){sw_answers[qi].placed[ci]=[];});
  }
  var bankItems=sw_answers[qi].bank||[];
  var chipBg=q.orTileBg||'#6366f1';
  var tableBg=q.clTableBg||'#f0f7ff';
  var clFs=(q.clFontSize||13)+'px';
  var clTxt=q.clTextColor||'#ffffff';

  var tableHtml='<table style="width:100%;border-collapse:collapse;margin-bottom:10px;background:'+tableBg+';table-layout:fixed">';
  tableHtml+='<thead><tr>';
  cols.forEach(function(col){
    tableHtml+='<th style="background:'+chipBg+';color:'+clTxt+';font-weight:900;font-size:'+clFs+';font-family:Tajawal,sans-serif;padding:7px 6px;text-align:center;border:2px solid rgba(255,255,255,.3)">'+(col||'—')+'</th>';
  });
  tableHtml+='</tr></thead><tbody>';
  for(var ri=0;ri<numRows;ri++){
    tableHtml+='<tr>';
    for(var ci=0;ci<numCols;ci++){
      var placed=sw_answers[qi].placed[ci]||[];
      var cellWord=placed[ri]||null;
      tableHtml+='<td id="cl-cell-'+qi+'-'+ci+'-'+ri+'" data-qi="'+qi+'" data-ci="'+ci+'" data-ri="'+ri+'"'
        +' ondragover="clCellDragOver(event)" ondrop="clCellDrop(event,'+qi+','+ci+','+ri+')" ondragleave="clCellDragLeave(event)"'
        +' style="border:2px dashed '+(cellWord?chipBg+'88':'#cbd5e1')+';border-radius:8px;min-height:40px;height:42px;text-align:center;vertical-align:middle;padding:4px;transition:.2s;cursor:default;background:'+(cellWord?chipBg+'18':tableBg)+'">';
      if(cellWord){
        tableHtml+='<div draggable="true" id="cl-placed-'+qi+'-'+ci+'-'+ri+'" data-word="'+cellWord+'" data-from-col="'+ci+'" data-from-row="'+ri+'"'
          +' ondragstart="clPlacedDragStart(event,'+qi+','+ci+','+ri+')" onclick="clPlacedClick('+qi+','+ci+','+ri+')"'
          +' style="background:'+chipBg+';color:'+clTxt+';border-radius:6px;padding:4px 10px;font-size:'+clFs+';font-weight:700;font-family:Tajawal,sans-serif;cursor:pointer;display:inline-block;user-select:none;max-width:100%;word-break:break-word">'+cellWord+'</div>';
      }
      tableHtml+='</td>';
    }
    tableHtml+='</tr>';
  }
  tableHtml+='</tbody></table>';

  return '<div style="background:linear-gradient(135deg,'+chipBg+','+chipBg+'cc);border-radius:10px;padding:10px;margin-bottom:10px">'
    +'<div style="font-size:9px;font-weight:800;color:rgba(255,255,255,.8);font-family:Montserrat,sans-serif;letter-spacing:1px;margin-bottom:8px;text-align:center">WORD BANK / مصدر المفردات</div>'
    +'<div id="cl-sw-bank-'+qi+'" ondragover="clBankDragOver(event)" ondrop="clBankDrop(event,'+qi+')" ondragleave="clBankDragLeave(event)" style="display:flex;flex-wrap:wrap;gap:6px;min-height:38px;justify-content:center;border-radius:8px;padding:5px;border:2px dashed rgba(255,255,255,.3);transition:.2s">'
    +bankItems.map(function(w,wi){
      return '<div draggable="true" class="cl-bank-chip" id="cl-chip-'+qi+'-'+wi+'" data-word="'+w+'" data-wi="'+wi+'"'
        +' ondragstart="clBankDragStart(event,'+qi+','+wi+')"'
        +' style="background:rgba(255,255,255,.22);color:'+clTxt+';border:2px solid rgba(255,255,255,.4);border-radius:8px;padding:5px 12px;font-size:'+clFs+';font-weight:700;cursor:grab;user-select:none;font-family:Tajawal,sans-serif;transition:.2s">'+w+'</div>';
    }).join('')
    +(bankItems.length===0?'<span style="color:rgba(255,255,255,.6);font-size:12px;font-family:Montserrat,sans-serif">✅ All words placed</span>':'')
    +'</div></div>'
    +tableHtml
    +'<div style="display:flex;gap:8px;justify-content:center">'
    +'<button onclick="clSwClear('+qi+')" style="background:rgba(239,68,68,.1);color:#ef4444;border:2px solid rgba(239,68,68,.3);border-radius:10px;padding:6px 16px;font-size:12px;font-weight:800;cursor:pointer;font-family:Tajawal,sans-serif">🗑 مسح الكل / Clear</button>'
    +'</div>'
    +'<div id="cl-sw-result-'+qi+'" style="display:none;margin-top:8px;text-align:center;padding:10px;border-radius:10px;font-size:13px;font-weight:800"></div>';
}
function showStudentInstructions(){
  sw_currentPage='instructions';
  sw_instructionsConfirmed=false;
  renderStudentWindowPage();
}
function swConfirmInstructions(){
  sw_instructionsConfirmed=true;
  sw_currentPage='domains';
  renderStudentWindowPage();
}
var _clDragData={};
function clBankDragStart(e,qi,wi){
  _clDragData={from:'bank',qi:qi,wi:wi,word:sw_answers[qi].bank[wi]};
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain',sw_answers[qi].bank[wi]);
}
function clPlacedDragStart(e,qi,ci,ri){
  _clDragData={from:'cell',qi:qi,ci:ci,ri:ri,word:sw_answers[qi].placed[ci][ri]};
  e.dataTransfer.effectAllowed='move';
  e.dataTransfer.setData('text/plain',sw_answers[qi].placed[ci][ri]);
}
// ١٢- الضغط على كلمة موضوعة في الجدول يُرجعها لصندوق البنك العلوي
function clPlacedClick(qi,ci,ri){
  var ans=sw_answers[qi];if(!ans||!ans.placed[ci])return;
  var word=ans.placed[ci][ri];if(!word)return;
  ans.bank.push(word);
  ans.placed[ci][ri]=null;
  ans.confirmed=false;
  updateProgressDots();renderStudentQuestion();
}
function clCellDragOver(e){e.preventDefault();e.currentTarget.style.background='#e0f2fe';e.currentTarget.style.borderColor='#0ea5e9';}
function clCellDragLeave(e){e.currentTarget.style.background='#f8fafc';e.currentTarget.style.borderColor='#cbd5e1';}
function clCellDrop(e,qi,ci,ri){
  e.preventDefault();
  var el=e.currentTarget;el.style.background='#f8fafc';el.style.borderColor='#cbd5e1';
  var d=_clDragData;if(!d||d.qi!==qi)return;
  var ans=sw_answers[qi];
  var incomingWord=d.word;
  var existingWord=ans.placed[ci]&&ans.placed[ci][ri]?ans.placed[ci][ri]:null;
  if(!ans.placed[ci])ans.placed[ci]=[];
  // إذا كانت الخلية مشغولة — الكلمة الأولى ترجع للبنك
  if(existingWord){
    ans.bank.push(existingWord);
  }
  if(d.from==='bank'){
    ans.bank=ans.bank.filter(function(w,i){return i!==d.wi;});
    ans.placed[ci][ri]=incomingWord;
  } else if(d.from==='cell'){
    // أزل من الخلية الأصلية
    if(ans.placed[d.ci])ans.placed[d.ci][d.ri]=null;
    ans.placed[ci][ri]=incomingWord;
  }
  // نظف nulls
  Object.keys(ans.placed).forEach(function(k){ans.placed[k]=ans.placed[k].filter(function(x){return x!==null;});});
  if(ans.bank.length===0)ans.confirmed=true;
  updateProgressDots();renderStudentQuestion();
}
function clBankDragOver(e){e.preventDefault();e.currentTarget.style.borderColor='rgba(255,255,255,.8)';}
function clBankDragLeave(e){e.currentTarget.style.borderColor='rgba(255,255,255,.3)';}
function clBankDrop(e,qi){
  e.preventDefault();
  e.currentTarget.style.borderColor='rgba(255,255,255,.3)';
  var d=_clDragData;if(!d||d.qi!==qi||d.from!=='cell')return;
  var ans=sw_answers[qi];
  if(ans.placed[d.ci])ans.placed[d.ci][d.ri]=null;
  ans.bank.push(d.word);
  Object.keys(ans.placed).forEach(function(k){ans.placed[k]=ans.placed[k].filter(function(x){return x!==null;});});
  ans.confirmed=false;
  updateProgressDots();renderStudentQuestion();
}
function clSwClear(qi){
  var d=testData.domains[sw_domainIdx];
  var q=sw_branchIdx>=0?(d.branches&&d.branches[sw_branchIdx]?d.branches[sw_branchIdx].questions[qi]:null):(d.questions?d.questions[qi]:null);
  if(!q)return;
  var initBank=q.items.map(function(it){return it.text;}).sort(function(){return Math.random()-.5;});
  var placed={};(q.columns||[]).forEach(function(_,ci){placed[ci]=[];});
  sw_answers[qi]={placed:placed,bank:initBank,confirmed:false};
  updateProgressDots();renderStudentQuestion();
}
// إبقاء الدوال القديمة للتوافق
function clSwDragStart(e,from,qi,idx){_clDragData={from:from==='bank'?'bank':'cell',qi:qi,wi:idx};e.dataTransfer.effectAllowed='move';}
function clSwDragOver(e){e.preventDefault();}
function clSwDragLeave(e){}
function clSwDrop(e,qi,ci){e.preventDefault();}
function clSwBankClick(qi,wi){}
function clSwReturn(qi,ci,wi){
  var w=sw_answers[qi].placed[ci][wi];if(!w)return;
  sw_answers[qi].placed[ci].splice(wi,1);
  sw_answers[qi].bank.push(w);
  updateProgressDots();renderStudentQuestion();
}
function clSwConfirm(qi){
  sw_answers[qi].confirmed=true;streamMarkAnswered&&streamMarkAnswered(qi);updateProgressDots();
}

// ══ Video Media Helpers ══
function attachVideoFile(event){
  var file=event.target.files?event.target.files[0]:null;if(!file)return;
  var reader=new FileReader();
  reader.onload=function(ev){
    var mp=document.getElementById('mediaPreviewArea');if(!mp)return;
    _insertVideoPreview(ev.target.result,'file');
  };
  reader.readAsDataURL(file);
  event.target.value='';
}
function attachVideoURL(){
  scPromptText('رابط الفيديو المباشر','Direct Video URL','رابط ملف فيديو مباشر (mp4, webm...) — لا يدعم روابط يوتيوب / Direct file link only — YouTube links are not supported','🎬').then(function(url){
    if(!url||!url.trim())return;
    url=url.trim();
    if(/youtube\.com|youtu\.be/i.test(url)){
      scWarn('روابط يوتيوب غير مدعومة لأنها قد لا تعمل داخل الاختبار وتتيح للطالب رابط خروج. ارفع ملف الفيديو من الجهاز بدلاً من ذلك','YouTube links are not supported — they may fail to embed and give students an exit link. Please upload the video file from your device instead');
      return;
    }
    _insertVideoPreview(url,'url');
  });
}
function _insertVideoPreview(src,type){
  var mp=document.getElementById('mediaPreviewArea');if(!mp)return;
  var wrap=document.createElement('div');
  wrap.style.cssText='margin-top:8px;border-radius:12px;overflow:hidden;background:#0a0a1a;border:1px solid rgba(124,58,237,.4)';
  // toolbar
  var tb=document.createElement('div');
  tb.style.cssText='display:flex;gap:6px;padding:8px 10px;background:#1a0535;border-bottom:1px solid rgba(124,58,237,.3);align-items:center';
  tb.innerHTML='<span style="font-size:11px;color:rgba(255,255,255,.6);font-family:Montserrat,sans-serif">🎬 Video (self-hosted)</span>';
  var confirmBtn=document.createElement('button');
  confirmBtn.textContent='✅ اعتماد / Approve';
  confirmBtn.style.cssText='margin-right:auto;background:rgba(34,197,94,.2);border:1px solid rgba(34,197,94,.4);color:#4ade80;border-radius:6px;padding:3px 10px;font-size:11px;cursor:pointer;font-family:Tajawal,sans-serif';
  confirmBtn.onclick=function(){tb.style.background='#052e16';confirmBtn.textContent='✅ معتمد / Approved';confirmBtn.disabled=true;};
  var removeBtn=document.createElement('button');
  removeBtn.textContent='✕';
  removeBtn.style.cssText='background:rgba(239,68,68,.2);border:1px solid rgba(239,68,68,.3);color:#f87171;border-radius:6px;padding:3px 8px;font-size:11px;cursor:pointer';
  removeBtn.onclick=function(){mp.innerHTML='';};
  tb.appendChild(confirmBtn);tb.appendChild(removeBtn);
  wrap.appendChild(tb);
  // فيديو مباشر فقط — بدون iframe يوتيوب، حتى يُحفظ ويُعرض داخل الاختبار مباشرة دون رابط خروج
  var mediaEl=document.createElement('video');
  mediaEl.src=src;mediaEl.controls=true;
  mediaEl.style.cssText='width:100%;max-height:240px;display:block;background:#000';
  wrap.appendChild(mediaEl);
  mp.innerHTML='';mp.appendChild(wrap);
}

function previewSingleQuestion(idx){
  var qs=_getCurrentQuestions();
  var q=qs[idx];if(!q) return;
  var d=testData.domains[currentDomainIndex];
  var score=q.score||(d.weight/Math.max(qs.length,1)).toFixed(2);
  var labels=['A','B','C','D','E','F'];
  var bodyHtml=buildQuestionBodyHtml(q,idx,labels,score);
  // Build preview modal
  var overlay=document.createElement('div');
  overlay.style.cssText='position:fixed;inset:0;z-index:99999;background:rgba(0,0,0,.85);backdrop-filter:blur(8px);display:flex;align-items:center;justify-content:center;padding:20px';
  var box=document.createElement('div');
  box.style.cssText='width:100%;max-width:680px;max-height:85vh;overflow-y:auto;border-radius:20px;overflow:hidden;box-shadow:0 0 0 3px #1e3a8a,0 0 0 6px rgba(250,204,21,.5)';
  box.innerHTML='<div style="background:#1e3a8a;padding:12px 16px;display:flex;align-items:center;justify-content:space-between">'
    +'<span style="color:#FACC15;font-weight:800;font-family:Montserrat,sans-serif;font-size:13px">👁 معاينة السؤال / Question Preview</span>'
    +'<button onclick="_rmVideoPreview(this)" style="background:rgba(255,255,255,.2);border:none;color:white;border-radius:8px;padding:4px 12px;cursor:pointer;font-size:13px">✕ إغلاق</button>'
    +'</div>'
    +'<div class="sw-q-header"><div class="sw-q-badge">Q.'+(idx+1)+'</div><div class="sw-q-stem" dir="auto">'+(q.stemHtml||'Question')+'</div><div class="sw-q-score">'+score+'%</div></div>'
    +(q.mediaHtml?'<div style="padding:14px 18px;background:white;border-bottom:2px solid #e2e8f0">'+q.mediaHtml+'</div>':'')
    +(q.bodyHtml?'<div style="padding:10px 18px;background:white;font-size:15px;line-height:1.8;color:#1a1a2e" dir="auto">'+q.bodyHtml+'</div>':'')
    +(q.questionText?'<div style="padding:10px 18px;background:#eff6ff;border-bottom:2px solid #bfdbfe"><div style="font-size:15px;font-weight:700;color:#1e3a8a" dir="auto">'+q.questionText+'</div></div>':'')
    +'<div style="padding:18px;background:#f0fdf4">'+bodyHtml+'</div>';
  overlay.appendChild(box);
  document.body.appendChild(overlay);
  overlay.onclick=function(e){if(e.target===overlay)overlay.remove();};
}

function removeAudioPreview(btn){var mp=document.getElementById("mediaPreviewArea");if(mp)mp.innerHTML="";}
function _rmVideoPreview(btn){try{var el=btn;for(var i=0;i<5;i++){el=el.parentElement;if(!el)break;}if(el)el.remove();}catch(e){var mp=document.getElementById("mediaPreviewArea");if(mp)mp.innerHTML="";}}

window.onload=function(){
  // أخفِ الـ loader فوراً
  var ldr=document.getElementById('ghLoader');
  if(ldr) ldr.style.display='none';
  try{
    ['adminPanel','supervisorPanel','supManager','schoolManager','generalReviewerPanel','schoolCoordPanel'].forEach(function(id){
      var el=document.getElementById(id); if(el) el.classList.add('hidden');
    });
    populateCountrySelects();
    animateValue('schoolsCount',0,520,2);
    animateValue('examsCount',0,20000,2);
  }catch(e){ console.warn('onload error:',e); }
  // GitHub في الخلفية بعد 200ms
  setTimeout(function(){
    initDataFromGitHub().then(function(){
      try{ populateReviewerDropdowns(); }catch(e){}
    }).catch(function(){
      var ldr2=document.getElementById('ghLoader');
      if(ldr2) ldr2.style.display='none';
    });
  },200);
};
</script>
</body>
</html>
