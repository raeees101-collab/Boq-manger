
import { useState, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── Theme System ──────────────────────────────────────────────────────────────
const THEMES = {
  light: {
    bg:"#F7F6F3", surface:"#FFFFFF", surfaceAlt:"#F0EDE8",
    border:"#E5E0D8", borderStrong:"#C8C0B5",
    text:"#1A1815", textSub:"#6B6560", textMuted:"#A09890",
    accent:"#B5894A", accentLight:"#F5ECD8", accentDark:"#8A6530",
    green:"#3D7A5F", greenLight:"#E8F4EE",
    red:"#C04A3F", redLight:"#FBE8E7",
    blue:"#3A6EA8", blueLight:"#E6EFF8",
  },
  dark: {
    bg:"#141210", surface:"#1E1B18", surfaceAlt:"#272320",
    border:"#3A342E", borderStrong:"#504840",
    text:"#F0EDE8", textSub:"#A09080", textMuted:"#6B6560",
    accent:"#C8964E", accentLight:"#3A2A14", accentDark:"#E5B870",
    green:"#4A9A72", greenLight:"#162A1F",
    red:"#D45A50", redLight:"#2A1210",
    blue:"#4A82C4", blueLight:"#0E1E30",
  }
};

const FONT = "'Noto Kufi Arabic','Cairo',sans-serif";
const UNITS = ["م²","م.ط","وحدة","قطعة","طقم","كجم","لتر","PCS","م³","SET","نقطة"];
const CATEGORIES = ["دهانات","إضاءة","نجارة","أرضيات","أسقف","ورق جدران","رخام","أثاث","ستائر","أعمال معدنية","جبس","كهرباء","سباكة","تمهيد"];

const mockProjects = [
  { id:1, name:"فيلا المهندس محمد الحاجري", client:"محمد الحاجري", location:"الرياض - حي الملقا", status:"جاري", date:"2025-03-15", items:42, specs:18, images:31 },
  { id:2, name:"مكتب شركة التقنية", client:"شركة ألفا للتقنية", location:"الرياض - طريق العليا", status:"مكتمل", date:"2025-01-10", items:28, specs:12, images:19 },
  { id:3, name:"شقة أبراج المملكة", client:"عبدالعزيز السعد", location:"الرياض - أبراج المملكة", status:"معلق", date:"2025-05-01", items:15, specs:6, images:8 },
];

const mockItems = [
  { id:1, category:"دهانات", name:"دهان جدران الغرف", spec:"دهان جوتن - لون أبيض كسر - مطفي - طبقتين بعد الأساس", supplier:"شركة الجزيرة للدهانات", qty:450, unit:"م²", unitPrice:"12", notes:"", image:null },
  { id:2, category:"إضاءة", name:"سبوت لايت LED سقف", spec:"سبوت لايت LED - 10W - 3000K - IP20 - فيليبس RS141B", supplier:"مؤسسة النور للإضاءة", qty:48, unit:"وحدة", unitPrice:"85", notes:"حسب مخطط الإضاءة", image:null },
  { id:3, category:"أرضيات", name:"رخام كرارا أبيض", spec:"رخام إيطالي كرارا - 60×60 سم - سماكة 2 سم - مصقول", supplier:"دار الرخام الفاخر", qty:320, unit:"م²", unitPrice:"220", notes:"مستورد من إيطاليا", image:null },
  { id:4, category:"نجارة", name:"باب غرفة رئيسية", spec:"باب MDF - قشرة جوز - 220×90 سم - بولي يوريثان لامع", supplier:"ورشة الأصالة للنجارة", qty:1, unit:"قطعة", unitPrice:"1800", notes:"", image:null },
  { id:5, category:"أسقف", name:"جبس بورد سقف المجلس", spec:"جبس بورد 12.5 مم - هيكل معدني - إضاءة مخفية", supplier:"مؤسسة الإتقان للجبس", qty:85, unit:"م²", unitPrice:"95", notes:"مع خط إضاءة LED", image:null },
];

const mockSuppliers = [
  { id:1, name:"شركة الجزيرة للدهانات", phone:"0112345678", email:"info@alzazirah.com", category:"دهانات" },
  { id:2, name:"مؤسسة النور للإضاءة", phone:"0556789012", email:"sales@nour.com", category:"إضاءة" },
  { id:3, name:"دار الرخام الفاخر", phone:"0501234567", email:"info@marble.com", category:"أرضيات" },
];

// ─── Export Excel ──────────────────────────────────────────────────────────────
function dlExcel(projectName, items, companyName) {
  const wb = XLSX.utils.book_new();
  const cn = companyName || "ركيز ديزاين";
  const coverWs = XLSX.utils.aoa_to_sheet([
    [""],["",`جدول الكميات — ${projectName||""}`],
    ["","الشركة:",cn],["","التاريخ:",new Date().toLocaleDateString("ar-SA")],
    ["","البنود:",items.length],[""]
  ]);
  coverWs["!cols"]=[{wch:4},{wch:22},{wch:36}];
  XLSX.utils.book_append_sheet(wb, coverWs, "Cover");

  const hdr = ["#","الفئة","اسم البند","المواصفة","الكمية","الوحدة","سعر الوحدة","الإجمالي","المورد","ملاحظات"];
  const cats = [...new Set(items.map(i=>i.category))];
  let n=1; const rows=[];
  cats.forEach(cat=>{
    rows.push(["","── "+cat+" ──","","","","","","","",""]);
    items.filter(i=>i.category===cat).forEach(it=>{
      const t=(parseFloat(it.qty)||0)*(parseFloat(it.unitPrice||0)||0);
      rows.push([n++,cat,it.name,it.spec||"",it.qty||"",it.unit,it.unitPrice||"",t||"",it.supplier||"",it.notes||""]);
    });
  });
  const grand = items.reduce((s,it)=>s+(parseFloat(it.qty)||0)*(parseFloat(it.unitPrice||0)||0),0);
  rows.push(["","","","","","","المجموع",grand,"",""]);
  const ws = XLSX.utils.aoa_to_sheet([hdr,...rows]);
  ws["!cols"]=[{wch:4},{wch:14},{wch:32},{wch:40},{wch:8},{wch:8},{wch:12},{wch:14},{wch:24},{wch:18}];
  XLSX.utils.book_append_sheet(wb, ws, "BOQ");
  cats.forEach(cat=>{
    const ci=items.filter(i=>i.category===cat);
    const cws=XLSX.utils.aoa_to_sheet([
      ["#","اسم البند","المواصفة","الكمية","الوحدة","سعر الوحدة","الإجمالي","المورد","ملاحظات"],
      ...ci.map((it,i)=>{const t=(parseFloat(it.qty)||0)*(parseFloat(it.unitPrice||0)||0);return[i+1,it.name,it.spec||"",it.qty||"",it.unit,it.unitPrice||"",t||"",it.supplier||"",it.notes||""];})
    ]);
    cws["!cols"]=[{wch:4},{wch:30},{wch:40},{wch:8},{wch:8},{wch:12},{wch:14},{wch:22},{wch:18}];
    XLSX.utils.book_append_sheet(wb, cws, cat.replace(/[\\/*?:[\]]/g,"").slice(0,25));
  });
  const out = XLSX.write(wb, {bookType:"xlsx",type:"base64"});
  const a = document.createElement("a");
  a.href = "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,"+out;
  a.download = `BOQ_${projectName||"مشروع"}_${new Date().toISOString().slice(0,10)}.xlsx`;
  a.click();
}

// ─── Export PDF ────────────────────────────────────────────────────────────────
function dlPDF(projectName, items, companyName) {
  const cn = companyName || "ركيز ديزاين";
  const rows = items.map((it,i)=>{
    const t=(parseFloat(it.qty)||0)*(parseFloat(it.unitPrice||0)||0);
    return `<tr style="background:${i%2?"#fff":"#F9F7F4"}"><td>${i+1}</td><td><span style="background:#F5ECD8;color:#8A6530;padding:2px 6px;border-radius:8px;font-size:9px;font-weight:600">${it.category}</span></td><td style="font-weight:600">${it.name}</td><td style="color:#6B6560;font-size:10px">${it.spec||""}</td><td style="text-align:center">${it.qty||""}</td><td style="text-align:center">${it.unit}</td><td style="text-align:center">${it.unitPrice?Number(it.unitPrice).toLocaleString("ar-SA"):""}</td><td style="text-align:center;font-weight:700;color:#3D7A5F">${t?t.toLocaleString("ar-SA"):""}</td><td style="font-size:9px">${it.supplier||""}</td></tr>`;
  }).join("");
  const grand = items.reduce((s,it)=>s+(parseFloat(it.qty)||0)*(parseFloat(it.unitPrice||0)||0),0);
  const html = `<!DOCTYPE html><html dir="rtl"><head><meta charset="UTF-8"><link href="https://fonts.googleapis.com/css2?family=Noto+Kufi+Arabic:wght@400;600;700&display=swap" rel="stylesheet"><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:'Noto Kufi Arabic',sans-serif;color:#1A1815}@page{size:A3 landscape;margin:12mm}.cover{background:#1A1815;color:#fff;padding:18px 24px;margin-bottom:14px;display:flex;justify-content:space-between;align-items:center}.cover h1{font-size:17px;color:#C8964E;margin-bottom:3px}.meta{font-size:10px;color:#999}table{width:100%;border-collapse:collapse;font-size:10.5px}th{background:#2A2520;color:#F5ECD8;padding:7px 8px;text-align:right;font-weight:600;border:1px solid #3A342E}td{padding:6px 8px;border-bottom:1px solid #E5E0D8;text-align:right;vertical-align:top}tfoot td{background:#F5ECD8;font-weight:700;color:#8A6530;border-top:2px solid #B5894A}</style></head><body><div class="cover"><div><h1>جدول الكميات — Bill of Quantities</h1><div class="meta">${cn} | ${projectName||""} | ${new Date().toLocaleDateString("ar-SA")} | ${items.length} بند</div></div></div><table><thead><tr><th>#</th><th>الفئة</th><th>اسم البند</th><th>المواصفة</th><th>الكمية</th><th>الوحدة</th><th>سعر الوحدة</th><th>الإجمالي</th><th>المورد</th></tr></thead><tbody>${rows}</tbody><tfoot><tr><td colspan="7">المجموع الكلي</td><td style="text-align:center">${grand?grand.toLocaleString("ar-SA"):""} SAR</td><td></td></tr></tfoot></table><script>setTimeout(()=>window.print(),600)<\/script></body></html>`;
  const a = document.createElement("a");
  a.href = "data:text/html;charset=utf-8," + encodeURIComponent(html);
  a.download = `BOQ_${projectName||"مشروع"}_${new Date().toISOString().slice(0,10)}.html`;
  a.click();
}

// ─── Shared UI ─────────────────────────────────────────────────────────────────
const Badge = ({ status, C }) => {
  const m = { "جاري":{bg:C.blueLight,c:C.blue}, "مكتمل":{bg:C.greenLight,c:C.green}, "معلق":{bg:C.accentLight,c:C.accentDark} };
  const s = m[status]||{bg:C.surfaceAlt,c:C.textSub};
  return <span style={{background:s.bg,color:s.c,borderRadius:20,padding:"3px 12px",fontSize:12,fontWeight:600}}>{status}</span>;
};

const Card = ({ children, style={}, onClick, C }) => (
  <div onClick={onClick} style={{background:C.surface,borderRadius:12,border:`1px solid ${C.border}`,padding:20,transition:"box-shadow .2s,transform .15s",cursor:onClick?"pointer":"default",...style}}
    onMouseEnter={e=>{if(onClick){e.currentTarget.style.boxShadow="0 4px 20px rgba(0,0,0,.08)";e.currentTarget.style.transform="translateY(-1px)";}}}
    onMouseLeave={e=>{e.currentTarget.style.boxShadow="none";e.currentTarget.style.transform="translateY(0)";}}>
    {children}
  </div>
);

const Btn = ({ children, variant="primary", onClick, style={}, disabled, sm, C }) => {
  const sz = sm ? {padding:"6px 13px",fontSize:12} : {padding:"9px 20px",fontSize:13};
  const vs = {
    primary:{background:C.accent,color:"#fff",border:"none"},
    secondary:{background:C.surfaceAlt,color:C.text,border:`1px solid ${C.border}`},
    ghost:{background:"transparent",color:C.textSub,border:"none"},
    danger:{background:C.redLight,color:C.red,border:"none"},
    success:{background:C.greenLight,color:C.green,border:"none"},
  };
  return <button style={{...sz,...(vs[variant]||vs.primary),borderRadius:8,fontWeight:600,fontFamily:FONT,cursor:disabled?"not-allowed":"pointer",opacity:disabled?.5:1,display:"inline-flex",alignItems:"center",gap:6,transition:"all .15s",...style}}
    onClick={onClick} disabled={disabled}
    onMouseEnter={e=>{if(!disabled)e.currentTarget.style.opacity=".82";}}
    onMouseLeave={e=>{e.currentTarget.style.opacity="1";}}>
    {children}
  </button>;
};

const Input = ({ label, value, onChange, placeholder, type="text", style={}, C }) => (
  <div style={{display:"flex",flexDirection:"column",gap:5}}>
    {label&&<label style={{fontSize:12,fontWeight:600,color:C.textSub}}>{label}</label>}
    <input type={type} value={value||""} onChange={e=>onChange(e.target.value)} placeholder={placeholder}
      style={{padding:"9px 12px",borderRadius:8,border:`1px solid ${C.border}`,background:C.surface,color:C.text,fontSize:13,fontFamily:FONT,outline:"none",...style}}
      onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor=C.border}/>
  </div>
);

const Textarea = ({ label, value, onChange, placeholder, rows=3, C }) => (
  <div style={{display:"flex",flexDirection:"column",gap:5}}>
    {label&&<label style={{fontSize:12,fontWeight:600,color:C.textSub}}>{label}</label>}
    <textarea value={value||""} onChange={e=>onChange(e.target.value)} rows={rows} placeholder={placeholder}
      style={{padding:"9px 12px",borderRadius:8,border:`1px solid ${C.border}`,background:C.surface,color:C.text,fontSize:13,fontFamily:FONT,resize:"vertical",outline:"none"}}
      onFocus={e=>e.target.style.borderColor=C.accent} onBlur={e=>e.target.style.borderColor=C.border}/>
  </div>
);

const Select = ({ label, value, onChange, options, C }) => (
  <div style={{display:"flex",flexDirection:"column",gap:5}}>
    {label&&<label style={{fontSize:12,fontWeight:600,color:C.textSub}}>{label}</label>}
    <select value={value||""} onChange={e=>onChange(e.target.value)}
      style={{padding:"9px 12px",borderRadius:8,border:`1px solid ${C.border}`,background:C.surface,color:C.text,fontSize:13,fontFamily:FONT,outline:"none"}}>
      {options.map(o=><option key={o.value||o} value={o.value||o}>{o.label||o}</option>)}
    </select>
  </div>
);

const Modal = ({ open, onClose, title, children, width=540, C }) => {
  if(!open) return null;
  return (
    <div onClick={onClose} style={{position:"fixed",inset:0,zIndex:1000,background:"rgba(0,0,0,.42)",display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div onClick={e=>e.stopPropagation()} style={{background:C.surface,borderRadius:16,width:"100%",maxWidth:width,maxHeight:"90vh",overflow:"auto",boxShadow:"0 20px 60px rgba(0,0,0,.22)"}}>
        <div style={{padding:"18px 24px",borderBottom:`1px solid ${C.border}`,display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,background:C.surface,zIndex:1}}>
          <h3 style={{fontSize:16,fontWeight:700,color:C.text}}>{title}</h3>
          <button onClick={onClose} style={{background:"none",border:"none",fontSize:20,color:C.textSub,cursor:"pointer",lineHeight:1}}>×</button>
        </div>
        <div style={{padding:24}}>{children}</div>
      </div>
    </div>
  );
};

// ─── Image Annotator ───────────────────────────────────────────────────────────
const ANNO_COLORS = ["#E53935","#FFA726","#43A047","#1E88E5","#8E24AA","#ffffff","#000000"];

const ImageAnnotator = ({ src, onSave, onClose }) => {
  const canvasRef = useRef();
  const imgRef = useRef(null);
  const [tool, setTool] = useState("rect");
  const [color, setColor] = useState("#E53935");
  const [shapes, setShapes] = useState([]);
  const [drawing, setDrawing] = useState(false);
  const [start, setStart] = useState(null);
  const [loaded, setLoaded] = useState(false);

  useEffect(()=>{
    const img = new Image();
    img.onload = ()=>{
      const cv = canvasRef.current; if(!cv) return;
      const MAX = Math.min(window.innerWidth*.82, 1100);
      const scale = Math.min(1, MAX/img.width);
      cv.width = img.width*scale; cv.height = img.height*scale;
      imgRef.current = img; setLoaded(true);
    };
    img.src = src;
  },[src]);

  useEffect(()=>{
    const cv = canvasRef.current; if(!cv||!imgRef.current) return;
    const ctx = cv.getContext("2d");
    ctx.clearRect(0,0,cv.width,cv.height);
    ctx.drawImage(imgRef.current,0,0,cv.width,cv.height);
    shapes.forEach(s=>{
      ctx.strokeStyle=s.color; ctx.fillStyle=s.color; ctx.lineWidth=3;
      if(s.type==="rect") ctx.strokeRect(s.x,s.y,s.w,s.h);
      if(s.type==="arrow"){
        ctx.beginPath(); ctx.moveTo(s.x1,s.y1); ctx.lineTo(s.x2,s.y2); ctx.stroke();
        const a=Math.atan2(s.y2-s.y1,s.x2-s.x1);
        [a-.4,a+.4].forEach(ang=>{ctx.beginPath();ctx.moveTo(s.x2,s.y2);ctx.lineTo(s.x2-14*Math.cos(ang),s.y2-14*Math.sin(ang));ctx.stroke();});
      }
    });
  },[shapes,loaded]);

  const gp = e=>{ const r=canvasRef.current.getBoundingClientRect(); return {x:e.clientX-r.left,y:e.clientY-r.top}; };
  const onDown=e=>{ setStart(gp(e)); setDrawing(true); };
  const onMove=e=>{
    if(!drawing||!start) return;
    const p=gp(e);
    if(tool==="rect") setShapes(s=>[...s.filter(x=>!x._p),{type:"rect",x:Math.min(start.x,p.x),y:Math.min(start.y,p.y),w:Math.abs(p.x-start.x),h:Math.abs(p.y-start.y),color,_p:true}]);
    if(tool==="arrow") setShapes(s=>[...s.filter(x=>!x._p),{type:"arrow",x1:start.x,y1:start.y,x2:p.x,y2:p.y,color,_p:true}]);
  };
  const onUp=()=>{ setDrawing(false); setShapes(s=>s.map(x=>({...x,_p:false}))); };
  const save=()=>{ onSave(canvasRef.current.toDataURL("image/jpeg",.92)); onClose(); };

  return (
    <div style={{position:"fixed",inset:0,zIndex:3000,background:"rgba(0,0,0,.93)",display:"flex",flexDirection:"column",alignItems:"center",padding:"10px 14px",gap:8}}>
      <div style={{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap",justifyContent:"center"}}>
        {[["rect","⬜","مستطيل"],["arrow","➡","سهم"]].map(([id,ic,lb])=>
          <button key={id} onClick={()=>setTool(id)} style={{padding:"6px 12px",borderRadius:7,border:`2px solid ${tool===id?"#B5894A":"transparent"}`,background:tool===id?"rgba(181,137,74,.2)":"rgba(255,255,255,.06)",color:"#fff",cursor:"pointer",fontFamily:FONT,fontSize:12,display:"flex",alignItems:"center",gap:4}}>{ic} {lb}</button>
        )}
        <div style={{width:1,height:26,background:"rgba(255,255,255,.2)"}}/>
        {ANNO_COLORS.map(col=><button key={col} onClick={()=>setColor(col)} style={{width:22,height:22,borderRadius:"50%",background:col,border:`3px solid ${color===col?"#B5894A":"transparent"}`,cursor:"pointer"}}/>)}
        <div style={{width:1,height:26,background:"rgba(255,255,255,.2)"}}/>
        <button onClick={()=>setShapes(s=>s.slice(0,-1))} style={{padding:"5px 11px",borderRadius:7,border:"1px solid rgba(255,255,255,.2)",background:"transparent",color:"#fff",cursor:"pointer",fontFamily:FONT,fontSize:12}}>↩ تراجع</button>
        <button onClick={onClose} style={{padding:"5px 11px",borderRadius:7,border:"1px solid rgba(255,255,255,.2)",background:"transparent",color:"#fff",cursor:"pointer",fontFamily:FONT,fontSize:12}}>✕ إغلاق</button>
        <button onClick={save} style={{padding:"6px 16px",borderRadius:8,border:"none",background:"#B5894A",color:"#fff",cursor:"pointer",fontFamily:FONT,fontSize:13,fontWeight:700}}>💾 حفظ</button>
      </div>
      <div style={{fontSize:10,color:"#888"}}>{tool==="rect"?"ارسم مستطيل تحديد":"ارسم سهم للإشارة"}</div>
      <div style={{overflow:"auto",maxHeight:"calc(100vh - 110px)"}}>
        <canvas ref={canvasRef} onMouseDown={onDown} onMouseMove={onMove} onMouseUp={onUp}
          style={{cursor:"crosshair",borderRadius:8,display:"block",maxWidth:"100%"}}/>
      </div>
    </div>
  );
};

// ─── Image Cell (single image per row) ─────────────────────────────────────────
const ImgCell = ({ image, onChange, C }) => {
  const r = useRef();
  const [ann, setAnn] = useState(false);
  const proc = files => {
    const f=files[0]; if(!f||!f.type.startsWith("image/")) return;
    const rd=new FileReader(); rd.onload=ev=>onChange(ev.target.result); rd.readAsDataURL(f);
  };
  useEffect(()=>{
    const h=e=>{ if(e.clipboardData?.files?.length) proc(e.clipboardData.files); };
    document.addEventListener("paste",h); return ()=>document.removeEventListener("paste",h);
  },[]);
  return (
    <>
      {image ? (
        <div style={{position:"relative",width:80,height:80,borderRadius:8,overflow:"hidden",border:`1.5px solid ${C.border}`,boxShadow:"0 2px 8px rgba(0,0,0,.1)"}}>
          <img src={image} alt="" onClick={()=>setAnn(true)} title="اضغط لتحرير الصورة"
            style={{width:"100%",height:"100%",objectFit:"cover",cursor:"pointer",display:"block"}}/>
          <div style={{position:"absolute",bottom:0,left:0,right:0,background:"rgba(0,0,0,.5)",display:"flex",justifyContent:"space-between",padding:"2px 5px"}}>
            <span onClick={()=>setAnn(true)} style={{color:"#fff",fontSize:9,cursor:"pointer"}}>✏ تحرير</span>
            <button onClick={()=>onChange(null)} style={{background:"none",border:"none",color:"#fff",fontSize:11,cursor:"pointer",lineHeight:1}}>×</button>
          </div>
        </div>
      ) : (
        <button onClick={()=>r.current.click()} title="+ صورة — أو Ctrl+V — أو اسحب"
          style={{width:80,height:80,borderRadius:8,border:`1.5px dashed ${C.accent}`,background:C.accentLight,color:C.accentDark,fontSize:22,cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:1}}>
          <span>+</span><span style={{fontSize:9,opacity:.7}}>صورة</span>
        </button>
      )}
      <input ref={r} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{proc(e.target.files);e.target.value="";}}/>
      {ann && image && <ImageAnnotator src={image} onClose={()=>setAnn(false)} onSave={d=>{onChange(d);setAnn(false);}}/>}
    </>
  );
};

// ─── Inline Cell Editor ────────────────────────────────────────────────────────
const IC = ({ value, onChange, placeholder="—", type="text", options, multiline, style={}, C }) => {
  const [ed, setEd] = useState(false);
  const [v, setV] = useState(value);
  const r = useRef();
  useEffect(()=>setV(value),[value]);
  useEffect(()=>{ if(ed&&r.current){ r.current.focus(); try{r.current.select();}catch(e){} } },[ed]);
  const commit = ()=>{ setEd(false); if(v!==value) onChange(v); };
  const onKey = e=>{
    if(e.key==="Enter"&&!multiline){ e.preventDefault(); commit(); }
    if(e.key==="Escape"){ setV(value); setEd(false); }
    if(e.key==="Tab"){ e.preventDefault(); commit(); }
  };
  if(!ed) return (
    <span onClick={e=>{e.stopPropagation();setEd(true);}} title="اضغط للتعديل"
      style={{cursor:"text",borderRadius:4,padding:"3px 6px",display:"block",lineHeight:1.5,minHeight:22,...style}}
      onMouseEnter={e=>e.currentTarget.style.background="rgba(181,137,74,.12)"}
      onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
      {value||<span style={{opacity:.3}}>{placeholder}</span>}
    </span>
  );
  const s = {width:"100%",padding:"3px 7px",borderRadius:5,border:`1.5px solid ${C.accent}`,fontSize:"inherit",fontFamily:FONT,outline:"none",background:C.surface,color:C.text,...style};
  if(options) return <select ref={r} value={v} onChange={e=>setV(e.target.value)} onBlur={commit} onKeyDown={onKey} style={s}>{options.map(o=><option key={o}>{o}</option>)}</select>;
  if(multiline) return <textarea ref={r} value={v} onChange={e=>setV(e.target.value)} onBlur={commit} onKeyDown={onKey} rows={2} style={{...s,resize:"vertical"}}/>;
  return <input ref={r} type={type} value={v} onChange={e=>setV(e.target.value)} onBlur={commit} onKeyDown={onKey} placeholder={placeholder} style={s}/>;
};

// ─── Quick Add Row ─────────────────────────────────────────────────────────────
const QAR = ({ onAdd, lastCat, lastUnit, C }) => {
  const blank = {category:lastCat,name:"",spec:"",qty:"",unit:lastUnit,unitPrice:"",supplier:"",notes:"",image:null};
  const [f, setF] = useState(blank);
  const nr = useRef();
  const commit = (andNew=false) => {
    if(!f.name.trim()) return;
    onAdd({...f, id:Date.now()});
    if(andNew){ setF({...blank,category:f.category,unit:f.unit}); setTimeout(()=>nr.current?.focus(),40); }
    else setF(blank);
  };
  const onKey = e=>{ if(e.key==="Enter"){e.preventDefault();commit(true);} if(e.key==="Escape")setF(blank); };
  const td = {padding:"6px 9px",borderTop:`2px solid ${C.accent}`,background:C.accentLight+"99"};
  const is = {background:"transparent",border:"none",outline:"none",width:"100%",fontSize:12,fontFamily:FONT,color:C.text};
  return (
    <tr onKeyDown={onKey}>
      <td style={{...td,textAlign:"center",color:C.accentDark,fontWeight:700,fontSize:16}}>+</td>
      <td style={td}><select value={f.category} onChange={e=>setF(p=>({...p,category:e.target.value}))} style={{...is,fontSize:11}}>{CATEGORIES.map(c=><option key={c}>{c}</option>)}</select></td>
      <td style={td} colSpan={2}><input ref={nr} value={f.name} onChange={e=>setF(p=>({...p,name:e.target.value}))} placeholder="اسم البند ← Enter للإضافة الفورية" style={{...is,fontWeight:600}} autoFocus/></td>
      <td style={td}><input type="number" value={f.qty} onChange={e=>setF(p=>({...p,qty:e.target.value}))} placeholder="0" style={{...is,textAlign:"center"}}/></td>
      <td style={td}><select value={f.unit} onChange={e=>setF(p=>({...p,unit:e.target.value}))} style={{...is,fontSize:11}}>{UNITS.map(u=><option key={u}>{u}</option>)}</select></td>
      <td style={td}><input type="number" value={f.unitPrice} onChange={e=>setF(p=>({...p,unitPrice:e.target.value}))} placeholder="السعر" style={{...is,textAlign:"center"}}/></td>
      <td style={td}><div style={{display:"flex",gap:5}}><Btn variant="primary" sm onClick={()=>commit(true)} C={C}>+ إضافة</Btn><Btn variant="ghost" sm onClick={()=>setF(blank)} C={C}>مسح</Btn></div></td>
    </tr>
  );
};

// ─── Dashboard ──────────────────────────────────────────────────────────────────
const DashboardPage = ({ projects, onSelectProject, C }) => {
  const total = { items:projects.reduce((s,p)=>s+p.items,0), specs:projects.reduce((s,p)=>s+p.specs,0), images:projects.reduce((s,p)=>s+p.images,0) };
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{marginBottom:24}}>
        <h1 style={{fontSize:22,fontWeight:700,color:C.text}}>لوحة التحكم</h1>
        <p style={{color:C.textSub,marginTop:4,fontSize:14}}>نظرة عامة على مشاريعك وبياناتك</p>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(185px,1fr))",gap:13,marginBottom:28}}>
        {[["إجمالي المشاريع",projects.length,"🏗️",C.accent],["إجمالي البنود",total.items,"📋",C.blue],["المواصفات",total.specs,"📄",C.green],["الصور",total.images,"🖼️","#9B59B6"]].map(([l,v,ic,col])=>(
          <Card key={l} C={C} style={{display:"flex",alignItems:"center",gap:13,padding:"16px 18px"}}>
            <div style={{width:42,height:42,borderRadius:10,background:col+"22",display:"flex",alignItems:"center",justifyContent:"center",fontSize:19}}>{ic}</div>
            <div><div style={{fontSize:22,fontWeight:700,color:C.text}}>{v}</div><div style={{fontSize:12,color:C.textSub,marginTop:2}}>{l}</div></div>
          </Card>
        ))}
      </div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
        <h2 style={{fontSize:16,fontWeight:700,color:C.text}}>المشاريع الأخيرة</h2>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:14}}>
        {projects.map(p=>(
          <Card key={p.id} onClick={()=>onSelectProject(p)} C={C}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
              <div><div style={{fontWeight:700,fontSize:15,marginBottom:4,color:C.text}}>{p.name}</div><div style={{fontSize:12,color:C.textSub}}>{p.client}</div></div>
              <Badge status={p.status} C={C}/>
            </div>
            <div style={{fontSize:12,color:C.textSub,marginBottom:12}}>📍 {p.location}</div>
            <div style={{display:"flex",gap:14,paddingTop:12,borderTop:`1px solid ${C.border}`}}>
              {[["بند",p.items,C.blue],["مواصفة",p.specs,C.green],["صورة",p.images,"#9B59B6"]].map(([l,v,col])=>(
                <div key={l} style={{textAlign:"center"}}><div style={{fontWeight:700,color:col,fontSize:16}}>{v}</div><div style={{fontSize:11,color:C.textMuted}}>{l}</div></div>
              ))}
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
};

// ─── Projects ───────────────────────────────────────────────────────────────────
const ProjectsPage = ({ projects, setProjects, onSelectProject, C }) => {
  const [showModal, setShowModal] = useState(false);
  const [search, setSearch] = useState("");
  const [filterStatus, setFilterStatus] = useState("الكل");
  const [form, setForm] = useState({name:"",client:"",location:"",status:"جاري"});

  const filtered = projects.filter(p=>
    (filterStatus==="الكل"||p.status===filterStatus)&&
    (p.name.includes(search)||p.client.includes(search)||p.location.includes(search))
  );
  const add = ()=>{
    if(!form.name) return;
    setProjects([...projects,{...form,id:Date.now(),date:new Date().toISOString().split("T")[0],items:0,specs:0,images:0}]);
    setForm({name:"",client:"",location:"",status:"جاري"}); setShowModal(false);
  };
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
        <div><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>المشاريع</h2><p style={{fontSize:13,color:C.textSub,marginTop:2}}>{filtered.length} مشروع</p></div>
        <Btn onClick={()=>setShowModal(true)} C={C}>+ مشروع جديد</Btn>
      </div>
      <div style={{display:"flex",gap:10,marginBottom:18,flexWrap:"wrap"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 بحث في المشاريع..."
          style={{flex:1,minWidth:200,padding:"9px 14px",borderRadius:8,border:`1px solid ${C.border}`,fontSize:13,fontFamily:FONT,background:C.surface,color:C.text,outline:"none"}}/>
        {["الكل","جاري","مكتمل","معلق"].map(s=>(
          <button key={s} onClick={()=>setFilterStatus(s)} style={{padding:"8px 16px",borderRadius:20,border:`1px solid ${filterStatus===s?C.accent:C.border}`,background:filterStatus===s?C.accentLight:C.surface,color:filterStatus===s?C.accentDark:C.textSub,fontFamily:FONT,fontSize:13,fontWeight:600,cursor:"pointer"}}>{s}</button>
        ))}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:14}}>
        {filtered.map(p=>(
          <Card key={p.id} onClick={()=>onSelectProject(p)} C={C}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:12}}>
              <div><div style={{fontWeight:700,fontSize:15,marginBottom:4,color:C.text}}>{p.name}</div><div style={{fontSize:12,color:C.textSub}}>👤 {p.client}</div></div>
              <Badge status={p.status} C={C}/>
            </div>
            <div style={{fontSize:12,color:C.textSub,marginBottom:4}}>📍 {p.location}</div>
            <div style={{fontSize:12,color:C.textMuted,marginBottom:14}}>📅 {p.date}</div>
            <div style={{display:"flex",gap:8,paddingTop:12,borderTop:`1px solid ${C.border}`}}>
              {[["📋",p.items,"بند"],["📄",p.specs,"مواصفة"],["🖼️",p.images,"صورة"]].map(([ic,v,l])=>(
                <div key={l} style={{flex:1,textAlign:"center",background:C.surfaceAlt,borderRadius:8,padding:"7px 4px"}}>
                  <div style={{fontSize:11}}>{ic}</div><div style={{fontWeight:700,fontSize:15,color:C.text}}>{v}</div><div style={{fontSize:10,color:C.textMuted}}>{l}</div>
                </div>
              ))}
            </div>
          </Card>
        ))}
      </div>
      <Modal open={showModal} onClose={()=>setShowModal(false)} title="مشروع جديد" C={C}>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <Input label="اسم المشروع *" value={form.name} onChange={v=>setForm({...form,name:v})} placeholder="مثال: فيلا الأستاذ عبدالله..." C={C}/>
          <Input label="اسم العميل" value={form.client} onChange={v=>setForm({...form,client:v})} placeholder="اسم العميل" C={C}/>
          <Input label="الموقع" value={form.location} onChange={v=>setForm({...form,location:v})} placeholder="المدينة والحي" C={C}/>
          <Select label="حالة المشروع" value={form.status} onChange={v=>setForm({...form,status:v})} options={["جاري","مكتمل","معلق"]} C={C}/>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:8}}>
            <Btn variant="secondary" onClick={()=>setShowModal(false)} C={C}>إلغاء</Btn>
            <Btn onClick={add} disabled={!form.name} C={C}>إنشاء المشروع</Btn>
          </div>
        </div>
      </Modal>
    </div>
  );
};

// ─── BOQ Page ───────────────────────────────────────────────────────────────────
const BOQPage = ({ project, companyName, C }) => {
  const [items, setItems] = useState(mockItems);
  const [search, setSearch] = useState("");
  const [filterCat, setFilterCat] = useState("الكل");
  const [sr, setSr] = useState({find:"",replace:"",active:false});
  const [sel, setSel] = useState(new Set());
  const [hist, setHist] = useState([mockItems]);
  const [hIdx, setHIdx] = useState(0);
  const [editItem, setEditItem] = useState(null);
  const [showEdit, setShowEdit] = useState(false);

  const push = next=>{ const h=hist.slice(0,hIdx+1); setHist([...h,next]); setHIdx(h.length); setItems(next); };
  const undo = ()=>{ if(hIdx>0){ setHIdx(i=>i-1); setItems(hist[hIdx-1]); } };
  const redo = ()=>{ if(hIdx<hist.length-1){ setHIdx(i=>i+1); setItems(hist[hIdx+1]); } };

  useEffect(()=>{
    const h=e=>{
      if((e.ctrlKey||e.metaKey)&&e.key==="z"&&!e.shiftKey){ e.preventDefault(); undo(); }
      if((e.ctrlKey||e.metaKey)&&(e.key==="y"||(e.key==="z"&&e.shiftKey))){ e.preventDefault(); redo(); }
    };
    window.addEventListener("keydown",h); return ()=>window.removeEventListener("keydown",h);
  },[hIdx,hist]);

  const upd = (id,f,v)=>push(items.map(it=>it.id===id?{...it,[f]:v}:it));
  const addItem = item=>push([...items,item]);
  const del = id=>{ push(items.filter(it=>it.id!==id)); setSel(s=>{const n=new Set(s);n.delete(id);return n;}); };
  const dup = item=>push([...items,{...item,id:Date.now(),name:item.name+" (نسخة)"}]);
  const delSel = ()=>{ push(items.filter(it=>!sel.has(it.id))); setSel(new Set()); };
  const doRepl = ()=>{ if(!sr.find)return; push(items.map(it=>({...it,name:it.name.replaceAll(sr.find,sr.replace),spec:(it.spec||"").replaceAll(sr.find,sr.replace)}))); setSr({find:"",replace:"",active:false}); };

  const cats = ["الكل",...new Set(items.map(i=>i.category))];
  const filtered = items.filter(it=>(filterCat==="الكل"||it.category===filterCat)&&(!search||it.name.includes(search)||(it.spec||"").includes(search)||it.category.includes(search)));
  const grand = items.reduce((s,it)=>s+(parseFloat(it.qty)||0)*(parseFloat(it.unitPrice||0)||0),0);

  const TH = {padding:"10px 12px",textAlign:"right",fontWeight:600,fontSize:12,color:C.textSub,borderBottom:`1px solid ${C.border}`,whiteSpace:"nowrap",background:C.surfaceAlt};
  const TD = (i,x={})=>({padding:"10px 11px",borderBottom:`1px solid ${C.border}`,background:i%2===0?C.surface:C.bg||C.surfaceAlt,fontSize:12,verticalAlign:"middle",...x});

  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      {/* Header + Export top-left */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16,flexWrap:"wrap",gap:10}}>
        <div>
          <h2 style={{fontSize:18,fontWeight:700,color:C.text}}>جدول الكميات — {project?.name||"المشروع"}</h2>
          <p style={{fontSize:13,color:C.textSub,marginTop:2}}>
            {filtered.length} بند
            {grand>0&&<span style={{marginRight:10,color:C.accentDark,fontWeight:700}}>{grand.toLocaleString("ar-SA")} SAR</span>}
          </p>
        </div>
        <div style={{display:"flex",gap:8}}>
          <Btn variant="success" sm onClick={()=>dlExcel(project?.name,items,companyName)} C={C}>⬇ Excel</Btn>
          <Btn variant="secondary" sm onClick={()=>dlPDF(project?.name,items,companyName)} C={C}>🖨 PDF</Btn>
        </div>
      </div>

      {/* Toolbar */}
      <div style={{display:"flex",gap:8,marginBottom:12,flexWrap:"wrap",alignItems:"center"}}>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="🔍 بحث في البنود..."
          style={{flex:1,minWidth:160,padding:"8px 13px",borderRadius:8,border:`1px solid ${C.border}`,fontSize:13,fontFamily:FONT,background:C.surface,color:C.text,outline:"none"}}/>
        {cats.map(c=><button key={c} onClick={()=>setFilterCat(c)} style={{padding:"6px 12px",borderRadius:18,border:`1px solid ${filterCat===c?C.accent:C.border}`,background:filterCat===c?C.accentLight:C.surface,color:filterCat===c?C.accentDark:C.textSub,fontFamily:FONT,fontSize:12,fontWeight:600,cursor:"pointer"}}>{c}</button>)}
        <div style={{marginRight:"auto",display:"flex",gap:6}}>
          <Btn variant="ghost" sm onClick={undo} disabled={hIdx<=0} C={C} title="Ctrl+Z">↩</Btn>
          <Btn variant="ghost" sm onClick={redo} disabled={hIdx>=hist.length-1} C={C} title="Ctrl+Y">↪</Btn>
          <Btn variant="secondary" sm onClick={()=>setSr(s=>({...s,active:!s.active}))} C={C}>🔄 استبدال</Btn>
          {sel.size>0&&<Btn variant="danger" sm onClick={delSel} C={C}>🗑 حذف ({sel.size})</Btn>}
        </div>
      </div>

      {/* Search/Replace */}
      {sr.active&&(
        <div style={{background:C.accentLight,borderRadius:10,border:`1px solid ${C.accent}`,padding:"12px 16px",marginBottom:12,display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
          <Input label="ابحث عن" value={sr.find} onChange={v=>setSr(s=>({...s,find:v}))} placeholder="النص القديم" style={{minWidth:150}} C={C}/>
          <Input label="استبدل بـ" value={sr.replace} onChange={v=>setSr(s=>({...s,replace:v}))} placeholder="النص الجديد" style={{minWidth:150}} C={C}/>
          <Btn variant="primary" sm onClick={doRepl} C={C}>تطبيق</Btn>
          <Btn variant="ghost" sm onClick={()=>setSr({find:"",replace:"",active:false})} C={C}>إلغاء</Btn>
        </div>
      )}

      {/* Table */}
      <Card C={C} style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead>
              <tr>
                <th style={{...TH,width:30,textAlign:"center"}}><input type="checkbox" checked={sel.size===filtered.length&&filtered.length>0} onChange={e=>setSel(e.target.checked?new Set(filtered.map(it=>it.id)):new Set())}/></th>
                <th style={{...TH,width:32}}>#</th>
                <th style={{...TH,width:100}}>الفئة</th>
                <th style={TH}>اسم البند</th>
                <th style={{...TH,maxWidth:250}}>المواصفة</th>
                <th style={{...TH,width:65,textAlign:"center"}}>الكمية</th>
                <th style={{...TH,width:65,textAlign:"center"}}>الوحدة</th>
                <th style={{...TH,width:80,textAlign:"center"}}>سعر الوحدة</th>
                <th style={{...TH,width:90,textAlign:"center"}}>الإجمالي</th>
                <th style={{...TH,width:96}}>صورة</th>
                <th style={{...TH,width:90}}>إجراءات</th>
              </tr>
            </thead>
            <tbody>
              {filtered.map((item,i)=>{
                const total=(parseFloat(item.qty)||0)*(parseFloat(item.unitPrice||0)||0);
                const isSel=sel.has(item.id);
                const rb=isSel?C.accentLight:(i%2===0?C.surface:"#FAFAF8");
                return (
                  <tr key={item.id} style={{background:rb}}>
                    <td style={TD(i,{background:rb,textAlign:"center"})}><input type="checkbox" checked={isSel} onChange={e=>{const s=new Set(sel);e.target.checked?s.add(item.id):s.delete(item.id);setSel(s);}}/></td>
                    <td style={TD(i,{background:rb,textAlign:"center",fontWeight:700,color:C.accentDark,fontSize:11})}>{i+1}</td>
                    <td style={TD(i,{background:rb})}><IC value={item.category} onChange={v=>upd(item.id,"category",v)} options={CATEGORIES} C={C}/></td>
                    <td style={TD(i,{background:rb,fontWeight:600})}><IC value={item.name} onChange={v=>upd(item.id,"name",v)} placeholder="اسم البند" C={C}/></td>
                    <td style={TD(i,{background:rb,maxWidth:250})}><IC value={item.spec||""} onChange={v=>upd(item.id,"spec",v)} multiline placeholder="المواصفة..." style={{fontSize:11,color:C.textSub}} C={C}/></td>
                    <td style={TD(i,{background:rb,textAlign:"center"})}><IC value={String(item.qty||"")} onChange={v=>upd(item.id,"qty",v)} type="number" placeholder="0" style={{textAlign:"center",fontWeight:700,color:C.accent}} C={C}/></td>
                    <td style={TD(i,{background:rb,textAlign:"center"})}><IC value={item.unit} onChange={v=>upd(item.id,"unit",v)} options={UNITS} C={C}/></td>
                    <td style={TD(i,{background:rb,textAlign:"center"})}><IC value={String(item.unitPrice||"")} onChange={v=>upd(item.id,"unitPrice",v)} type="number" placeholder="—" style={{textAlign:"center"}} C={C}/></td>
                    <td style={TD(i,{background:rb,textAlign:"center",fontWeight:700,color:total>0?C.green:C.textMuted})}>{total>0?total.toLocaleString("ar-SA"):"—"}</td>
                    <td style={TD(i,{background:rb,padding:"5px 8px"})}><ImgCell image={item.image} onChange={v=>upd(item.id,"image",v)} C={C}/></td>
                    <td style={TD(i,{background:rb})}>
                      <div style={{display:"flex",gap:4}}>
                        <button onClick={()=>{setEditItem({...item});setShowEdit(true);}} style={{padding:"4px 8px",borderRadius:5,border:`1px solid ${C.border}`,background:"none",cursor:"pointer",fontSize:11,color:C.blue}}>✏</button>
                        <button onClick={()=>dup(item)} style={{padding:"4px 8px",borderRadius:5,border:`1px solid ${C.border}`,background:"none",cursor:"pointer",fontSize:11,color:C.textSub}}>⧉</button>
                        <button onClick={()=>del(item.id)} style={{padding:"4px 8px",borderRadius:5,border:"none",background:C.redLight,cursor:"pointer",fontSize:11,color:C.red}}>✕</button>
                      </div>
                    </td>
                  </tr>
                );
              })}
              <QAR onAdd={addItem} lastCat={items[items.length-1]?.category||"دهانات"} lastUnit={items[items.length-1]?.unit||"م²"} C={C}/>
            </tbody>
            {grand>0&&(
              <tfoot>
                <tr style={{background:C.accentLight}}>
                  <td colSpan={8} style={{padding:"10px 14px",fontWeight:700,color:C.accentDark,fontSize:13}}>المجموع الكلي</td>
                  <td style={{padding:"10px 14px",fontWeight:800,color:C.green,fontSize:14,textAlign:"center"}}>{grand.toLocaleString("ar-SA")} SAR</td>
                  <td colSpan={2}/>
                </tr>
              </tfoot>
            )}
          </table>
        </div>
      </Card>

      {/* Edit Modal */}
      <Modal open={showEdit} onClose={()=>setShowEdit(false)} title="تعديل تفصيلي للبند" width={640} C={C}>
        {editItem&&(
          <div style={{display:"flex",flexDirection:"column",gap:14}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
              <Select label="الفئة" value={editItem.category} onChange={v=>setEditItem(e=>({...e,category:v}))} options={CATEGORIES} C={C}/>
              <Input label="اسم البند *" value={editItem.name} onChange={v=>setEditItem(e=>({...e,name:v}))} C={C}/>
            </div>
            <Textarea label="المواصفة الفنية" value={editItem.spec||""} onChange={v=>setEditItem(e=>({...e,spec:v}))} placeholder="اكتب المواصفة..." C={C}/>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:12}}>
              <Input label="الكمية" value={editItem.qty||""} onChange={v=>setEditItem(e=>({...e,qty:v}))} type="number" C={C}/>
              <Select label="الوحدة" value={editItem.unit} onChange={v=>setEditItem(e=>({...e,unit:v}))} options={UNITS} C={C}/>
              <Input label="سعر الوحدة" value={editItem.unitPrice||""} onChange={v=>setEditItem(e=>({...e,unitPrice:v}))} type="number" C={C}/>
              <div style={{display:"flex",flexDirection:"column",gap:5}}>
                <label style={{fontSize:12,fontWeight:600,color:C.textSub}}>الإجمالي</label>
                <div style={{padding:"9px 12px",borderRadius:8,background:C.greenLight,color:C.green,fontWeight:700,fontSize:14}}>{((parseFloat(editItem.qty)||0)*(parseFloat(editItem.unitPrice||0)||0)).toLocaleString("ar-SA")||"—"}</div>
              </div>
            </div>
            <Input label="المورد" value={editItem.supplier||""} onChange={v=>setEditItem(e=>({...e,supplier:v}))} C={C}/>
            <Input label="ملاحظات" value={editItem.notes||""} onChange={v=>setEditItem(e=>({...e,notes:v}))} C={C}/>
            <div>
              <label style={{fontSize:12,fontWeight:600,color:C.textSub,display:"block",marginBottom:8}}>صورة البند</label>
              <ImgCell image={editItem.image} onChange={v=>setEditItem(e=>({...e,image:v}))} C={C}/>
            </div>
            <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:8}}>
              <Btn variant="secondary" onClick={()=>setShowEdit(false)} C={C}>إلغاء</Btn>
              <Btn onClick={()=>{push(items.map(it=>it.id===editItem.id?editItem:it));setShowEdit(false);}} C={C}>💾 حفظ</Btn>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
};

// ─── Library ─────────────────────────────────────────────────────────────────
const LibraryPage = ({ C }) => {
  const [tab, setTab] = useState("دهانات");
  const paints=[{company:"جوتن",name:"أبيض كسر",code:"YY0900",type:"داخلي",finish:"مطفي"},{company:"الجزيرة",name:"رمادي فاتح",code:"GT1220",type:"داخلي",finish:"نصف لامع"},{company:"دولوكس",name:"كريمي دافئ",code:"DX3310",type:"داخلي",finish:"مطفي"}];
  const lights=[{type:"سبوت لايت",company:"فيليبس",model:"RS141B",power:"10W",cct:"3000K",ip:"IP20"},{type:"داون لايت",company:"أوسرام",model:"DN150B",power:"15W",cct:"4000K",ip:"IP44"}];
  const cct={"2700K":{color:"#FFD06A",desc:"دافئ جداً — الأماكن الهادئة"},"3000K":{color:"#FFE097",desc:"دافئ — المساكن والفنادق"},"4000K":{color:"#FFFBE0",desc:"محايد — المكاتب"},"5000K":{color:"#FFFFFF",desc:"بارد — الورش"}};
  const tabSt=a=>({padding:"10px 20px",borderRadius:"8px 8px 0 0",border:`1px solid ${a?C.accent:"transparent"}`,borderBottom:a?`2px solid ${C.accent}`:"none",background:a?C.accentLight:"transparent",color:a?C.accentDark:C.textSub,fontFamily:FONT,fontSize:13,fontWeight:600,cursor:"pointer",marginBottom:-1});
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{marginBottom:20}}><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>مكتبة المواد</h2><p style={{fontSize:13,color:C.textSub,marginTop:2}}>مرجع الأصناف والمواصفات الجاهزة</p></div>
      <div style={{display:"flex",gap:8,borderBottom:`1px solid ${C.border}`,marginBottom:24}}>
        {["دهانات","إضاءة","درجات اللون"].map(t=><button key={t} onClick={()=>setTab(t)} style={tabSt(tab===t)}>{t}</button>)}
      </div>
      {tab==="دهانات"&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(260px,1fr))",gap:14}}>{paints.map((p,i)=><Card key={i} C={C}><div style={{fontWeight:700,fontSize:15,marginBottom:6,color:C.text}}>{p.company}</div><div style={{fontSize:14,marginBottom:8,color:C.text}}>{p.name}</div><code style={{fontSize:12,background:C.surfaceAlt,padding:"4px 8px",borderRadius:6,color:C.accent}}>{p.code}</code><div style={{display:"flex",gap:8,marginTop:10}}>{[p.type,p.finish].map(t=><span key={t} style={{fontSize:11,padding:"3px 8px",borderRadius:12,background:C.surfaceAlt,color:C.textSub}}>{t}</span>)}</div></Card>)}</div>}
      {tab==="إضاءة"&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:14}}>{lights.map((l,i)=><Card key={i} C={C}><div style={{display:"flex",justifyContent:"space-between",marginBottom:12}}><div><div style={{fontWeight:700,fontSize:15,color:C.text}}>{l.company} — {l.model}</div><div style={{fontSize:12,color:C.textSub,marginTop:3}}>{l.type}</div></div><span style={{fontSize:24}}>💡</span></div><div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>{[["القدرة",l.power],["درجة اللون",l.cct],["الحماية",l.ip]].map(([k,v])=><div key={k} style={{background:C.surfaceAlt,padding:"8px 10px",borderRadius:8}}><div style={{fontSize:10,color:C.textMuted}}>{k}</div><div style={{fontSize:13,fontWeight:600,marginTop:2,color:C.text}}>{v}</div></div>)}</div></Card>)}</div>}
      {tab==="درجات اللون"&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(220px,1fr))",gap:14}}>{Object.entries(cct).map(([k,d])=><Card key={k} C={C}><div style={{height:60,borderRadius:8,marginBottom:12,background:`radial-gradient(circle at 30% 50%,${d.color},#f0e0b0)`,border:`1px solid ${C.border}`}}/><div style={{fontWeight:700,fontSize:16,marginBottom:6,color:C.text}}>{k}</div><div style={{fontSize:12,color:C.textSub,lineHeight:1.6}}>{d.desc}</div></Card>)}</div>}
    </div>
  );
};

// ─── Suppliers ────────────────────────────────────────────────────────────────
const SuppliersPage = ({ C }) => {
  const [suppliers, setSuppliers] = useState(mockSuppliers);
  const [showModal, setShowModal] = useState(false);
  const [form, setForm] = useState({name:"",phone:"",email:"",category:"دهانات"});
  const add=()=>{if(!form.name)return;setSuppliers([...suppliers,{...form,id:Date.now()}]);setForm({name:"",phone:"",email:"",category:"دهانات"});setShowModal(false);};
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
        <div><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>الموردون</h2><p style={{fontSize:13,color:C.textSub,marginTop:2}}>{suppliers.length} مورد مسجل</p></div>
        <Btn onClick={()=>setShowModal(true)} C={C}>+ إضافة مورد</Btn>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:14}}>
        {suppliers.map(s=>(
          <Card key={s.id} C={C}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:14}}>
              <div><div style={{fontWeight:700,fontSize:15,marginBottom:4,color:C.text}}>{s.name}</div><span style={{background:C.accentLight,color:C.accentDark,padding:"3px 10px",borderRadius:12,fontSize:11,fontWeight:600}}>{s.category}</span></div>
              <div style={{width:40,height:40,borderRadius:"50%",background:C.accentLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18}}>🏢</div>
            </div>
            <div style={{fontSize:13,color:C.textSub}}>📞 {s.phone}</div>
            <div style={{fontSize:13,color:C.textSub,marginTop:4}}>✉️ {s.email}</div>
          </Card>
        ))}
      </div>
      <Modal open={showModal} onClose={()=>setShowModal(false)} title="إضافة مورد جديد" C={C}>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <Input label="اسم المورد *" value={form.name} onChange={v=>setForm({...form,name:v})} placeholder="اسم الشركة" C={C}/>
          <Input label="رقم الجوال" value={form.phone} onChange={v=>setForm({...form,phone:v})} placeholder="05XXXXXXXX" C={C}/>
          <Input label="البريد الإلكتروني" value={form.email} onChange={v=>setForm({...form,email:v})} C={C}/>
          <Select label="التخصص" value={form.category} onChange={v=>setForm({...form,category:v})} options={CATEGORIES} C={C}/>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:8}}>
            <Btn variant="secondary" onClick={()=>setShowModal(false)} C={C}>إلغاء</Btn>
            <Btn onClick={add} C={C}>إضافة</Btn>
          </div>
        </div>
      </Modal>
    </div>
  );
};

// ─── AI ───────────────────────────────────────────────────────────────────────
const AIPage = ({ C }) => {
  const [input,setInput]=useState("");const[result,setResult]=useState("");const[loading,setLoading]=useState(false);
  const gen=async()=>{
    if(!input.trim())return;setLoading(true);setResult("");
    try{const res=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,system:"أنت مواصفاتي داخلي محترف. اكتب مواصفة فنية احترافية للبند المطلوب بالعربية. اشمل: الوصف العام، المواد، المقاسات، التشطيب، التركيب.",messages:[{role:"user",content:`اكتب مواصفة فنية لـ: ${input}`}]})});const d=await res.json();setResult(d.content?.[0]?.text||"");}
    catch{setResult("⚠️ حدث خطأ في الاتصال.");}setLoading(false);
  };
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{marginBottom:20}}><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>مولّد المواصفات ✨</h2><p style={{fontSize:13,color:C.textSub,marginTop:2}}>اكتب اسم البند وسيتولى الذكاء الاصطناعي كتابة المواصفة الكاملة</p></div>
      <Card C={C} style={{marginBottom:20}}>
        <div style={{display:"flex",gap:10,marginBottom:14}}>
          <input value={input} onChange={e=>setInput(e.target.value)} onKeyDown={e=>e.key==="Enter"&&gen()}
            placeholder="مثال: سبوت لايت LED — أو — رخام كرارا للأرضيات..."
            style={{flex:1,padding:"12px 16px",borderRadius:10,border:`1px solid ${C.border}`,fontSize:14,fontFamily:FONT,background:C.surface,color:C.text,outline:"none"}}/>
          <Btn onClick={gen} disabled={loading} style={{padding:"12px 24px"}} C={C}>{loading?"⏳ جاري...":"✨ توليد"}</Btn>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {["سبوت لايت 10W 3000K","دهان مطفي أبيض كسر","رخام كرارا 60×60","باب MDF قشرة جوز"].map(s=>(
            <button key={s} onClick={()=>setInput(s)} style={{padding:"5px 12px",borderRadius:20,border:`1px solid ${C.border}`,background:C.surfaceAlt,fontFamily:FONT,fontSize:12,cursor:"pointer",color:C.textSub}}>{s}</button>
          ))}
        </div>
      </Card>
      {result&&<Card C={C} style={{borderRight:`4px solid ${C.accent}`}}><div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}><div style={{fontWeight:700,color:C.accent}}>✅ المواصفة الفنية المولّدة</div><Btn variant="secondary" sm onClick={()=>navigator.clipboard.writeText(result)} C={C}>📋 نسخ</Btn></div><pre style={{whiteSpace:"pre-wrap",fontFamily:FONT,fontSize:14,lineHeight:1.8,color:C.text}}>{result}</pre></Card>}
    </div>
  );
};

// ─── Reports ──────────────────────────────────────────────────────────────────
const ReportsPage = ({ project, companyName, C }) => {
  const cats = CATEGORIES.slice(0,6).map((c,i)=>({name:c,count:[5,12,8,4,7,3][i]}));
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
        <div><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>التقارير</h2><p style={{fontSize:13,color:C.textSub,marginTop:2}}>{project?.name}</p></div>
        <div style={{display:"flex",gap:8}}>
          <Btn variant="success" sm onClick={()=>dlExcel(project?.name,mockItems,companyName)} C={C}>⬇ Excel</Btn>
          <Btn variant="secondary" sm onClick={()=>dlPDF(project?.name,mockItems,companyName)} C={C}>🖨 PDF</Btn>
        </div>
      </div>
      <Card C={C}>
        <div style={{fontWeight:700,fontSize:15,marginBottom:16,color:C.text}}>توزيع البنود حسب الفئة</div>
        {cats.map(c=>(
          <div key={c.name} style={{display:"flex",alignItems:"center",gap:12,marginBottom:14}}>
            <div style={{width:90,textAlign:"right",fontSize:13,color:C.textSub}}>{c.name}</div>
            <div style={{flex:1,height:10,background:C.surfaceAlt,borderRadius:10,overflow:"hidden"}}><div style={{width:`${(c.count/12)*100}%`,height:"100%",background:C.accent,borderRadius:10}}/></div>
            <div style={{width:28,fontWeight:700,fontSize:13,color:C.accent}}>{c.count}</div>
          </div>
        ))}
      </Card>
    </div>
  );
};

// ─── Settings ─────────────────────────────────────────────────────────────────
const SettingsPage = ({ settings, setSettings, C }) => {
  const [tab, setTab] = useState("عام");
  const logoRef = useRef();
  const tabs = ["عام","معلومات الشركة","الطباعة","حول البرنامج"];
  const TS = a=>({padding:"10px 14px",border:"none",borderRadius:8,background:a?C.accentLight:"transparent",color:a?C.accentDark:C.textSub,fontFamily:FONT,fontSize:13,fontWeight:a?700:500,cursor:"pointer",textAlign:"right",width:"100%",display:"block",marginBottom:2});
  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{marginBottom:20}}><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>⚙️ الإعدادات</h2></div>
      <div style={{display:"flex",gap:22}}>
        <div style={{width:180,flexShrink:0}}>{tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={TS(tab===t)}>{t}</button>)}</div>
        <div style={{flex:1}}>
          {tab==="عام"&&(
            <Card C={C}>
              <div style={{fontWeight:700,fontSize:15,marginBottom:18,color:C.text}}>الإعدادات العامة</div>
              <div style={{display:"flex",flexDirection:"column",gap:20}}>
                <div>
                  <label style={{fontSize:13,fontWeight:600,color:C.textSub,display:"block",marginBottom:10}}>المظهر</label>
                  <div style={{display:"flex",gap:10}}>
                    {[["light","☀️ فاتح"],["dark","🌙 داكن"]].map(([val,lbl])=>(
                      <button key={val} onClick={()=>setSettings(s=>({...s,theme:val}))} style={{padding:"10px 22px",borderRadius:10,border:`2px solid ${settings.theme===val?C.accent:C.border}`,background:settings.theme===val?C.accentLight:C.surface,color:settings.theme===val?C.accentDark:C.text,fontFamily:FONT,fontSize:14,fontWeight:settings.theme===val?700:400,cursor:"pointer"}}>{lbl}</button>
                    ))}
                  </div>
                </div>
                <div>
                  <label style={{fontSize:13,fontWeight:600,color:C.textSub,display:"block",marginBottom:10}}>اللغة / Language</label>
                  <div style={{display:"flex",gap:10}}>
                    {[["ar","🇸🇦 العربية"],["en","🇬🇧 English"]].map(([val,lbl])=>(
                      <button key={val} onClick={()=>setSettings(s=>({...s,lang:val}))} style={{padding:"10px 22px",borderRadius:10,border:`2px solid ${settings.lang===val?C.accent:C.border}`,background:settings.lang===val?C.accentLight:C.surface,color:settings.lang===val?C.accentDark:C.text,fontFamily:FONT,fontSize:14,fontWeight:settings.lang===val?700:400,cursor:"pointer"}}>{lbl}</button>
                    ))}
                  </div>
                </div>
                <Select label="العملة" value={settings.currency||"SAR"} onChange={v=>setSettings(s=>({...s,currency:v}))} options={["SAR","USD","AED","KWD","QAR"]} C={C}/>
              </div>
            </Card>
          )}
          {tab==="معلومات الشركة"&&(
            <Card C={C}>
              <div style={{fontWeight:700,fontSize:15,marginBottom:18,color:C.text}}>بيانات الشركة</div>
              <div style={{display:"flex",flexDirection:"column",gap:16}}>
                <div>
                  <label style={{fontSize:13,fontWeight:600,color:C.textSub,display:"block",marginBottom:10}}>شعار الشركة</label>
                  <div style={{display:"flex",alignItems:"center",gap:16}}>
                    <div onClick={()=>logoRef.current.click()} style={{width:74,height:74,borderRadius:10,border:`2px dashed ${C.accent}`,background:C.accentLight,display:"flex",alignItems:"center",justifyContent:"center",overflow:"hidden",cursor:"pointer"}}>
                      {settings.logo?<img src={settings.logo} alt="logo" style={{width:"100%",height:"100%",objectFit:"contain"}}/>:<span style={{fontSize:28,color:C.accentDark}}>+</span>}
                    </div>
                    <div>
                      <Btn variant="secondary" sm onClick={()=>logoRef.current.click()} C={C}>📁 رفع شعار</Btn>
                      {settings.logo&&<Btn variant="ghost" sm onClick={()=>setSettings(s=>({...s,logo:null}))} style={{marginRight:8}} C={C}>حذف</Btn>}
                      <div style={{fontSize:12,color:C.textMuted,marginTop:6}}>يظهر في الشريط الجانبي</div>
                    </div>
                    <input ref={logoRef} type="file" accept="image/*" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(!f)return;const rd=new FileReader();rd.onload=ev=>setSettings(s=>({...s,logo:ev.target.result}));rd.readAsDataURL(f);}}/>
                  </div>
                </div>
                <Input label="اسم الشركة" value={settings.companyName||""} onChange={v=>setSettings(s=>({...s,companyName:v}))} placeholder="ركيز ديزاين" C={C}/>
                <Input label="البريد الإلكتروني" value={settings.email||""} onChange={v=>setSettings(s=>({...s,email:v}))} C={C}/>
                <Input label="رقم الجوال" value={settings.phone||""} onChange={v=>setSettings(s=>({...s,phone:v}))} C={C}/>
              </div>
            </Card>
          )}
          {tab==="الطباعة"&&<Card C={C}><div style={{fontWeight:700,fontSize:15,marginBottom:16,color:C.text}}>إعدادات التصدير</div><div style={{display:"flex",flexDirection:"column",gap:14}}><Select label="حجم الورق" value={settings.paperSize||"A3"} onChange={v=>setSettings(s=>({...s,paperSize:v}))} options={["A4","A3","A4 Landscape","A3 Landscape"]} C={C}/><div style={{background:C.greenLight,borderRadius:8,padding:"12px 14px",color:C.green,fontSize:13}}>✅ الإعدادات تُحفظ تلقائياً</div></div></Card>}
          {tab==="حول البرنامج"&&(
            <Card C={C}>
              <div style={{textAlign:"center",padding:"20px 0"}}>
                <div style={{width:68,height:68,borderRadius:18,background:C.accent,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:800,fontSize:28,margin:"0 auto 16px"}}>R</div>
                <h2 style={{fontSize:20,fontWeight:800,color:C.text}}>ركيز ديزاين</h2>
                <p style={{color:C.textSub,fontSize:14,margin:"5px 0 3px"}}>نظام إدارة جداول الكميات والمواصفات</p>
                <p style={{color:C.textMuted,fontSize:12}}>Rakeez Design — BOQ Manager</p>
                <div style={{display:"inline-flex",background:C.accentLight,borderRadius:20,padding:"4px 16px",margin:"14px 0",fontSize:13,color:C.accentDark,fontWeight:700}}>الإصدار 2.1</div>
              </div>
            </Card>
          )}
        </div>
      </div>
    </div>
  );
};

// ─── Templates ────────────────────────────────────────────────────────────────
const BUILTIN_TEMPLATES = [
  {id:"t1",name:"فيلا سكنية",nameEn:"Residential Villa",icon:"🏡",desc:"قالب شامل للفلل والمنازل",divisions:[
    {name:"الأرضيات",items:[{category:"أرضيات",name:"رخام للصالات",spec:"رخام 80×80 سم مصقول",qty:"",unit:"م²",supplier:"",notes:"",image:null}]},
    {name:"الأسقف",items:[{category:"أسقف",name:"سقف جبس بورد",spec:"جبس 12.5 مم",qty:"",unit:"م²",supplier:"",notes:"",image:null}]},
    {name:"الدهانات",items:[{category:"دهانات",name:"دهان جدران داخلي",spec:"دهان مطفي طبقتين",qty:"",unit:"م²",supplier:"",notes:"",image:null}]},
    {name:"الإضاءة",items:[{category:"إضاءة",name:"سبوت لايت LED",spec:"10W 3000K IP20",qty:"",unit:"وحدة",supplier:"",notes:"",image:null}]}
  ]},
  {id:"t2",name:"مكتب تجاري",nameEn:"Commercial Office",icon:"🏢",desc:"قالب المكاتب والشركات",divisions:[
    {name:"الأرضيات",items:[{category:"أرضيات",name:"سجاد مكتبي",spec:"سجاد كتل 50×50",qty:"",unit:"م²",supplier:"",notes:"",image:null}]},
    {name:"الإضاءة",items:[{category:"إضاءة",name:"بانل LED 60×60",spec:"36W 4000K",qty:"",unit:"وحدة",supplier:"",notes:"",image:null}]}
  ]},
  {id:"t3",name:"مساحة ترفيهية",nameEn:"Entertainment",icon:"🎡",desc:"ملاهي — مستوحى من Malahi Carnival",divisions:[
    {name:"الأرضيات",items:[{category:"أرضيات",name:"فينيل تجاري ملون",spec:"فينيل تجاري",qty:"",unit:"م²",supplier:"gerflor.com",notes:"",image:null}]},
    {name:"اللافتات",items:[{category:"لافتات",name:"بوابة مدخل مضيئة",spec:"أكريليك مضيء",qty:"",unit:"قطعة",supplier:"",notes:"",image:null}]}
  ]},
];

const TemplatesPage = ({ onApply, C }) => {
  const [prev, setPrev] = useState(null);
  const [imp, setImp] = useState([]);
  const [importing, setImporting] = useState(false);
  const fr = useRef();
  const all = [...BUILTIN_TEMPLATES,...imp];
  const ti = t=>t.divisions.reduce((s,d)=>s+d.items.length,0);

  const handleImp = e=>{
    const file=e.target.files[0]; if(!file) return; setImporting(true);
    const rd=new FileReader();
    rd.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"binary"}); const divs=[];
        wb.SheetNames.forEach(sn=>{
          if(/(COVER|SUMMARY|الغلاف|الملخص|غلاف|Cover)/i.test(sn)) return;
          const ws=wb.Sheets[sn]; const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
          if(rows.length<2) return;
          let hr=0; for(let i=0;i<Math.min(rows.length,5);i++){if(rows[i].some(c=>String(c).trim()==="#"||String(c).trim()==="NO")){hr=i;break;}}
          const hdrs=rows[hr].map(c=>String(c).toLowerCase());
          const ix=(...k)=>{for(const kk of k){const i=hdrs.findIndex(h=>h.includes(kk));if(i>=0)return i;}return -1;};
          const iAr=ix("عربي","ar","الوصف"),iEn=ix("english","en","desc"),iQty=ix("qty","كمية"),iUnit=ix("unit","وحدة"),iSup=ix("supplier","مورد"),iPrice=ix("price","سعر","unit price");
          const items=[];
          for(let r=hr+1;r<rows.length;r++){
            const row=rows[r]; const ar=iAr>=0?String(row[iAr]||"").trim():""; const en=iEn>=0?String(row[iEn]||"").trim():"";
            if(!ar&&!en) continue; if(/إجمالي|مجموع|TOTAL/i.test(ar+en)) continue;
            items.push({category:sn,name:ar||en,spec:en&&ar?en:"",qty:iQty>=0?String(row[iQty]||""):"",unit:iUnit>=0?String(row[iUnit]||"وحدة"):"وحدة",unitPrice:iPrice>=0?String(row[iPrice]||""):"",supplier:iSup>=0?String(row[iSup]||""):"",notes:"",image:null});
          }
          if(items.length>0) divs.push({name:sn,items});
        });
        if(divs.length>0) setImp(p=>[...p,{id:"imp-"+Date.now(),name:file.name.replace(/\.[^.]+$/,""),nameEn:file.name.replace(/\.[^.]+$/,""),icon:"📥",desc:"مستورد من Excel",divisions:divs,imported:true}]);
        else alert("لم يتم العثور على بنود");
      }catch(err){alert("خطأ: "+err.message);}
      setImporting(false); e.target.value="";
    };
    rd.readAsBinaryString(file);
  };

  return (
    <div style={{animation:"fadeIn .25s ease"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
        <div><h2 style={{fontSize:18,fontWeight:700,color:C.text}}>مكتبة القوالب 🗂</h2><p style={{fontSize:13,color:C.textSub,marginTop:2}}>قوالب جاهزة لأنواع المشاريع</p></div>
        <div style={{display:"flex",gap:8}}>
          <Btn variant="secondary" sm onClick={()=>fr.current.click()} disabled={importing} C={C}>{importing?"⏳ استيراد...":"📥 استيراد Excel"}</Btn>
          <input ref={fr} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={handleImp}/>
        </div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(280px,1fr))",gap:14}}>
        {all.map(t=>(
          <Card key={t.id} C={C} style={{position:"relative"}}>
            {t.imported&&<span style={{position:"absolute",top:10,left:10,background:C.blueLight,color:C.blue,fontSize:11,fontWeight:700,padding:"2px 8px",borderRadius:10}}>مستورد</span>}
            <div style={{display:"flex",gap:12,marginBottom:12}}>
              <div style={{width:44,height:44,borderRadius:11,background:C.accentLight,display:"flex",alignItems:"center",justifyContent:"center",fontSize:22,flexShrink:0}}>{t.icon}</div>
              <div><div style={{fontWeight:700,fontSize:15,color:C.text}}>{t.name}</div><div style={{fontSize:12,color:C.textSub}}>{t.nameEn}</div></div>
            </div>
            {t.desc&&<div style={{fontSize:13,color:C.textSub,marginBottom:12,lineHeight:1.6}}>{t.desc}</div>}
            <div style={{display:"flex",gap:8,marginBottom:14}}>
              <span style={{background:C.accentLight,color:C.accentDark,padding:"3px 10px",borderRadius:12,fontSize:12,fontWeight:600}}>{t.divisions.length} قسم</span>
              <span style={{background:C.blueLight,color:C.blue,padding:"3px 10px",borderRadius:12,fontSize:12,fontWeight:600}}>{ti(t)} بند</span>
            </div>
            <div style={{display:"flex",gap:8}}>
              <Btn variant="secondary" sm onClick={()=>setPrev(t)} style={{flex:1,justifyContent:"center"}} C={C}>👁 معاينة</Btn>
              <Btn sm onClick={()=>onApply(t)} style={{flex:1,justifyContent:"center"}} C={C}>✅ استخدام</Btn>
            </div>
          </Card>
        ))}
      </div>

      <Modal open={!!prev} onClose={()=>setPrev(null)} title={prev?"معاينة: "+prev.name:""} width={600} C={C}>
        {prev&&<>
          <div style={{display:"flex",gap:10,marginBottom:16}}>
            <div style={{background:C.accentLight,borderRadius:9,padding:"9px 14px"}}><div style={{fontSize:11,color:C.accentDark}}>الأقسام</div><div style={{fontWeight:700,fontSize:18,color:C.text}}>{prev.divisions.length}</div></div>
            <div style={{background:C.blueLight,borderRadius:9,padding:"9px 14px"}}><div style={{fontSize:11,color:C.blue}}>البنود</div><div style={{fontWeight:700,fontSize:18,color:C.text}}>{ti(prev)}</div></div>
          </div>
          <div style={{display:"flex",flexDirection:"column",gap:8,maxHeight:320,overflow:"auto"}}>
            {prev.divisions.map((div,i)=>(
              <div key={i} style={{border:`1px solid ${C.border}`,borderRadius:9,overflow:"hidden"}}>
                <div style={{background:C.surfaceAlt,padding:"9px 13px",display:"flex",justifyContent:"space-between"}}>
                  <span style={{fontWeight:700,fontSize:13,color:C.text}}>{div.name}</span>
                  <span style={{fontSize:12,background:C.accentLight,color:C.accentDark,padding:"2px 9px",borderRadius:10}}>{div.items.length} بند</span>
                </div>
                <div style={{padding:"7px 13px"}}>
                  {div.items.slice(0,3).map((it,j)=><div key={j} style={{fontSize:12,color:C.textSub,padding:"3px 0",borderBottom:j<2?`1px solid ${C.border}`:"none"}}>• {it.name}</div>)}
                  {div.items.length>3&&<div style={{fontSize:11,color:C.textMuted}}>+ {div.items.length-3} بنود...</div>}
                </div>
              </div>
            ))}
          </div>
          <div style={{display:"flex",gap:9,justifyContent:"flex-end",marginTop:16}}>
            <Btn variant="secondary" onClick={()=>setPrev(null)} C={C}>إغلاق</Btn>
            <Btn onClick={()=>{onApply(prev);setPrev(null);}} C={C}>✅ تطبيق على مشروع</Btn>
          </div>
        </>}
      </Modal>
    </div>
  );
};

// ─── App ──────────────────────────────────────────────────────────────────────
const NAV = [
  { id:"dashboard", label:"الرئيسية", icon:"⊞" },
  { id:"projects",  label:"المشاريع", icon:"🏗" },
  { id:"boq",       label:"جدول الكميات", icon:"📋" },
  { id:"templates", label:"القوالب", icon:"🗂" },
  { id:"library",   label:"المكتبة", icon:"📚" },
  { id:"suppliers", label:"الموردون", icon:"🏢" },
  { id:"ai",        label:"مولّد المواصفات", icon:"✨" },
  { id:"reports",   label:"التقارير", icon:"📊" },
];

export default function App() {
  const [page, setPage]       = useState("dashboard");
  const [projects, setProjects] = useState(mockProjects);
  const [selProj, setSelProj]  = useState(mockProjects[0]);
  const [sbOpen, setSbOpen]    = useState(true);
  const [settings, setSettings] = useState({ theme:"light", lang:"ar", currency:"SAR", companyName:"ركيز ديزاين", logo:null });
  const [tpl, setTpl]          = useState(null);
  const [showTpl, setShowTpl]  = useState(false);
  const [tplForm, setTplForm]  = useState({ name:"", client:"", location:"", status:"جاري" });

  const C = THEMES[settings.theme] || THEMES.light;

  const css = `
    @import url('https://fonts.googleapis.com/css2?family=Noto+Kufi+Arabic:wght@300;400;500;600;700&display=swap');
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    html { direction: ${settings.lang==="ar"?"rtl":"ltr"}; }
    body { font-family: ${FONT}; background: ${C.bg}; color: ${C.text}; }
    ::-webkit-scrollbar { width: 5px; height: 5px; }
    ::-webkit-scrollbar-track { background: ${C.surfaceAlt}; }
    ::-webkit-scrollbar-thumb { background: ${C.borderStrong}; border-radius: 10px; }
    input, textarea, select, button { font-family: ${FONT}; }
    @keyframes fadeIn { from { opacity:0; transform: translateY(5px); } to { opacity:1; transform: translateY(0); } }
  `;

  const onSelect = p=>{ setSelProj(projects.find(pp=>pp.id===p.id)||p); setPage("boq"); };
  const onApply  = t=>{ setTpl(t); setShowTpl(true); };
  const confirmTpl = ()=>{
    if(!tplForm.name) return;
    setProjects(ps=>[...ps,{...tplForm,id:Date.now(),date:new Date().toISOString().slice(0,10),items:tpl.divisions.reduce((s,d)=>s+d.items.length,0),specs:0,images:0}]);
    setShowTpl(false); setTpl(null); setTplForm({name:"",client:"",location:"",status:"جاري"}); setPage("projects");
  };

  const render = ()=>{
    switch(page){
      case "dashboard":  return <DashboardPage projects={projects} onSelectProject={onSelect} C={C}/>;
      case "projects":   return <ProjectsPage projects={projects} setProjects={setProjects} onSelectProject={onSelect} C={C}/>;
      case "boq":        return <BOQPage project={selProj} companyName={settings.companyName} C={C}/>;
      case "templates":  return <TemplatesPage onApply={onApply} C={C}/>;
      case "library":    return <LibraryPage C={C}/>;
      case "suppliers":  return <SuppliersPage C={C}/>;
      case "ai":         return <AIPage C={C}/>;
      case "reports":    return <ReportsPage project={selProj} companyName={settings.companyName} C={C}/>;
      case "settings":   return <SettingsPage settings={settings} setSettings={setSettings} C={C}/>;
      default: return null;
    }
  };

  return (
    <>
      <style>{css}</style>
      <div style={{display:"flex",minHeight:"100vh",background:C.bg}}>
        {/* Sidebar */}
        <aside style={{width:sbOpen?240:60,minWidth:sbOpen?240:60,background:C.surface,borderLeft:`1px solid ${C.border}`,display:"flex",flexDirection:"column",position:"sticky",top:0,height:"100vh",transition:"all .25s",overflow:"hidden"}}>
          {/* Logo */}
          <div style={{padding:"18px 14px 15px",borderBottom:`1px solid ${C.border}`,display:"flex",alignItems:"center",gap:10}}>
            {settings.logo
              ?<img src={settings.logo} alt="logo" style={{width:36,height:36,borderRadius:9,objectFit:"contain",flexShrink:0,border:`1px solid ${C.border}`}}/>
              :<div style={{width:36,height:36,borderRadius:10,background:C.accent,display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",fontWeight:800,fontSize:16,flexShrink:0}}>R</div>}
            {sbOpen&&<div>
              <div style={{fontWeight:800,fontSize:15,color:C.text}}>{settings.companyName||"ركيز ديزاين"}</div>
              <div style={{fontSize:10,color:C.textMuted}}>نظام BOQ</div>
            </div>}
          </div>
          {/* Nav */}
          <nav style={{flex:1,padding:"12px 8px",overflow:"auto"}}>
            {NAV.map(n=>(
              <button key={n.id} onClick={()=>setPage(n.id)} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:8,border:"none",fontFamily:FONT,fontSize:13,cursor:"pointer",marginBottom:4,background:page===n.id?C.accentLight:"transparent",color:page===n.id?C.accentDark:C.textSub,fontWeight:page===n.id?700:500,justifyContent:sbOpen?"flex-start":"center",transition:"all .15s",textAlign:"right"}}>
                <span style={{fontSize:16,flexShrink:0}}>{n.icon}</span>
                {sbOpen&&<span>{n.label}</span>}
              </button>
            ))}
          </nav>
          {/* Current project + Settings */}
          <div style={{borderTop:`1px solid ${C.border}`}}>
            {sbOpen&&selProj&&<div style={{padding:"10px 14px",background:C.surfaceAlt}}>
              <div style={{fontSize:10,color:C.textMuted,marginBottom:3}}>المشروع الحالي</div>
              <div style={{fontSize:12,fontWeight:600,color:C.text,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{selProj.name}</div>
              <Badge status={selProj.status} C={C}/>
            </div>}
            <button onClick={()=>setPage("settings")} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 12px",border:"none",fontFamily:FONT,fontSize:13,cursor:"pointer",background:page==="settings"?C.accentLight:"transparent",color:page==="settings"?C.accentDark:C.textSub,fontWeight:page==="settings"?700:500,justifyContent:sbOpen?"flex-start":"center",transition:"all .15s"}}>
              <span style={{fontSize:16}}>⚙️</span>
              {sbOpen&&<span>الإعدادات</span>}
              {sbOpen&&<span style={{marginRight:"auto",fontSize:11,opacity:.55}}>{settings.lang==="ar"?"🇸🇦":"🇬🇧"} {settings.theme==="dark"?"🌙":"☀️"}</span>}
            </button>
            <button onClick={()=>setSbOpen(x=>!x)} style={{width:"100%",padding:11,border:"none",borderTop:`1px solid ${C.border}`,background:"none",cursor:"pointer",color:C.textMuted,fontSize:16}}>
              {sbOpen?"→":"←"}
            </button>
          </div>
        </aside>

        {/* Main */}
        <main style={{flex:1,overflow:"auto",padding:"28px 32px",minWidth:0,background:C.bg}}>
          {render()}
        </main>
      </div>

      {/* Template apply modal */}
      <Modal open={showTpl} onClose={()=>setShowTpl(false)} title={"إنشاء مشروع من: "+(tpl?.name||"")} width={490} C={C}>
        <div style={{display:"flex",flexDirection:"column",gap:14}}>
          <div style={{background:C.accentLight,borderRadius:9,padding:"11px 14px",display:"flex",gap:10,alignItems:"center"}}>
            <span style={{fontSize:22}}>{tpl?.icon}</span>
            <div><div style={{fontWeight:700,fontSize:14,color:C.text}}>{tpl?.name}</div><div style={{fontSize:12,color:C.accentDark}}>{tpl?.divisions?.length} قسم</div></div>
          </div>
          <Input label="اسم المشروع *" value={tplForm.name} onChange={v=>setTplForm(f=>({...f,name:v}))} C={C}/>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12}}>
            <Input label="العميل" value={tplForm.client} onChange={v=>setTplForm(f=>({...f,client:v}))} C={C}/>
            <Input label="الموقع" value={tplForm.location} onChange={v=>setTplForm(f=>({...f,location:v}))} C={C}/>
          </div>
          <div style={{display:"flex",gap:10,justifyContent:"flex-end",marginTop:6}}>
            <Btn variant="secondary" onClick={()=>setShowTpl(false)} C={C}>إلغاء</Btn>
            <Btn onClick={confirmTpl} disabled={!tplForm.name} C={C}>🚀 إنشاء</Btn>
          </div>
        </div>
      </Modal>
    </>
  );
}
