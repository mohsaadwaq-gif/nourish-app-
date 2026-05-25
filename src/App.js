import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── Constants ────────────────────────────────────────────────────────────────
const MEAL_TYPES = ["Breakfast", "Lunch", "Dinner", "Snack"];
const mealEmoji = { Breakfast: "🌅", Lunch: "☀️", Dinner: "🌙", Snack: "🍎" };
const pastelBg = { breakfast:"#FFD6A544", lunch:"#CAFFBF44", dinner:"#A0C4FF44", snack:"#FFC6FF44" };
const CAL_GOAL = 2000;

const DEFAULT_PRESETS = [
  { id:"p1",  name:"Black Coffee",       emoji:"☕", protein:0,  carbs:0,  fat:0,  notes:"No milk/sugar" },
  { id:"p2",  name:"Latte (Flat White)", emoji:"🥛", protein:4,  carbs:8,  fat:5,  notes:"Whole milk" },
  { id:"p3",  name:"Protein Shake",      emoji:"🥤", protein:25, carbs:5,  fat:3,  notes:"1 scoop whey + water" },
  { id:"p4",  name:"Protein Bar",        emoji:"🍫", protein:20, carbs:22, fat:8,  notes:"~220 kcal bar" },
  { id:"p5",  name:"Banana",             emoji:"🍌", protein:1,  carbs:27, fat:0,  notes:"Medium banana" },
  { id:"p6",  name:"Greek Yogurt",       emoji:"🥣", protein:17, carbs:6,  fat:0,  notes:"170g non-fat" },
  { id:"p7",  name:"Boiled Eggs (x2)",   emoji:"🥚", protein:13, carbs:1,  fat:10, notes:"Hard boiled" },
  { id:"p8",  name:"Oatmeal",            emoji:"🥣", protein:5,  carbs:27, fat:3,  notes:"40g oats + water" },
  { id:"p9",  name:"Chicken Breast",     emoji:"🍗", protein:31, carbs:0,  fat:3,  notes:"100g grilled" },
  { id:"p10", name:"Brown Rice",         emoji:"🍚", protein:3,  carbs:44, fat:1,  notes:"180g cooked" },
  { id:"p11", name:"Avocado Toast",      emoji:"🥑", protein:5,  carbs:28, fat:12, notes:"1 slice sourdough" },
  { id:"p12", name:"Almonds (30g)",      emoji:"🌰", protein:6,  carbs:5,  fat:15, notes:"Small handful" },
];

const generateId = () => Math.random().toString(36).slice(2,10);
const getTodayKey = () => new Date().toISOString().slice(0,10);

function calcCalories(p,c,f) {
  const pn=Number(p)||0, cn=Number(c)||0, fn=Number(f)||0;
  if (!pn&&!cn&&!fn) return "";
  return String(Math.round(pn*4+cn*4+fn*9));
}
function formatDateLong(d) {
  const [y,m,day] = d.split("-");
  return new Date(y,m-1,day).toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"});
}
function getDayTotals(meals=[]) {
  return meals.reduce((a,m)=>({
    calories:a.calories+(Number(m.calories)||0),
    protein:a.protein+(Number(m.protein)||0),
    carbs:a.carbs+(Number(m.carbs)||0),
    fat:a.fat+(Number(m.fat)||0),
  }),{calories:0,protein:0,carbs:0,fat:0});
}
function buildCalendarGrid(baseDate) {
  const [y,m] = baseDate.split("-").map(Number);
  const first = new Date(y,m-1,1);
  const last  = new Date(y,m,0);
  const cells = [];
  for (let i=0;i<first.getDay();i++) cells.push(null);
  for (let d=1;d<=last.getDate();d++)
    cells.push(`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  while (cells.length%7!==0) cells.push(null);
  return cells;
}

const EMPTY_FORM   = {type:"Breakfast",name:"",calories:"",protein:"",carbs:"",fat:"",notes:""};
const EMPTY_PRESET = {name:"",emoji:"🍽",protein:"",carbs:"",fat:"",notes:""};

const inputStyle = {
  width:"100%",padding:"10px 14px",borderRadius:10,
  border:"1.5px solid #d0c4f0",fontSize:14,color:"#3d1f6b",
  background:"#faf8ff",outline:"none",fontFamily:"Georgia,serif",
  marginBottom:10,boxSizing:"border-box",
};
const smallInput = {...inputStyle,fontSize:13,padding:"8px 10px",marginBottom:0};

function exportPDF(allData, dateRange) {
  const rows = dateRange.flatMap(date => {
    const meals = allData[date] || [];
    if (!meals.length) return [];
    const t = getDayTotals(meals);
    return [
      `<tr style="background:#f0ebff"><td colspan="7" style="padding:6px 8px;font-weight:700;color:#3d1f6b">${formatDateLong(date)}</td></tr>`,
      ...meals.map(m=>`<tr><td style="padding:4px 8px">${m.type}</td><td style="padding:4px 8px;font-weight:600">${m.name}</td><td style="padding:4px 8px;text-align:right">${m.calories||""}</td><td style="padding:4px 8px;text-align:right">${m.protein||""}</td><td style="padding:4px 8px;text-align:right">${m.carbs||""}</td><td style="padding:4px 8px;text-align:right">${m.fat||""}</td><td style="padding:4px 8px;color:#666;font-size:11px">${m.notes||""}</td></tr>`),
      `<tr style="background:#faf8ff;font-style:italic;font-size:11px"><td colspan="2" style="padding:4px 8px;color:#9b87c2">Day total</td><td style="padding:4px 8px;text-align:right;color:#8b5cf6;font-weight:700">${t.calories}</td><td style="padding:4px 8px;text-align:right;color:#FF6B6B">${t.protein}g</td><td style="padding:4px 8px;text-align:right;color:#4ECDC4">${t.carbs}g</td><td style="padding:4px 8px;text-align:right;color:#FFD93D">${t.fat}g</td><td></td></tr>`,
    ];
  });
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Nourish Report</title><style>body{font-family:Georgia,serif;padding:28px;color:#2d1b55;font-size:13px}h1{font-size:22px;color:#3d1f6b}table{width:100%;border-collapse:collapse}th{background:#3d1f6b;color:white;padding:7px 8px;text-align:left;font-size:12px}tr:nth-child(even){background:#faf8ff}@media print{button{display:none}}</style></head><body><h1>🍽 Nourish — Meal Report</h1><p>Generated ${new Date().toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"})}</p><table><thead><tr><th>Meal Type</th><th>Food</th><th>Calories</th><th>Protein</th><th>Carbs</th><th>Fat</th><th>Notes</th></tr></thead><tbody>${rows.join("")}</tbody></table></body></html>`;
  const win=window.open("","_blank");
  win.document.write(html);
  win.document.close();
  setTimeout(()=>win.print(),400);
}

function exportExcel(allData, dateRange) {
  const rows=[["Date","Meal Type","Food Name","Calories (kcal)","Protein (g)","Carbs (g)","Fat (g)","Notes"]];
  dateRange.forEach(date=>{
    const meals=allData[date]||[];
    if(!meals.length) return;
    meals.forEach(m=>rows.push([formatDateLong(date),m.type,m.name,Number(m.calories)||"",Number(m.protein)||"",Number(m.carbs)||"",Number(m.fat)||"",m.notes||""]));
    const t=getDayTotals(meals);
    rows.push([`Day Total (${formatDateLong(date)})`, "","",t.calories,t.protein,t.carbs,t.fat,""]);
    rows.push([]);
  });
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=[18,12,28,16,12,10,8,30].map(w=>({wch:w}));
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,"Meal Log");
  XLSX.writeFile(wb,`nourish-report-${getTodayKey()}.xlsx`);
}

export default function App() {
  const [allData,setAllData]=useState({});
  const [presets,setPresets]=useState(DEFAULT_PRESETS);
  const [selDate,setSelDate]=useState(getTodayKey());
  const [activeTab,setActiveTab]=useState("log");
  const today=getTodayKey();
  const [calMonth,setCalMonth]=useState(today.slice(0,7));
  const [showForm,setShowForm]=useState(false);
  const [form,setForm]=useState(EMPTY_FORM);
  const [editId,setEditId]=useState(null);
  const [presetSearch,setPresetSearch]=useState("");
  const [showLibrary,setShowLibrary]=useState(false);
  const [presetForm,setPresetForm]=useState(EMPTY_PRESET);
  const [editPresetId,setEditPresetId]=useState(null);
  const [libSearch,setLibSearch]=useState("");
  const [showReport,setShowReport]=useState(false);
  const [reportRange,setReportRange]=useState("7");
  const [reportFrom,setReportFrom]=useState("");
  const [reportTo,setReportTo]=useState(today);
  const [cameraStep,setCameraStep]=useState("idle");
  const [capturedImage,setCapturedImage]=useState(null);
  const [analysisError,setAnalysisError]=useState("");
  const videoRef=useRef(null);
  const streamRef=useRef(null);
  const fileInputRef=useRef(null);

  useEffect(()=>{
    const d=localStorage.getItem("nourish_data_v3");
    const p=localStorage.getItem("nourish_presets_v1");
    if(d) try{setAllData(JSON.parse(d));}catch{}
    if(p) try{setPresets(JSON.parse(p));}catch{}
  },[]);
  useEffect(()=>{localStorage.setItem("nourish_data_v3",JSON.stringify(allData));},[allData]);
  useEffect(()=>{localStorage.setItem("nourish_presets_v1",JSON.stringify(presets));},[presets]);

  const stopCamera=useCallback(()=>{streamRef.current?.getTracks().forEach(t=>t.stop());streamRef.current=null;},[]);
  useEffect(()=>{if(!showForm){stopCamera();setCameraStep("idle");setCapturedImage(null);setAnalysisError("");}},[showForm,stopCamera]);

  async function startCamera(){
    setAnalysisError("");setCameraStep("preview");
    try{
      const s=await navigator.mediaDevices.getUserMedia({video:{facingMode:"environment"},audio:false});
      streamRef.current=s;
      if(videoRef.current){videoRef.current.srcObject=s;videoRef.current.play();}
    }catch{setAnalysisError("Camera access denied. Use upload instead.");setCameraStep("idle");}
  }
  function captureFromCamera(){
    const v=videoRef.current;if(!v)return;
    const c=document.createElement("canvas");c.width=v.videoWidth;c.height=v.videoHeight;
    c.getContext("2d").drawImage(v,0,0);
    const url=c.toDataURL("image/jpeg",0.85);
    stopCamera();setCapturedImage(url);analyzeImage(url);
  }
  function handleFileUpload(e){
    const file=e.target.files[0];if(!file)return;
    const r=new FileReader();
    r.onload=ev=>{setCapturedImage(ev.target.result);analyzeImage(ev.target.result);};
    r.readAsDataURL(file);e.target.value="";
  }
  async function analyzeImage(dataUrl){
    setCameraStep("analyzing");setAnalysisError("");
    const base64=dataUrl.split(",")[1];
    const mediaType=dataUrl.startsWith("data:image/png")?"image/png":"image/jpeg";
    try{
      const res=await fetch("https://api.anthropic.com/v1/messages",{
        method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:1000,
          messages:[{role:"user",content:[
            {type:"image",source:{type:"base64",media_type:mediaType,data:base64}},
            {type:"text",text:`Analyze this food photo. Return ONLY valid JSON, no markdown:\n{"name":"short meal name","protein":number,"carbs":number,"fat":number,"notes":"brief note"}`}
          ]}]})
      });
      const data=await res.json();
      const text=data.content?.map(b=>b.text||"").join("")||"";
      const parsed=JSON.parse(text.replace(/```json|```/g,"").trim());
      const protein=String(parsed.protein||""),carbs=String(parsed.carbs||""),fat=String(parsed.fat||"");
      setForm(f=>({...f,name:parsed.name||f.name,protein,carbs,fat,calories:calcCalories(protein,carbs,fat),notes:parsed.notes||f.notes}));
      setCameraStep("done");
    }catch{setAnalysisError("Could not analyse. Fill macros manually.");setCameraStep("done");}
  }

  const todayMeals=allData[selDate]||[];
  const totals=getDayTotals(todayMeals);
  const calPct=Math.min((totals.calories/CAL_GOAL)*100,100);

  function updateMacro(key,val){setForm(f=>{const n={...f,[key]:val};n.calories=calcCalories(n.protein,n.carbs,n.fat);return n;});}
  function openAdd(){setForm(EMPTY_FORM);setEditId(null);setPresetSearch("");setCameraStep("idle");setCapturedImage(null);setAnalysisError("");setShowForm(true);}
  function openEdit(meal){setForm({...meal});setEditId(meal.id);setPresetSearch("");setCameraStep("idle");setCapturedImage(null);setAnalysisError("");setShowForm(true);}
  function applyPreset(p){
    const pr=String(p.protein||""),cr=String(p.carbs||""),ft=String(p.fat||"");
    setForm(f=>({...f,name:p.name,protein:pr,carbs:cr,fat:ft,calories:calcCalories(pr,cr,ft),notes:p.notes||""}));
    setPresetSearch("");
  }
  function saveMeal(){
    if(!form.name.trim())return;
    const u={...allData};const meals=[...(u[selDate]||[])];
    if(editId){const i=meals.findIndex(m=>m.id===editId);if(i!==-1)meals[i]={...form,id:editId};}
    else meals.push({...form,id:generateId()});
    u[selDate]=meals;setAllData(u);setShowForm(false);setEditId(null);
  }
  function deleteMeal(id){const u={...allData};u[selDate]=(u[selDate]||[]).filter(m=>m.id!==id);setAllData(u);}

  function savePreset(){
    if(!presetForm.name.trim())return;
    if(editPresetId)setPresets(ps=>ps.map(p=>p.id===editPresetId?{...presetForm,id:editPresetId}:p));
    else setPresets(ps=>[...ps,{...presetForm,id:"u"+generateId()}]);
    setPresetForm(EMPTY_PRESET);setEditPresetId(null);
  }
  function editPreset(p){setPresetForm({...p});setEditPresetId(p.id);}
  function deletePreset(id){setPresets(ps=>ps.filter(p=>p.id!==id));if(editPresetId===id){setPresetForm(EMPTY_PRESET);setEditPresetId(null);}}
  function cancelPresetEdit(){setPresetForm(EMPTY_PRESET);setEditPresetId(null);}

  const filteredPresets=presets.filter(p=>!presetSearch||p.name.toLowerCase().includes(presetSearch.toLowerCase()));
  const libFiltered=presets.filter(p=>!libSearch||p.name.toLowerCase().includes(libSearch.toLowerCase()));

  const calGrid=buildCalendarGrid(calMonth+"-01");
  function prevMonth(){const [y,m]=calMonth.split("-").map(Number);const d=new Date(y,m-2,1);setCalMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}
  function nextMonth(){const [y,m]=calMonth.split("-").map(Number);const d=new Date(y,m,1);setCalMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}
  const monthLabel=new Date(calMonth+"-01").toLocaleDateString("en-US",{month:"long",year:"numeric"});

  function getReportDates(){
    if(reportRange==="custom"){
      if(!reportFrom||!reportTo)return[];
      const days=[];let cur=new Date(reportFrom);const end=new Date(reportTo);
      while(cur<=end){days.push(cur.toISOString().slice(0,10));cur.setDate(cur.getDate()+1);}
      return days;
    }
    const n=Number(reportRange);const days=[];
    for(let i=n-1;i>=0;i--){const d=new Date();d.setDate(d.getDate()-i);days.push(d.toISOString().slice(0,10));}
    return days;
  }
  const reportDates=getReportDates();
  const reportDaysWithData=reportDates.filter(d=>(allData[d]||[]).length>0);
  const reportTotals=reportDates.reduce((a,d)=>{const t=getDayTotals(allData[d]||[]);return{calories:a.calories+t.calories,protein:a.protein+t.protein,carbs:a.carbs+t.carbs,fat:a.fat+t.fat};},{calories:0,protein:0,carbs:0,fat:0});
  const avgCalories=reportDaysWithData.length>0?Math.round(reportTotals.calories/reportDaysWithData.length):0;

  return (
    <div style={{minHeight:"100vh",background:"linear-gradient(135deg,#f8f0ff 0%,#e8f4fd 50%,#f0fff4 100%)",fontFamily:"Georgia,serif"}}>

      {/* Header */}
      <div style={{background:"rgba(255,255,255,0.9)",backdropFilter:"blur(12px)",borderBottom:"1.5px solid #e2d9f3",padding:"12px 16px 10px",position:"sticky",top:0,zIndex:50,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:19,fontWeight:700,color:"#3d1f6b"}}>🍽 Nourish</div>
          <div style={{fontSize:9,color:"#9b87c2",letterSpacing:"0.08em",textTransform:"uppercase"}}>Daily Meals Tracker</div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:7}}>
          <button onClick={()=>setShowLibrary(true)} style={{padding:"5px 10px",borderRadius:9,border:"1.5px solid #d0c4f0",background:"white",color:"#6b5b9e",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>📚 Foods</button>
          <button onClick={()=>setShowReport(true)} style={{padding:"5px 10px",borderRadius:9,border:"1.5px solid #d0c4f0",background:"white",color:"#6b5b9e",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>📊 Report</button>
          <input type="date" value={selDate} onChange={e=>setSelDate(e.target.value)} style={{border:"1.5px solid #d0c4f0",borderRadius:9,padding:"4px 8px",fontSize:11,color:"#3d1f6b",background:"white",outline:"none",fontFamily:"inherit"}}/>
        </div>
      </div>

      {/* Tabs */}
      <div style={{display:"flex",padding:"11px 16px 0",gap:5}}>
        {[["log","📋 Daily Log"],["calendar","📅 Calendar"]].map(([tab,label])=>(
          <button key={tab} onClick={()=>setActiveTab(tab)} style={{padding:"6px 14px",borderRadius:20,border:"1.5px solid",borderColor:activeTab===tab?"#8b5cf6":"#d0c4f0",background:activeTab===tab?"#8b5cf6":"white",color:activeTab===tab?"white":"#6b5b9e",fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>{label}</button>
        ))}
      </div>

      {/* Daily Log */}
      {activeTab==="log"&&(
        <div style={{padding:"13px 16px 80px"}}>
          <div style={{fontSize:13,color:"#6b5b9e",fontStyle:"italic",marginBottom:11}}>{selDate===today?"Today — ":""}{formatDateLong(selDate)}</div>
          <div style={{background:"white",borderRadius:16,padding:"14px 16px",boxShadow:"0 4px 20px rgba(139,92,246,0.08)",marginBottom:14,display:"flex",gap:16,alignItems:"center"}}>
            <div style={{position:"relative",width:66,height:66,flexShrink:0}}>
              <svg width="66" height="66" viewBox="0 0 66 66">
                <circle cx="33" cy="33" r="25" fill="none" stroke="#f0e8ff" strokeWidth="7"/>
                <circle cx="33" cy="33" r="25" fill="none" stroke="#8b5cf6" strokeWidth="7" strokeDasharray={`${calPct*1.571} 157.1`} strokeLinecap="round" transform="rotate(-90 33 33)" style={{transition:"stroke-dasharray 0.6s ease"}}/>
              </svg>
              <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                <span style={{fontSize:13,fontWeight:700,color:"#3d1f6b",lineHeight:1}}>{totals.calories}</span>
                <span style={{fontSize:8,color:"#9b87c2"}}>kcal</span>
              </div>
            </div>
            <div style={{flex:1,display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:"6px 8px"}}>
              {[{l:"Protein",v:totals.protein,c:"#FF6B6B"},{l:"Carbs",v:totals.carbs,c:"#4ECDC4"},{l:"Fat",v:totals.fat,c:"#FFD93D"}].map(m=>(
                <div key={m.l} style={{textAlign:"center"}}>
                  <div style={{fontSize:14,fontWeight:700,color:m.c}}>{m.v}<span style={{fontSize:9}}>g</span></div>
                  <div style={{fontSize:9,color:"#9b87c2",textTransform:"uppercase",letterSpacing:"0.04em"}}>{m.l}</div>
                </div>
              ))}
              <div style={{gridColumn:"1/-1",fontSize:10,color:"#b8a9d9",textAlign:"center"}}>Goal {CAL_GOAL} kcal · {Math.round(calPct)}% reached</div>
            </div>
          </div>
          {MEAL_TYPES.map(type=>{
            const meals=todayMeals.filter(m=>m.type===type);
            return(
              <div key={type} style={{marginBottom:11}}>
                <div style={{display:"flex",alignItems:"center",gap:6,marginBottom:4,color:"#3d1f6b",fontWeight:700,fontSize:12}}>{mealEmoji[type]} {type} <span style={{fontSize:10,color:"#9b87c2",fontWeight:400,fontStyle:"italic"}}>{meals.length>0?`${meals.reduce((a,m)=>a+(Number(m.calories)||0),0)} kcal`:"nothing yet"}</span></div>
                {meals.map(meal=>(
                  <div key={meal.id} style={{background:pastelBg[type.toLowerCase()]||"#f9f6ff",border:"1.5px solid #e2d9f3",borderRadius:11,padding:"8px 11px",marginBottom:4,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <div>
                      <div style={{fontWeight:600,color:"#2d1b55",fontSize:12}}>{meal.name}</div>
                      <div style={{fontSize:10,color:"#7b6aab",marginTop:2}}>{meal.calories?`${meal.calories} kcal`:""}{meal.protein?` · ${meal.protein}g P`:""}{meal.carbs?` · ${meal.carbs}g C`:""}{meal.fat?` · ${meal.fat}g F`:""}{meal.notes?` · ${meal.notes}`:""}</div>
                    </div>
                    <div style={{display:"flex",gap:3}}>
                      <button onClick={()=>openEdit(meal)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,padding:3}}>✏️</button>
                      <button onClick={()=>deleteMeal(meal.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:12,padding:3}}>🗑</button>
                    </div>
                  </div>
                ))}
              </div>
            );
          })}
          {todayMeals.length===0&&<div style={{textAlign:"center",color:"#b8a9d9",fontSize:13,fontStyle:"italic",marginTop:28,padding:20}}>No meals logged yet.<br/>Tap + to add your first meal!</div>}
        </div>
      )}

      {/* Calendar */}
      {activeTab==="calendar"&&(
        <div style={{padding:"13px 16px 80px"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
            <button onClick={prevMonth} style={{background:"none",border:"1.5px solid #d0c4f0",borderRadius:9,padding:"5px 12px",cursor:"pointer",color:"#6b5b9e",fontSize:14,fontFamily:"inherit"}}>‹</button>
            <div style={{fontWeight:700,color:"#3d1f6b",fontSize:15}}>{monthLabel}</div>
            <button onClick={nextMonth} style={{background:"none",border:"1.5px solid #d0c4f0",borderRadius:9,padding:"5px 12px",cursor:"pointer",color:"#6b5b9e",fontSize:14,fontFamily:"inherit"}}>›</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:3}}>
            {["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].map(d=><div key={d} style={{textAlign:"center",fontSize:10,color:"#9b87c2",fontWeight:600,padding:"3px 0"}}>{d}</div>)}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
            {calGrid.map((dateKey,i)=>{
              if(!dateKey)return<div key={i}/>;
              const meals=allData[dateKey]||[];
              const t=getDayTotals(meals);
              const isToday=dateKey===today,isSelected=dateKey===selDate,hasMeals=meals.length>0;
              const pct=Math.min(t.calories/CAL_GOAL,1);
              const ringColor=pct>1?"#FF6B6B":pct>0.6?"#4ECDC4":"#8b5cf6";
              return(
                <div key={dateKey} onClick={()=>{setSelDate(dateKey);setActiveTab("log");}} style={{background:isSelected?"#8b5cf6":isToday?"#f0ebff":"white",border:`1.5px solid ${isSelected?"#8b5cf6":isToday?"#c4a8f0":"#e2d9f3"}`,borderRadius:10,padding:"6px 4px 5px",cursor:"pointer",minHeight:56,display:"flex",flexDirection:"column",alignItems:"center",gap:2,transition:"all 0.15s"}}>
                  <div style={{fontSize:12,fontWeight:isToday||isSelected?700:500,color:isSelected?"white":isToday?"#8b5cf6":"#3d1f6b"}}>{Number(dateKey.slice(8))}</div>
                  {hasMeals&&(<>
                    <svg width="28" height="28" viewBox="0 0 28 28">
                      <circle cx="14" cy="14" r="10" fill="none" stroke={isSelected?"rgba(255,255,255,0.3)":"#f0e8ff"} strokeWidth="3.5"/>
                      <circle cx="14" cy="14" r="10" fill="none" stroke={isSelected?"white":ringColor} strokeWidth="3.5" strokeDasharray={`${pct*62.8} 62.8`} strokeLinecap="round" transform="rotate(-90 14 14)"/>
                    </svg>
                    <div style={{fontSize:9,color:isSelected?"rgba(255,255,255,0.85)":"#8b5cf6",fontWeight:600,lineHeight:1}}>{t.calories}</div>
                    <div style={{fontSize:8,color:isSelected?"rgba(255,255,255,0.6)":"#b8a9d9",lineHeight:1}}>kcal</div>
                  </>)}
                  {!hasMeals&&<div style={{width:6,height:6,borderRadius:"50%",background:isSelected?"rgba(255,255,255,0.4)":"#e8e0f8",marginTop:2}}/>}
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",gap:12,marginTop:12,justifyContent:"center",flexWrap:"wrap"}}>
            {[["#8b5cf6","On track"],["#4ECDC4","60–100%"],["#FF6B6B","Over goal"],["#e8e0f8","No entries"]].map(([c,l])=>(
              <div key={l} style={{display:"flex",alignItems:"center",gap:4}}><div style={{width:10,height:10,borderRadius:"50%",background:c}}/><span style={{fontSize:10,color:"#9b87c2"}}>{l}</span></div>
            ))}
          </div>
          {selDate&&(
            <div style={{marginTop:16,background:"white",borderRadius:14,padding:"13px 14px",boxShadow:"0 2px 12px rgba(139,92,246,0.08)"}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
                <div style={{fontWeight:700,color:"#3d1f6b",fontSize:13}}>{selDate===today?"Today — ":""}{formatDateLong(selDate)}</div>
                <button onClick={()=>setActiveTab("log")} style={{background:"#8b5cf6",border:"none",borderRadius:8,padding:"4px 10px",color:"white",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>View →</button>
              </div>
              {(allData[selDate]||[]).length===0?<div style={{fontSize:12,color:"#b8a9d9",fontStyle:"italic"}}>No meals logged.</div>:(allData[selDate]||[]).map(m=>(
                <div key={m.id} style={{display:"flex",justifyContent:"space-between",borderBottom:"1px solid #f0ebff",padding:"4px 0",fontSize:12}}>
                  <span style={{color:"#3d1f6b"}}>{mealEmoji[m.type]} {m.name}</span>
                  <span style={{color:"#8b5cf6",fontWeight:600}}>{m.calories||"—"} kcal</span>
                </div>
              ))}
            </div>
          )}
        </div>
      )}

      {/* FAB */}
      <button onClick={openAdd} style={{position:"fixed",bottom:24,right:20,width:52,height:52,borderRadius:"50%",background:"linear-gradient(135deg,#8b5cf6,#6366f1)",border:"none",color:"white",fontSize:24,cursor:"pointer",boxShadow:"0 6px 24px rgba(139,92,246,0.4)",display:"flex",alignItems:"center",justifyContent:"center",zIndex:100}}>+</button>
      <input ref={fileInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleFileUpload}/>

      {/* Add/Edit Meal Modal */}
      {showForm&&(
        <div style={{position:"fixed",inset:0,background:"rgba(30,10,60,0.45)",backdropFilter:"blur(4px)",zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget)setShowForm(false);}}>
          <div style={{background:"white",borderRadius:"22px 22px 0 0",padding:"20px 16px 34px",width:"100%",maxWidth:480,boxShadow:"0 -8px 40px rgba(139,92,246,0.18)",maxHeight:"92vh",overflowY:"auto"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
              <div style={{fontSize:15,fontWeight:700,color:"#3d1f6b"}}>{editId?"Edit Meal":"Add Meal"}</div>
              <button onClick={()=>setShowForm(false)} style={{background:"none",border:"none",fontSize:17,cursor:"pointer",color:"#9b87c2"}}>✕</button>
            </div>
            {!editId&&(
              <div style={{marginBottom:12}}>
                <div style={{fontSize:11,color:"#9b87c2",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:6}}>⚡ Quick Pick</div>
                <input placeholder="Search foods…" value={presetSearch} onChange={e=>setPresetSearch(e.target.value)} style={{...inputStyle,marginBottom:6,fontSize:12,padding:"7px 10px"}}/>
                <div style={{display:"flex",gap:5,flexWrap:"wrap",maxHeight:80,overflowY:"auto"}}>
                  {filteredPresets.map(p=>(
                    <button key={p.id} onClick={()=>applyPreset(p)} style={{padding:"5px 10px",borderRadius:18,border:"1.5px solid #d0c4f0",background:"#faf8ff",color:"#3d1f6b",fontSize:11,fontWeight:500,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",gap:3,whiteSpace:"nowrap"}}>
                      {p.emoji} {p.name} <span style={{color:"#9b87c2",fontSize:9}}>{calcCalories(p.protein,p.carbs,p.fat)||"0"}kcal</span>
                    </button>
                  ))}
                </div>
                <div style={{height:1,background:"#ede8f8",margin:"10px 0"}}/>
              </div>
            )}
            <div style={{display:"flex",gap:5,marginBottom:10,flexWrap:"wrap"}}>
              {MEAL_TYPES.map(t=>(
                <button key={t} onClick={()=>setForm(f=>({...f,type:t}))} style={{padding:"4px 10px",borderRadius:16,border:"1.5px solid",borderColor:form.type===t?"#8b5cf6":"#d0c4f0",background:form.type===t?"#8b5cf6":"white",color:form.type===t?"white":"#6b5b9e",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>{mealEmoji[t]} {t}</button>
              ))}
            </div>
            {!editId&&(
              <div style={{marginBottom:10}}>
                {cameraStep==="idle"&&(
                  <div style={{display:"flex",gap:6}}>
                    <button onClick={startCamera} style={{flex:1,padding:"8px 0",borderRadius:10,border:"1.5px dashed #8b5cf6",background:"#faf8ff",color:"#8b5cf6",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:4}}>📷 Camera</button>
                    <button onClick={()=>fileInputRef.current?.click()} style={{flex:1,padding:"8px 0",borderRadius:10,border:"1.5px dashed #6366f1",background:"#f8f8ff",color:"#6366f1",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",display:"flex",alignItems:"center",justifyContent:"center",gap:4}}>🖼 Upload</button>
                  </div>
                )}
                {cameraStep==="preview"&&(
                  <div style={{borderRadius:12,overflow:"hidden",background:"#000"}}>
                    <video ref={videoRef} autoPlay playsInline muted style={{width:"100%",maxHeight:180,objectFit:"cover",display:"block"}}/>
                    <div style={{display:"flex",gap:6,padding:"8px",background:"rgba(0,0,0,0.6)"}}>
                      <button onClick={()=>{stopCamera();setCameraStep("idle");}} style={{flex:1,padding:"7px",borderRadius:8,border:"none",background:"rgba(255,255,255,0.2)",color:"white",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
                      <button onClick={captureFromCamera} style={{flex:2,padding:"7px",borderRadius:8,border:"none",background:"#8b5cf6",color:"white",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📸 Capture</button>
                    </div>
                  </div>
                )}
                {(cameraStep==="analyzing"||cameraStep==="done")&&capturedImage&&(
                  <div style={{borderRadius:12,overflow:"hidden"}}>
                    <img src={capturedImage} alt="meal" style={{width:"100%",maxHeight:120,objectFit:"cover",display:"block"}}/>
                    {cameraStep==="analyzing"&&<div style={{padding:"8px",background:"rgba(61,31,107,0.7)",display:"flex",alignItems:"center",gap:7}}><span style={{fontSize:16}}>🔍</span><span style={{color:"white",fontSize:11}}>Analysing nutrition…</span></div>}
                    {cameraStep==="done"&&(
                      <div style={{padding:"5px 8px",background:analysisError?"#fff0f0":"#f0fff4",display:"flex",alignItems:"center",gap:5}}>
                        <span style={{fontSize:12}}>{analysisError?"⚠️":"✅"}</span>
                        <span style={{fontSize:11,color:analysisError?"#c0392b":"#27ae60"}}>{analysisError||"Macros pre-filled"}</span>
                        <button onClick={()=>{setCameraStep("idle");setCapturedImage(null);setAnalysisError("");}} style={{marginLeft:"auto",background:"none",border:"none",fontSize:10,color:"#9b87c2",cursor:"pointer"}}>Retake</button>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}
            <input placeholder="Meal name" value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={inputStyle}/>
            <div style={{fontSize:10,color:"#9b87c2",marginBottom:4,fontStyle:"italic"}}>Protein + Carbs + Fat → calories auto-calculated</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:6}}>
              {[{key:"protein",ph:"Protein g",col:"#FF6B6B"},{key:"carbs",ph:"Carbs g",col:"#4ECDC4"},{key:"fat",ph:"Fat g",col:"#FFD93D"}].map(({key,ph,col})=>(
                <input key={key} type="number" min="0" placeholder={ph} value={form[key]} onChange={e=>updateMacro(key,e.target.value)} style={{...smallInput,borderColor:form[key]?col:"#d0c4f0"}}/>
              ))}
            </div>
            <div style={{background:"#faf0ff",border:"1.5px solid #c4a8f0",borderRadius:9,padding:"8px 12px",display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
              <span style={{fontSize:11,color:"#8b5cf6",fontWeight:600}}>🔥 Calories</span>
              <input type="number" min="0" value={form.calories} onChange={e=>setForm(f=>({...f,calories:e.target.value}))} placeholder="auto" style={{border:"none",background:"transparent",fontSize:14,fontWeight:700,color:"#3d1f6b",width:70,textAlign:"right",outline:"none",fontFamily:"inherit"}}/>
              <span style={{fontSize:10,color:"#9b87c2"}}>kcal</span>
            </div>
            <input placeholder="Notes (optional)" value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} style={inputStyle}/>
            <div style={{display:"flex",gap:8}}>
              <button onClick={()=>setShowForm(false)} style={{flex:1,padding:"10px",borderRadius:10,border:"1.5px solid #d0c4f0",background:"white",color:"#6b5b9e",fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
              <button onClick={saveMeal} style={{flex:2,padding:"10px",borderRadius:10,border:"none",background:form.name.trim()?"linear-gradient(135deg,#8b5cf6,#6366f1)":"#d0c4f0",color:"white",fontSize:12,fontWeight:700,cursor:form.name.trim()?"pointer":"default",fontFamily:"inherit"}}>{editId?"Save Changes":"Add Meal"}</button>
            </div>
          </div>
        </div>
      )}

      {/* Library Modal */}
      {showLibrary&&(
        <div style={{position:"fixed",inset:0,background:"rgba(30,10,60,0.45)",backdropFilter:"blur(4px)",zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget){setShowLibrary(false);cancelPresetEdit();}}}>
          <div style={{background:"white",borderRadius:"22px 22px 0 0",padding:"20px 16px 34px",width:"100%",maxWidth:480,boxShadow:"0 -8px 40px rgba(139,92,246,0.18)",maxHeight:"92vh",overflowY:"auto"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div><div style={{fontSize:15,fontWeight:700,color:"#3d1f6b"}}>📚 My Foods Library</div><div style={{fontSize:11,color:"#9b87c2"}}>{presets.length} items</div></div>
              <button onClick={()=>{setShowLibrary(false);cancelPresetEdit();}} style={{background:"none",border:"none",fontSize:17,cursor:"pointer",color:"#9b87c2"}}>✕</button>
            </div>
            <div style={{background:"#faf6ff",border:"1.5px solid #d0c4f0",borderRadius:13,padding:"12px 12px 9px",marginBottom:14}}>
              <div style={{fontSize:11,fontWeight:700,color:"#3d1f6b",marginBottom:8}}>{editPresetId?"✏️ Edit Food":"＋ Add New Food"}</div>
              <div style={{display:"flex",gap:6,marginBottom:7}}>
                <input value={presetForm.emoji} onChange={e=>setPresetForm(f=>({...f,emoji:e.target.value}))} style={{...smallInput,width:46,textAlign:"center",fontSize:18,padding:"5px 3px",flexShrink:0}} maxLength={2}/>
                <input placeholder="Food name" value={presetForm.name} onChange={e=>setPresetForm(f=>({...f,name:e.target.value}))} style={{...smallInput,flex:1}}/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:6,marginBottom:6}}>
                {[{k:"protein",ph:"Protein g",c:"#FF6B6B"},{k:"carbs",ph:"Carbs g",c:"#4ECDC4"},{k:"fat",ph:"Fat g",c:"#FFD93D"}].map(({k,ph,c})=>(
                  <input key={k} type="number" min="0" placeholder={ph} value={presetForm[k]} onChange={e=>setPresetForm(f=>({...f,[k]:e.target.value}))} style={{...smallInput,borderColor:presetForm[k]?c:"#d0c4f0"}}/>
                ))}
              </div>
              <div style={{fontSize:11,color:"#8b5cf6",fontWeight:600,textAlign:"right",marginBottom:6}}>🔥 {calcCalories(presetForm.protein,presetForm.carbs,presetForm.fat)||"0"} kcal</div>
              <input placeholder="Notes" value={presetForm.notes} onChange={e=>setPresetForm(f=>({...f,notes:e.target.value}))} style={{...smallInput,width:"100%",boxSizing:"border-box"}}/>
              <div style={{display:"flex",gap:6,marginTop:8}}>
                {editPresetId&&<button onClick={cancelPresetEdit} style={{flex:1,padding:"8px",borderRadius:9,border:"1.5px solid #d0c4f0",background:"white",color:"#6b5b9e",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>}
                <button onClick={savePreset} style={{flex:2,padding:"8px",borderRadius:9,border:"none",background:presetForm.name.trim()?"linear-gradient(135deg,#8b5cf6,#6366f1)":"#d0c4f0",color:"white",fontSize:11,fontWeight:700,cursor:presetForm.name.trim()?"pointer":"default",fontFamily:"inherit"}}>{editPresetId?"Save Changes":"Add to Library"}</button>
              </div>
            </div>
            <input placeholder="Search…" value={libSearch} onChange={e=>setLibSearch(e.target.value)} style={{...inputStyle,fontSize:12,padding:"7px 10px"}}/>
            {libFiltered.map(p=>(
              <div key={p.id} style={{background:editPresetId===p.id?"#f0ebff":"white",border:`1.5px solid ${editPresetId===p.id?"#8b5cf6":"#e2d9f3"}`,borderRadius:11,padding:"9px 11px",marginBottom:6,display:"flex",alignItems:"center",gap:9}}>
                <span style={{fontSize:20}}>{p.emoji}</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:600,color:"#2d1b55",fontSize:12}}>{p.name}</div>
                  <div style={{fontSize:10,color:"#7b6aab"}}>{calcCalories(p.protein,p.carbs,p.fat)||"0"} kcal{p.protein?` · ${p.protein}g P`:""}{p.carbs?` · ${p.carbs}g C`:""}{p.fat?` · ${p.fat}g F`:""}</div>
                </div>
                <button onClick={()=>editPreset(p)} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,padding:2}}>✏️</button>
                <button onClick={()=>deletePreset(p.id)} style={{background:"none",border:"none",cursor:"pointer",fontSize:13,padding:2}}>🗑</button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* Report Modal */}
      {showReport&&(
        <div style={{position:"fixed",inset:0,background:"rgba(30,10,60,0.45)",backdropFilter:"blur(4px)",zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}} onClick={e=>{if(e.target===e.currentTarget)setShowReport(false);}}>
          <div style={{background:"white",borderRadius:"22px 22px 0 0",padding:"20px 16px 36px",width:"100%",maxWidth:480,boxShadow:"0 -8px 40px rgba(139,92,246,0.18)",maxHeight:"90vh",overflowY:"auto"}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
              <div><div style={{fontSize:15,fontWeight:700,color:"#3d1f6b"}}>📊 Download Report</div><div style={{fontSize:11,color:"#9b87c2"}}>Export your meal history</div></div>
              <button onClick={()=>setShowReport(false)} style={{background:"none",border:"none",fontSize:17,cursor:"pointer",color:"#9b87c2"}}>✕</button>
            </div>
            <div style={{marginBottom:14}}>
              <div style={{fontSize:11,color:"#9b87c2",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:7}}>Date Range</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {[["7","Last 7 days"],["14","Last 14 days"],["30","Last 30 days"],["custom","Custom"]].map(([val,label])=>(
                  <button key={val} onClick={()=>setReportRange(val)} style={{padding:"6px 12px",borderRadius:18,border:"1.5px solid",borderColor:reportRange===val?"#8b5cf6":"#d0c4f0",background:reportRange===val?"#8b5cf6":"white",color:reportRange===val?"white":"#6b5b9e",fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>{label}</button>
                ))}
              </div>
            </div>
            {reportRange==="custom"&&(
              <div style={{display:"flex",gap:8,marginBottom:14,alignItems:"center"}}>
                <div style={{flex:1}}><div style={{fontSize:10,color:"#9b87c2",marginBottom:4}}>From</div><input type="date" value={reportFrom} onChange={e=>setReportFrom(e.target.value)} style={{...smallInput,fontSize:12,width:"100%",boxSizing:"border-box"}}/></div>
                <div style={{color:"#9b87c2",marginTop:14}}>→</div>
                <div style={{flex:1}}><div style={{fontSize:10,color:"#9b87c2",marginBottom:4}}>To</div><input type="date" value={reportTo} onChange={e=>setReportTo(e.target.value)} style={{...smallInput,fontSize:12,width:"100%",boxSizing:"border-box"}}/></div>
              </div>
            )}
            <div style={{background:"#faf6ff",border:"1.5px solid #e2d9f3",borderRadius:13,padding:"13px 14px",marginBottom:16}}>
              <div style={{fontSize:11,color:"#9b87c2",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:10}}>Period Summary</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                {[{label:"Days with entries",val:reportDaysWithData.length,unit:"days",color:"#8b5cf6"},{label:"Avg daily calories",val:avgCalories,unit:"kcal",color:"#FF6B6B"},{label:"Total protein",val:reportTotals.protein,unit:"g",color:"#4ECDC4"},{label:"Total meals logged",val:reportDates.reduce((a,d)=>a+(allData[d]||[]).length,0),unit:"meals",color:"#FFD93D"}].map(s=>(
                  <div key={s.label} style={{background:"white",borderRadius:10,padding:"9px 11px"}}>
                    <div style={{fontSize:18,fontWeight:700,color:s.color}}>{s.val}<span style={{fontSize:11}}>{s.unit}</span></div>
                    <div style={{fontSize:10,color:"#9b87c2"}}>{s.label}</div>
                  </div>
                ))}
              </div>
            </div>
            <div style={{fontSize:11,color:"#9b87c2",fontWeight:600,textTransform:"uppercase",letterSpacing:"0.06em",marginBottom:8}}>Export As</div>
            <div style={{display:"flex",gap:9}}>
              <button onClick={()=>exportExcel(allData,reportDates)} disabled={reportDaysWithData.length===0} style={{flex:1,padding:"13px 8px",borderRadius:12,border:"none",background:reportDaysWithData.length>0?"linear-gradient(135deg,#1D6F42,#21A366)":"#d0c4f0",color:"white",fontSize:13,fontWeight:700,cursor:reportDaysWithData.length>0?"pointer":"default",fontFamily:"inherit",display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                <span style={{fontSize:22}}>📗</span><span>Excel (.xlsx)</span><span style={{fontSize:10,opacity:0.8}}>Full detail table</span>
              </button>
              <button onClick={()=>exportPDF(allData,reportDates)} disabled={reportDaysWithData.length===0} style={{flex:1,padding:"13px 8px",borderRadius:12,border:"none",background:reportDaysWithData.length>0?"linear-gradient(135deg,#c0392b,#e74c3c)":"#d0c4f0",color:"white",fontSize:13,fontWeight:700,cursor:reportDaysWithData.length>0?"pointer":"default",fontFamily:"inherit",display:"flex",flexDirection:"column",alignItems:"center",gap:4}}>
                <span style={{fontSize:22}}>📕</span><span>PDF / Print</span><span style={{fontSize:10,opacity:0.8}}>Printable report</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
