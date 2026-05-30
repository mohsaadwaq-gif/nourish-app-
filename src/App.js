import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── Design Tokens — Dark Premium Theme ──────────────────────────────────────
const C = {
  bg:        "#0d0d14",
  bg1:       "#13131f",
  bg2:       "#1a1a2e",
  bg3:       "#22223a",
  border:    "#2a2a45",
  borderHi:  "#a78bfa44",
  text:      "#f0eeff",
  textSub:   "#7c7a9e",
  textMuted: "#3d3b5c",
  accent:    "#a78bfa",
  accentDim: "#7c3aed",
  accentGlow:"rgba(167,139,250,0.25)",
  green:     "#34d399",
  red:       "#f87171",
  amber:     "#fbbf24",
  cyan:      "#67e8f9",
};

const MEAL_TYPES = ["Breakfast","Lunch","Dinner","Snack"];
const mealEmoji  = {Breakfast:"🌅",Lunch:"☀️",Dinner:"🌙",Snack:"🍎"};
const mealColor  = {Breakfast:"#f59e0b",Lunch:"#34d399",Dinner:"#a78bfa",Snack:"#f87171"};
const CAL_GOAL   = 2000; // fallback only — algorithm overrides this

// ─── Adaptive TDEE Algorithm ─────────────────────────────────────────────────
// How it works:
//   1. Take last 14 days of meal logs → calculate average daily calories
//   2. Take first + last weight entries in that window → calculate kg change
//   3. Convert weight change to kcal (1 kg fat ≈ 7700 kcal)
//   4. Derive true TDEE: avg_cal_eaten + (kcal_change / days)
//   5. Apply goal adjustment: lose = TDEE−300, maintain = TDEE, gain = TDEE+300
function calcAdaptiveTDEE(allData, weightLog, userGoal) {
  // Need at least 7 days of meal data + 2 weight entries
  const today = getTodayKey();
  const window = 14;
  const days = [];
  for (let i = 0; i < window; i++) {
    const d = new Date(); d.setDate(d.getDate() - i);
    days.push(d.toISOString().slice(0, 10));
  }

  // Filter days that actually have meal logs
  const loggedDays = days.filter(d => (allData[d] || []).length > 0);
  if (loggedDays.length < 7) return null; // not enough data yet

  // Average daily calories over logged days
  const totalCals = loggedDays.reduce((sum, d) => sum + getDayTotals(allData[d] || []).calories, 0);
  const avgCals = totalCals / loggedDays.length;

  // Weight entries within this window
  const sorted = [...weightLog].sort((a, b) => a.date.localeCompare(b.date));
  const windowWeights = sorted.filter(e => days.includes(e.date));
  if (windowWeights.length < 2) return null; // need at least 2 weigh-ins

  const firstW = windowWeights[0].weight;
  const lastW  = windowWeights[windowWeights.length - 1].weight;
  const daysBetween = Math.max(
    (new Date(windowWeights[windowWeights.length-1].date) - new Date(windowWeights[0].date)) / 86400000,
    1
  );

  // Weight change → kcal/day equivalent
  // Gaining weight means you ate MORE than TDEE → TDEE is lower than avgCals
  // Losing weight means you ate LESS than TDEE → TDEE is higher than avgCals
  const kgChange = lastW - firstW;
  const kcalPerDayFromWeight = (kgChange * 7700) / daysBetween;
  const estimatedTDEE = Math.round(avgCals - kcalPerDayFromWeight);

  // Apply goal adjustment
  const GOAL_ADJUSTMENTS = { lose: -500, maintain: 0, gain: 300 };
  const adjustment = GOAL_ADJUSTMENTS[userGoal] ?? -500;
  const recommendedGoal = Math.max(1200, estimatedTDEE + adjustment);

  return {
    tdee: estimatedTDEE,
    recommendedGoal,
    avgCals: Math.round(avgCals),
    kgChange: Math.round(kgChange * 10) / 10,
    daysAnalysed: loggedDays.length,
    confidence: loggedDays.length >= 10 ? "high" : loggedDays.length >= 7 ? "medium" : "low",
  };
}

const DEFAULT_PRESETS = [
  {id:"p1", name:"Black Coffee",      emoji:"☕",protein:0, carbs:0, fat:0, notes:"No milk/sugar"},
  {id:"p2", name:"Latte (Flat White)",emoji:"🥛",protein:4, carbs:8, fat:5, notes:"Whole milk"},
  {id:"p3", name:"Protein Shake",     emoji:"🥤",protein:25,carbs:5, fat:3, notes:"1 scoop whey + water"},
  {id:"p4", name:"Protein Bar",       emoji:"🍫",protein:20,carbs:22,fat:8, notes:"~220 kcal bar"},
  {id:"p5", name:"Banana",            emoji:"🍌",protein:1, carbs:27,fat:0, notes:"Medium banana"},
  {id:"p6", name:"Greek Yogurt",      emoji:"🥣",protein:17,carbs:6, fat:0, notes:"170g non-fat"},
  {id:"p7", name:"Boiled Eggs (x2)",  emoji:"🥚",protein:13,carbs:1, fat:10,notes:"Hard boiled"},
  {id:"p8", name:"Oatmeal",           emoji:"🥣",protein:5, carbs:27,fat:3, notes:"40g oats + water"},
  {id:"p9", name:"Chicken Breast",    emoji:"🍗",protein:31,carbs:0, fat:3, notes:"100g grilled"},
  {id:"p10",name:"Brown Rice",        emoji:"🍚",protein:3, carbs:44,fat:1, notes:"180g cooked"},
  {id:"p11",name:"Avocado Toast",     emoji:"🥑",protein:5, carbs:28,fat:12,notes:"1 slice sourdough"},
  {id:"p12",name:"Almonds (30g)",     emoji:"🌰",protein:6, carbs:5, fat:15,notes:"Small handful"},
];

// ─── Helpers ──────────────────────────────────────────────────────────────────
const generateId  = () => Math.random().toString(36).slice(2,10);
const getTodayKey = () => new Date().toISOString().slice(0,10);

function calcCalories(p,c,f){
  const pn=Number(p)||0,cn=Number(c)||0,fn=Number(f)||0;
  if(!pn&&!cn&&!fn) return "";
  return String(Math.round(pn*4+cn*4+fn*9));
}
function formatDateLong(d){
  const [y,m,day]=d.split("-");
  return new Date(y,m-1,day).toLocaleDateString("en-US",{weekday:"short",month:"short",day:"numeric"});
}
function formatDateShort(d){
  const [y,m,day]=d.split("-");
  return new Date(y,m-1,day).toLocaleDateString("en-US",{month:"short",day:"numeric"});
}
function getDayTotals(meals=[]){
  return meals.reduce((a,m)=>({
    calories:a.calories+(Number(m.calories)||0),
    protein: a.protein +(Number(m.protein)||0),
    carbs:   a.carbs   +(Number(m.carbs)||0),
    fat:     a.fat     +(Number(m.fat)||0),
  }),{calories:0,protein:0,carbs:0,fat:0});
}
function buildCalendarGrid(baseDate){
  const [y,m]=baseDate.split("-").map(Number);
  const first=new Date(y,m-1,1), last=new Date(y,m,0);
  const cells=[];
  for(let i=0;i<first.getDay();i++) cells.push(null);
  for(let d=1;d<=last.getDate();d++)
    cells.push(`${y}-${String(m).padStart(2,"0")}-${String(d).padStart(2,"0")}`);
  while(cells.length%7!==0) cells.push(null);
  return cells;
}
function calcStreak(allData){
  let streak=0, d=new Date();
  while(true){
    const key=d.toISOString().slice(0,10);
    if((allData[key]||[]).length>0){ streak++; d.setDate(d.getDate()-1); }
    else break;
  }
  return streak;
}
function getGreeting(){
  const h=new Date().getHours();
  if(h<12) return "Good morning";
  if(h<17) return "Good afternoon";
  return "Good evening";
}

const EMPTY_FORM   = {type:"Breakfast",name:"",calories:"",protein:"",carbs:"",fat:"",notes:""};
const EMPTY_PRESET = {name:"",emoji:"🍽",protein:"",carbs:"",fat:"",notes:""};

// ─── Shared input style (dark) ────────────────────────────────────────────────
const inp = {
  width:"100%",padding:"11px 14px",borderRadius:10,
  border:`1px solid ${C.border}`,fontSize:14,color:C.text,
  background:C.bg2,outline:"none",fontFamily:"'SF Pro Display',system-ui,sans-serif",
  marginBottom:10,boxSizing:"border-box",
};

// ─── Macro bar component ──────────────────────────────────────────────────────
function MacroBar({label,val,goal=200,color}){
  const pct=Math.min(val/goal,1)*100;
  return(
    <div style={{marginBottom:10}}>
      <div style={{display:"flex",justifyContent:"space-between",marginBottom:5}}>
        <span style={{fontSize:11,color:C.textSub,fontWeight:600,letterSpacing:"0.06em",textTransform:"uppercase"}}>{label}</span>
        <span style={{fontSize:11,color,fontWeight:700}}>{val}g</span>
      </div>
      <div style={{height:4,borderRadius:4,background:C.bg3,overflow:"hidden"}}>
        <div style={{height:"100%",width:`${pct}%`,borderRadius:4,background:color,
          boxShadow:`0 0 8px ${color}88`,transition:"width 0.6s ease"}}/>
      </div>
    </div>
  );
}

// ─── Weight Chart ─────────────────────────────────────────────────────────────
function WeightChart({entries}){
  if(entries.length<2) return null;
  const W=320,H=130,PAD={top:14,right:14,bottom:26,left:34};
  const vals=entries.map(e=>e.weight);
  const minV=Math.min(...vals),maxV=Math.max(...vals);
  const range=maxV-minV||1;
  const cW=W-PAD.left-PAD.right, cH=H-PAD.top-PAD.bottom;
  const px=i=>PAD.left+(i/(entries.length-1))*cW;
  const py=v=>PAD.top+cH-((v-minV)/range)*cH;
  const pts=entries.map((e,i)=>`${px(i)},${py(e.weight)}`).join(" ");
  const area=`${px(0)},${PAD.top+cH} ${pts} ${px(entries.length-1)},${PAD.top+cH}`;
  const yL=[minV,(minV+maxV)/2,maxV].map(v=>Math.round(v*10)/10);
  const xI=[0,Math.floor((entries.length-1)/2),entries.length-1];
  return(
    <svg width="100%" viewBox={`0 0 ${W} ${H}`} style={{overflow:"visible"}}>
      <defs>
        <linearGradient id="wg" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stopColor={C.accent} stopOpacity="0.3"/>
          <stop offset="100%" stopColor={C.accent} stopOpacity="0.0"/>
        </linearGradient>
      </defs>
      <polygon points={area} fill="url(#wg)"/>
      {yL.map((v,i)=>(
        <line key={i} x1={PAD.left} y1={py(v)} x2={W-PAD.right} y2={py(v)}
          stroke={C.border} strokeWidth="1" strokeDasharray="4,3"/>
      ))}
      <polyline points={pts} fill="none" stroke={C.accent} strokeWidth="2.5"
        strokeLinejoin="round" strokeLinecap="round"
        style={{filter:`drop-shadow(0 0 4px ${C.accent})`}}/>
      {entries.map((e,i)=>(
        <circle key={i} cx={px(i)} cy={py(e.weight)} r="4"
          fill={C.bg1} stroke={C.accent} strokeWidth="2.5"/>
      ))}
      {yL.map((v,i)=>(
        <text key={i} x={PAD.left-5} y={py(v)+4} textAnchor="end" fontSize="9" fill={C.textSub}>{v}</text>
      ))}
      {xI.map(i=>(
        <text key={i} x={px(i)} y={H-4} textAnchor="middle" fontSize="9" fill={C.textSub}>
          {formatDateShort(entries[i].date)}
        </text>
      ))}
    </svg>
  );
}

// ─── Calorie Ring ─────────────────────────────────────────────────────────────
function CalRing({calories,goal}){  // goal passed in from activeCalGoal
  const pct=Math.min(calories/goal,1);
  const R=54, circ=2*Math.PI*R;
  const over=calories>goal;
  const color=over?C.red:C.accent;
  return(
    <div style={{position:"relative",width:128,height:128,flexShrink:0}}>
      <svg width="128" height="128" viewBox="0 0 128 128">
        <circle cx="64" cy="64" r={R} fill="none" stroke={C.bg3} strokeWidth="10"/>
        <circle cx="64" cy="64" r={R} fill="none" stroke={color} strokeWidth="10"
          strokeDasharray={`${pct*circ} ${circ}`} strokeLinecap="round"
          transform="rotate(-90 64 64)"
          style={{filter:`drop-shadow(0 0 8px ${color})`,transition:"stroke-dasharray 0.6s ease"}}/>
      </svg>
      <div style={{position:"absolute",inset:0,display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
        <span style={{fontSize:26,fontWeight:800,color:C.text,lineHeight:1,fontFamily:"'SF Pro Display',system-ui,sans-serif"}}>{calories}</span>
        <span style={{fontSize:10,color:C.textSub,marginTop:2,letterSpacing:"0.08em"}}>KCAL</span>
        <span style={{fontSize:9,color:C.textMuted,marginTop:1}}>{Math.round(pct*100)}% of {goal}</span>
      </div>
    </div>
  );
}

// ─── Streak Card ──────────────────────────────────────────────────────────────
function StreakCard({streak}){
  return(
    <div style={{background:`linear-gradient(135deg, #1e1b3a, #2a1f4e)`,
      border:`1px solid ${C.borderHi}`,borderRadius:18,padding:"18px 20px",
      display:"flex",alignItems:"center",justifyContent:"space-between",
      boxShadow:`0 0 30px ${C.accentGlow}`,marginBottom:14}}>
      <div>
        <div style={{fontSize:11,color:C.textSub,fontWeight:600,letterSpacing:"0.1em",textTransform:"uppercase",marginBottom:4}}>
          Current Streak
        </div>
        <div style={{display:"flex",alignItems:"baseline",gap:6}}>
          <span style={{fontSize:44,fontWeight:800,color:C.text,lineHeight:1,fontFamily:"'SF Pro Display',system-ui,sans-serif"}}>{streak}</span>
          <span style={{fontSize:16,color:C.textSub,fontWeight:500}}>days</span>
        </div>
        <div style={{fontSize:11,color:streak>0?C.accent:C.textMuted,marginTop:4}}>
          {streak===0?"Log today to start your streak!":streak===1?"Great start — keep it going!":streak<7?`${7-streak} days until a week!`:`🔥 ${streak} day streak — incredible!`}
        </div>
      </div>
      <div style={{fontSize:52,filter:streak>0?`drop-shadow(0 0 12px #f59e0b)`:"grayscale(1) opacity(0.3)"}}>
        🔥
      </div>
    </div>
  );
}

// ─── Export helpers ───────────────────────────────────────────────────────────
function exportPDF(allData,dateRange){
  const rows=dateRange.flatMap(date=>{
    const meals=allData[date]||[]; if(!meals.length) return [];
    const t=getDayTotals(meals);
    return[
      `<tr style="background:#1a1a2e"><td colspan="7" style="padding:6px 8px;font-weight:700;color:#a78bfa">${formatDateLong(date)}</td></tr>`,
      ...meals.map(m=>`<tr><td style="padding:4px 8px">${m.type}</td><td style="padding:4px 8px;font-weight:600">${m.name}</td><td style="padding:4px 8px;text-align:right">${m.calories||""}</td><td style="padding:4px 8px;text-align:right">${m.protein||""}</td><td style="padding:4px 8px;text-align:right">${m.carbs||""}</td><td style="padding:4px 8px;text-align:right">${m.fat||""}</td><td style="padding:4px 8px;color:#999;font-size:11px">${m.notes||""}</td></tr>`),
      `<tr style="font-style:italic;font-size:11px"><td colspan="2" style="padding:4px 8px;color:#666">Day total</td><td style="padding:4px 8px;text-align:right;color:#a78bfa;font-weight:700">${t.calories}</td><td style="padding:4px 8px;text-align:right;color:#f87171">${t.protein}g</td><td style="padding:4px 8px;text-align:right;color:#34d399">${t.carbs}g</td><td style="padding:4px 8px;text-align:right;color:#fbbf24">${t.fat}g</td><td></td></tr>`,
    ];
  });
  const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Nourish Report</title><style>body{font-family:system-ui,sans-serif;padding:28px;color:#1a1a2e;font-size:13px}h1{font-size:22px;color:#7c3aed}table{width:100%;border-collapse:collapse}th{background:#0d0d14;color:white;padding:7px 8px;text-align:left;font-size:12px}tr:nth-child(even){background:#f8f8ff}@media print{button{display:none}}</style></head><body><h1>🍽 Nourish — Meal Report</h1><p>Generated ${new Date().toLocaleDateString("en-US",{year:"numeric",month:"long",day:"numeric"})}</p><table><thead><tr><th>Type</th><th>Food</th><th>Calories</th><th>Protein</th><th>Carbs</th><th>Fat</th><th>Notes</th></tr></thead><tbody>${rows.join("")}</tbody></table></body></html>`;
  const win=window.open("","_blank");
  win.document.write(html); win.document.close();
  setTimeout(()=>win.print(),400);
}
function exportExcel(allData,dateRange){
  const rows=[["Date","Meal Type","Food Name","Calories (kcal)","Protein (g)","Carbs (g)","Fat (g)","Notes"]];
  dateRange.forEach(date=>{
    const meals=allData[date]||[]; if(!meals.length) return;
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

// ═══════════════════════════════════════════════════════════════════════════════
export default function App(){
  const [allData,  setAllData]  = useState({});
  const [presets,  setPresets]  = useState(DEFAULT_PRESETS);
  const [selDate,  setSelDate]  = useState(getTodayKey());
  const [activeTab,setActiveTab]= useState("log");
  const today = getTodayKey();
  const [calMonth, setCalMonth] = useState(today.slice(0,7));

  // Weight
  const [weightLog,   setWeightLog]   = useState([]);
  const [weightInput, setWeightInput] = useState("");
  const [weightDate,  setWeightDate]  = useState(today);
  const [weightGoal,  setWeightGoal]  = useState("");
  const [editWeightId,setEditWeightId]= useState(null);

  // ── User Goal + Profile (for adaptive algorithm) ─────────────────────────
  const [userGoal,    setUserGoal]    = useState("lose");   // lose | maintain | gain
  const [showGoals,   setShowGoals]   = useState(false);

  // Meal modal
  const [showForm,     setShowForm]     = useState(false);
  const [form,         setForm]         = useState(EMPTY_FORM);
  const [editId,       setEditId]       = useState(null);
  const [presetSearch, setPresetSearch] = useState("");

  // Library
  const [showLibrary,   setShowLibrary]   = useState(false);
  const [presetForm,    setPresetForm]    = useState(EMPTY_PRESET);
  const [editPresetId,  setEditPresetId]  = useState(null);
  const [libSearch,     setLibSearch]     = useState("");

  // Report
  const [showReport,  setShowReport]  = useState(false);
  const [reportRange, setReportRange] = useState("7");
  const [reportFrom,  setReportFrom]  = useState("");
  const [reportTo,    setReportTo]    = useState(today);

  // Camera
  const [cameraStep,    setCameraStep]    = useState("idle");
  const [capturedImage, setCapturedImage] = useState(null);
  const [analysisError, setAnalysisError] = useState("");
  const videoRef     = useRef(null);
  const streamRef    = useRef(null);
  const fileInputRef = useRef(null);

  // ── Persist ──────────────────────────────────────────────────────────────
  useEffect(()=>{
    const d=localStorage.getItem("nourish_data_v3");
    const p=localStorage.getItem("nourish_presets_v1");
    const w=localStorage.getItem("nourish_weight_v1");
    const wg=localStorage.getItem("nourish_weight_goal_v1");
    if(d) try{setAllData(JSON.parse(d));}catch{}
    if(p) try{setPresets(JSON.parse(p));}catch{}
    if(w) try{setWeightLog(JSON.parse(w));}catch{}
    if(wg) setWeightGoal(wg);
    const ug=localStorage.getItem("nourish_user_goal_v1");
    if(ug) setUserGoal(ug);
  },[]);
  useEffect(()=>{localStorage.setItem("nourish_data_v3",JSON.stringify(allData));},[allData]);
  useEffect(()=>{localStorage.setItem("nourish_presets_v1",JSON.stringify(presets));},[presets]);
  useEffect(()=>{localStorage.setItem("nourish_weight_v1",JSON.stringify(weightLog));},[weightLog]);
  useEffect(()=>{localStorage.setItem("nourish_weight_goal_v1",weightGoal);},[weightGoal]);
  useEffect(()=>{localStorage.setItem("nourish_user_goal_v1",userGoal);},[userGoal]);

  // ── Camera ────────────────────────────────────────────────────────────────
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
            {type:"text",text:`Analyze this food photo. Return ONLY valid JSON:\n{"name":"short meal name","protein":number,"carbs":number,"fat":number,"notes":"brief note"}`}
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

  // ── Meals ─────────────────────────────────────────────────────────────────
  const todayMeals=allData[selDate]||[];
  const totals=getDayTotals(todayMeals);
  const streak=calcStreak(allData);

  // ── Adaptive Algorithm ────────────────────────────────────────────────────
  const adaptive = calcAdaptiveTDEE(allData, weightLog, userGoal);
  const activeCalGoal = adaptive ? adaptive.recommendedGoal : CAL_GOAL;

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

  // ── Presets ───────────────────────────────────────────────────────────────
  function savePreset(){
    if(!presetForm.name.trim())return;
    if(editPresetId) setPresets(ps=>ps.map(p=>p.id===editPresetId?{...presetForm,id:editPresetId}:p));
    else setPresets(ps=>[...ps,{...presetForm,id:"u"+generateId()}]);
    setPresetForm(EMPTY_PRESET);setEditPresetId(null);
  }
  function editPreset(p){setPresetForm({...p});setEditPresetId(p.id);}
  function deletePreset(id){setPresets(ps=>ps.filter(p=>p.id!==id));if(editPresetId===id){setPresetForm(EMPTY_PRESET);setEditPresetId(null);}}
  function cancelPresetEdit(){setPresetForm(EMPTY_PRESET);setEditPresetId(null);}
  const filteredPresets=presets.filter(p=>!presetSearch||p.name.toLowerCase().includes(presetSearch.toLowerCase()));
  const libFiltered=presets.filter(p=>!libSearch||p.name.toLowerCase().includes(libSearch.toLowerCase()));

  // ── Weight ────────────────────────────────────────────────────────────────
  const sortedWeight=[...weightLog].sort((a,b)=>a.date.localeCompare(b.date));
  const latestWeight=sortedWeight.length>0?sortedWeight[sortedWeight.length-1].weight:null;
  const firstWeight =sortedWeight.length>0?sortedWeight[0].weight:null;
  const weightChange=(latestWeight&&firstWeight)?Math.round((latestWeight-firstWeight)*10)/10:null;

  function saveWeight(){
    if(!weightInput||isNaN(Number(weightInput)))return;
    const w=Number(weightInput);
    if(editWeightId){
      setWeightLog(wl=>wl.map(e=>e.id===editWeightId?{...e,weight:w,date:weightDate}:e));
      setEditWeightId(null);
    } else {
      const exists=weightLog.find(e=>e.date===weightDate);
      if(exists) setWeightLog(wl=>wl.map(e=>e.date===weightDate?{...e,weight:w}:e));
      else setWeightLog(wl=>[...wl,{id:generateId(),date:weightDate,weight:w}]);
    }
    setWeightInput("");setWeightDate(today);
  }
  function deleteWeight(id){setWeightLog(wl=>wl.filter(e=>e.id!==id));}
  function startEditWeight(e){setEditWeightId(e.id);setWeightInput(String(e.weight));setWeightDate(e.date);}
  function cancelEditWeight(){setEditWeightId(null);setWeightInput("");setWeightDate(today);}

  // ── Calendar ──────────────────────────────────────────────────────────────
  const calGrid=buildCalendarGrid(calMonth+"-01");
  function prevMonth(){const [y,m]=calMonth.split("-").map(Number);const d=new Date(y,m-2,1);setCalMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}
  function nextMonth(){const [y,m]=calMonth.split("-").map(Number);const d=new Date(y,m,1);setCalMonth(`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`);}
  const monthLabel=new Date(calMonth+"-01").toLocaleDateString("en-US",{month:"long",year:"numeric"});

  // ── Report ────────────────────────────────────────────────────────────────
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

  // ─── Dark modal base ──────────────────────────────────────────────────────
  const modalBase = {background:C.bg1,borderRadius:"24px 24px 0 0",padding:"22px 18px 36px",
    width:"100%",maxWidth:480,boxShadow:"0 -12px 60px rgba(0,0,0,0.8)",
    maxHeight:"92vh",overflowY:"auto",border:`1px solid ${C.border}`,borderBottom:"none"};

  const darkInp = {...inp,background:C.bg2,border:`1px solid ${C.border}`,color:C.text};
  const darkInpSm = {...darkInp,fontSize:13,padding:"10px 12px",marginBottom:0};

  // ─────────────────────────────────────────────────────────────────────────
  return(
    <div style={{minHeight:"100vh",background:C.bg,fontFamily:"'SF Pro Display',system-ui,-apple-system,sans-serif",color:C.text}}>

      {/* ── Header ── */}
      <div style={{background:"rgba(13,13,20,0.95)",backdropFilter:"blur(20px)",
        borderBottom:`1px solid ${C.border}`,padding:"14px 18px 12px",
        position:"sticky",top:0,zIndex:50,
        display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:20,fontWeight:800,letterSpacing:"-0.5px",
            background:"linear-gradient(135deg,#a78bfa,#7c3aed)",
            WebkitBackgroundClip:"text",WebkitTextFillColor:"transparent"}}>
            🍽 Nourish
          </div>
          <div style={{fontSize:10,color:C.textSub,letterSpacing:"0.12em",textTransform:"uppercase",marginTop:1}}>
            {getGreeting()}
          </div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <button onClick={()=>setShowGoals(true)} style={{padding:"6px 11px",borderRadius:9,border:`1px solid ${adaptive?C.green:C.border}`,background:adaptive?`${C.green}22`:C.bg2,color:adaptive?C.green:C.textSub,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>{adaptive?"🧬":"🎯"}</button>
          <button onClick={()=>setShowLibrary(true)} style={{padding:"6px 11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.bg2,color:C.textSub,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>📚</button>
          <button onClick={()=>setShowReport(true)} style={{padding:"6px 11px",borderRadius:9,border:`1px solid ${C.border}`,background:C.bg2,color:C.textSub,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>📊</button>
          <input type="date" value={selDate} onChange={e=>setSelDate(e.target.value)}
            style={{border:`1px solid ${C.border}`,borderRadius:9,padding:"5px 9px",fontSize:11,
              color:C.textSub,background:C.bg2,outline:"none",fontFamily:"inherit"}}/>
        </div>
      </div>

      {/* ── Tabs ── */}
      <div style={{display:"flex",padding:"12px 18px 0",gap:6,overflowX:"auto"}}>
        {[["log","Log"],["calendar","Calendar"],["weight","Weight"]].map(([tab,label])=>(
          <button key={tab} onClick={()=>setActiveTab(tab)} style={{
            padding:"7px 16px",borderRadius:20,border:`1px solid ${activeTab===tab?C.accent:C.border}`,
            background:activeTab===tab?C.accentDim:"transparent",
            color:activeTab===tab?C.text:C.textSub,
            fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit",
            whiteSpace:"nowrap",flexShrink:0,
            boxShadow:activeTab===tab?`0 0 14px ${C.accentGlow}`:"none",
            transition:"all 0.2s",
          }}>{label}</button>
        ))}
      </div>

      {/* ══════════════════ DAILY LOG ══════════════════ */}
      {activeTab==="log"&&(
        <div style={{padding:"14px 18px 90px"}}>

          {/* Streak */}
          <StreakCard streak={streak}/>

          {/* Adaptive insight card */}
          {adaptive ? (
            <div style={{background:`linear-gradient(135deg,${C.bg1},${C.bg2})`,
              border:`1px solid ${C.green}44`,borderRadius:16,padding:"12px 16px",
              marginBottom:14,display:"flex",alignItems:"center",gap:12,
              boxShadow:`0 0 20px ${C.green}18`}}>
              <div style={{fontSize:26}}>🧬</div>
              <div style={{flex:1}}>
                <div style={{fontSize:11,color:C.green,fontWeight:700,letterSpacing:"0.06em",textTransform:"uppercase",marginBottom:3}}>
                  Adaptive Algorithm Active
                </div>
                <div style={{fontSize:12,color:C.text}}>
                  Your true TDEE is <span style={{color:C.cyan,fontWeight:700}}>{adaptive.tdee} kcal</span>
                  {" · "}Goal: <span style={{color:C.accent,fontWeight:700}}>{activeCalGoal} kcal</span>
                  {" · "}<span style={{color:C.textSub}}>{adaptive.daysAnalysed} days of data</span>
                </div>
                <div style={{fontSize:10,color:C.textMuted,marginTop:3}}>
                  {userGoal==="lose"?"Eating 500 kcal below your TDEE for steady fat loss":
                   userGoal==="gain"?"Eating 300 kcal above your TDEE for muscle gain":
                   "Eating at your TDEE to maintain weight"}
                   {" · "}Confidence: <span style={{color:adaptive.confidence==="high"?C.green:adaptive.confidence==="medium"?C.amber:C.red}}>{adaptive.confidence}</span>
                </div>
              </div>
              <button onClick={()=>setShowGoals(true)} style={{background:C.bg3,border:`1px solid ${C.border}`,borderRadius:8,padding:"5px 10px",color:C.textSub,fontSize:11,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>Edit</button>
            </div>
          ) : (
            <div style={{background:C.bg1,border:`1px solid ${C.amber}33`,borderRadius:16,padding:"11px 14px",
              marginBottom:14,display:"flex",alignItems:"center",gap:10,cursor:"pointer"}}
              onClick={()=>setShowGoals(true)}>
              <div style={{fontSize:20}}>🎯</div>
              <div style={{flex:1}}>
                <div style={{fontSize:12,color:C.amber,fontWeight:600}}>Set your goal to activate smart calorie targets</div>
                <div style={{fontSize:10,color:C.textMuted,marginTop:2}}>
                  Log 7+ days of meals + 2 weigh-ins → algorithm calculates your real TDEE
                </div>
              </div>
              <div style={{fontSize:14,color:C.textMuted}}>›</div>
            </div>
          )}

          {/* Calorie ring + macros */}
          <div style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:20,
            padding:"20px",marginBottom:14,
            boxShadow:`0 4px 30px rgba(0,0,0,0.4)`}}>
            <div style={{display:"flex",gap:20,alignItems:"center",marginBottom:16}}>
              <CalRing calories={totals.calories} goal={activeCalGoal}/>
              <div style={{flex:1}}>
                <div style={{fontSize:11,color:C.textSub,fontWeight:600,letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:12}}>
                  {selDate===today?"Today":"Selected Day"}
                </div>
                <MacroBar label="Protein" val={totals.protein} goal={150} color={C.red}/>
                <MacroBar label="Carbs"   val={totals.carbs}   goal={250} color={C.green}/>
                <MacroBar label="Fat"     val={totals.fat}     goal={65}  color={C.amber}/>
              </div>
            </div>
          </div>

          {/* Meals by type */}
          {MEAL_TYPES.map(type=>{
            const meals=todayMeals.filter(m=>m.type===type);
            const color=mealColor[type];
            return(
              <div key={type} style={{marginBottom:10}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                  <div style={{width:3,height:16,borderRadius:2,background:color,boxShadow:`0 0 6px ${color}`}}/>
                  <span style={{fontSize:12,fontWeight:700,color:C.text,letterSpacing:"0.02em"}}>{mealEmoji[type]} {type}</span>
                  <span style={{fontSize:10,color:C.textSub,fontWeight:400,marginLeft:2}}>
                    {meals.length>0?`${meals.reduce((a,m)=>a+(Number(m.calories)||0),0)} kcal`:"—"}
                  </span>
                </div>
                {meals.map(meal=>(
                  <div key={meal.id} style={{background:C.bg1,border:`1px solid ${C.border}`,
                    borderRadius:14,padding:"11px 14px",marginBottom:5,
                    display:"flex",justifyContent:"space-between",alignItems:"center",
                    transition:"border-color 0.2s"}}>
                    <div>
                      <div style={{fontWeight:600,color:C.text,fontSize:13}}>{meal.name}</div>
                      <div style={{fontSize:10,color:C.textSub,marginTop:3,display:"flex",gap:8}}>
                        {meal.calories&&<span style={{color:C.accent}}>{meal.calories} kcal</span>}
                        {meal.protein&&<span style={{color:C.red}}>{meal.protein}g P</span>}
                        {meal.carbs&&<span style={{color:C.green}}>{meal.carbs}g C</span>}
                        {meal.fat&&<span style={{color:C.amber}}>{meal.fat}g F</span>}
                      </div>
                    </div>
                    <div style={{display:"flex",gap:4}}>
                      <button onClick={()=>openEdit(meal)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,cursor:"pointer",fontSize:11,padding:"5px 8px",color:C.textSub}}>✏️</button>
                      <button onClick={()=>deleteMeal(meal.id)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,cursor:"pointer",fontSize:11,padding:"5px 8px",color:C.textSub}}>🗑</button>
                    </div>
                  </div>
                ))}
                {meals.length===0&&(
                  <div style={{border:`1px dashed ${C.border}`,borderRadius:14,padding:"12px 14px",
                    fontSize:11,color:C.textMuted,fontStyle:"italic"}}>Nothing logged yet</div>
                )}
              </div>
            );
          })}
        </div>
      )}

      {/* ══════════════════ CALENDAR ══════════════════ */}
      {activeTab==="calendar"&&(
        <div style={{padding:"14px 18px 90px"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:14}}>
            <button onClick={prevMonth} style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:10,
              padding:"7px 14px",cursor:"pointer",color:C.text,fontSize:14,fontFamily:"inherit"}}>‹</button>
            <div style={{fontWeight:700,color:C.text,fontSize:15}}>{monthLabel}</div>
            <button onClick={nextMonth} style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:10,
              padding:"7px 14px",cursor:"pointer",color:C.text,fontSize:14,fontFamily:"inherit"}}>›</button>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3,marginBottom:4}}>
            {["S","M","T","W","T","F","S"].map((d,i)=>(
              <div key={i} style={{textAlign:"center",fontSize:10,color:C.textMuted,fontWeight:600,padding:"3px 0"}}>{d}</div>
            ))}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(7,1fr)",gap:3}}>
            {calGrid.map((dateKey,i)=>{
              if(!dateKey)return<div key={i}/>;
              const meals=allData[dateKey]||[];
              const t=getDayTotals(meals);
              const isToday=dateKey===today,isSelected=dateKey===selDate,hasMeals=meals.length>0;
              const pct=Math.min(t.calories/activeCalGoal,1);
              const ringColor=pct>1?C.red:pct>0.6?C.green:C.accent;
              return(
                <div key={dateKey} onClick={()=>{setSelDate(dateKey);setActiveTab("log");}}
                  style={{background:isSelected?C.accentDim:isToday?C.bg2:C.bg1,
                    border:`1px solid ${isSelected?C.accent:isToday?C.borderHi:C.border}`,
                    borderRadius:10,padding:"6px 3px 5px",cursor:"pointer",
                    minHeight:54,display:"flex",flexDirection:"column",alignItems:"center",gap:2,
                    boxShadow:isSelected?`0 0 14px ${C.accentGlow}`:"none",
                    transition:"all 0.15s"}}>
                  <div style={{fontSize:11,fontWeight:isToday||isSelected?700:400,
                    color:isSelected?C.text:isToday?C.accent:C.textSub}}>
                    {Number(dateKey.slice(8))}
                  </div>
                  {hasMeals&&(
                    <>
                      <svg width="26" height="26" viewBox="0 0 26 26">
                        <circle cx="13" cy="13" r="9" fill="none" stroke={C.bg3} strokeWidth="3"/>
                        <circle cx="13" cy="13" r="9" fill="none" stroke={ringColor} strokeWidth="3"
                          strokeDasharray={`${pct*56.5} 56.5`} strokeLinecap="round"
                          transform="rotate(-90 13 13)"
                          style={{filter:`drop-shadow(0 0 3px ${ringColor})`}}/>
                      </svg>
                      <div style={{fontSize:8,color:isSelected?C.text:C.accent,fontWeight:600,lineHeight:1}}>{t.calories}</div>
                    </>
                  )}
                  {!hasMeals&&<div style={{width:5,height:5,borderRadius:"50%",background:C.bg3,marginTop:3}}/>}
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",gap:12,marginTop:14,justifyContent:"center",flexWrap:"wrap"}}>
            {[[C.accent,"On track"],[C.green,"60–100%"],[C.red,"Over goal"],[C.bg3,"No entries"]].map(([c,l])=>(
              <div key={l} style={{display:"flex",alignItems:"center",gap:5}}>
                <div style={{width:8,height:8,borderRadius:"50%",background:c,boxShadow:`0 0 5px ${c}`}}/>
                <span style={{fontSize:10,color:C.textSub}}>{l}</span>
              </div>
            ))}
          </div>
          {selDate&&(
            <div style={{marginTop:14,background:C.bg1,borderRadius:16,padding:"14px",border:`1px solid ${C.border}`}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10}}>
                <div style={{fontWeight:700,color:C.text,fontSize:13}}>{selDate===today?"Today — ":""}{formatDateLong(selDate)}</div>
                <button onClick={()=>setActiveTab("log")} style={{background:C.accentDim,border:"none",borderRadius:8,
                  padding:"5px 12px",color:C.text,fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",
                  boxShadow:`0 0 10px ${C.accentGlow}`}}>View →</button>
              </div>
              {(allData[selDate]||[]).length===0
                ?<div style={{fontSize:12,color:C.textMuted,fontStyle:"italic"}}>No meals logged.</div>
                :(allData[selDate]||[]).map(m=>(
                  <div key={m.id} style={{display:"flex",justifyContent:"space-between",
                    borderBottom:`1px solid ${C.border}`,padding:"6px 0",fontSize:12}}>
                    <span style={{color:C.textSub}}>{mealEmoji[m.type]} {m.name}</span>
                    <span style={{color:C.accent,fontWeight:600}}>{m.calories||"—"} kcal</span>
                  </div>
                ))
              }
            </div>
          )}
        </div>
      )}

      {/* ══════════════════ WEIGHT ══════════════════ */}
      {activeTab==="weight"&&(
        <div style={{padding:"14px 18px 90px"}}>

          {/* Stats */}
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:14}}>
            {[
              {label:"Current",val:latestWeight?`${latestWeight}`:"-",unit:latestWeight?"kg":"",color:C.accent},
              {label:"Change",val:weightChange!==null?(weightChange>0?`+${weightChange}`:String(weightChange)):"-",unit:weightChange!==null?"kg":"",color:weightChange===null?C.textSub:weightChange<0?C.green:weightChange>0?C.red:C.textSub},
              {label:"Goal",val:weightGoal?`${weightGoal}`:"-",unit:weightGoal?"kg":"",color:C.amber},
            ].map(s=>(
              <div key={s.label} style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:14,
                padding:"13px 10px",textAlign:"center",
                boxShadow:s.val!=="-"?`0 0 20px ${s.color}22`:"none"}}>
                <div style={{fontSize:22,fontWeight:800,color:s.color,lineHeight:1}}>
                  {s.val}<span style={{fontSize:11,fontWeight:500}}>{s.unit}</span>
                </div>
                <div style={{fontSize:10,color:C.textSub,marginTop:4,letterSpacing:"0.06em",textTransform:"uppercase"}}>{s.label}</div>
              </div>
            ))}
          </div>

          {/* Goal input */}
          <div style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:14,
            padding:"13px 16px",marginBottom:12,display:"flex",alignItems:"center",gap:12}}>
            <span style={{fontSize:13,color:C.textSub,fontWeight:600,whiteSpace:"nowrap"}}>🎯 Goal</span>
            <input type="number" min="0" step="0.1" placeholder="e.g. 75" value={weightGoal}
              onChange={e=>setWeightGoal(e.target.value)}
              style={{...darkInpSm,flex:1,marginBottom:0}}/>
            <span style={{fontSize:12,color:C.textMuted}}>kg</span>
          </div>

          {/* Log form */}
          <div style={{background:C.bg1,border:`1px solid ${C.borderHi}`,borderRadius:16,
            padding:"16px",marginBottom:14,
            boxShadow:`0 0 20px ${C.accentGlow}`}}>
            <div style={{fontSize:12,fontWeight:700,color:C.text,marginBottom:14,letterSpacing:"0.02em"}}>
              {editWeightId?"✏️ Edit Entry":"➕ Log Weight"}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:12}}>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                <div style={{fontSize:10,fontWeight:600,color:C.textSub,textTransform:"uppercase",letterSpacing:"0.08em"}}>Date</div>
                <input type="date" value={weightDate} onChange={e=>setWeightDate(e.target.value)}
                  style={{...darkInpSm,width:"100%",boxSizing:"border-box",height:42}}/>
              </div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                <div style={{fontSize:10,fontWeight:600,color:C.textSub,textTransform:"uppercase",letterSpacing:"0.08em"}}>Weight (kg)</div>
                <input type="number" min="0" step="0.1" placeholder="78.5" value={weightInput}
                  onChange={e=>setWeightInput(e.target.value)}
                  style={{...darkInpSm,width:"100%",boxSizing:"border-box",height:42,
                    borderColor:weightInput?C.accent:C.border,
                    boxShadow:weightInput?`0 0 10px ${C.accentGlow}`:"none"}}/>
              </div>
            </div>
            <div style={{display:"flex",gap:8}}>
              {editWeightId&&(
                <button onClick={cancelEditWeight} style={{flex:1,padding:"11px",borderRadius:11,
                  border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,
                  fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
              )}
              <button onClick={saveWeight} style={{flex:2,padding:"11px",borderRadius:11,border:"none",
                background:weightInput?`linear-gradient(135deg,${C.accent},${C.accentDim})`:"#1e1e2e",
                color:weightInput?C.text:C.textMuted,fontSize:12,fontWeight:700,
                cursor:weightInput?"pointer":"default",fontFamily:"inherit",
                boxShadow:weightInput?`0 0 16px ${C.accentGlow}`:"none",
                transition:"all 0.2s"}}>
                {editWeightId?"Save Changes":"Log Weight"}
              </button>
            </div>
          </div>

          {/* Chart */}
          {sortedWeight.length>=2&&(
            <div style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:16,
              padding:"16px",marginBottom:14}}>
              <div style={{fontSize:12,fontWeight:700,color:C.text,marginBottom:12}}>📈 Trend</div>
              <WeightChart entries={sortedWeight}/>
              {weightGoal&&latestWeight&&(
                <div style={{marginTop:10,fontSize:11,color:C.textSub,textAlign:"center"}}>
                  <span style={{color:latestWeight>Number(weightGoal)?C.red:C.green,fontWeight:600}}>
                    {Math.abs(Math.round((latestWeight-Number(weightGoal))*10)/10)} kg
                  </span>
                  {" "}{latestWeight>Number(weightGoal)?"to reach goal":"below goal"} · target {weightGoal} kg
                </div>
              )}
            </div>
          )}
          {sortedWeight.length===1&&(
            <div style={{textAlign:"center",color:C.textSub,fontSize:12,fontStyle:"italic",padding:"10px 0"}}>
              Log one more entry to see your trend.
            </div>
          )}

          {/* List */}
          {sortedWeight.length>0&&(
            <div style={{background:C.bg1,border:`1px solid ${C.border}`,borderRadius:16,padding:"14px"}}>
              <div style={{fontSize:12,fontWeight:700,color:C.text,marginBottom:12}}>All Entries</div>
              {[...sortedWeight].reverse().map((e,i,arr)=>{
                const prev=arr[i+1];
                const diff=prev?Math.round((e.weight-prev.weight)*10)/10:null;
                return(
                  <div key={e.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",
                    padding:"8px 0",borderBottom:`1px solid ${C.border}`}}>
                    <div>
                      <div style={{fontSize:14,fontWeight:700,color:C.text}}>{e.weight} <span style={{fontSize:11,color:C.textSub}}>kg</span></div>
                      <div style={{fontSize:10,color:C.textSub}}>{formatDateLong(e.date)}</div>
                    </div>
                    <div style={{display:"flex",alignItems:"center",gap:10}}>
                      {diff!==null&&(
                        <span style={{fontSize:12,fontWeight:700,
                          color:diff<0?C.green:diff>0?C.red:C.textSub}}>
                          {diff>0?"+":""}{diff} kg
                        </span>
                      )}
                      <button onClick={()=>startEditWeight(e)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:7,cursor:"pointer",fontSize:11,padding:"4px 8px",color:C.textSub}}>✏️</button>
                      <button onClick={()=>deleteWeight(e.id)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:7,cursor:"pointer",fontSize:11,padding:"4px 8px",color:C.textSub}}>🗑</button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}
          {sortedWeight.length===0&&(
            <div style={{textAlign:"center",color:C.textMuted,fontSize:13,fontStyle:"italic",marginTop:24,padding:20}}>
              No weight entries yet.<br/>Log your first weigh-in above!
            </div>
          )}
        </div>
      )}

      {/* FAB */}
      {activeTab!=="weight"&&(
        <button onClick={openAdd} style={{position:"fixed",bottom:26,right:20,
          width:54,height:54,borderRadius:"50%",
          background:`linear-gradient(135deg,${C.accent},${C.accentDim})`,
          border:"none",color:"white",fontSize:26,cursor:"pointer",
          boxShadow:`0 4px 20px ${C.accentGlow}, 0 0 30px ${C.accentGlow}`,
          display:"flex",alignItems:"center",justifyContent:"center",zIndex:100,
          fontWeight:300,transition:"transform 0.15s"}}
          onTouchStart={e=>e.currentTarget.style.transform="scale(0.92)"}
          onTouchEnd={e=>e.currentTarget.style.transform="scale(1)"}>+</button>
      )}
      <input ref={fileInputRef} type="file" accept="image/*" style={{display:"none"}} onChange={handleFileUpload}/>

      {/* ══ ADD/EDIT MEAL MODAL ══ */}
      {showForm&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)",
          zIndex:200,display:"flex",alignItems:"flex-end",justifyContent:"center"}}
          onClick={e=>{if(e.target===e.currentTarget)setShowForm(false);}}>
          <div style={modalBase}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div style={{fontSize:16,fontWeight:700,color:C.text}}>{editId?"Edit Meal":"Add Meal"}</div>
              <button onClick={()=>setShowForm(false)} style={{background:C.bg2,border:`1px solid ${C.border}`,
                borderRadius:8,width:30,height:30,display:"flex",alignItems:"center",justifyContent:"center",
                cursor:"pointer",color:C.textSub,fontSize:14}}>✕</button>
            </div>

            {!editId&&(
              <div style={{marginBottom:14}}>
                <div style={{fontSize:10,color:C.textSub,fontWeight:600,textTransform:"uppercase",
                  letterSpacing:"0.1em",marginBottom:8}}>⚡ Quick Pick</div>
                <input placeholder="Search foods…" value={presetSearch}
                  onChange={e=>setPresetSearch(e.target.value)}
                  style={{...darkInp,marginBottom:8,fontSize:12,padding:"9px 12px"}}/>
                <div style={{display:"flex",gap:6,flexWrap:"wrap",maxHeight:82,overflowY:"auto"}}>
                  {filteredPresets.map(p=>(
                    <button key={p.id} onClick={()=>applyPreset(p)} style={{
                      padding:"5px 11px",borderRadius:20,border:`1px solid ${C.border}`,
                      background:C.bg2,color:C.textSub,fontSize:11,fontWeight:500,
                      cursor:"pointer",fontFamily:"inherit",
                      display:"flex",alignItems:"center",gap:4,whiteSpace:"nowrap"}}>
                      {p.emoji} {p.name}
                      <span style={{color:C.accent,fontSize:9}}>{calcCalories(p.protein,p.carbs,p.fat)||"0"} kcal</span>
                    </button>
                  ))}
                </div>
                <div style={{height:1,background:C.border,margin:"12px 0"}}/>
              </div>
            )}

            {/* Meal type */}
            <div style={{display:"flex",gap:6,marginBottom:12,flexWrap:"wrap"}}>
              {MEAL_TYPES.map(t=>{
                const active=form.type===t;
                const color=mealColor[t];
                return(
                  <button key={t} onClick={()=>setForm(f=>({...f,type:t}))} style={{
                    padding:"5px 12px",borderRadius:18,border:`1px solid ${active?color:C.border}`,
                    background:active?`${color}22`:"transparent",
                    color:active?color:C.textSub,
                    fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",
                    boxShadow:active?`0 0 10px ${color}44`:"none",transition:"all 0.2s"}}>
                    {mealEmoji[t]} {t}
                  </button>
                );
              })}
            </div>

            {/* Camera */}
            {!editId&&(
              <div style={{marginBottom:12}}>
                {cameraStep==="idle"&&(
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={startCamera} style={{flex:1,padding:"9px 0",borderRadius:11,
                      border:`1px dashed ${C.accent}`,background:C.bg2,color:C.accent,
                      fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",
                      display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>📷 Camera</button>
                    <button onClick={()=>fileInputRef.current?.click()} style={{flex:1,padding:"9px 0",borderRadius:11,
                      border:`1px dashed ${C.borderHi}`,background:C.bg2,color:C.textSub,
                      fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",
                      display:"flex",alignItems:"center",justifyContent:"center",gap:5}}>🖼 Upload</button>
                  </div>
                )}
                {cameraStep==="preview"&&(
                  <div style={{borderRadius:14,overflow:"hidden",background:"#000"}}>
                    <video ref={videoRef} autoPlay playsInline muted style={{width:"100%",maxHeight:180,objectFit:"cover",display:"block"}}/>
                    <div style={{display:"flex",gap:8,padding:"9px",background:"rgba(0,0,0,0.7)"}}>
                      <button onClick={()=>{stopCamera();setCameraStep("idle");}} style={{flex:1,padding:"8px",borderRadius:9,border:"none",background:"rgba(255,255,255,0.1)",color:"white",fontSize:11,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
                      <button onClick={captureFromCamera} style={{flex:2,padding:"8px",borderRadius:9,border:"none",background:C.accentDim,color:"white",fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📸 Capture</button>
                    </div>
                  </div>
                )}
                {(cameraStep==="analyzing"||cameraStep==="done")&&capturedImage&&(
                  <div style={{borderRadius:14,overflow:"hidden"}}>
                    <img src={capturedImage} alt="meal" style={{width:"100%",maxHeight:120,objectFit:"cover",display:"block"}}/>
                    {cameraStep==="analyzing"&&(
                      <div style={{padding:"9px",background:"rgba(0,0,0,0.8)",display:"flex",alignItems:"center",gap:8}}>
                        <span style={{fontSize:16}}>🔍</span>
                        <span style={{color:C.text,fontSize:11}}>Analysing nutrition…</span>
                      </div>
                    )}
                    {cameraStep==="done"&&(
                      <div style={{padding:"7px 10px",background:analysisError?"#2a1010":"#0a2a1a",
                        display:"flex",alignItems:"center",gap:6}}>
                        <span style={{fontSize:12}}>{analysisError?"⚠️":"✅"}</span>
                        <span style={{fontSize:11,color:analysisError?C.red:C.green}}>{analysisError||"Macros pre-filled"}</span>
                        <button onClick={()=>{setCameraStep("idle");setCapturedImage(null);setAnalysisError("");}}
                          style={{marginLeft:"auto",background:"none",border:"none",fontSize:10,color:C.textSub,cursor:"pointer"}}>Retake</button>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            <input placeholder="Meal name" value={form.name} onChange={e=>setForm(f=>({...f,name:e.target.value}))} style={darkInp}/>
            <div style={{fontSize:10,color:C.textMuted,marginBottom:6,fontStyle:"italic"}}>Protein + Carbs + Fat → calories auto-calculated</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
              {[{key:"protein",ph:"Protein g",col:C.red},{key:"carbs",ph:"Carbs g",col:C.green},{key:"fat",ph:"Fat g",col:C.amber}].map(({key,ph,col})=>(
                <input key={key} type="number" min="0" placeholder={ph} value={form[key]}
                  onChange={e=>updateMacro(key,e.target.value)}
                  style={{...darkInpSm,borderColor:form[key]?col:C.border,
                    boxShadow:form[key]?`0 0 8px ${col}44`:"none"}}/>
              ))}
            </div>
            <div style={{background:C.bg2,border:`1px solid ${C.borderHi}`,borderRadius:10,
              padding:"10px 14px",display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,
              boxShadow:`0 0 12px ${C.accentGlow}`}}>
              <span style={{fontSize:11,color:C.accent,fontWeight:700,letterSpacing:"0.06em"}}>🔥 CALORIES</span>
              <input type="number" min="0" value={form.calories} onChange={e=>setForm(f=>({...f,calories:e.target.value}))}
                placeholder="auto" style={{border:"none",background:"transparent",fontSize:18,fontWeight:800,
                color:C.text,width:80,textAlign:"right",outline:"none",fontFamily:"inherit"}}/>
              <span style={{fontSize:11,color:C.textMuted}}>kcal</span>
            </div>
            <input placeholder="Notes (optional)" value={form.notes} onChange={e=>setForm(f=>({...f,notes:e.target.value}))} style={darkInp}/>
            <div style={{display:"flex",gap:9}}>
              <button onClick={()=>setShowForm(false)} style={{flex:1,padding:"12px",borderRadius:12,
                border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,
                fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
              <button onClick={saveMeal} style={{flex:2,padding:"12px",borderRadius:12,border:"none",
                background:form.name.trim()?`linear-gradient(135deg,${C.accent},${C.accentDim})`:"#1e1e2e",
                color:form.name.trim()?C.text:C.textMuted,fontSize:13,fontWeight:700,
                cursor:form.name.trim()?"pointer":"default",fontFamily:"inherit",
                boxShadow:form.name.trim()?`0 0 16px ${C.accentGlow}`:"none"}}>
                {editId?"Save Changes":"Add Meal"}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ══ LIBRARY MODAL ══ */}
      {showLibrary&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)",
          zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}}
          onClick={e=>{if(e.target===e.currentTarget){setShowLibrary(false);cancelPresetEdit();}}}>
          <div style={modalBase}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:C.text}}>📚 My Foods Library</div>
                <div style={{fontSize:11,color:C.textSub}}>{presets.length} items</div>
              </div>
              <button onClick={()=>{setShowLibrary(false);cancelPresetEdit();}} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,width:30,height:30,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",color:C.textSub,fontSize:14}}>✕</button>
            </div>
            <div style={{background:C.bg2,border:`1px solid ${C.borderHi}`,borderRadius:14,
              padding:"14px",marginBottom:14,boxShadow:`0 0 16px ${C.accentGlow}`}}>
              <div style={{fontSize:11,fontWeight:700,color:C.text,marginBottom:10}}>
                {editPresetId?"✏️ Edit Food":"＋ Add New Food"}
              </div>
              <div style={{display:"flex",gap:8,marginBottom:9}}>
                <input value={presetForm.emoji} onChange={e=>setPresetForm(f=>({...f,emoji:e.target.value}))}
                  style={{...darkInpSm,width:46,textAlign:"center",fontSize:18,padding:"7px 4px",flexShrink:0,marginBottom:0}} maxLength={2}/>
                <input placeholder="Food name" value={presetForm.name}
                  onChange={e=>setPresetForm(f=>({...f,name:e.target.value}))}
                  style={{...darkInpSm,flex:1,marginBottom:0}}/>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:8}}>
                {[{k:"protein",ph:"Protein",c:C.red},{k:"carbs",ph:"Carbs",c:C.green},{k:"fat",ph:"Fat",c:C.amber}].map(({k,ph,c})=>(
                  <input key={k} type="number" min="0" placeholder={ph} value={presetForm[k]}
                    onChange={e=>setPresetForm(f=>({...f,[k]:e.target.value}))}
                    style={{...darkInpSm,borderColor:presetForm[k]?c:C.border}}/>
                ))}
              </div>
              <div style={{fontSize:11,color:C.accent,fontWeight:700,textAlign:"right",marginBottom:8}}>
                🔥 {calcCalories(presetForm.protein,presetForm.carbs,presetForm.fat)||"0"} kcal
              </div>
              <input placeholder="Notes" value={presetForm.notes}
                onChange={e=>setPresetForm(f=>({...f,notes:e.target.value}))}
                style={{...darkInpSm,width:"100%",boxSizing:"border-box",marginBottom:10}}/>
              <div style={{display:"flex",gap:8}}>
                {editPresetId&&(
                  <button onClick={cancelPresetEdit} style={{flex:1,padding:"9px",borderRadius:10,
                    border:`1px solid ${C.border}`,background:"transparent",color:C.textSub,
                    fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Cancel</button>
                )}
                <button onClick={savePreset} style={{flex:2,padding:"9px",borderRadius:10,border:"none",
                  background:presetForm.name.trim()?`linear-gradient(135deg,${C.accent},${C.accentDim})`:"#1e1e2e",
                  color:presetForm.name.trim()?C.text:C.textMuted,fontSize:11,fontWeight:700,
                  cursor:presetForm.name.trim()?"pointer":"default",fontFamily:"inherit"}}>
                  {editPresetId?"Save Changes":"Add to Library"}
                </button>
              </div>
            </div>
            <input placeholder="Search foods…" value={libSearch} onChange={e=>setLibSearch(e.target.value)}
              style={{...darkInp,fontSize:12,padding:"9px 12px"}}/>
            {libFiltered.map(p=>(
              <div key={p.id} style={{background:editPresetId===p.id?C.bg3:C.bg2,
                border:`1px solid ${editPresetId===p.id?C.accent:C.border}`,
                borderRadius:12,padding:"10px 12px",marginBottom:7,
                display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:22}}>{p.emoji}</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontWeight:600,color:C.text,fontSize:12}}>{p.name}</div>
                  <div style={{fontSize:10,color:C.textSub,display:"flex",gap:8,marginTop:2}}>
                    <span style={{color:C.accent}}>{calcCalories(p.protein,p.carbs,p.fat)||"0"} kcal</span>
                    {p.protein?<span style={{color:C.red}}>{p.protein}g P</span>:""}
                    {p.carbs?<span style={{color:C.green}}>{p.carbs}g C</span>:""}
                    {p.fat?<span style={{color:C.amber}}>{p.fat}g F</span>:""}
                  </div>
                </div>
                <button onClick={()=>editPreset(p)} style={{background:C.bg3,border:`1px solid ${C.border}`,borderRadius:7,cursor:"pointer",fontSize:11,padding:"4px 8px",color:C.textSub}}>✏️</button>
                <button onClick={()=>deletePreset(p.id)} style={{background:C.bg3,border:`1px solid ${C.border}`,borderRadius:7,cursor:"pointer",fontSize:11,padding:"4px 8px",color:C.textSub}}>🗑</button>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* ══ GOALS MODAL ══ */}
      {showGoals&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)",
          zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}}
          onClick={e=>{if(e.target===e.currentTarget)setShowGoals(false);}}>
          <div style={{background:C.bg1,borderRadius:"24px 24px 0 0",padding:"22px 18px 36px",
            width:"100%",maxWidth:480,boxShadow:"0 -12px 60px rgba(0,0,0,0.8)",
            maxHeight:"92vh",overflowY:"auto",border:`1px solid ${C.border}`,borderBottom:"none"}}>

            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:18}}>
              <div>
                <div style={{fontSize:16,fontWeight:700,color:C.text}}>🎯 Your Goal</div>
                <div style={{fontSize:11,color:C.textSub}}>Drives your adaptive calorie target</div>
              </div>
              <button onClick={()=>setShowGoals(false)} style={{background:C.bg2,border:`1px solid ${C.border}`,
                borderRadius:8,width:30,height:30,display:"flex",alignItems:"center",justifyContent:"center",
                cursor:"pointer",color:C.textSub,fontSize:14}}>✕</button>
            </div>

            {/* Goal selector */}
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:10,marginBottom:20}}>
              {[
                {val:"lose",   icon:"🔥",label:"Lose Weight", desc:"−500 kcal/day", color:"#f87171"},
                {val:"maintain",icon:"⚖️",label:"Maintain",   desc:"At your TDEE",  color:"#a78bfa"},
                {val:"gain",   icon:"💪",label:"Build Muscle",desc:"+300 kcal/day", color:"#34d399"},
              ].map(g=>(
                <div key={g.val} onClick={()=>setUserGoal(g.val)}
                  style={{background:userGoal===g.val?`${g.color}22`:C.bg2,
                    border:`1.5px solid ${userGoal===g.val?g.color:C.border}`,
                    borderRadius:14,padding:"14px 10px",textAlign:"center",cursor:"pointer",
                    boxShadow:userGoal===g.val?`0 0 16px ${g.color}33`:"none",
                    transition:"all 0.2s"}}>
                  <div style={{fontSize:24,marginBottom:6}}>{g.icon}</div>
                  <div style={{fontSize:12,fontWeight:700,color:userGoal===g.val?g.color:C.text,marginBottom:3}}>{g.label}</div>
                  <div style={{fontSize:10,color:C.textSub}}>{g.desc}</div>
                </div>
              ))}
            </div>

            {/* Algorithm result */}
            {adaptive ? (
              <div style={{background:C.bg2,border:`1px solid ${C.green}44`,borderRadius:14,padding:"16px",marginBottom:16}}>
                <div style={{fontSize:11,color:C.green,fontWeight:700,letterSpacing:"0.08em",textTransform:"uppercase",marginBottom:12}}>
                  🧬 Algorithm Result
                </div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:12}}>
                  {[
                    {label:"Your Real TDEE",   val:adaptive.tdee,         unit:"kcal", color:C.cyan},
                    {label:"Recommended Goal", val:adaptive.recommendedGoal,unit:"kcal",color:C.accent},
                    {label:"Avg Calories Logged",val:adaptive.avgCals,   unit:"kcal", color:C.amber},
                    {label:"Weight Change",    val:`${adaptive.kgChange>0?"+":""}${adaptive.kgChange}`,unit:"kg",color:adaptive.kgChange<0?C.green:adaptive.kgChange>0?C.red:C.textSub},
                  ].map(s=>(
                    <div key={s.label} style={{background:C.bg1,borderRadius:10,padding:"10px 12px",border:`1px solid ${C.border}`}}>
                      <div style={{fontSize:18,fontWeight:800,color:s.color}}>{s.val}<span style={{fontSize:11,fontWeight:500}}> {s.unit}</span></div>
                      <div style={{fontSize:10,color:C.textSub,marginTop:2}}>{s.label}</div>
                    </div>
                  ))}
                </div>
                <div style={{fontSize:11,color:C.textSub,lineHeight:1.5}}>
                  Based on <span style={{color:C.text,fontWeight:600}}>{adaptive.daysAnalysed} days</span> of meal logs
                  {" "}and your weight data. Confidence:{" "}
                  <span style={{color:adaptive.confidence==="high"?C.green:adaptive.confidence==="medium"?C.amber:C.red,fontWeight:600}}>
                    {adaptive.confidence}
                  </span>
                  {adaptive.confidence!=="high"&&" — log more days for higher accuracy."}
                </div>
              </div>
            ) : (
              <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:14,padding:"16px",marginBottom:16}}>
                <div style={{fontSize:12,fontWeight:700,color:C.amber,marginBottom:8}}>⏳ Not enough data yet</div>
                <div style={{fontSize:11,color:C.textSub,lineHeight:1.6}}>
                  The algorithm needs:<br/>
                  <span style={{color:C.text}}>✓ At least 7 days of meal logs</span><br/>
                  <span style={{color:C.text}}>✓ At least 2 weight entries</span><br/><br/>
                  Until then, your calorie goal is set to <span style={{color:C.accent,fontWeight:600}}>{CAL_GOAL} kcal</span>.
                  Keep logging and it will activate automatically.
                </div>
              </div>
            )}

            <button onClick={()=>setShowGoals(false)} style={{width:"100%",padding:"13px",borderRadius:12,border:"none",
              background:`linear-gradient(135deg,${C.accent},${C.accentDim})`,
              color:C.text,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit",
              boxShadow:`0 0 16px ${C.accentGlow}`}}>
              Got it
            </button>
          </div>
        </div>
      )}

      {/* ══ REPORT MODAL ══ */}
      {showReport&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.75)",backdropFilter:"blur(8px)",
          zIndex:300,display:"flex",alignItems:"flex-end",justifyContent:"center"}}
          onClick={e=>{if(e.target===e.currentTarget)setShowReport(false);}}>
          <div style={modalBase}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
              <div>
                <div style={{fontSize:15,fontWeight:700,color:C.text}}>📊 Download Report</div>
                <div style={{fontSize:11,color:C.textSub}}>Export your meal history</div>
              </div>
              <button onClick={()=>setShowReport(false)} style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:8,width:30,height:30,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",color:C.textSub,fontSize:14}}>✕</button>
            </div>
            <div style={{marginBottom:14}}>
              <div style={{fontSize:10,color:C.textSub,fontWeight:600,textTransform:"uppercase",letterSpacing:"0.08em",marginBottom:8}}>Date Range</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap"}}>
                {[["7","7 days"],["14","14 days"],["30","30 days"],["custom","Custom"]].map(([val,label])=>(
                  <button key={val} onClick={()=>setReportRange(val)} style={{
                    padding:"6px 13px",borderRadius:18,border:`1px solid ${reportRange===val?C.accent:C.border}`,
                    background:reportRange===val?C.accentDim:"transparent",
                    color:reportRange===val?C.text:C.textSub,
                    fontSize:11,fontWeight:600,cursor:"pointer",fontFamily:"inherit",
                    boxShadow:reportRange===val?`0 0 10px ${C.accentGlow}`:"none"}}>{label}</button>
                ))}
              </div>
            </div>
            {reportRange==="custom"&&(
              <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center"}}>
                <div style={{flex:1}}>
                  <div style={{fontSize:10,color:C.textSub,marginBottom:5}}>From</div>
                  <input type="date" value={reportFrom} onChange={e=>setReportFrom(e.target.value)} style={{...darkInpSm,width:"100%",boxSizing:"border-box"}}/>
                </div>
                <div style={{color:C.textMuted,marginTop:14}}>→</div>
                <div style={{flex:1}}>
                  <div style={{fontSize:10,color:C.textSub,marginBottom:5}}>To</div>
                  <input type="date" value={reportTo} onChange={e=>setReportTo(e.target.value)} style={{...darkInpSm,width:"100%",boxSizing:"border-box"}}/>
                </div>
              </div>
            )}
            <div style={{background:C.bg2,border:`1px solid ${C.border}`,borderRadius:14,
              padding:"14px",marginBottom:16}}>
              <div style={{fontSize:10,color:C.textSub,fontWeight:600,textTransform:"uppercase",
                letterSpacing:"0.08em",marginBottom:12}}>Period Summary</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                {[
                  {label:"Days logged",val:reportDaysWithData.length,unit:"days",color:C.accent},
                  {label:"Avg calories",val:avgCalories,unit:"kcal",color:C.red},
                  {label:"Total protein",val:reportTotals.protein,unit:"g",color:C.green},
                  {label:"Total meals",val:reportDates.reduce((a,d)=>a+(allData[d]||[]).length,0),unit:"",color:C.amber},
                ].map(s=>(
                  <div key={s.label} style={{background:C.bg1,borderRadius:10,padding:"10px 12px",border:`1px solid ${C.border}`}}>
                    <div style={{fontSize:20,fontWeight:800,color:s.color}}>{s.val}<span style={{fontSize:11,fontWeight:500}}> {s.unit}</span></div>
                    <div style={{fontSize:10,color:C.textSub,marginTop:2}}>{s.label}</div>
                  </div>
                ))}
              </div>
            </div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>exportExcel(allData,reportDates)} disabled={reportDaysWithData.length===0}
                style={{flex:1,padding:"14px 8px",borderRadius:13,border:"none",
                  background:reportDaysWithData.length>0?"linear-gradient(135deg,#064e3b,#065f46)":"#1e1e2e",
                  color:reportDaysWithData.length>0?C.text:C.textMuted,fontSize:12,fontWeight:700,
                  cursor:reportDaysWithData.length>0?"pointer":"default",fontFamily:"inherit",
                  display:"flex",flexDirection:"column",alignItems:"center",gap:5}}>
                <span style={{fontSize:24}}>📗</span>
                <span>Excel</span>
                <span style={{fontSize:10,opacity:0.7}}>.xlsx file</span>
              </button>
              <button onClick={()=>exportPDF(allData,reportDates)} disabled={reportDaysWithData.length===0}
                style={{flex:1,padding:"14px 8px",borderRadius:13,border:"none",
                  background:reportDaysWithData.length>0?"linear-gradient(135deg,#7f1d1d,#991b1b)":"#1e1e2e",
                  color:reportDaysWithData.length>0?C.text:C.textMuted,fontSize:12,fontWeight:700,
                  cursor:reportDaysWithData.length>0?"pointer":"default",fontFamily:"inherit",
                  display:"flex",flexDirection:"column",alignItems:"center",gap:5}}>
                <span style={{fontSize:24}}>📕</span>
                <span>PDF / Print</span>
                <span style={{fontSize:10,opacity:0.7}}>Printable</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
