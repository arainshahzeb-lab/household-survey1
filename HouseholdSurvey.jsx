import { useState } from "react";
import * as XLSX from "xlsx";

const C = {
  bg: "#060D1A", card: "#0D1B2E", card2: "#0F2035",
  border: "#1A3050", accent: "#3B82F6", accentDim: "#1E3A5F",
  green: "#10B981", red: "#EF4444", muted: "#4B6A8A",
  text: "#E2EAF4", textDim: "#8BAAC8",
};

const EDU_OPTIONS = [
  "None","No Education","Primary (1-5)","Middle (6-8)","Matric (9-10)",
  "Intermediate (11-12)","Graduate","Post-Graduate","Hafiz-e-Quran","Other",
];
const DISABILITY_TYPES = [
  "Physical","Visual","Hearing","Speech","Intellectual",
  "Autism","Multiple Disabilities","Other",
];
const OCCUPATION_OPTIONS = [
  "Farmer","Labor / Daily Wage","Labour","Government Job","Private Job",
  "Business / Self-Employed","Student","Housewife","Unemployed","Other",
];
const TRANSPORT_OPTIONS = [
  "None","Bicycle","Motorcycle","Car / Van",
  "Tractor","Public Transport","Animal Cart","Other",
];

const MAX_CHILDREN = 15;
const SECTIONS = [
  "Household Head","Spouse","Children Count",
  "Ages","Education","Disability","Other Details","Review",
];

const emptyForm = () => ({
  hhDeceased: false,
  hhName: "", hhCnic: "", hhContact: "",
  spouseName: "", spouseCnic: "", spouseContact: "",
  maleChildCount: 0, femaleChildCount: 0,
  hhAge: "", spouseAge: "",
  maleChildAges: [], femaleChildAges: [],
  hhEducation: "None", spouseEducation: "None",
  maleChildEducations: [], femaleChildEducations: [],
  disabledCount: 0, disabilityTypes: [],
  occupation: "", earners: "", hpNo: "", toiletNo: "",
  transport: "", income: "", omCapacity: "",
});

// Format CNIC: 12345-1234567-1
const formatCnic = (raw) => {
  const digits = raw.replace(/\D/g, "").slice(0, 13);
  if (digits.length <= 5) return digits;
  if (digits.length <= 12) return `${digits.slice(0,5)}-${digits.slice(5)}`;
  return `${digits.slice(0,5)}-${digits.slice(5,12)}-${digits.slice(12)}`;
};

// Format contact: 1234-1234567
const formatContact = (raw) => {
  const digits = raw.replace(/\D/g, "").slice(0, 11);
  if (digits.length <= 4) return digits;
  return `${digits.slice(0,4)}-${digits.slice(4)}`;
};

const S = {
  root: { minHeight:"100vh", background:C.bg, fontFamily:"'Inter','Segoe UI',sans-serif", color:C.text, display:"flex", flexDirection:"column", alignItems:"center" },
  header: { width:"100%", background:C.card, borderBottom:`1px solid ${C.border}`, padding:"18px 20px 14px", boxSizing:"border-box" },
  htag: { fontSize:9, letterSpacing:2, color:C.accent, textTransform:"uppercase", marginBottom:4 },
  htitle: { margin:0, fontSize:22, fontWeight:800, letterSpacing:-0.5, fontFamily:"'Georgia',serif", color:C.text },
  hsub: { fontSize:11, color:C.muted, marginTop:3 },
  progress: { width:"100%", height:3, background:C.accentDim, marginTop:10, borderRadius:2 },
  progressFill: (pct) => ({ height:"100%", width:`${pct}%`, background:C.accent, borderRadius:2, transition:"width .4s ease" }),
  wrap: { width:"100%", maxWidth:480, padding:"0 14px 100px", boxSizing:"border-box" },
  sectionLabel: { fontSize:10, letterSpacing:2, color:C.accent, textTransform:"uppercase", marginTop:20, marginBottom:10, display:"flex", alignItems:"center", gap:8 },
  sectionLine: { flex:1, height:1, background:C.border },
  card: { background:C.card, border:`1px solid ${C.border}`, borderRadius:14, padding:"16px", marginBottom:10 },
  label: { fontSize:10, color:C.muted, letterSpacing:1, textTransform:"uppercase", marginBottom:5, display:"block" },
  input: { width:"100%", background:"#060D1A", border:`1px solid ${C.border}`, borderRadius:9, padding:"10px 12px", color:C.text, fontSize:14, fontFamily:"inherit", boxSizing:"border-box", outline:"none" },
  select: { width:"100%", background:"#060D1A", border:`1px solid ${C.border}`, borderRadius:9, padding:"10px 12px", color:C.text, fontSize:14, fontFamily:"inherit", boxSizing:"border-box", outline:"none", appearance:"none" },
  row2: { display:"grid", gridTemplateColumns:"1fr 1fr", gap:10 },
  row3: { display:"grid", gridTemplateColumns:"1fr 1fr 1fr", gap:10 },
  counter: { display:"flex", alignItems:"center", gap:12, background:"#060D1A", border:`1px solid ${C.border}`, borderRadius:9, padding:"8px 12px" },
  cBtn: (color) => ({ width:32, height:32, borderRadius:8, border:"none", background:color, color:"#fff", fontSize:20, lineHeight:1, cursor:"pointer", fontWeight:700, display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0 }),
  cVal: { flex:1, textAlign:"center", fontSize:20, fontWeight:700, color:C.text },
  navRow: { display:"flex", gap:10, marginTop:16 },
  prevBtn: { flex:1, padding:"13px", background:"transparent", border:`1.5px solid ${C.border}`, borderRadius:11, color:C.textDim, fontSize:14, fontWeight:600, cursor:"pointer", fontFamily:"inherit" },
  nextBtn: { flex:2, padding:"13px", background:C.accent, border:"none", borderRadius:11, color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit" },
  addBtn: { width:"100%", padding:"13px", background:C.green, border:"none", borderRadius:11, color:"#fff", fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit" },
  exportBtn: { width:"100%", padding:"13px", background:"transparent", border:`2px solid ${C.accent}`, borderRadius:11, color:C.accent, fontSize:14, fontWeight:700, cursor:"pointer", fontFamily:"inherit", marginTop:8 },
  recordCard: { background:C.card, border:`1px solid ${C.border}`, borderRadius:12, padding:"13px 14px", marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center" },
  chip: { background:C.accentDim, color:C.accent, fontSize:10, padding:"3px 8px", borderRadius:6, fontWeight:700 },
  delBtn: { background:"none", border:"none", color:C.red, cursor:"pointer", fontSize:16, padding:"2px 6px" },
  toast: (show) => ({ position:"fixed", bottom:24, left:"50%", transform:`translateX(-50%) translateY(${show?0:80}px)`, background:C.green, color:"#fff", padding:"10px 22px", borderRadius:50, fontSize:13, fontWeight:700, boxShadow:`0 4px 20px ${C.green}55`, transition:"transform .3s cubic-bezier(.34,1.56,.64,1)", pointerEvents:"none", zIndex:999, whiteSpace:"nowrap" }),
  tabs: { display:"flex", gap:4, margin:"16px 0 14px", background:C.card, borderRadius:11, padding:4, border:`1px solid ${C.border}` },
  tab: (a) => ({ flex:1, padding:"8px 0", borderRadius:7, border:"none", cursor:"pointer", fontSize:12, fontWeight:600, background:a?C.accent:"transparent", color:a?"#fff":C.muted, fontFamily:"inherit", transition:"all .2s" }),
  reviewRow: { display:"flex", justifyContent:"space-between", padding:"6px 0", borderBottom:`1px solid ${C.border}`, fontSize:12 },
  reviewKey: { color:C.muted },
  reviewVal: { color:C.text, fontWeight:600, textAlign:"right", maxWidth:"60%" },
  // Deceased toggle
  deceasedBox: (on) => ({ display:"flex", alignItems:"center", gap:10, padding:"12px 14px", borderRadius:10, border:`1.5px solid ${on?"#EF4444":C.border}`, background: on?"#2A0A0A":"#060D1A", cursor:"pointer", marginBottom:14 }),
  deceasedDot: (on) => ({ width:22, height:22, borderRadius:6, border:`2px solid ${on?"#EF4444":C.border}`, background:on?"#EF4444":"transparent", display:"flex", alignItems:"center", justifyContent:"center", flexShrink:0, transition:"all .2s" }),
};

function SectionHead({ title }) {
  return (
    <div style={S.sectionLabel}>
      <span>{title}</span><span style={S.sectionLine}/>
    </div>
  );
}
function Field({ label, children, style }) {
  return <div style={{ marginBottom:12, ...style }}><label style={S.label}>{label}</label>{children}</div>;
}
function Counter({ label, value, onChange, max=MAX_CHILDREN }) {
  return (
    <div>
      <label style={S.label}>{label}</label>
      <div style={S.counter}>
        <button style={S.cBtn("#1A3050")} onClick={()=>onChange(Math.max(0,value-1))}>−</button>
        <span style={S.cVal}>{value}</span>
        <button style={S.cBtn(C.accent)} onClick={()=>onChange(Math.min(max,value+1))}>+</button>
      </div>
    </div>
  );
}

export default function App() {
  const [mainTab, setMainTab] = useState("form");
  const [step, setStep] = useState(0);
  const [form, setForm] = useState(emptyForm());
  const [records, setRecords] = useState([]);
  const [toast, setToast] = useState(false);
  const [toastMsg, setToastMsg] = useState("");

  const showToast = (msg) => { setToastMsg(msg); setToast(true); setTimeout(()=>setToast(false),2400); };
  const set = (key, val) => setForm(f=>({...f,[key]:val}));

  const setMaleCount = (n) => setForm(f=>({
    ...f, maleChildCount:n,
    maleChildAges: Array.from({length:n},(_,i)=>f.maleChildAges[i]??""),
    maleChildEducations: Array.from({length:n},(_,i)=>f.maleChildEducations[i]??"None"),
  }));
  const setFemaleCount = (n) => setForm(f=>({
    ...f, femaleChildCount:n,
    femaleChildAges: Array.from({length:n},(_,i)=>f.femaleChildAges[i]??""),
    femaleChildEducations: Array.from({length:n},(_,i)=>f.femaleChildEducations[i]??"None"),
  }));

  const toggleDeceased = () => setForm(f=>({...f, hhDeceased:!f.hhDeceased, hhAge:"", spouseAge:""}));

  const totalMale = 1 + form.maleChildCount;
  const totalFemale = 1 + form.femaleChildCount;
  const progress = ((step+1)/SECTIONS.length)*100;

  // ── label adjustments based on deceased ──
  const hhLabel = form.hhDeceased ? "Spouse Name (acting HH)" : "HH Name";
  const hhCnicLabel = form.hhDeceased ? "Spouse CNIC (acting HH)" : "HH CNIC";
  const hhContactLabel = form.hhDeceased ? "Spouse Contact (acting HH)" : "HH Contact";

  const steps = [
    // 0 — HH
    <div key="hh">
      <SectionHead title="Household Head"/>
      <div style={S.card}>
        {/* Deceased toggle */}
        <div style={S.deceasedBox(form.hhDeceased)} onClick={toggleDeceased}>
          <div style={S.deceasedDot(form.hhDeceased)}>
            {form.hhDeceased && <span style={{color:"#fff",fontSize:14,fontWeight:900}}>✓</span>}
          </div>
          <div>
            <div style={{fontSize:13,fontWeight:700,color:form.hhDeceased?"#EF4444":C.text}}>HH is Deceased</div>
            <div style={{fontSize:10,color:C.muted,marginTop:2}}>Spouse becomes acting head; ages adjusted automatically</div>
          </div>
        </div>

        <Field label={hhLabel}>
          <input style={S.input} value={form.hhName} placeholder="Full name" onChange={e=>set("hhName",e.target.value)}/>
        </Field>
        <Field label={hhCnicLabel}>
          <input style={S.input} value={form.hhCnic} placeholder="12345-1234567-1"
            onChange={e=>set("hhCnic", formatCnic(e.target.value))}
            maxLength={15}/>
        </Field>
        <Field label={hhContactLabel} style={{marginBottom:0}}>
          <input style={S.input} value={form.hhContact} placeholder="1234-1234567"
            onChange={e=>set("hhContact", formatContact(e.target.value))}
            maxLength={12}/>
        </Field>
      </div>
    </div>,

    // 1 — Spouse (hidden if deceased since spouse is now HH)
    <div key="sp">
      <SectionHead title={form.hhDeceased ? "Deceased HH Info" : "Spouse"}/>
      <div style={S.card}>
        {form.hhDeceased
          ? <div style={{fontSize:12,color:C.muted,padding:"8px 0",lineHeight:1.7}}>
              Since the HH is deceased, the spouse has been recorded as the acting household head in the previous step.<br/><br/>
              You can record the deceased HH's info below (optional).
            </div>
          : null
        }
        <Field label={form.hhDeceased ? "Deceased HH Name (optional)" : "Spouse Name"}>
          <input style={S.input} value={form.spouseName} placeholder="Full name" onChange={e=>set("spouseName",e.target.value)}/>
        </Field>
        <Field label={form.hhDeceased ? "Deceased HH CNIC (optional)" : "Spouse CNIC"}>
          <input style={S.input} value={form.spouseCnic} placeholder="12345-1234567-1"
            onChange={e=>set("spouseCnic", formatCnic(e.target.value))} maxLength={15}/>
        </Field>
        <Field label={form.hhDeceased ? "Deceased HH Contact (optional)" : "Spouse Contact"} style={{marginBottom:0}}>
          <input style={S.input} value={form.spouseContact} placeholder="1234-1234567"
            onChange={e=>set("spouseContact", formatContact(e.target.value))} maxLength={12}/>
        </Field>
      </div>
    </div>,

    // 2 — Children Count
    <div key="count">
      <SectionHead title="Children Count"/>
      <div style={S.card}>
        <div style={S.row2}>
          <Counter label="Male Children" value={form.maleChildCount} onChange={setMaleCount}/>
          <Counter label="Female Children" value={form.femaleChildCount} onChange={setFemaleCount}/>
        </div>
        <div style={{...S.row2, marginTop:16}}>
          <div style={{background:"#060D1A",borderRadius:9,padding:"10px 12px",textAlign:"center"}}>
            <div style={{fontSize:22,fontWeight:800,color:C.accent}}>{totalMale}</div>
            <div style={{fontSize:10,color:C.muted,textTransform:"uppercase",letterSpacing:1}}>Total Male</div>
            <div style={{fontSize:10,color:C.muted}}>(HH + {form.maleChildCount} children)</div>
          </div>
          <div style={{background:"#060D1A",borderRadius:9,padding:"10px 12px",textAlign:"center"}}>
            <div style={{fontSize:22,fontWeight:800,color:"#A78BFA"}}>{totalFemale}</div>
            <div style={{fontSize:10,color:C.muted,textTransform:"uppercase",letterSpacing:1}}>Total Female</div>
            <div style={{fontSize:10,color:C.muted}}>(Spouse + {form.femaleChildCount} children)</div>
          </div>
        </div>
      </div>
    </div>,

    // 3 — Ages
    <div key="ages">
      <SectionHead title="Ages"/>
      <div style={S.card}>
        <div style={S.row2}>
          {/* If deceased: HH age label = "Acting HH (Spouse) Age", recorded in female age col */}
          <div>
            <label style={S.label}>{form.hhDeceased ? "Acting HH Age (Spouse)" : "HH Age"}</label>
            <input style={S.input} type="number" placeholder="e.g. 40" value={form.hhAge} onChange={e=>set("hhAge",e.target.value)}/>
          </div>
          <div>
            <label style={S.label}>{form.hhDeceased ? "Deceased HH Age (optional)" : "Spouse Age"}</label>
            <input style={S.input} type="number" placeholder="e.g. 35" value={form.spouseAge} onChange={e=>set("spouseAge",e.target.value)}/>
          </div>
        </div>
        {form.hhDeceased && (
          <div style={{fontSize:10,color:"#EF4444",marginTop:4,marginBottom:8,padding:"6px 10px",background:"#2A0A0A",borderRadius:7}}>
            ⚠ Acting HH age will be recorded in the Female Age column as per survey rules.
          </div>
        )}

        {form.maleChildCount > 0 && (
          <>
            <div style={{fontSize:10,color:C.accent,letterSpacing:1,textTransform:"uppercase",margin:"10px 0 8px"}}>Male Children Ages</div>
            <div style={S.row3}>
              {form.maleChildAges.map((age,i)=>(
                <div key={i}>
                  <label style={S.label}>Boy {i+1}</label>
                  <input style={S.input} type="number" placeholder="age" value={age}
                    onChange={e=>{const a=[...form.maleChildAges];a[i]=e.target.value;set("maleChildAges",a);}}/>
                </div>
              ))}
            </div>
          </>
        )}

        {form.femaleChildCount > 0 && (
          <>
            <div style={{fontSize:10,color:"#A78BFA",letterSpacing:1,textTransform:"uppercase",margin:"10px 0 8px"}}>Female Children Ages</div>
            <div style={S.row3}>
              {form.femaleChildAges.map((age,i)=>(
                <div key={i}>
                  <label style={S.label}>Girl {i+1}</label>
                  <input style={S.input} type="number" placeholder="age" value={age}
                    onChange={e=>{const a=[...form.femaleChildAges];a[i]=e.target.value;set("femaleChildAges",a);}}/>
                </div>
              ))}
            </div>
          </>
        )}
      </div>
    </div>,

    // 4 — Education
    <div key="edu">
      <SectionHead title="Education"/>
      <div style={S.card}>
        <div style={S.row2}>
          <div>
            <label style={S.label}>{form.hhDeceased?"Acting HH Education":"HH Education"}</label>
            <select style={S.select} value={form.hhEducation} onChange={e=>set("hhEducation",e.target.value)}>
              <option value="">Select…</option>
              {EDU_OPTIONS.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
          <div>
            <label style={S.label}>{form.hhDeceased?"Deceased HH Edu":"Spouse Education"}</label>
            <select style={S.select} value={form.spouseEducation} onChange={e=>set("spouseEducation",e.target.value)}>
              <option value="">Select…</option>
              {EDU_OPTIONS.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
        </div>
        {form.maleChildCount>0&&(
          <>
            <div style={{fontSize:10,color:C.accent,letterSpacing:1,textTransform:"uppercase",margin:"10px 0 8px"}}>Male Children Education</div>
            {form.maleChildEducations.map((edu,i)=>(
              <div key={i} style={{marginBottom:10}}>
                <label style={S.label}>Boy {i+1}</label>
                <select style={S.select} value={edu} onChange={e=>{const a=[...form.maleChildEducations];a[i]=e.target.value;set("maleChildEducations",a);}}>
                  <option value="">Select…</option>
                  {EDU_OPTIONS.map(o=><option key={o}>{o}</option>)}
                </select>
              </div>
            ))}
          </>
        )}
        {form.femaleChildCount>0&&(
          <>
            <div style={{fontSize:10,color:"#A78BFA",letterSpacing:1,textTransform:"uppercase",margin:"10px 0 8px"}}>Female Children Education</div>
            {form.femaleChildEducations.map((edu,i)=>(
              <div key={i} style={{marginBottom:10}}>
                <label style={S.label}>Girl {i+1}</label>
                <select style={S.select} value={edu} onChange={e=>{const a=[...form.femaleChildEducations];a[i]=e.target.value;set("femaleChildEducations",a);}}>
                  <option value="">Select…</option>
                  {EDU_OPTIONS.map(o=><option key={o}>{o}</option>)}
                </select>
              </div>
            ))}
          </>
        )}
      </div>
    </div>,

    // 5 — Disability
    <div key="dis">
      <SectionHead title="Disability"/>
      <div style={S.card}>
        <Counter label="No. of Disabled People" value={form.disabledCount}
          onChange={n=>{set("disabledCount",n);set("disabilityTypes",Array.from({length:n},(_,i)=>form.disabilityTypes[i]??""));}}
          max={totalMale+totalFemale}/>
        {form.disabledCount>0&&(
          <div style={{marginTop:14}}>
            {form.disabilityTypes.map((dt,i)=>(
              <div key={i} style={{marginBottom:10}}>
                <label style={S.label}>Person {i+1} — Disability Type</label>
                <select style={S.select} value={dt} onChange={e=>{const a=[...form.disabilityTypes];a[i]=e.target.value;set("disabilityTypes",a);}}>
                  <option value="">Select…</option>
                  {DISABILITY_TYPES.map(o=><option key={o}>{o}</option>)}
                </select>
              </div>
            ))}
          </div>
        )}
      </div>
    </div>,

    // 6 — Other Details (OM Capacity next to Income, HP No as text)
    <div key="other">
      <SectionHead title="Other Details"/>
      <div style={S.card}>
        <div style={S.row2}>
          <div>
            <label style={S.label}>Occupation</label>
            <select style={S.select} value={form.occupation} onChange={e=>set("occupation",e.target.value)}>
              <option value="">Select…</option>
              {OCCUPATION_OPTIONS.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
          <div>
            <label style={S.label}>Means of Transport</label>
            <select style={S.select} value={form.transport} onChange={e=>set("transport",e.target.value)}>
              <option value="">Select…</option>
              {TRANSPORT_OPTIONS.map(o=><option key={o}>{o}</option>)}
            </select>
          </div>
        </div>
        <div style={{...S.row2, marginTop:2}}>
          <Field label="No. of Earners">
            <input style={S.input} type="number" placeholder="e.g. 2" value={form.earners} onChange={e=>set("earners",e.target.value)}/>
          </Field>
          <Field label="Toilet No.">
            <input style={S.input} type="number" placeholder="0" value={form.toiletNo} onChange={e=>set("toiletNo",e.target.value)}/>
          </Field>
        </div>
        {/* HP No as text field */}
        <Field label="HP No. (text or number)">
          <input style={S.input} type="text" placeholder="e.g. HP-001 or 5" value={form.hpNo} onChange={e=>set("hpNo",e.target.value)}/>
        </Field>
        {/* Income and OM Capacity side by side */}
        <div style={S.row2}>
          <Field label="Income (PKR)" style={{marginBottom:0}}>
            <input style={S.input} type="number" placeholder="e.g. 25000" value={form.income} onChange={e=>set("income",e.target.value)}/>
          </Field>
          <Field label="OM Capacity" style={{marginBottom:0}}>
            <input style={S.input} type="number" placeholder="0" value={form.omCapacity} onChange={e=>set("omCapacity",e.target.value)}/>
          </Field>
        </div>
      </div>
    </div>,

    // 7 — Review
    <div key="review">
      <SectionHead title="Review & Save"/>
      <div style={S.card}>
        {[
          ["HH Deceased", form.hhDeceased?"Yes":"No"],
          [form.hhDeceased?"Acting HH (Spouse) Name":"HH Name", form.hhName],
          ["HH CNIC", form.hhCnic],
          ["HH Contact", form.hhContact],
          [form.hhDeceased?"Deceased HH Name":"Spouse Name", form.spouseName],
          ["Spouse CNIC", form.spouseCnic],
          ["Spouse Contact", form.spouseContact],
          ["Total Male", totalMale],
          ["Total Female", totalFemale],
          ["Male Children", form.maleChildCount],
          ["Female Children", form.femaleChildCount],
          [form.hhDeceased?"Acting HH Age (Female col)":"HH Age", form.hhAge],
          [form.hhDeceased?"Deceased HH Age":"Spouse Age", form.spouseAge],
          ...(form.maleChildAges.map((a,i)=>[`Boy ${i+1} Age`,a])),
          ...(form.femaleChildAges.map((a,i)=>[`Girl ${i+1} Age`,a])),
          ["HH Education", form.hhEducation],
          ["Spouse Education", form.spouseEducation],
          ...(form.maleChildEducations.map((e,i)=>[`Boy ${i+1} Edu`,e])),
          ...(form.femaleChildEducations.map((e,i)=>[`Girl ${i+1} Edu`,e])),
          ["Disabled People", form.disabledCount],
          ...(form.disabilityTypes.map((t,i)=>[`Disability ${i+1}`,t])),
          ["Occupation", form.occupation],
          ["No. of Earners", form.earners],
          ["HP No.", form.hpNo],
          ["Toilet No.", form.toiletNo],
          ["Transport", form.transport],
          ["Income (PKR)", form.income],
          ["OM Capacity", form.omCapacity],
        ].map(([k,v])=>(
          <div key={k} style={S.reviewRow}>
            <span style={S.reviewKey}>{k}</span>
            <span style={S.reviewVal}>{v||"—"}</span>
          </div>
        ))}
      </div>
    </div>,
  ];

  const handleSave = () => {
    if (!form.hhName.trim()) { showToast("⚠ HH Name is required"); return; }
    setRecords(r=>[{...form,id:Date.now(),totalMale,totalFemale},...r]);
    setForm(emptyForm()); setStep(0);
    showToast("✓ Household saved!"); setMainTab("records");
  };

  const handleExport = () => {
    if (!records.length) { showToast("No records to export"); return; }
    const rows = buildRows();

    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"]=Object.keys(rows[0]).map(()=>({wch:20}));
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,ws,"Household Survey");
    const wbOut = XLSX.write(wb, { bookType:"xlsx", type:"array" });
    const blob = new Blob([wbOut], { type:"application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `household_survey_${new Date().toISOString().slice(0,10)}.xlsx`;
    document.body.appendChild(a); a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); document.body.removeChild(a); }, 1000);
    showToast("📥 Excel downloaded!");
  };

  const buildRows = () => {
    const maxM = Math.max(...records.map(r=>r.maleChildCount),0);
    const maxF = Math.max(...records.map(r=>r.femaleChildCount),0);
    const maxD = Math.max(...records.map(r=>r.disabledCount),0);
    return records.map(r=>{
      const row = {
        "HH Deceased": r.hhDeceased?"Yes":"No",
        "HH Name": r.hhName,
        "HH CNIC": r.hhCnic,
        "HH Contact": r.hhContact,
        [r.hhDeceased?"Deceased HH Name":"Spouse Name"]: r.spouseName,
        "Spouse CNIC": r.spouseCnic,
        "Spouse Contact": r.spouseContact,
        "Total Male": r.totalMale,
        "Total Female": r.totalFemale,
        "No. of Male Children": r.maleChildCount,
        "No. of Female Children": r.femaleChildCount,
        "HH Age (Male)": r.hhDeceased?"":r.hhAge,
        "Acting HH Age (Female)": r.hhDeceased?r.hhAge:"",
        "Spouse Age": r.spouseAge,
      };
      for(let i=0;i<maxM;i++) row[`Boy ${i+1} Age`]=r.maleChildAges[i]??"";
      for(let i=0;i<maxF;i++) row[`Girl ${i+1} Age`]=r.femaleChildAges[i]??"";
      row["HH Education"]=r.hhEducation;
      row["Spouse Education"]=r.spouseEducation;
      for(let i=0;i<maxM;i++) row[`Boy ${i+1} Education`]=r.maleChildEducations[i]??"";
      for(let i=0;i<maxF;i++) row[`Girl ${i+1} Education`]=r.femaleChildEducations[i]??"";
      row["No. of Disabled"]=r.disabledCount;
      for(let i=0;i<maxD;i++) row[`Disability Type ${i+1}`]=r.disabilityTypes[i]??"";
      row["Occupation"]=r.occupation;
      row["No. of Earners"]=r.earners?Number(r.earners):"";
      row["HP No."]=r.hpNo;
      row["Toilet No."]=r.toiletNo?Number(r.toiletNo):"";
      row["Means of Transport"]=r.transport;
      row["Income (PKR)"]=r.income?Number(r.income):"";
      row["OM Capacity"]=r.omCapacity?Number(r.omCapacity):"";
      return row;
    });
  };

  const handleExportCsv = () => {
    if (!records.length) { showToast("No records to export"); return; }
    const rows = buildRows();
    const headers = Object.keys(rows[0]);
    const escape = (v) => {
      const s = String(v ?? "");
      return s.includes(",") || s.includes('"') || s.includes("\n")
        ? `"${s.replace(/"/g,'""')}"` : s;
    };
    const csv = [
      headers.map(escape).join(","),
      ...rows.map(r => headers.map(h => escape(r[h])).join(",")),
    ].join("\n");
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `household_survey_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(a);
    a.click();
    setTimeout(()=>{ URL.revokeObjectURL(url); document.body.removeChild(a); }, 1000);
    showToast("📥 CSV downloaded!");
  };

  return (
    <div style={S.root}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
        *{box-sizing:border-box;}
        input:focus,select:focus{border-color:${C.accent}!important;outline:none;box-shadow:0 0 0 3px ${C.accent}22;}
        input[type=number]::-webkit-inner-spin-button{opacity:0.4;}
        select option{background:#0D1B2E;}
      `}</style>

      <div style={S.header}>
        <div style={S.htag}>Household Survey Tool</div>
        <h1 style={S.htitle}>HH Data Collector</h1>
        <p style={S.hsub}>{records.length} household{records.length!==1?"s":""} saved</p>
        <div style={S.progress}><div style={S.progressFill(mainTab==="form"?progress:100)}/></div>
      </div>

      <div style={S.wrap}>
        <div style={S.tabs}>
          <button style={S.tab(mainTab==="form")} onClick={()=>setMainTab("form")}>+ Add</button>
          <button style={S.tab(mainTab==="records")} onClick={()=>setMainTab("records")}>Records ({records.length})</button>
          <button style={S.tab(mainTab==="export")} onClick={()=>setMainTab("export")}>Export</button>
        </div>

        {mainTab==="form"&&(
          <>
            <div style={{display:"flex",gap:4,marginBottom:16}}>
              {SECTIONS.map((_,i)=>(
                <div key={i} style={{flex:1,height:4,borderRadius:2,background:i<=step?C.accent:C.border,transition:"background .3s"}}/>
              ))}
            </div>
            <div style={{fontSize:12,color:C.muted,marginBottom:4}}>
              Step {step+1} of {SECTIONS.length} — <span style={{color:C.text,fontWeight:600}}>{SECTIONS[step]}</span>
            </div>
            {steps[step]}
            <div style={S.navRow}>
              {step>0&&<button style={S.prevBtn} onClick={()=>setStep(s=>s-1)}>← Back</button>}
              {step<SECTIONS.length-1
                ?<button style={S.nextBtn} onClick={()=>setStep(s=>s+1)}>Next →</button>
                :<button style={S.addBtn} onClick={handleSave}>✓ Save Household</button>
              }
            </div>
          </>
        )}

        {mainTab==="records"&&(
          <>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
              {[
                ["Households",records.length,C.accent],
                ["Total Male",records.reduce((s,r)=>s+r.totalMale,0),"#60A5FA"],
                ["Total Female",records.reduce((s,r)=>s+r.totalFemale,0),"#A78BFA"],
              ].map(([l,v,color])=>(
                <div key={l} style={{background:C.card,border:`1px solid ${C.border}`,borderRadius:11,padding:"11px 8px",textAlign:"center"}}>
                  <div style={{fontSize:22,fontWeight:800,color}}>{v}</div>
                  <div style={{fontSize:9,color:C.muted,textTransform:"uppercase",letterSpacing:1}}>{l}</div>
                </div>
              ))}
            </div>
            {records.length===0
              ?<div style={{textAlign:"center",color:C.muted,padding:"48px 20px"}}>
                  <div style={{fontSize:36,marginBottom:10}}>🏠</div>
                  No households yet.<br/>Fill the form to get started.
                </div>
              :records.map(r=>(
                <div key={r.id} style={S.recordCard}>
                  <div>
                    <div style={{fontWeight:700,fontSize:14,display:"flex",alignItems:"center",gap:6}}>
                      {r.hhName||"—"}
                      {r.hhDeceased&&<span style={{fontSize:9,background:"#2A0A0A",color:"#EF4444",padding:"2px 6px",borderRadius:4,border:"1px solid #EF444433"}}>DECEASED</span>}
                    </div>
                    <div style={{fontSize:11,color:C.muted,marginTop:2}}>{r.hhCnic}{r.occupation?` · ${r.occupation}`:""}</div>
                    <div style={{display:"flex",gap:6,marginTop:6,flexWrap:"wrap"}}>
                      <span style={{...S.chip,color:"#60A5FA",background:"#1E3A5F"}}>♂ {r.totalMale}</span>
                      <span style={{...S.chip,color:"#A78BFA",background:"#2D1B4E"}}>♀ {r.totalFemale}</span>
                      {r.income&&<span style={S.chip}>PKR {Number(r.income).toLocaleString()}</span>}
                      {r.hpNo&&<span style={{...S.chip,color:C.textDim,background:C.accentDim}}>HP: {r.hpNo}</span>}
                    </div>
                  </div>
                  <button style={S.delBtn} onClick={()=>setRecords(rs=>rs.filter(x=>x.id!==r.id))}>✕</button>
                </div>
              ))
            }
          </>
        )}

        {mainTab==="export"&&(
          <>
            <div style={{...S.card,textAlign:"center",padding:"32px 20px"}}>
              <div style={{fontSize:48,marginBottom:14}}>📊</div>
              <div style={{fontSize:20,fontWeight:800,marginBottom:8}}>Export to Excel</div>
              <div style={{fontSize:13,color:C.muted,lineHeight:1.7}}>
                Download <span style={{color:C.accent,fontWeight:700}}>{records.length} household{records.length!==1?"s":""}</span> as
                Excel (.xlsx) or CSV — one row per household, with dynamic child columns.
              </div>
              {records.length>0&&(
                <div style={{marginTop:18,background:"#060D1A",borderRadius:10,border:`1px solid ${C.border}`,padding:"12px 14px",textAlign:"left"}}>
                  {[
                    ["Households",records.length],
                    ["Total People (M)",records.reduce((s,r)=>s+r.totalMale,0)],
                    ["Total People (F)",records.reduce((s,r)=>s+r.totalFemale,0)],
                    ["Deceased HHs",records.filter(r=>r.hhDeceased).length],
                    ["Total Income",`PKR ${records.reduce((s,r)=>s+(Number(r.income)||0),0).toLocaleString()}`],
                    ["File Formats",".xlsx · .csv"],
                  ].map(([k,v])=>(
                    <div key={k} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",fontSize:12,borderBottom:`1px solid ${C.border}`}}>
                      <span style={{color:C.muted}}>{k}</span>
                      <span style={{color:C.text,fontWeight:700}}>{v}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
            <button style={S.exportBtn} onClick={handleExport}>↓ Download Excel (.xlsx)</button>
            <button style={{...S.exportBtn, color:"#10B981", borderColor:"#10B981", marginTop:8}} onClick={handleExportCsv}>↓ Download CSV (.csv)</button>
            <button style={{...S.addBtn,background:C.accentDim,color:C.accent,marginTop:8}}
              onClick={()=>{setMainTab("form");setStep(0);}}>
              + Add Another Household
            </button>
          </>
        )}
      </div>

      <div style={S.toast(toast)}>{toastMsg}</div>
    </div>
  );
}
