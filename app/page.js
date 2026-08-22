'use client';

import {useEffect,useMemo,useState} from 'react';
import * as XLSX from 'xlsx';

const API='https://ehfyhvfmdtbjipgqpvoq.supabase.co/functions/v1/payflow-api';
const empty={payments:[]};
const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
const money=n=>'₩'+Number(n||0).toLocaleString('en-US');
const localDate=()=>{const d=new Date();d.setMinutes(d.getMinutes()-d.getTimezoneOffset());return d.toISOString().slice(0,10)};
const dateText=p=>p.payment_date||new Date(p.paid_at).toLocaleDateString('en-CA');
const methodLabel=v=>v==='cash'?'Cash':'Card';
const formatAmount=v=>Number(v||0).toLocaleString('en-US');

async function prepareReceipt(file){
  if(!file)return null;
  if(!['image/jpeg','image/png','image/webp'].includes(file.type))throw new Error('Receipt must be JPG, PNG, or WEBP');
  const src=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(new Error('Could not read image'));r.readAsDataURL(file)});
  const img=await new Promise((resolve,reject)=>{const i=new Image();i.onload=()=>resolve(i);i.onerror=()=>reject(new Error('Could not open image'));i.src=src});
  const scale=Math.min(1,1600/Math.max(img.width,img.height));
  const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(img.width*scale));canvas.height=Math.max(1,Math.round(img.height*scale));
  canvas.getContext('2d').drawImage(img,0,0,canvas.width,canvas.height);
  let q=.84,data=canvas.toDataURL('image/jpeg',q);
  while(data.length>6_000_000&&q>.5){q-=.08;data=canvas.toDataURL('image/jpeg',q)}
  if(data.length>7_000_000)throw new Error('Receipt image is too large');
  return data;
}

export default function Home(){
  const [data,setData]=useState(empty);
  const [tab,setTab]=useState('dashboard');
  const [form,setForm]=useState({company:'',method:'card',ourCard:'',clientCard:'',date:localDate(),amount:''});
  const [receipt,setReceipt]=useState(null);
  const [receiptName,setReceiptName]=useState('');
  const [query,setQuery]=useState('');
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [notice,setNotice]=useState('');
  const [editing,setEditing]=useState(null);
  const [deletingId,setDeletingId]=useState('');
  const [preview,setPreview]=useState(null);
  const [year,setYear]=useState(localDate().slice(0,4));
  const [month,setMonth]=useState(localDate().slice(0,7));

  const flash=t=>{setNotice(t);setTimeout(()=>setNotice(''),2400)};
  async function request(method='GET',body){const r=await fetch(API,{method,headers:{'Content-Type':'application/json'},body:body?JSON.stringify(body):undefined,cache:'no-store'});const j=await r.json().catch(()=>({}));if(!r.ok)throw new Error(j.error||'Database request failed');return j}
  async function load(){try{setData(await request())}catch(e){flash(e.message)}finally{setLoading(false)}}
  async function mutate(body){const j=await request('POST',body);setData(j);return j}
  useEffect(()=>{load()},[]);

  const stats=useMemo(()=>{const today=localDate(),m=today.slice(0,7);let t=0,mt=0,all=0;data.payments.forEach(p=>{const v=Number(p.amount||0),d=dateText(p);all+=v;if(d===today)t+=v;if(d.startsWith(m))mt+=v});return{today:t,month:mt,total:all,count:data.payments.length}},[data.payments]);
  const rows=data.payments.filter(p=>!query||Object.values(p).join(' ').toLowerCase().includes(query.toLowerCase()));
  const years=useMemo(()=>{const s=new Set([localDate().slice(0,4)]);data.payments.forEach(p=>s.add(dateText(p).slice(0,4)));return[...s].sort((a,b)=>b.localeCompare(a))},[data.payments]);
  const monthStats=useMemo(()=>MONTHS.map((name,i)=>{const key=`${year}-${String(i+1).padStart(2,'0')}`;const list=data.payments.filter(p=>dateText(p).startsWith(key));return{name,key,count:list.length,total:list.reduce((s,p)=>s+Number(p.amount||0),0)}}),[data.payments,year]);
  const dayGroups=useMemo(()=>{const map=new Map();data.payments.filter(p=>dateText(p).startsWith(month)).forEach(p=>{const d=dateText(p);if(!map.has(d))map.set(d,[]);map.get(d).push(p)});return[...map.entries()].sort((a,b)=>b[0].localeCompare(a[0]))},[data.payments,month]);

  function exportExcel(){
    const sorted=[...data.payments].sort((a,b)=>new Date(b.created_at||b.paid_at)-new Date(a.created_at||a.paid_at));
    const rows=sorted.map((p,i)=>({'No.':i+1,'Date':dateText(p),'Company Name':p.company_name,'Payment Method':methodLabel(p.our_payment_method),'Our Card':p.our_payment_method==='card'?(p.our_account||'—'):'—','Client Card':p.our_payment_method==='card'?(p.company_account||'—'):'—','Amount':Number(p.amount)}));
    const wb=XLSX.utils.book_new();const ws=XLSX.utils.json_to_sheet(rows);const last=Math.max(rows.length+1,2);ws['!cols']=[{wch:7},{wch:13},{wch:24},{wch:18},{wch:28},{wch:28},{wch:16}];ws['!autofilter']={ref:`A1:G${last}`};for(let r=2;r<=last;r++)if(ws[`G${r}`])ws[`G${r}`].z='₩#,##0';XLSX.utils.book_append_sheet(wb,ws,'Payments');XLSX.writeFile(wb,'payflow-payments.xlsx',{compression:true});
  }

  async function chooseReceipt(file){try{setReceipt(await prepareReceipt(file));setReceiptName(file?.name||'Receipt')}catch(e){flash(e.message)}}
  async function submit(e){e.preventDefault();const amount=Number(form.amount.replace(/\D/g,''));const cardsOk=form.method==='cash'||(form.ourCard.trim()&&form.clientCard.trim());if(!form.company.trim()||!cardsOk||!form.date||!amount){flash('Please complete all required fields');return}setSaving(true);try{const fresh=await mutate({action:'payment',company_name:form.company.trim(),our_payment_method:form.method,our_card_account_text:form.method==='card'?form.ourCard.trim():null,company_card_account_text:form.method==='card'?form.clientCard.trim():null,payment_date:form.date,amount,receipt_data_url:receipt});setForm({company:'',method:form.method,ourCard:'',clientCard:'',date:localDate(),amount:''});setReceipt(null);setReceiptName('');flash('Payment saved');setTimeout(exportExcel,100)}catch(e){flash(e.message)}finally{setSaving(false)}}
  async function removePayment(item){if(!window.confirm(`Delete ${item.company_name} payment ${money(item.amount)}?`))return;setDeletingId(item.id);try{await mutate({action:'payment_delete',id:item.id});flash('Payment deleted')}catch(e){flash(e.message)}finally{setDeletingId('')}}

  if(loading)return <div className="loading-screen"><div className="loader"/><strong>Loading PayFlow...</strong></div>;
  return <div className="app-shell">
    <header className="topbar"><div className="logo-wrap"><div className="logo">P</div><div><strong>PayFlow</strong><span>Payment management</span></div></div><nav className="tabs"><button className={tab==='dashboard'?'active':''} onClick={()=>setTab('dashboard')}>Dashboard</button><button className={tab==='calendar'?'active':''} onClick={()=>setTab('calendar')}>Calendar</button></nav><div className="top-actions"><span className="connected"><i/>Supabase connected</span><button className="secondary" onClick={exportExcel}>Download Excel</button></div></header>
    <main className="page">
      {tab==='dashboard'?<Dashboard data={data} form={form} setForm={setForm} receipt={receipt} receiptName={receiptName} chooseReceipt={chooseReceipt} clearReceipt={()=>{setReceipt(null);setReceiptName('')}} submit={submit} saving={saving} stats={stats} rows={rows} query={query} setQuery={setQuery} setEditing={setEditing} removePayment={removePayment} deletingId={deletingId} setPreview={setPreview}/>:<Calendar payments={data.payments} years={years} year={year} setYear={y=>{setYear(y);setMonth(`${y}-01`)}} monthStats={monthStats} month={month} setMonth={setMonth} dayGroups={dayGroups} setPreview={setPreview}/>} 
    </main>
    {editing&&<EditModal item={editing} onClose={()=>setEditing(null)} setPreview={setPreview} onSave={async body=>{await mutate(body);setEditing(null);flash('Payment updated')}}/>}
    {preview&&<ImageModal src={preview} onClose={()=>setPreview(null)}/>} 
    {notice&&<div className="toast">✓ {notice}</div>}
  </div>;
}

function Dashboard({form,setForm,receipt,receiptName,chooseReceipt,clearReceipt,submit,saving,stats,rows,query,setQuery,setEditing,removePayment,deletingId,setPreview}){
  const set=(key,value)=>setForm(v=>({...v,[key]:value,...(key==='method'&&value==='cash'?{ourCard:'',clientCard:''}:{})}));
  return <><div className="page-heading"><div><h1>Payments</h1><p>Enter payment details directly. Receipt image is optional.</p></div></div><section className="stats-grid"><Stat label="Today" value={money(stats.today)}/><Stat label="This month" value={money(stats.month)}/><Stat label="All-time total" value={money(stats.total)}/><Stat label="Payments" value={stats.count}/></section><section className="workspace">
    <div className="panel"><div className="panel-title"><h2>New payment</h2><p>Card: Our Card → Client Card. Cash: card fields disappear.</p></div><form className="payment-form" onSubmit={submit}>
      <Field label="Company name"><input value={form.company} onChange={e=>set('company',e.target.value)} placeholder="Enter company name"/></Field>
      <Field label="Payment method"><SelectControl><select value={form.method} onChange={e=>set('method',e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>
      {form.method==='card'&&<CardFields ourCard={form.ourCard} clientCard={form.clientCard} setOur={v=>set('ourCard',v)} setClient={v=>set('clientCard',v)}/>} 
      <div className="form-row"><Field label="Date"><input type="date" value={form.date} onChange={e=>set('date',e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input inputMode="numeric" value={form.amount} placeholder="1,500,000" onChange={e=>{const v=e.target.value.replace(/\D/g,'');set('amount',v?Number(v).toLocaleString('en-US'):'')}}/></div></Field></div>
      <ReceiptPicker receipt={receipt} receiptName={receiptName} chooseReceipt={chooseReceipt} clearReceipt={clearReceipt} setPreview={setPreview}/>
      <button className="save-button" disabled={saving}>{saving?'Saving payment...':'Save payment'}</button>
    </form></div>
    <History rows={rows} total={rows.length} query={query} setQuery={setQuery} setEditing={setEditing} removePayment={removePayment} deletingId={deletingId} setPreview={setPreview}/>
  </section></>;
}

function CardFields({ourCard,clientCard,setOur,setClient}){return <><Field label="Our card"><input value={ourCard} onChange={e=>setOur(e.target.value)} placeholder="e.g. Shinhan 1010101010"/></Field><Field label="Client card"><input value={clientCard} onChange={e=>setClient(e.target.value)} placeholder="e.g. Toss Bank 1010101010"/></Field></>}

function ReceiptPicker({receipt,receiptName,chooseReceipt,clearReceipt,setPreview}){return <><div className="receipt-field"><div><strong>Receipt image</strong><span>Optional · JPG, PNG or WEBP</span></div><label className="upload-button">{receipt?'Change image':'Add image'}<input type="file" accept="image/jpeg,image/png,image/webp" onChange={e=>chooseReceipt(e.target.files?.[0])}/></label></div>{receipt&&<div className="receipt-selected"><img src={receipt} alt="Receipt" onClick={()=>setPreview(receipt)}/><div><strong>{receiptName}</strong><button type="button" onClick={clearReceipt}>Remove</button></div></div>}</>}

function History({rows,total,query,setQuery,setEditing,removePayment,deletingId,setPreview}){return <div className="panel history-panel"><div className="history-head"><div><h2>Payment history</h2><p>{total} records · newest first</p></div><input className="search" placeholder="Search company or card..." value={query} onChange={e=>setQuery(e.target.value)}/></div><div className="table-wrap"><table><thead><tr><th>Company</th><th>Method</th><th>Our card</th><th>Client card</th><th>Receipt</th><th>Date</th><th>Amount</th><th>Actions</th></tr></thead><tbody>{rows.map(item=><tr key={item.id}><td><strong>{item.company_name}</strong></td><td>{methodLabel(item.our_payment_method)}</td><td>{item.our_payment_method==='card'?(item.our_account||'—'):'—'}</td><td>{item.our_payment_method==='card'?(item.company_account||'—'):'—'}</td><td>{item.receipt_url?<img className="receipt-thumb" src={item.receipt_url} alt="Receipt" onClick={()=>setPreview(item.receipt_url)}/>:<span>—</span>}</td><td>{dateText(item)}</td><td className="amount-cell">{money(item.amount)}</td><td><div className="row-actions"><button className="edit-btn" onClick={()=>setEditing(item)}>Edit</button><button className="delete-btn" onClick={()=>removePayment(item)} disabled={deletingId===item.id}>{deletingId===item.id?'Deleting...':'Delete'}</button></div></td></tr>)}</tbody></table></div><div className="mobile-history">{rows.map(item=><article className="payment-card-mobile" key={item.id}><div className="payment-card-top"><strong>{item.company_name}</strong><b>{money(item.amount)}</b></div><div className="payment-card-line"><span>Method</span><strong>{methodLabel(item.our_payment_method)}</strong></div>{item.our_payment_method==='card'&&<><div className="payment-card-line"><span>Our card</span><strong>{item.our_account||'—'}</strong></div><div className="payment-card-line"><span>Client card</span><strong>{item.company_account||'—'}</strong></div></>}<div className="payment-card-line"><span>Date</span><strong>{dateText(item)}</strong></div>{item.receipt_url&&<button className="mobile-receipt" onClick={()=>setPreview(item.receipt_url)}><img src={item.receipt_url} alt="Receipt"/><span>View receipt</span></button>}<div className="mobile-card-actions"><button className="edit-btn" onClick={()=>setEditing(item)}>Edit</button><button className="delete-btn" onClick={()=>removePayment(item)}>Delete</button></div></article>)}</div></div>}

function Calendar({payments,years,year,setYear,monthStats,month,setMonth,dayGroups,setPreview}){const mi=Math.max(0,Number(month.slice(5,7))-1);const monthPayments=payments.filter(p=>dateText(p).startsWith(month));return <><div className="calendar-heading"><div><h1>Payment Calendar</h1><p>Choose a year, then a month to see daily payment history.</p></div><SelectControl className="year-select"><select value={year} onChange={e=>setYear(e.target.value)}>{years.map(y=><option key={y}>{y}</option>)}</select></SelectControl></div><section className="month-grid">{monthStats.map(m=><button key={m.key} className={`month-card ${month===m.key?'selected':''}`} onClick={()=>setMonth(m.key)}><span>{m.name}</span><strong>{money(m.total)}</strong><small>{m.count} payment{m.count===1?'':'s'}</small></button>)}</section><section className="panel calendar-history"><div className="calendar-history-head"><div><h2>{MONTHS[mi]} {year}</h2><p>{monthPayments.length} payments</p></div><strong>{money(monthPayments.reduce((s,p)=>s+Number(p.amount||0),0))}</strong></div><div className="day-groups">{dayGroups.length?dayGroups.map(([day,items])=><div className="day-group" key={day}><div className="day-head"><div><strong>{new Date(`${day}T12:00:00`).toLocaleDateString('en-US',{weekday:'long',month:'short',day:'numeric'})}</strong><span>{items.length} payment{items.length===1?'':'s'}</span></div><b>{money(items.reduce((s,p)=>s+Number(p.amount||0),0))}</b></div><div className="day-payments">{items.map(p=><div className="calendar-payment" key={p.id}><div className="calendar-company"><strong>{p.company_name}</strong><span>{methodLabel(p.our_payment_method)}{p.our_payment_method==='card'&&p.our_account?` · ${p.our_account}`:''}</span></div>{p.receipt_url?<img className="receipt-thumb" src={p.receipt_url} alt="Receipt" onClick={()=>setPreview(p.receipt_url)}/>:<span/>}<b>{money(p.amount)}</b></div>)}</div></div>):<div className="calendar-empty">No payments for {MONTHS[mi]} {year}.</div>}</div></section></>}

function EditModal({item,onClose,onSave,setPreview}){
  const [f,setF]=useState({company:item.company_name||'',method:item.our_payment_method||'card',our:item.our_account||'',client:item.company_account||'',date:dateText(item),amount:formatAmount(item.amount)});
  const [newReceipt,setNewReceipt]=useState(null);const [removed,setRemoved]=useState(false);const [saving,setSaving]=useState(false);
  const set=(k,v)=>setF(x=>({...x,[k]:v,...(k==='method'&&v==='cash'?{our:'',client:''}:{})}));
  async function pick(file){try{setNewReceipt(await prepareReceipt(file));setRemoved(false)}catch(e){alert(e.message)}}
  async function submit(e){e.preventDefault();const a=Number(f.amount.replace(/\D/g,''));if(!f.company.trim()||!f.date||!a||(f.method==='card'&&(!f.our.trim()||!f.client.trim())))return;setSaving(true);try{await onSave({action:'payment_update',id:item.id,company_name:f.company.trim(),our_payment_method:f.method,our_card_account_text:f.method==='card'?f.our.trim():null,company_card_account_text:f.method==='card'?f.client.trim():null,payment_date:f.date,amount:a,receipt_data_url:newReceipt,remove_receipt:removed})}finally{setSaving(false)}}
  const shown=newReceipt||(!removed?item.receipt_url:null);
  return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="edit-modal"><div className="modal-head"><div><h2>Edit payment</h2><p>Change details or receipt.</p></div><button className="modal-close" onClick={onClose}>×</button></div><form className="edit-form" onSubmit={submit}><Field label="Company name"><input value={f.company} onChange={e=>set('company',e.target.value)}/></Field><Field label="Payment method"><SelectControl><select value={f.method} onChange={e=>set('method',e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>{f.method==='card'&&<CardFields ourCard={f.our} clientCard={f.client} setOur={v=>set('our',v)} setClient={v=>set('client',v)}/>}<div className="form-row"><Field label="Date"><input type="date" value={f.date} onChange={e=>set('date',e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input value={f.amount} onChange={e=>{const v=e.target.value.replace(/\D/g,'');set('amount',v?Number(v).toLocaleString('en-US'):'')}}/></div></Field></div><div className="receipt-field"><div><strong>Receipt image</strong><span>Optional</span></div><label className="upload-button">{shown?'Replace':'Add image'}<input type="file" accept="image/jpeg,image/png,image/webp" onChange={e=>pick(e.target.files?.[0])}/></label></div>{shown&&<div className="receipt-selected"><img src={shown} alt="Receipt" onClick={()=>setPreview(shown)}/><div><strong>Receipt attached</strong><button type="button" onClick={()=>{setNewReceipt(null);setRemoved(true)}}>Remove</button></div></div>}<div className="modal-actions"><button type="button" className="cancel-btn" onClick={onClose}>Cancel</button><button className="save-edit-btn" disabled={saving}>{saving?'Saving...':'Save changes'}</button></div></form></div></div>;
}

function ImageModal({src,onClose}){return <div className="image-modal" onClick={onClose}><button onClick={onClose}>×</button><img src={src} alt="Payment receipt" onClick={e=>e.stopPropagation()}/></div>}
function Stat({label,value}){return <article className="stat-card"><span>{label}</span><strong>{value}</strong></article>}
function Field({label,children}){return <label className="field"><span>{label}</span>{children}</label>}
function SelectControl({children,className=''}){return <div className={`select-control ${className}`}>{children}<span className="select-arrow">⌄</span></div>}
