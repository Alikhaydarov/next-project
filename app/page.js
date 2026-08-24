'use client';

import {useEffect,useMemo,useState} from 'react';
import * as XLSX from 'xlsx';

const API='https://ehfyhvfmdtbjipgqpvoq.supabase.co/functions/v1/payflow-api';
const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];
const money=n=>'₩'+Number(n||0).toLocaleString('en-US');
const localDate=()=>{const d=new Date();d.setMinutes(d.getMinutes()-d.getTimezoneOffset());return d.toISOString().slice(0,10)};
const dateText=p=>p.payment_date||new Date(p.paid_at).toLocaleDateString('en-CA');
const methodLabel=v=>v==='cash'?'Cash':'Card';
const formatAmount=v=>Number(v||0).toLocaleString('en-US');
const validDate=v=>/^\d{4}-\d{2}-\d{2}$/.test(v);
const receiptList=item=>Array.isArray(item?.receipt_urls)?item.receipt_urls:[];

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

async function prepareReceipts(files){
  const list=[];
  for(const file of Array.from(files||[]))list.push({name:file.name,data:await prepareReceipt(file)});
  return list;
}

export default function Home(){
  const [data,setData]=useState({payments:[]});
  const [tab,setTab]=useState('dashboard');
  const [form,setForm]=useState({company:'',method:'card',ourCard:'',clientCard:'',date:localDate(),amount:'',description:''});
  const [receipts,setReceipts]=useState([]);
  const [query,setQuery]=useState('');
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [notice,setNotice]=useState('');
  const [editing,setEditing]=useState(null);
  const [details,setDetails]=useState(null);
  const [deletingId,setDeletingId]=useState('');
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
    const excelRows=sorted.map((p,i)=>({'No.':i+1,'Date':dateText(p),'Company Name':p.company_name,'Payment Method':methodLabel(p.our_payment_method),'Our Card':p.our_payment_method==='card'?(p.our_account||'—'):'—','Client Card':p.our_payment_method==='card'?(p.company_account||'—'):'—','Amount':Number(p.amount)}));
    const wb=XLSX.utils.book_new(),ws=XLSX.utils.json_to_sheet(excelRows),last=Math.max(excelRows.length+1,2);
    ws['!cols']=[{wch:7},{wch:13},{wch:24},{wch:18},{wch:28},{wch:28},{wch:16}];ws['!autofilter']={ref:`A1:G${last}`};for(let r=2;r<=last;r++)if(ws[`G${r}`])ws[`G${r}`].z='₩#,##0';XLSX.utils.book_append_sheet(wb,ws,'Payments');XLSX.writeFile(wb,'payflow-payments.xlsx',{compression:true});
  }

  async function addReceiptFiles(files){try{const next=await prepareReceipts(files);setReceipts(v=>[...v,...next])}catch(e){flash(e.message)}}
  async function submit(e){
    e.preventDefault();const amount=Number(form.amount.replace(/\D/g,''));const cardsOk=form.method==='cash'||(form.ourCard.trim()&&form.clientCard.trim());
    if(!form.company.trim()||!cardsOk||!validDate(form.date)||!amount){flash('Complete required fields. Date must be YYYY-MM-DD');return}
    setSaving(true);
    try{
      const created=await request('POST',{action:'payment',company_name:form.company.trim(),our_payment_method:form.method,our_card_account_text:form.method==='card'?form.ourCard.trim():null,company_card_account_text:form.method==='card'?form.clientCard.trim():null,payment_date:form.date,amount,description:form.description.trim()||null});
      const id=created.created_payment_id;
      for(const image of receipts)await request('POST',{action:'receipt_add',id,receipt_data_url:image.data});
      await load();
      setForm({company:'',method:form.method,ourCard:'',clientCard:'',date:localDate(),amount:'',description:''});setReceipts([]);flash('Payment saved');setTimeout(exportExcel,100);
    }catch(e){flash(e.message)}finally{setSaving(false)}
  }
  async function removePayment(item){if(!window.confirm(`Delete ${item.company_name} payment ${money(item.amount)}?`))return;setDeletingId(item.id);try{await mutate({action:'payment_delete',id:item.id});if(details?.id===item.id)setDetails(null);flash('Payment deleted')}catch(e){flash(e.message)}finally{setDeletingId('')}}

  if(loading)return <div className="loading-screen"><div className="loader"/><strong>Loading PayFlow...</strong></div>;
  return <div className="app-shell">
    <header className="topbar"><div className="logo-wrap"><div className="logo">P</div><div><strong>PayFlow</strong><span>Payment management</span></div></div><nav className="tabs"><button className={tab==='dashboard'?'active':''} onClick={()=>setTab('dashboard')}>Dashboard</button><button className={tab==='calendar'?'active':''} onClick={()=>setTab('calendar')}>Calendar</button></nav><div className="top-actions"><span className="connected"><i/>Supabase connected</span><button className="secondary" onClick={exportExcel}>Download Excel</button></div></header>
    <main className="page">{tab==='dashboard'?<Dashboard form={form} setForm={setForm} receipts={receipts} addReceiptFiles={addReceiptFiles} setReceipts={setReceipts} submit={submit} saving={saving} stats={stats} rows={rows} query={query} setQuery={setQuery} setEditing={setEditing} setDetails={setDetails} removePayment={removePayment} deletingId={deletingId}/>:<Calendar years={years} year={year} setYear={y=>{setYear(y);setMonth(`${y}-01`)}} monthStats={monthStats} month={month} setMonth={setMonth} dayGroups={dayGroups} setDetails={setDetails}/>}</main>
    {editing&&<EditModal item={editing} onClose={()=>setEditing(null)} request={request} reload={load} flash={flash}/>} 
    {details&&<PaymentDetails item={details} onClose={()=>setDetails(null)} onEdit={()=>{setEditing(details);setDetails(null)}}/>}
    {notice&&<div className="toast">✓ {notice}</div>}
  </div>;
}

function Dashboard({form,setForm,receipts,addReceiptFiles,setReceipts,submit,saving,stats,rows,query,setQuery,setEditing,setDetails,removePayment,deletingId}){
  const set=(key,value)=>setForm(v=>({...v,[key]:value,...(key==='method'&&value==='cash'?{ourCard:'',clientCard:''}:{})}));
  return <><div className="page-heading"><div><h1>Payments</h1><p>Images and description are optional. Click a payment to view them.</p></div></div><section className="stats-grid"><Stat label="Today" value={money(stats.today)}/><Stat label="This month" value={money(stats.month)}/><Stat label="All-time total" value={money(stats.total)}/><Stat label="Payments" value={stats.count}/></section><section className="workspace">
    <div className="panel"><div className="panel-title"><h2>New payment</h2><p>Card: Our Card → Client Card. Cash: card fields disappear.</p></div><form className="payment-form" onSubmit={submit}>
      <Field label="Company name"><input value={form.company} onChange={e=>set('company',e.target.value)} placeholder="Enter company name"/></Field>
      <Field label="Payment method"><SelectControl><select value={form.method} onChange={e=>set('method',e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>
      {form.method==='card'&&<><Field label="Our card"><input value={form.ourCard} onChange={e=>set('ourCard',e.target.value)} placeholder="e.g. Shinhan 1010101010"/></Field><Field label="Client card"><input value={form.clientCard} onChange={e=>set('clientCard',e.target.value)} placeholder="e.g. Toss Bank 1010101010"/></Field></>}
      <div className="form-row"><Field label="Date (YYYY-MM-DD)"><input className="date-text-input" inputMode="numeric" maxLength={10} value={form.date} onChange={e=>set('date',e.target.value)} placeholder="2026-08-23"/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input inputMode="numeric" value={form.amount} placeholder="1,500,000" onChange={e=>{const v=e.target.value.replace(/\D/g,'');set('amount',v?Number(v).toLocaleString('en-US'):'')}}/></div></Field></div>
      <Field label="Description (optional)"><textarea className="description-input" rows={3} value={form.description} onChange={e=>set('description',e.target.value)} placeholder="Add a note about this payment..."/></Field>
      <MultiReceiptPicker receipts={receipts} onFiles={addReceiptFiles} onRemove={index=>setReceipts(v=>v.filter((_,i)=>i!==index))}/>
      <button className="save-button" disabled={saving}>{saving?'Saving payment...':'Save payment'}</button>
    </form></div>
    <History rows={rows} query={query} setQuery={setQuery} setEditing={setEditing} setDetails={setDetails} removePayment={removePayment} deletingId={deletingId}/>
  </section></>;
}

function MultiReceiptPicker({receipts,onFiles,onRemove}){return <><div className="receipt-field"><div><strong>Receipt images</strong><span>Optional · add as many as you need</span></div><label className="upload-button">Add images<input type="file" multiple accept="image/jpeg,image/png,image/webp" onChange={e=>{onFiles(e.target.files);e.target.value=''}}/></label></div>{receipts.length>0&&<div className="receipt-gallery draft-gallery">{receipts.map((r,i)=><div className="receipt-gallery-item" key={`${r.name}-${i}`}><img src={r.data} alt={`Receipt ${i+1}`}/><button type="button" onClick={()=>onRemove(i)}>×</button><span>{i+1}</span></div>)}</div>}</>}

function History({rows,query,setQuery,setEditing,setDetails,removePayment,deletingId}){return <div className="panel history-panel"><div className="history-head"><div><h2>Payment history</h2><p>{rows.length} records · click any payment for details</p></div><input className="search" placeholder="Search company or card..." value={query} onChange={e=>setQuery(e.target.value)}/></div><div className="table-wrap"><table><thead><tr><th>Company</th><th>Method</th><th>Our card</th><th>Client card</th><th>Date</th><th>Amount</th><th>Actions</th></tr></thead><tbody>{rows.map(item=><tr className="clickable-payment" key={item.id} onClick={()=>setDetails(item)}><td><strong>{item.company_name}</strong></td><td>{methodLabel(item.our_payment_method)}</td><td>{item.our_payment_method==='card'?(item.our_account||'—'):'—'}</td><td>{item.our_payment_method==='card'?(item.company_account||'—'):'—'}</td><td>{dateText(item)}</td><td className="amount-cell">{money(item.amount)}</td><td><div className="row-actions"><button className="edit-btn" onClick={e=>{e.stopPropagation();setEditing(item)}}>Edit</button><button className="delete-btn" onClick={e=>{e.stopPropagation();removePayment(item)}} disabled={deletingId===item.id}>{deletingId===item.id?'Deleting...':'Delete'}</button></div></td></tr>)}</tbody></table></div><div className="mobile-history">{rows.map(item=><article className="payment-card-mobile clickable-payment" key={item.id} onClick={()=>setDetails(item)}><div className="payment-card-top"><strong>{item.company_name}</strong><b>{money(item.amount)}</b></div><div className="payment-card-line"><span>Method</span><strong>{methodLabel(item.our_payment_method)}</strong></div>{item.our_payment_method==='card'&&<><div className="payment-card-line"><span>Our card</span><strong>{item.our_account||'—'}</strong></div><div className="payment-card-line"><span>Client card</span><strong>{item.company_account||'—'}</strong></div></>}<div className="payment-card-line"><span>Date</span><strong>{dateText(item)}</strong></div><div className="mobile-card-actions"><button className="edit-btn" onClick={e=>{e.stopPropagation();setEditing(item)}}>Edit</button><button className="delete-btn" onClick={e=>{e.stopPropagation();removePayment(item)}}>Delete</button></div></article>)}</div></div>}

function Calendar({years,year,setYear,monthStats,month,setMonth,dayGroups,setDetails}){const mi=Math.max(0,Number(month.slice(5,7))-1);return <><div className="calendar-heading"><div><h1>Payment Calendar</h1><p>Choose a year, then a month. Click a payment to open its details.</p></div><SelectControl className="year-select"><select value={year} onChange={e=>setYear(e.target.value)}>{years.map(y=><option key={y}>{y}</option>)}</select></SelectControl></div><section className="month-grid">{monthStats.map(m=><button key={m.key} className={`month-card ${month===m.key?'active':''}`} onClick={()=>setMonth(m.key)}><span>{m.name}</span><strong>{money(m.total)}</strong><small>{m.count} payments</small></button>)}</section><section className="panel calendar-list"><div className="history-head"><div><h2>{MONTHS[mi]} {year}</h2><p>Daily payment history</p></div></div>{dayGroups.length?dayGroups.map(([day,list])=><div className="day-group" key={day}><div className="day-head"><strong>{day}</strong><span>{list.length} payments · {money(list.reduce((s,p)=>s+Number(p.amount||0),0))}</span></div><div className="day-payments">{list.map(p=><button className="calendar-payment" key={p.id} onClick={()=>setDetails(p)}><div><strong>{p.company_name}</strong><span>{methodLabel(p.our_payment_method)}</span></div><b>{money(p.amount)}</b></button>)}</div></div>):<div className="calendar-empty">No payments in this month.</div>}</section></>}

function EditModal({item,onClose,request,reload,flash}){
  const [form,setForm]=useState({company:item.company_name||'',method:item.our_payment_method||'card',ourCard:item.our_account||'',clientCard:item.company_account||'',date:dateText(item),amount:formatAmount(item.amount),description:item.description||''});
  const [existing,setExisting]=useState(receiptList(item));
  const [newImages,setNewImages]=useState([]);
  const [saving,setSaving]=useState(false);
  const set=(k,v)=>setForm(x=>({...x,[k]:v,...(k==='method'&&v==='cash'?{ourCard:'',clientCard:''}:{})}));
  async function addFiles(files){try{const next=await prepareReceipts(files);setNewImages(v=>[...v,...next])}catch(e){flash(e.message)}}
  async function removeExisting(image){if(!window.confirm('Remove this image?'))return;try{await request('POST',{action:'receipt_remove',id:item.id,path:image.path});setExisting(v=>v.filter(x=>x.path!==image.path));flash('Image removed')}catch(e){flash(e.message)}}
  async function submit(e){
    e.preventDefault();const amount=Number(form.amount.replace(/\D/g,''));const cardsOk=form.method==='cash'||(form.ourCard.trim()&&form.clientCard.trim());if(!form.company.trim()||!cardsOk||!validDate(form.date)||!amount)return;
    setSaving(true);
    try{
      await request('POST',{action:'payment_update',id:item.id,company_name:form.company.trim(),our_payment_method:form.method,our_card_account_text:form.method==='card'?form.ourCard.trim():null,company_card_account_text:form.method==='card'?form.clientCard.trim():null,payment_date:form.date,amount,description:form.description.trim()||null});
      for(const image of newImages)await request('POST',{action:'receipt_add',id:item.id,receipt_data_url:image.data});
      await reload();flash('Payment updated');onClose();
    }catch(e){flash(e.message)}finally{setSaving(false)}
  }
  return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="edit-modal"><div className="modal-head"><div><h2>Edit payment</h2><p>Add or remove any number of images.</p></div><button className="modal-close" onClick={onClose}>×</button></div><form className="edit-form" onSubmit={submit}><Field label="Company name"><input value={form.company} onChange={e=>set('company',e.target.value)}/></Field><Field label="Payment method"><SelectControl><select value={form.method} onChange={e=>set('method',e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>{form.method==='card'&&<><Field label="Our card"><input value={form.ourCard} onChange={e=>set('ourCard',e.target.value)}/></Field><Field label="Client card"><input value={form.clientCard} onChange={e=>set('clientCard',e.target.value)}/></Field></>}<div className="form-row"><Field label="Date (YYYY-MM-DD)"><input className="date-text-input" value={form.date} maxLength={10} onChange={e=>set('date',e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input value={form.amount} inputMode="numeric" onChange={e=>{const v=e.target.value.replace(/\D/g,'');set('amount',v?Number(v).toLocaleString('en-US'):'')}}/></div></Field></div><Field label="Description (optional)"><textarea className="description-input" rows={3} value={form.description} onChange={e=>set('description',e.target.value)}/></Field><div className="receipt-field"><div><strong>Receipt images</strong><span>{existing.length+newImages.length} attached / selected</span></div><label className="upload-button">Add more<input type="file" multiple accept="image/jpeg,image/png,image/webp" onChange={e=>{addFiles(e.target.files);e.target.value=''}}/></label></div>{existing.length>0&&<div className="receipt-gallery">{existing.map((r,i)=><div className="receipt-gallery-item" key={r.path}><img src={r.url} alt={`Receipt ${i+1}`}/><button type="button" onClick={()=>removeExisting(r)}>×</button><span>{i+1}</span></div>)}</div>}{newImages.length>0&&<div className="receipt-gallery draft-gallery">{newImages.map((r,i)=><div className="receipt-gallery-item" key={`${r.name}-${i}`}><img src={r.data} alt={`New receipt ${i+1}`}/><button type="button" onClick={()=>setNewImages(v=>v.filter((_,x)=>x!==i))}>×</button><span>New</span></div>)}</div>}<div className="modal-actions"><button type="button" className="cancel-btn" onClick={onClose}>Cancel</button><button className="save-edit-btn" disabled={saving}>{saving?'Saving...':'Save changes'}</button></div></form></div></div>;
}

function PaymentDetails({item,onClose,onEdit}){
  const images=receiptList(item);
  const [active,setActive]=useState(images[0]?.url||null);
  return <div className="modal-backdrop details-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="payment-details-modal"><div className="modal-head"><div><h2>{item.company_name}</h2><p>{dateText(item)} · {methodLabel(item.our_payment_method)}</p></div><button className="modal-close" onClick={onClose}>×</button></div><div className="details-body"><div className="details-amount">{money(item.amount)}</div><div className="details-grid"><Detail label="Date" value={dateText(item)}/><Detail label="Payment method" value={methodLabel(item.our_payment_method)}/>{item.our_payment_method==='card'&&<><Detail label="Our card" value={item.our_account||'—'}/><Detail label="Client card" value={item.company_account||'—'}/></>}</div><div className="details-section"><span>Description</span><p>{item.description||'No description'}</p></div>{images.length>0&&<div className="details-section"><span>Receipt images · {images.length}</span>{active&&<img className="details-receipt" src={active} alt="Payment receipt"/>}<div className="receipt-gallery details-gallery">{images.map((r,i)=><button className={`detail-thumb ${active===r.url?'active':''}`} key={r.path||i} onClick={()=>setActive(r.url)}><img src={r.url} alt={`Receipt ${i+1}`}/><span>{i+1}</span></button>)}</div></div>}</div><div className="modal-actions details-actions"><button className="cancel-btn" onClick={onClose}>Close</button><button className="save-edit-btn" onClick={onEdit}>Edit payment</button></div></div></div>
}
function Detail({label,value}){return <div className="detail-item"><span>{label}</span><strong>{value}</strong></div>}
function Stat({label,value}){return <article className="stat-card"><span>{label}</span><strong>{value}</strong></article>}
function Field({label,children}){return <label className="field"><span>{label}</span>{children}</label>}
function SelectControl({children,className=''}){return <div className={`select-control ${className}`}>{children}<span className="select-arrow">⌄</span></div>}
