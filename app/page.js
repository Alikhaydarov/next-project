'use client';

import {useEffect,useMemo,useState} from 'react';
import * as XLSX from 'xlsx';

const API='https://ehfyhvfmdtbjipgqpvoq.supabase.co/functions/v1/payflow-api';
const empty={mode:'database',companies:[],our_accounts:[],company_accounts:[],payments:[]};
const money=n=>'₩'+Number(n||0).toLocaleString('en-US');
const localDate=()=>{const d=new Date();d.setMinutes(d.getMinutes()-d.getTimezoneOffset());return d.toISOString().slice(0,10)};
const dateText=item=>item.payment_date||new Date(item.paid_at).toLocaleDateString('en-CA');
const formatAmount=value=>Number(value||0).toLocaleString('en-US');
const methodLabel=value=>value==='cash'?'Cash':'Card';
const MONTHS=['January','February','March','April','May','June','July','August','September','October','November','December'];

async function prepareReceipt(file){
  if(!file)return null;
  if(!['image/jpeg','image/png','image/webp'].includes(file.type))throw new Error('Receipt must be JPG, PNG, or WEBP');
  const source=await new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(new Error('Could not read image'));r.readAsDataURL(file)});
  const image=await new Promise((resolve,reject)=>{const img=new Image();img.onload=()=>resolve(img);img.onerror=()=>reject(new Error('Could not open image'));img.src=source});
  const max=1600,scale=Math.min(1,max/Math.max(image.width,image.height));
  const canvas=document.createElement('canvas');canvas.width=Math.max(1,Math.round(image.width*scale));canvas.height=Math.max(1,Math.round(image.height*scale));
  const ctx=canvas.getContext('2d');ctx.drawImage(image,0,0,canvas.width,canvas.height);
  let quality=.84,data=canvas.toDataURL('image/jpeg',quality);
  while(data.length>6_000_000&&quality>.5){quality-=.08;data=canvas.toDataURL('image/jpeg',quality)}
  if(data.length>7_000_000)throw new Error('Receipt image is too large');
  return data;
}

export default function Home(){
  const [data,setData]=useState(empty);
  const [tab,setTab]=useState('dashboard');
  const [companyName,setCompanyName]=useState('');
  const [paymentMethod,setPaymentMethod]=useState('card');
  const [ourCard,setOurCard]=useState('');
  const [clientCard,setClientCard]=useState('');
  const [paidAt,setPaidAt]=useState(localDate());
  const [amount,setAmount]=useState('');
  const [receiptData,setReceiptData]=useState(null);
  const [receiptName,setReceiptName]=useState('');
  const [query,setQuery]=useState('');
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [notice,setNotice]=useState('');
  const [editing,setEditing]=useState(null);
  const [deletingId,setDeletingId]=useState('');
  const [imagePreview,setImagePreview]=useState(null);
  const [selectedYear,setSelectedYear]=useState(localDate().slice(0,4));
  const [selectedMonth,setSelectedMonth]=useState(localDate().slice(0,7));

  const flash=text=>{setNotice(text);setTimeout(()=>setNotice(''),2400)};
  async function request(method='GET',body){
    const response=await fetch(API,{method,headers:{'Content-Type':'application/json'},body:body?JSON.stringify(body):undefined,cache:'no-store'});
    const result=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(result.error||'Database request failed');
    return result;
  }
  async function load(){try{setData(await request())}catch(e){flash(e.message)}finally{setLoading(false)}}
  useEffect(()=>{load()},[]);
  useEffect(()=>{if(paymentMethod==='cash'){setOurCard('');setClientCard('')}},[paymentMethod]);

  const stats=useMemo(()=>{
    const today=localDate(),month=today.slice(0,7);let todayTotal=0,monthTotal=0,total=0;
    data.payments.forEach(item=>{const value=Number(item.amount||0),day=dateText(item);total+=value;if(day===today)todayTotal+=value;if(day.slice(0,7)===month)monthTotal+=value});
    return {todayTotal,monthTotal,total,count:data.payments.length};
  },[data.payments]);
  const rows=data.payments.filter(item=>!query||Object.values(item).join(' ').toLowerCase().includes(query.toLowerCase()));
  const years=useMemo(()=>{
    const set=new Set([localDate().slice(0,4)]);data.payments.forEach(p=>set.add(dateText(p).slice(0,4)));return [...set].sort((a,b)=>b.localeCompare(a));
  },[data.payments]);
  const monthStats=useMemo(()=>MONTHS.map((name,index)=>{
    const key=`${selectedYear}-${String(index+1).padStart(2,'0')}`;
    const payments=data.payments.filter(p=>dateText(p).startsWith(key));
    return {name,key,count:payments.length,total:payments.reduce((s,p)=>s+Number(p.amount||0),0)};
  }),[data.payments,selectedYear]);
  const calendarGroups=useMemo(()=>{
    const map=new Map();
    data.payments.filter(p=>dateText(p).startsWith(selectedMonth)).forEach(p=>{
      const day=dateText(p);if(!map.has(day))map.set(day,[]);map.get(day).push(p);
    });
    return [...map.entries()].sort((a,b)=>b[0].localeCompare(a[0]));
  },[data.payments,selectedMonth]);

  async function mutate(body){const result=await request('POST',body);setData(result);return result}

  function exportExcel(payments=data.payments){
    const sorted=[...payments].sort((a,b)=>new Date(b.created_at||b.paid_at)-new Date(a.created_at||a.paid_at));
    const excelRows=sorted.map((p,index)=>({
      'No.':index+1,'Date':dateText(p),'Company Name':p.company_name,'Payment Method':methodLabel(p.our_payment_method),
      'Our Card':p.our_payment_method==='card'?(p.our_account||'—'):'—','Client Card':p.our_payment_method==='card'?(p.company_account||'—'):'—','Amount':Number(p.amount)
    }));
    const workbook=XLSX.utils.book_new();
    const paymentsSheet=XLSX.utils.json_to_sheet(excelRows);const lastRow=Math.max(excelRows.length+1,2);
    paymentsSheet['!cols']=[{wch:7},{wch:13},{wch:24},{wch:18},{wch:28},{wch:28},{wch:16}];paymentsSheet['!autofilter']={ref:`A1:G${lastRow}`};paymentsSheet['!rows']=[{hpt:24}];
    for(let row=2;row<=lastRow;row++)if(paymentsSheet[`G${row}`])paymentsSheet[`G${row}`].z='₩#,##0';XLSX.utils.book_append_sheet(workbook,paymentsSheet,'Payments');
    const companyMap=new Map();sorted.forEach(p=>{const c=companyMap.get(p.company_name)||{count:0,total:0,lastDate:''};c.count++;c.total+=Number(p.amount||0);if(!c.lastDate||dateText(p)>c.lastDate)c.lastDate=dateText(p);companyMap.set(p.company_name,c)});
    const companyRows=[...companyMap.entries()].map(([name,v])=>({'Company':name,'Payments':v.count,'Total Amount':v.total,'Latest Payment':v.lastDate})).sort((a,b)=>b['Total Amount']-a['Total Amount']);
    const companySheet=XLSX.utils.json_to_sheet(companyRows);companySheet['!cols']=[{wch:28},{wch:12},{wch:18},{wch:16}];companySheet['!autofilter']={ref:`A1:D${Math.max(companyRows.length+1,2)}`};for(let row=2;row<=companyRows.length+1;row++)if(companySheet[`C${row}`])companySheet[`C${row}`].z='₩#,##0';XLSX.utils.book_append_sheet(workbook,companySheet,'Company Summary');
    const monthMap=new Map();sorted.forEach(p=>{const m=dateText(p).slice(0,7),c=monthMap.get(m)||{count:0,total:0};c.count++;c.total+=Number(p.amount||0);monthMap.set(m,c)});
    const monthRows=[...monthMap.entries()].map(([month,v])=>({'Month':month,'Payments':v.count,'Total Amount':v.total})).sort((a,b)=>b.Month.localeCompare(a.Month));const monthSheet=XLSX.utils.json_to_sheet(monthRows);monthSheet['!cols']=[{wch:14},{wch:12},{wch:18}];monthSheet['!autofilter']={ref:`A1:C${Math.max(monthRows.length+1,2)}`};for(let row=2;row<=monthRows.length+1;row++)if(monthSheet[`C${row}`])monthSheet[`C${row}`].z='₩#,##0';XLSX.utils.book_append_sheet(workbook,monthSheet,'Monthly Summary');
    XLSX.writeFile(workbook,'payflow-payments.xlsx',{compression:true});
  }

  async function pickReceipt(file){try{const prepared=await prepareReceipt(file);setReceiptData(prepared);setReceiptName(file.name)}catch(e){flash(e.message)}}
  async function submitPayment(e){
    e.preventDefault();const numericAmount=Number(amount.replace(/\D/g,''));const cardsOk=paymentMethod==='cash'||(ourCard.trim()&&clientCard.trim());
    if(!companyName.trim()||!cardsOk||!paidAt||!numericAmount){flash('Please complete all required fields');return}
    setSaving(true);
    try{
      const fresh=await mutate({action:'payment',company_name:companyName.trim(),our_payment_method:paymentMethod,our_card_account_text:paymentMethod==='card'?ourCard.trim():null,company_card_account_text:paymentMethod==='card'?clientCard.trim():null,payment_date:paidAt,amount:numericAmount,receipt_data_url:receiptData});
      setCompanyName('');setOurCard('');setClientCard('');setAmount('');setPaidAt(localDate());setReceiptData(null);setReceiptName('');flash('Payment saved');setTimeout(()=>exportExcel(fresh.payments),100);
    }catch(e){flash(e.message)}finally{setSaving(false)}
  }
  async function deletePayment(item){if(!window.confirm(`Delete ${item.company_name} payment ${money(item.amount)}?`))return;setDeletingId(item.id);try{await mutate({action:'payment_delete',id:item.id});flash('Payment deleted')}catch(e){flash(e.message)}finally{setDeletingId('')}}

  if(loading)return <div className="loading-screen"><div className="loader"/><strong>Loading PayFlow...</strong></div>;

  return <div className="app-shell">
    <header className="topbar">
      <div className="logo-wrap"><div className="logo">P</div><div><strong>PayFlow</strong><span>Payment management</span></div></div>
      <nav className="tabs"><button className={tab==='dashboard'?'active':''} onClick={()=>setTab('dashboard')}>Dashboard</button><button className={tab==='calendar'?'active':''} onClick={()=>setTab('calendar')}>Calendar</button></nav>
      <div className="top-actions"><span className="connected"><i/>Supabase connected</span><button className="secondary" onClick={()=>exportExcel()}>Download Excel</button></div>
    </header>

    <main className="page">
      {tab==='dashboard'?<>
        <div className="page-heading"><div><h1>Payments</h1><p>Enter payment details directly. Receipt image is optional.</p></div></div>
        <section className="stats-grid"><Stat label="Today" value={money(stats.todayTotal)}/><Stat label="This month" value={money(stats.monthTotal)}/><Stat label="All-time total" value={money(stats.total)}/><Stat label="Payments" value={stats.count}/></section>
        <section className="workspace">
          <div className="panel payment-panel">
            <div className="panel-title"><div><h2>New payment</h2><p>Card: Our Card → Client Card. Cash: card fields disappear.</p></div></div>
            <form className="payment-form" onSubmit={submitPayment}>
              <Field label="Company name"><input value={companyName} onChange={e=>setCompanyName(e.target.value)} placeholder="Enter company name" autoComplete="off"/></Field>
              <Field label="Payment method"><SelectControl><select value={paymentMethod} onChange={e=>setPaymentMethod(e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>
              {paymentMethod==='card'&&<><Field label="Our card"><input value={ourCard} onChange={e=>setOurCard(e.target.value)} placeholder="e.g. Shinhan 1010101010"/></Field><Field label="Client card"><input value={clientCard} onChange={e=>setClientCard(e.target.value)} placeholder="e.g. Toss Bank 1010101010"/></Field></>}
              <div className="form-row"><Field label="Date"><input type="date" value={paidAt} onChange={e=>setPaidAt(e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input inputMode="numeric" value={amount} placeholder="1,500,000" onChange={e=>{const v=e.target.value.replace(/\D/g,'');setAmount(v?Number(v).toLocaleString('en-US'):'')}}/></div></Field></div>
              <div className="receipt-field"><div><strong>Receipt image</strong><span>Optional · JPG, PNG or WEBP</span></div><label className="upload-button">{receiptData?'Change image':'Add image'}<input type="file" accept="image/jpeg,image/png,image/webp" onChange={e=>pickReceipt(e.target.files?.[0])}/></label></div>
              {receiptData&&<div className="receipt-selected"><img src={receiptData} alt="Receipt preview" onClick={()=>setImagePreview(receiptData)}/><div><strong>{receiptName||'Receipt image'}</strong><button type="button" onClick={()=>{setReceiptData(null);setReceiptName('')}}>Remove</button></div></div>}
              <button className="save-button" disabled={saving}>{saving?'Saving payment...':'Save payment'}</button>
            </form>
          </div>
          <HistoryPanel rows={rows} total={data.payments.length} query={query} setQuery={setQuery} onEdit={setEditing} onDelete={deletePayment} deletingId={deletingId} onImage={setImagePreview}/>
        </section>
      </>:<CalendarView payments={data.payments} years={years} selectedYear={selectedYear} setSelectedYear={year=>{setSelectedYear(year);setSelectedMonth(`${year}-01`)}} monthStats={monthStats} selectedMonth={selectedMonth} setSelectedMonth={setSelectedMonth} groups={calendarGroups} onImage={setImagePreview}/>} 
    </main>

    {editing&&<EditPaymentModal item={editing} onClose={()=>setEditing(null)} onImage={setImagePreview} onSave={async body=>{const fresh=await mutate(body);setEditing(null);flash('Payment updated');setTimeout(()=>exportExcel(fresh.payments),100)}}/>}
    {imagePreview&&<ImageModal src={imagePreview} onClose={()=>setImagePreview(null)}/>} 
    {notice&&<div className="toast">✓ {notice}</div>}
  </div>;
}

function HistoryPanel({rows,total,query,setQuery,onEdit,onDelete,deletingId,onImage}){
  return <div className="panel history-panel"><div className="history-head"><div><h2>Payment history</h2><p>{total} total records · newest entry first</p></div><input className="search" placeholder="Search company or card..." value={query} onChange={e=>setQuery(e.target.value)}/></div>
    <div className="table-wrap"><table><thead><tr><th>Company</th><th>Method</th><th>Our card</th><th>Client card</th><th>Receipt</th><th>Date</th><th>Amount</th><th>Actions</th></tr></thead><tbody>{rows.length?rows.map(item=><tr key={item.id}><td><strong>{item.company_name}</strong></td><td>{methodLabel(item.our_payment_method)}</td><td>{item.our_payment_method==='card'?(item.our_account||'—'):'—'}</td><td>{item.our_payment_method==='card'?(item.company_account||'—'):'—'}</td><td>{item.receipt_url?<img className="receipt-thumb" src={item.receipt_url} alt="Receipt" onClick={()=>onImage(item.receipt_url)}/>:<span className="no-receipt">—</span>}</td><td>{dateText(item)}</td><td className="amount-cell">{money(item.amount)}</td><td><div className="row-actions"><button className="edit-btn" onClick={()=>onEdit(item)}>Edit</button><button className="delete-btn" disabled={deletingId===item.id} onClick={()=>onDelete(item)}>{deletingId===item.id?'Deleting...':'Delete'}</button></div></td></tr>):<tr><td colSpan="8" className="empty-row">No payments yet.</td></tr>}</tbody></table></div>
    <div className="mobile-history">{rows.length?rows.map(item=><article className="payment-card-mobile" key={item.id}><div className="payment-card-top"><strong>{item.company_name}</strong><b>{money(item.amount)}</b></div><div className="payment-card-line"><span>Method</span><strong>{methodLabel(item.our_payment_method)}</strong></div>{item.our_payment_method==='card'&&<><div className="payment-card-line"><span>Our card</span><strong>{item.our_account||'—'}</strong></div><div className="payment-card-line"><span>Client card</span><strong>{item.company_account||'—'}</strong></div></>}<div className="payment-card-line"><span>Date</span><strong>{dateText(item)}</strong></div>{item.receipt_url&&<button className="mobile-receipt" onClick={()=>onImage(item.receipt_url)}><img src={item.receipt_url} alt="Receipt"/><span>View receipt</span></button>}<div className="mobile-card-actions"><button className="edit-btn" onClick={()=>onEdit(item)}>Edit</button><button className="delete-btn" disabled={deletingId===item.id} onClick={()=>onDelete(item)}>{deletingId===item.id?'Deleting...':'Delete'}</button></div></article>):<div className="mobile-empty">No payments yet.</div>}</div>
  </div>;
}

function CalendarView({payments,years,selectedYear,setSelectedYear,monthStats,selectedMonth,setSelectedMonth,groups,onImage}){
  const monthIndex=Number(selectedMonth.slice(5,7))-1;
  return <><div className="calendar-heading"><div><h1>Payment Calendar</h1><p>Choose a year, then a month to see daily payment history.</p></div><SelectControl className="year-select"><select value={selectedYear} onChange={e=>setSelectedYear(e.target.value)}>{years.map(y=><option key={y}>{y}</option>)}</select></SelectControl></div>
    <section className="month-grid">{monthStats.map(m=><button key={m.key} className={`month-card ${selectedMonth===m.key?'selected':''}`} onClick={()=>setSelectedMonth(m.key)}><span>{m.name}</span><strong>{money(m.total)}</strong><small>{m.count} payment{m.count===1?'':'s'}</small></button>)}</section>
    <section className="panel calendar-history"><div className="calendar-history-head"><div><h2>{MONTHS[monthIndex]} {selectedYear}</h2><p>{payments.filter(p=>dateText(p).startsWith(selectedMonth)).length} payments in this month</p></div><strong>{money(payments.filter(p=>dateText(p).startsWith(selectedMonth)).reduce((s,p)=>s+Number(p.amount||0),0))}</strong></div>
      <div className="day-groups">{groups.length?groups.map(([day,items])=><div className="day-group" key={day}><div className="day-head"><div><strong>{new Date(`${day}T12:00:00`).toLocaleDateString('en-US',{weekday:'long',month:'short',day:'numeric'})}</strong><span>{items.length} payment{items.length===1?'':'s'}</span></div><b>{money(items.reduce((s,p)=>s+Number(p.amount||0),0))}</b></div><div className="day-payments">{items.map(item=><div className="calendar-payment" key={item.id}><div className="calendar-company"><strong>{item.company_name}</strong><span>{methodLabel(item.our_payment_method)}{item.our_payment_method==='card'&&item.our_account?` · ${item.our_account}`:''}</span></div>{item.receipt_url?<img className="receipt-thumb" src={item.receipt_url} alt="Receipt" onClick={()=>onImage(item.receipt_url)}/>:<span/>}<b>{money(item.amount)}</b></div>)}</div></div>):<div className="calendar-empty">No payments for {MONTHS[monthIndex]} {selectedYear}.</div>}</div>
    </section></>;
}

function EditPaymentModal({item,onClose,onSave,onImage}){
  const [companyName,setCompanyName]=useState(item.company_name||'');const [paymentMethod,setPaymentMethod]=useState(String(item.our_payment_method||'card'));const [ourCard,setOurCard]=useState(item.our_account||'');const [clientCard,setClientCard]=useState(item.company_account||'');const [date,setDate]=useState(dateText(item));const [amount,setAmount]=useState(formatAmount(item.amount));const [receiptData,setReceiptData]=useState(null);const [removeReceipt,setRemoveReceipt]=useState(false);const [saving,setSaving]=useState(false);
  useEffect(()=>{if(paymentMethod==='cash'){setOurCard('');setClientCard('')}},[paymentMethod]);
  async function pick(file){try{setReceiptData(await prepareReceipt(file));setRemoveReceipt(false)}catch(e){alert(e.message)}}
  async function submit(e){e.preventDefault();const numericAmount=Number(amount.replace(/\D/g,''));const cardsOk=paymentMethod==='cash'||(ourCard.trim()&&clientCard.trim());if(!companyName.trim()||!cardsOk||!date||!numericAmount)return;setSaving(true);try{await onSave({action:'payment_update',id:item.id,company_name:companyName.trim(),our_payment_method:paymentMethod,our_card_account_text:paymentMethod==='card'?ourCard.trim():null,company_card_account_text:paymentMethod==='card'?clientCard.trim():null,payment_date:date,amount:numericAmount,receipt_data_url:receiptData,remove_receipt:removeReceipt})}finally{setSaving(false)}}
  const shownReceipt=receiptData||(!removeReceipt?item.receipt_url:null);
  return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="edit-modal"><div className="modal-head"><div><h2>Edit payment</h2><p>Update details or replace the receipt image.</p></div><button className="modal-close" onClick={onClose}>×</button></div><form className="edit-form" onSubmit={submit}><Field label="Company name"><input value={companyName} onChange={e=>setCompanyName(e.target.value)}/></Field><Field label="Payment method"><SelectControl><select value={paymentMethod} onChange={e=>setPaymentMethod(e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>{paymentMethod==='card'&&<><Field label="Our card"><input value={ourCard} onChange={e=>setOurCard(e.target.value)}/></Field><Field label="Client card"><input value={clientCard} onChange={e=>setClientCard(e.target.value)}/></>}<div className="form-row"><Field label="Date"><input type="date" value={date} onChange={e=>setDate(e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input inputMode="numeric" value={amount} onChange={e=>{const v=e.target.value.replace(/\D/g,'');setAmount(v?Number(v).toLocaleString('en-US'):'')}}/></div></Field></div><div className="receipt-field"><div><strong>Receipt image</strong><span>Optional</span></div><label className="upload-button">{shownReceipt?'Replace':'Add image'}<input type="file" accept="image/jpeg,image/png,image/webp" onChange={e=>pick(e.target.files?.[0])}/></label></div>{shownReceipt&&<div className="receipt-selected"><img src={shownReceipt} alt="Receipt" onClick={()=>onImage(shownReceipt)}/><div><strong>Receipt attached</strong><button type="button" onClick={()=>{setReceiptData(null);setRemoveReceipt(true)}}>Remove</button></div></div>}<div className="modal-actions"><button type="button" className="cancel-btn" onClick={onClose}>Cancel</button><button className="save-edit-btn" disabled={saving}>{saving?'Saving...':'Save changes'}</button></div></form></div></div>;
}

function ImageModal({src,onClose}){return <div className="image-modal" onClick={onClose}><button onClick={onClose}>×</button><img src={src} alt="Payment receipt" onClick={e=>e.stopPropagation()}/></div>}
function Stat({label,value}){return <article className="stat-card"><span>{label}</span><strong>{value}</strong></article>}
function Field({label,children}){return <label className="field"><span>{label}</span>{children}</label>}
function SelectControl({children,className=''}){return <div className={`select-control ${className}`}>{children}<span className="select-arrow">⌄</span></div>}
