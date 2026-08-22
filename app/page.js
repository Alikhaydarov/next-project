'use client';

import {useEffect,useMemo,useState} from 'react';
import * as XLSX from 'xlsx';

const API='https://ehfyhvfmdtbjipgqpvoq.supabase.co/functions/v1/payflow-api';
const empty={mode:'database',companies:[],our_accounts:[],company_accounts:[],payments:[]};
const money=n=>'₩'+Number(n||0).toLocaleString('en-US');
const localDate=()=>{const d=new Date();d.setMinutes(d.getMinutes()-d.getTimezoneOffset());return d.toISOString().slice(0,10)};
const dateText=item=>item.payment_date||new Date(item.paid_at).toLocaleDateString('en-CA');
const formatAmount=value=>Number(value||0).toLocaleString('en-US');
const methodLabel=value=>value==='cash'?'Cash':value==='card'?'Card':'—';
const accountLabel=(method,account)=>method==='card'?(account||'—'):'—';

export default function Home(){
  const [data,setData]=useState(empty);
  const [companyName,setCompanyName]=useState('');
  const [paymentMethod,setPaymentMethod]=useState('card');
  const [ourCard,setOurCard]=useState('');
  const [clientCard,setClientCard]=useState('');
  const [paidAt,setPaidAt]=useState(localDate());
  const [amount,setAmount]=useState('');
  const [query,setQuery]=useState('');
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [notice,setNotice]=useState('');
  const [editing,setEditing]=useState(null);
  const [deletingId,setDeletingId]=useState('');

  const flash=text=>{setNotice(text);setTimeout(()=>setNotice(''),2200)};

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
  async function mutate(body){const result=await request('POST',body);setData(result);return result}

  function exportExcel(payments=data.payments){
    const sorted=[...payments].sort((a,b)=>new Date(b.created_at||b.paid_at)-new Date(a.created_at||a.paid_at));
    const excelRows=sorted.map((p,index)=>({
      'No.':index+1,
      'Date':dateText(p),
      'Company Name':p.company_name,
      'Payment Method':methodLabel(p.our_payment_method),
      'Our Card':accountLabel(p.our_payment_method,p.our_account),
      'Client Card':p.our_payment_method==='card'?(p.company_account||'—'):'—',
      'Amount':Number(p.amount)
    }));

    const workbook=XLSX.utils.book_new();
    const paymentsSheet=XLSX.utils.json_to_sheet(excelRows);
    const lastRow=Math.max(excelRows.length+1,2);
    paymentsSheet['!cols']=[{wch:7},{wch:13},{wch:24},{wch:18},{wch:28},{wch:28},{wch:16}];
    paymentsSheet['!autofilter']={ref:`A1:G${lastRow}`};
    paymentsSheet['!rows']=[{hpt:24}];
    for(let row=2;row<=lastRow;row++)if(paymentsSheet[`G${row}`])paymentsSheet[`G${row}`].z='₩#,##0';
    XLSX.utils.book_append_sheet(workbook,paymentsSheet,'Payments');

    const companyMap=new Map();
    sorted.forEach(p=>{const current=companyMap.get(p.company_name)||{count:0,total:0,lastDate:''};current.count+=1;current.total+=Number(p.amount||0);if(!current.lastDate||dateText(p)>current.lastDate)current.lastDate=dateText(p);companyMap.set(p.company_name,current)});
    const companyRows=[...companyMap.entries()].map(([name,v])=>({'Company':name,'Payments':v.count,'Total Amount':v.total,'Latest Payment':v.lastDate})).sort((a,b)=>b['Total Amount']-a['Total Amount']);
    const companySheet=XLSX.utils.json_to_sheet(companyRows);
    companySheet['!cols']=[{wch:28},{wch:12},{wch:18},{wch:16}];
    companySheet['!autofilter']={ref:`A1:D${Math.max(companyRows.length+1,2)}`};
    for(let row=2;row<=companyRows.length+1;row++)if(companySheet[`C${row}`])companySheet[`C${row}`].z='₩#,##0';
    XLSX.utils.book_append_sheet(workbook,companySheet,'Company Summary');

    const monthMap=new Map();
    sorted.forEach(p=>{const month=dateText(p).slice(0,7),current=monthMap.get(month)||{count:0,total:0};current.count+=1;current.total+=Number(p.amount||0);monthMap.set(month,current)});
    const monthRows=[...monthMap.entries()].map(([month,v])=>({'Month':month,'Payments':v.count,'Total Amount':v.total})).sort((a,b)=>b.Month.localeCompare(a.Month));
    const monthSheet=XLSX.utils.json_to_sheet(monthRows);
    monthSheet['!cols']=[{wch:14},{wch:12},{wch:18}];
    monthSheet['!autofilter']={ref:`A1:C${Math.max(monthRows.length+1,2)}`};
    for(let row=2;row<=monthRows.length+1;row++)if(monthSheet[`C${row}`])monthSheet[`C${row}`].z='₩#,##0';
    XLSX.utils.book_append_sheet(workbook,monthSheet,'Monthly Summary');

    const summary=XLSX.utils.aoa_to_sheet([
      ['PAYFLOW REPORT',''],['Generated',new Date().toLocaleString('en-GB')],['Total Payments',sorted.length],['All-Time Amount',sorted.reduce((sum,p)=>sum+Number(p.amount||0),0)],['This Month',stats.monthTotal],['Today',stats.todayTotal]
    ]);
    summary['!cols']=[{wch:22},{wch:70}];
    ['B4','B5','B6'].forEach(cell=>{if(summary[cell])summary[cell].z='₩#,##0'});
    XLSX.utils.book_append_sheet(workbook,summary,'Summary');
    XLSX.writeFile(workbook,'payflow-payments.xlsx',{compression:true});
  }

  async function submitPayment(e){
    e.preventDefault();
    const numericAmount=Number(amount.replace(/\D/g,''));
    const cardsOk=paymentMethod==='cash'||(ourCard.trim()&&clientCard.trim());
    if(!companyName.trim()||!cardsOk||!paidAt||!numericAmount){flash('Please complete all required fields');return}
    setSaving(true);
    try{
      const fresh=await mutate({
        action:'payment',
        company_name:companyName.trim(),
        our_payment_method:paymentMethod,
        our_card_account_text:paymentMethod==='card'?ourCard.trim():null,
        company_card_account_text:paymentMethod==='card'?clientCard.trim():null,
        payment_date:paidAt,
        amount:numericAmount
      });
      setCompanyName('');setOurCard('');setClientCard('');setAmount('');setPaidAt(localDate());flash('Payment saved');setTimeout(()=>exportExcel(fresh.payments),100);
    }catch(e){flash(e.message)}finally{setSaving(false)}
  }

  async function deletePayment(item){
    if(!window.confirm(`Delete ${item.company_name} payment ${money(item.amount)}?`))return;
    setDeletingId(item.id);
    try{await mutate({action:'payment_delete',id:item.id});flash('Payment deleted')}catch(e){flash(e.message)}finally{setDeletingId('')}
  }

  if(loading)return <div className="loading-screen"><div className="loader"/><strong>Loading PayFlow...</strong></div>;

  return <div className="app-shell">
    <header className="topbar">
      <div className="logo-wrap"><div className="logo">P</div><div><strong>PayFlow</strong><span>Payment management</span></div></div>
      <nav className="tabs"><button className="active">Dashboard</button></nav>
      <div className="top-actions"><span className="connected"><i/>Supabase connected</span><button className="secondary" onClick={()=>exportExcel()}>Download Excel</button></div>
    </header>

    <main className="page">
      <div className="page-heading"><div><h1>Payments</h1><p>Type the company and card details directly, then save.</p></div></div>
      <section className="stats-grid"><Stat label="Today" value={money(stats.todayTotal)}/><Stat label="This month" value={money(stats.monthTotal)}/><Stat label="All-time total" value={money(stats.total)}/><Stat label="Payments" value={stats.count}/></section>

      <section className="workspace">
        <div className="panel payment-panel">
          <div className="panel-title"><div><h2>New payment</h2><p>When Card is selected, enter our card first and client card below it.</p></div></div>
          <form className="payment-form" onSubmit={submitPayment}>
            <Field label="Company name"><input value={companyName} onChange={e=>setCompanyName(e.target.value)} placeholder="Enter company name" autoComplete="off"/></Field>
            <Field label="Payment method"><SelectControl><select value={paymentMethod} onChange={e=>setPaymentMethod(e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>
            {paymentMethod==='card'&&<>
              <Field label="Our card"><input value={ourCard} onChange={e=>setOurCard(e.target.value)} placeholder="e.g. Shinhan •••• 1234"/></Field>
              <Field label="Client card"><input value={clientCard} onChange={e=>setClientCard(e.target.value)} placeholder="e.g. KB Bank •••• 8821"/></Field>
            </>}
            <div className="form-row"><Field label="Date"><input type="date" value={paidAt} onChange={e=>setPaidAt(e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input inputMode="numeric" value={amount} placeholder="1,500,000" onChange={e=>{const value=e.target.value.replace(/\D/g,'');setAmount(value?Number(value).toLocaleString('en-US'):'')}}/></div></Field></div>
            <button className="save-button" disabled={saving}>{saving?'Saving payment...':'Save payment'}</button>
            <p className="form-note">No company or card setup is required beforehand.</p>
          </form>
        </div>

        <div className="panel history-panel">
          <div className="history-head"><div><h2>Payment history</h2><p>{data.payments.length} total records · newest entry first</p></div><input className="search" placeholder="Search company or card..." value={query} onChange={e=>setQuery(e.target.value)}/></div>
          <div className="table-wrap"><table><thead><tr><th>Company</th><th>Method</th><th>Our card</th><th>Client card</th><th>Date</th><th>Amount</th><th>Actions</th></tr></thead><tbody>{rows.length?rows.map(item=><tr key={item.id}><td><strong>{item.company_name}</strong></td><td>{methodLabel(item.our_payment_method)}</td><td>{accountLabel(item.our_payment_method,item.our_account)}</td><td>{item.our_payment_method==='card'?(item.company_account||'—'):'—'}</td><td>{dateText(item)}</td><td className="amount-cell">{money(item.amount)}</td><td><div className="row-actions"><button className="edit-btn" onClick={()=>setEditing(item)}>Edit</button><button className="delete-btn" disabled={deletingId===item.id} onClick={()=>deletePayment(item)}>{deletingId===item.id?'Deleting...':'Delete'}</button></div></td></tr>):<tr><td colSpan="7" className="empty-row">No payments yet.</td></tr>}</tbody></table></div>
          <div className="mobile-history">{rows.length?rows.map(item=><article className="payment-card-mobile" key={item.id}><div className="payment-card-top"><strong>{item.company_name}</strong><b>{money(item.amount)}</b></div><div className="payment-card-line"><span>Method</span><strong>{methodLabel(item.our_payment_method)}</strong></div>{item.our_payment_method==='card'&&<><div className="payment-card-line"><span>Our card</span><strong>{item.our_account||'—'}</strong></div><div className="payment-card-line"><span>Client card</span><strong>{item.company_account||'—'}</strong></div></>}<div className="payment-card-line"><span>Date</span><strong>{dateText(item)}</strong></div><div className="mobile-card-actions"><button className="edit-btn" onClick={()=>setEditing(item)}>Edit</button><button className="delete-btn" disabled={deletingId===item.id} onClick={()=>deletePayment(item)}>{deletingId===item.id?'Deleting...':'Delete'}</button></div></article>):<div className="mobile-empty">No payments yet.</div>}</div>
        </div>
      </section>
    </main>

    {editing&&<EditPaymentModal item={editing} onClose={()=>setEditing(null)} onSave={async body=>{const fresh=await mutate(body);setEditing(null);flash('Payment updated');setTimeout(()=>exportExcel(fresh.payments),100)}}/>}
    {notice&&<div className="toast">✓ {notice}</div>}
  </div>;
}

function EditPaymentModal({item,onClose,onSave}){
  const [companyName,setCompanyName]=useState(item.company_name||'');
  const [paymentMethod,setPaymentMethod]=useState(String(item.our_payment_method||'card'));
  const [ourCard,setOurCard]=useState(item.our_account||'');
  const [clientCard,setClientCard]=useState(item.company_account||'');
  const [date,setDate]=useState(dateText(item));
  const [amount,setAmount]=useState(formatAmount(item.amount));
  const [saving,setSaving]=useState(false);

  useEffect(()=>{if(paymentMethod==='cash'){setOurCard('');setClientCard('')}},[paymentMethod]);

  async function submit(e){
    e.preventDefault();
    const numericAmount=Number(amount.replace(/\D/g,''));
    const cardsOk=paymentMethod==='cash'||(ourCard.trim()&&clientCard.trim());
    if(!companyName.trim()||!cardsOk||!date||!numericAmount)return;
    setSaving(true);
    try{await onSave({action:'payment_update',id:item.id,company_name:companyName.trim(),our_payment_method:paymentMethod,our_card_account_text:paymentMethod==='card'?ourCard.trim():null,company_card_account_text:paymentMethod==='card'?clientCard.trim():null,payment_date:date,amount:numericAmount})}finally{setSaving(false)}
  }

  return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="edit-modal"><div className="modal-head"><div><h2>Edit payment</h2><p>Change only what you need.</p></div><button className="modal-close" onClick={onClose}>×</button></div><form className="edit-form" onSubmit={submit}>
    <Field label="Company name"><input value={companyName} onChange={e=>setCompanyName(e.target.value)}/></Field>
    <Field label="Payment method"><SelectControl><select value={paymentMethod} onChange={e=>setPaymentMethod(e.target.value)}><option value="card">Card</option><option value="cash">Cash</option></select></SelectControl></Field>
    {paymentMethod==='card'&&<><Field label="Our card"><input value={ourCard} onChange={e=>setOurCard(e.target.value)}/></Field><Field label="Client card"><input value={clientCard} onChange={e=>setClientCard(e.target.value)}/></Field></>}
    <div className="form-row"><Field label="Date"><input type="date" value={date} onChange={e=>setDate(e.target.value)}/></Field><Field label="Amount"><div className="amount-input"><span>₩</span><input inputMode="numeric" value={amount} onChange={e=>{const value=e.target.value.replace(/\D/g,'');setAmount(value?Number(value).toLocaleString('en-US'):'')}}/></div></Field></div>
    <div className="modal-actions"><button type="button" className="cancel-btn" onClick={onClose}>Cancel</button><button className="save-edit-btn" disabled={saving}>{saving?'Saving...':'Save changes'}</button></div>
  </form></div></div>;
}

function Stat({label,value}){return <article className="stat-card"><span>{label}</span><strong>{value}</strong></article>}
function Field({label,children}){return <label className="field"><span>{label}</span>{children}</label>}
function SelectControl({children,className=''}){return <div className={`select-control ${className}`}>{children}<span className="select-arrow">⌄</span></div>}
