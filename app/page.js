'use client';

import {useEffect,useMemo,useState} from 'react';
import * as XLSX from 'xlsx';

const API='https://ehfyhvfmdtbjipgqpvoq.supabase.co/functions/v1/payflow-api';
const empty={mode:'database',companies:[],our_accounts:[],company_accounts:[],payments:[]};
const money=n=>'₩'+Number(n||0).toLocaleString('en-US');
const localNow=()=>{const d=new Date();d.setMinutes(d.getMinutes()-d.getTimezoneOffset());return d.toISOString().slice(0,16)};

export default function Home(){
  const [data,setData]=useState(empty);
  const [tab,setTab]=useState('dashboard');
  const [company,setCompany]=useState('');
  const [ourAccount,setOurAccount]=useState('');
  const [companyAccount,setCompanyAccount]=useState('');
  const [paidAt,setPaidAt]=useState(localNow());
  const [amount,setAmount]=useState('');
  const [query,setQuery]=useState('');
  const [loading,setLoading]=useState(true);
  const [saving,setSaving]=useState(false);
  const [notice,setNotice]=useState('');

  const flash=(text)=>{setNotice(text);setTimeout(()=>setNotice(''),2200)};

  async function request(method='GET',body){
    const response=await fetch(API,{
      method,
      headers:{'Content-Type':'application/json'},
      body:body?JSON.stringify(body):undefined,
      cache:'no-store'
    });
    const result=await response.json().catch(()=>({}));
    if(!response.ok)throw new Error(result.error||'Database request failed');
    return result;
  }

  async function load(){
    try{setData(await request())}
    catch(e){flash(e.message)}
    finally{setLoading(false)}
  }

  useEffect(()=>{load()},[]);
  useEffect(()=>setCompanyAccount(''),[company]);

  const availableCompanyAccounts=data.company_accounts.filter(x=>String(x.company_id)===String(company));

  const stats=useMemo(()=>{
    const today=new Date();
    let todayTotal=0,monthTotal=0,total=0;
    data.payments.forEach(item=>{
      const value=Number(item.amount||0);
      const date=new Date(item.paid_at);
      total+=value;
      if(date.toDateString()===today.toDateString())todayTotal+=value;
      if(date.getMonth()===today.getMonth()&&date.getFullYear()===today.getFullYear())monthTotal+=value;
    });
    return {todayTotal,monthTotal,total,count:data.payments.length};
  },[data.payments]);

  const rows=data.payments.filter(item=>!query||Object.values(item).join(' ').toLowerCase().includes(query.toLowerCase()));

  async function mutate(body){
    const result=await request('POST',body);
    setData(result);
    return result;
  }

  async function addItem(type,value,companyId){
    if(!value.trim())return;
    try{
      if(type==='company')await mutate({action:'company',name:value.trim()});
      if(type==='our')await mutate({action:'our',label:value.trim()});
      if(type==='their')await mutate({action:'their',company_id:companyId,label:value.trim()});
      flash('Saved');
    }catch(e){flash(e.message)}
  }

  function exportExcel(payments=data.payments){
    const rows=payments.map((p,index)=>({
      'ID':index+1,
      'Company Name':p.company_name,
      'Our Card Account':p.our_account,
      'Company Card Account':p.company_account,
      'Date & Time':new Date(p.paid_at).toLocaleString('en-GB'),
      'Amount':Number(p.amount)
    }));
    const workbook=XLSX.utils.book_new();
    const paymentsSheet=XLSX.utils.json_to_sheet(rows);
    paymentsSheet['!cols']=[8,24,24,26,22,16].map(wch=>({wch}));
    XLSX.utils.book_append_sheet(workbook,paymentsSheet,'Payments');
    const summary=XLSX.utils.aoa_to_sheet([
      ['PayFlow Summary',''],
      ['Total Payments',stats.count],
      ['All-Time Amount',stats.total],
      ['This Month',stats.monthTotal],
      ['Today',stats.todayTotal]
    ]);
    XLSX.utils.book_append_sheet(workbook,summary,'Summary');
    XLSX.writeFile(workbook,'payments.xlsx');
  }

  async function submitPayment(e){
    e.preventDefault();
    const numericAmount=Number(amount.replace(/\D/g,''));
    if(!company||!ourAccount||!companyAccount||!paidAt||!numericAmount){flash('Please complete all fields');return}
    setSaving(true);
    try{
      const fresh=await mutate({
        action:'payment',
        company_id:company,
        our_account_id:ourAccount,
        company_account_id:companyAccount,
        paid_at:new Date(paidAt).toISOString(),
        amount:numericAmount
      });
      setAmount('');
      setPaidAt(localNow());
      flash('Payment saved');
      setTimeout(()=>exportExcel(fresh.payments),100);
    }catch(e){flash(e.message)}
    finally{setSaving(false)}
  }

  if(loading)return <div className="loading-screen"><div className="loader"/><strong>Loading PayFlow...</strong></div>;

  return <div className="app-shell">
    <header className="topbar">
      <div className="logo-wrap">
        <div className="logo">P</div>
        <div><strong>PayFlow</strong><span>Payment management</span></div>
      </div>
      <nav className="tabs">
        <button className={tab==='dashboard'?'active':''} onClick={()=>setTab('dashboard')}>Dashboard</button>
        <button className={tab==='settings'?'active':''} onClick={()=>setTab('settings')}>Settings</button>
      </nav>
      <div className="top-actions">
        <span className="connected"><i/>Supabase connected</span>
        <button className="secondary" onClick={()=>exportExcel()}>Download Excel</button>
      </div>
    </header>

    <main className="page">
      {tab==='dashboard'?<Dashboard
        data={data}
        stats={stats}
        company={company}
        setCompany={setCompany}
        ourAccount={ourAccount}
        setOurAccount={setOurAccount}
        companyAccount={companyAccount}
        setCompanyAccount={setCompanyAccount}
        availableCompanyAccounts={availableCompanyAccounts}
        paidAt={paidAt}
        setPaidAt={setPaidAt}
        amount={amount}
        setAmount={setAmount}
        saving={saving}
        submitPayment={submitPayment}
        query={query}
        setQuery={setQuery}
        rows={rows}
      />:<Settings data={data} addItem={addItem}/>} 
    </main>

    {notice&&<div className="toast">{notice}</div>}
  </div>;
}

function Dashboard({data,stats,company,setCompany,ourAccount,setOurAccount,companyAccount,setCompanyAccount,availableCompanyAccounts,paidAt,setPaidAt,amount,setAmount,saving,submitPayment,query,setQuery,rows}){
  return <>
    <div className="page-heading">
      <div><h1>Payments</h1><p>Record payments and keep everything synced with Supabase.</p></div>
    </div>

    <section className="stats-grid">
      <Stat label="Today" value={money(stats.todayTotal)}/>
      <Stat label="This month" value={money(stats.monthTotal)}/>
      <Stat label="Total" value={money(stats.total)}/>
      <Stat label="Payments" value={stats.count}/>
    </section>

    <section className="workspace">
      <div className="panel payment-panel">
        <div className="panel-title"><div><h2>New payment</h2><p>Add a payment in a few seconds.</p></div></div>
        <form className="payment-form" onSubmit={submitPayment}>
          <Field label="Company">
            <select value={company} onChange={e=>setCompany(e.target.value)}>
              <option value="">Select company</option>
              {data.companies.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}
            </select>
          </Field>
          <Field label="Our account">
            <select value={ourAccount} onChange={e=>setOurAccount(e.target.value)}>
              <option value="">Select our account</option>
              {data.our_accounts.map(x=><option key={x.id} value={x.id}>{x.label}</option>)}
            </select>
          </Field>
          <Field label="Company account">
            <select value={companyAccount} onChange={e=>setCompanyAccount(e.target.value)} disabled={!company}>
              <option value="">Select company account</option>
              {availableCompanyAccounts.map(x=><option key={x.id} value={x.id}>{x.label}</option>)}
            </select>
          </Field>
          <div className="form-row">
            <Field label="Date & time">
              <input type="datetime-local" value={paidAt} onChange={e=>setPaidAt(e.target.value)}/>
            </Field>
            <Field label="Amount">
              <div className="amount-input"><span>₩</span><input value={amount} placeholder="1,500,000" onChange={e=>{const value=e.target.value.replace(/\D/g,'');setAmount(value?Number(value).toLocaleString('en-US'):'')}}/></div>
            </Field>
          </div>
          <button className="save-button" disabled={saving}>{saving?'Saving...':'Save payment'}</button>
          <p className="form-note">Excel is downloaded automatically after saving.</p>
        </form>
      </div>

      <div className="panel history-panel">
        <div className="history-head">
          <div><h2>Payment history</h2><p>{data.payments.length} total records</p></div>
          <input className="search" placeholder="Search payments..." value={query} onChange={e=>setQuery(e.target.value)}/>
        </div>
        <div className="table-wrap">
          <table>
            <thead><tr><th>Company</th><th>Our account</th><th>Company account</th><th>Date</th><th>Amount</th></tr></thead>
            <tbody>{rows.length?rows.map(item=><tr key={item.id}>
              <td><strong>{item.company_name}</strong></td>
              <td>{item.our_account}</td>
              <td>{item.company_account}</td>
              <td>{new Date(item.paid_at).toLocaleString('en-GB')}</td>
              <td className="amount-cell">{money(item.amount)}</td>
            </tr>):<tr><td colSpan="5" className="empty-row">No payments yet.</td></tr>}</tbody>
          </table>
        </div>
      </div>
    </section>
  </>;
}

function Settings({data,addItem}){
  const [company,setCompany]=useState('');
  const [ourAccount,setOurAccount]=useState('');
  const [selectedCompany,setSelectedCompany]=useState('');
  const [companyAccount,setCompanyAccount]=useState('');

  return <>
    <div className="page-heading"><div><h1>Settings</h1><p>Manage companies and bank accounts.</p></div></div>
    <section className="settings-grid">
      <SettingCard title="Companies" count={data.companies.length}>
        <form className="quick-add" onSubmit={e=>{e.preventDefault();addItem('company',company);setCompany('')}}>
          <input placeholder="Company name" value={company} onChange={e=>setCompany(e.target.value)}/><button>Add</button>
        </form>
        <SimpleList items={data.companies.map(x=>({id:x.id,title:x.name,meta:'Company'}))}/>
      </SettingCard>

      <SettingCard title="Our accounts" count={data.our_accounts.length}>
        <form className="quick-add" onSubmit={e=>{e.preventDefault();addItem('our',ourAccount);setOurAccount('')}}>
          <input placeholder="Shinhan •••• 1234" value={ourAccount} onChange={e=>setOurAccount(e.target.value)}/><button>Add</button>
        </form>
        <SimpleList items={data.our_accounts.map(x=>({id:x.id,title:x.label,meta:'Our account'}))}/>
      </SettingCard>

      <SettingCard title="Company accounts" count={data.company_accounts.length}>
        <select className="wide-select" value={selectedCompany} onChange={e=>setSelectedCompany(e.target.value)}>
          <option value="">Select company</option>
          {data.companies.map(x=><option key={x.id} value={x.id}>{x.name}</option>)}
        </select>
        <form className="quick-add" onSubmit={e=>{e.preventDefault();if(!selectedCompany)return;addItem('their',companyAccount,selectedCompany);setCompanyAccount('')}}>
          <input placeholder="KB •••• 8821" value={companyAccount} onChange={e=>setCompanyAccount(e.target.value)}/><button>Add</button>
        </form>
        <SimpleList items={data.company_accounts.map(x=>({id:x.id,title:x.label,meta:data.companies.find(c=>String(c.id)===String(x.company_id))?.name||'Company'}))}/>
      </SettingCard>
    </section>
  </>;
}

function Stat({label,value}){return <article className="stat-card"><span>{label}</span><strong>{value}</strong></article>}
function Field({label,children}){return <label className="field"><span>{label}</span>{children}</label>}
function SettingCard({title,count,children}){return <div className="panel setting-card"><div className="setting-title"><h2>{title}</h2><span>{count}</span></div>{children}</div>}
function SimpleList({items}){return <div className="simple-list">{items.length?items.map(item=><div className="simple-item" key={item.id}><strong>{item.title}</strong><span>{item.meta}</span></div>):<div className="empty-list">Nothing added yet.</div>}</div>}
