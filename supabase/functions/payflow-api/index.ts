import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import postgres from "npm:postgres@3.4.7";

const sql = postgres(Deno.env.get("SUPABASE_DB_URL")!, { prepare: false, max: 1 });
const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const BUCKET = "payflow-receipts";
const allowedOrigins = new Set([
  "https://next-project-ashy.vercel.app",
  "https://next-project-alikhaydarovs-projects.vercel.app",
  "https://next-project-git-main-alikhaydarovs-projects.vercel.app",
  "https://crm-chayhana.vercel.app",
  "http://localhost:3000",
  "http://localhost:3010",
]);

function cors(req: Request) {
  const origin = req.headers.get("origin") || "";
  const allowed = allowedOrigins.has(origin);
  return {
    "Access-Control-Allow-Origin": allowed ? origin : "https://next-project-ashy.vercel.app",
    "Access-Control-Allow-Headers": "authorization, apikey, content-type",
    "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
    "Content-Type": "application/json",
    "Vary": "Origin",
  };
}
function originAllowed(req: Request) { return allowedOrigins.has(req.headers.get("origin") || ""); }
function storageHeaders(contentType?: string) {
  return { apikey: SERVICE_KEY, Authorization: `Bearer ${SERVICE_KEY}`, ...(contentType ? { "Content-Type": contentType } : {}) };
}
function extensionFor(type: string) { if (type === "image/png") return "png"; if (type === "image/webp") return "webp"; return "jpg"; }
function decodeBase64(value: string) { const binary=atob(value); const bytes=new Uint8Array(binary.length); for(let i=0;i<binary.length;i++)bytes[i]=binary.charCodeAt(i); return bytes; }

async function uploadReceipt(dataUrl: string | null | undefined, paymentId: string) {
  if (!dataUrl) throw new Error("Image is required");
  const match = dataUrl.match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/);
  if (!match) throw new Error("Receipt must be JPG, PNG, or WEBP");
  const type=match[1], bytes=decodeBase64(match[2]);
  if (bytes.byteLength > 5*1024*1024) throw new Error("Each image must be under 5 MB");
  const path=`${paymentId}/${crypto.randomUUID()}.${extensionFor(type)}`;
  const response=await fetch(`${SUPABASE_URL}/storage/v1/object/${BUCKET}/${path}`,{method:"POST",headers:{...storageHeaders(type),"x-upsert":"false"},body:bytes});
  if(!response.ok)throw new Error("Image upload failed");
  return path;
}
async function deleteReceipt(path?: string|null) {
  if(!path)return;
  await fetch(`${SUPABASE_URL}/storage/v1/object/${BUCKET}/${path}`,{method:"DELETE",headers:storageHeaders("application/json")}).catch(()=>null);
}
async function signedReceiptUrl(path?: string|null) {
  if(!path)return null;
  const response=await fetch(`${SUPABASE_URL}/storage/v1/object/sign/${BUCKET}/${path}`,{method:"POST",headers:storageHeaders("application/json"),body:JSON.stringify({expiresIn:3600})});
  if(!response.ok)return null;
  const result=await response.json(); const signed=result.signedURL||result.signedUrl;
  return signed?`${SUPABASE_URL}/storage/v1${signed}`:null;
}
async function signedReceipts(paths: string[]) {
  return (await Promise.all(paths.map(async path=>({path,url:await signedReceiptUrl(path)})))).filter(x=>x.url);
}

async function snapshot() {
  const payments=await sql`
    select p.id,p.company_id,p.our_payment_method,p.company_payment_method,
           p.our_card_account_text,p.company_card_account_text,p.description,
           p.receipt_path,p.receipt_paths,
           c.name as company_name,
           coalesce(p.our_card_account_text,oa.label) as our_account,
           coalesce(p.company_card_account_text,ca.label) as company_account,
           to_char(p.paid_at at time zone 'Asia/Seoul','YYYY-MM-DD') as payment_date,
           p.paid_at,p.created_at,p.amount::float8 as amount
    from payflow.payments p
    join payflow.companies c on c.id=p.company_id
    left join payflow.accounts oa on oa.id=p.our_account_id
    left join payflow.accounts ca on ca.id=p.company_account_id
    order by p.created_at desc`;
  const enriched=await Promise.all(payments.map(async (p:any)=>{
    const paths=(Array.isArray(p.receipt_paths)&&p.receipt_paths.length?p.receipt_paths:(p.receipt_path?[p.receipt_path]:[])).filter(Boolean);
    const receipts=await signedReceipts(paths);
    return {...p,receipt_paths:paths,receipt_urls:receipts,receipt_url:receipts[0]?.url||null};
  }));
  return {mode:"database",companies:[],our_accounts:[],company_accounts:[],payments:enriched};
}

async function resolveCompanyId(body:any) {
  const companyName=String(body.company_name||"").trim();
  if(companyName){
    const [existing]=await sql`select id from payflow.companies where lower(name)=lower(${companyName}) order by created_at asc limit 1`;
    if(existing)return existing.id;
    const [created]=await sql`insert into payflow.companies(name) values(${companyName}) on conflict(name) do update set name=excluded.name returning id`;
    return created.id;
  }
  if(body.company_id){const [existing]=await sql`select id from payflow.companies where id=${body.company_id}::uuid`;if(existing)return existing.id;}
  throw new Error("Company name is required");
}
function validatePayment(body:any){
  const amount=Number(body.amount),paymentDate=String(body.payment_date||"").trim(),ourMethod=String(body.our_payment_method||"").toLowerCase();
  let ourCardText=String(body.our_card_account_text||"").trim()||null,clientCardText=String(body.company_card_account_text||"").trim()||null;
  const description=String(body.description||"").trim()||null;
  if(!paymentDate||!(amount>0))throw new Error("Complete all required fields");
  if(!/^\d{4}-\d{2}-\d{2}$/.test(paymentDate))throw new Error("Invalid payment date");
  if(!['card','cash'].includes(ourMethod))throw new Error("Invalid payment method");
  if(ourMethod==='cash'){ourCardText=null;clientCardText=null}else{if(!ourCardText)throw new Error("Enter our card");if(!clientCardText)throw new Error("Enter client card")}
  return {amount,paymentDate,ourMethod,ourCardText,clientCardText,description};
}

Deno.serve(async(req:Request)=>{
  const headers=cors(req);
  if(req.method==="OPTIONS"){if(!originAllowed(req))return new Response("Forbidden",{status:403,headers});return new Response("ok",{headers});}
  if(!originAllowed(req))return new Response(JSON.stringify({error:"Origin not allowed"}),{status:403,headers});
  try{
    if(req.method==="GET")return new Response(JSON.stringify(await snapshot()),{headers});
    if(req.method!=="POST")return new Response(JSON.stringify({error:"Method not allowed"}),{status:405,headers});
    const body=await req.json();
    let createdPaymentId:string|null=null;

    if(body.action==="payment"){
      const companyId=await resolveCompanyId(body),v=validatePayment(body),paymentId=crypto.randomUUID();createdPaymentId=paymentId;
      await sql`insert into payflow.payments(id,company_id,our_account_id,company_account_id,our_payment_method,company_payment_method,our_card_account_text,company_card_account_text,description,receipt_path,receipt_paths,amount,paid_at)
        values(${paymentId}::uuid,${companyId}::uuid,null,null,${v.ourMethod},${v.ourMethod==='card'?'card':null},${v.ourCardText},${v.clientCardText},${v.description},null,'{}'::text[],${v.amount},((${v.paymentDate}::date+time '12:00') at time zone 'Asia/Seoul'))`;
      if(body.receipt_data_url){
        const path=await uploadReceipt(body.receipt_data_url,paymentId);
        await sql`update payflow.payments set receipt_paths=array_append(receipt_paths,${path}),receipt_path=${path} where id=${paymentId}::uuid`;
      }
    } else if(body.action==="receipt_add"){
      if(!body.id)throw new Error("Payment id is required");
      const [payment]=await sql`select id from payflow.payments where id=${body.id}::uuid`;if(!payment)throw new Error("Payment not found");
      const path=await uploadReceipt(body.receipt_data_url,body.id);
      try{await sql`update payflow.payments set receipt_paths=array_append(receipt_paths,${path}),receipt_path=coalesce(receipt_path,${path}),updated_at=now() where id=${body.id}::uuid`;}catch(e){await deleteReceipt(path);throw e;}
    } else if(body.action==="receipt_remove"){
      if(!body.id||!body.path)throw new Error("Payment id and image path are required");
      const [payment]=await sql`select receipt_paths from payflow.payments where id=${body.id}::uuid`;if(!payment)throw new Error("Payment not found");
      if(!Array.isArray(payment.receipt_paths)||!payment.receipt_paths.includes(body.path))throw new Error("Image not found");
      await sql`update payflow.payments set receipt_paths=array_remove(receipt_paths,${body.path}),updated_at=now() where id=${body.id}::uuid`;
      await sql`update payflow.payments set receipt_path=case when cardinality(receipt_paths)>0 then receipt_paths[1] else null end where id=${body.id}::uuid`;
      await deleteReceipt(body.path);
    } else if(body.action==="payment_update"){
      if(!body.id)throw new Error("Payment id is required");
      const companyId=await resolveCompanyId(body),v=validatePayment(body);
      const updated=await sql`update payflow.payments set company_id=${companyId}::uuid,our_account_id=null,company_account_id=null,our_payment_method=${v.ourMethod},company_payment_method=${v.ourMethod==='card'?'card':null},our_card_account_text=${v.ourCardText},company_card_account_text=${v.clientCardText},description=${v.description},amount=${v.amount},paid_at=((${v.paymentDate}::date+time '12:00') at time zone 'Asia/Seoul'),updated_at=now() where id=${body.id}::uuid returning id`;
      if(!updated.length)throw new Error("Payment not found");
      if(body.remove_receipt){
        const [cur]=await sql`select receipt_paths,receipt_path from payflow.payments where id=${body.id}::uuid`;
        const paths=(cur?.receipt_paths?.length?cur.receipt_paths:(cur?.receipt_path?[cur.receipt_path]:[]));
        await sql`update payflow.payments set receipt_paths='{}'::text[],receipt_path=null where id=${body.id}::uuid`;
        await Promise.all(paths.map((p:string)=>deleteReceipt(p)));
      }
      if(body.receipt_data_url){
        const path=await uploadReceipt(body.receipt_data_url,body.id);
        await sql`update payflow.payments set receipt_paths=array_append(receipt_paths,${path}),receipt_path=coalesce(receipt_path,${path}) where id=${body.id}::uuid`;
      }
    } else if(body.action==="payment_delete"){
      if(!body.id)throw new Error("Payment id is required");
      const [current]=await sql`select receipt_paths,receipt_path from payflow.payments where id=${body.id}::uuid`;
      const deleted=await sql`delete from payflow.payments where id=${body.id}::uuid returning id`;if(!deleted.length)throw new Error("Payment not found");
      const paths=(current?.receipt_paths?.length?current.receipt_paths:(current?.receipt_path?[current.receipt_path]:[]));
      await Promise.all(paths.map((p:string)=>deleteReceipt(p)));
    } else throw new Error("Unknown action");

    const snap=await snapshot();
    return new Response(JSON.stringify(createdPaymentId?{...snap,created_payment_id:createdPaymentId}:snap),{headers});
  }catch(e){console.error(e);return new Response(JSON.stringify({error:e instanceof Error?e.message:"Server error"}),{status:400,headers});}
});
