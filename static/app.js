const PAYMENT_THRESHOLD=50000;
const RULE_COUNT=9;
let state={records:[],filter:'all',admin:null,permissions:[],stats:{},imports:[]};
let loading=false;

const $=(selector,root=document)=>root.querySelector(selector);
const $$=(selector,root=document)=>Array.from(root.querySelectorAll(selector));
const money=value=>'¥'+Number(value||0).toLocaleString('zh-CN',{minimumFractionDigits:2,maximumFractionDigits:2});
const number=value=>Number(value||0).toLocaleString('zh-CN',{maximumFractionDigits:2});
const esc=value=>String(value??'').replace(/[&<>"']/g,character=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[character]));
const escAttr=value=>esc(value).replace(/\r?\n/g,' ');
// JSON plus HTML escaping keeps inline action arguments safe for imported identifiers.
const idArg=value=>escAttr(JSON.stringify(String(value??'')));
const domId=(prefix,value)=>`${prefix}-${String(value??'').replace(/[^a-zA-Z0-9_-]/g,'_')}`;
const names={overview:'数据总览',ingest:'单据录入与导入',verify:'核验中心',approval:'付款审批',payment:'付款中心',ledger:'合规台账',admin:'后台管理'};

function hasPermission(permission){return state.permissions.includes(permission)}
function ruleCount(record){return Array.isArray(record?.checks)&&record.checks.length?record.checks.length:RULE_COUNT}
function passedCount(record){return Number(record?.passed||0)}
function isRisk(record){return passedCount(record)<ruleCount(record)&&record?.status!=='已付款'}
function isPassed(record){return passedCount(record)>=ruleCount(record)}
function riskBadge(record){
  const total=ruleCount(record),passed=passedCount(record),failed=total-passed;
  if(failed<=0)return{cls:'success',text:'规则全通过'};
  const level=record.risk?.level||'中风险';
  const cls=level==='高风险'?'danger':level==='中风险'?'warning':'neutral';
  return{cls,level,text:`${level} · ${failed} 项异常`};
}
function displayStage(record){
  if(record?.current_stage)return record.current_stage;
  if(record?.status==='已付款')return '已付款';
  if(isRisk(record))return record.status==='审批驳回'?'审批驳回':'风险待驳回';
  if(!record?.verified)return '待核验';
  if(Number(record.amount||0)>=PAYMENT_THRESHOLD&&!record.financial_approved)return '待审批';
  return '待付款';
}
function statusClass(record){
  const stage=displayStage(record);
  if(stage==='已付款'||stage==='待付款')return 'success';
  if(stage==='待审批')return 'warning';
  if(isRisk(record))return 'danger';
  return 'neutral';
}
function toast(message){const element=$('#toast');if(!element)return;element.textContent=message||'';element.classList.add('show');setTimeout(()=>element.classList.remove('show'),2600)}

function showView(id){
  const requested=names[id]?id:'overview';
  const target=hasPermission(requested)?requested:(hasPermission('overview')?'overview':state.permissions[0]);
  if(!target){toast('当前账号没有可访问的板块');return}
  if(requested!==target)toast('当前账号无权访问此板块');
  $$('.view').forEach(view=>{view.classList.toggle('active',view.id===target);view.hidden=!hasPermission(view.id)});
  $$('nav a').forEach(link=>link.classList.toggle('active',link.dataset.view===target));
  if($('#breadcrumb'))$('#breadcrumb').innerHTML=`<span>绿循智审</span><b>${names[target]||target}</b>`;
  history.replaceState(null,'','#'+target);
}

function applyPermissions(){
  $$('[data-permission]').forEach(element=>{const allowed=hasPermission(element.dataset.permission);element.hidden=!allowed;element.setAttribute('aria-hidden',String(!allowed))});
  $$('.view').forEach(view=>{const allowed=hasPermission(view.id);view.hidden=!allowed;if(!allowed)view.classList.remove('active')});
}

async function readJson(response){
  let data={};
  try{data=await response.json()}catch(error){data={error:'服务器返回了无效响应'}}
  if(response.status===401){location.href='/login';return {ok:false,data}}
  return {ok:response.ok,data}
}
async function loadData(){
  if(loading)return;
  loading=true;
  try{
    const response=await fetch('/api/dashboard',{headers:{Accept:'application/json'}});
    const result=await readJson(response);
    if(!result.ok){if(response.status!==401)toast(result.data.error||'数据加载失败');return}
    state={...state,...result.data,permissions:Array.isArray(result.data.permissions)?result.data.permissions:state.permissions};
    applyPermissions();
    render();
    if(hasPermission('admin'))await loadAdmin();
    const hash=location.hash.slice(1);
    showView(hash&&names[hash]?hash:'overview');
  }catch(error){console.error(error);toast('数据加载失败，请稍后重试')}finally{loading=false}
}
async function loadAdmin(){
  try{const response=await fetch('/api/admin',{headers:{Accept:'application/json'}});const result=await readJson(response);if(result.ok){state.admin=result.data;renderAdmin()}}catch(error){console.error(error)}
}

function render(){
  const stats=state.stats||{};
  const total=Number(stats.total||state.records.length||0);
  $('#stats').innerHTML=[['已导入业务',total],['规则全通过',Number(stats.passed||0)],['风险待处置',Number(stats.risk||0)],['交易总金额',money(stats.amount||0)]].map(item=>`<div class="stat"><div class="stat-top"><span>${item[0]}</span></div><b>${item[1]}</b></div>`).join('');
  if($('#riskNav'))$('#riskNav').textContent=Number(stats.risk||0);
  if($('#approvalNav'))$('#approvalNav').textContent=state.records.filter(record=>isRisk(record)||(displayStage(record)==='待审批')).length;
  if($('#paymentNav'))$('#paymentNav').textContent=state.records.filter(record=>paymentEligible(record)).length;
  renderCharts();renderQueue();renderRecords();renderApproval();renderPayments();renderLedger();renderImports();
}

function renderCharts(){
  const records=state.records||[];
  // 格1：供货方金额 TOP5
  const suppliers={};records.forEach(record=>{const key=String(record.supplier||'未填写').trim()||'未填写';suppliers[key]=(suppliers[key]||0)+Number(record.amount||0)});
  const topSuppliers=Object.entries(suppliers).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const supplierMax=topSuppliers[0]?.[1]||1;
  if($('#supplierChart'))$('#supplierChart').innerHTML=topSuppliers.map(([name,sum])=>`<div class="risk-bar"><div><span title="${escAttr(name)}">${esc(name)}</span><b>${money(sum)}</b></div><i><em style="width:${Math.max(4,sum/supplierMax*100)}%;background:#3d7ea6"></em></i></div>`).join('')||'<div class="empty">暂无供货方数据</div>';
  // 格2：风险等级分布（高/中/低）
  const levelColors={高风险:'#bf4b3a',中风险:'#e0a356',低风险:'#93b3a4'};
  const levels={高风险:0,中风险:0,低风险:0};
  records.forEach(record=>{if(passedCount(record)<ruleCount(record)){const level=record.risk?.level;levels[level&&level in levels?level:'中风险']++}});
  const riskTotal=levels.高风险+levels.中风险+levels.低风险;
  if($('#riskLevelTotal'))$('#riskLevelTotal').textContent=`${riskTotal} 笔`;
  if($('#levelChart'))$('#levelChart').innerHTML=riskTotal?`<div class="level-stack">${['高风险','中风险','低风险'].map(level=>levels[level]?`<i style="width:${levels[level]/riskTotal*100}%;background:${levelColors[level]}"></i>`:'').join('')}</div>${['高风险','中风险','低风险'].map(level=>`<div class="level-row"><i class="level-dot" style="background:${levelColors[level]}"></i><span>${level}</span><small>${riskTotal?Math.round(levels[level]/riskTotal*100):0}%</small><b>${levels[level]} 笔</b></div>`).join('')}`:'<div class="empty">当前没有风险记录</div>';
  // 格3：货物品类金额 TOP5
  const goods={};records.forEach(record=>{const key=String(record.goods||'未填写').trim()||'未填写';goods[key]=(goods[key]||0)+Number(record.amount||0)});
  const topGoods=Object.entries(goods).sort((a,b)=>b[1]-a[1]).slice(0,5);
  const goodsMax=topGoods[0]?.[1]||1;
  if($('#goodsChart'))$('#goodsChart').innerHTML=topGoods.map(([name,sum])=>`<div class="risk-bar"><div><span>${esc(name)}</span><b>${money(sum)}</b></div><i><em style="width:${Math.max(4,sum/goodsMax*100)}%;background:#167a59"></em></i></div>`).join('')||'<div class="empty">暂无品类数据</div>';
  // 格4：规则命中 TOP
  const failures={};records.forEach(record=>(record.checks||[]).filter(check=>!check.passed).forEach(check=>{failures[check.name]=(failures[check.name]||0)+1}));
  const totalFailures=Object.values(failures).reduce((sum,value)=>sum+value,0)||1;
  if($('#ruleChart'))$('#ruleChart').innerHTML=Object.entries(failures).sort((a,b)=>b[1]-a[1]).slice(0,5).map(([name,count])=>`<div class="risk-bar"><div><span>${esc(name)}</span><b>${count} 笔</b></div><i><em style="width:${count/totalFailures*100}%"></em></i></div>`).join('')||'<div class="empty">当前没有规则风险命中</div>';
  if($('#funnel')){
    const stages=[['已导入',records.length],['已核验',records.filter(record=>record.verified).length],['待审批',records.filter(record=>displayStage(record)==='待审批').length],['已付款',records.filter(record=>record.status==='已付款').length]];
    const peak=Math.max(...stages.map(item=>item[1]),1);
    const total=stages[0][1]||1;
    const colors=['#14684a','#1d7a58','#2c8f6a','#45a37c'];
    $('#funnel').innerHTML=`<div class="histo"><div class="histo-bars">${stages.map((item,index)=>`<div class="histo-col" title="${item[0]}：${item[1]} 笔"><span>${item[1]}</span><i style="height:${Math.max(5,item[1]/peak*100)}%;background:linear-gradient(180deg,${colors[index]},${colors[index]}cc)"></i></div>`).join('')}</div><div class="histo-axis">${stages.map((item,index)=>`<span><b>${item[0]}</b><small>${Math.round(item[1]/total*100)}%</small></span>`).join('')}</div></div>`;
  }
}

const QUEUE_PREVIEW_COUNT=5;
let queueExpanded=false;
function toggleQueue(){queueExpanded=!queueExpanded;renderQueue()}
function renderQueue(){
  const rows=state.records.filter(record=>isRisk(record));
  if(!$('#riskQueue'))return;
  const item=record=>`<div class="queue-item"><i class="risk-pip"></i><div class="queue-main"><b>${esc(record.id)} · ${esc(record.supplier)}</b><p>${esc(record.risk?.level||'风险')}（${esc(record.risk?.score||0)} 分）· ${esc(record.goods)} · ${(record.checks||[]).filter(check=>!check.passed).map(check=>esc(check.name)).slice(0,2).join('、')}</p></div><span class="tag ${riskBadge(record).cls}">${esc(displayStage(record))}</span>${hasPermission('approval')?`<button class="text-button" onclick="showView('approval')">处置 →</button>`:hasPermission('verify')?`<button class="text-button" onclick="openDrawer(${idArg(record.id)})">处置 →</button>`:''}</div>`;
  const visible=queueExpanded?rows:rows.slice(0,QUEUE_PREVIEW_COUNT);
  let html=visible.map(item).join('');
  if(rows.length>QUEUE_PREVIEW_COUNT){
    const hidden=rows.length-QUEUE_PREVIEW_COUNT;
    html+=`<button class="queue-toggle" onclick="toggleQueue()">${queueExpanded?`收起列表 ↑`:`展开全部 ${rows.length} 条（还有 ${hidden} 条）↓`}</button>`;
  }
  $('#riskQueue').innerHTML=html||'<div class="queue-item">没有待处理业务</div>';
}

function recordMatchesFilter(record){
  if(state.filter==='risk')return isRisk(record);
  if(state.filter==='passed')return isPassed(record);
  if(state.filter==='pending')return !record.verified&&record.status!=='已付款';
  return true;
}
function renderRecords(){
  if(!$('#records'))return;
  const query=($('#search')?.value||'').trim().toLowerCase();
  const rows=state.records.filter(record=>recordMatchesFilter(record)).filter(record=>[record.id,record.batch,record.supplier].join(' ').toLowerCase().includes(query));
  $('#records').innerHTML=rows.map(record=>{
    const total=ruleCount(record),passed=passedCount(record),badge=riskBadge(record);
    const rejectAction=isRisk(record)&&(hasPermission('verify')||hasPermission('approval'))?`<button class="reject compact" onclick="rejectAll(${idArg(record.id)})">一键驳回</button>`:'';
    const action=passed>=total&&!record.verified&&hasPermission('verify')?`<button class="primary compact" onclick="verifyBiz(${idArg(record.id)})">确认已核验</button>`:`<button class="view-link" onclick="openDrawer(${idArg(record.id)})">查看详情</button>`;
    return `<tr><td><div class="business-cell"><b>${esc(record.id)}</b><small>${esc(record.batch)} · ${esc(record.supplier)} · ${esc(record.supplier_type||'自然人')}</small></div></td><td>${esc(record.goods)}<small class="block">${money(record.amount)}</small></td><td><div class="rule-dots">${(record.checks||[]).map(check=>`<i class="rule-dot ${check.passed?'':'fail'}" title="${escAttr(check.name)}">${check.passed?'✓':'×'}</i>`).join('')}</div></td><td><span class="tag ${badge.cls}" title="${escAttr(record.risk?.summary||'')}">${badge.text}</span></td><td><span class="tag ${statusClass(record)}">${esc(displayStage(record))}</span></td><td>${rejectAction}${action}</td></tr>`;
  }).join('')||'<tr><td colspan="6">暂无匹配记录</td></tr>';
}

function approvalRows(){return state.records.filter(record=>isRisk(record)||(isPassed(record)&&record.verified&&Number(record.amount||0)>=PAYMENT_THRESHOLD&&!record.financial_approved&&record.status!=='已付款'))}
function renderApproval(){
  if(!$('#approvalRows'))return;
  const rows=approvalRows();
  if($('#approvalCount'))$('#approvalCount').textContent=`${rows.length} 笔待处置`;
  $('#approvalRows').innerHTML=rows.map(record=>{
    const risk=isRisk(record), failed=(record.checks||[]).filter(check=>!check.passed), handled=failed.filter(check=>record.risk_decisions?.[check.code]?.decision==='rejected').length;
    const riskInfo=risk?`<div class="risk-reasons">${failed.map(check=>`<span class="risk-chip ${record.risk_decisions?.[check.code]?.decision==='rejected'?'handled':''}">! ${esc(check.code)} ${esc(check.name)} ${record.risk_decisions?.[check.code]?.decision==='rejected'?'· 已驳回':''}</span>`).join('')}</div><p class="approval-progress">已手动驳回 ${handled}/${failed.length} 项风险依据${record.status==='审批驳回'?'，等待补件修改':''}</p>`:'';
    const canHandleRisk=hasPermission('approval')||hasPermission('verify');
    const rejectAction=risk&&canHandleRisk&&record.status!=='审批驳回'?`<button class="reject" onclick="rejectAll(${idArg(record.id)})">一键驳回</button>`:'';
    const individualActions=risk&&canHandleRisk?failed.filter(check=>record.risk_decisions?.[check.code]?.decision!=='rejected').map(check=>`<button class="reject" onclick="rejectRisk(${idArg(record.id)},${idArg(check.code)})">驳回 ${esc(check.code)}</button>`).join(''):'';
    const editAction=record.status==='审批驳回'&&hasPermission('admin')?`<button class="primary" onclick="openEdit(${idArg(record.id)})">补件并再次审核</button>`:'';
    const approveAction=!risk&&hasPermission('approval')&&!record.financial_approved?`<button class="primary" onclick="approveBiz(${idArg(record.id)})">财务审批</button>`:'';
    return `<article class="approval-card"><div class="approval-top"><div><span class="caption">${risk?'RISK CASE':'AMOUNT REVIEW'} · ${esc(record.id)}</span><h2>${esc(record.supplier)} · ${esc(record.goods)}</h2><p>${money(record.amount)} · ${number(record.net)} Kg · ${esc(displayStage(record))}</p></div><span class="tag ${risk?riskBadge(record).cls:'warning'}">${risk?riskBadge(record).text:'待财务判断'}</span></div>${riskInfo}<div class="approval-actions"><button class="secondary" onclick="openDrawer(${idArg(record.id)})">查看核验依据</button>${rejectAction}${individualActions}${editAction}${approveAction}</div></article>`;
  }).join('')||'<div class="empty-panel">当前没有待处置业务。</div>';
}

function paymentEligible(record){return isPassed(record)&&Boolean(record.verified)&&record.status!=='已付款'&&(Number(record.amount||0)<PAYMENT_THRESHOLD||Boolean(record.financial_approved))}
function localMask(value){const text=String(value||'');return text.length>8?`${text.slice(0,4)}****${text.slice(-4)}`:(text||'未填写')}
function renderPayments(){
  if(!$('#paymentRows'))return;
  const rows=state.records.filter(record=>paymentEligible(record));
  if($('#paymentCount'))$('#paymentCount').textContent=`${rows.length} 笔待确认`;
  $('#paymentRows').innerHTML=rows.map(record=>{const masked=record.bank_account_masked||localMask(record.bank_account);return `<div class="payment-row"><div class="grow"><b>${esc(record.id)} · ${esc(record.supplier)}</b><p>${esc(record.goods)} · ${money(record.amount)} · ${esc(displayStage(record))}</p><div class="bank-account"><span id="${domId('bank',record.id)}">${esc(masked)}</span><button type="button" title="查看完整银行卡号" aria-label="查看完整银行卡号" aria-pressed="false" onclick="toggleBank(${idArg(record.id)},this)">◉</button></div></div><button class="primary" onclick="pay(${idArg(record.id)})">确认付款</button></div>`}).join('')||'<div class="payment-row">暂无符合付款条件的业务</div>';
  renderPaymentGuide(rows);
}
function renderPaymentGuide(rows){
  const pendingAmount=rows.reduce((sum,record)=>sum+Number(record.amount||0),0);
  const paid=state.records.filter(record=>record.status==='已付款');
  const paidAmount=paid.reduce((sum,record)=>sum+Number(record.paid_amount||record.amount||0),0);
  if($('#payKpis'))$('#payKpis').innerHTML=[['待确认付款',`${rows.length} 笔`,money(pendingAmount)],['已确认付款',`${paid.length} 笔`,money(paidAmount)],['付款阈值',`${PAYMENT_THRESHOLD/10000} 万`,`小额直付线`]].map(item=>`<div><small>${item[0]}</small><b>${item[1]}</b><span>${item[2]}</span></div>`).join('');
  if($('#paySplit')){
    const smallRows=rows.filter(record=>Number(record.amount||0)<PAYMENT_THRESHOLD);
    const bigRows=rows.filter(record=>Number(record.amount||0)>=PAYMENT_THRESHOLD);
    const bigWaiting=state.records.filter(record=>isPassed(record)&&record.verified&&Number(record.amount||0)>=PAYMENT_THRESHOLD&&!record.financial_approved&&record.status!=='已付款');
    const base=Math.max(pendingAmount,1);
    const bar=(count,sum,label,note)=>`<div class="pay-split-row"><span>${label}</span><i><em style="width:${Math.max(4,sum/base*100)}%"></em></i><b>${count} 笔</b></div><small>${note}</small>`;
    $('#paySplit').innerHTML=`<b>金额自动分流（待确认）</b>${bar(smallRows.length,smallRows.reduce((sum,record)=>sum+Number(record.amount||0),0),'小额直付','金额 < 5 万，核验通过后可直接确认付款')}${bar(bigRows.length,bigRows.reduce((sum,record)=>sum+Number(record.amount||0),0),'大额已审批','金额 ≥ 5 万，已获财务负责人批准')}${bar(bigWaiting.length,bigWaiting.reduce((sum,record)=>sum+Number(record.amount||0),0),'大额待审批','金额 ≥ 5 万，等待财务负责人判断')}`;
  }
  if($('#payRecent')){
    const recent=paid.slice().sort((a,b)=>String(b.payment_time||'').localeCompare(String(a.payment_time||''))).slice(0,3);
    $('#payRecent').innerHTML=`<b>最近确认付款</b>${recent.map(record=>`<div class="pay-recent-item"><b>${esc(record.id)} · ${esc(record.supplier)}</b><span>${money(record.paid_amount||record.amount)} · ${esc(record.payment_time||'时间未记录')}</span></div>`).join('')||'<p class="pay-recent-empty">还没有确认付款记录</p>'}`;
  }
}

function snapshotFields(snapshot){
  if(!snapshot||typeof snapshot!=='object')return '<p class="empty">暂无单据快照</p>';
  const fields=[['业务编号','id'],['批次号','batch'],['供货方','supplier'],['货物','goods'],['毛重','gross'],['皮重','tare'],['净重','net'],['扣杂数量','inspect_qty'],['入库数量','inbound_qty'],['金额','amount'],['过磅时间','weigh_time'],['质检时间','inspect_time'],['入库时间','inbound_time'],['付款时间','payment_time'],['银行卡号','bank_account']];
  return `<dl>${fields.filter(([,key])=>snapshot[key]!==undefined&&snapshot[key]!=='').map(([label,key])=>`<div><dt>${label}</dt><dd>${esc(String(snapshot[key]))}</dd></div>`).join('')}</dl>`;
}
function historyMarkup(record){
  const history=Array.isArray(record.rejection_history)?record.rejection_history:[];
  if(!history.length)return '<p class="empty">该业务没有驳回历史。</p>';
  return history.map(item=>{const original=item.original||item.original_documents;const revised=item.revised||item.revised_documents;return `<div class="history-item"><b>第 ${esc(item.cycle||'?')} 次驳回 · ${esc(item.rejected_code||'ALL')}</b><p>驳回意见：${esc(item.opinion||'未填写')}</p><small>驳回时间：${esc(item.rejected_at||'')}</small><div class="source-grid"><div><small>原始单据</small>${snapshotFields(original)}</div><div><small>修改后的新版单据</small>${revised?snapshotFields(revised):'<p class="empty">尚未补件</p>'}</div></div></div>`}).join('');
}
function attachmentMarkup(record){return `<div class="source-grid"><div><small>过磅单</small><b>${esc(record.id)}</b><p>毛重 ${number(record.gross)} Kg · 皮重 ${number(record.tare)} Kg · 净重 ${number(record.net)} Kg<br>过磅时间：${esc(record.weigh_time||'未填写')}</p></div><div><small>质检单</small><b>${esc(record.inspect_id||'缺失')}</b><p>${esc(record.spec||'未填写')} · ${esc(record.grade||'未填写')} · 扣杂 ${number(record.inspect_qty)} Kg<br>质检时间：${esc(record.inspect_time||'未填写')}</p></div><div><small>入库单</small><b>${esc(record.inbound_id||'缺失')}</b><p>入库 ${number(record.inbound_qty)} Kg<br>入库时间：${esc(record.inbound_time||'未填写')}</p></div><div><small>收款银行卡</small><b>${esc(record.bank_account_masked||localMask(record.bank_account))}</b><p>${esc(record.bank_name||'虚拟收款户名')} · ${esc(record.bank_branch||'虚拟银行支行')}</p></div></div><h4>驳回历史附件</h4>${historyMarkup(record)}`}
function ledgerCompletedTime(record){return String(record.payment_time||record.updated_at||'').replace('T',' ')}
function ledgerAmount(record){return Number(record.paid_amount||record.amount||0)||0}
function sortLedgerRows(rows){
  const sort=$('#ledgerSort')?.value||'completed_desc';
  const sorted=rows.slice();
  if(sort==='completed_asc') sorted.sort((a,b)=>ledgerCompletedTime(a).localeCompare(ledgerCompletedTime(b)));
  else if(sort==='amount_asc') sorted.sort((a,b)=>ledgerAmount(a)-ledgerAmount(b));
  else if(sort==='amount_desc') sorted.sort((a,b)=>ledgerAmount(b)-ledgerAmount(a));
  else sorted.sort((a,b)=>ledgerCompletedTime(b).localeCompare(ledgerCompletedTime(a)));
  return sorted;
}
function renderLedger(){
  if(!$('#ledgerRows'))return;
  const rows=sortLedgerRows(state.records.filter(record=>record.ledger_eligible===true));
  if($('#ledgerNote'))$('#ledgerNote').textContent=`当前已归档 ${rows.length} 笔`;
  $('#ledgerRows').innerHTML=rows.map(record=>{const detailId=domId('ledger-detail',record.id);return `<tr class="ledger-row" data-ledger-id="${escAttr(record.id)}"><td>HG${esc(String(record.id||'').replace(/^GB/,''))}</td><td>${esc(record.id)}</td><td>${esc(record.supplier)}</td><td>${money(record.paid_amount||record.amount)}</td><td>${esc(ledgerCompletedTime(record)||'—')}</td><td><span class="tag success">已付款</span></td><td><button class="view-link" type="button" aria-expanded="false" aria-controls="${detailId}" onclick="toggleLedgerAttachment(${idArg(record.id)},this)">展开附件</button></td></tr><tr id="${detailId}" class="ledger-detail-row" hidden><td colspan="7"><div class="ledger-attachments">${attachmentMarkup(record)}</div></td></tr>`}).join('')||'<tr><td colspan="7">暂无已完整付款的合规业务</td></tr>';
}
function toggleLedgerAttachment(id,button){const detail=$(`#${domId('ledger-detail',id)}`);if(!detail)return;const open=!detail.hidden;detail.hidden=open;button.setAttribute('aria-expanded',String(!open));button.textContent=open?'展开附件':'收起附件'}

function renderImports(){
  if(!$('#importHistory'))return;
  const imports=Array.isArray(state.imports)?state.imports:[];
  $('#importHistory').innerHTML=imports.map(item=>`<div class="import-row"><span class="import-file">XLSX</span><div><b>${esc(item.filename)}</b><p>${esc(String(item.created_at||'').replace('T',' '))} · ${number(item.rows_count)} 笔业务</p></div><span class="tag success">${esc(item.result||'已导入')}</span></div>`).join('')||'<div class="import-row">暂无导入记录</div>';
}
function renderAdmin(){
  const admin=state.admin;if(!admin)return;
  if($('#adminSummary'))$('#adminSummary').innerHTML=[['应用后端',admin.backend||'Flask · Python'],['数据存储',admin.storage||'SQLite'],['规则数量',`${(admin.rules||[]).length} 项固定规则`]].map(item=>`<div class="admin-stat"><small>${item[0]}</small><b>${esc(item[1])}</b><span>运行正常</span></div>`).join('');
  if($('#adminRules'))$('#adminRules').innerHTML=(admin.rules||[]).map(rule=>`<div class="admin-row"><div><b>${esc(rule.code)} · ${esc(rule.name)}</b><p>${esc(rule.owner||'合规风控部')} · 最近更新 ${esc(rule.updated||'')}</p></div><span class="tag success">${rule.enabled===false?'停用':'启用'}</span></div>`).join('');
  if($('#adminAccounts'))$('#adminAccounts').innerHTML=(admin.accounts||[]).map(account=>`<div class="admin-row"><div><b>${esc(account.name)}</b><p>账号 ${esc(account.account)} · ${esc(account.role)}</p><small>${(account.permissions||[]).map(permission=>esc(names[permission]||permission)).join('、')}</small></div><span class="tag success">${esc(account.status||'启用')}</span></div>`).join('');
}

async function requestJson(url,options={}){const response=await fetch(url,options);const result=await readJson(response);return {response,...result}}
async function action(url,method='POST',body){
  const options={method,headers:body?{'Content-Type':'application/json'}:undefined,body:body?JSON.stringify(body):undefined};
  try{const result=await requestJson(url,options);toast(result.data.message||result.data.error|| (result.ok?'操作完成':'操作失败'));if(result.ok)await loadData();return result.ok}catch(error){console.error(error);toast('操作失败，请稍后重试');return false}
}
async function verifyBiz(id){const ok=await action(`/api/business/${encodeURIComponent(id)}/verify`);if(ok)closeDrawer()}
async function approveBiz(id){const ok=await action(`/api/business/${encodeURIComponent(id)}/approve`);if(ok)closeDrawer()}
async function pay(id){const ok=await action(`/api/business/${encodeURIComponent(id)}/pay`);if(ok)closeDrawer()}
async function rejectAll(id){const ok=await action(`/api/business/${encodeURIComponent(id)}/reject`,'POST',{opinion:'风险业务一键驳回，退回补件'});if(ok)closeDrawer()}
async function rejectRisk(id,code){const ok=await action(`/api/business/${encodeURIComponent(id)}/risk/${encodeURIComponent(code)}/reject`,'POST',{opinion:'未通过风险项，退回补件'});if(ok)closeDrawer()}
async function clearBusinessData(){if(!confirm('确定清空全部业务数据吗？清空后仅能通过重新导入 Excel 恢复。'))return;await action('/api/admin/clear-data')}
async function clearImportHistory(){await action('/api/import-history','DELETE')}
async function syncLocalTemplate(){await action('/api/local-sync')}
function toggleBank(id,button){const record=state.records.find(item=>String(item.id)===String(id));if(!record)return;const element=$(`#${domId('bank',id)}`);if(!element)return;const open=button.dataset.open==='1';element.textContent=open?(record.bank_account_masked||localMask(record.bank_account)):(record.bank_account||'未填写');button.dataset.open=open?'0':'1';button.setAttribute('aria-pressed',String(!open));button.title=open?'查看完整银行卡号':'隐藏银行卡号';button.textContent=open?'◉':'◎'}

function openEdit(id){
  if(!hasPermission('admin')){toast('当前账号无权修改单据');return}
  const record=state.records.find(item=>String(item.id)===String(id));if(!record)return;
  const fields=[['supplier','供货方名称'],['supplier_type','供货方类别（自然人/公司）'],['id_no','身份证号'],['registered','个体户登记标记（是/否）'],['plate','车牌号'],['goods','过磅货物品名'],['gross','毛重 (Kg)'],['tare','皮重 (Kg)'],['method','过磅方式'],['weigher','过磅员签字'],['driver','司机签字'],['weigh_date','业务日期'],['weigh_time','过磅时间'],['inspect_id','质检单编号'],['inspect_goods','质检货物品名'],['inspect_date','质检日期'],['inspect_time','质检时间'],['spec','货品规格'],['grade','质检等级'],['inspect_qty','质检扣杂数量 (Kg)'],['inbound_id','入库单编号'],['inbound_goods','入库货物品名'],['inbound_date','入库日期'],['inbound_time','入库时间'],['inbound_qty','入库数量 (Kg)'],['settlement_qty','扣杂后结算数量 (Kg)'],['payment_time','付款时间'],['price','计价单价'],['amount','计价金额'],['inspector','质检员签字'],['reviewer','复核人签字'],['invoice','票据方式（reverse/self/none）'],['frequency','近30日交易笔数'],['annual','年度累计金额'],['bank_name','收款户名'],['bank_account','银行卡号'],['bank_branch','开户行']];
  $('#drawerContent').innerHTML=`<p class="eyebrow dark">SUPPLEMENT · ${esc(record.id)}</p><h2>补件修改后再次审核</h2><p class="drawer-sub">后台管理员可以修正资料并重新执行全部规则；财务负责人只能查看和驳回。</p><form id="editForm" class="edit-form">${fields.map(([key,label])=>`<div class="form-field"><label>${label}</label><input name="${key}" value="${escAttr(record[key])}"></div>`).join('')}<div class="manual-actions"><button class="primary" type="submit">提交补件并重新审核</button></div></form>`;
  $('#drawer').classList.add('open');
  $('#editForm').onsubmit=async event=>{event.preventDefault();const data=Object.fromEntries(new FormData(event.target).entries());const ok=await action(`/api/business/${encodeURIComponent(id)}/resubmit`,'POST',data);if(ok)closeDrawer()};
}

function openDrawer(id){
  const record=state.records.find(item=>String(item.id)===String(id));if(!record)return;
  const total=ruleCount(record),passed=passedCount(record),failed=(record.checks||[]).filter(check=>!check.passed);
  const riskActions=failed.length&&(hasPermission('approval')||hasPermission('verify'))?`<div class="decision-box"><b>风险处置</b><p>风险业务可一键驳回；也可逐项记录驳回意见。全部规则通过后才能进入付款流程。</p><div>${record.status!=='审批驳回'?`<button class="reject" onclick="rejectAll(${idArg(record.id)})">一键驳回</button>`:''}${failed.filter(check=>record.risk_decisions?.[check.code]?.decision!=='rejected').map(check=>`<button class="reject" onclick="rejectRisk(${idArg(record.id)},${idArg(check.code)})">驳回 ${esc(check.code)}</button>`).join('')}${record.status==='审批驳回'&&hasPermission('admin')?`<button class="primary" onclick="openEdit(${idArg(record.id)})">补件并再次审核</button>`:''}</div></div>`:'';
  const verifyAction=passed>=total&&!record.verified&&hasPermission('verify')?`<div class="decision-box"><b>规则全部通过</b><p>确认已核对全部资料后，业务将进入待付款或待审批环节。</p><div><button class="primary" onclick="verifyBiz(${idArg(record.id)})">确认已核验</button></div></div>`:'';
  const approvalAction=passed>=total&&record.verified&&Number(record.amount||0)>=PAYMENT_THRESHOLD&&!record.financial_approved&&hasPermission('approval')?`<div class="decision-box"><b>付款审批</b><p>当前业务等待财务负责人判断。</p><div><button class="primary" onclick="approveBiz(${idArg(record.id)})">财务审批</button></div></div>`:'';
  $('#drawerContent').innerHTML=`<p class="eyebrow dark">BUSINESS EVIDENCE · ${esc(record.id)}</p><h2>${esc(record.supplier)} · ${esc(record.goods)}</h2><p class="drawer-sub">过磅 ${esc(record.weigh_time||record.weigh_date||'未填写')} · 质检 ${esc(record.inspect_time||record.inspect_date||'未填写')} · 入库 ${esc(record.inbound_time||record.inbound_date||'未填写')}</p><div class="detail-summary"><span class="tag ${passed>=total?'success':riskBadge(record).cls}">${passed}/${total} 通过${passed<total?` · ${riskBadge(record).text}`:''}</span><span class="tag ${statusClass(record)}">${esc(displayStage(record))}</span><span class="tag neutral">${esc(record.approval||'无需审批')}</span></div><div class="security-detail"><b>安全指纹已记录</b><span>算法：${esc(record.hash_algorithm||'SHA-256')}</span><small>整单指纹：${esc(record.integrity_hash||'待生成')}</small></div>${attachmentMarkup(record)}<h3>规则执行证据</h3>${(record.checks||[]).map(check=>`<div class="rule-detail"><div class="rule-detail-head"><i class="rule-dot ${check.passed?'':'fail'}">${check.passed?'✓':'×'}</i><b>${esc(check.code)} ${esc(check.name)}</b><span>${check.passed?'通过':record.risk_decisions?.[check.code]?.decision==='rejected'?'已驳回':'待处置'}</span></div><p>${esc(check.reason)}</p></div>`).join('')}${riskActions}${approvalAction}${verifyAction}`;
  $('#drawer').classList.add('open');
}
function closeDrawer(){$('#drawer')?.classList.remove('open')}

document.querySelectorAll('nav a').forEach(link=>link.addEventListener('click',event=>{event.preventDefault();showView(link.dataset.view)}));
document.querySelectorAll('.filter-row button.filter').forEach(button=>button.addEventListener('click',()=>{$$('.filter-row button.filter').forEach(item=>item.classList.remove('active'));button.classList.add('active');state.filter=button.dataset.filter;renderRecords()}));
$('#search')?.addEventListener('input',renderRecords);
$('#ledgerSort')?.addEventListener('change',event=>{const sort=encodeURIComponent(event.target.value);if($('#ledgerExport'))$('#ledgerExport').href=`/api/export?sort=${sort}`;renderLedger()});
$('#ledgerRows')?.addEventListener('click',event=>{const target=event.target;const row=target?.closest?.('tr.ledger-row');if(!row||target?.closest?.('button'))return;const button=row.querySelector('button[aria-controls]');if(button)toggleLedgerAttachment(row.dataset.ledgerId,button)});
$('#manualForm')?.addEventListener('input',event=>{if(event.target.name==='gross'||event.target.name==='tare'){const gross=Number($('#manualForm [name="gross"]')?.value||0),tare=Number($('#manualForm [name="tare"]')?.value||0);if($('#manualNet'))$('#manualNet').value=gross>tare?(gross-tare).toFixed(2):''}});
$('#manualForm')?.addEventListener('submit',async event=>{event.preventDefault();const data=Object.fromEntries(new FormData(event.target).entries());try{const result=await requestJson('/api/business/manual',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});if($('#manualResult')){$('#manualResult').textContent=result.data.message||result.data.error;$('#manualResult').className=result.ok?'success-text':'error-text'}if(result.ok){toast(`${result.data.record?.id||''} 已保存`);event.target.reset();if($('#manualNet'))$('#manualNet').value='';await loadData()}}catch(error){console.error(error);toast('保存失败，请稍后重试')}});
/* ---------- 导入流水线：可视化分步执行 ---------- */
const IMPORT_STEPS=[
  {title:'建立文件读取通道',note:'校验 .xlsx 格式，载入工作簿对象并锁定文件'},
  {title:'解析工作表结构',note:'识别过磅单 / 质检单 / 入库单 表头字段与数据行'},
  {title:'自动匹配三单关联',note:'按业务编号与时间链关联过磅、质检、入库单据'},
  {title:'写入合规数据库',note:'SQLite 事务提交，生成业务主记录与三单快照'},
  {title:'执行 9 项合规规则核验',note:'时间链、数量链、票据方式、交易频次与限额、定价一致性'},
  {title:'生成 SHA-256 数据指纹',note:'整单摘要留痕，便于发现数据被事后修改'}
];
const IMPORT_STEP_MS=[520,620,560,640,880,520];
const sleep=ms=>new Promise(resolve=>setTimeout(resolve,ms));
let importRunning=false;

function importStepMeta(index,ctx){
  if(index===0)return `${ctx.files} 个文件 · ${ctx.kb} KB`;
  if(!ctx.ready)return '处理中';
  if(index===1)return `${ctx.weigh} / ${ctx.inspect} / ${ctx.inbound} 行`;
  if(index===2)return `质检 ${ctx.matched_inspect} · 入库 ${ctx.matched_inbound}`;
  if(index===3)return `${ctx.count} 笔主记录`;
  if(index===4)return `通过 ${ctx.passed} · 风险 ${ctx.risk}`;
  return `${ctx.hashed} 条指纹`;
}
function renderImportSteps(activeIndex,ctx,status='run',failedIndex=-1){
  const list=$('#importSteps');if(!list)return;
  list.innerHTML=IMPORT_STEPS.map((step,index)=>{
    const state_=index<activeIndex?'done':index===activeIndex?(status==='fail'?'fail':'active'):'pending';
    const reached=index<=activeIndex?'reached':'';
    const meta=index<activeIndex||index===activeIndex?importStepMeta(index,ctx):'排队中';
    const icon=state_==='done'?'✓':state_==='fail'?'!':index+1;
    return `<li class="import-step ${state_} ${reached}"><span class="dot">${icon}</span><div><b>${index+1}. ${esc(step.title)}</b><p>${esc(step.note)}</p></div><span class="meta">${esc(meta)}</span></li>`;
  }).join('');
}
function setImportBadge(text,cls){const badge=$('#importModalBadge');if(!badge)return;badge.textContent=text;badge.className=`tag ${cls||'neutral'}`}
function setImportProgress(percent){
  const bar=$('#importProgressBar'),label=$('#importModalPercent');
  const value=Math.max(0,Math.min(100,Math.round(percent)));
  if(bar)bar.style.width=`${value}%`;
  if(label)label.textContent=`${value}%`;
}
function setImportStage(text){const stage=$('#importModalStage');if(stage)stage.textContent=text}
// 百分比在每一步内平滑推进，而不是步末跳变。
async function animateImportStep(index,total){
  const start=Math.round(index/total*100),end=Math.round((index+1)/total*100);
  const duration=IMPORT_STEP_MS[index],frames=Math.max(6,Math.round(duration/80));
  for(let frame=1;frame<=frames;frame++){
    setImportProgress(start+(end-start)*frame/frames);
    await sleep(Math.round(duration/frames));
  }
  setImportProgress(end);
}
// 点击右上角 × 立即关闭，不做任何二次确认；后台写入不受影响。
function closeImportModal(){$('#importModal')?.classList.remove('open')}
function importSummaryMarkup(ctx){
  const levels=ctx.levels||{};
  const levelText=Object.entries(levels).map(([level,count])=>`${level} ${count}`).join(' · ')||'无';
  return `<div class="import-summary"><div><small>写入业务</small><b>${number(ctx.count)}</b><span>笔</span></div><div><small>规则全通过</small><b>${number(ctx.passed)}</b><span>笔</span></div><div><small>触发风险</small><b>${number(ctx.risk)}</b><span>${esc(levelText)}</span></div></div>`;
}
async function runImportPipeline(formData,fileList){
  importRunning=true;
  const files=Array.from(fileList||[]);
  const ctx={files:files.length,kb:Math.max(1,Math.round(files.reduce((sum,file)=>sum+(file.size||0),0)/1024)),weigh:0,inspect:0,inbound:0,count:0,passed:0,risk:0,levels:{},matched_inspect:0,matched_inbound:0,hashed:0,ready:false};
  const submit=$('#importForm button[type="submit"]');if(submit)submit.disabled=true;
  closeImportModal();$('#importModal')?.classList.add('open');
  $('#importModalClose')&&($('#importModalClose').hidden=true);
  $('#importModalNote')&&($('#importModalNote').textContent='导入进行中，完成后需手动关闭本窗口');
  $('#importModalTitle')&&($('#importModalTitle').textContent='正在导入并核验业务单据');
  const summary=$('#importSummaryHost');if(summary)summary.remove();
  setImportBadge('准备中','neutral');setImportProgress(0);setImportStage('正在建立文件读取通道');
  renderImportSteps(0,ctx);

  let settled=false,result=null;
  const upload=requestJson('/api/import',{method:'POST',body:formData}).then(value=>{
    const data=value.data||{};const stats=data.stats||{};
    ctx.weigh=data.weigh||0;ctx.inspect=data.inspect||0;ctx.inbound=data.inbound||0;ctx.count=data.count||0;
    ctx.passed=stats.passed||0;ctx.risk=stats.risk||0;ctx.levels=stats.levels||{};
    ctx.matched_inspect=stats.matched_inspect||0;ctx.matched_inbound=stats.matched_inbound||0;ctx.hashed=stats.hashed||0;
    ctx.ready=true;ctx.ok=value.ok!==false;ctx.message=data.message||data.error||'';settled=true;result=value;return value;
  }).catch(error=>{console.error(error);ctx.ready=true;ctx.ok=false;ctx.message='导入失败，请稍后重试';settled=true;result={ok:false,data:{error:ctx.message}};return result});

  for(let index=0;index<IMPORT_STEPS.length;index++){
    renderImportSteps(index,ctx);
    setImportBadge(`第 ${index+1} / ${IMPORT_STEPS.length} 步`,'neutral');
    setImportStage(IMPORT_STEPS[index].title);
    await animateImportStep(index,IMPORT_STEPS.length);
    // 涉及真实结果的步骤必须等到服务端返回后再标记完成，避免进度条与数据不符。
    let waited=0;while(!settled&&waited<8000){await sleep(120);waited+=120}
    if(settled&&!ctx.ok){
      setImportStage('导入失败');
      renderImportSteps(index,ctx,'fail',index);
      setImportBadge('导入失败','danger');
      $('#importModalNote')&&($('#importModalNote').textContent=ctx.message||'导入失败');
      $('#importModalClose')&&($('#importModalClose').hidden=false);
      finishImportTodo(result,submit);return;
    }
    const next=$('.import-step.pending');if(next)next.classList.add('reached');
  }
  await upload;
  renderImportSteps(IMPORT_STEPS.length,ctx);
  setImportProgress(100);setImportBadge('导入成功','success');setImportStage('导入成功，正在收尾');
  $('#importModalTitle')&&($('#importModalTitle').textContent='导入成功');
  $('#importModalNote')&&($('#importModalNote').textContent=ctx.message||'已完成导入与合规核验');
  const steps=$('#importSteps');
  if(steps)steps.insertAdjacentHTML('afterend',`<div id="importSummaryHost">${importSummaryMarkup(ctx)}</div>`);
  toast(`已导入 ${ctx.count} 笔业务`);
  // 不自动关闭：由用户点击右上角 ×（或「完成」按钮）手动退出，便于逐项核对步骤与数据。
  $('#importModalNote')&&($('#importModalNote').textContent=`${ctx.message||'导入已完成'} · 核对无误后点击右上角 × 关闭`);
  finishImportTodo(result,submit);
}
function finishImportTodo(result,submit){
  if(submit)submit.disabled=false;
  importRunning=false;
  const box=$('#importResult');
  if(box&&result){box.textContent=result.data?.message||result.data?.error||'';box.className=result.ok?'success-text':'error-text'}
  loadData();
}
$('#importForm')?.addEventListener('submit',async event=>{
  event.preventDefault();
  if(importRunning){toast('已有导入任务正在执行');return}
  const files=$('#files')?.files;
  if(!files?.length){toast('请先选择 Excel 文件');return}
  const formData=new FormData();Array.from(files).forEach(file=>formData.append('files',file));
  await runImportPipeline(formData,files);
});
$('#drawer')?.addEventListener('click',event=>{if(event.target.id==='drawer')closeDrawer()});
loadData();
setInterval(()=>{if(document.visibilityState==='visible')loadData()},5000);
