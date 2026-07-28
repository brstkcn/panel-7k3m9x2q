# =====================================================================
#  Kripto Panosu — VERİ TAZELEME (çıpa bazlı kâr/zarar)
#  Çalıştır:  powershell -ExecutionPolicy Bypass -File .\kripto-panel.ps1
#  Üretir:    kripto-panel.html (görsel)  +  kripto-fiyat.json (Cowork okur, ~2.5KB)
#  Kaynaklar: Bitfinex (USD fiyat + geçmiş) • BtcTürk (TRY pariteleri) • Yahoo (USDTRY)
# =====================================================================
param([switch]$NoOpen,[switch]$Cloud)   # -Cloud: GitHub Actions modu (Cowork/açma atlanır, kripto.html üretir)

# ----------------- BURAYI DÜZENLE -----------------
# adet = toplam miktar (Bitfinex+BtcTürk birleşik), cipa = çıpa fiyatı (USD)
$ASSETS = @(
    @{ code="ETH";  sym="tETHUSD"; adet=1.4898;     cipa=1635.00 }
    @{ code="BTC";  sym="tBTCUSD"; adet=0.00939634; cipa=61996.00 }
    @{ code="USDT"; sym="tUSTUSD"; adet=435.74;     cipa=1.0 }
    @{ code="SOL";  sym="tSOLUSD"; adet=5.1493;     cipa=65.00 }
)
$CIPA_DT = "2026-06-11"                        # çıpa tarihi
$BTCTURK_PAIRS = @("BTCTRY","ETHTRY","SOLTRY","USDTTRY","BTCUSDT","ETHUSDT","SOLUSDT")
$COWORK_DIR = "C:\Users\Barış\Documents\Claude\Projects\Yatırım Danışmanlığı"   # Cowork kripto dashboard projesi ("" = kopyalama)
# --------------------------------------------------

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.WebRequest]::DefaultWebProxy = $null
$UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OUTHTML = if($Cloud){ "kripto.html" } else { "kripto-panel.html" }   # bulutta kripto.html
if($Cloud){ $COWORK_DIR = "" }                                        # bulutta Cowork kopyasi yok
$PERIODS = @( @{k="1H";days=7}, @{k="1A";days=30}, @{k="1Y";days=365} )
$errs=@()

function NowMs { [int64]([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) }
function Norm($pts){ if($pts.Count -lt 2){ return @() }; $base=$pts[0].v; @($pts | ForEach-Object { [pscustomobject]@{ t=$_.t; p=[math]::Round(($_.v/$base-1)*100,3) } }) }
function SliceNorm($pts,$days){ if($pts.Count -lt 2){ return @() }; $cut=(NowMs)-([int64]$days*86400000); $w=@($pts | Where-Object { $_.t -ge $cut }); if($w.Count -lt 2){ $w=$pts }; Norm $w }

# ---- Bitfinex: ticker + mum geçmişi ----
# ticker: [SYMBOL,BID,...,DAILY_CHANGE(5),DAILY_CHANGE_REL(6),LAST(7),VOL,HIGH,LOW]
# mum:    [MTS,OPEN,CLOSE,HIGH,LOW,VOLUME] (yeni->eski)
$tickers=@{}
try{
    $syms=($ASSETS | ForEach-Object { $_.sym }) -join ","
    $r=Invoke-RestMethod -Uri "https://api-pub.bitfinex.com/v2/tickers?symbols=$syms" -UserAgent $UA -TimeoutSec 15
    foreach($t in $r){ $tickers[$t[0]]=$t }
}catch{ $errs+="Bitfinex ticker: $($_.Exception.Message)" }

function BfxCandles($sym,$tf,$limit){
    try{
        $r=Invoke-RestMethod -Uri "https://api-pub.bitfinex.com/v2/candles/trade:${tf}:${sym}/hist?limit=$limit" -UserAgent $UA -TimeoutSec 20
        ,@($r | Sort-Object { $_[0] } | ForEach-Object { [pscustomobject]@{ t=[int64]$_[0]; v=[double]$_[2] } })
    }catch{ ,@() }
}

$items=@()
$dailyMap=@{}   # code -> günlük mumlar (mini grafik için)
foreach($a in $ASSETS){
    $tk=$tickers[$a.sym]
    if(-not $tk){ $errs+="$($a.code): ticker yok"; continue }
    $last=[double]$tk[7]; $prev=$last-[double]$tk[5]
    $chg=[math]::Round([double]$tk[6]*100,2)
    $daily=BfxCandles $a.sym "1D" 400
    $intra=BfxCandles $a.sym "15m" 96
    $dailyMap[$a.code]=$daily
    $smap=[ordered]@{}
    if($intra.Count -ge 2){ $smap["1G"]=Norm $intra }
    foreach($p in $PERIODS){ $s=SliceNorm $daily $p.days; if($s.Count -ge 2){ $smap[$p.k]=$s } }
    $adet=[double]$a.adet; $cipa=[double]$a.cipa
    $items += [pscustomobject]@{
        code=$a.code; name=$a.code+"/USD"; price=$last; prev=$prev; currency="USD"; changePct=$chg; series=$smap
        adet=$adet; bas=$cipa
        maliyet=[math]::Round($adet*$cipa,2); deger=[math]::Round($adet*$last,2)
        totalTL=[math]::Round($adet*($last-$cipa),2); totalPct= if($cipa){[math]::Round(($last/$cipa-1)*100,2)}else{0}
        dayTL=[math]::Round($adet*($last-$prev),2);   dayPct= if($prev){[math]::Round(($last/$prev-1)*100,2)}else{0}
    }
    Write-Host ("  {0} OK ({1} USD)" -f $a.code,$last) -ForegroundColor DarkGray
}

# ---- BtcTürk (ince JSON için) ----
$btcturk=@()
try{
    $r=Invoke-RestMethod -Uri "https://api.btcturk.com/api/v2/ticker" -UserAgent $UA -TimeoutSec 15
    foreach($p in $r.data){ if($p.pair -in $BTCTURK_PAIRS){ $btcturk += [pscustomobject]@{ pair=$p.pair; last=[double]$p.last; changePct24h=[double]$p.dailyPercent } } }
}catch{ $errs+="BtcTurk: $($_.Exception.Message)" }

# ---- USDTRY ----
$usdtry=$null
try{ $usdtry=[double](Invoke-RestMethod -Uri "https://query1.finance.yahoo.com/v8/finance/chart/USDTRY=X?interval=1d&range=1d" -UserAgent $UA -TimeoutSec 15).chart.result[0].meta.regularMarketPrice }
catch{ $alt=$btcturk | Where-Object { $_.pair -eq "USDTTRY" }; if($alt){ $usdtry=[double]$alt.last; $errs+="USDTRY: BtcTurk yedegi kullanildi" } else { $errs+="USDTRY: $($_.Exception.Message)" } }

# ---- Portföy (USD bazlı, çıpaya göre) ----
$pMaliyet=0.0; $pDeger=0.0; $pPrev=0.0
foreach($it in $items){ $pMaliyet+=$it.maliyet; $pDeger+=$it.deger; $pPrev+=$it.adet*$it.prev }
$portfolio=[ordered]@{
    maliyet=[math]::Round($pMaliyet,2); deger=[math]::Round($pDeger,2)
    degerTRY= if($usdtry){[math]::Round($pDeger*$usdtry,2)}else{$null}
    usdtry= if($usdtry){[math]::Round($usdtry,4)}else{$null}
    totalTL=[math]::Round($pDeger-$pMaliyet,2); totalPct= if($pMaliyet){[math]::Round(($pDeger-$pMaliyet)/$pMaliyet*100,2)}else{0}
    dayTL=[math]::Round($pDeger-$pPrev,2); dayPct= if($pPrev){[math]::Round(($pDeger-$pPrev)/$pPrev*100,2)}else{0}
}

# ---- Mini grafik: çıpadan beri günlük toplam K/Z % ----
$cipaMs=[int64]([DateTimeOffset][datetime]::ParseExact($CIPA_DT,"yyyy-MM-dd",$null)).ToUnixTimeMilliseconds()
$allT=@(); foreach($c in $dailyMap.Keys){ $allT+=@($dailyMap[$c] | Where-Object { $_.t -ge $cipaMs } | ForEach-Object { $_.t }) }
$allT=@($allT | Sort-Object -Unique)
$plSeries=@()
foreach($tt in $allT){
    $val=0.0; $ok=$true
    foreach($it in $items){
        $arr=$dailyMap[$it.code]
        $pt=@($arr | Where-Object { $_.t -le $tt })
        if($pt.Count){ $val+=$it.adet*$pt[-1].v } else { $val+=$it.adet*$it.bas }
    }
    if($pMaliyet -gt 0){ $plSeries += [pscustomobject]@{ t=$tt; p=[math]::Round(($val/$pMaliyet-1)*100,3) } }
}
$portfolio.series=$plSeries

$groups=@( [pscustomobject]@{ id="crypto"; title="🪙 Kripto (Bitfinex)"; sub="USD bazında — seçili döneme göre yüzde değişim"; items=$items } )
$data=[ordered]@{ updated=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); cipa=$CIPA_DT; portfolio=$portfolio; groups=$groups; errors=$errs }
$json=$data | ConvertTo-Json -Depth 14
Set-Content -Path (Join-Path $dir "kripto-veri.json") -Value $json -Encoding UTF8

# ---- İnce JSON (Cowork) — kripto-fiyat.ps1 ile aynı format ----
$bfxThin=@($items | ForEach-Object { [pscustomobject]@{ coin=$_.code; usd=$_.price; changePct24h=$_.changePct } })
$thin=[ordered]@{ updated=$data.updated; usdtry=$portfolio.usdtry; bitfinex=$bfxThin; btcturk=$btcturk
    portfolio=[ordered]@{ degerUSD=$portfolio.deger; degerTRY=$portfolio.degerTRY; cipadanPct=$portfolio.totalPct; gunlukPct=$portfolio.dayPct }
    errors=$errs }
$thinJson=$thin | ConvertTo-Json -Depth 6
Set-Content -Path (Join-Path $dir "kripto-fiyat.json") -Value $thinJson -Encoding UTF8

# ---- HTML ----
$tpl = @'
<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex,nofollow"><title>Kripto Panosu</title><style>
:root{--bg:#0b0e11;--card:#161a1e;--card-h:#1c2127;--border:#2a2f36;--text:#eaecef;--muted:#848e9c;--green:#16c784;--red:#ea3943}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:30px 16px}
h1{font-size:23px;font-weight:700}.sub{color:var(--muted);font-size:13px;margin:6px 0 16px}.wrap{width:100%;max-width:900px}.section{margin-bottom:34px}
.pf{display:flex;gap:14px;margin-bottom:18px;flex-wrap:wrap}
.pf-stats{display:flex;flex-direction:column;gap:10px;flex:1 1 240px;min-width:220px}
.pf-row{background:var(--card);border:1px solid var(--border);border-radius:11px;padding:11px 15px;display:flex;justify-content:space-between;align-items:baseline}
.pf-l{color:var(--muted);font-size:12px}.pf-v{font-size:17px;font-weight:700;font-variant-numeric:tabular-nums}.pf-pct{font-size:12.5px;margin-left:7px;opacity:.9}
.pf-chart{flex:2 1 320px;min-width:260px;background:var(--card);border:1px solid var(--border);border-radius:13px;padding:10px 14px 6px;display:flex;flex-direction:column}
.pc-t{font-size:12px;color:var(--muted);margin-bottom:2px}.pf-chart canvas{display:block;width:100%;flex:1;min-height:170px}
.pbar{display:flex;gap:6px;justify-content:center;flex-wrap:wrap;margin-bottom:22px;position:sticky;top:6px;z-index:10}
.pbtn{background:var(--card);border:1px solid var(--border);color:var(--muted);padding:7px 15px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer;transition:.15s}
.pbtn:hover{background:var(--card-h);color:var(--text)}.pbtn.active{background:#f7931a;border-color:#f7931a;color:#241503}
.sec-title{font-size:17px;font-weight:700;margin-bottom:4px}.sec-sub{color:var(--muted);font-size:12px;margin-bottom:14px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:12px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--border);border-radius:13px;padding:15px;transition:.2s}.card:hover{background:var(--card-h);transform:translateY(-2px)}
.c-top{display:flex;align-items:center;gap:9px;margin-bottom:10px}.c-badge{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:14px;color:#fff}
.c-name{font-weight:600;font-size:14px;line-height:1.2}.c-sub{color:var(--muted);font-size:11px;margin-top:1px}
.price-row{display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.c-price{font-size:20px;font-weight:700;font-variant-numeric:tabular-nums}
.c-buy{font-size:12px;color:#7aa7da;font-variant-numeric:tabular-nums;white-space:nowrap}
.gains{margin-top:10px;border-top:1px solid var(--border);padding-top:9px}
.gtot{font-size:13.5px;font-weight:700;font-variant-numeric:tabular-nums}
.gday{font-size:11.5px;color:var(--muted);margin-top:4px;font-variant-numeric:tabular-nums}.gday b{font-weight:600}
.up{color:var(--green)}.down{color:var(--red)}.chart-card{background:var(--card);border:1px solid var(--border);border-radius:13px;padding:16px}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:10px}.lg{display:flex;align-items:center;gap:6px;font-size:12px;font-variant-numeric:tabular-nums}
.sw{width:11px;height:11px;border-radius:3px}.cw{position:relative;width:100%}canvas.big{display:block;width:100%;height:300px}
.banner{margin-top:12px;background:#23200e;border:1px solid #50461c;color:#f5d98b;border-radius:9px;padding:10px 12px;font-size:12.5px;line-height:1.5}
</style></head><body>
<h1>🪙 Kripto Panosu</h1><div class="sub">Çıpa: __CIPA__ • Güncelleme: __UPDATED__ • USD/TRY: __FX__</div>
<div class="wrap"><div class="pf" id="pf"></div><div class="pbar" id="pbar"></div><div id="groups"></div></div>
<script>
const DATA = __DATA__;
const COLORS={BTC:"#f7931a",ETH:"#627eea",SOL:"#9945ff",USDT:"#26a17b"};
const ICONS={BTC:"₿",ETH:"Ξ",SOL:"S",USDT:"₮"};
const PALETTE=["#627eea","#f7931a","#26a17b","#9945ff","#e84393","#00cec9"];
const PLABEL={"1G":"Günlük","1H":"Haftalık","1A":"Aylık","1Y":"Yıllık"};
const ORDER=["1G","1H","1A","1Y"];
function arr(x){return x==null?[]:[].concat(x);}
function col(it,i){return COLORS[it.code]||PALETTE[i%PALETTE.length];}
function fmt(p){const d=p>=100?2:(p>=1?2:4);return "$"+p.toLocaleString("tr-TR",{minimumFractionDigits:d,maximumFractionDigits:Math.max(d,6)});}
function fUSD(v){return (v>=0?'+':'−')+"$"+Math.abs(v).toLocaleString("tr-TR",{maximumFractionDigits:2});}
function fPc(v){return (v>=0?'+':'')+v.toFixed(2)+'%';}
function presentKeys(){const s=new Set();arr(DATA.groups).forEach(g=>arr(g.items).forEach(it=>Object.keys(it.series||{}).forEach(k=>s.add(k))));return ORDER.filter(k=>s.has(k));}
let period=(function(){const k=presentKeys();return k.includes("1A")?"1A":(k[0]||"1A");})();
function ser(it){const m=it.series||{};return arr(m[period]||m["1A"]||m[Object.keys(m)[0]]);}
function cards(items){let h='<div class="grid">';items.forEach((it,i)=>{const c=col(it,i);const hasH=it.totalTL!=null;h+=`<div class="card"><div class="c-top"><div class="c-badge" style="background:${c}">${ICONS[it.code]||it.code[0]}</div><div><div class="c-name">${it.code}</div><div class="c-sub">${(it.adet||"")+" adet"}</div></div></div><div class="price-row"><div class="c-price">${fmt(it.price)}</div>${it.bas!=null?`<div class="c-buy">Çıpa ${fmt(it.bas)}</div>`:''}</div>`+(hasH?`<div class="gains"><div class="gtot ${it.totalTL>=0?'up':'down'}">Toplam ${fUSD(it.totalTL)} · ${fPc(it.totalPct)}</div><div class="gday">Bugün <b class="${it.dayTL>=0?'up':'down'}">${fUSD(it.dayTL)} · ${fPc(it.dayPct)}</b></div></div>`:'')+`</div>`;});return h+'</div>';}
function legendHTML(items){let h='<div class="legend">';items.forEach((it,i)=>{const c=col(it,i);const sd=ser(it);const s=sd.length?sd[sd.length-1].p:null;const txt=s==null?"—":(s>=0?"+":"")+s.toFixed(2)+"%";const cc=s==null?"var(--muted)":(s>=0?"#16c784":"#ea3943");h+=`<div class="lg"><span class="sw" style="background:${c}"></span>${it.code} <span style="color:${cc};font-weight:600">${txt}</span></div>`;});return h+'</div>';}
function draw(cv,items){const ctx=cv.getContext("2d");const W=cv.clientWidth,H=cv.clientHeight,dpr=window.devicePixelRatio||1;cv.width=W*dpr;cv.height=H*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);ctx.clearRect(0,0,W,H);const padL=50,padR=14,padT=12,padB=24,plotW=W-padL-padR,plotH=H-padT-padB;const series=items.map((it,i)=>({color:col(it,i),data:ser(it)})).filter(s=>s.data.length);if(!series.length){ctx.fillStyle="#848e9c";ctx.font="13px sans-serif";ctx.textAlign="center";ctx.fillText("Bu dönem için veri yok",W/2,H/2);return;}let tMin=1/0,tMax=-1/0,pMin=1/0,pMax=-1/0;for(const s of series)for(const d of s.data){if(d.t<tMin)tMin=d.t;if(d.t>tMax)tMax=d.t;if(d.p<pMin)pMin=d.p;if(d.p>pMax)pMax=d.p;}if(pMin===pMax){pMin-=.5;pMax+=.5;}const yr=pMax-pMin;pMin-=yr*.12;pMax+=yr*.12;const xP=t=>padL+(t-tMin)/((tMax-tMin)||1)*plotW,yP=p=>padT+(pMax-p)/(pMax-pMin)*plotH;ctx.font="11px sans-serif";ctx.textAlign="right";ctx.textBaseline="middle";for(let i=0;i<=5;i++){const p=pMin+(pMax-pMin)*i/5,y=yP(p);ctx.strokeStyle="#222831";ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(padL,y);ctx.lineTo(W-padR,y);ctx.stroke();ctx.fillStyle="#848e9c";ctx.fillText(p.toFixed(1)+"%",padL-8,y);}const yz=yP(0);if(yz>=padT&&yz<=padT+plotH){ctx.strokeStyle="#566";ctx.beginPath();ctx.moveTo(padL,yz);ctx.lineTo(W-padR,yz);ctx.stroke();}ctx.textAlign="center";ctx.textBaseline="top";ctx.fillStyle="#848e9c";const span=tMax-tMin;for(let i=0;i<=5;i++){const t=tMin+span*i/5,dt=new Date(t);let lbl;if(span>2*864e5){lbl=(span>400*864e5)?((dt.getMonth()+1)+"."+(dt.getFullYear()%100)):(dt.getDate()+"."+(dt.getMonth()+1));}else{lbl=String(dt.getHours()).padStart(2,"0")+":"+String(dt.getMinutes()).padStart(2,"0");}ctx.fillText(lbl,xP(t),H-padB+6);}for(const s of series){ctx.strokeStyle=s.color;ctx.lineWidth=2;ctx.lineJoin="round";ctx.beginPath();s.data.forEach((d,i)=>{const x=xP(d.t),y=yP(d.p);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.stroke();const l=s.data[s.data.length-1];ctx.fillStyle=s.color;ctx.beginPath();ctx.arc(xP(l.t),yP(l.p),3,0,7);ctx.fill();}}
const pf=DATA.portfolio;
if(pf){const tryStr=pf.degerTRY!=null?("₺"+pf.degerTRY.toLocaleString("tr-TR",{maximumFractionDigits:0})):"—";document.getElementById("pf").innerHTML=`<div class="pf-stats"><div class="pf-row"><span class="pf-l">Toplam Değer (USD)</span><span class="pf-v">$${pf.deger.toLocaleString("tr-TR",{maximumFractionDigits:2})}</span></div><div class="pf-row"><span class="pf-l">Toplam Değer (TRY)</span><span class="pf-v">${tryStr}</span></div><div class="pf-row"><span class="pf-l">Çıpadan K/Z</span><span class="pf-v ${pf.totalTL>=0?'up':'down'}">${fUSD(pf.totalTL)}<span class="pf-pct">${fPc(pf.totalPct)}</span></span></div><div class="pf-row"><span class="pf-l">Bugünkü K/Z</span><span class="pf-v ${pf.dayTL>=0?'up':'down'}">${fUSD(pf.dayTL)}<span class="pf-pct">${fPc(pf.dayPct)}</span></span></div></div><div class="pf-chart"><div class="pc-t">Çıpadan K/Z (%) — günlük</div><canvas id="pfcanvas"></canvas></div>`;}
const groupsEl=document.getElementById("groups");const cards_=[];
arr(DATA.groups).forEach(g=>{const items=arr(g.items);const sec=document.createElement("div");sec.className="section";sec.innerHTML=`<div class="sec-title">${g.title}</div><div class="sec-sub">${g.sub||""}</div>`+cards(items)+`<div class="chart-card"><div class="cw"><canvas class="big"></canvas></div><div class="lgwrap"></div></div>`;groupsEl.appendChild(sec);cards_.push({cv:sec.querySelector("canvas"),lg:sec.querySelector(".lgwrap"),items});});
const pbar=document.getElementById("pbar");
presentKeys().forEach(k=>{const b=document.createElement("button");b.className="pbtn"+(k===period?" active":"");b.textContent=PLABEL[k]||k;b.onclick=()=>{period=k;document.querySelectorAll(".pbtn").forEach(x=>x.classList.toggle("active",x===b));redraw();};pbar.appendChild(b);});
if(DATA.errors&&DATA.errors.length){const b=document.createElement("div");b.className="banner";b.innerHTML="ⓘ Not: "+DATA.errors.join("<br>");document.querySelector(".wrap").appendChild(b);}
function drawMini(){const cv=document.getElementById("pfcanvas");if(!cv||!DATA.portfolio)return;const data=arr(DATA.portfolio.series);
const ctx=cv.getContext("2d");const W=cv.clientWidth,H=cv.clientHeight,dpr=window.devicePixelRatio||1;cv.width=W*dpr;cv.height=H*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);ctx.clearRect(0,0,W,H);if(data.length<1){ctx.fillStyle="#848e9c";ctx.font="12px sans-serif";ctx.textAlign="center";ctx.fillText("Veri yok",W/2,H/2);return;}
const padL=42,padR=12,padT=10,padB=20,plotW=W-padL-padR,plotH=H-padT-padB;let tMin=1/0,tMax=-1/0,pMin=1/0,pMax=-1/0;for(const d of data){if(d.t<tMin)tMin=d.t;if(d.t>tMax)tMax=d.t;if(d.p<pMin)pMin=d.p;if(d.p>pMax)pMax=d.p;}if(pMin===pMax){pMin-=.5;pMax+=.5;}const yr=pMax-pMin;pMin-=yr*.18;pMax+=yr*.18;if(pMin>0)pMin=0;if(pMax<0)pMax=0;
const xP=t=>padL+(t-tMin)/((tMax-tMin)||1)*plotW,yP=p=>padT+(pMax-p)/(pMax-pMin)*plotH;
ctx.font="10px sans-serif";ctx.textAlign="right";ctx.textBaseline="middle";[pMax,(pMax+pMin)/2,pMin].forEach(p=>{const y=yP(p);ctx.setLineDash([]);ctx.strokeStyle="#222831";ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(padL,y);ctx.lineTo(W-padR,y);ctx.stroke();ctx.fillStyle="#848e9c";ctx.fillText(p.toFixed(1)+"%",padL-6,y);});
const step=data.length<=16?1:Math.ceil(data.length/12);ctx.textAlign="center";ctx.textBaseline="top";data.forEach((d,i)=>{if(!(i%step===0||i===data.length-1))return;const x=xP(d.t);ctx.setLineDash([3,3]);ctx.strokeStyle="#2f343c";ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(x,padT);ctx.lineTo(x,padT+plotH);ctx.stroke();ctx.setLineDash([]);const dt=new Date(d.t);ctx.fillStyle="#848e9c";ctx.fillText(dt.getDate()+"."+(dt.getMonth()+1),x,H-padB+5);});
const y0=yP(0);ctx.setLineDash([]);ctx.strokeStyle="#5a6470";ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(padL,y0);ctx.lineTo(W-padR,y0);ctx.stroke();
const last=data[data.length-1].p;const c=last>=0?"#16c784":"#ea3943";if(data.length>=2){ctx.beginPath();data.forEach((d,i)=>{const x=xP(d.t),y=yP(d.p);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.lineTo(xP(tMax),y0);ctx.lineTo(xP(tMin),y0);ctx.closePath();ctx.fillStyle=last>=0?"rgba(22,199,132,.13)":"rgba(234,57,67,.13)";ctx.fill();ctx.beginPath();data.forEach((d,i)=>{const x=xP(d.t),y=yP(d.p);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.strokeStyle=c;ctx.lineWidth=2;ctx.lineJoin="round";ctx.stroke();}
ctx.fillStyle=c;ctx.beginPath();ctx.arc(xP(tMax),yP(last),3.5,0,7);ctx.fill();}
function redraw(){cards_.forEach(o=>{draw(o.cv,o.items);o.lg.innerHTML=legendHTML(o.items);});drawMini();}
redraw();window.addEventListener("resize",redraw);
</script></body></html>
'@
$fxStr= if($usdtry){ "{0:N4}" -f $usdtry } else { "—" }
$html=$tpl.Replace("__DATA__",$json).Replace("__UPDATED__",$data.updated).Replace("__CIPA__",$CIPA_DT).Replace("__FX__",$fxStr)
Set-Content -Path (Join-Path $dir $OUTHTML) -Value $html -Encoding UTF8

# ---- Cowork köprüsü ----
if($COWORK_DIR -and (Test-Path $COWORK_DIR)){
    Copy-Item (Join-Path $dir "kripto-fiyat.json") $COWORK_DIR -Force
    Copy-Item (Join-Path $dir "kripto-panel.html") $COWORK_DIR -Force
    Write-Host "Cowork klasörüne kopyalandı." -ForegroundColor Green
}

Write-Host ""
Write-Host ("Bitti. {0} varlık. Portföy: {1:N2}$ ({2:N0}₺)  çıpadan K/Z: {3:N2}$ (%{4})" -f $items.Count,$portfolio.deger,$portfolio.degerTRY,$portfolio.totalTL,$portfolio.totalPct) -ForegroundColor Green
if($errs.Count){ Write-Host ("Notlar: "+($errs -join "; ")) -ForegroundColor Yellow }
if(-not $NoOpen){ try { Start-Process (Join-Path $dir $OUTHTML) } catch {} }
