# =====================================================================
#  Takip Panosu — VERİ TAZELEME (portföy kâr/zarar dahil)
#  Çalıştır:  powershell -ExecutionPolicy Bypass -File .\panel-guncelle.ps1
#  Üretir:    panel.html (açıp bakarsın)  +  veri.json (Cowork okur)
#  Kaynaklar: Fonlar=fvt.com.tr • Hisseler=Yahoo • Madenler=Yahoo(GC=F/SI=F×USDTRY)
# =====================================================================
param([switch]$NoOpen,[switch]$Cloud)   # -Cloud: GitHub Actions modu (Cowork/açma atlanır, index.html üretir). Anahtarsız = PC modu, aynen.

# ----------------- BURAYI DÜZENLE -----------------
$FUND_CODES  = @("DFI","GTL","PHE","YAY","JET","DTZ")   # fvt.com.tr fon kodları
$BIST_CODES  = @("KCHOL","THYAO","SAHOL")      # BIST hisse kodları (.IS otomatik)
$METALS      = @(                              # kıymetli maden (Yahoo sembolü + ad)
    @{ code="ALTIN"; ysym="GC=F"; name="Gram Altın" },
    @{ code="GUMUS"; ysym="SI=F"; name="Gram Gümüş" }
)
# Sahip olduğun adet + başlangıç (alış) fiyatı — kâr/zarar bundan hesaplanır
$HOLDINGS = @{
    DFI   = @{ adet=251518;   bas=4.136739; dt="2026-06-16" }
    GTL   = @{ adet=643000;   bas=0.125455; dt="2026-06-16" }
    PHE   = @{ adet=27211.06; bas=3.674862; dt="2026-06-18" }
    KCHOL = @{ adet=155;      bas=200.10;   dt="2026-06-16" }
    THYAO = @{ adet=94;       bas=326.50;   dt="2026-06-16" }
    SAHOL = @{ adet=304;      bas=100.80;   dt="2026-06-16" }
    ALTIN = @{ adet=57.5;     bas=6439.00;  dt="2026-06-16" }
    GUMUS = @{ adet=1752.15;  bas=104.06;   dt="2026-06-16" }
    YAY   = @{ adet=65;       bas=1832.38;  dt="2026-06-30" }
    JET   = @{ adet=45461;    bas=2.1997;   dt="2026-06-30" }
    DTZ   = @{ adet=16348;    bas=4.8934;   dt="2026-06-30" }
}
$CASH = 100900   # NAKIT (TL) — portföy toplamına dahil, K/Z'si yok
$MANUAL = @(   # fvt/Yahoo'dan otomatik çekilemeyenler — güncel fiyatı ELLE gir (değişince güncelle)
    @{ code="DMLKT"; name="Damla Kent GMS · Sertifika"; adet=8532; bas=6.12; price=6.06; dt="2026-06-16" }
)
$COWORK_DIR = "C:\Users\Barış\Claude\Projects\yani yatırım danışmanlığı"   # Cowork'ün okuyabildiği klasör (veri.json + panel.html buraya da kopyalanır). Boş "" bırakırsan kopyalamaz.
# --------------------------------------------------

$ErrorActionPreference = "Stop"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[System.Net.WebRequest]::DefaultWebProxy = $null
$UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/124 Safari/537.36"
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$OUTHTML = if($Cloud){ "index.html" } else { "panel.html" }   # bulutta kok sayfa index.html
if($Cloud){ $COWORK_DIR = "" }                                 # bulutta Cowork kopyasi yok
$PERIODS = @( @{k="1H";days=7}, @{k="1A";days=30}, @{k="1Y";days=365} )

function To-Ms([datetime]$d) { [int64]([DateTimeOffset]$d).ToUnixTimeMilliseconds() }
function NowMs { [int64]([DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()) }
function Y-Chart($sym,$iv,$rg){ $u="https://query1.finance.yahoo.com/v8/finance/chart/$sym"+"?interval=$iv&range=$rg"; (Invoke-RestMethod -Uri $u -UserAgent $UA -TimeoutSec 25).chart.result[0] }
function PtsFromRes($res){ $ts=$res.timestamp; $cl=$res.indicators.quote[0].close; $p=@(); if($ts){ $p=@(for($i=0;$i -lt $ts.Count;$i++){ $c=$cl[$i]; if($c -ne $null){ [pscustomobject]@{ t=[int64]$ts[$i]*1000; v=[double]$c } } }) }; ,$p }
function Y-Pts($sym,$iv,$rg){ try{ ,(PtsFromRes (Y-Chart $sym $iv $rg)) }catch{ ,@() } }

function Norm($pts){ if($pts.Count -lt 2){ return @() }; $base=$pts[0].v; @($pts | ForEach-Object { [pscustomobject]@{ t=$_.t; p=[math]::Round(($_.v/$base-1)*100,3) } }) }
function SliceNorm($pts,$days){ if($pts.Count -lt 2){ return @() }; $cut=(NowMs)-([int64]$days*86400000); $w=@($pts | Where-Object { $_.t -ge $cut }); if($w.Count -lt 2){ $w=$pts }; Norm $w }
function Downsample($pts,$max){ if($pts.Count -le $max){ return ,$pts }; $step=[int][math]::Ceiling($pts.Count/$max); $o=@(for($i=0;$i -lt $pts.Count;$i+=$step){ $pts[$i] }); if($o[$o.Count-1].t -ne $pts[$pts.Count-1].t){ $o+=$pts[$pts.Count-1] }; ,$o }
function Make-SeriesMap($daily,$long,$intraday){
    $o=[ordered]@{}
    if($intraday -and $intraday.Count -ge 2){ $o["1G"]=Norm $intraday }
    foreach($p in $PERIODS){ $s=SliceNorm $daily $p.days; if($s.Count -ge 2){ $o[$p.k]=$s } }
    $o
}

# ----------------- Hisseler (Yahoo) -----------------
function Get-Stock($code){
    $sym = if ($code -match '\.') { $code } else { "$code.IS" }
    $resI = Y-Chart $sym "5m" "1d"; $meta=$resI.meta
    $intra = PtsFromRes $resI
    $daily = Y-Pts $sym "1d" "1y"
    $long  = Y-Pts $sym "1wk" "5y"
    $price=[double]$meta.regularMarketPrice
    $name = if($meta.shortName){$meta.shortName}else{$code}
    # günlük = son seansın hareketi: güncel fiyat vs bir önceki günün kapanışı (seriden, meta güvenilmez)
    $prevC = if($daily.Count -ge 2){ [double]$daily[$daily.Count-2].v } elseif($meta.chartPreviousClose){ [double]$meta.chartPreviousClose } else { $price }
    $chg  = if($prevC -ne 0){($price/$prevC-1)*100}else{0}
    [pscustomobject]@{ code=$code; name=$name; price=$price; prev=$prevC; currency="TRY"; changePct=[math]::Round($chg,2); series=(Make-SeriesMap $daily $long $intra); _daily=$daily }
}

# ----------------- Madenler (Yahoo: ons USD/31.1035 × USDTRY) -----------------
$script:fxCache=@{}
function FxMap($iv,$rg){   # USDTRY -> zaman damgasina gore sirali dizi (tarih-string eslesmesi gun dusuruyordu)
    $key="$iv|$rg"; if($script:fxCache.ContainsKey($key)){ return $script:fxCache[$key] }
    $res=Y-Chart "USDTRY=X" $iv $rg; $ts=$res.timestamp; $cl=$res.indicators.quote[0].close
    $a=@(); for($i=0;$i -lt $ts.Count;$i++){ if($cl[$i] -ne $null){ $a+=[pscustomobject]@{ t=[int64]$ts[$i]*1000; v=[double]$cl[$i] } } }
    $a=@($a | Sort-Object t)
    $script:fxCache[$key]=$a; ,$a
}
function GramPts($res,$fxArr){   # her maden barina EN YAKIN FX (gun dusurmez) -> ons/31.1035 * USDTRY
    $ts=$res.timestamp; $cl=$res.indicators.quote[0].close; $p=@()
    if($ts -and $fxArr.Count){
        $j=0; $fn=$fxArr.Count
        for($i=0;$i -lt $ts.Count;$i++){
            $c=$cl[$i]; if($c -eq $null){ continue }
            $mt=[int64]$ts[$i]*1000
            while($j+1 -lt $fn -and [math]::Abs([int64]$fxArr[$j+1].t-$mt) -le [math]::Abs([int64]$fxArr[$j].t-$mt)){ $j++ }
            $p+=[pscustomobject]@{ t=$mt; v=($c/31.1035*[double]$fxArr[$j].v) }
        }
    }
    ,$p
}
function Get-Metal($m){
    $resD=Y-Chart $m.ysym "1d" "1y"; $resL=Y-Chart $m.ysym "1wk" "5y"
    $daily=GramPts $resD (FxMap "1d" "1y")
    $long =GramPts $resL (FxMap "1wk" "5y")
    if($daily.Count -lt 2){ throw "maden verisi hizalanamadı" }
    $cur=$daily[$daily.Count-1].v
    $prev= if($daily.Count -gt 1){ $daily[$daily.Count-2].v } else { $cur }
    $chg= if($prev -ne 0){ ($cur/$prev-1)*100 } else { 0 }
    [pscustomobject]@{ code=$m.code; name=$m.name; price=[math]::Round($cur,2); prev=$prev; currency="TRY"; changePct=[math]::Round($chg,2); series=(Make-SeriesMap $daily $long $null); _daily=$daily }
}

# ----------------- Fonlar (fvt.com.tr) -----------------
$script:fvtToken=$null
function Get-FvtToken { if($script:fvtToken){return $script:fvtToken}; $h=@{"User-Agent"=$UA;"Accept"="application/json";"Origin"="https://fvt.com.tr";"Referer"="https://fvt.com.tr/"}; $script:fvtToken=(Invoke-RestMethod -Uri "https://fvt.com.tr/api/app-token" -Headers $h -TimeoutSec 20).data.token; $script:fvtToken }
function Get-Fund($code){
    $tok=Get-FvtToken
    $h=@{"User-Agent"=$UA;"Accept"="application/json";"Origin"="https://fvt.com.tr";"Referer"="https://fvt.com.tr/fonlar/yatirim-fonlari/$code";"Authorization"="Bearer $tok"}
    $info=(Invoke-RestMethod -Uri "https://fvt.com.tr/api/funds/$code" -Headers $h -TimeoutSec 25).data.fund
    $prices=(Invoke-RestMethod -Uri "https://fvt.com.tr/api/funds/$code/prices" -Headers $h -TimeoutSec 25).data
    $daily=@($prices | Sort-Object { $_.tarih } | ForEach-Object { [pscustomobject]@{ t=(To-Ms ([datetime]::ParseExact($_.tarih,"yyyy-MM-dd",$null))); v=[double]$_.fiyat } })
    $long = Downsample $daily 320
    $price=[double]$info.fiyat
    $prevF= if($daily.Count -ge 2){ $daily[$daily.Count-2].v } else { $price }
    [pscustomobject]@{ code=$code; name=$info.fonAdi; price=$price; prev=$prevF; currency="TRY"; changePct=[math]::Round([double]$info.getiri,2); series=(Make-SeriesMap $daily $long $null); _daily=$daily }
}

# ----------------- Topla -----------------
Write-Host "Veri çekiliyor..." -ForegroundColor Cyan
$errs=@()
$stocks=@(); foreach($c in $BIST_CODES){ try{ $stocks+=(Get-Stock $c); Write-Host "  HISSE $c OK" -ForegroundColor DarkGray }catch{ $errs+="HISSE $c : $($_.Exception.Message)" } }
$funds=@();  foreach($c in $FUND_CODES){ try{ $funds +=(Get-Fund  $c); Write-Host "  FON   $c OK" -ForegroundColor DarkGray }catch{ $errs+="FON $c : $($_.Exception.Message)" } }
$metalsOut=@(); foreach($m in $METALS){ try{ $metalsOut+=(Get-Metal $m); Write-Host "  MADEN $($m.code) OK" -ForegroundColor DarkGray }catch{ $errs+="MADEN $($m.code) : $($_.Exception.Message)" } }

# ----------------- Elle girilen varlıklar (DMLKT gibi) -----------------
$nowMsM = NowMs
$manualOut=@(foreach($x in $MANUAL){
    $HOLDINGS[$x.code]=@{ adet=$x.adet; bas=$x.bas; dt=$x.dt }
    $dtMsM=(To-Ms ([datetime]::ParseExact($x.dt,"yyyy-MM-dd",$null)))
    [pscustomobject]@{ code=$x.code; name=$x.name; price=[double]$x.price; prev=[double]$x.price; currency="TRY"; changePct=0; series=@{}; _daily=@([pscustomobject]@{t=$dtMsM;v=[double]$x.price},[pscustomobject]@{t=$nowMsM;v=[double]$x.price}) }
})

# ----------------- Portföy kâr/zarar -----------------
$pMaliyet=0.0; $pDeger=0.0; $pPrevDeger=0.0
foreach($it in @($stocks + $funds + $metalsOut + $manualOut)){
    if($HOLDINGS.ContainsKey($it.code)){
        $hh=$HOLDINGS[$it.code]; $adet=[double]$hh.adet; $bas=[double]$hh.bas
        $price=[double]$it.price; $prev=[double]$it.prev
        $maliyet=$adet*$bas; $deger=$adet*$price
        $tTL=$deger-$maliyet; $tPct= if($bas){ ($price/$bas-1)*100 } else {0}
        $dTL=$adet*($price-$prev); $dPct= if($prev){ ($price/$prev-1)*100 } else {0}
        Add-Member -InputObject $it -Force -NotePropertyMembers @{
            adet=$adet; bas=$bas
            maliyet=[math]::Round($maliyet,2); deger=[math]::Round($deger,2)
            totalTL=[math]::Round($tTL,2); totalPct=[math]::Round($tPct,2)
            dayTL=[math]::Round($dTL,2);   dayPct=[math]::Round($dPct,2)
        }
        $pMaliyet+=$maliyet; $pDeger+=$deger; $pPrevDeger+=$adet*$prev
    }
}
if($CASH){ $pMaliyet+=$CASH; $pDeger+=$CASH; $pPrevDeger+=$CASH }
$portfolio=[ordered]@{
    maliyet=[math]::Round($pMaliyet,2); deger=[math]::Round($pDeger,2)
    totalTL=[math]::Round($pDeger-$pMaliyet,2); totalPct= if($pMaliyet){[math]::Round(($pDeger-$pMaliyet)/$pMaliyet*100,2)}else{0}
    dayTL=[math]::Round($pDeger-$pPrevDeger,2); dayPct= if($pPrevDeger){[math]::Round(($pDeger-$pPrevDeger)/$pPrevDeger*100,2)}else{0}
}

# ---- portföyün günlük toplam K/Z % geçmişi (küçük grafik için) ----
$held=@(@($stocks+$funds+$metalsOut+$manualOut) | Where-Object { $HOLDINGS.ContainsKey($_.code) })
$assetSeries=@{}
foreach($it in $held){
    $lst=@(); foreach($pt in $it._daily){ $d=([DateTimeOffset]::FromUnixTimeMilliseconds([int64]$pt.t)).LocalDateTime.ToString("yyyy-MM-dd"); $lst+=[pscustomobject]@{ d=$d; v=[double]$pt.v } }
    $assetSeries[$it.code]=@($lst | Sort-Object d)
}
$minDt=@($HOLDINGS.Values | ForEach-Object { $_.dt } | Sort-Object)[0]
$allDates=@($held | ForEach-Object { $assetSeries[$_.code].d } | Where-Object { $_ -ge $minDt } | Sort-Object -Unique)
$ptr=@{}; $last=@{}; foreach($it in $held){ $ptr[$it.code]=0; $last[$it.code]=$null }
$plSeries=@()
foreach($date in $allDates){
    $cost=0.0; $val=0.0
    foreach($it in $held){
        $code=$it.code; $hh=$HOLDINGS[$code]
        if($hh.dt -gt $date){ continue }
        $arr=$assetSeries[$code]
        while($ptr[$code] -lt $arr.Count -and $arr[$ptr[$code]].d -le $date){ $last[$code]=$arr[$ptr[$code]].v; $ptr[$code]++ }
        $price= if($last[$code] -ne $null){ $last[$code] } else { [double]$hh.bas }
        $cost+=[double]$hh.adet*[double]$hh.bas; $val+=[double]$hh.adet*$price
    }
    if($CASH){ $cost+=$CASH; $val+=$CASH }
    if($cost -gt 0){ $plSeries+=[pscustomobject]@{ t=(To-Ms ([datetime]::ParseExact($date,"yyyy-MM-dd",$null))); p=[math]::Round(($val/$cost-1)*100,3) } }
}
$portfolio.series=$plSeries
foreach($it in @($stocks+$funds+$metalsOut+$manualOut)){ $it.PSObject.Properties.Remove('_daily') | Out-Null }

$cashItems = if($CASH){ @([pscustomobject]@{ code="NAKIT"; name="Nakit"; price=[double]$CASH; currency="TRY"; changePct=0; series=@{} }) } else { @() }
$groups=@(
  [pscustomobject]@{ id="stocks"; title="📈 Hisseler (BIST)";     sub="Seçili döneme göre yüzde değişim"; items=$stocks }
  [pscustomobject]@{ id="funds";  title="💼 Fonlar (fvt.com.tr)"; sub="Seçili döneme göre yüzde değişim"; items=$funds }
  [pscustomobject]@{ id="metals"; title="🥇 Kıymetli Maden";       sub="Gram fiyat (uluslararası ons × USDTRY) — seçili döneme göre %"; items=$metalsOut }
)
if(@($manualOut).Count){ $groups += [pscustomobject]@{ id="manual"; title="📜 Sertifika (elle)"; sub="fvt'den otomatik çekilemez — fiyat elle güncellenir"; items=@($manualOut) } }
if($CASH){ $groups += [pscustomobject]@{ id="cash"; title="💵 Nakit"; sub="Portföydeki nakit — K/Z yok"; items=@($cashItems) } }
$data=[ordered]@{ updated=(Get-Date).ToString("yyyy-MM-dd HH:mm:ss"); portfolio=$portfolio; groups=$groups; errors=$errs }
$json=$data | ConvertTo-Json -Depth 14
Set-Content -Path (Join-Path $dir "veri.json") -Value $json -Encoding UTF8

# ----------------- panel.html üret -----------------
$tpl = @'
<!DOCTYPE html><html lang="tr"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta name="robots" content="noindex,nofollow"><title>Takip Panosu</title><style>
:root{--bg:#0b0e11;--card:#161a1e;--card-h:#1c2127;--border:#2a2f36;--text:#eaecef;--muted:#848e9c;--green:#16c784;--red:#ea3943}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",Roboto,sans-serif;background:var(--bg);color:var(--text);min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:30px 16px}
h1{font-size:23px;font-weight:700}.sub{color:var(--muted);font-size:13px;margin:6px 0 16px}.wrap{width:100%;max-width:900px}.section{margin-bottom:34px}
.pf{display:flex;gap:14px;margin-bottom:18px;flex-wrap:wrap}
.pf-stats{display:flex;flex-direction:column;gap:10px;flex:1 1 230px;min-width:210px}
.pf-row{background:var(--card);border:1px solid var(--border);border-radius:11px;padding:11px 15px;display:flex;justify-content:space-between;align-items:baseline}
.pf-l{color:var(--muted);font-size:12px}.pf-v{font-size:17px;font-weight:700;font-variant-numeric:tabular-nums}.pf-pct{font-size:12.5px;margin-left:7px;opacity:.9}
.pf-chart{flex:2 1 320px;min-width:260px;background:var(--card);border:1px solid var(--border);border-radius:13px;padding:10px 14px 6px;display:flex;flex-direction:column}
.pc-t{font-size:12px;color:var(--muted);margin-bottom:2px}
.pf-chart canvas{display:block;width:100%;flex:1;min-height:160px}
.pbar{display:flex;gap:6px;justify-content:center;flex-wrap:wrap;margin-bottom:22px;position:sticky;top:6px;z-index:10}
.pbtn{background:var(--card);border:1px solid var(--border);color:var(--muted);padding:7px 15px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer;transition:.15s}
.pbtn:hover{background:var(--card-h);color:var(--text)}.pbtn.active{background:var(--green);border-color:var(--green);color:#06281c}
.sec-title{font-size:17px;font-weight:700;margin-bottom:4px}.sec-sub{color:var(--muted);font-size:12px;margin-bottom:14px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(185px,1fr));gap:12px;margin-bottom:16px}
.card{background:var(--card);border:1px solid var(--border);border-radius:13px;padding:15px;transition:.2s}.card:hover{background:var(--card-h);transform:translateY(-2px)}
.c-top{display:flex;align-items:center;gap:9px;margin-bottom:10px}.c-badge{width:32px;height:32px;border-radius:8px;display:flex;align-items:center;justify-content:center;font-weight:700;font-size:11px;color:#fff}
.c-name{font-weight:600;font-size:14px;line-height:1.2}.c-sub{color:var(--muted);font-size:11px;margin-top:1px}
.price-row{display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.c-price{font-size:20px;font-weight:700;font-variant-numeric:tabular-nums}
.c-buy{font-size:12px;color:#7aa7da;font-variant-numeric:tabular-nums;white-space:nowrap}
.gains{margin-top:10px;border-top:1px solid var(--border);padding-top:9px}
.gtot{font-size:13.5px;font-weight:700;font-variant-numeric:tabular-nums}
.gday{font-size:11.5px;color:var(--muted);margin-top:4px;font-variant-numeric:tabular-nums}.gday b{font-weight:600}
.up{color:var(--green)}.down{color:var(--red)}.chart-card{background:var(--card);border:1px solid var(--border);border-radius:13px;padding:16px}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin-top:10px}.lg{display:flex;align-items:center;gap:6px;font-size:12px;font-variant-numeric:tabular-nums}
.sw{width:11px;height:11px;border-radius:3px}.cw{position:relative;width:100%}canvas{display:block;width:100%;height:300px}
.banner{margin-top:12px;background:#23200e;border:1px solid #50461c;color:#f5d98b;border-radius:9px;padding:10px 12px;font-size:12.5px;line-height:1.5}
</style></head><body>
<h1>📊 Takip Panosu</h1><div class="sub">Veri yerelde çekilip gömüldü • Güncelleme: __UPDATED__</div>
<div class="wrap"><div class="pf" id="pf"></div><div class="pbar" id="pbar"></div><div id="groups"></div></div>
<script>
const DATA = __DATA__;
const PALETTE=["#f7931a","#627eea","#14f195","#e84393","#00cec9","#fdcb6e","#a29bfe","#ff7675","#55efc4","#74b9ff"];
const PLABEL={"1G":"Günlük","1H":"Haftalık","1A":"Aylık","1Y":"Yıllık"};
const ORDER=["1G","1H","1A","1Y"];
function arr(x){return x==null?[]:[].concat(x);}
function fmt(p,cur){const s=p.toLocaleString("tr-TR",{minimumFractionDigits:2,maximumFractionDigits:6});return s+" "+(cur==="TRY"?"₺":cur);}
function fTL(v){return (v>=0?'+':'−')+Math.abs(v).toLocaleString("tr-TR",{maximumFractionDigits:0})+' ₺';}
function fPc(v){return (v>=0?'+':'')+v.toFixed(2)+'%';}
function gline(lbl,tl,pc){const up=tl>=0;return `<div class="gline"><span class="gl-l">${lbl}</span><span class="gl-v ${up?'up':'down'}">${fTL(tl)} · ${fPc(pc)}</span></div>`;}
function presentKeys(){const s=new Set();arr(DATA.groups).forEach(g=>arr(g.items).forEach(it=>Object.keys(it.series||{}).forEach(k=>s.add(k))));return ORDER.filter(k=>s.has(k));}
let period=(function(){const k=presentKeys();return k.includes("1A")?"1A":(k[0]||"1A");})();
function ser(it){const m=it.series||{};return arr(m[period]||m["3A"]||m["1A"]||m[Object.keys(m)[0]]);}
function cards(items){let h='<div class="grid">';items.forEach((it,i)=>{const col=PALETTE[i%PALETTE.length];const hasH=it.totalTL!=null;h+=`<div class="card"><div class="c-top"><div class="c-badge" style="background:${col}">${it.code.slice(0,5)}</div><div><div class="c-name">${it.code}</div><div class="c-sub" title="${it.name||""}">${(it.name||"").slice(0,24)}</div></div></div><div class="price-row"><div class="c-price">${it.price!=null?fmt(it.price,it.currency):"—"}</div>${it.bas!=null?`<div class="c-buy">Alış ${fmt(it.bas,it.currency)}</div>`:''}</div>`+(hasH?`<div class="gains"><div class="gtot ${it.totalTL>=0?'up':'down'}">Toplam ${fTL(it.totalTL)} · ${fPc(it.totalPct)}</div><div class="gday">Bugün <b class="${it.dayTL>=0?'up':'down'}">${fTL(it.dayTL)} · ${fPc(it.dayPct)}</b></div></div>`:'')+`</div>`;});return h+'</div>';}
function legendHTML(items){let h='<div class="legend">';items.forEach((it,i)=>{const col=PALETTE[i%PALETTE.length];const sd=ser(it);const s=sd.length?sd[sd.length-1].p:null;const txt=s==null?"—":(s>=0?"+":"")+s.toFixed(2)+"%";const c=s==null?"var(--muted)":(s>=0?"#16c784":"#ea3943");h+=`<div class="lg"><span class="sw" style="background:${col}"></span>${it.code} <span style="color:${c};font-weight:600">${txt}</span></div>`;});return h+'</div>';}
function draw(cv,items){const ctx=cv.getContext("2d");const W=cv.clientWidth,H=cv.clientHeight,dpr=window.devicePixelRatio||1;cv.width=W*dpr;cv.height=H*dpr;ctx.setTransform(dpr,0,0,dpr,0,0);ctx.clearRect(0,0,W,H);const padL=50,padR=14,padT=12,padB=24,plotW=W-padL-padR,plotH=H-padT-padB;const series=items.map((it,i)=>({color:PALETTE[i%PALETTE.length],data:ser(it)})).filter(s=>s.data.length);if(!series.length){ctx.fillStyle="#848e9c";ctx.font="13px sans-serif";ctx.textAlign="center";ctx.fillText("Bu dönem için veri yok",W/2,H/2);return;}let tMin=1/0,tMax=-1/0,pMin=1/0,pMax=-1/0;for(const s of series)for(const d of s.data){if(d.t<tMin)tMin=d.t;if(d.t>tMax)tMax=d.t;if(d.p<pMin)pMin=d.p;if(d.p>pMax)pMax=d.p;}if(pMin===pMax){pMin-=.5;pMax+=.5;}const yr=pMax-pMin;pMin-=yr*.12;pMax+=yr*.12;const xP=t=>padL+(t-tMin)/((tMax-tMin)||1)*plotW,yP=p=>padT+(pMax-p)/(pMax-pMin)*plotH;ctx.font="11px sans-serif";ctx.textAlign="right";ctx.textBaseline="middle";for(let i=0;i<=5;i++){const p=pMin+(pMax-pMin)*i/5,y=yP(p);ctx.strokeStyle="#222831";ctx.lineWidth=1;ctx.beginPath();ctx.moveTo(padL,y);ctx.lineTo(W-padR,y);ctx.stroke();ctx.fillStyle="#848e9c";ctx.fillText(p.toFixed(1)+"%",padL-8,y);}const yz=yP(0);if(yz>=padT&&yz<=padT+plotH){ctx.strokeStyle="#566";ctx.beginPath();ctx.moveTo(padL,yz);ctx.lineTo(W-padR,yz);ctx.stroke();}ctx.textAlign="center";ctx.textBaseline="top";ctx.fillStyle="#848e9c";const span=tMax-tMin;for(let i=0;i<=5;i++){const t=tMin+span*i/5,dt=new Date(t);let lbl;if(span>2*864e5){lbl=(span>400*864e5)?((dt.getMonth()+1)+"."+(dt.getFullYear()%100)):(dt.getDate()+"."+(dt.getMonth()+1));}else{lbl=String(dt.getHours()).padStart(2,"0")+":"+String(dt.getMinutes()).padStart(2,"0");}ctx.fillText(lbl,xP(t),H-padB+6);}for(const s of series){ctx.strokeStyle=s.color;ctx.lineWidth=2;ctx.lineJoin="round";ctx.beginPath();s.data.forEach((d,i)=>{const x=xP(d.t),y=yP(d.p);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.stroke();const l=s.data[s.data.length-1];ctx.fillStyle=s.color;ctx.beginPath();ctx.arc(xP(l.t),yP(l.p),3,0,7);ctx.fill();}}
const pf=DATA.portfolio;
if(pf){document.getElementById("pf").innerHTML=`<div class="pf-stats"><div class="pf-row"><span class="pf-l">Güncel Değer</span><span class="pf-v">${pf.deger.toLocaleString("tr-TR",{maximumFractionDigits:0})} ₺</span></div><div class="pf-row"><span class="pf-l">Toplam K/Z</span><span class="pf-v ${pf.totalTL>=0?'up':'down'}">${fTL(pf.totalTL)}<span class="pf-pct">${fPc(pf.totalPct)}</span></span></div><div class="pf-row"><span class="pf-l">Bugünkü K/Z</span><span class="pf-v ${pf.dayTL>=0?'up':'down'}">${fTL(pf.dayTL)}<span class="pf-pct">${fPc(pf.dayPct)}</span></span></div><div class="pf-row"><span class="pf-l">Toplam Maliyet</span><span class="pf-v">${pf.maliyet.toLocaleString("tr-TR",{maximumFractionDigits:0})} ₺</span></div></div><div class="pf-chart"><div class="pc-t">Toplam K/Z (%) — günlük artan</div><canvas id="pfcanvas"></canvas></div>`;}
const groupsEl=document.getElementById("groups");const cards_=[];
arr(DATA.groups).forEach(g=>{const items=arr(g.items);const hasChart=items.some(it=>Object.keys(it.series||{}).length);const sec=document.createElement("div");sec.className="section";sec.innerHTML=`<div class="sec-title">${g.title}</div><div class="sec-sub">${g.sub||""}</div>`+cards(items)+(hasChart?`<div class="chart-card"><div class="cw"><canvas></canvas></div><div class="lgwrap"></div></div>`:'');groupsEl.appendChild(sec);if(hasChart)cards_.push({cv:sec.querySelector("canvas"),lg:sec.querySelector(".lgwrap"),items});});
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
const last=data[data.length-1].p;const col=last>=0?"#16c784":"#ea3943";if(data.length>=2){ctx.beginPath();data.forEach((d,i)=>{const x=xP(d.t),y=yP(d.p);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.lineTo(xP(tMax),y0);ctx.lineTo(xP(tMin),y0);ctx.closePath();ctx.fillStyle=last>=0?"rgba(22,199,132,.13)":"rgba(234,57,67,.13)";ctx.fill();ctx.beginPath();data.forEach((d,i)=>{const x=xP(d.t),y=yP(d.p);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.strokeStyle=col;ctx.lineWidth=2;ctx.lineJoin="round";ctx.stroke();}
const lx=xP(tMax),ly=yP(last);ctx.fillStyle=col;ctx.beginPath();ctx.arc(lx,ly,3.5,0,7);ctx.fill();}
function redraw(){cards_.forEach(o=>{draw(o.cv,o.items);o.lg.innerHTML=legendHTML(o.items);});drawMini();}
redraw();window.addEventListener("resize",redraw);
</script></body></html>
'@
$html=$tpl.Replace("__DATA__",$json).Replace("__UPDATED__",$data.updated)
Set-Content -Path (Join-Path $dir $OUTHTML) -Value $html -Encoding UTF8

# ----------------- Cowork köprüsü: çıktıyı Cowork'ün eriştiği klasöre de kopyala -----------------
if($COWORK_DIR){
    try{
        if(-not (Test-Path $COWORK_DIR)){ New-Item -ItemType Directory -Force -Path $COWORK_DIR | Out-Null }
        Copy-Item (Join-Path $dir "veri.json")  $COWORK_DIR -Force
        Copy-Item (Join-Path $dir "panel.html") $COWORK_DIR -Force
        Write-Host "Cowork klasörüne kopyalandı: $COWORK_DIR" -ForegroundColor Green
    }catch{ Write-Host "Cowork kopyalama hatası: $($_.Exception.Message)" -ForegroundColor Yellow }
}

Write-Host ""
Write-Host ("Bitti. Hisse:{0} Fon:{1} Maden:{2}" -f $stocks.Count,$funds.Count,$metalsOut.Count) -ForegroundColor Green
Write-Host ("Portföy: maliyet={0:N0}₺  değer={1:N0}₺  toplam K/Z={2:N0}₺ (%{3})  bugün={4:N0}₺ (%{5})" -f $portfolio.maliyet,$portfolio.deger,$portfolio.totalTL,$portfolio.totalPct,$portfolio.dayTL,$portfolio.dayPct) -ForegroundColor Green
if($errs.Count){ Write-Host "Notlar: $($errs -join '; ')" -ForegroundColor Yellow }
# ---- Cowork için İNCE özet (geçmiş seri YOK, ~2 KB → kırpılmadan senkron olur) ----
$ozetItems = foreach($it in @($stocks + $funds + $metalsOut + $manualOut)){
    [pscustomobject]@{ code=$it.code; price=$it.price; prev=[math]::Round([double]$it.prev,6); changePct=$it.changePct }
}
$ozet = [ordered]@{
    updated   = $data.updated
    portfolio = [ordered]@{
        maliyet=$portfolio.maliyet; deger=$portfolio.deger
        totalTL=$portfolio.totalTL; totalPct=$portfolio.totalPct
        dayTL=$portfolio.dayTL;     dayPct=$portfolio.dayPct
    }
    prices = $ozetItems
}
$ozetJson = $ozet | ConvertTo-Json -Depth 6
Set-Content -Path (Join-Path $dir "ozet.json") -Value $ozetJson -Encoding UTF8
if($COWORK_DIR){ Copy-Item (Join-Path $dir "ozet.json") (Join-Path $COWORK_DIR "ozet.json") -Force }

if(-not $NoOpen){ try { Start-Process (Join-Path $dir $OUTHTML) } catch {} }
