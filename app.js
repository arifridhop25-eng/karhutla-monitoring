// Kalimantan Fire Monitoring System - client-side prototype
const mapOptions = { center: [-1.5, 114.0], zoom: 5, minZoom: 3 };
const maps = {};
const layerRegistry = new Map();
let nextLayerId = 1;
let chart = null;
let selectedFeatures = [];

function makeMap(id) {
  const m = L.map(id, { preferCanvas: true }).setView(mapOptions.center, mapOptions.zoom);
  const osm = L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
    maxZoom: 19, attribution: '&copy; OpenStreetMap contributors'
  });
  const imagery = L.tileLayer('https://server.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}', {
    maxZoom: 18, attribution: 'Tiles &copy; Esri'
  });
  osm.addTo(m);
  L.control.layers({ 'OpenStreetMap': osm, 'Citra satelit': imagery }, {}, { collapsed: true }).addTo(m);
  return m;
}
maps.dashboard = makeMap('mapDashboard');
maps.full = makeMap('mapFull');

function notify(message, error=false) {
  const el = document.getElementById('toast');
  el.textContent = message;
  el.style.background = error ? '#492d31' : '#1d3e38';
  el.classList.add('show');
  setTimeout(() => el.classList.remove('show'), 3200);
}
function setUploadMessage(message) {
  document.getElementById('uploadMessage').textContent = message;
}
function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}
function featurePopup(feature, layer) {
  const props = feature?.properties || {};
  const rows = Object.entries(props).slice(0, 40).map(([k,v]) =>
    `<tr><td>${escapeHtml(k)}</td><td>${escapeHtml(v)}</td></tr>`).join('');
  layer.bindPopup(`<b>Informasi fitur</b><table class="attribute-table">${rows || '<tr><td colspan="2">Tidak ada atribut</td></tr>'}</table>`);
  layer.on('click', () => {
    document.getElementById('featureDetails').innerHTML =
      `<table class="attribute-table">${rows || '<tr><td colspan="2">Tidak ada atribut</td></tr>'}</table>`;
  });
}
function inferType(data) {
  const geom = data?.features?.[0]?.geometry?.type || data?.geometry?.type;
  if (geom) return `Vektor · ${geom}`;
  return 'Vektor';
}
function addLayer(name, type, leafletLayer, count=0, meta={}) {
  const id = `layer-${nextLayerId++}`;
  const record = { id, name, type, layer: leafletLayer, secondLayer: meta.secondLayer || null, count, visible: true, meta };
  layerRegistry.set(id, record);
  leafletLayer.addTo(maps.dashboard);
  if (record.secondLayer) record.secondLayer.addTo(maps.full);
  refreshUI();
  const bounds = leafletLayer.getBounds?.();
  if (bounds?.isValid?.()) {
    maps.dashboard.fitBounds(bounds.pad(0.08));
    maps.full.fitBounds(bounds.pad(0.08));
  }
  notify(`${name} berhasil dimuat.`);
  return id;
}
function makeGeoJSONLayer(data) {
  return L.geoJSON(data, {
    pointToLayer: (feature, latlng) => L.circleMarker(latlng, {
      radius: 5, color: '#fff1df', weight: 1, fillColor: '#f05e4e', fillOpacity: .9
    }),
    style: feature => {
      const p = feature?.properties || {};
      const cls = String(p.kelas || p.class || p.risk || '').toLowerCase();
      const color = cls.includes('tinggi') || cls.includes('high') ? '#b94f70' : '#ed9848';
      return { color, weight: 1.6, fillColor: color, fillOpacity: .32 };
    },
    onEachFeature: featurePopup
  });
}
function addGeoJSON(data, name) {
  if (!data || (!data.type && !data.features)) throw new Error('Struktur GeoJSON tidak dikenali.');
  const layer = makeGeoJSONLayer(data);
  const secondLayer = makeGeoJSONLayer(data);
  const count = data.type === 'FeatureCollection' ? (data.features || []).length : 1;
  return addLayer(name, inferType(data), layer, count, { data, secondLayer });
}
async function readGeoJSONFile(file) {
  const data = JSON.parse(await file.text());
  addGeoJSON(data, file.name);
}
async function readShapefile(file) {
  const buffer = await file.arrayBuffer();
  const result = await shp(buffer);
  if (result.type === 'FeatureCollection') addGeoJSON(result, file.name);
  else if (result && typeof result === 'object') {
    Object.entries(result).forEach(([name, geojson]) => addGeoJSON(geojson, `${file.name} · ${name}`));
  } else throw new Error('Tidak dapat membaca isi ZIP Shapefile.');
}
function normalizeKey(row, candidates) {
  const keys = Object.keys(row);
  return keys.find(k => candidates.includes(k.trim().toLowerCase()));
}
function rowsToGeoJSON(rows, name) {
  if (!rows.length) throw new Error('Tabel kosong.');
  const latKey = normalizeKey(rows[0], ['latitude','lat','y']);
  const lonKey = normalizeKey(rows[0], ['longitude','lon','long','lng','x']);
  if (!latKey || !lonKey) throw new Error('Kolom koordinat tidak ditemukan. Gunakan nama latitude/longitude atau lat/lon.');
  const features = [];
  rows.forEach(row => {
    const lat = Number(row[latKey]), lon = Number(row[lonKey]);
    if (!Number.isFinite(lat) || !Number.isFinite(lon) || lat < -90 || lat > 90 || lon < -180 || lon > 180) return;
    features.push({ type:'Feature', geometry:{type:'Point',coordinates:[lon,lat]}, properties:row });
  });
  if (!features.length) throw new Error('Tidak ditemukan baris dengan koordinat valid.');
  addGeoJSON({type:'FeatureCollection',features}, name);
}
async function readTable(file) {
  const ext = file.name.split('.').pop().toLowerCase();
  let rows;
  if (ext === 'csv') {
    rows = await new Promise((resolve,reject) => Papa.parse(file, {
      header:true, skipEmptyLines:true, dynamicTyping:false,
      complete:r=>resolve(r.data), error:reject
    }));
  } else {
    const workbook = XLSX.read(await file.arrayBuffer(), {type:'array'});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(sheet, {defval:''});
  }
  rowsToGeoJSON(rows, file.name);
}
async function readGeoTIFF(file) {
  if (typeof parseGeoraster !== 'function' || typeof GeoRasterLayer === 'undefined')
    throw new Error('Pustaka GeoTIFF belum berhasil dimuat. Periksa koneksi internet.');
  const raster = await parseGeoraster(await file.arrayBuffer());
  const makeRasterLayer = () => new GeoRasterLayer({
    georaster: raster, opacity: 0.72, resolution: 256,
    pixelValuesToColorFn: values => {
      const v = values[0];
      if (v === null || v === undefined || Number.isNaN(v)) return null;
      // Default grayscale stretch; sesuaikan color ramp dengan klasifikasi raster milikmu.
      const min = raster.mins?.[0] ?? 0, max = raster.maxs?.[0] ?? 255;
      const t = Math.max(0, Math.min(1, (v-min) / ((max-min) || 1)));
      const g = Math.round(35 + t*210);
      return `rgba(${g},${g},${g},0.85)`;
    }
  });
  const rasterLayer = makeRasterLayer();
  const secondLayer = makeRasterLayer();
  addLayer(file.name, 'Raster · GeoTIFF', rasterLayer, 0, { raster:true, secondLayer });
}
async function readImageOverlay(file) {
  const north = Number(document.getElementById('north').value);
  const south = Number(document.getElementById('south').value);
  const east = Number(document.getElementById('east').value);
  const west = Number(document.getElementById('west').value);
  if (![north,south,east,west].every(Number.isFinite) || north <= south || east <= west)
    throw new Error('Masukkan batas koordinat valid: utara > selatan dan timur > barat.');
  const url = URL.createObjectURL(file);
  const bounds = [[south,west],[north,east]];
  const overlay1 = L.imageOverlay(url,bounds,{opacity:.75});
  const overlay2 = L.imageOverlay(url,bounds,{opacity:.75});
  addLayer(file.name,'Image overlay · PNG/JPG',overlay1,1,{image:true,url,secondLayer:overlay2});
}
function updateStats() {
  let pointCount=0, provinces=new Set(), areaHa=0, areaKm2=0, hasHa=false, hasKm=false;
  for (const rec of layerRegistry.values()) {
    if (!rec.meta.data) continue;
    const features = rec.meta.data.features || [];
    features.forEach(f => {
      if (f.geometry?.type === 'Point') pointCount++;
      const p=f.properties||{};
      const provinceKey=normalizeKey(p,['provinsi','province','nama_provinsi']);
      if (provinceKey && p[provinceKey]) provinces.add(String(p[provinceKey]));
      const haKey=normalizeKey(p,['area_ha','luas_ha','luas_hektar']);
      const kmKey=normalizeKey(p,['area_km2','luas_km2']);
      if (haKey && Number.isFinite(Number(p[haKey]))) {areaHa+=Number(p[haKey]);hasHa=true;}
      else if (kmKey && Number.isFinite(Number(p[kmKey]))) {areaKm2+=Number(p[kmKey]);hasKm=true;}
    });
  }
  document.getElementById('hotspotTotal').textContent=pointCount.toLocaleString('id-ID');
  document.getElementById('burnedTotal').textContent=hasHa ? `${areaHa.toLocaleString('id-ID',{maximumFractionDigits:2})} ha` : hasKm ? `${areaKm2.toLocaleString('id-ID',{maximumFractionDigits:2})} km²` : '—';
  document.getElementById('layerTotal').textContent=layerRegistry.size;
  document.getElementById('provinceTotal').textContent=provinces.size;
  updateChart();
}
function updateChart() {
  const counts = {};
  for (const rec of layerRegistry.values()) counts[rec.type] = (counts[rec.type]||0) + (rec.count || (rec.meta.raster ? 1 : 0));
  const labels=Object.keys(counts), values=Object.values(counts);
  document.getElementById('chartFallback').style.display=labels.length?'none':'block';
  const canvas=document.getElementById('compositionChart');
  if (chart) chart.destroy();
  if (labels.length && typeof Chart !== 'undefined') {
    chart=new Chart(canvas,{type:'doughnut',data:{labels,datasets:[{data:values,backgroundColor:['#38c995','#f09a48','#ed6259','#6c91c2','#b45b79','#8b9ca5'] ,borderColor:'#14232d',borderWidth:3}]},options:{responsive:true,plugins:{legend:{position:'bottom',labels:{color:'#b7c8cf',font:{size:10},padding:13}}}}});
  }
}
function refreshUI() {
  const list=document.getElementById('layerList'), summary=document.getElementById('summaryTable');
  if (!layerRegistry.size) {
    list.innerHTML='<div class="empty-state">Belum ada layer yang dimuat.</div>';
    summary.innerHTML='<tr><td colspan="4" class="empty-cell">Belum ada data.</td></tr>';
  } else {
    list.innerHTML='';
    summary.innerHTML='';
    for (const rec of layerRegistry.values()) {
      const row=document.createElement('div'); row.className='layer-row';
      const meta=document.createElement('div'); meta.className='layer-meta';
      meta.innerHTML=`<div class="layer-name">${escapeHtml(rec.name)}</div><div class="layer-sub">${escapeHtml(rec.type)} · ${rec.count ? rec.count.toLocaleString('id-ID')+' fitur' : 'raster/overlay'}</div>`;
      const actions=document.createElement('div'); actions.className='layer-actions';
      const label=document.createElement('label'); const check=document.createElement('input'); check.type='checkbox'; check.checked=rec.visible;
      check.addEventListener('change',()=>toggleLayer(rec.id,check.checked)); label.append(check,document.createTextNode(' Tampil'));
      const del=document.createElement('button'); del.className='delete-button'; del.textContent='Hapus'; del.addEventListener('click',()=>removeLayer(rec.id));
      actions.append(label,del); row.append(meta,actions); list.append(row);
      const tr=document.createElement('tr'); tr.innerHTML=`<td>${escapeHtml(rec.name)}</td><td>${escapeHtml(rec.type)}</td><td>${rec.count.toLocaleString('id-ID')}</td><td>${rec.visible?'Aktif':'Nonaktif'}</td>`; summary.append(tr);
    }
  }
  updateStats(); updateFilterOptions();
}
function toggleLayer(id, visible) {
  const rec=layerRegistry.get(id); if(!rec)return;
  rec.visible=visible;
  [maps.dashboard,maps.full].forEach((m,i)=>{
    const lyr=(i===1 && rec.secondLayer) ? rec.secondLayer : rec.layer;
    if(visible && !m.hasLayer(lyr)) lyr.addTo(m);
    if(!visible && m.hasLayer(lyr)) m.removeLayer(lyr);
  });
  refreshUI();
}
function removeLayer(id) {
  const rec=layerRegistry.get(id); if(!rec)return;
  [maps.dashboard,maps.full].forEach((m,i)=>{
    const lyr=(i===1 && rec.secondLayer) ? rec.secondLayer : rec.layer;
    if(m.hasLayer(lyr))m.removeLayer(lyr);
  });
  if(rec.meta.url)URL.revokeObjectURL(rec.meta.url);
  layerRegistry.delete(id); refreshUI(); notify('Layer dihapus dari sesi.');
}
function updateFilterOptions() {
  const years=new Set(), provinces=new Set();
  for(const rec of layerRegistry.values()) {
    (rec.meta.data?.features||[]).forEach(f=>{
      const p=f.properties||{};
      const y=normalizeKey(p,['tahun','year','acq_year']);
      const pr=normalizeKey(p,['provinsi','province','nama_provinsi']);
      if(y&&p[y])years.add(String(p[y]));
      if(pr&&p[pr])provinces.add(String(p[pr]));
    });
  }
  const fill=(id,values,placeholder)=>{
    const el=document.getElementById(id),old=el.value;
    el.innerHTML=`<option value="">${placeholder}</option>`;
    [...values].sort().forEach(v=>{const o=document.createElement('option');o.value=v;o.textContent=v;el.append(o);});
    if([...values].includes(old))el.value=old;
  };
  fill('yearFilter',years,'Semua tahun'); fill('provinceFilter',provinces,'Semua provinsi');
}
function applyFilters() {
  const year=document.getElementById('yearFilter').value;
  const province=document.getElementById('provinceFilter').value;
  for(const rec of layerRegistry.values()) {
    if(!rec.meta.data) continue;
    const filtered={type:'FeatureCollection',features:(rec.meta.data.features||[]).filter(f=>{
      const p=f.properties||{},yk=normalizeKey(p,['tahun','year','acq_year']),pk=normalizeKey(p,['provinsi','province','nama_provinsi']);
      return (!year || !yk || String(p[yk])===year) && (!province || !pk || String(p[pk])===province);
    })};
    [rec.layer,rec.secondLayer].filter(Boolean).forEach(target=>{
      target.clearLayers();
      makeGeoJSONLayer(filtered).eachLayer(l=>target.addLayer(l));
    });
    rec.count=filtered.features.length;
  }
  refreshUI(); notify('Filter diterapkan pada layer vektor.');
}
function resetFilters() {
  document.getElementById('yearFilter').value='';document.getElementById('provinceFilter').value='';
  for(const rec of layerRegistry.values()) {
    if(!rec.meta.data)continue;
    [rec.layer,rec.secondLayer].filter(Boolean).forEach(target=>{
      target.clearLayers();
      makeGeoJSONLayer(rec.meta.data).eachLayer(l=>target.addLayer(l));
    });
    rec.count=rec.meta.data.features?.length||0;
  }
  refreshUI();notify('Filter direset.');
}
function exportSummary() {
  const rows=[['Nama layer','Jenis','Jumlah fitur','Status']];
  for(const r of layerRegistry.values())rows.push([r.name,r.type,r.count,r.visible?'Aktif':'Nonaktif']);
  const csv=rows.map(row=>row.map(v=>`"${String(v).replace(/"/g,'""')}"`).join(',')).join('\n');
  const blob=new Blob([csv],{type:'text/csv;charset=utf-8;'});
  const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download='ringkasan_layer_karhutla.csv';a.click();URL.revokeObjectURL(url);
}
function navigate(section) {
  document.querySelectorAll('.page-section').forEach(s=>s.classList.toggle('active',s.id===section));
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.toggle('active',b.dataset.section===section));
  if(section==='mapview')setTimeout(()=>maps.full.invalidateSize(),100);
  if(section==='dashboard')setTimeout(()=>maps.dashboard.invalidateSize(),100);
}
document.querySelectorAll('.nav-item').forEach(b=>b.addEventListener('click',()=>navigate(b.dataset.section)));
document.querySelectorAll('[data-goto]').forEach(b=>b.addEventListener('click',()=>navigate(b.dataset.goto)));
document.getElementById('loadGeojson').addEventListener('click',async()=>{try{const f=document.getElementById('geojsonInput').files[0];if(!f)throw new Error('Pilih file GeoJSON terlebih dahulu.');await readGeoJSONFile(f);setUploadMessage('GeoJSON berhasil dimuat.');}catch(e){notify(e.message,true);setUploadMessage(e.message);}});
document.getElementById('loadShp').addEventListener('click',async()=>{try{const f=document.getElementById('shpInput').files[0];if(!f)throw new Error('Pilih ZIP Shapefile terlebih dahulu.');await readShapefile(f);setUploadMessage('Shapefile ZIP berhasil dimuat.');}catch(e){notify(e.message,true);setUploadMessage(e.message);}});
document.getElementById('loadTiff').addEventListener('click',async()=>{try{const f=document.getElementById('tiffInput').files[0];if(!f)throw new Error('Pilih GeoTIFF terlebih dahulu.');await readGeoTIFF(f);setUploadMessage('GeoTIFF berhasil dimuat.');}catch(e){notify(e.message,true);setUploadMessage(e.message);}});
document.getElementById('loadTable').addEventListener('click',async()=>{try{const f=document.getElementById('tableInput').files[0];if(!f)throw new Error('Pilih CSV atau Excel terlebih dahulu.');await readTable(f);setUploadMessage('Tabel koordinat berhasil dimuat.');}catch(e){notify(e.message,true);setUploadMessage(e.message);}});
document.getElementById('loadImage').addEventListener('click',async()=>{try{const f=document.getElementById('imageInput').files[0];if(!f)throw new Error('Pilih PNG/JPG terlebih dahulu.');await readImageOverlay(f);setUploadMessage('Image overlay berhasil ditambahkan.');}catch(e){notify(e.message,true);setUploadMessage(e.message);}});
document.getElementById('applyFilters').addEventListener('click',applyFilters);
document.getElementById('resetFilters').addEventListener('click',resetFilters);
document.getElementById('exportBtn').addEventListener('click',exportSummary);
document.getElementById('todayLabel').textContent=new Date().toLocaleDateString('id-ID',{day:'numeric',month:'long',year:'numeric'});
refreshUI();
