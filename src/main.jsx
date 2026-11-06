import React, { useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';
import './style.css';

const SUPABASE_URL = import.meta.env.VITE_SUPABASE_URL;
const SUPABASE_KEY = import.meta.env.VITE_SUPABASE_PUBLISHABLE_KEY;
const supabase = SUPABASE_URL && SUPABASE_KEY ? createClient(SUPABASE_URL, SUPABASE_KEY) : null;
const CLOUD_TABLES = { items: 'inventory_items', surplus: 'inventory_surplus', movements: 'inventory_movements' };

const KEYS = { items: 'inventory-react-items-v1', surplus: 'inventory-react-surplus-v1', movements: 'inventory-react-movements-v1' };
const readStore = (key, fallback = []) => { try { const v = localStorage.getItem(key); return v ? JSON.parse(v) : fallback; } catch { return fallback; } };
const makeId = () => `${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
const fmt = n => Number(n || 0).toLocaleString('vi-VN');
function App() {
  const [page, setPage] = useState('dashboard');
  const [items, setItems] = useState(() => readStore(KEYS.items));
  const [surplus, setSurplus] = useState(() => readStore(KEYS.surplus));
  const [movements, setMovements] = useState(() => readStore(KEYS.movements));
  const [cloudReady, setCloudReady] = useState(false);
  const [query, setQuery] = useState('');
  const [notice, setNotice] = useState('');
  const [itemModal, setItemModal] = useState(null);
  const [surplusModal, setSurplusModal] = useState(null);
  const fileRef = useRef(null);
  const surplusFileRef = useRef(null);
  useEffect(() => {
    let alive = true;
    async function loadCloud() {
      if (!supabase) { setCloudReady(true); return; }
      try {
        const names = Object.keys(CLOUD_TABLES);
        const results = await Promise.all(names.map(k => supabase.from(CLOUD_TABLES[k]).select('data').eq('id', 'shared').maybeSingle()));
        const hasErrors = results.some(r => r.error);
        if (hasErrors) throw results.find(r => r.error).error;
        if (!alive) return;
        if (results[0].data?.data) setItems(results[0].data.data);
        if (results[1].data?.data) setSurplus(results[1].data.data);
        if (results[2].data?.data) setMovements(results[2].data.data);
        setCloudReady(true);
        setNotice('Đã kết nối dữ liệu kho online.');
      } catch (err) { console.error(err); if (alive) { setCloudReady(true); setNotice('Không tải được dữ liệu online. Kiểm tra URL, key và quyền Supabase.'); } }
    }
    loadCloud();
    return () => { alive = false; };
  }, []);
  useEffect(() => { localStorage.setItem(KEYS.items, JSON.stringify(items)); if (cloudReady && supabase) supabase.from(CLOUD_TABLES.items).upsert({ id: 'shared', data: items }).then(({error}) => { if (error) console.error('Cloud items save failed:', error); }); }, [items, cloudReady]);
  useEffect(() => { localStorage.setItem(KEYS.surplus, JSON.stringify(surplus)); if (cloudReady && supabase) supabase.from(CLOUD_TABLES.surplus).upsert({ id: 'shared', data: surplus }).then(({error}) => { if (error) console.error('Cloud surplus save failed:', error); }); }, [surplus, cloudReady]);
  useEffect(() => { localStorage.setItem(KEYS.movements, JSON.stringify(movements)); if (cloudReady && supabase) supabase.from(CLOUD_TABLES.movements).upsert({ id: 'shared', data: movements }).then(({error}) => { if (error) console.error('Cloud movements save failed:', error); }); }, [movements, cloudReady]);
  useEffect(() => { if (!supabase || !cloudReady) return; const timer = setInterval(async () => { try { const names = Object.keys(CLOUD_TABLES); const results = await Promise.all(names.map(k => supabase.from(CLOUD_TABLES[k]).select('data').eq('id', 'shared').maybeSingle())); if (results.some(r => r.error)) return; if (results[0].data?.data) setItems(prev => JSON.stringify(prev) === JSON.stringify(results[0].data.data) ? prev : results[0].data.data); if (results[1].data?.data) setSurplus(prev => JSON.stringify(prev) === JSON.stringify(results[1].data.data) ? prev : results[1].data.data); if (results[2].data?.data) setMovements(prev => JSON.stringify(prev) === JSON.stringify(results[2].data.data) ? prev : results[2].data.data); } catch (err) { console.error(err); } }, 5000); return () => clearInterval(timer); }, [cloudReady]);
  useEffect(() => { if (!notice) return; const t = setTimeout(() => setNotice(''), 2800); return () => clearTimeout(t); }, [notice]);

  const totalStock = items.reduce((s, x) => s + Number(x.stock || 0), 0);
  const filteredItems = useMemo(() => items.filter(x => `${x.code} ${x.name} ${x.specification}`.toLowerCase().includes(query.toLowerCase())), [items, query]);
  const filteredSurplus = useMemo(() => surplus.filter(x => `${x.code} ${x.name} ${x.note || ''}`.toLowerCase().includes(query.toLowerCase())), [surplus, query]);

  function saveItem(form) {
    const code = form.code.trim();
    if (!code || !form.name.trim()) return setNotice('Vui lòng nhập mã hàng và tên hàng.');
    const duplicate = items.some(x => x.code.toLowerCase() === code.toLowerCase() && x.id !== form.id);
    if (duplicate) return setNotice('Mã hàng đã tồn tại.');
    if (form.id) setItems(prev => prev.map(x => x.id === form.id ? { ...form, code, name: form.name.trim(), stock: Number(form.stock || 0) } : x));
    else setItems(prev => [...prev, { ...form, id: makeId(), code, name: form.name.trim(), stock: Number(form.stock || 0) }]);
    setItemModal(null); setNotice('Đã lưu thông tin hàng hóa.');
  }
  function deleteItem(item) {
    if (!confirm(`Xóa hàng hóa ${item.code} - ${item.name}?`)) return;
    setItems(prev => prev.filter(x => x.id !== item.id)); setNotice('Đã xóa hàng hóa.');
  }
  function saveSurplus(form) {
    if (!form.code.trim() || !form.name.trim()) return setNotice('Vui lòng nhập mã liệu và tên liệu.');
    const actualIssued = Math.max(0, Number(form.actualIssued || 0));
    const slipQty = Math.max(0, Number(form.slipQty || 0));
    const excess = Math.max(0, actualIssued - slipQty);
    const entry = { ...form, actualIssued, slipQty, excess, issued: actualIssued, returned: excess, date: form.date || new Date().toISOString().slice(0,10) };
    if (form.id) setSurplus(prev => prev.map(x => x.id === form.id ? entry : x));
    else setSurplus(prev => [{ ...entry, id: makeId() }, ...prev]);
    setSurplusModal(null); setNotice('Đã lưu phiếu hàng dư.');
  }
  function exportExcel(type) {
    let data = type === 'items' ? items.map(({id, ...x}) => x) : type === 'surplus' ? surplus.map(({id, ...x}) => x) : movements.map(({id, ...x}) => x);
    if (!data.length) return setNotice('Chưa có dữ liệu để xuất.');
    const ws = XLSX.utils.json_to_sheet(data); const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, type === 'items' ? 'HangHoa' : type === 'surplus' ? 'HangDu' : 'LichSu');
    XLSX.writeFile(wb, `${type === 'items' ? 'danh-sach-hang-hoa' : type === 'surplus' ? 'hang-du' : 'lich-su-kho'}.xlsx`);
  }
  function importExcel(file, type) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array' });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: '' });
        if (!rows.length) return setNotice('File Excel không có dữ liệu.');
        const normalize = row => Object.fromEntries(Object.entries(row).map(([k,v]) => [String(k).trim().toLowerCase().replace(/[ _-]/g,''), v]));
        const normalized = rows.map(normalize);
        if (type === 'items') {
          const mapped = normalized.map(r => ({ id: makeId(), code: String(r.code ?? r.mahang ?? '').trim(), name: String(r.name ?? r.tenhang ?? '').trim(), specification: String(r.specification ?? r.quycach ?? ''), unit: String(r.unit ?? r.donvi ?? ''), stock: Number(r.stock ?? r.tonkho ?? 0) })).filter(x => x.code && x.name);
          if (!mapped.length) return setNotice('Không tìm thấy cột mã hàng/code và tên hàng/name.');
          setItems(prev => { const map = new Map(prev.map(x => [x.code.toLowerCase(), x])); mapped.forEach(x => map.set(x.code.toLowerCase(), { ...map.get(x.code.toLowerCase()), ...x, id: map.get(x.code.toLowerCase())?.id || x.id })); return [...map.values()]; });
          setNotice(`Đã nhập ${mapped.length} dòng hàng hóa.`);
        } else {
          const mapped = normalized.map(r => ({ id: makeId(), date: String(r.date ?? r.ngay ?? new Date().toISOString().slice(0,10)), code: String(r.code ?? r.mahang ?? '').trim(), name: String(r.name ?? r.tenhang ?? '').trim(), issued: Number(r.issued ?? r.soluongxuat ?? 0), excess: Number(r.excess ?? r.soluongdu ?? 0), returned: Number(r.returned ?? r.soluongtrave ?? 0), note: String(r.note ?? r.ghichu ?? ''), received: false })).filter(x => x.code && x.name);
          if (!mapped.length) return setNotice('Không tìm thấy cột mã hàng/code và tên hàng/name.');
          setSurplus(prev => [...mapped, ...prev]); setNotice(`Đã nhập ${mapped.length} dòng hàng dư.`);
        }
      } catch { setNotice('Không đọc được file Excel. Hãy thử lưu file dưới dạng .xlsx.'); }
    };
    reader.readAsArrayBuffer(file);
  }
  function backup() {
    const blob = new Blob([JSON.stringify({ items, surplus, movements, exportedAt: new Date().toISOString() }, null, 2)], {type:'application/json'});
    const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = `sao-luu-kho-${new Date().toISOString().slice(0,10)}.json`; a.click(); URL.revokeObjectURL(a.href); setNotice('Đã tạo file sao lưu.');
  }
  function restore(file) {
    if (!file) return; const reader = new FileReader(); reader.onload = e => { try { const d = JSON.parse(e.target.result); if (!Array.isArray(d.items) || !Array.isArray(d.surplus)) throw new Error(); if (!confirm('Khôi phục sẽ thay thế dữ liệu hiện tại. Tiếp tục?')) return; setItems(d.items); setSurplus(d.surplus); setMovements(d.movements || []); setNotice('Đã khôi phục dữ liệu.'); } catch { setNotice('File sao lưu không hợp lệ.'); } }; reader.readAsText(file);
  }

  const nav = [{id:'dashboard',label:'Tổng quan',icon:'▦'},{id:'items',label:'Danh mục hàng hóa',icon:'▤'},{id:'surplus',label:'Theo dõi hàng dư',icon:'↻'},{id:'movements',label:'Lịch sử kho',icon:'☷'}];
  return <div className="app-shell">
    <aside className="sidebar"><div className="brand"><div className="brand-mark">K</div><div><strong>Kho của tôi</strong><small>QUẢN LÝ KHO HÀNG</small></div></div><div className="nav-label">MENU CHÍNH</div><nav>{nav.map(n=><button key={n.id} className={page===n.id?'nav-item active':'nav-item'} onClick={()=>{setPage(n.id);setQuery('')}}><span className="nav-icon">{n.icon}</span>{n.label}</button>)}</nav><div className="sidebar-bottom"><div className="local-note"><span className="status-dot"/> {supabase ? (cloudReady ? 'Dữ liệu online' : 'Đang kết nối...') : 'Chưa cấu hình online'}</div><button className="nav-item" onClick={backup}><span className="nav-icon">⇩</span>Sao lưu dữ liệu</button><label className="nav-item file-label"><span className="nav-icon">⇧</span>Khôi phục dữ liệu<input type="file" accept=".json,application/json" hidden onChange={e=>{restore(e.target.files[0]);e.target.value=''}}/></label></div></aside>
    <main className="main"><header className="topbar"><div><div className="eyebrow">HỆ THỐNG QUẢN LÝ</div><h1>{nav.find(n=>n.id===page)?.label}</h1></div><div className="top-right"><span className="live-pill"><i/> Đang hoạt động</span><div className="avatar">QL</div></div></header>
      {page==='dashboard' && <section className="content"><div className="welcome"><div><span className="welcome-kicker">BẢNG ĐIỀU KHIỂN</span><h2>Chào mừng trở lại! 👋</h2><p>Theo dõi hàng hóa, tồn kho và các phiếu hàng dư tại một nơi.</p></div><div className="welcome-art">▥</div></div><div className="stats-grid"><Stat title="Mặt hàng" value={fmt(items.length)} icon="▤" tone="blue" detail="Mã hàng đang quản lý"/><Stat title="Tổng tồn kho" value={fmt(totalStock)} icon="▦" tone="green" detail="Tổng số lượng hiện có"/><Stat title="Phiếu hàng dư" value={fmt(surplus.length)} icon="↻" tone="orange" detail="Tổng phiếu đã ghi nhận"/><Stat title="Chưa nhập trả" value={fmt(surplus.filter(x=>!x.received && x.returned>0).length)} icon="◷" tone="purple" detail="Phiếu cần cập nhật tồn kho"/></div><div className="dashboard-grid"><div className="panel"><div className="panel-heading"><div><h3>Hàng hóa gần đây</h3><p>Các mặt hàng mới trong danh mục</p></div><button className="text-button" onClick={()=>setPage('items')}>Xem tất cả →</button></div>{items.length ? <div className="table-wrap"><table><thead><tr><th>Mã hàng</th><th>Tên hàng</th><th>Quy cách</th><th className="right">Tồn kho</th></tr></thead><tbody>{items.slice(-5).reverse().map(x=><tr key={x.id}><td><span className="code-pill">{x.code}</span></td><td className="primary-cell">{x.name}</td><td>{x.specification || '—'}</td><td className="right"><b>{fmt(x.stock)}</b> {x.unit}</td></tr>)}</tbody></table></div> : <Empty text="Chưa có hàng hóa. Hãy thêm mặt hàng đầu tiên." action="Thêm hàng hóa" onClick={()=>setItemModal({})}/>}</div><div className="panel quick-panel"><div className="panel-heading"><div><h3>Thao tác nhanh</h3><p>Các chức năng thường dùng</p></div></div><button className="quick-action" onClick={()=>setItemModal({})}><span className="quick-icon blue">＋</span><span><b>Thêm hàng hóa</b><small>Tạo mã hàng mới</small></span><span className="arrow">→</span></button><button className="quick-action" onClick={()=>setSurplusModal({})}><span className="quick-icon orange">↻</span><span><b>Tạo phiếu hàng dư</b><small>Ghi nhận hàng xuất dư</small></span><span className="arrow">→</span></button><button className="quick-action" onClick={()=>exportExcel('items')}><span className="quick-icon green">⇩</span><span><b>Xuất Excel hàng hóa</b><small>Tải danh sách hiện tại</small></span><span className="arrow">→</span></button><div className="tip-box"><b>💡 Mẹo sử dụng</b><p>Hãy sao lưu dữ liệu định kỳ để tránh mất dữ liệu khi xóa dữ liệu trình duyệt.</p></div></div></div></section>}
      {page==='items' && <section className="content"><div className="toolbar"><div className="searchbox"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Tìm theo mã hàng, tên hàng..."/></div><div className="toolbar-actions"><button className="btn secondary" onClick={()=>fileRef.current?.click()}>⇧ Nhập Excel</button><input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>{importExcel(e.target.files[0],'items');e.target.value=''}}/><button className="btn secondary" onClick={()=>exportExcel('items')}>⇩ Xuất Excel</button><button className="btn primary" onClick={()=>setItemModal({})}>＋ Thêm hàng hóa</button></div></div><div className="panel table-panel"><div className="panel-heading"><div><h3>Danh sách hàng hóa</h3><p>Đang hiển thị {filteredItems.length} / {items.length} mặt hàng</p></div></div><div className="table-wrap"><table><thead><tr><th>#</th><th>Mã hàng</th><th>Tên hàng</th><th>Quy cách</th><th>Đơn vị</th><th className="right">Tồn kho</th><th className="right">Thao tác</th></tr></thead><tbody>{filteredItems.map((x,i)=><tr key={x.id}><td className="muted">{i+1}</td><td><span className="code-pill">{x.code}</span></td><td className="primary-cell">{x.name}</td><td>{x.specification||'—'}</td><td>{x.unit||'—'}</td><td className="right"><b>{fmt(x.stock)}</b></td><td className="right actions"><button title="Sửa" onClick={()=>setItemModal({...x})}>✎</button><button title="Xóa" onClick={()=>deleteItem(x)}>×</button></td></tr>)}</tbody></table>{!filteredItems.length&&<Empty text={items.length?'Không tìm thấy hàng hóa phù hợp.':'Chưa có dữ liệu hàng hóa.'} action="Thêm hàng hóa" onClick={()=>setItemModal({})}/>}</div></div><p className="footnote">Định dạng Excel nhập: <b>code, name, specification, unit, stock</b> hoặc các cột tiếng Việt tương đương (mahang, tenhang, quycach, donvi, tonkho).</p></section>}
      {page==='surplus' && <section className="content"><div className="toolbar"><div className="searchbox"><span>⌕</span><input value={query} onChange={e=>setQuery(e.target.value)} placeholder="Tìm theo mã hàng, tên hàng..."/></div><div className="toolbar-actions"><button className="btn secondary" onClick={()=>surplusFileRef.current?.click()}>⇧ Nhập Excel</button><input ref={surplusFileRef} type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>{importExcel(e.target.files[0],'surplus');e.target.value=''}}/><button className="btn secondary" onClick={()=>exportExcel('surplus')}>⇩ Xuất Excel</button><button className="btn primary" onClick={()=>setSurplusModal({})}>＋ Tạo phiếu hàng dư</button></div></div><div className="panel table-panel"><div className="panel-heading"><div><h3>Danh sách phiếu hàng dư</h3><p>So sánh số lượng phát thực tế với số lượng trên phiếu</p></div></div><div className="table-wrap"><table><thead><tr><th>Ngày</th><th>Mã hàng</th><th>Tên hàng</th><th className="right">SL phát thực tế</th><th className="right">SL trên phiếu</th><th className="right">Còn thừa</th><th className="right">Thao tác</th></tr></thead><tbody>{filteredSurplus.map(x=><tr key={x.id}><td>{x.date||'—'}</td><td><span className="code-pill">{x.code}</span></td><td className="primary-cell">{x.name}</td><td className="right">{fmt(x.actualIssued ?? x.issued)}</td><td className="right">{fmt(x.slipQty ?? 0)}</td><td className="right"><b>{fmt(x.excess ?? Math.max(0, Number(x.issued||0)-Number(x.slipQty||0)))}</b></td><td className="right actions"><button title="Sửa" onClick={()=>setSurplusModal({...x})}>✎</button><button title="Xóa" onClick={()=>{if(confirm('Xóa phiếu hàng dư này?'))setSurplus(p=>p.filter(y=>y.id!==x.id))}}>×</button></td></tr>)}</tbody></table>{!filteredSurplus.length&&<Empty text="Chưa có phiếu hàng dư." action="Tạo phiếu hàng dư" onClick={()=>setSurplusModal({})}/>}</div></div><p className="footnote">Số lượng còn thừa được tự động tính bằng: số lượng phát thực tế − số lượng trên phiếu (không nhỏ hơn 0).</p></section>}
      {page==='movements' && <section className="content"><div className="toolbar"><div><h3 className="section-title">Lịch sử biến động kho</h3><p className="sub-title">Các lần nhập trả hàng dư đã cộng vào tồn kho</p></div><button className="btn secondary" onClick={()=>exportExcel('movements')}>⇩ Xuất Excel</button></div><div className="panel table-panel"><div className="table-wrap"><table><thead><tr><th>Thời gian</th><th>Mã hàng</th><th>Tên hàng</th><th>Loại giao dịch</th><th className="right">Số lượng</th><th>Ghi chú</th></tr></thead><tbody>{movements.map(x=><tr key={x.id}><td>{new Date(x.date).toLocaleString('vi-VN')}</td><td><span className="code-pill">{x.code}</span></td><td className="primary-cell">{x.name}</td><td><span className="badge success">{x.type}</span></td><td className="right"><b>+{fmt(x.qty)}</b></td><td>{x.note}</td></tr>)}</tbody></table>{!movements.length&&<Empty text="Chưa có giao dịch kho nào."/>}</div></div></section>}
      <footer className="footer">Quản lý kho hàng <span>•</span> {supabase ? 'Dữ liệu đồng bộ online khoảng mỗi 5 giây' : 'Chưa cấu hình kết nối cơ sở dữ liệu online'}</footer>
    </main>
    {itemModal && <ItemDialog initial={itemModal} onClose={()=>setItemModal(null)} onSave={saveItem}/>}
    {surplusModal && <SurplusDialog initial={surplusModal} items={items} onClose={()=>setSurplusModal(null)} onSave={saveSurplus}/>}
    {notice&&<div className="toast">✓ {notice}</div>}
  </div>;
}
function Stat({title,value,icon,tone,detail}) { return <div className="stat-card"><div className={`stat-icon ${tone}`}>{icon}</div><div className="stat-title">{title}</div><div className="stat-value">{value}</div><div className="stat-detail">{detail}</div></div>; }
function Empty({text,action,onClick}) { return <div className="empty"><div className="empty-icon">▤</div><p>{text}</p>{action&&<button className="btn primary" onClick={onClick}>{action}</button>}</div>; }
function ItemDialog({initial,onClose,onSave}) { const [form,setForm]=useState({code:'',name:'',specification:'',unit:'',stock:0,...initial}); const set=(k,v)=>setForm(p=>({...p,[k]:v})); return <Modal title={initial.id?'Cập nhật hàng hóa':'Thêm hàng hóa mới'} onClose={onClose}><form onSubmit={e=>{e.preventDefault();onSave(form)}}><div className="form-grid"><Field label="Mã hàng *"><input required value={form.code} onChange={e=>set('code',e.target.value)} placeholder="VD: HH001"/></Field><Field label="Tên hàng *"><input required value={form.name} onChange={e=>set('name',e.target.value)} placeholder="Nhập tên hàng"/></Field><Field label="Quy cách"><input value={form.specification||''} onChange={e=>set('specification',e.target.value)} placeholder="VD: 10 x 20 cm"/></Field><Field label="Đơn vị tính"><input value={form.unit||''} onChange={e=>set('unit',e.target.value)} placeholder="VD: cái, hộp, kg"/></Field><Field label="Tồn kho ban đầu" wide><input type="number" min="0" step="any" value={form.stock??0} onChange={e=>set('stock',e.target.value)}/></Field></div><div className="modal-actions"><button type="button" className="btn secondary" onClick={onClose}>Hủy</button><button className="btn primary" type="submit">Lưu hàng hóa</button></div></form></Modal>; }
function SurplusDialog({initial,items,onClose,onSave}) {
  const [form,setForm]=useState({date:new Date().toISOString().slice(0,10),code:'',name:'',actualIssued:0,slipQty:0,excess:0,note:'',...initial});
  const set=(k,v)=>setForm(p=>{const next={...p,[k]:v}; if(k==='actualIssued'||k==='slipQty') next.excess=Math.max(0,Number(k==='actualIssued'?v:next.actualIssued||0)-Number(k==='slipQty'?v:next.slipQty||0)); return next;});
  function pickCode(code) { const item=items.find(x=>x.code===code); setForm(p=>({...p,code,name:item?.name||p.name})); }
  const excess=Math.max(0,Number(form.actualIssued||0)-Number(form.slipQty||0));
  return <Modal title={initial.id?'Cập nhật phiếu hàng dư':'Tạo phiếu hàng dư'} onClose={onClose}><form onSubmit={e=>{e.preventDefault();onSave({...form,excess})}}><div className="form-grid">
    <Field label="Mã liệu *"><input required list="item-codes" value={form.code} onChange={e=>pickCode(e.target.value)} placeholder="Nhập hoặc chọn mã liệu"/><datalist id="item-codes">{items.map(x=><option key={x.id} value={x.code}>{x.name}</option>)}</datalist></Field>
    <Field label="Tên liệu *" wide><input required value={form.name} onChange={e=>set('name',e.target.value)} placeholder="Tên liệu"/></Field>
    <Field label="Số lượng phát thực tế"><input type="number" min="0" step="any" value={form.actualIssued??0} onChange={e=>set('actualIssued',e.target.value)}/></Field>
    <Field label="Số lượng trên phiếu"><input type="number" min="0" step="any" value={form.slipQty??0} onChange={e=>set('slipQty',e.target.value)}/></Field>
    <Field label="Số lượng còn thừa (tự tính)" wide><input type="number" value={excess} readOnly/><small className="field-help">Còn thừa = số lượng phát thực tế − số lượng trên phiếu</small></Field>
  </div><div className="modal-actions"><button type="button" className="btn secondary" onClick={onClose}>Hủy</button><button className="btn primary" type="submit">Lưu phiếu</button></div></form></Modal>;
}
function Field({label,children,wide}) { return <label className={`field ${wide?'wide':''}`}><span>{label}</span>{children}</label>; }
function Modal({title,children,onClose}) { return <div className="modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)onClose()}}><div className="modal"><div className="modal-head"><h2>{title}</h2><button className="close" onClick={onClose}>×</button></div>{children}</div></div>; }

createRoot(document.getElementById('root')).render(<App/>);
