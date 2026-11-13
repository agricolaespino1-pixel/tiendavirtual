// ===== Tienda virtual · Google Apps Script =====
// 1) Crea una hoja de cálculo nueva, abre Extensiones > Apps Script y pega este código.
// 2) Implementar > Nueva implementación > Aplicación web
//    Ejecutar como: Yo   |   Quién tiene acceso: Cualquier persona
// 3) Copia la URL que termina en /exec y pégala en URL_SHEETS dentro de index.html.
const SS = SpreadsheetApp.getActive();
const sh = n => SS.getSheetByName(n) || SS.insertSheet(n);
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const sum = (a, f) => a.reduce((s, x) => s + f(x), 0);

function load() {
  const s = sh('_estado'), n = s.getLastRow();
  if (!n) return { cfg: {}, users: [{ u: 'admin', p: 'admin123', r: 'admin' }], prods: [], sales: [], orders: [], buys: [], exp: [], banners: [], promos: [], closes: [] };
  return JSON.parse(s.getRange(1, 1, n, 1).getValues().map(r => r[0]).join(''));
}

function store(st) {
  // Copia maestra (oculta): aquí viven las claves, por eso la hoja se oculta
  const s = sh('_estado'), j = JSON.stringify(st), rows = [];
  for (let i = 0; i < j.length; i += 40000) rows.push([j.slice(i, i + 40000)]);
  s.clearContents(); s.getRange(1, 1, rows.length, 1).setValues(rows);

  put('Productos', ['Producto', 'Categoría', 'Compra', 'Venta', 'Stock', 'Margen unit.', 'Valor a compra', 'Valor a venta'],
    st.prods.map(p => [p.n, p.cat, p.pc, p.pv, p.st, p.pv - p.pc, p.pc * p.st, p.pv * p.st]), true);
  put('Ventas', ['Fecha', 'Vendedor', 'Cliente', 'Pago', 'Detalle', 'Total', 'Costo', 'Ganancia', 'Mes'],
    st.sales.map(v => [v.f.slice(0, 16).replace('T', ' '), v.vd, v.cl || '', v.py, v.it.map(i => i.q + 'x ' + i.n).join('; '), v.tot, v.cs, v.tot - v.cs, v.f.slice(0, 7)]), true);
  put('Pedidos', ['Fecha', 'Cliente', 'Celular', 'Pago', 'Detalle', 'Total', 'Estado'],
    st.orders.map(o => [o.f.slice(0, 16).replace('T', ' '), o.cl, o.tl, o.py, o.it.map(i => i.q + 'x ' + i.n).join('; '), o.tot, o.es]), true);
  put('Compras', ['Fecha', 'Producto', 'Cantidad', 'Costo unit.', 'Total', 'Origen del dinero'],
    st.buys.map(b => [b.f.slice(0, 10), b.n, b.q, b.pc, b.tot, b.src === 'inv' ? 'Inversión' : 'Caja']));
  put('Gastos', ['Fecha', 'Concepto', 'Monto', 'Origen del dinero'],
    st.exp.map(e => [e.f.slice(0, 10), e.c, e.m, e.src === 'inv' ? 'Inversión' : 'Caja']));
  put('Usuarios', ['Usuario', 'Rol'], st.users.map(u => [u.u, u.r]));
  put('Promociones', ['Promoción', 'Precio combo', 'Productos', 'Costo'],
    st.promos.map(p => [p.n, p.pr, p.items.map(i => ((st.prods.find(x => x.id === i.id) || {}).n || '?') + ' x' + i.q).join(' + '),
      sum(p.items, i => ((st.prods.find(x => x.id === i.id) || {}).pc || 0) * i.q)]));
  put('Cierres de caja', ['Fecha', 'Vendedor', 'Periodo', 'Total vendido'],
    st.closes.map(c => [c.f.slice(0, 16).replace('T', ' '), c.vd, c.per, c.tot]));
  caja(st);
  try { pivots(); } catch (e) { Logger.log('Pivots: ' + e); }
  try { dashboard(); } catch (e) { Logger.log('Dashboard: ' + e); }
  try { sh('_estado').hideSheet(); } catch (e) {}
}

function put(name, head, rows, filter) {
  const s = sh(name);
  if (s.getFilter()) s.getFilter().remove();
  s.clearContents();
  s.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#0E5A47').setFontColor('#ffffff');
  if (rows.length) s.getRange(2, 1, rows.length, head.length).setValues(rows);
  s.setFrozenRows(1);
  if (filter) s.getRange(1, 1, Math.max(2, rows.length + 1), head.length).createFilter(); // filtros listos en cada columna
}

function caja(st) {
  const cap = +st.cfg.cap || 0, ing = sum(st.sales, s => s.tot),
    cb = sum(st.buys.filter(b => b.src !== 'inv'), b => b.tot), cg = sum(st.exp.filter(e => e.src !== 'inv'), e => e.m),
    inv = sum(st.buys.filter(b => b.src === 'inv'), b => b.tot) + sum(st.exp.filter(e => e.src === 'inv'), e => e.m),
    mc = sum(st.prods, p => p.pc * p.st), mv = sum(st.prods, p => p.pv * p.st),
    gan = ing - sum(st.sales, s => s.cs) - sum(st.exp, e => e.m);
  const rows = [['Capital inicial', cap], ['Ventas cobradas (efectivo + Yape)', ing], ['Compras pagadas con dinero de la tienda', -cb],
    ['Gastos pagados con dinero de la tienda', -cg], ['CAJA CHICA', cap + ing - cb - cg], ['Inversión aportada (otro dinero)', inv],
    ['Mercadería a precio de compra', mc], ['Mercadería a precio de venta', mv], ['Ganancia si se vende todo', mv - mc],
    ['Ganancia / pérdida acumulada', gan]];
  put('Caja', ['Concepto', 'Monto (S/)'], rows);
  sh('Caja').getRange('A6:B6').setFontWeight('bold').setBackground('#FFC83D');
}

// Tablas dinámicas reales (se crean una sola vez y se actualizan solas)
function pivots() {
  const S = SpreadsheetApp.PivotTableSummarizeFunction;
  const t1 = sh('Tabla dinámica productos');
  if (!t1.getPivotTables().length) {
    const p = sh('Productos'), pt = t1.getRange('A1').createPivotTable(p.getRange(1, 1, p.getMaxRows(), 8));
    pt.addRowGroup(2);  // Categoría
    pt.addRowGroup(1);  // Producto
    pt.addPivotValue(5, S.SUM).setDisplayName('Stock');
    pt.addPivotValue(7, S.SUM).setDisplayName('Valor a compra');
    pt.addPivotValue(8, S.SUM).setDisplayName('Valor a venta');
  }
  const t2 = sh('Tabla dinámica ventas');
  if (!t2.getPivotTables().length) {
    const v = sh('Ventas'), pt = t2.getRange('A1').createPivotTable(v.getRange(1, 1, v.getMaxRows(), 9));
    pt.addRowGroup(9);     // Mes
    pt.addRowGroup(2);     // Vendedor
    pt.addColumnGroup(4);  // Pago
    pt.addPivotValue(6, S.SUM).setDisplayName('Vendido');
    pt.addPivotValue(8, S.SUM).setDisplayName('Ganancia');
  }
}

function dashboard() {
  const s = sh('Dashboard');
  if (s.getRange('A1').getValue()) return;
  s.getRange('A1:D1').setValues([['Mes', 'Ingresos', 'Costo', 'Ganancia']]).setFontWeight('bold').setBackground('#0E5A47').setFontColor('#ffffff');
  s.getRange('A2').setFormula('=SORT(UNIQUE(FILTER(Ventas!I2:I,Ventas!I2:I<>"")))');
  s.getRange('B2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(Ventas!I:I,A2:A,Ventas!F:F)))');
  s.getRange('C2').setFormula('=ARRAYFORMULA(IF(A2:A="","",SUMIF(Ventas!I:I,A2:A,Ventas!G:G)))');
  s.getRange('D2').setFormula('=ARRAYFORMULA(IF(A2:A="","",B2:B-C2:C))');
  s.insertChart(s.newChart().setChartType(Charts.ChartType.COLUMN).addRange(s.getRange('A1:D13')).setPosition(6, 1, 0, 0).setOption('title', 'Ventas por mes').build());
}

// Lo que ve cualquier cliente (sin costos, sin usuarios, sin ventas)
function publico(st) {
  const c = st.cfg || {};
  return { cfg: { name: c.name, slogan: c.slogan, bannerTxt: c.bannerTxt, banner: c.banner, qr: c.qr, wa: c.wa, color: c.color, logo: c.logo },
    prods: (st.prods || []).map(p => ({ id: p.id, n: p.n, cat: p.cat, pv: p.pv, st: p.st, img: p.img })),
    banners: st.banners || [], promos: st.promos || [] };
}

const union = (inc, old) => { const ids = {}; inc.forEach(x => ids[x.id] = 1); return inc.concat(old.filter(x => !ids[x.id])); };

function doGet() { return out(publico(load())); }

function doPost(e) {
  const b = JSON.parse(e.postData.contents), lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const st = load();
    if (b.a === 'login') {
      const ok = st.users.some(x => x.u === b.u && x.p === b.p);
      return out(ok ? { ok: true, state: st } : { ok: false });
    }
    if (b.a === 'pedido') { st.orders = union([b.o], st.orders); store(st); return out({ ok: true }); }
    if (b.a === 'estado') {
      const me = st.users.find(x => x.u === b.u && x.p === b.p);
      if (!me) return out({ ok: false });
      const n = b.state;
      n.orders = union(n.orders || [], st.orders);   // nunca se pierden pedidos de clientes
      n.sales = union(n.sales || [], st.sales);      // ni ventas de otros vendedores
      n.closes = union(n.closes || [], st.closes || []);
      if (me.r !== 'admin') {                        // el vendedor no cambia usuarios, precios ni ajustes
        ['users', 'cfg', 'banners', 'promos', 'buys', 'exp'].forEach(k => n[k] = st[k]);
        const stock = {}; (b.state.prods || []).forEach(p => stock[p.id] = p.st);
        n.prods = st.prods.map(p => Object.assign({}, p, { st: p.id in stock ? stock[p.id] : p.st }));
      }
      store(n); return out({ ok: true, orders: n.orders });
    }
  } finally { lock.releaseLock(); }
  return out({ ok: false });
}
