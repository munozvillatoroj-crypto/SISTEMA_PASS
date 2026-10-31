/**
 * Sistema PAS Alimentos — Versión Excel (sin XAMPP)
 * Datos en localStorage + importar/exportar Excel
 * Reportes HTML descargables + reporte mensual automático el último día del mes
 */
(function () {
  "use strict";

  // ========== CONSTANTES ==========
  const STORAGE_KEY = "pas_alimentos_v1";
  const SESSION_KEY = "pas_session";
  const MESES = [
    "", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
  ];

  // ========== ESTADO ==========
  let DB = {
    usuarios: [{ id: 1, username: "admin", password: "admin123", nombre: "Administrador", rol: "Administrador" }],
    categorias: [
      { id: 1, nombre: "Granos Básicos" },
      { id: 2, nombre: "Lácteos" },
      { id: 3, nombre: "Carnes" },
      { id: 4, nombre: "Verduras" },
      { id: 5, nombre: "Frutas" },
      { id: 6, nombre: "Aceites y Grasas" },
      { id: 7, nombre: "Bebidas" },
      { id: 8, nombre: "Otros" }
    ],
    proveedores: [{ id: 1, nombre: "Proveedor General", telefono: "", email: "" }],
    productos: [],
    entradas: [],
    consumos: [],
    bitacora: [],
    meta: { lastMonthlyReport: null }
  };

  let session = null;
  let charts = { mov: null, cat: null };

  // ========== PERSISTENCIA ==========
  function loadDB() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        DB = { ...DB, ...parsed };
        if (!DB.usuarios || !DB.usuarios.length) {
          DB.usuarios = [{ id: 1, username: "admin", password: "admin123", nombre: "Administrador", rol: "Administrador" }];
        }
      }
    } catch (e) {
      console.warn("No se pudo cargar datos locales", e);
    }
  }

  function saveDB() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(DB));
    } catch (e) {
      toast("Error al guardar datos locales. Espacio insuficiente.", "err");
    }
  }

  function nextId(arr) {
    if (!arr || !arr.length) return 1;
    return Math.max(...arr.map((x) => x.id || 0)) + 1;
  }

  function logBitacora(accion) {
    DB.bitacora.unshift({
      id: nextId(DB.bitacora),
      usuario: session ? session.username : "sistema",
      accion,
      fecha_hora: new Date().toISOString(),
      ip: "local"
    });
    if (DB.bitacora.length > 500) DB.bitacora = DB.bitacora.slice(0, 500);
    saveDB();
  }

  // ========== UTILIDADES ==========
  function $(sel) { return document.querySelector(sel); }
  function $all(sel) { return document.querySelectorAll(sel); }
  function today() {
    const d = new Date();
    return d.toISOString().slice(0, 10);
  }
  function fmtDate(iso) {
    if (!iso) return "-";
    const d = new Date(iso);
    if (isNaN(d)) return iso;
    return d.toLocaleString("es-SV", { day: "2-digit", month: "2-digit", year: "numeric", hour: "2-digit", minute: "2-digit" });
  }
  function fmtNum(n) {
    return Number(n || 0).toLocaleString("es-SV", { maximumFractionDigits: 2 });
  }
  function toast(msg, type) {
    const el = $("#toast");
    el.textContent = msg;
    el.className = "toast show " + (type || "ok");
    setTimeout(() => el.classList.remove("show"), 3200);
  }
  function lastDayOfMonth(y, m) {
    return new Date(y, m, 0).getDate(); // m is 1-12
  }
  function isLastDayOfMonth(dateStr) {
    const d = dateStr ? new Date(dateStr + "T12:00:00") : new Date();
    const last = lastDayOfMonth(d.getFullYear(), d.getMonth() + 1);
    return d.getDate() === last;
  }

  // ========== AUTH ==========
  function tryLogin(user, pass) {
    const u = DB.usuarios.find((x) => x.username === user && x.password === pass);
    if (!u) return false;
    session = { id: u.id, username: u.username, nombre: u.nombre, rol: u.rol };
    sessionStorage.setItem(SESSION_KEY, JSON.stringify(session));
    logBitacora("Inicio de sesión");
    return true;
  }

  function logout() {
    if (session) logBitacora("Cierre de sesión");
    session = null;
    sessionStorage.removeItem(SESSION_KEY);
    showLogin();
  }

  function restoreSession() {
    try {
      const s = sessionStorage.getItem(SESSION_KEY);
      if (s) session = JSON.parse(s);
    } catch (e) {}
  }

  function showLogin() {
    $("#loginScreen").style.display = "flex";
    $("#app").classList.remove("show");
  }

  function showApp() {
    $("#loginScreen").style.display = "none";
    $("#app").classList.add("show");
    $("#userLabel").textContent = session.nombre || session.username;
    navigate("dashboard");
    checkMonthlyAutoReport();
  }


  // ========== MENÚ RESPONSIVE ==========
  function openSidebar() {
    const sb = $("#sidebar");
    const ov = $("#sidebarOverlay");
    if (sb) sb.classList.add("open");
    if (ov) ov.classList.add("show");
    document.body.style.overflow = "hidden";
  }
  function closeSidebar() {
    const sb = $("#sidebar");
    const ov = $("#sidebarOverlay");
    if (sb) sb.classList.remove("open");
    if (ov) ov.classList.remove("show");
    document.body.style.overflow = "";
  }
  function toggleSidebar() {
    const sb = $("#sidebar");
    if (sb && sb.classList.contains("open")) closeSidebar();
    else openSidebar();
  }

  // ========== NAVEGACIÓN ==========
  function navigate(view) {
    closeSidebar();
    $all(".view").forEach((v) => v.classList.remove("active"));
    $all(".nav-item").forEach((n) => n.classList.remove("active"));
    const el = $("#view-" + view);
    if (el) el.classList.add("active");
    const nav = document.querySelector('.nav-item[data-view="' + view + '"]');
    if (nav) nav.classList.add("active");
    const titles = {
      dashboard: "Dashboard",
      entradas: "Ingresar Entradas",
      consumo: "Consumo Diario",
      bodega: "Existencia Bodega",
      proveedores: "Proveedores",
      reportes: "Reportes de Auditoría",
      bitacora: "Bitácora",
      datos: "Base de datos (Excel)"
    };
    $("#pageTitle").textContent = titles[view] || view;
    if (view === "dashboard") renderDashboard();
    if (view === "bodega") renderBodega();
    if (view === "entradas") renderEntradasForm();
    if (view === "consumo") renderConsumoForm();
    if (view === "proveedores") renderProveedores();
    if (view === "reportes") renderReportes();
    if (view === "bitacora") renderBitacora();
    if (view === "datos") renderDatos();
  }

  // ========== DASHBOARD ==========
  function renderDashboard() {
    const prods = DB.productos || [];
    const stockBajo = prods.filter((p) => Number(p.stock_actual) <= Number(p.stock_minimo)).length;
    const vencidos = prods.filter((p) => p.fecha_vencimiento && p.fecha_vencimiento <= today()).length;
    $("#kpiProductos").textContent = prods.length;
    $("#kpiStockBajo").textContent = stockBajo;
    $("#kpiVencidos").textContent = vencidos;
    $("#kpiUsuarios").textContent = (DB.usuarios || []).length;

    // Banner último día del mes
    const banner = $("#monthBanner");
    if (isLastDayOfMonth()) {
      banner.style.display = "flex";
      banner.className = "banner banner-warn";
      banner.innerHTML = "<strong>📅 Hoy es el último día del mes.</strong> Se recomienda generar el <b>reporte mensual</b> de auditoría desde el menú Reportes.";
    } else {
      banner.style.display = "none";
    }

    // Charts
    const anio = new Date().getFullYear();
    const entradasMes = Array(12).fill(0);
    const salidasMes = Array(12).fill(0);
    (DB.entradas || []).forEach((e) => {
      const d = new Date(e.fecha + "T12:00:00");
      if (d.getFullYear() === anio) entradasMes[d.getMonth()] += Number(e.cantidad) || 0;
    });
    (DB.consumos || []).forEach((c) => {
      const d = new Date(c.fecha + "T12:00:00");
      if (d.getFullYear() === anio) salidasMes[d.getMonth()] += Number(c.cantidad) || 0;
    });

    const labels = MESES.slice(1);
    if (charts.mov) charts.mov.destroy();
    charts.mov = new Chart($("#chartMovimientos"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          { label: "Entradas", data: entradasMes, backgroundColor: "#2563EB" },
          { label: "Consumos", data: salidasMes, backgroundColor: "#059669" }
        ]
      },
      options: {
        responsive: true,
        plugins: { legend: { position: "bottom" } },
        scales: { y: { beginAtZero: true } }
      }
    });

    const catMap = {};
    (DB.categorias || []).forEach((c) => { catMap[c.id] = { nombre: c.nombre, total: 0 }; });
    prods.forEach((p) => {
      if (catMap[p.categoria_id]) catMap[p.categoria_id].total++;
    });
    const catLabels = Object.values(catMap).map((c) => c.nombre);
    const catData = Object.values(catMap).map((c) => c.total);
    if (charts.cat) charts.cat.destroy();
    charts.cat = new Chart($("#chartCategorias"), {
      type: "doughnut",
      data: {
        labels: catLabels,
        datasets: [{
          data: catData,
          backgroundColor: ["#2563EB", "#059669", "#D97706", "#DC2626", "#8B5CF6", "#06B6D4", "#EC4899", "#64748B"]
        }]
      },
      options: { responsive: true, plugins: { legend: { position: "bottom" } } }
    });
  }

  // ========== BODEGA ==========
  function renderBodega() {
    const catSel = $("#prodCategoria");
    const provSel = $("#prodProveedor");
    catSel.innerHTML = (DB.categorias || []).map((c) => `<option value="${c.id}">${c.nombre}</option>`).join("");
    provSel.innerHTML = (DB.proveedores || []).map((p) => `<option value="${p.id}">${p.nombre}</option>`).join("");

    const tbody = $("#tablaProductos tbody");
    const prods = [...(DB.productos || [])].sort((a, b) => (b.id || 0) - (a.id || 0));
    if (!prods.length) {
      tbody.innerHTML = `<tr><td colspan="7" class="empty">Sin productos. Registre el primero a la izquierda.</td></tr>`;
      return;
    }
    tbody.innerHTML = prods.map((p) => {
      const bajo = Number(p.stock_actual) <= Number(p.stock_minimo);
      const badge = bajo
        ? `<span class="badge badge-low">Stock Bajo</span>`
        : `<span class="badge badge-ok">Disponible</span>`;
      return `<tr>
        <td>${esc(p.codigo)}</td>
        <td>${esc(p.nombre)}</td>
        <td>${fmtNum(p.stock_actual)}</td>
        <td>${fmtNum(p.stock_minimo)}</td>
        <td>${esc(p.unidad_medida)}</td>
        <td>${p.fecha_vencimiento || "-"}</td>
        <td>${badge}</td>
      </tr>`;
    }).join("");
  }

  function guardarProducto(e) {
    e.preventDefault();
    const codigo = $("#prodCodigo").value.trim();
    const nombre = $("#prodNombre").value.trim();
    if (!codigo || !nombre) { toast("Código y nombre son obligatorios", "err"); return; }
    if ((DB.productos || []).some((p) => p.codigo.toLowerCase() === codigo.toLowerCase())) {
      toast("El código ya existe. Use otro.", "err"); return;
    }
    const prod = {
      id: nextId(DB.productos),
      codigo,
      nombre,
      categoria_id: Number($("#prodCategoria").value) || 1,
      proveedor_id: Number($("#prodProveedor").value) || 1,
      unidad_medida: $("#prodUnidad").value || "Unidad",
      stock_actual: Number($("#prodStock").value) || 0,
      stock_minimo: Number($("#prodMinimo").value) || 5,
      fecha_vencimiento: $("#prodVence").value || null,
      fecha_ingreso: today()
    };
    DB.productos.push(prod);
    logBitacora("Se registró un nuevo alimento: " + nombre + " (Código: " + codigo + ")");
    saveDB();
    $("#formProducto").reset();
    toast("Producto guardado correctamente");
    renderBodega();
  }

  // ========== ENTRADAS ==========
  function renderEntradasForm() {
    const sel = $("#entProducto");
    sel.innerHTML = (DB.productos || []).map((p) =>
      `<option value="${p.id}">${esc(p.codigo)} — ${esc(p.nombre)} (Stock: ${fmtNum(p.stock_actual)})</option>`
    ).join("") || `<option value="">Sin productos</option>`;
    $("#entFecha").value = today();
  }

  function guardarEntrada(e) {
    e.preventDefault();
    const pid = Number($("#entProducto").value);
    const cant = Number($("#entCantidad").value);
    const fecha = $("#entFecha").value || today();
    const vence = $("#entVence").value || null;
    if (!pid || cant <= 0) { toast("Seleccione producto y cantidad mayor a 0", "err"); return; }
    const prod = DB.productos.find((p) => p.id === pid);
    if (!prod) { toast("Producto no encontrado", "err"); return; }

    DB.entradas.push({
      id: nextId(DB.entradas),
      producto_id: pid,
      cantidad: cant,
      fecha,
      usuario: session.username,
      fecha_registro: new Date().toISOString()
    });
    prod.stock_actual = Number(prod.stock_actual) + cant;
    if (vence) prod.fecha_vencimiento = vence;
    logBitacora("Entrada de " + cant + " unidades del producto " + prod.nombre);
    saveDB();
    toast("Entrada registrada. Nuevo stock: " + fmtNum(prod.stock_actual));
    $("#formEntrada").reset();
    renderEntradasForm();
  }

  // ========== CONSUMO ==========
  function renderConsumoForm() {
    const sel = $("#conProducto");
    sel.innerHTML = (DB.productos || []).map((p) =>
      `<option value="${p.id}">${esc(p.codigo)} — ${esc(p.nombre)} (Stock: ${fmtNum(p.stock_actual)})</option>`
    ).join("") || `<option value="">Sin productos</option>`;
    $("#conFecha").value = today();
  }

  function guardarConsumo(e) {
    e.preventDefault();
    const pid = Number($("#conProducto").value);
    const cant = Number($("#conCantidad").value);
    const fecha = $("#conFecha").value || today();
    if (!pid || cant <= 0) { toast("Seleccione producto y cantidad mayor a 0", "err"); return; }
    const prod = DB.productos.find((p) => p.id === pid);
    if (!prod) { toast("Producto no encontrado", "err"); return; }
    if (Number(prod.stock_actual) < cant) {
      toast("Stock insuficiente. Disponible: " + fmtNum(prod.stock_actual), "err");
      return;
    }
    DB.consumos.push({
      id: nextId(DB.consumos),
      producto_id: pid,
      cantidad: cant,
      fecha,
      usuario: session.username,
      fecha_registro: new Date().toISOString()
    });
    prod.stock_actual = Number(prod.stock_actual) - cant;
    logBitacora("Consumo de " + cant + " unidades de " + prod.nombre);
    saveDB();
    toast("Consumo registrado. Stock restante: " + fmtNum(prod.stock_actual));
    $("#formConsumo").reset();
    renderConsumoForm();
  }

  // ========== PROVEEDORES ==========
  function renderProveedores() {
    const tbody = $("#tablaProveedores tbody");
    const list = DB.proveedores || [];
    if (!list.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty">Sin proveedores</td></tr>`;
      return;
    }
    tbody.innerHTML = list.map((p) =>
      `<tr><td>${p.id}</td><td>${esc(p.nombre)}</td><td>${esc(p.telefono || "-")}</td><td>${esc(p.email || "-")}</td></tr>`
    ).join("");
  }

  function guardarProveedor(e) {
    e.preventDefault();
    const nombre = $("#provNombre").value.trim();
    if (!nombre) { toast("Nombre obligatorio", "err"); return; }
    DB.proveedores.push({
      id: nextId(DB.proveedores),
      nombre,
      telefono: $("#provTel").value.trim(),
      email: $("#provEmail").value.trim()
    });
    logBitacora("Se registró proveedor: " + nombre);
    saveDB();
    $("#formProveedor").reset();
    toast("Proveedor guardado");
    renderProveedores();
  }

  // ========== BITÁCORA ==========
  function renderBitacora() {
    const tbody = $("#tablaBitacora tbody");
    const logs = (DB.bitacora || []).slice(0, 100);
    if (!logs.length) {
      tbody.innerHTML = `<tr><td colspan="4" class="empty">Sin movimientos registrados</td></tr>`;
      return;
    }
    tbody.innerHTML = logs.map((l) =>
      `<tr>
        <td><small><b>${fmtDate(l.fecha_hora)}</b></small></td>
        <td><span class="badge badge-dark">${esc(l.usuario)}</span></td>
        <td>${esc(l.accion)}</td>
        <td><code>${esc(l.ip || "-")}</code></td>
      </tr>`
    ).join("");
  }

  // ========== REPORTES ==========
  function renderReportes() {
    const anio = new Date().getFullYear();
    const mesActual = new Date().getMonth() + 1;
    const grid = $("#monthsGrid");
    grid.innerHTML = MESES.slice(1).map((nombre, i) => {
      const m = i + 1;
      const cur = m === mesActual ? " current" : "";
      return `<div class="month-card${cur}" data-mes="${m}" data-anio="${anio}">
        <div class="ico">📄</div>
        <div class="name">${nombre}</div>
        <div class="year">${anio}</div>
        ${m === mesActual ? '<span class="badge badge-ok" style="margin-top:4px">Actual</span>' : ""}
      </div>`;
    }).join("");
    grid.querySelectorAll(".month-card").forEach((card) => {
      card.addEventListener("click", () => {
        generarReporte("mensual", null, Number(card.dataset.mes), Number(card.dataset.anio));
      });
    });
    $("#repFechaDiaria").value = today();
  }

  function nombreProducto(id) {
    const p = (DB.productos || []).find((x) => x.id === id);
    return p ? p.nombre : "—";
  }
  function codigoProducto(id) {
    const p = (DB.productos || []).find((x) => x.id === id);
    return p ? p.codigo : "—";
  }

  function generarReporte(tipo, fecha, mes, anio) {
    const now = new Date();
    anio = anio || now.getFullYear();
    mes = mes || (now.getMonth() + 1);
    fecha = fecha || today();

    let entradas = [], consumos = [], bitacora = [];
    let tituloPeriodo = "", nombreArchivo = "";

    if (tipo === "diario") {
      entradas = (DB.entradas || []).filter((e) => e.fecha === fecha);
      consumos = (DB.consumos || []).filter((c) => c.fecha === fecha);
      bitacora = (DB.bitacora || []).filter((b) => (b.fecha_hora || "").slice(0, 10) === fecha);
      tituloPeriodo = "Día " + fecha.split("-").reverse().join("/");
      nombreArchivo = "Auditoria_Diaria_" + fecha + ".html";
    } else {
      entradas = (DB.entradas || []).filter((e) => {
        const d = new Date(e.fecha + "T12:00:00");
        return d.getMonth() + 1 === mes && d.getFullYear() === anio;
      });
      consumos = (DB.consumos || []).filter((c) => {
        const d = new Date(c.fecha + "T12:00:00");
        return d.getMonth() + 1 === mes && d.getFullYear() === anio;
      });
      tituloPeriodo = MESES[mes] + " " + anio;
      nombreArchivo = "Reporte_Mensual_" + MESES[mes] + "_" + anio + ".html";
    }

    const prods = DB.productos || [];
    const totalEnt = entradas.reduce((s, e) => s + Number(e.cantidad), 0);
    const totalCon = consumos.reduce((s, c) => s + Number(c.cantidad), 0);
    const totalStock = prods.reduce((s, p) => s + Number(p.stock_actual), 0);
    const stockBajo = prods.filter((p) => Number(p.stock_actual) <= Number(p.stock_minimo)).length;

    const rowsProd = prods.map((p, i) => {
      const bajo = Number(p.stock_actual) <= Number(p.stock_minimo);
      return `<tr>
        <td>${i + 1}</td><td>${esc(p.codigo)}</td><td>${esc(p.nombre)}</td>
        <td>${fmtNum(p.stock_actual)}</td><td>${fmtNum(p.stock_minimo)}</td>
        <td>${esc(p.unidad_medida)}</td><td>${p.fecha_vencimiento || "-"}</td>
        <td style="color:${bajo ? "#dc2626" : "#16a34a"};font-weight:700">${bajo ? "⚠ Stock Bajo" : "Disponible"}</td>
      </tr>`;
    }).join("") || `<tr><td colspan="8" style="text-align:center;color:#64748b">Sin productos</td></tr>`;

    const rowsEnt = entradas.map((e, i) =>
      `<tr><td>${i + 1}</td><td>${esc(codigoProducto(e.producto_id))}</td><td>${esc(nombreProducto(e.producto_id))}</td>
       <td>${fmtNum(e.cantidad)}</td><td>${e.fecha}</td><td>${esc(e.usuario || "-")}</td></tr>`
    ).join("") || `<tr><td colspan="6" style="text-align:center;color:#64748b">No hubo entradas en este período</td></tr>`;

    const rowsCon = consumos.map((c, i) =>
      `<tr><td>${i + 1}</td><td>${esc(codigoProducto(c.producto_id))}</td><td>${esc(nombreProducto(c.producto_id))}</td>
       <td>${fmtNum(c.cantidad)}</td><td>${c.fecha}</td><td>${esc(c.usuario || "-")}</td></tr>`
    ).join("") || `<tr><td colspan="6" style="text-align:center;color:#64748b">No hubo consumos en este período</td></tr>`;

    let bitHtml = "";
    if (tipo === "diario" && bitacora.length) {
      const rowsBit = bitacora.map((b, i) =>
        `<tr><td>${i + 1}</td><td>${fmtDate(b.fecha_hora)}</td><td>${esc(b.usuario)}</td><td>${esc(b.accion)}</td></tr>`
      ).join("");
      bitHtml = `<h3>📋 Bitácora del día</h3><table><thead><tr><th>#</th><th>Hora</th><th>Usuario</th><th>Acción</th></tr></thead><tbody>${rowsBit}</tbody></table>`;
    }

    const html = `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>Reporte ${tipo} — PAS Alimentos</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f1f5f9;margin:0;padding:24px;color:#1e293b}
.card{background:#fff;border-radius:12px;box-shadow:0 4px 20px rgba(0,0,0,.08);padding:28px;max-width:1100px;margin:0 auto}
h1{text-align:center;color:#0f172a;margin:0 0 4px;font-size:1.7rem}
h2{text-align:center;color:#2563eb;margin:0 0 18px;font-size:1.15rem}
.meta{display:flex;flex-wrap:wrap;gap:12px;margin-bottom:20px;font-size:.92rem}
.meta span{background:#f8fafc;padding:8px 14px;border-radius:8px;border:1px solid #e2e8f0}
.kpi{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;margin-bottom:24px}
.kpi-item{background:#0f172a;color:#fff;padding:14px;border-radius:10px;text-align:center}
.kpi-item .num{font-size:1.5rem;font-weight:700}
.kpi-item .lbl{font-size:.75rem;opacity:.8;text-transform:uppercase}
h3{background:#2563eb;color:#fff;padding:10px 14px;border-radius:8px 8px 0 0;margin:0;font-size:.95rem}
table{width:100%;border-collapse:collapse;margin-bottom:24px;font-size:.88rem}
th{background:#0f172a;color:#fff;padding:9px 7px;text-align:left}
td{padding:7px;border:1px solid #e2e8f0}
tr:nth-child(even){background:#f8fafc}
.footer{text-align:center;margin-top:28px;color:#64748b;font-size:.82rem;border-top:1px solid #e2e8f0;padding-top:14px}
.btn-print{display:block;margin:18px auto 0;padding:12px 28px;background:#2563eb;color:#fff;border:none;border-radius:8px;font-size:1rem;cursor:pointer}
@media print{.btn-print{display:none}body{background:#fff;padding:0}.card{box-shadow:none}}
</style></head><body><div class="card">
<h1>SISTEMA PAS — Alimentos</h1>
<h2>Reporte ${tipo.toUpperCase()} de Auditoría<br>${tituloPeriodo}</h2>
<div class="meta">
<span><b>Generado:</b> ${fmtDate(new Date().toISOString())}</span>
<span><b>Usuario:</b> ${esc(session ? session.username : "-")}</span>
<span><b>Tipo:</b> ${tipo}</span>
</div>
<div class="kpi">
<div class="kpi-item"><div class="num">${prods.length}</div><div class="lbl">Productos</div></div>
<div class="kpi-item"><div class="num">${fmtNum(totalEnt)}</div><div class="lbl">Entradas</div></div>
<div class="kpi-item"><div class="num">${fmtNum(totalCon)}</div><div class="lbl">Consumos</div></div>
<div class="kpi-item"><div class="num">${fmtNum(totalStock)}</div><div class="lbl">Stock Total</div></div>
<div class="kpi-item" style="background:#b45309"><div class="num">${stockBajo}</div><div class="lbl">Stock Bajo</div></div>
</div>
<h3>📦 Inventario Actual</h3>
<table><thead><tr><th>#</th><th>Código</th><th>Producto</th><th>Stock</th><th>Mínimo</th><th>Unidad</th><th>Vence</th><th>Estado</th></tr></thead>
<tbody>${rowsProd}</tbody></table>
<h3>📥 Entradas del período</h3>
<table><thead><tr><th>#</th><th>Código</th><th>Producto</th><th>Cantidad</th><th>Fecha</th><th>Usuario</th></tr></thead>
<tbody>${rowsEnt}</tbody></table>
<h3>📤 Consumos / Salidas del período</h3>
<table><thead><tr><th>#</th><th>Código</th><th>Producto</th><th>Cantidad</th><th>Fecha</th><th>Usuario</th></tr></thead>
<tbody>${rowsCon}</tbody></table>
${bitHtml}
<div class="footer"><b>Sistema PAS Alimentos — Escuela</b><br>
Documento generado automáticamente para auditoría · ${fmtDate(new Date().toISOString())}</div>
<button class="btn-print" onclick="window.print()">🖨 Imprimir / Guardar PDF</button>
</div></body></html>`;

    // Descargar
    const blob = new Blob([html], { type: "text/html;charset=utf-8" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = nombreArchivo;
    a.click();
    URL.revokeObjectURL(a.href);

    logBitacora("Generó reporte " + tipo + " (" + tituloPeriodo + ")");
    if (tipo === "mensual") {
      DB.meta = DB.meta || {};
      DB.meta.lastMonthlyReport = anio + "-" + String(mes).padStart(2, "0");
      saveDB();
    }
    toast("Reporte descargado: " + nombreArchivo);
  }

  function checkMonthlyAutoReport() {
    if (!isLastDayOfMonth()) return;
    const now = new Date();
    const key = now.getFullYear() + "-" + String(now.getMonth() + 1).padStart(2, "0");
    if (DB.meta && DB.meta.lastMonthlyReport === key) return;
    // Mostrar aviso (no forzar descarga automática para no molestar, pero sí avisar)
    setTimeout(() => {
      if (confirm("Hoy es el último día del mes (" + MESES[now.getMonth() + 1] + ").\n\n¿Desea generar y descargar ahora el reporte mensual de auditoría?\n\nIncluye todo lo que entró y salió del mes.")) {
        generarReporte("mensual", null, now.getMonth() + 1, now.getFullYear());
      }
    }, 800);
  }

  // ========== EXCEL (base de datos) ==========
  function renderDatos() {
    $("#statProductos").textContent = (DB.productos || []).length;
    $("#statEntradas").textContent = (DB.entradas || []).length;
    $("#statConsumos").textContent = (DB.consumos || []).length;
    $("#statBitacora").textContent = (DB.bitacora || []).length;
  }

  function exportarExcel() {
    if (typeof XLSX === "undefined") {
      toast("Librería Excel no cargada. Revise su conexión a internet (CDN).", "err");
      return;
    }
    const wb = XLSX.utils.book_new();

    // Productos
    const prodRows = (DB.productos || []).map((p) => ({
      id: p.id,
      codigo: p.codigo,
      nombre: p.nombre,
      categoria_id: p.categoria_id,
      proveedor_id: p.proveedor_id,
      unidad_medida: p.unidad_medida,
      stock_actual: p.stock_actual,
      stock_minimo: p.stock_minimo,
      fecha_vencimiento: p.fecha_vencimiento || "",
      fecha_ingreso: p.fecha_ingreso || ""
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(prodRows.length ? prodRows : [{ id: "" }]), "Productos");

    // Entradas
    const entRows = (DB.entradas || []).map((e) => ({
      id: e.id,
      producto_id: e.producto_id,
      producto: nombreProducto(e.producto_id),
      cantidad: e.cantidad,
      fecha: e.fecha,
      usuario: e.usuario || "",
      fecha_registro: e.fecha_registro || ""
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(entRows.length ? entRows : [{ id: "" }]), "Entradas");

    // Consumos
    const conRows = (DB.consumos || []).map((c) => ({
      id: c.id,
      producto_id: c.producto_id,
      producto: nombreProducto(c.producto_id),
      cantidad: c.cantidad,
      fecha: c.fecha,
      usuario: c.usuario || "",
      fecha_registro: c.fecha_registro || ""
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(conRows.length ? conRows : [{ id: "" }]), "Consumos");

    // Proveedores
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(DB.proveedores || [{ id: 1, nombre: "Proveedor General" }]), "Proveedores");

    // Categorías
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(DB.categorias || []), "Categorias");

    // Bitácora
    const bitRows = (DB.bitacora || []).map((b) => ({
      id: b.id,
      usuario: b.usuario,
      accion: b.accion,
      fecha_hora: b.fecha_hora,
      ip: b.ip || ""
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(bitRows.length ? bitRows : [{ id: "" }]), "Bitacora");

    // Usuarios (sin contraseñas en export sensible - include for restore)
    const userRows = (DB.usuarios || []).map((u) => ({
      id: u.id, username: u.username, password: u.password, nombre: u.nombre, rol: u.rol
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(userRows), "Usuarios");

    const nombre = "PAS_Inventario_" + today() + ".xlsx";
    XLSX.writeFile(wb, nombre);
    logBitacora("Exportó base de datos a Excel: " + nombre);
    toast("Excel descargado: " + nombre);
  }

  function importarExcel(file) {
    if (!file) return;
    if (typeof XLSX === "undefined") {
      toast("Librería Excel no cargada.", "err");
      return;
    }
    const reader = new FileReader();
    reader.onload = function (ev) {
      try {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: "array" });

        function sheetToJson(name) {
          const sh = wb.Sheets[name];
          if (!sh) return null;
          return XLSX.utils.sheet_to_json(sh, { defval: "" });
        }

        const prods = sheetToJson("Productos");
        const ents = sheetToJson("Entradas");
        const cons = sheetToJson("Consumos");
        const provs = sheetToJson("Proveedores");
        const cats = sheetToJson("Categorias");
        const bits = sheetToJson("Bitacora");
        const users = sheetToJson("Usuarios");

        if (prods && prods.length && prods[0].codigo !== undefined) {
          DB.productos = prods.filter((r) => r.codigo).map((r) => ({
            id: Number(r.id) || nextId(DB.productos),
            codigo: String(r.codigo),
            nombre: String(r.nombre || ""),
            categoria_id: Number(r.categoria_id) || 1,
            proveedor_id: Number(r.proveedor_id) || 1,
            unidad_medida: String(r.unidad_medida || "Unidad"),
            stock_actual: Number(r.stock_actual) || 0,
            stock_minimo: Number(r.stock_minimo) || 5,
            fecha_vencimiento: r.fecha_vencimiento || null,
            fecha_ingreso: r.fecha_ingreso || today()
          }));
        }
        if (ents && ents.length && ents[0].producto_id !== undefined) {
          DB.entradas = ents.filter((r) => r.producto_id).map((r) => ({
            id: Number(r.id) || 1,
            producto_id: Number(r.producto_id),
            cantidad: Number(r.cantidad) || 0,
            fecha: String(r.fecha || today()),
            usuario: String(r.usuario || ""),
            fecha_registro: r.fecha_registro || ""
          }));
        }
        if (cons && cons.length && cons[0].producto_id !== undefined) {
          DB.consumos = cons.filter((r) => r.producto_id).map((r) => ({
            id: Number(r.id) || 1,
            producto_id: Number(r.producto_id),
            cantidad: Number(r.cantidad) || 0,
            fecha: String(r.fecha || today()),
            usuario: String(r.usuario || ""),
            fecha_registro: r.fecha_registro || ""
          }));
        }
        if (provs && provs.length) {
          DB.proveedores = provs.filter((r) => r.nombre).map((r) => ({
            id: Number(r.id) || 1,
            nombre: String(r.nombre),
            telefono: String(r.telefono || ""),
            email: String(r.email || "")
          }));
        }
        if (cats && cats.length) {
          DB.categorias = cats.filter((r) => r.nombre).map((r) => ({
            id: Number(r.id) || 1,
            nombre: String(r.nombre)
          }));
        }
        if (bits && bits.length) {
          DB.bitacora = bits.filter((r) => r.accion).map((r) => ({
            id: Number(r.id) || 1,
            usuario: String(r.usuario || ""),
            accion: String(r.accion),
            fecha_hora: r.fecha_hora || new Date().toISOString(),
            ip: String(r.ip || "local")
          }));
        }
        if (users && users.length) {
          DB.usuarios = users.filter((r) => r.username).map((r) => ({
            id: Number(r.id) || 1,
            username: String(r.username),
            password: String(r.password || "1234"),
            nombre: String(r.nombre || r.username),
            rol: String(r.rol || "Bodeguero")
          }));
        }

        saveDB();
        logBitacora("Importó base de datos desde Excel: " + file.name);
        toast("Datos importados desde Excel correctamente");
        renderDatos();
        navigate("dashboard");
      } catch (err) {
        console.error(err);
        toast("Error al leer el Excel. Use un archivo exportado por este sistema.", "err");
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function esc(s) {
    if (s == null) return "";
    return String(s)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  // ========== INIT ==========
  function bindEvents() {
    $("#formLogin").addEventListener("submit", function (e) {
      e.preventDefault();
      const u = $("#loginUser").value.trim();
      const p = $("#loginPass").value;
      if (tryLogin(u, p)) {
        $("#loginError").style.display = "none";
        showApp();
      } else {
        $("#loginError").style.display = "block";
      }
    });

    $all(".nav-item[data-view]").forEach((n) => {
      n.addEventListener("click", () => navigate(n.dataset.view));
    });
    $("#btnLogout").addEventListener("click", logout);

    const btnMenu = $("#btnMenu");
    const overlay = $("#sidebarOverlay");
    if (btnMenu) btnMenu.addEventListener("click", toggleSidebar);
    if (overlay) overlay.addEventListener("click", closeSidebar);
    window.addEventListener("resize", function () {
      if (window.innerWidth > 768) closeSidebar();
    });

    $("#formProducto").addEventListener("submit", guardarProducto);
    $("#formEntrada").addEventListener("submit", guardarEntrada);
    $("#formConsumo").addEventListener("submit", guardarConsumo);
    $("#formProveedor").addEventListener("submit", guardarProveedor);

    $("#btnReporteHoy").addEventListener("click", () => {
      generarReporte("diario", $("#repFechaDiaria").value || today());
    });
    $("#btnReporteDiario").addEventListener("click", () => {
      generarReporte("diario", $("#repFechaDiaria").value || today());
    });

    $("#btnExportExcel").addEventListener("click", exportarExcel);
    $("#btnImportExcel").addEventListener("click", () => $("#fileExcel").click());
    $("#fileExcel").addEventListener("change", function () {
      if (this.files && this.files[0]) {
        if (confirm("Importar reemplazará los datos actuales con el contenido del Excel. ¿Continuar?")) {
          importarExcel(this.files[0]);
        }
        this.value = "";
      }
    });
    $("#btnBackupAhora").addEventListener("click", exportarExcel);
  }

  function init() {
    loadDB();
    restoreSession();
    bindEvents();
    if (session) showApp();
    else showLogin();
  }

  document.addEventListener("DOMContentLoaded", init);
})();
