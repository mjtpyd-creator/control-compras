"use strict";

const $ = (s) => document.querySelector(s);
const MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio",
               "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];

const VISTAS = { resumen: "#vista-resumen", compras: "#vista-compras", nueva: "#vista-nueva" };
let vistaActual = "resumen";
let listasData = null;
let fotoDataUrl = null;

// ---------------------------------------------------------------- utilidades
async function api(path, options) {
  const resp = await fetch(path, options);
  let datos = {};
  try { datos = await resp.json(); } catch (_) { /* sin cuerpo */ }
  if (!resp.ok || datos.ok === false) {
    throw new Error(datos.error || ("Error " + resp.status));
  }
  return datos;
}

function moneda(n) {
  const v = Number(n || 0);
  return v.toLocaleString("es-AR", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function aviso(texto, clase) {
  const el = $("#aviso");
  el.textContent = texto;
  el.className = "aviso " + (clase || "");
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.add("oculta"), 3500);
}

function ir(vista) {
  vistaActual = vista;
  for (const [k, sel] of Object.entries(VISTAS)) {
    $(sel).classList.toggle("oculta", k !== vista);
  }
  document.querySelectorAll(".tab").forEach((t) =>
    t.classList.toggle("activo", t.dataset.vista === vista));
  if (vista === "resumen") cargarResumen();
  if (vista === "compras") cargarCompras();
  if (vista === "nueva") cargarNueva();
  if (vista === "prediccion") cargarPrediccion();
}

// ---------------------------------------------------------------- resumen
async function cargarResumen() {
  const el = $("#vista-resumen");
  el.innerHTML = '<div class="vacio">Cargando…</div>';
  try {
    const r = await api("/api/resumen");
    const sub = $("#subtitulo");
    if (sub) sub.textContent = `Ejercicio ${r.ejercicio} · ARS`;
    const mesActual = new Date().getMonth(); // 0..11
    const mesNombre = MESES[mesActual];
    const gastoMes = r.por_mes && r.por_mes[mesNombre] ? r.por_mes[mesNombre] : 0;
    const maxMes = Math.max(1, ...Object.values(r.por_mes || {}));

    let html = `<div class="kpis">
      <div class="kpi destacado"><div class="valor moneda">$${moneda(gastoMes)}</div><div class="etiqueta">Este mes (${mesNombre})</div></div>
      <div class="kpi"><div class="valor moneda">$${moneda(r.gasto_total)}</div><div class="etiqueta">Ejercicio ${r.ejercicio}</div></div>
      <div class="kpi"><div class="valor">${r.tickets}</div><div class="etiqueta">Tickets</div></div>
      <div class="kpi"><div class="valor moneda">$${moneda(r.iva_incluido)}</div><div class="etiqueta">IVA incluido</div></div>
    </div>`;

    let bloqueMes = "";
    if (r.por_mes && Object.keys(r.por_mes).length) {
      bloqueMes = `<div class="tarjeta"><h2>Gasto por mes</h2><div class="barras">`;
      for (const mes of MESES) {
        const v = r.por_mes[mes] || 0;
        const h = Math.round((v / maxMes) * 90);
        bloqueMes += `<div class="barra"><div class="col" style="height:${h}px" title="${mes}: $${moneda(v)}"></div><div class="mes">${mes.slice(0, 3)}</div></div>`;
      }
      bloqueMes += `</div></div>`;
    }

    let bloqueCat = "";
    if (r.por_categoria && Object.keys(r.por_categoria).length) {
      bloqueCat = `<div class="tarjeta"><h2>Por categoría</h2>`;
      for (const [cat, v] of Object.entries(r.por_categoria)) {
        bloqueCat += `<div class="fila"><div><div class="nombre">${cat}</div></div><div class="monto">$${moneda(v)}</div></div>`;
      }
      bloqueCat += `</div>`;
    }
    if (bloqueMes || bloqueCat) html += `<div class="grid-dos">${bloqueMes}${bloqueCat}</div>`;

    const avisos = (r.problemas || []).map((p) => ({ t: p, warn: false }))
      .concat((r.avisos || []).map((a) => ({ t: a, warn: true })));
    if (avisos.length) {
      html += `<div class="tarjeta"><h2>A revisar (${avisos.length})</h2><div class="aviso-lista">`;
      for (const a of avisos) html += `<div class="item ${a.warn ? "warn" : ""}">• ${a.t}</div>`;
      html += `</div></div>`;
    }
    if (!r.tickets) html += `<div class="vacio">Todavía no hay compras cargadas.<br>Usá “＋ Nueva” para registrar la primera.</div>`;

    el.innerHTML = html;
  } catch (e) {
    el.innerHTML = `<div class="vacio">No se pudo cargar: ${e.message}</div>`;
  }
}

// ---------------------------------------------------------------- compras
async function cargarCompras() {
  const el = $("#vista-compras");
  el.innerHTML = '<div class="vacio">Cargando…</div>';
  try {
    const lista = await api("/api/tickets");
    if (!lista.length) {
      el.innerHTML = '<div class="vacio">No hay compras cargadas.</div>';
      return;
    }
    let html = "";
    for (const t of lista) {
      const chip = t.estado === "OK"
        ? '<span class="chip ok">OK</span>'
        : '<span class="chip revisar">REVISAR</span>';
      html += `<div class="ticket">
        <div class="cab" data-toggle="${t.id}">
          <div class="izq">
            <div class="comercio">${t.comercio}</div>
            <div class="meta">${t.fecha} · ${t.forma_pago}${t.numero && t.numero !== "-" ? " · " + t.numero : ""}</div>
          </div>
          <div class="der"><div class="total">$${moneda(t.total)}</div>${chip}</div>
        </div>
        <div class="lineas oculta" data-lineas="${t.id}">`;
      for (const l of t.lineas) {
        html += `<div class="linea"><span class="desc">${l.descripcion}</span><span class="monto">$${moneda(l.importe)}</span></div>`;
      }
      html += `</div></div>`;
    }
    el.innerHTML = html;
    document.querySelectorAll("[data-toggle]").forEach((cab) => {
      cab.addEventListener("click", () => {
        const id = cab.dataset.toggle;
        document.querySelector(`[data-lineas="${id}"]`).classList.toggle("oculta");
      });
    });
  } catch (e) {
    el.innerHTML = `<div class="vacio">No se pudo cargar: ${e.message}</div>`;
  }
}

// ---------------------------------------------------------------- prediccion
// Predicción de compras para días específicos (ej: martes y sábado)
async function cargarPrediccion() {
  const el = $("#vista-prediccion");
  el.innerHTML = '<div class="vacio">Cargando…</div>';
  try {
    const r = await api("/api/prediccion", {
      method: "GET",
      params: {
        dias_semana: [1, 5],  // martes (1) y sábado (5)
        horizonte: 7
      }
    });
    const s = r;
    if (!s) {
      el.innerHTML = '<div class="vacio">No se pudieron calcular las predicciones.</div>';
      return;
    }
    el.innerHTML = `
      <div class="tarjeta">
        <h2>🔮 Previsiones para Martes y Sábado</h2>
        <div class="kpis">
          <div class="kpi destacado">
            <div class="valor moneda">$${moneda(s["hoy"])}</div>
            <div class="etiqueta">Ejercicio ${s["hoy"]}</div>
          </div>
          <div class="kpi">
            <div class="valor moneda">$${moneda(s["proximamente"]["dias_semana"][0])}</div>
            <div class="etiqueta">Próxima compra (martes)</div>
          </div>
          <div class="kpi">
            <div class="valor moneda">$${moneda(s["proximadamente"]["dias_semana"][1])}</div>
            <div class="etiqueta">Próxima compra (sábado)</div>
          </div>
        </div>
        <div class="tarjeta">
          <h2>📅 Martes</h2>
          ${JSON.stringify(s["por_dia"][1]?.products || [])}
        </div>
        <div class="tarjeta">
          <h2>📅 Sábado</h2>
          ${JSON.stringify(s["por_dia"][5]?.products || [])}
        </div>
      </div>`;
  } catch (e) {
    el.innerHTML = `<div class="vacio">Error al cargar predicciones: ${e.message}</div>`;
  }
}

// ---------------------------------------------------------------- nueva compra
async function cargarNueva() {
  const el = $("#vista-nueva");
  if (!listasData) {
    try { listasData = await api("/api/listas"); }
    catch (e) { listasData = { categorias: [], formas_pago: [], unidades: [] }; }
  }
  const hoy = new Date().toISOString().slice(0, 10);
  fotoDataUrl = null;

  el.innerHTML = `
    <div class="grid-nueva">
      <div class="tarjeta">
        <h2>Nueva compra</h2>
        <div class="fila-dos">
          <div class="campo"><label>Fecha</label><input id="c-fecha" type="date" value="${hoy}"></div>
          <div class="campo"><label>Total</label><input id="c-total" type="number" inputmode="decimal" step="0.01" min="0" placeholder="0,00"></div>
        </div>
        <div class="campo"><label>Comercio</label><input id="c-comercio" type="text" placeholder="Supermercado, kiosco…"></div>
        <div class="fila-dos">
          <div class="campo"><label>Categoría</label><select id="c-categoria"></select></div>
          <div class="campo"><label>Forma de pago</label><select id="c-pago"></select></div>
        </div>
        <div class="campo"><label>Nº de ticket (opcional)</label><input id="c-numero" type="text" placeholder="—"></div>
        <div class="campo"><label>Notas (opcional)</label><input id="c-notas" type="text"></div>
        <div class="campo"><label>Pegá texto de IA (opcional)</label><textarea id="c-ia" rows="3" placeholder="Ej: Supermercado Central TOTAL 1250.50 ..."></textarea></div>
        <button id="btn-parse-ia" class="boton boton-secundario" style="margin-top:4px">Deducir del texto</button>
      </div>

      <div>
        <div class="tarjeta">
          <h2>Foto del ticket</h2>
          <div class="foto-zone">
            <button id="btn-foto" class="boton boton-secundario">📷 Sacar / elegir foto</button>
            <img id="foto-previa" class="vista-previa oculta" alt="">
            <input id="c-foto" type="file" accept="image/*" capture="environment" class="oculta">
            <div id="foto-estado" style="font-size:12px;color:var(--gris);margin-top:6px;"></div>
          </div>
        </div>

        <div class="tarjeta">
          <h2>Detalle (opcional)</h2>
          <div id="lineas"></div>
          <button id="btn-linea" class="boton boton-secundario">＋ Añadir línea</button>
        </div>
      </div>
    </div>

    <button id="btn-guardar" class="boton boton-primario boton-ancho">Guardar compra</button>
  `;

  // Cargar productos históricos para desplegable de líneas
  try {
    const ts = await api("/api/tickets");
    const descs = [...new Set((ts || []).flatMap(t => (t.lineas || []).map(l => l.descripcion)).filter(Boolean))];
    let dlProd = document.getElementById("productos-hist");
    if (!dlProd) {
      dlProd = document.createElement("datalist");
      dlProd.id = "productos-hist";
      document.body.appendChild(dlProd);
    }
    dlProd.innerHTML = descs.map(d => `<option value="${d}">`).join("");
  } catch (e) { /* ignorar */ }

  // Cargar comercios históricos para desplegable
  try {
    const ts = await api("/api/tickets");
    const comercios = [...new Set((ts || []).map(t => t.comercio).filter(Boolean))];
    const selCom = $("#c-comercio");
    // Si es input text, no hay desplegable; si queremos sugerencias, usar datalist
    // Agregar datalist al input de comercio para sugerir
    let dl = document.getElementById("comercios-list");
    if (!dl) {
      dl = document.createElement("datalist");
      dl.id = "comercios-list";
      selCom.appendChild(dl);
    }
    dl.innerHTML = comercios.map(c => `<option value="${c}">`).join("");
    selCom.setAttribute("list", "comercios-list");
  } catch (e) { /* ignorar */ }

  // Cargar categorías
  const selCat = $("#c-categoria");
  for (const c of listasData.categorias) selCat.appendChild(new Option(c, c));
  const selPago = $("#c-pago");
  for (const p of listasData.formas_pago) selPago.appendChild(new Option(p, p));
  if (listasData.categorias.includes("Supermercado")) selCat.value = "Supermercado";
  if (listasData.formas_pago.includes("Tarjeta debito")) selPago.value = "Tarjeta debito";

  $("#btn-foto").addEventListener("click", () => $("#c-foto").click());
  $("#c-foto").addEventListener("change", manejarFoto);
  $("#btn-linea").addEventListener("click", () => agregarLinea());
  $("#btn-guardar").addEventListener("click", guardar);
  $("#btn-parse-ia").addEventListener("click", () => {
    const texto = $("#c-ia").value;
    if (!texto) return;
    // Detectar comercio: primera línea con letras, o "BUENAS MIGAS" etc.
    const lineas = texto.split(/\r?\n/);
    let comercio = "";
    for (const l of lineas) {
      const trimmed = l.trim();
      if (trimmed.length > 3 && trimmed.length < 60 && !/^\d/.test(trimmed) && !trimmed.includes("---") && !trimmed.includes("Total") && !trimmed.includes("Subtotal") && !trimmed.includes("Fecha")) {
        comercio = trimmed;
        break;
      }
    }
    $("#c-comercio").value = comercio || "BUENAS MIGAS";
    // Buscar total: línea que empieza con TOTAL o número grande con puntos/miles
    let total = 0;
    for (const l of lineas) {
      const t = l.match(/Total[:\s]*([0-9\.,]+)/i);
      if (t) {
        total = parseFloat(t[1].replace(/\./g, "").replace(",", "."));
        break;
      }
    }
    if (total > 0) $("#c-total").value = total.toFixed(2);
    // Buscar items con formato "Descripcion — Price" o "Descripcion cantidad precio"
    // Buscar items con formato real (ejemplo del usuario)
    lineas.forEach((l) => {
      const line = l.trim();
      // Formato "C CAFE CORTADO DOBLE — 3500,00"
      const m = line.match(/(.+?)\s+[-–—]\s*([0-9]{1,3}(?:\.[0-9]{3})*,[0-9]{2})/);
      if (m) {
        const desc = m[1].trim();
        const precioStr = m[2];
        const precio = parseFloat(precioStr.replace(/\./g, "").replace(",", "."));
        agregarLinea({ desc: desc, cant: 1, precio: precio });
        return;
      }
      // Formato cantidad x precio: "0,06 x 32000" (precio total calculado = 0,06*32000=1920) — usar precio como total de la línea
      const m2 = line.match(/([0-9]{1,3}(?:\.[0-9]{3})?,[0-9]+)\s*[xX]\s*([0-9]+(?:\.[0-9]+)?)/);
      if (m2) {
        const cant = parseFloat(m2[1].replace(".", "").replace(",", "."));
        const precioUnitario = parseFloat(m2[2]);
        const desc = line.split("x")[0].trim();
        const precioTotal = cant * precioUnitario;
        agregarLinea({ desc: desc || "MASAS FINAS POR KG", cant: Math.round(cant*100)/100 || 1, precio: precioTotal });
        return;
      }
      // Intento con precio al final y descripción antes: "MASAS FINAS POR KG — 1920,00"
      const m3 = line.match(/(.+?)\s+[-–—]\s*([0-9]+(?:\.[0-9]{2})?)/);
      if (m3) {
        const desc = m3[1].trim();
        const precioStr = m3[2];
        const precio = parseFloat(precioStr.replace(",", "."));
        agregarLinea({ desc: desc, cant: 1, precio: precio });
      }
    });
  });
}

function agregarLinea(prefill) {
  const cont = $("#lineas");
  const div = document.createElement("div");
  div.className = "linea-item";
  div.innerHTML = `
    <input class="l-desc" type="text" list="productos-hist" placeholder="Producto" value="${prefill && prefill.desc || ""}">
    <input class="l-cant" type="number" inputmode="decimal" step="0.01" min="0" placeholder="Cant." value="${prefill && prefill.cant || 1}">
    <input class="l-precio" type="number" inputmode="decimal" step="0.01" placeholder="Precio" value="${prefill && prefill.precio || ""}">
    <button class="quitar" title="Quitar">✕</button>`;
  div.querySelector(".quitar").addEventListener("click", () => div.remove());
  cont.appendChild(div);
}

function manejarFoto(e) {
  const file = e.target.files && e.target.files[0];
  if (!file) return;
  const img = new Image();
  const url = URL.createObjectURL(file);
  img.onload = () => {
    const maxLado = 1600;
    let { width, height } = img;
    const factor = Math.min(1, maxLado / Math.max(width, height));
    const canvas = document.createElement("canvas");
    canvas.width = Math.round(width * factor);
    canvas.height = Math.round(height * factor);
    canvas.getContext("2d").drawImage(img, 0, 0, canvas.width, canvas.height);
    fotoDataUrl = canvas.toDataURL("image/jpeg", 0.72);
    const previa = $("#foto-previa");
    previa.src = fotoDataUrl;
    previa.classList.remove("oculta");
    $("#foto-estado").textContent = "Foto lista";
    URL.revokeObjectURL(url);
    // Enviar a OCR para sugerir comercio y total
    api("/api/foto", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ data_url: fotoDataUrl }),
    }).then(r => {
      if (r.comercio_sugerido) {
        $("#c-comercio").value = r.comercio_sugerido;
        $("#foto-estado").textContent = "OCR: comercio detectado — " + (r.total_sugerido ? "$" + r.total_sugerido : "");
      }
      if (r.total_sugerido) {
        $("#c-total").value = r.total_sugerido;
      }
      // Si OCR detectó líneas de producto, cargarlas
      if (r.ocr && r.ocr.items && r.ocr.items.length) {
        const cont = $("#lineas");
        cont.innerHTML = "";
        r.ocr.items.forEach((it) => {
          agregarLinea({ desc: it.descripcion, cant: it.cantidad, precio: it.precio });
        });
        $("#foto-estado").textContent += " · " + r.ocr.items.length + " productos detectados";
      }
    }).catch(() => { /* si OCR falla, sigue manual */ });
  };
  img.src = url;
}

async function guardar() {
  const btn = $("#btn-guardar");
  const comercio = $("#c-comercio").value.trim();
  // Calcular líneas primero para obtener total automático
  const lineas = [];
  let totalCalculado = 0;
  document.querySelectorAll("#lineas .linea-item").forEach((item) => {
    const desc = item.querySelector(".l-desc").value.trim();
    const cant = parseFloat(item.querySelector(".l-cant").value) || 1;
    const precio = parseFloat(item.querySelector(".l-precio").value) || 0;
    if (desc && precio) {
      lineas.push({ descripcion: desc, cantidad: cant, precio: precio, categoria: $("#c-categoria").value });
      totalCalculado += cant * precio;
    }
  });
  if (lineas.length && totalCalculado > 0) {
    $("#c-total").value = totalCalculado.toFixed(2);
  }
  const total = parseFloat($("#c-total").value) || totalCalculado;
  if (!comercio) return aviso("Poné el comercio", "error");
  if (!(total > 0) && !(lineas.length > 0)) return aviso("Poné un total mayor que cero", "error");

  btn.disabled = true;
  btn.textContent = "Guardando…";
  try {
    let archivo = null;
    if (fotoDataUrl) {
      const r = await api("/api/foto", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ data_url: fotoDataUrl }),
      });
      archivo = r.path;
    }

    const payload = {
      fecha: $("#c-fecha").value,
      comercio,
      total,
      forma_pago: $("#c-pago").value,
      moneda: "ARS",
      numero: $("#c-numero").value.trim() || "-",
      notas: $("#c-notas").value.trim(),
      lineas: lineas.length
        ? lineas
        : [{ descripcion: comercio, cantidad: 1, precio: total, categoria: $("#c-categoria").value }],
    };
    if (archivo) payload.archivo = archivo;

    const r = await api("/api/tickets", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const tk = r.tickets[0];
    const dif = tk.diferencia;
    aviso(`Guardado ${tk.id}` + (Math.abs(dif) > 0.02 ? " (revisar diferencia)" : ""),
          Math.abs(dif) > 0.02 ? "error" : "ok");
    ir("resumen");
  } catch (e) {
    aviso(e.message, "error");
  } finally {
    btn.disabled = false;
    btn.textContent = "Guardar compra";
  }
}

// ---------------------------------------------------------------- arranque
document.querySelectorAll(".tab").forEach((t) =>
  t.addEventListener("click", () => ir(t.dataset.vista)));
$("#btnNueva").addEventListener("click", () => ir("nueva"));
ir("resumen");
