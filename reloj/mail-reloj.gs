/**
 * PROlimpio – Resumen mensual del reloj por mail.
 * Va en la planilla del reloj: Extensiones → Apps Script → borrar lo que hay, pegar esto y Guardar.
 * Ejecutar UNA vez la función "instalarReloj". Para probar, ejecutar "probarReloj" (manda el mes en curso).
 */

const RELOJ_LOCAL = "Colonia";                                                   // nombre del local
const RELOJ_DESTINOS = "coloniaadm@prolimpio.com.uy,colonia@prolimpio.com.uy";  // mails que reciben el resumen, separados por coma

const RELOJ_MESES = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Setiembre","Octubre","Noviembre","Diciembre"];
const RELOJ_DIAS = ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"];
const RELOJ_TZ = "America/Montevideo";

/** Ejecutar UNA vez: manda el resumen el día 1 de cada mes, a eso de las 10 de la mañana. */
function instalarReloj() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === "enviarResumenReloj") ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger("enviarResumenReloj").timeBased().onMonthDay(1).atHour(10).inTimezone(RELOJ_TZ).create();
  MailApp.sendEmail(RELOJ_DESTINOS, "✅ Resumen mensual del reloj activado",
    "Desde ahora, el día 1 de cada mes vas a recibir el resumen del reloj de PROlimpio " + RELOJ_LOCAL + ": días trabajados, faltas, llegadas tarde y horas extra de cada funcionario.");
}

/** Prueba: manda el resumen del mes en curso hasta hoy. */
function probarReloj() {
  const hoy = new Date();
  enviarMesReloj(hoy.getFullYear(), hoy.getMonth(), true);
}

/** Lo ejecuta el reloj automático el día 1: manda el mes anterior. */
function enviarResumenReloj() {
  const hoy = new Date();
  const ant = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 15);
  enviarMesReloj(ant.getFullYear(), ant.getMonth(), false);
}

function relojP2(n) { return ("0" + n).slice(-2); }
function relojMin(hhmm) { const p = String(hhmm).split(":"); return Number(p[0]) * 60 + Number(p[1]); }
function relojHm(m) { m = Math.round(m); return Math.floor(m / 60) + ":" + relojP2(m % 60); }
function relojFmt(ts, f) { return Utilities.formatDate(new Date(ts), RELOJ_TZ, f); }

function relojCalcular(y, m) {
  const ss = SpreadsheetApp.getActive();
  const cfgSh = ss.getSheetByName("RelojConfig"), marSh = ss.getSheetByName("RelojMarcas");
  if (!cfgSh || !marSh) throw new Error("No encuentro las hojas del reloj en esta planilla.");
  const R = JSON.parse(cfgSh.getRange("A1").getValue() || "{}");
  const tol = Number(R.tolerancia) || 0, fer = R.feriados || [];
  const pref = y + "-" + relojP2(m + 1);
  const hoyIso = relojFmt(Date.now(), "yyyy-MM-dd");
  const marcas = marSh.getLastRow() < 2 ? [] : marSh.getRange(2, 1, marSh.getLastRow() - 1, 9).getValues()
    .filter(r => r[6] !== "sí" && Number(r[7]))
    .map(r => ({ fecha: relojFmt(Number(r[7]), "yyyy-MM-dd"), t: Number(r[7]), manual: r[5] === "sí", emp: String(r[8]) }))
    .filter(x => x.fecha.indexOf(pref) === 0);
  const ult = new Date(y, m + 1, 0).getDate();
  const emps = (R.emps || []).filter(e => e.activo !== false || marcas.some(x => x.emp === e.id));
  return emps.map(e => {
    const r = { nombre: e.nombre, trab: 0, faltas: 0, tardes: 0, min: 0, extra: 0, inc: 0, dias: [] };
    for (let d = 1; d <= ult; d++) {
      const f = pref + "-" + relojP2(d); if (f > hoyIso) break;
      const w = new Date(y, m, d, 12).getDay();
      const tr = w >= 1 && w <= 5 ? (R.horario.lv || []) : w === 6 ? (R.horario.sab || []) : [];
      const lab = tr.length > 0 && fer.indexOf(f) < 0;
      const its = marcas.filter(x => x.emp === e.id && x.fecha === f).sort((a, b) => a.t - b.t);
      let trab = 0; for (let i = 0; i + 1 < its.length; i += 2) trab += (its[i + 1].t - its[i].t) / 60000;
      const inc = its.length % 2 === 1 && f < hoyIso;
      const esp = lab ? tr.reduce((s, t) => s + relojMin(t[1]) - relojMin(t[0]), 0) : 0;
      let tardes = 0, minT = 0;
      if (lab) tr.forEach((t, k) => { const en = its[2 * k]; if (!en) return; const x = relojMin(relojFmt(en.t, "HH:mm")) - relojMin(t[0]); if (x > tol) { tardes++; minT += x; } });
      let extra = 0; if (its.length && its.length % 2 === 0 && f !== hoyIso) { if (!lab) extra = trab; else if (trab - esp > tol) extra = trab - esp; }
      const falta = lab && !its.length && f < hoyIso && f >= (e.desde || "0000-00-00");
      if (its.length) r.trab++; if (falta) r.faltas++; r.tardes += tardes; r.min += trab; r.extra += extra; if (inc) r.inc++;
      if (lab || its.length) r.dias.push([e.nombre, f.slice(8) + "/" + f.slice(5, 7) + "/" + f.slice(0, 4), RELOJ_DIAS[w],
        its.map((it, i) => (i % 2 ? "S " : "E ") + relojFmt(it.t, "HH:mm") + (it.manual ? "*" : "")).join("  "),
        its.length ? relojHm(trab) : "", extra ? relojHm(extra) : "", tardes ? relojHm(minT) : "",
        falta ? "Falta" : inc ? "Falta una marca" : (!its.length && fer.indexOf(f) >= 0 ? "Feriado" : "")]);
    }
    return r;
  });
}

function enviarMesReloj(y, m, prueba) {
  const ss = SpreadsheetApp.getActive();
  const filas = relojCalcular(y, m);
  const nombre = "Reloj " + RELOJ_MESES[m] + " " + y;
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre); else sh.clear();
  const v = [["PROlimpio " + RELOJ_LOCAL + " – Planilla de " + RELOJ_MESES[m] + " " + y, "", "", "", "", "", "", ""], ["", "", "", "", "", "", "", ""],
    ["Funcionario", "Días trabajados", "Faltas", "Llegadas tarde", "Horas trabajadas", "Horas extra", "", ""]];
  filas.forEach(r => v.push([r.nombre, r.trab, r.faltas, r.tardes, relojHm(r.min), relojHm(r.extra), "", ""]));
  v.push(["", "", "", "", "", "", "", ""], ["Detalle por día", "", "", "", "", "", "", ""], ["Funcionario", "Fecha", "Día", "Marcas", "Trabajó", "Extra", "Tarde", "Estado"]);
  filas.forEach(r => r.dias.forEach(d => v.push(d)));
  v.push(["", "", "", "", "", "", "", ""], ["* marca agregada a mano. Generado " + relojFmt(Date.now(), "dd/MM/yyyy HH:mm"), "", "", "", "", "", "", ""]);
  sh.getRange(1, 1, v.length, 8).setNumberFormat("@").setValues(v);
  sh.getRange(1, 1).setFontWeight("bold").setFontSize(13);
  sh.getRange(3, 1, 1, 6).setFontWeight("bold");
  sh.autoResizeColumns(1, 8);

  const url = ss.getUrl() + "#gid=" + sh.getSheetId();
  const td = (x, r) => `<td style="padding:6px 8px;border-bottom:1px solid #F7E2CC;${r ? "text-align:right" : ""}">${x}</td>`;
  const tabla = filas.map(r => `<tr>${td(r.nombre + (r.inc ? ` <span style="color:#7F1D1D;font-size:12px">(${r.inc} día${r.inc > 1 ? "s" : ""} sin cerrar)</span>` : ""))}${td(r.trab, 1)}${td(r.faltas ? `<b style="color:#7F1D1D">${r.faltas}</b>` : 0, 1)}${td(r.tardes, 1)}${td(relojHm(r.min), 1)}${td("<b>" + relojHm(r.extra) + "</b>", 1)}</tr>`).join("");
  const html = `<div style="font-family:Arial,sans-serif;max-width:600px;color:#1F0D05">
    <div style="background:#E84400;color:#fff;padding:14px 16px;border-radius:10px 10px 0 0;font-size:18px;font-weight:bold">PROlimpio ${RELOJ_LOCAL} · Reloj de ${RELOJ_MESES[m]} ${y}</div>
    <div style="border:1px solid #F2C9A0;border-top:0;padding:12px 10px;border-radius:0 0 10px 10px">
      ${prueba ? '<div style="background:#FEF3C7;color:#92400E;padding:8px 10px;border-radius:8px;margin-bottom:8px">Prueba: mes en curso, hasta hoy.</div>' : ""}
      <table style="width:100%;border-collapse:collapse;font-size:14px">
        <tr style="color:#7A2E0A;font-size:12px;text-transform:uppercase"><th style="text-align:left;padding:6px 8px">Funcionario</th><th style="text-align:right;padding:6px 8px">Días</th><th style="text-align:right;padding:6px 8px">Faltas</th><th style="text-align:right;padding:6px 8px">Tarde</th><th style="text-align:right;padding:6px 8px">Horas</th><th style="text-align:right;padding:6px 8px">Extra</th></tr>
        ${tabla || '<tr><td colspan="6" style="padding:8px">No hay funcionarios cargados.</td></tr>'}
      </table>
      <p style="margin:14px 8px 4px"><a href="${url}" style="background:#E84400;color:#fff;padding:10px 14px;border-radius:8px;text-decoration:none;font-weight:bold">Ver la planilla con el detalle día por día</a></p>
    </div></div>`;
  MailApp.sendEmail({ to: RELOJ_DESTINOS, subject: (prueba ? "🧪 Prueba · " : "🕘 ") + "Reloj " + RELOJ_MESES[m] + " " + y + " · PROlimpio " + RELOJ_LOCAL, htmlBody: html });
}
