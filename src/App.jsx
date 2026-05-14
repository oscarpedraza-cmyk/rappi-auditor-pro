/**
 * ╔══════════════════════════════════════════════════════════════════════════════╗
 * ║   RAPPI PAIDLOT AUDITOR PRO  v4.7  —  Reingeniería Integral                 ║
 * ║   + logQueryToSheets · AdsAlertBanner · Tax % labels                      ║
 * ║   + Riesgo Fiscal proactivo · PDF mejorado · Fallbacks positivos           ║
 * ╚══════════════════════════════════════════════════════════════════════════════╝
 */

import { useState, useMemo, useCallback, useEffect, useRef, memo } from "react";
import * as XLSX from "xlsx";

// ─────────────────────────────────────────────────────────────────────────────
// § 0. SERVICE LAYER  (pure functions — no React, fully testable in isolation)
// ─────────────────────────────────────────────────────────────────────────────

// ── KPI_CONFIG — canonical KPI registry ─────────────────────────────────────
// Single source of truth for KPI keys and labels used in KPIGrid + exports.
const KPI_CONFIG = [
  { key: "ventaBruta",          label: "Ventas y descuentos para llegar a la venta neta", icon: "💰", color: "#22c55e" },
  { key: "darInversionTotal",   label: "Descuentos asumidos por Rappi",                   icon: "🎯", color: "#f97316" },
  { key: "descuentosVenta",     label: "Descuentos sobre la venta",                       icon: "🔄", color: "#8b5cf6" },
  { key: "comision",            label: "Uso y alquiler de plataforma Rappi y tasas Rappi",icon: "🏢", color: "#ef4444" },
  { key: "impuestosTotalExacto",label: "Impuestos",                                       icon: "🧾", color: "#0ea5e9", scrollTo: "section-impuestos" },
  { key: "totalAPagar",         label: "Total a Pagar",                                   icon: "✅", color: "#f59e0b" },
  { key: "otrosDescuentos",     label: "Otros Descuentos",                                icon: "📋", color: "#475569" },
  { key: "prestamos",           label: "Prestamos",                                       icon: "🏦", color: "#0f172a" },
];

// ── TAX_RULES — country → exact column name matches for tax line detection ──
// Extends COUNTRY_TAX_DETAIL with runtime column-matching logic.
// Each rule: { name, match } where match is the EXACT column header string
// as it appears in Rappi paidlots for that country.
// Used by: processPaidlot() to sum taxes by rule + by TaxRulesDebugPanel.
const TAX_RULES = {
  Argentina: {
    taxes: [
      { name: "IVA Plataforma",              pct: "21%",   match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "Percepción IVA",              pct: "10.5%", match: "Percepción de IVA" },
      { name: "Percepción Gral (BA)",        pct: "var.",  match: "Percepcion" },
      { name: "Percepción IIBB CABA",        pct: "3%–5%", match: "CABA" },
      { name: "Percepción IIBB CBDA",        pct: "3%",    match: "CBDA" },
      { name: "Percepción IIBB Cba.",        pct: "3%",    match: "Percepcion Cordoba" },
      { name: "Percepción IIBB Santa Fe",    pct: "3.5%",  match: "SANTA FE" },
      { name: "Percepción IIBB Tucumán",     pct: "3%",    match: "Percepción Tucuman" },
      { name: "Percepción IIBB Corrientes",  pct: "3%",    match: "Perceptión Corrientes" },
      { name: "Ret. Ganancias",              pct: "6%",    match: "Retencion Ganancias" },
      { name: "Ret. Buenos Aires",           pct: "3%",    match: "Retencion Buenos Aires" },
      { name: "Ret. Córdoba",                pct: "3%",    match: "Retencion Cordoba" },
      { name: "Ret. Tucumán",                pct: "3%",    match: "Retención Tucuman" },
      { name: "IVA RappiAds",               pct: "21%",   match: "IVA Rappi Ads" },
      { name: "ReteIVA RappiAds",           pct: "10.5%", match: "ReteIVA Rappi Ads" },
      { name: "IVA Campañas",               pct: "21%",   match: "IVA Campañas" },
      { name: "Percepción Campañas",        pct: "var.",  match: "Percepción Campañas" },
      { name: "Percepción Campañas Cba.",   pct: "3%",    match: "Percepción Campañas Cordoba" },
      { name: "Percepción Campañas Tucumán",pct: "3%",    match: "Percepción Campañas Tucuman" },
      { name: "Percepción Campañas Ctes.",  pct: "3%",    match: "Percepción Campañas Corrientes" },
      { name: "IVA descuento DAR",          pct: "21%",   match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  Colombia: {
    taxes: [
      { name: "IVA Plataforma",        pct: "19%",    match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "ReteIVA",               pct: "15%",    match: "Reteiva Uso y alquiler de plataforma Rappi" },
      { name: "Retefuente",            pct: "var.",   match: "Retefuente Uso y alquiler de plataforma Rappi" },
      { name: "ReteICA",               pct: "0.414%", match: "ReteICA Uso y alquiler de plataforma Rappi" },
      { name: "IVA RappiAds",          pct: "19%",    match: "IVA Rappi Ads" },
      { name: "ReteIVA RappiAds",      pct: "15%",    match: "ReteIVA Rappi Ads" },
      { name: "Impoconsumo",           pct: "8%",     match: "Impoconsumo / IVA de la venta (informativo)" },
      { name: "IVA Gasto Bancario",    pct: "19%",    match: "IVA Gasto Bancario" },
      { name: "Ret. Gasto Bancario",   pct: "var.",   match: "Retefuente Gasto Bancario" },
      { name: "IVA descuento DAR",     pct: "19%",    match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  México: {
    taxes: [
      { name: "IVA Plataforma",    pct: "16%",   match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "ISR Retención",     pct: "1%",    match: "ISR" },
      { name: "IVA RappiAds",      pct: "16%",   match: "IVA Rappi Ads" },
      { name: "IVA descuento DAR", pct: "16%",   match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  Chile: {
    taxes: [
      { name: "IVA Plataforma",    pct: "19%",  match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "IVA descuento DAR", pct: "19%",  match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  Perú: {
    taxes: [
      { name: "IGV Plataforma",    pct: "18%",  match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "IGV descuento DAR", pct: "18%",  match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  Ecuador: {
    taxes: [
      { name: "IVA Plataforma",    pct: "15%",  match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "IVA descuento DAR", pct: "15%",  match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  Uruguay: {
    taxes: [
      { name: "IVA Plataforma",    pct: "22%",  match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "IVA descuento DAR", pct: "22%",  match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  "Costa Rica": {
    taxes: [
      { name: "IVA Plataforma",    pct: "13%",  match: "IVA Uso y alquiler de plataforma Rappi" },
      { name: "IVA descuento DAR", pct: "13%",  match: "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR" },
    ],
  },
  Nicaragua: {
    taxes: [
      { name: "IVA/IR Plataforma", pct: "15%",  match: "IVA Uso y alquiler de plataforma Rappi" },
    ],
  },
};

/**
 * processPaidlot(colTotals, country) — Pure service function.
 * Takes the already-computed colTotals map (column name → sum) and a country string.
 * Returns a structured KPI object broken down by semantic category.
 *
 * Design decision: colTotals comes from parseWorkbook() which already did the
 * row-level aggregation. This function does the semantic grouping — exactly what
 * processPaidlot() did in the Pro Version, but using exact column matches from
 * TAX_RULES instead of loose string.includes() checks.
 *
 * @param {Object} colTotals  — { [columnName]: number }
 * @param {string} country    — country string matching CONFIG.countries keys
 * @returns {Object}          — { ventas, comision, compensaciones, darTotal, impuestosPorRegla, impuestosTotal, neto }
 */
function processPaidlot(colTotals, country) {
  const rules = TAX_RULES[country] ?? { taxes: [] };

  // Sum taxes using EXACT column matches from TAX_RULES (no loose includes)
  const impuestosPorRegla = rules.taxes.map(rule => ({
    name: rule.name,
    pct: rule.pct,
    match: rule.match,
    value: Math.abs(colTotals[rule.match] ?? 0),
  }));
  const impuestosTotal = impuestosPorRegla.reduce((s, r) => s + r.value, 0);

  return {
    ventas:           Math.abs(colTotals["Venta Bruta"] ?? 0),
    comision:         Math.abs(colTotals["Uso y alquiler de plataforma Rappi"] ?? 0),
    compensaciones:   Math.abs(colTotals["Compensaciones"] ?? 0),
    darTotal:         Math.abs(colTotals["Descuentos por inversión de Rappi DAR"] ?? 0),
    impuestosPorRegla,
    impuestosTotal,
  };
}

/**
 * buildConciliation(topKpis) — Pure service function.
 * Takes topKpis (from parseWorkbook) and returns a conciliation summary.
 * Mirrors buildConciliation() from the Pro Version but uses shadow reconciliation
 * (declared vs row-level sum) instead of a naive formula.
 *
 * @param {Object} topKpis
 * @param {Object} reconciliation  — { declared, shadow, diff, ok }
 * @returns {Object}               — { total, status, diff, declared, shadow, breakdown }
 */
function buildConciliation(topKpis, reconciliation) {
  const netCalculado = round2(
    topKpis.ventaBruta
    - topKpis.comision
    - topKpis.totalImpuestos
    - topKpis.ajustesTotal
    + topKpis.darBeneficioTotal
    - topKpis.compensaciones
  );
  return {
    total:      reconciliation.shadow,
    declared:   reconciliation.declared,
    shadow:     reconciliation.shadow,
    diff:       reconciliation.diff,
    status:     reconciliation.ok ? "OK" : "REVISAR",
    netCalculado,
    breakdown: {
      "Venta Bruta":       topKpis.ventaBruta,
      "(-) Comisión":     -topKpis.comision,
      "(-) Impuestos":    -topKpis.totalImpuestos,
      "(-) Ajustes":      -topKpis.ajustesTotal,
      "(+) DAR Beneficio": topKpis.darBeneficioTotal,
      "(-) Compensaciones":-topKpis.compensaciones,
    },
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// § 1. CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────────

const DAR_CONFIG = {
  Argentina: {
    flag: "🇦🇷", organismo: "AFIP", norma: "RG AFIP 4540/2019",
    plazoNC: "15 días calendario desde la factura original",
    tiposNC: ["NC Individual referenciada a la factura", "NC Global por usuario (Consumidor Final) agrupando múltiples facturas"],
    restricciones: "La NC global solo puede agrupar operaciones de un mismo usuario. Requiere trazabilidad completa: N° pedido, factura original, monto del descuento.",
    pasos: ["Verificar el monto DAR en el Portal Aliados o en la relación de ventas del paidlot.", "Emitir NC electrónica ante AFIP referenciando cada factura afectada.", "Para 'Consumidor Final': agrupar todas las facturas del mismo usuario en una NC global.", "Plazo máximo: 15 días calendario desde la emisión de la factura original.", "Si el aliado tiene integración POS: el descuento ya debe estar en la factura directamente."],
    ivaTip: "El descuento afecta la base de IVA. Al reducir la venta en $100, el IVA facturado también baja proporcionalmente.",
    urlOficial: "https://www.afip.gob.ar",
  },
  Colombia: {
    flag: "🇨🇴", organismo: "DIAN", norma: "Resolución DIAN 165/2023 · E.T. Art. 454",
    plazoNC: "Hasta 30 días calendario (NC no referenciada / global)",
    tiposNC: ["NC referenciada a factura específica (CUFE)", "NC no referenciada / global (sin asociar a factura, máx. 30 días de operaciones)"],
    restricciones: "Para que el descuento afecte la base gravable de IVA, debe ser: efectivo, constar en documento soporte, no estar condicionado, y ser normal según la costumbre comercial (Art. 454 E.T.).",
    pasos: ["Confirmar el valor DAR en la relación de ventas del paidlot.", "Opción A — Integración: el descuento ya aparece en la factura electrónica del POS.", "Opción B — NC referenciada: emitir nota crédito electrónica (CUFE) asociada a cada factura.", "Opción C — NC global: emitir una nota crédito no referenciada cubriendo hasta 30 días de operaciones.", "Verificar que el sistema de facturación transmita la NC a la DIAN en el mismo período gravable."],
    ivaTip: "La DIAN confirmó en Oficio 000206/2020 que los descuentos vía NC pueden afectar la base del IVA. Aplica también al Impuesto al Consumo.",
    urlOficial: "https://www.dian.gov.co",
  },
  México: {
    flag: "🇲🇽", organismo: "SAT", norma: "LISR Art. 25 Fracc. I · CFF Art. 29 · CFDI tipo E (Egreso)",
    plazoNC: "⚠️ MISMO MES en que se emitió la factura de ingresos",
    tiposNC: ["CFDI tipo E relacionado a un comprobante (NC individual)", "CFDI tipo E relacionado a varios comprobantes (NC global)", "CFDI tipo E a 'público general' (cuando el usuario no solicitó factura)"],
    restricciones: "CRÍTICO: La NC debe emitirse en el mismo mes que la factura de ingresos. Si se emite en mes posterior, el IVA no se descuenta en el mismo período. Para pagos >$2,000 MXN requiere transferencia electrónica.",
    pasos: ["Verificar el monto DAR en el Portal Partners o en la relación de ventas.", "Para usuario con RFC: emitir CFDI tipo E (Egreso) relacionado a su factura de ingresos.", "Para 'público general': emitir CFDI tipo E global relacionando la factura a 'público general'.", "⚠️ PLAZO CRÍTICO: emitir en el mismo mes calendario que la factura original.", "Guardar módulo JSON: product_id, amount_by_rappi, type y total_to_pay para auditoría SAT."],
    ivaTip: "El CFDI tipo E (Egreso) reduce tanto el ingreso fiscal como la base del IVA. Si se emite en el mes siguiente, el IVA solo se ajusta en ese mes posterior, no en el original.",
    urlOficial: "https://www.sat.gob.mx",
  },
  Chile: {
    flag: "🇨🇱", organismo: "SII", norma: "Ley de Facturación Electrónica · RUT 76.837.223-3 (Rappi Chile SPA)",
    plazoNC: "Sin plazo máximo explícito definido por norma",
    tiposNC: ["NC electrónica referenciada al folio de la boleta/factura original"],
    restricciones: "El descuento debe reflejarse restando sobre el precio unitario del servicio de plataforma en la misma factura electrónica de Rappi.",
    pasos: ["Verificar el monto en la relación de ventas del paidlot.", "Rappi Chile emite la factura con una segunda línea de descuento.", "El aliado puede emitir NC electrónica al SII referenciando el folio original.", "Con integración POS: el ticket al cliente ya refleja el precio neto."],
    ivaTip: "En Chile el IVA es 19%. El descuento sobre la comisión reduce la base imponible del IVA de Rappi al aliado.",
    urlOficial: "https://www.sii.cl",
  },
  Perú: {
    flag: "🇵🇪", organismo: "SUNAT", norma: "Factura Electrónica SUNAT · RUC 20602985971 (Rappi Perú SAC)",
    plazoNC: "Sin plazo máximo explícito; se recomienda dentro del mismo período",
    tiposNC: ["NC electrónica referenciada a la Factura Electrónica original (F001-XXXXXX)"],
    restricciones: "Las operaciones están sujetas al sistema SPOT (detracciones) al 10%.",
    pasos: ["Verificar el monto DAR en el Portal Partners.", "Rappi Perú emite la factura con una línea adicional de descuento (valor negativo).", "El aliado puede emitir NC electrónica ante SUNAT referenciando la factura original.", "Verificar que la base del IGV (18%) se calcule sobre el neto después del descuento.", "Recordar la detracción SPOT: aplica sobre el valor neto de la operación."],
    ivaTip: "El IGV es 18% en Perú. El descuento DAR reduce la base del IGV.",
    urlOficial: "https://www.sunat.gob.pe",
  },
  Ecuador: {
    flag: "🇪🇨", organismo: "SRI", norma: "Ley de Régimen Tributario Interno · RUC 1793010105001 (Rappiec S.A.S.)",
    plazoNC: "Sin plazo máximo explícito",
    tiposNC: ["NC electrónica referenciada a la factura original emitida por Rappi Ecuador"],
    restricciones: "IVA en Ecuador es 15%. El descuento aparece como línea separada en la factura de Rappi.",
    pasos: ["Verificar el monto DAR en la relación de ventas.", "Rappi Ecuador emite la factura con la línea '// DESCUENTO USO DE PLATAFORMA'.", "El aliado puede emitir NC electrónica al SRI referenciando el número de autorización.", "Verificar el impacto en IVA (15%) e IRBPNR sobre la base neta."],
    ivaTip: "El IVA ecuatoriano es 15%. El descuento reduce la base gravable del IVA.",
    urlOficial: "https://www.sri.gob.ec",
  },
  Uruguay: {
    flag: "🇺🇾", organismo: "DGI", norma: "e-Factura DGI · RUT 21817500000 (Rappi UY)",
    plazoNC: "Sin plazo máximo explícito",
    tiposNC: ["e-Nota de Crédito referenciada a la e-Factura original (Serie A)"],
    restricciones: "IVA en Uruguay es 22%.",
    pasos: ["Verificar el monto DAR en la relación de ventas del paidlot.", "Rappi UY emite la e-factura con la segunda línea de descuento (valor negativo).", "El aliado puede emitir e-Nota de Crédito ante DGI referenciando el número de e-Factura original."],
    ivaTip: "IVA Uruguay 22%. La reducción en comisión reduce la base del IVA.",
    urlOficial: "https://www.dgi.gub.uy",
  },
  "Costa Rica": {
    flag: "🇨🇷", organismo: "Ministerio de Hacienda", norma: "Factura Electrónica DGII · Cédula Jurídica 3101768820",
    plazoNC: "Sin plazo máximo explícito",
    tiposNC: ["Factura Electrónica con línea de descuento negativa", "NC electrónica referenciada"],
    restricciones: "IVA Costa Rica 13%.",
    pasos: ["Verificar el monto DAR en el Portal Partners.", "Rappi Pura Vida emite la factura con la segunda línea de descuento.", "El aliado puede emitir NC electrónica referenciando el consecutivo de la factura original."],
    ivaTip: "IVA Costa Rica 13%. El descuento reduce la base imponible del IVA.",
    urlOficial: "https://www.hacienda.go.cr",
  },
  Nicaragua: {
    flag: "🇳🇮", organismo: "DGI Nicaragua", norma: "Facturación local DGI",
    plazoNC: "Sin plazo máximo explícito",
    tiposNC: ["NC referenciada a factura original"],
    restricciones: "Verificar con contador local el tratamiento del IVA e IR sobre el descuento.",
    pasos: ["Verificar el monto DAR en la relación de ventas.", "Consultar con el equipo contable local el mecanismo de nota crédito aplicable."],
    ivaTip: "Consultar con el equipo contable la tasa de IVA e IR aplicable en Nicaragua.",
    urlOficial: "https://www.dgi.gob.ni",
  },
  "No detectado": {
    flag: "🌎", organismo: "—", norma: "Selecciona el país para ver la guía",
    plazoNC: "—", tiposNC: [], restricciones: "—", pasos: [], ivaTip: "—", urlOficial: "#",
  },
};

export const CONFIG = {
  countries: {
    Argentina: { currency: "ARS", locale: "es-AR", flag: "🇦🇷", taxAgency: "AFIP", taxLinks: [{ label: "AFIP", url: "https://www.afip.gob.ar" }, { label: "ARBA", url: "https://www.arba.gov.ar" }, { label: "AGIP", url: "https://www.agip.gob.ar" }], taxConcepts: "Percepciones IVA · Ingresos Brutos · CABA · Santa Fe · Córdoba · Tucumán · Corrientes" },
    Colombia: { currency: "COP", locale: "es-CO", flag: "🇨🇴", taxAgency: "DIAN", taxLinks: [{ label: "DIAN", url: "https://www.dian.gov.co" }, { label: "Hacienda Bogotá", url: "https://www.haciendabogota.gov.co" }], taxConcepts: "IVA · ReteIVA · ReteICA · Retefuente · Impoconsumo" },
    México: { currency: "MXN", locale: "es-MX", flag: "🇲🇽", taxAgency: "SAT", taxLinks: [{ label: "SAT", url: "https://www.sat.gob.mx" }, { label: "PRODECON", url: "https://www.prodecon.gob.mx" }], taxConcepts: "IVA · ISR · Retenciones SAT" },
    Chile: { currency: "CLP", locale: "es-CL", flag: "🇨🇱", taxAgency: "SII", taxLinks: [{ label: "SII", url: "https://www.sii.cl" }], taxConcepts: "IVA 19% · Retenciones SII" },
    Perú: { currency: "PEN", locale: "es-PE", flag: "🇵🇪", taxAgency: "SUNAT", taxLinks: [{ label: "SUNAT", url: "https://www.sunat.gob.pe" }], taxConcepts: "IGV 18% · Detracciones SPOT" },
    Ecuador: { currency: "USD", locale: "es-EC", flag: "🇪🇨", taxAgency: "SRI", taxLinks: [{ label: "SRI", url: "https://www.sri.gob.ec" }], taxConcepts: "IVA 15% · IRBPNR" },
    Uruguay: { currency: "UYU", locale: "es-UY", flag: "🇺🇾", taxAgency: "DGI", taxLinks: [{ label: "DGI", url: "https://www.dgi.gub.uy" }], taxConcepts: "IVA 22% · IRAE" },
    "Costa Rica": { currency: "CRC", locale: "es-CR", flag: "🇨🇷", taxAgency: "Hacienda", taxLinks: [{ label: "Hacienda", url: "https://www.hacienda.go.cr" }], taxConcepts: "IVA 13%" },
    Nicaragua: { currency: "NIO", locale: "es-NI", flag: "🇳🇮", taxAgency: "DGI NI", taxLinks: [{ label: "DGI NI", url: "https://www.dgi.gob.ni" }], taxConcepts: "IVA · IR" },
    "No detectado": { currency: "USD", locale: "es-US", flag: "🌎", taxAgency: "—", taxLinks: [], taxConcepts: "—" },
  },
  fingerprints: [
    { patterns: ["caba", "iibb", "percepcioncordoba", "retencioncordoba"], country: "Argentina" },
    { patterns: ["reteica", "impoconsumoivadelaventa", "retefuenteusoyalquiler", "gastobancario"], country: "Colombia" },
    { patterns: ["isr", "ivadescuentopordomiciliogratuito"], country: "México" },
    { patterns: ["igv"], country: "Perú" },
    { patterns: ["sii", "rut"], country: "Chile" },
    { patterns: ["sri"], country: "Ecuador" },
  ],
  groups: [
    {
      key: "ventas", label: "Ventas y descuentos para llegar a la venta neta", icon: "💰", color: "#22c55e", bg: "#dcfce7",
      tooltip: "Suma de Venta Bruta de todas las órdenes menos descuentos directos de producto. Es la base desde la que se calculan comisiones e impuestos.",
      cols: ["Venta Bruta", "Descuento de Producto asumido por el aliado", "Descuento de Producto", "Descuento en créditos"],
    },
    {
      key: "dar", label: "Descuentos asumidos por Rappi (DAR)", icon: "🎯", color: "#f97316", bg: "#fff7ed",
      tooltip: "DAR = Descuento Asumido por Rappi. Rappi financia el 100% de este descuento al cliente. El ingreso neto del aliado NO cambia.",
      cols: ["Descuentos por inversión de Rappi DAR"],
    },
    {
      key: "descuentosVenta", label: "Descuentos sobre la venta", icon: "🔄", color: "#8b5cf6", bg: "#ede9fe",
      tooltip: "Descuentos de domicilio, vouchers, pagos al repartidor, compensaciones y costos de canceladas que se descuentan de la venta.",
      cols: [
        "Costo de Domicilio - Propinas (marketplace)",
        "Meal Vouchers",
        "Total pagado por Repartidor independiente al Aliado en Efectivo",
        "Total pagado por el Usuario al Aliado  (marketplace)",
        "Total pagado por el Usuario al Aliado (marketplace)",
        "Descuento por Domicilio gratis",
        "Compensaciones",
        "Devolucion de Compensaciones",
        "Devolucion Compensaciones",
        "Costo Canceladas",
      ],
    },
    {
      key: "plataforma", label: "Uso y alquiler de plataforma Rappi y tasas Rappi", icon: "🏢", color: "#ef4444", bg: "#fee2e2",
      tooltip: "Comisión de plataforma + tarifas de servicio + inversión publicitaria RappiAds. Es el costo principal del aliado por operar en Rappi.",
      cols: [
        "Uso y alquiler de plataforma Rappi",
        "Descuento por inversión de Rappi a aplicar sobre Uso y alquiler de plataforma Rappi DAR",
        "Descuento por Service Fee",
        "Prime Uso y alquiler de plataforma Rappi",
        "Uso y alquiler de plataforma Rappi Prime",
        "Tarifa de Integration",
        "Tarifa por demora",
        "Tarifa Transaccional",
        "Tarifa transaccional",
        "Tarifa por activación (marketplace)",
        "Tarifa de servicio al usuario",
        "Servicios de Entrega y Recolección por cargo",
        "Contracargos",
        "Cuota de RappiAds",
        "Servicio de Cargo",
        "Descuento por pago anticipado",
      ],
    },
    {
      key: "impuestos", label: "Impuestos", icon: "🧾", color: "#0ea5e9", bg: "#e0f2fe",
      tooltip: "Impuestos aplicables según el país. El DAR también reduce la base de IVA/IGV sobre la comisión.",
      cols: [
        "IVA Uso y alquiler de plataforma Rappi",
        "IVA Campañas", "IVA Rappi Ads",
        "Reteiva Uso y alquiler de plataforma Rappi",
        "Retefuente Uso y alquiler de plataforma Rappi",
        "ReteIVA Rappi Ads", "ReteICA Uso y alquiler de plataforma Rappi",
        "ISR",
        "Percepcion", "Percepción de IVA",
        "Percepcion Cordoba", "Percepción Tucuman", "Perceptión Corrientes",
        "CABA", "CBDA", "SANTA FE",
        "Retencion Ganancias", "Retencion Buenos Aires",
        "Retencion Cordoba", "Retención Tucuman",
        "Percepción Campañas", "Percepción Campañas Cordoba", "Percepción Campañas Tucuman", "Percepción Campañas Corrientes",
        "IVA Descuento por Service Fee", "Retefuente Descuento por Service Fee",
        "IVA Servicio de Cargo", "Percepción Servicio de Cargo",
        "Percepción Córdoba Servicio de Cargo", "Percepción Corrientes Servicio de Cargo", "Percepción Tucuman Servicio de Cargo",
        "Percepción activación fee",
        "Percepción tarifa de servicio", "Percepción Cordoba Tarifa de servicio",
        "Percepción Current sobre Tarifa de servicio", "Percepción tucuman sobre Tarifa de servicio",
        "Impoconsumo / IVA de la venta (informativo)",
        "IVA Gasto Bancario", "Retefuente Gasto Bancario",
        "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR",
      ],
    },
    {
      key: "otrosDescuentos", label: "Otros Descuentos", icon: "📋", color: "#475569", bg: "#e2e8f0",
      tooltip: "Cashbacks asumidos por el aliado, desafíos de créditos Rappi y otros descuentos sobre la facturación.",
      cols: [
        "Cashback en Rappi creditos asumido por el aliado",
        "Challenge Rappi créditos asumidos por el aliado",
        "Cashback 15 mis o gratis",
        "Gasto bancario",
        "IVA Gasto Bancario",
        "Retefuente Gasto Bancario",
      ],
    },
    {
      key: "prestamos", label: "Préstamos", icon: "🏦", color: "#0f172a", bg: "#f1f5f9",
      tooltip: "Cuotas de préstamos y financiamientos Rappi descontados en este período.",
      cols: ["Cuota de préstamo", "Cuota de Prestamo", "Cuota prestamo"],
    },
    {
      key: "ajustes", label: "Ajustes y Deudas", icon: "⚖️", color: "#64748b", bg: "#f1f5f9",
      tooltip: "Ajustes manuales y deudas de períodos anteriores. Revisar con equipo financiero ante cualquier valor inesperado.",
      cols: ["Valor Ajustes Manuales", "Deuda Periodos Anteriores"],
    },
  ],
};

// ─────────────────────────────────────────────────────────────────────────────
// § 2. UTILITIES
// ─────────────────────────────────────────────────────────────────────────────

const slugify = (s) => String(s ?? "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
const cleanNum = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  if (v instanceof Date) return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  const s = String(v).trim();
  const lastComma = s.lastIndexOf(","), lastDot = s.lastIndexOf(".");
  let n = lastComma > lastDot ? s.replace(/\./g, "").replace(",", ".") : s.replace(/,/g, "");
  n = Number(n.replace(/[^0-9.-]/g, ""));
  return isFinite(n) ? n : 0;
};
const fmt = (value, country) => {
  const cfg = CONFIG.countries[country] ?? CONFIG.countries["No detectado"];
  try { return new Intl.NumberFormat(cfg.locale, { style: "currency", currency: cfg.currency, maximumFractionDigits: 2 }).format(value ?? 0); }
  catch { return `${cfg.currency} ${(value ?? 0).toFixed(2)}`; }
};
const fmtPct = (v) => `${(v * 100).toFixed(1)}%`;
const fmtDate = (v) => { if (!v) return "—"; if (v instanceof Date) return v.toISOString().split("T")[0]; const s = String(v).trim(); if (s.includes("T")) return s.split("T")[0]; if (s.includes(",")) return s.split(",")[0]; return s; };
const safeId = (v) => String(v ?? "").replace(/\.0$/, "").trim() || "—";
const round2 = (n) => Math.round(n * 100) / 100;

// ─────────────────────────────────────────────────────────────────────────────
// § 3. COUNTRY DETECTION
// ─────────────────────────────────────────────────────────────────────────────

function detectCountry(headers) {
  const slugs = new Set(headers.map(slugify));
  for (const { patterns, country } of CONFIG.fingerprints) {
    if (patterns.some(p => slugs.has(p) || [...slugs].some(s => s.includes(p)))) return { country, confidence: "high" };
  }
  return { country: "No detectado", confidence: "low" };
}

// ─────────────────────────────────────────────────────────────────────────────
// § 4. PARSER
// ─────────────────────────────────────────────────────────────────────────────

function parseWorkbook(wb, countryOverride = null) {
  if (!wb.Sheets["Resumen"]) throw new Error('Falta la pestaña "Resumen".');
  if (!wb.Sheets["Detalle"]) throw new Error('Falta la pestaña "Detalle".');

  const rsMatrix = XLSX.utils.sheet_to_json(wb.Sheets["Resumen"], { header: 1, defval: null, raw: true, cellDates: true });
  // findRS: locates label in any column, returns first non-null cell to its right
  // Handles Rappi layouts where value may be in col D, E, or further right
  const findCell = (r, col) => rsMatrix[r]?.[col] ?? null;
  const findRS = (label) => {
    for (const row of rsMatrix) {
      const idx = row.findIndex(cell => String(cell ?? "").trim() === label);
      if (idx >= 0) {
        for (let i = idx + 1; i < row.length; i++) {
          if (row[i] !== null && row[i] !== "" && row[i] !== undefined) return row[i];
        }
      }
    }
    return null;
  };
  // Loose label search: case-insensitive includes match, returns first numeric cell to the right
  const findRSPartial = (fragment) => {
    const fLow = fragment.toLowerCase();
    for (const row of rsMatrix) {
      const idx = row.findIndex(cell => String(cell ?? "").toLowerCase().trim().includes(fLow));
      if (idx >= 0) {
        for (let i = idx + 1; i < row.length; i++) {
          const v = row[i];
          if (v !== null && v !== "" && v !== undefined && cleanNum(v) !== 0) return v;
        }
        // Value might live in the SAME cell after a colon, e.g. "Total a pagar: 12345"
        const same = String(row[idx] ?? "");
        const colonIdx = same.indexOf(":");
        if (colonIdx >= 0) {
          const after = same.slice(colonIdx + 1).trim();
          if (after && cleanNum(after) !== 0) return after;
        }
      }
    }
    return null;
  };
  // findColCD: scan every row for a label in column C (idx 2), return value from column D (idx 3).
  // Rappi Resumen layouts consistently put the label in col C and the figure in col D.
  const findColCD = (label) => {
    const lLow = label.toLowerCase();
    for (const row of rsMatrix) {
      const cellC = String(row[2] ?? "").trim().toLowerCase();
      if (cellC === lLow || cellC.includes(lLow)) {
        const valD = row[3];
        if (valD !== null && valD !== "" && valD !== undefined) return valD;
      }
    }
    return null;
  };

  // Dates: D4 = row-index 3 col D (idx 3), D5 = row-index 4 col D — primary for real paidlots
  const inicioRaw =
    findCell(3, 3) ??
    findRS("Inicio Período de venta") ?? findRS("Inicio Periodo de venta") ??
    findRS("Inicio Período") ?? findRS("Inicio periodo") ?? "—";
  const finRaw =
    findCell(4, 3) ??
    findRS("Fin Período de venta") ?? findRS("Fin Periodo de venta") ??
    findRS("Fin Período") ?? findRS("Fin periodo") ?? "—";

  // Total declarado: column C/D scan is primary; other strategies are fallbacks
  const totalDeclaradoRaw =
    findColCD("Valor total a transferir") ??
    findRS("Valor total a transferir") ??
    findRS("Total a transferir") ??
    findRS("Valor neto a transferir") ??
    findRS("Total a pagar") ??
    findRS("Neto a pagar") ??
    findRS("Valor a transferir") ??
    findRSPartial("total a transferir") ??
    findRSPartial("total a pagar") ??
    findRSPartial("neto a pagar") ??
    findRSPartial("valor neto") ??
    findCell(6,3) ?? findCell(7,3) ?? findCell(8,3) ?? findCell(9,3) ??
    findCell(10,3) ?? findCell(11,3) ?? findCell(12,3) ?? null;
  const resumen = { inicio: inicioRaw, fin: finRaw, fechaPago: findRS("Fecha de pago establecida por contrato") ?? "—", aliado: findRS("Nombre del aliado") ?? "—", idPago: safeId(findRS("ID de pago")), totalDeclarado: cleanNum(totalDeclaradoRaw) };

  const compSheet = wb.Sheets["Compensaciones"];
  const compMatrix = compSheet ? XLSX.utils.sheet_to_json(compSheet, { header: 1, defval: null, raw: false }) : [];
  const compRows = [];
  for (let i = 2; i < compMatrix.length; i++) {
    const r = compMatrix[i];
    if (!r[0]) continue;
    compRows.push({ orderId: safeId(r[0]), fecha: r[1] ? String(r[1]).split("T")[0] : "—", razon: r[2] ?? "—", monto: cleanNum(r[4]), productos: String(r[5] ?? "").replace(/['"]/g, ""), comentario: String(r[7] ?? "").slice(0, 120) });
  }

  const detMatrix = XLSX.utils.sheet_to_json(wb.Sheets["Detalle"], { header: 1, defval: null, raw: true, cellDates: true });
  const rawHeaders = detMatrix[1] ?? [];
  const headers = rawHeaders.map(h => h ? String(h).trim() : "");
  const colByName = {}, colBySlug = {};
  headers.forEach((h, i) => { if (!h) return; colByName[h] = i; colBySlug[slugify(h)] = i; });
  const getCol = (row, ...candidates) => { for (const c of candidates) { if (colByName[c] !== undefined) return row[colByName[c]]; if (colBySlug[slugify(c)] !== undefined) return row[colBySlug[slugify(c)]]; } return null; };

  const detection = countryOverride ? { country: countryOverride, confidence: "manual" } : detectCountry(headers);
  const tipoIdx = colByName["Tipo de transacción"] ?? colBySlug[slugify("Tipo de transacción")];
  const dataRows = detMatrix.slice(2).filter(r => r[tipoIdx]);
  if (!dataRows.length) throw new Error("Sin transacciones en Detalle.");

  const firstRow = dataRows[0] ?? [];
  const meta = { tienda: String(getCol(firstRow, "Nombre de la tienda") ?? resumen.aliado ?? "—"), tiendaId: safeId(getCol(firstRow, "ID de la tienda")), paidlotId: safeId(getCol(firstRow, "ID del paidlot") ?? resumen.idPago) };

  const colTotals = {};
  let shadowTotal = 0;
  for (const row of dataRows) {
    headers.forEach((h, i) => { if (!h) return; colTotals[h] = (colTotals[h] ?? 0) + cleanNum(row[i]); });
    shadowTotal += cleanNum(getCol(row, "Valor a transferir", "Valor Neto", "Valor transferir", "Neto a transferir", "Neto", "Total a transferir", "Valor Total a Transferir"));
  }

  // Fallback: compute net from signed sum of all monetary group columns
  // Covers Excel files generated programmatically with formula cells (no cached values)
  if (Math.abs(shadowTotal) <= 0.01) {
    const allGroupCols = new Set(CONFIG.groups.flatMap(g => g.cols.map(slugify)));
    shadowTotal = Object.entries(colTotals)
      .filter(([col]) => allGroupCols.has(slugify(col)))
      .reduce((sum, [, val]) => sum + val, 0);
  }

  const shadowRounded = round2(shadowTotal), declaredRounded = round2(resumen.totalDeclarado);
  const reconciliationDiff = round2(shadowRounded - declaredRounded);
  const reconciliationOk = declaredRounded === 0 || Math.abs(reconciliationDiff) < 0.02;

  const byType = {};
  for (const row of dataRows) {
    const t = String(row[tipoIdx] ?? "").trim().toUpperCase();
    if (!byType[t]) byType[t] = [];
    byType[t].push(row);
  }
  const ordenRows = byType["ORDEN"] ?? [];
  const compDetRows = byType["COMPENSACIÓN"] ?? byType["COMPENSACION"] ?? [];
  const extraRows = byType["EXTRA SERVICE"] ?? [];

  const DAR_COL = "Descuentos por inversión de Rappi DAR";
  const DAR_COMP_COL = "Descuento por inversión de Rappi  a aplicar sobre Uso y alquiler de plataforma Rappi DAR";
  const DAR_IVA_COL = "Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR";
  const totalDARInversion = Math.abs(colTotals[DAR_COL] ?? 0);
  const totalDARComision = Math.abs(colTotals[DAR_COMP_COL] ?? 0);
  const totalDARIVA = Math.abs(colTotals[DAR_IVA_COL] ?? 0);
  const totalDARBeneficio = totalDARInversion + totalDARComision + totalDARIVA;
  const hasDar = totalDARInversion > 0;

  const ordersTable = ordenRows.map(r => {
    const ventaBruta = cleanNum(getCol(r, "Venta Bruta"));
    const darInversion = Math.abs(cleanNum(getCol(r, DAR_COL)));
    const darComision = Math.abs(cleanNum(getCol(r, DAR_COMP_COL)));
    const darIva = Math.abs(cleanNum(getCol(r, DAR_IVA_COL)));
    const darTotal = darInversion + darComision + darIva;
    return { fecha: fmtDate(getCol(r, "Fecha de creación orden")), ordenId: safeId(getCol(r, "ID de la órden", "ID de la orden")), tienda: String(getCol(r, "Nombre de la tienda") ?? "—"), metodoPago: String(getCol(r, "Método de pago") ?? "—"), prime: String(getCol(r, "Prime  ", "Prime") ?? "") === "true", ventaBruta, darInversion, darComision, darIva, darTotal, darPct: ventaBruta > 0 && darInversion > 0 ? darInversion / ventaBruta : 0, comision: cleanNum(getCol(r, "Uso y alquiler de plataforma Rappi")), subtotal: cleanNum(getCol(r, "Subtotal antes de impuestos")), neto: cleanNum(getCol(r, "Valor a transferir", "Valor Neto")), pctComision: cleanNum(getCol(r, "Porcentaje de Uso y alquiler de plataforma Rappi")) };
  });

  const extrasTable = extraRows.map(r => ({ fecha: fmtDate(getCol(r, "Fecha de creación orden")), ordenId: safeId(getCol(r, "ID de la órden", "ID de la orden")), tipo: String(getCol(r, "Tipo de transacción") ?? "EXTRA SERVICE"), ventaBruta: cleanNum(getCol(r, "Venta Bruta")), neto: cleanNum(getCol(r, "Valor a transferir", "Valor Neto")) }));

  const ajustesRows = dataRows.filter(r => Math.abs(cleanNum(getCol(r, "Valor Ajustes Manuales"))) > 0 || Math.abs(cleanNum(getCol(r, "Deuda Periodos Anteriores"))) > 0).map(r => ({ fecha: fmtDate(getCol(r, "Fecha de creación orden")), ordenId: safeId(getCol(r, "ID de la órden", "ID de la orden")), razon: String(getCol(r, "Razon (Ajuste / RADs)", "Razon") ?? "—"), descripcion: String(getCol(r, "Descripción o comentarios (Ajustes  / RADs)  ", "Descripcion") ?? "—").slice(0, 100), ajuste: cleanNum(getCol(r, "Valor Ajustes Manuales")), deuda: cleanNum(getCol(r, "Deuda Periodos Anteriores")) }));

  const groups = CONFIG.groups.map(g => {
    const seenResolved = new Set(); // dedup by resolved Excel column slug
    const seenLabelVal = new Set(); // dedup by (label prefix + value) — catches near-duplicate config entries
    const items = g.cols.map(col => {
      // Resolve the actual Excel column that this config col maps to
      const resolvedKey = (colTotals[col] !== undefined && colTotals[col] !== 0)
        ? col
        : (headers.find(h => slugify(h) === slugify(col)) ?? null);
      if (!resolvedKey) return null;
      const v = colTotals[resolvedKey] ?? 0;
      return { col, label: col, value: v, resolvedKey };
    }).filter(item => {
      if (!item) return false;
      if (Math.abs(item.value) <= 0.005) return false;
      // Primary dedup: by resolved Excel column slug
      const rs = slugify(item.resolvedKey);
      if (seenResolved.has(rs)) return false;
      // Secondary dedup: same label prefix + same value → treat as duplicate
      // Catches: "Tarifa Transaccional" vs "Tarifa transaccional", double-space DAR variants, etc.
      const labelPrefix = slugify(item.label).slice(0, 22);
      const lvKey = `${labelPrefix}_${Math.abs(item.value).toFixed(2)}`;
      if (seenLabelVal.has(lvKey)) return false;
      seenResolved.add(rs);
      seenLabelVal.add(lvKey);
      return true;
    }).map(({ col, label, value }) => ({ col, label, value }));
    const total = items.reduce((s, item) => s + Math.abs(item.value), 0);
    return { ...g, items, total };
  }).filter(g => g.total > 0.005);

  // comisionTotal = full platform group sum (base + ads + tarifas)
  const comisionBase = Math.abs(colTotals["Uso y alquiler de plataforma Rappi"] ?? 0);
  const plataformaGroup = groups.find(g => g.key === "plataforma");
  const comisionTotal = plataformaGroup ? plataformaGroup.total : comisionBase;
  // totalImpuestos = full impuestos group sum (all items, same pattern as comisionTotal)
  const impuestosGroupFull = groups.find(g => g.key === "impuestos");
  const totalImpuestos = impuestosGroupFull
    ? impuestosGroupFull.total
    : (Math.abs(colTotals["IVA Uso y alquiler de plataforma Rappi"] ?? 0) + Math.abs(colTotals["Reteiva Uso y alquiler de plataforma Rappi"] ?? 0) + Math.abs(colTotals["IVA Campañas"] ?? 0) + Math.abs(colTotals["ReteICA Uso y alquiler de plataforma Rappi"] ?? 0) + Math.abs(colTotals["ISR"] ?? 0));
  const ventaBrutaTotal = Math.abs(colTotals["Venta Bruta"] ?? 0);

  // Run service layer: exact-column tax matching via TAX_RULES
  const serviceKpis = processPaidlot(colTotals, detection.country);

  const topKpis = {
    ventaBruta: ventaBrutaTotal,
    comision: comisionTotal,
    compensaciones: Math.abs(colTotals["Compensaciones"] ?? 0),
    // KPI: Descuentos sobre la venta = Devolución Compensaciones + Costo Canceladas
    descuentosVenta: round2(
      Math.abs(colTotals["Devolucion de Compensaciones"] ?? 0) +
      Math.abs(colTotals["Devolucion Compensaciones"] ?? 0) +
      Math.abs(colTotals["Costo Canceladas"] ?? 0) +
      Math.abs(colTotals["Costo de Domicilio - Propinas (marketplace)"] ?? 0) +
      Math.abs(colTotals["Meal Vouchers"] ?? 0) +
      Math.abs(colTotals["Total pagado por Repartidor independiente al Aliado en Efectivo"] ?? 0) +
      Math.abs(colTotals["Total pagado por el Usuario al Aliado  (marketplace)"] ?? 0) +
      Math.abs(colTotals["Total pagado por el Usuario al Aliado (marketplace)"] ?? 0) +
      Math.abs(colTotals["Descuento por Domicilio gratis"] ?? 0)
    ),
    // KPI: Otros Descuentos = cashbacks + challenges + gasto bancario
    otrosDescuentos: round2(
      Math.abs(colTotals["Cashback en Rappi creditos asumido por el aliado"] ?? 0) +
      Math.abs(colTotals["Challenge Rappi créditos asumidos por el aliado"] ?? 0) +
      Math.abs(colTotals["Cashback 15 mis o gratis"] ?? 0) +
      Math.abs(colTotals["Gasto bancario"] ?? 0) +
      Math.abs(colTotals["IVA Gasto Bancario"] ?? 0) +
      Math.abs(colTotals["Retefuente Gasto Bancario"] ?? 0)
    ),
    // KPI: Préstamos = cuota de préstamo
    prestamos: round2(
      Math.abs(colTotals["Cuota de préstamo"] ?? 0) +
      Math.abs(colTotals["Cuota de Prestamo"] ?? 0) +
      Math.abs(colTotals["Cuota prestamo"] ?? 0)
    ),
    // neto / totalAPagar: declared Resumen value is primary (most accurate)
    // shadowTotal is fallback when Resumen doesn't have the value
    neto: Math.abs(resumen.totalDeclarado) > 0.01 ? resumen.totalDeclarado : (Math.abs(shadowTotal) > 0.01 ? shadowTotal : 0),
    totalAPagar: Math.abs(resumen.totalDeclarado) > 0.01 ? resumen.totalDeclarado : (Math.abs(shadowTotal) > 0.01 ? shadowTotal : 0),
    ordenes: ordenRows.length,
    ajustesTotal: Math.abs(colTotals["Valor Ajustes Manuales"] ?? 0) + Math.abs(colTotals["Deuda Periodos Anteriores"] ?? 0),
    // darInversionTotal = beneficio completo (producto + comisión + IVA) — mismo valor que el banner
    darInversionTotal: totalDARBeneficio,
    darInversionProducto: totalDARInversion,
    darComisionTotal: totalDARComision,
    darIvaTotal: totalDARIVA,
    darBeneficioTotal: totalDARBeneficio,
    hasDar,
    darPctSobreVentas: ventaBrutaTotal > 0 ? totalDARInversion / ventaBrutaTotal : 0,
    rappiAdsCollection: Math.abs(colTotals["Descuento rappi_ads_invoiced_collection"] ?? 0),
    cuotaRappiAds: Math.abs(colTotals["Cuota de RappiAds"] ?? 0),
    totalImpuestos,
    get effectiveFee() { return this.ventaBruta <= 0 ? 0 : round2((this.comision + this.totalImpuestos) / this.ventaBruta); },
    // Service layer output — exact tax breakdown by TAX_RULES column matches
    impuestosPorRegla: serviceKpis.impuestosPorRegla,
    impuestosTotalExacto: serviceKpis.impuestosTotal,
  };

  return {
    id: `${meta.paidlotId}-${Date.now()}`,
    resumen, meta, detection, groups, ordersTable,
    compRows: compRows.length > 0 ? compRows : compDetRows.map(r => ({ orderId: safeId(getCol(r, "ID de la órden", "ID de la orden")), fecha: fmtDate(getCol(r, "Fecha de creación orden")), razon: "—", monto: Math.abs(cleanNum(getCol(r, "Compensaciones"))), productos: "—", comentario: "—" })),
    extrasTable, ajustesRows, topKpis, colTotals,
    shadowTotal: shadowRounded,
    reconciliation: { declared: declaredRounded, shadow: shadowRounded, diff: reconciliationDiff, ok: reconciliationOk },
    conciliationService: buildConciliation(
      { ventaBruta: ventaBrutaTotal, comision: comisionBase, totalImpuestos: totalImpuestos, ajustesTotal: Math.abs(colTotals["Valor Ajustes Manuales"] ?? 0) + Math.abs(colTotals["Deuda Periodos Anteriores"] ?? 0), darBeneficioTotal: totalDARBeneficio, compensaciones: Math.abs(colTotals["Compensaciones"] ?? 0) },
      { declared: declaredRounded, shadow: shadowRounded, diff: reconciliationDiff, ok: reconciliationOk }
    ),
    headers, loadedAt: Date.now(),
  };
}

// ─────────────────────────────────────────────────────────────────────────────
// § 5. localStorage
// ─────────────────────────────────────────────────────────────────────────────

const LS_KEY = "rappi_paidlots_v47";
const savePaidlots = (list) => { try { localStorage.setItem(LS_KEY, JSON.stringify(list.map(p => ({ ...p, headers: [] })).slice(0, 20))); } catch {} };
const loadPaidlots = () => { try { const r = localStorage.getItem(LS_KEY); return r ? JSON.parse(r) : []; } catch { return []; } };

// ─────────────────────────────────────────────────────────────────────────────
// § 6. ALERTS + PDF EXPORT
// ─────────────────────────────────────────────────────────────────────────────

function detectAlerts(kpi, country) {
  const alerts = [];
  const taxTotal = kpi.impuestosTotalExacto ?? kpi.totalImpuestos ?? 0;
  const taxPct   = kpi.ventaBruta > 0 ? taxTotal / kpi.ventaBruta : 0;
  const adsPct   = kpi.ventaBruta > 0 && kpi.cuotaRappiAds > 0 ? kpi.cuotaRappiAds / kpi.ventaBruta : 0;
  const compPct  = kpi.ventaBruta > 0 && kpi.compensaciones > 0 ? kpi.compensaciones / kpi.ventaBruta : 0;

  if (kpi.effectiveFee > 0.35)
    alerts.push({ type: "danger", icon: "🚨", title: `Solo el ${(100 - kpi.effectiveFee*100).toFixed(1)}% del total queda disponible para el aliado`, msg: `Del total de ventas facturadas, solo el ${(100 - kpi.effectiveFee*100).toFixed(1)}% llega al aliado (tarifa efectiva: ${(kpi.effectiveFee*100).toFixed(1)}%). Revisar certificados de exención o deducciones aplicables.` });
  if (taxPct > 0.15)
    alerts.push({ type: "danger", icon: "🧾", title: "Carga impositiva alta", msg: `Impuestos representan el ${(taxPct*100).toFixed(1)}% de las ventas brutas (umbral recomendado: 15%). Verificar con el contador si existen regímenes especiales.` });
  if (compPct > 0.05)
    alerts.push({ type: "warning", icon: "⚠️", title: "Compensaciones elevadas", msg: `Compensaciones equivalen al ${(compPct*100).toFixed(1)}% de ventas (umbral: 5%). Revisar causas: stock, tiempos de preparación o pedidos incompletos.` });
  if (adsPct > 0.20)
    alerts.push({ type: "warning", icon: "📺", title: "RappiAds supera el 20% de ventas", msg: `Inversión en ADS = ${(adsPct*100).toFixed(1)}% de ventas brutas. Verificar que el ROI justifique esta inversión antes de renovar.` });
  if (kpi.hasDar && kpi.darInversionTotal > 0)
    alerts.push({ type: "info", icon: "🎯", title: "DAR activo — emitir Nota de Crédito", msg: `Hay inversión DAR activa. Recordar al aliado emitir la NC fiscal correspondiente para optimizar la carga impositiva.` });
  if (!kpi.hasDar && kpi.ventaBruta > 0)
    alerts.push({ type: "info", icon: "💡", title: "Sin DAR activo — oportunidad", msg: `El aliado no tiene DAR activo este período. Activar campañas DAR aumenta la demanda sin afectar el neto.` });
  if (!kpi.cuotaRappiAds && kpi.ventaBruta > 0)
    alerts.push({ type: "info", icon: "📺", title: "Sin pauta RappiAds", msg: `No hay inversión en publicidad este período. Una pauta bien segmentada puede generar 3x–5x de retorno.` });
  if (kpi.ajustesTotal > kpi.ventaBruta * 0.03)
    alerts.push({ type: "warning", icon: "⚖️", title: "Ajustes contables significativos", msg: `Ajustes del período = ${(kpi.ajustesTotal/kpi.ventaBruta*100).toFixed(1)}% de ventas. Validar con el equipo contable si corresponden a correcciones de liquidaciones anteriores.` });

  return alerts;
}

function exportPDF(paidlot, country, aiInsights = "", allSelected = []) {
  if (!paidlot) return;
  const p = paidlot;
  const kpi = p.topKpis;
  const cfg = CONFIG.countries[country] ?? CONFIG.countries["No detectado"];
  const fmtV = (v) => new Intl.NumberFormat(cfg.locale, { style: "currency", currency: cfg.currency, maximumFractionDigits: 2 }).format(v ?? 0);
  const isMultiPeriod = allSelected.length > 1;
  const sortedPeriods = isMultiPeriod ? [...allSelected].sort((a, b) => (a.resumen.inicio ?? "").localeCompare(b.resumen.inicio ?? "")) : [];

  // ── Aggregated multi-period totals (para narrativa del farmer) ────────────
  const multiTotals = isMultiPeriod ? {
    ventaBruta:   sortedPeriods.reduce((s, sp) => s + (sp.topKpis.ventaBruta ?? 0), 0),
    totalAPagar:  sortedPeriods.reduce((s, sp) => s + (sp.topKpis.totalAPagar ?? sp.topKpis.neto ?? 0), 0),
    comision:     sortedPeriods.reduce((s, sp) => s + (sp.topKpis.comision ?? 0), 0),
    impuestos:    sortedPeriods.reduce((s, sp) => s + (sp.topKpis.impuestosTotalExacto ?? sp.topKpis.totalImpuestos ?? 0), 0),
    ordenes:      sortedPeriods.reduce((s, sp) => s + sp.ordersTable.length, 0),
    darInversion: sortedPeriods.reduce((s, sp) => s + (sp.topKpis.darInversionTotal ?? 0), 0),
    ads:          sortedPeriods.reduce((s, sp) => s + (sp.topKpis.cuotaRappiAds ?? 0), 0),
    inicio:       sortedPeriods[0].resumen.inicio,
    fin:          sortedPeriods[sortedPeriods.length - 1].resumen.fin,
  } : null;

  // ── Derived values ────────────────────────────────────────────────────────
  const totalAPagar   = kpi.totalAPagar ?? kpi.neto;
  const totalImpuestos = kpi.impuestosTotalExacto ?? kpi.totalImpuestos ?? 0;
  const adsPct        = kpi.ventaBruta > 0 && kpi.cuotaRappiAds > 0 ? kpi.cuotaRappiAds / kpi.ventaBruta : 0;
  const taxPctVentas  = kpi.ventaBruta > 0 ? totalImpuestos / kpi.ventaBruta : 0;
  const feePct        = (kpi.effectiveFee * 100).toFixed(1);
  const darCfg        = DAR_CONFIG[country]        ?? DAR_CONFIG["No detectado"];
  const taxCfg        = COUNTRY_TAX_DETAIL[country] ?? COUNTRY_TAX_DETAIL["No detectado"];
  const taxRows       = (kpi.impuestosPorRegla ?? []).filter(r => r.value > 0);
  const pdfSafe = (s) => String(s ?? "").replace(/[^ -ÿ]/g, "").replace(/\s+/g, " ").trim();

  // ── Recomendaciones de Facturación (data-driven) ─────────────────────────
  const recosFact = [];
  if (!kpi.hasDar && kpi.ventaBruta > 0)
    recosFact.push({ prio: "media", titulo: "Aliado sin DAR — evaluar elegibilidad", texto: `Este aliado no tiene DAR activo. Con ${fmtV(kpi.ventaBruta)} en ventas brutas, puede ser candidato según los criterios de coinversión de Rappi. Conversarlo con el equipo comercial para evaluar si aplica al programa.` });
  if (kpi.hasDar)
    recosFact.push({ prio: "media", titulo: "DAR activo — revisar performance", texto: `La inversión DAR activa de ${fmtV(kpi.darInversionTotal)} está generando demanda. Revisar con el equipo comercial si el rango de productos elegibles o los descuentos están optimizados para este período.` });
  if (!kpi.cuotaRappiAds || kpi.cuotaRappiAds === 0)
    recosFact.push({ prio: "media", titulo: "Iniciar pauta RappiAds", texto: `El aliado no tiene inversión activa en RappiAds. Una pauta bien segmentada en hora pico puede triplicar la visibilidad y generar un retorno de 3x a 5x sobre la inversión.` });
  if (kpi.cuotaRappiAds > 0 && adsPct < 0.10)
    recosFact.push({ prio: "baja", titulo: "Incrementar RappiAds", texto: `La cuota de ${fmtV(kpi.cuotaRappiAds)} (${(adsPct*100).toFixed(1)}% de ventas) está en zona saludable con margen de crecimiento. Aumentar la pauta puede capturar más tráfico sin comprometer el flujo de caja.` });
  if (adsPct > 0.20)
    recosFact.push({ prio: "alta", titulo: "Revisar ROI de RappiAds", texto: `La cuota de RappiAds representa el ${(adsPct*100).toFixed(1)}% de ventas — supera el umbral crítico del 20%. Antes de renovar, verificar que el incremento en pedidos justifique esta inversión.` });
  else if (adsPct > 0.10)
    recosFact.push({ prio: "media", titulo: "Monitorear ROI de RappiAds", texto: `La cuota de RappiAds (${(adsPct*100).toFixed(1)}% de ventas) supera el 10%. Verificar que las campañas estén generando retorno visible en pedidos y GMV.` });
  if (kpi.compensaciones > kpi.ventaBruta * 0.05)
    recosFact.push({ prio: "alta", titulo: "Reducir compensaciones", texto: `Las compensaciones (${fmtV(kpi.compensaciones)}, ${(kpi.compensaciones/kpi.ventaBruta*100).toFixed(1)}% de ventas) superan el umbral del 5%. Revisar stock, tiempos de preparación y pedidos incompletos con el aliado.` });

  // ── Recomendaciones Fiscales (data-driven) ───────────────────────────────
  const recosFisc = [];
  if (kpi.hasDar)
    recosFisc.push({ prio: "alta", titulo: "Emitir Nota de Crédito DAR", texto: `Con DAR activo, el aliado debe emitir la Nota de Crédito correspondiente ante ${darCfg.organismo} (${darCfg.norma}) dentro del plazo: ${darCfg.plazoNC}. Esto reduce la base del ${taxCfg.iva} y optimiza la carga fiscal del período.` });
  if (kpi.effectiveFee > 0.35)
    recosFisc.push({ prio: "alta", titulo: "Tarifa efectiva elevada — acción requerida", texto: `La tarifa efectiva del ${feePct}% supera el umbral recomendado del 35%. Verificar si el aliado tiene certificados de exención o reducción de retenciones vigentes ante ${darCfg.organismo}.` });
  if (taxPctVentas > 0.15)
    recosFisc.push({ prio: "alta", titulo: "Carga impositiva alta", texto: `Los impuestos representan el ${(taxPctVentas*100).toFixed(1)}% de las ventas brutas — por encima del promedio del sector. Revisar con el contador del aliado si existen regímenes especiales aplicables en ${country}.` });
  if (kpi.ajustesTotal && kpi.ajustesTotal > kpi.ventaBruta * 0.05)
    recosFisc.push({ prio: "media", titulo: "Ajustes contables elevados", texto: `Los ajustes del período (${fmtV(kpi.ajustesTotal)}) superan el 5% de las ventas. Validar con el equipo contable si corresponden a correcciones de liquidaciones anteriores.` });
  if (!recosFisc.length)
    recosFisc.push({ prio: "baja", titulo: "Situación fiscal en orden", texto: `Los indicadores fiscales del período se encuentran dentro de rangos normales para ${country}. Se recomienda mantener la documentación actualizada.` });

  // ── Helpers de prioridad ─────────────────────────────────────────────────
  const prioLabel  = { alta: "PRIORITARIO", media: "RECOMENDADO", baja: "PREVENTIVO" };
  const prioCss    = { alta: "reco-alta",   media: "reco-media",  baja: "reco-baja"  };

  // ── Render HTML ───────────────────────────────────────────────────────────
  const win = window.open("", "_blank");
  if (!win) { alert("El navegador bloqueó la ventana emergente. Permite popups para este sitio."); return; }

  const bars = [
    { label: "Ventas Brutas",       val: kpi.ventaBruta,            color: "#22c55e" },
    { label: "Total a Pagar",       val: totalAPagar,               color: "#f59e0b" },
    { label: "Comisión Plataforma", val: Math.abs(kpi.comision),    color: "#ef4444" },
    { label: "Total Impuestos",     val: totalImpuestos,            color: "#0ea5e9" },
    ...(kpi.hasDar           ? [{ label: "Inversión DAR",   val: kpi.darInversionTotal, color: "#f97316" }] : []),
    ...(kpi.cuotaRappiAds > 0? [{ label: "Cuota RappiAds", val: kpi.cuotaRappiAds,     color: "#7c3aed" }] : []),
    ...(kpi.compensaciones>0 ? [{ label: "Compensaciones",  val: kpi.compensaciones,    color: "#8b5cf6" }] : []),
  ];
  const base = kpi.ventaBruta || 1;

  win.document.write(`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">
<title>Informe Financiero — ${pdfSafe(p.meta.tienda)}</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:'Segoe UI',system-ui,Arial,sans-serif;font-size:13px;color:#0f172a;background:#f1f5f9}
.page{max-width:900px;margin:0 auto;background:white}

/* ── Header ── */
.hdr{background:linear-gradient(135deg,#ff441f 0%,#ff6b47 55%,#c2410c 100%);padding:36px 44px 32px;color:white;position:relative;overflow:hidden}
.hdr::after{content:'';position:absolute;top:-60px;right:-60px;width:260px;height:260px;border-radius:50%;background:rgba(255,255,255,0.05);pointer-events:none}
.hdr-top{display:flex;align-items:center;gap:14px;margin-bottom:22px}
.hdr-mark{height:48px;border-radius:12px;background:white;display:flex;align-items:center;justify-content:center;flex-shrink:0;padding:0 12px;box-shadow:0 2px 8px rgba(0,0,0,0.15)}
.hdr-brand{font-size:11px;font-weight:800;opacity:.8;letter-spacing:.1em;text-transform:uppercase}
.hdr h1{font-size:28px;font-weight:900;letter-spacing:-0.02em;line-height:1.1;margin-bottom:10px}
.hdr-meta{font-size:12px;opacity:.85;line-height:2.1}
.hdr-tags{display:flex;gap:8px;flex-wrap:wrap;margin-top:14px}
.tag{display:inline-flex;align-items:center;gap:5px;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.25);border-radius:20px;padding:4px 14px;font-size:11px;font-weight:700}

/* ── Body ── */
.body{padding:0 44px 48px}

/* ── Section ── */
.sec{margin-top:36px}
.sec-title{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:#94a3b8;margin-bottom:16px;display:flex;align-items:center;gap:10px}
.sec-title::after{content:'';flex:1;height:1px;background:#e2e8f0}

/* ── Executive summary ── */
.exec{background:linear-gradient(135deg,#fff7ed,#fffbf5);border:1.5px solid #fed7aa;border-radius:14px;padding:22px 26px;font-size:13px;color:#1e293b;line-height:1.85}
.exec strong{color:#c2410c}

/* ── KPI grid ── */
.kpi-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:12px}
.kpi{border:1.5px solid #e2e8f0;border-radius:13px;padding:16px 18px;background:#f8fafc}
.kpi-label{font-size:9px;font-weight:800;color:#94a3b8;text-transform:uppercase;letter-spacing:.08em;margin-bottom:7px}
.kpi-val{font-size:19px;font-weight:900;line-height:1;letter-spacing:-0.01em}
.kpi-sub{font-size:10px;color:#94a3b8;margin-top:5px;font-weight:500}
.kpi-warn{background:#fff7ed;border-color:#fbd38d}

/* ── Bar chart ── */
.bars{display:flex;flex-direction:column;gap:10px}
.bar-row{display:grid;grid-template-columns:180px 1fr 120px;align-items:center;gap:14px}
.bar-lbl{font-size:11px;color:#64748b;text-align:right;font-weight:500}
.bar-bg{background:#f1f5f9;border-radius:6px;height:16px;overflow:hidden}
.bar-fill{height:100%;border-radius:6px}
.bar-val{font-size:12px;font-weight:800;text-align:right}

/* ── Fiscal ── */
.fiscal-pct{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px}
.fpct{background:#f0f9ff;border:1.5px solid #bae6fd;border-radius:14px;padding:18px;text-align:center}
.fpct-val{font-size:32px;font-weight:900;color:#0369a1;line-height:1}
.fpct-lbl{font-size:11px;color:#0284c7;margin-top:6px;font-weight:600}
.tax-row{display:flex;justify-content:space-between;align-items:center;padding:11px 16px;background:#f8fafc;border-radius:10px;border:1px solid #e2e8f0;margin-bottom:7px}
.tax-name{font-size:12px;color:#334155;font-weight:500}
.tax-pct-badge{font-size:10px;background:#e0f2fe;color:#0369a1;border-radius:12px;padding:2px 8px;font-weight:700;margin-left:8px}
.tax-val{font-size:13px;font-weight:800;color:#0369a1}
.tax-total{background:#e0f2fe;border-color:#7dd3fc}
.tax-total .tax-name{font-weight:800;color:#0369a1;font-size:13px}
.tax-total .tax-val{font-size:15px}

/* ── Recommendations ── */
.reco{border-radius:13px;padding:18px 22px;margin-bottom:10px;border-left:4px solid}
.reco-head{display:flex;justify-content:space-between;align-items:center;margin-bottom:7px}
.reco-title{font-size:13px;font-weight:800;color:#0f172a}
.reco-badge{font-size:9px;font-weight:800;padding:3px 10px;border-radius:20px;letter-spacing:.05em}
.reco-text{font-size:12px;color:#334155;line-height:1.7}
.reco-alta{background:#fef2f2;border-color:#ef4444}
.reco-alta .reco-badge{background:#fee2e2;color:#dc2626}
.reco-media{background:#fff7ed;border-color:#f97316}
.reco-media .reco-badge{background:#ffedd5;color:#c2410c}
.reco-baja{background:#f0fdf4;border-color:#22c55e}
.reco-baja .reco-badge{background:#dcfce7;color:#166534}

/* ── Footer ── */
.ftr{margin:0 44px;padding:20px 0;border-top:1px solid #e2e8f0;display:flex;justify-content:space-between;align-items:center;font-size:10px;color:#94a3b8}
.ftr-brand{font-weight:800;color:#ff441f;font-size:11px}

/* ── Print ── */
.no-print{margin:20px 44px 0;display:flex;gap:10px}
@media print{
  body{background:white}
  .no-print{display:none!important}
  .page{max-width:100%}
  .body{padding:0 32px 32px}
  .hdr{padding:24px 32px 22px}
  .kpi-grid{grid-template-columns:repeat(4,1fr)}
  .sec{margin-top:24px}
}
.internal-banner{background:#1e293b;color:white;padding:10px 28px;display:flex;align-items:center;gap:12px;font-size:11px;border-bottom:3px solid #ff441f}
.internal-banner .ib-badge{background:#ff441f;color:white;font-weight:900;font-size:10px;padding:3px 10px;border-radius:4px;white-space:nowrap;letter-spacing:.04em;flex-shrink:0}
.internal-banner .ib-text{font-weight:700;letter-spacing:.03em;text-transform:uppercase}
.internal-banner .ib-sub{color:#94a3b8;font-weight:400;font-size:10px;text-transform:none;letter-spacing:0;margin-left:8px}
@media print{.internal-banner{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style></head><body>
<div class="page">

<!-- BANNER INTERNO — NO ENVIAR AL ALIADO -->
<div class="internal-banner">
  <span class="ib-badge">⚠ USO INTERNO</span>
  <span class="ib-text">NO ENVIAR ESTE DOCUMENTO AL ALIADO — NO ES UN DOCUMENTO OFICIAL RAPPI<span class="ib-sub">Exclusivo para análisis interno del farmer · No compartir con el aliado</span></span>
</div>

<!-- HEADER -->
<div class="hdr">
  <div class="hdr-top">
    <div class="hdr-mark"><svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 130 44" width="104" height="34"><defs><linearGradient id="rg" x1="0" y1="0" x2="1" y2="0"><stop offset="0%" stop-color="#ff441f"/><stop offset="100%" stop-color="#ff6b35"/></linearGradient></defs><text x="2" y="36" font-family="Arial Black,Arial,sans-serif" font-weight="900" font-size="38" fill="url(#rg)" letter-spacing="-1">rappi</text></svg></div>
    <div class="hdr-brand">Rappi Paidlot Auditor Pro</div>
  </div>
  <h1>${pdfSafe(p.meta.tienda)}</h1>
  <div class="hdr-meta">
    ${cfg.flag ?? ""} ${country} &nbsp;·&nbsp; Paidlot <strong>${pdfSafe(p.meta.paidlotId)}</strong> &nbsp;·&nbsp; Período: ${p.resumen.inicio} → ${p.resumen.fin}<br>
    Fecha de pago: <strong>${p.resumen.fechaPago}</strong> &nbsp;·&nbsp; Generado el ${new Date().toLocaleDateString("es", { day:"2-digit", month:"long", year:"numeric" })}
  </div>
  <div class="hdr-tags">
    ${kpi.hasDar ? '<span class="tag">DAR Activo</span>' : '<span class="tag">Sin DAR</span>'}
    ${kpi.cuotaRappiAds > 0 ? '<span class="tag">RappiAds Activo</span>' : ''}
    <span class="tag">${kpi.ordenes} ordenes</span>
    <span class="tag">Tarifa efectiva ${feePct}%${kpi.effectiveFee > 0.35 ? ' ALTA' : ''}</span>
  </div>
</div>

<!-- PRINT BUTTON -->
<div class="no-print">
  <button onclick="window.print()" style="padding:10px 22px;background:#ff441f;color:white;border:none;border-radius:9px;font-size:13px;font-weight:700;cursor:pointer">Imprimir / Guardar como PDF</button>
</div>

<div class="body">

<!-- NOTAS PARA EL FARMER -->
<div class="sec" style="margin-top:20px">
  <div class="sec-title" style="display:flex;align-items:center;gap:8px">
    📋 Notas para el Farmer
    <span style="background:#1e293b;color:white;font-size:9px;font-weight:800;padding:2px 8px;border-radius:4px;letter-spacing:.05em">USO INTERNO</span>
  </div>
  <div style="background:#f8fafc;border:1.5px solid #e2e8f0;border-radius:12px;padding:16px 20px;font-size:12px;line-height:1.8;color:#334155">
    ${isMultiPeriod ? `
    <p style="margin:0 0 10px"><strong>Análisis consolidado de ${sortedPeriods.length} períodos consecutivos</strong> — ${pdfSafe(p.meta.tienda)} &nbsp;·&nbsp; ${cfg.flag ?? ""} ${country}</p>
    <p style="margin:0 0 8px">Este informe cubre desde <strong>${multiTotals.inicio}</strong> hasta <strong>${multiTotals.fin}</strong>, sumando un total de <strong>${sortedPeriods.length} liquidaciones</strong> del mismo aliado. Se pueden analizar como un único período comercial para evaluar el desempeño mensual o quincenal completo.</p>
    <table style="width:100%;border-collapse:collapse;margin:10px 0;font-size:12px">
      <tr style="background:#fff"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569;width:40%">Ventas brutas totales</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:800;color:#0f172a">${fmtV(multiTotals.ventaBruta)}</td></tr>
      <tr style="background:#f8fafc"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569">Total a pagar al aliado</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:800;color:#16a34a">${fmtV(multiTotals.totalAPagar)}</td></tr>
      <tr style="background:#fff"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569">Comisión total</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:400;color:#0f172a">${fmtV(multiTotals.comision)}</td></tr>
      <tr style="background:#f8fafc"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569">Impuestos totales</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:400;color:#0f172a">${fmtV(multiTotals.impuestos)} (${multiTotals.ventaBruta > 0 ? ((multiTotals.impuestos/multiTotals.ventaBruta)*100).toFixed(1) : 0}% s/ventas)</td></tr>
      <tr style="background:#fff"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569">Inversión DAR total</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:400;color:#0f172a">${multiTotals.darInversion > 0 ? fmtV(multiTotals.darInversion) : "Sin DAR activo"}</td></tr>
      <tr style="background:#f8fafc"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569">RappiAds total</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:400;color:#0f172a">${multiTotals.ads > 0 ? fmtV(multiTotals.ads) : "Sin inversión"}</td></tr>
      <tr style="background:#fff"><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:600;color:#475569">Órdenes totales</td><td style="padding:5px 10px;border:1px solid #e2e8f0;font-weight:400;color:#0f172a">${multiTotals.ordenes} órdenes en ${sortedPeriods.length} períodos</td></tr>
    </table>
    <p style="margin:8px 0 0;font-size:11px;color:#64748b">⬇ El detalle por período aparece en la tabla de comparación. El resumen ejecutivo corresponde al período más reciente cargado (${p.resumen.inicio} → ${p.resumen.fin}).</p>
    ` : `
    <p style="margin:0 0 8px"><strong>Período único</strong> — ${pdfSafe(p.meta.tienda)} &nbsp;·&nbsp; ${cfg.flag ?? ""} ${country}</p>
    <p style="margin:0 0 8px">Este informe corresponde a la liquidación del período <strong>${p.resumen.inicio}</strong> al <strong>${p.resumen.fin}</strong>, con fecha de pago <strong>${p.resumen.fechaPago}</strong>. Incluye <strong>${kpi.ordenes} órdenes</strong> procesadas, ventas brutas de <strong>${fmtV(kpi.ventaBruta)}</strong> y un neto a pagar de <strong>${fmtV(totalAPagar)}</strong>.</p>
    <p style="margin:0;font-size:11px;color:#64748b">Los KPIs, recomendaciones y alertas detalladas se encuentran en las secciones siguientes de este informe.</p>
    `}
  </div>
</div>

<!-- ANÁLISIS IA (si hay) -->
${aiInsights ? `
<div class="sec">
  <div class="sec-title">Análisis IA · Groq Llama 3.3</div>
  <div class="exec" style="background:linear-gradient(135deg,#eff6ff,#f0fdf4);border-color:#93c5fd">
    <div style="font-size:10px;font-weight:800;color:#2563eb;letter-spacing:.08em;text-transform:uppercase;margin-bottom:8px">🤖 Análisis generado por IA</div>
    ${aiInsights.split("\n").filter(l => l.trim()).map(l => `<p style="margin:0 0 6px;font-size:12px;line-height:1.7;color:#1e293b">${l}</p>`).join("")}
  </div>
</div>` : ""}

<!-- COMPARACIÓN DE PERÍODOS (multi-paidlot) -->
${isMultiPeriod ? `
<div class="sec">
  <div class="sec-title">Comparación de Períodos (${sortedPeriods.length} períodos)</div>
  <table style="width:100%;border-collapse:collapse;font-size:11px">
    <thead>
      <tr style="background:#f8fafc">
        <th style="padding:8px 10px;text-align:left;font-weight:700;color:#64748b;border-bottom:1.5px solid #e2e8f0">Indicador</th>
        ${sortedPeriods.map(sp => `<th style="padding:8px 10px;text-align:right;font-weight:700;color:#0f172a;border-bottom:1.5px solid #e2e8f0;white-space:nowrap">${sp.resumen.inicio}<br><span style="color:#94a3b8;font-weight:400">→ ${sp.resumen.fin}</span></th>`).join("")}
        <th style="padding:8px 10px;text-align:right;font-weight:800;color:#ff441f;border-bottom:1.5px solid #e2e8f0;background:#fff7ed">Δ Total</th>
      </tr>
    </thead>
    <tbody>
      ${[
        ["Ventas Brutas",   sp => sp.topKpis.ventaBruta,                                          true],
        ["Total a Pagar",   sp => sp.topKpis.totalAPagar ?? sp.topKpis.neto,                      true],
        ["Comisión",        sp => sp.topKpis.comision,                                            false],
        ["Impuestos",       sp => sp.topKpis.impuestosTotalExacto ?? sp.topKpis.totalImpuestos ?? 0, false],
        ["Inversión DAR",   sp => sp.topKpis.darInversionTotal,                                   true],
        ["Órdenes",         sp => sp.ordersTable.length,                                          true],
      ].map(([label, fn, higherIsBetter], ri) => {
        const vals = sortedPeriods.map(fn);
        const first = vals[0], last = vals[vals.length-1];
        const delta = typeof first === "number" && first !== 0 ? ((last - first) / Math.abs(first)) * 100 : null;
        const color = delta === null ? "#94a3b8" : (higherIsBetter ? (delta >= 0 ? "#10b981" : "#ef4444") : (delta <= 0 ? "#10b981" : "#ef4444"));
        return `<tr style="background:${ri%2===0?"white":"#fafafa"};border-bottom:1px solid #f1f5f9">
          <td style="padding:7px 10px;font-weight:600;color:#475569">${label}</td>
          ${sortedPeriods.map((sp, i) => `<td style="padding:7px 10px;text-align:right;font-weight:${i===sortedPeriods.length-1?800:400}">${typeof vals[i]==="number"&&vals[i]>1?fmtV(vals[i]):vals[i]}</td>`).join("")}
          <td style="padding:7px 10px;text-align:right;background:#fff7ed;font-weight:800;color:${color}">${delta!==null?(delta>=0?"+":"")+delta.toFixed(1)+"%":"—"}</td>
        </tr>`;
      }).join("")}
    </tbody>
  </table>
</div>` : ""}

<!-- RESUMEN EJECUTIVO -->
<div class="sec">
  <div class="sec-title">Resumen Ejecutivo</div>
  <div class="exec">
    Durante el período del <strong>${p.resumen.inicio}</strong> al <strong>${p.resumen.fin}</strong>,
    <strong>${pdfSafe(p.meta.tienda)}</strong> registró <strong>${kpi.ordenes} órdenes</strong> con ventas brutas de
    <strong>${fmtV(kpi.ventaBruta)}</strong>. El importe neto a transferir es <strong>${fmtV(totalAPagar)}</strong>,
    con una tarifa efectiva de <strong style="color:${kpi.effectiveFee > 0.35 ? "#dc2626" : "#166534"}">${feePct}%</strong>
    ${kpi.effectiveFee > 0.35 ? "— <strong style='color:#dc2626'>por encima del umbral recomendado del 35%</strong>" : "— dentro del rango saludable"}.
    ${kpi.hasDar ? `Rappi invirtió <strong>${fmtV(kpi.darInversionTotal)}</strong> en descuentos DAR, generando demanda sin afectar el neto del aliado.` : "No hay inversión DAR activa en este período."}
    ${totalImpuestos > 0 ? `La carga impositiva es de <strong>${fmtV(totalImpuestos)}</strong>, representando el <strong>${(taxPctVentas*100).toFixed(1)}%</strong> de las ventas brutas.` : ""}
  </div>
</div>

<!-- RESUMEN FINANCIERO -->
<div class="sec">
  <div class="sec-title">Resumen Financiero</div>
  <div class="kpi-grid">
    <div class="kpi">
      <div class="kpi-label">Ventas Brutas</div>
      <div class="kpi-val" style="color:#22c55e">${fmtV(kpi.ventaBruta)}</div>
      <div class="kpi-sub">${kpi.ordenes} órdenes</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Total a Pagar</div>
      <div class="kpi-val" style="color:#f59e0b">${fmtV(totalAPagar)}</div>
      <div class="kpi-sub">Neto transferido</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Comisión Plataforma</div>
      <div class="kpi-val" style="color:#ef4444">${fmtV(Math.abs(kpi.comision))}</div>
      <div class="kpi-sub">Uso y alquiler Rappi</div>
    </div>
    <div class="kpi ${kpi.effectiveFee > 0.35 ? "kpi-warn" : ""}">
      <div class="kpi-label">Tarifa Efectiva</div>
      <div class="kpi-val" style="color:${kpi.effectiveFee > 0.35 ? "#c2410c" : "#10b981"}">${feePct}%</div>
      <div class="kpi-sub">${kpi.effectiveFee > 0.35 ? "⚠ Supera el 35%" : "✓ Rango normal"}</div>
    </div>
    <div class="kpi">
      <div class="kpi-label">Total Impuestos</div>
      <div class="kpi-val" style="color:#0ea5e9">${fmtV(totalImpuestos)}</div>
      <div class="kpi-sub">${(taxPctVentas*100).toFixed(1)}% de ventas brutas</div>
    </div>
    <div class="kpi" style="${kpi.hasDar ? "background:#fff7ed;border-color:#fbd38d" : ""}">
      <div class="kpi-label">Inversión DAR</div>
      <div class="kpi-val" style="color:#f97316">${fmtV(kpi.darInversionTotal)}</div>
      <div class="kpi-sub">${kpi.hasDar ? "Activo este período" : "Sin DAR activo"}</div>
    </div>
    ${kpi.cuotaRappiAds > 0 ? `<div class="kpi"><div class="kpi-label">Cuota RappiAds</div><div class="kpi-val" style="color:#7c3aed">${fmtV(kpi.cuotaRappiAds)}</div><div class="kpi-sub">${(adsPct*100).toFixed(1)}% de ventas</div></div>` : ""}
    ${kpi.compensaciones > 0 ? `<div class="kpi"><div class="kpi-label">Compensaciones</div><div class="kpi-val" style="color:#8b5cf6">${fmtV(kpi.compensaciones)}</div><div class="kpi-sub">${(kpi.compensaciones/kpi.ventaBruta*100).toFixed(1)}% de ventas</div></div>` : ""}
  </div>
</div>

<!-- COMPOSICIÓN -->
<div class="sec">
  <div class="sec-title">Composición sobre Venta Bruta</div>
  <div class="bars">
    ${bars.map(b => `
    <div class="bar-row">
      <div class="bar-lbl">${b.label}</div>
      <div class="bar-bg"><div class="bar-fill" style="width:${Math.min(100,(b.val/base*100)).toFixed(1)}%;background:${b.color}"></div></div>
      <div class="bar-val" style="color:${b.color}">${fmtV(b.val)}</div>
    </div>`).join("")}
  </div>
</div>

<!-- ANÁLISIS FISCAL -->
<div class="sec">
  <div class="sec-title">Análisis Fiscal</div>
  <div class="fiscal-pct">
    <div class="fpct">
      <div class="fpct-val">${(taxPctVentas*100).toFixed(1)}%</div>
      <div class="fpct-lbl">de las Ventas Brutas</div>
    </div>
    <div class="fpct">
      <div class="fpct-val">${totalAPagar !== 0 ? (totalImpuestos/Math.abs(totalAPagar)*100).toFixed(1) : "0.0"}%</div>
      <div class="fpct-lbl">del Total a Pagar</div>
    </div>
  </div>
  ${taxRows.length > 0
    ? taxRows.map(r => `<div class="tax-row"><div class="tax-name">${pdfSafe(r.name)}<span class="tax-pct-badge">${r.pct ?? ""}</span></div><div class="tax-val">${fmtV(r.value)}</div></div>`).join("") +
      `<div class="tax-row tax-total"><div class="tax-name">TOTAL IMPUESTOS</div><div class="tax-val">${fmtV(totalImpuestos)}</div></div>`
    : `<div class="tax-row"><div class="tax-name" style="color:#94a3b8">No se detectaron impuestos en este período</div></div>`
  }
</div>

<!-- RECOMENDACIONES DE FACTURACIÓN -->
<div class="sec">
  <div class="sec-title">Recomendaciones de Facturación</div>
  ${recosFact.length === 0
    ? `<div class="reco reco-baja"><div class="reco-head"><div class="reco-title">Operación en parámetros óptimos</div><span class="reco-badge">PREVENTIVO</span></div><div class="reco-text">Los indicadores de facturación se encuentran dentro de rangos saludables. Mantener la estrategia actual.</div></div>`
    : recosFact.map(r => `<div class="reco ${prioCss[r.prio]}"><div class="reco-head"><div class="reco-title">${r.titulo}</div><span class="reco-badge">${prioLabel[r.prio]}</span></div><div class="reco-text">${r.texto}</div></div>`).join("")}
</div>

<!-- RECOMENDACIONES FISCALES -->
<div class="sec">
  <div class="sec-title">Recomendaciones Fiscales</div>
  ${recosFisc.map(r => `<div class="reco ${prioCss[r.prio]}"><div class="reco-head"><div class="reco-title">${r.titulo}</div><span class="reco-badge">${prioLabel[r.prio]}</span></div><div class="reco-text">${r.texto}</div></div>`).join("")}
</div>

</div><!-- /body -->

<!-- FOOTER -->
<div class="ftr">
  <div><span class="ftr-brand">Rappi Paidlot Auditor Pro</span> &nbsp;·&nbsp; ${new Date().toLocaleString("es")}</div>
  <div>Documento de uso interno &nbsp;·&nbsp; No reemplaza la liquidación oficial de Rappi</div>
</div>

</div><!-- /page -->
</body></html>`);
  win.document.close();
}



// ── Google Sheets logging (fire-and-forget) ───────────────────────────────────
async function logQueryToSheets({ aliado, pais, pregunta, respuesta }) {
  const url = import.meta.env.VITE_SHEETS_SCRIPT_URL;
  if (!url || url.includes("TU_SCRIPT_ID")) return;
  try {
    await fetch(url, {
      method: "POST",
      mode: "no-cors",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ aliado, pais, pregunta, respuesta, ts: new Date().toISOString() }),
    });
  } catch {}
}

// ── AutoAlertsBanner — muestra alertas automáticas del paidlot activo ─────────
const AUTO_ALERT_COLORS = {
  danger:  { bg: "#fef2f2", border: "#fca5a5", text: "#dc2626", badge: "#fee2e2" },
  warning: { bg: "#fffbeb", border: "#fcd34d", text: "#d97706", badge: "#fef3c7" },
  info:    { bg: "#eff6ff", border: "#93c5fd", text: "#2563eb", badge: "#dbeafe" },
};

const AutoAlertsBanner = memo(({ kpi, country }) => {
  const [open, setOpen] = useState(true);
  const alerts = useMemo(() => detectAlerts(kpi, country), [kpi, country]);
  if (!alerts.length || !open) return null;
  const dangers  = alerts.filter(a => a.type === "danger");
  const warnings = alerts.filter(a => a.type === "warning");
  const infos    = alerts.filter(a => a.type === "info");
  return (
    <div style={{ marginBottom: 12, borderRadius: 14, border: "1.5px solid #e2e8f0", overflow: "hidden", background: "white" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "10px 16px", background: "#f8fafc", borderBottom: "1px solid #e2e8f0", cursor: "pointer" }} onClick={() => setOpen(v => !v)}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 16 }}>🔔</span>
          <span style={{ fontWeight: 800, fontSize: 13, color: "#0f172a" }}>Alertas automáticas</span>
          {dangers.length  > 0 && <span style={{ background: "#fee2e2", color: "#dc2626", fontSize: 10, fontWeight: 800, padding: "2px 8px", borderRadius: 20 }}>{dangers.length} crítica{dangers.length > 1 ? "s" : ""}</span>}
          {warnings.length > 0 && <span style={{ background: "#fef3c7", color: "#d97706", fontSize: 10, fontWeight: 800, padding: "2px 8px", borderRadius: 20 }}>{warnings.length} alerta{warnings.length > 1 ? "s" : ""}</span>}
          {infos.length    > 0 && <span style={{ background: "#dbeafe", color: "#2563eb", fontSize: 10, fontWeight: 800, padding: "2px 8px", borderRadius: 20 }}>{infos.length} oportunidad{infos.length > 1 ? "es" : ""}</span>}
        </div>
        <span style={{ fontSize: 11, color: "#94a3b8", fontWeight: 600 }}>{open ? "▲ Ocultar" : "▼ Ver"}</span>
      </div>
      {open && (
        <div style={{ padding: "12px 16px", display: "flex", flexDirection: "column", gap: 8 }}>
          {alerts.map((a, i) => {
            const c = AUTO_ALERT_COLORS[a.type];
            return (
              <div key={i} style={{ display: "flex", alignItems: "flex-start", gap: 10, background: c.bg, border: `1px solid ${c.border}`, borderRadius: 10, padding: "10px 14px" }}>
                <span style={{ fontSize: 18, flexShrink: 0, marginTop: 1 }}>{a.icon}</span>
                <div>
                  <div style={{ fontWeight: 700, fontSize: 12, color: c.text, marginBottom: 2 }}>{a.title}</div>
                  <div style={{ fontSize: 11, color: "#475569", lineHeight: 1.6 }}>{a.msg}</div>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
});

// ── ComparisonPanel — tabla de KPIs comparando múltiples períodos ─────────────
const ComparisonPanel = memo(({ paidlots, country }) => {
  if (paidlots.length < 2) return null;
  const cfg = CONFIG.countries[country] ?? CONFIG.countries["No detectado"];
  const fmtV = (v) => { try { return new Intl.NumberFormat(cfg.locale, { style: "currency", currency: cfg.currency, maximumFractionDigits: 0 }).format(v ?? 0); } catch { return (v ?? 0).toFixed(0); } };
  const pct = (v) => `${(v * 100).toFixed(1)}%`;

  const sorted = [...paidlots].sort((a, b) => (a.resumen.inicio ?? "").localeCompare(b.resumen.inicio ?? ""));
  const base = sorted[0];

  const rows = [
    { label: "Ventas Brutas",      key: p => p.topKpis.ventaBruta,                          fmt: fmtV },
    { label: "Total a Pagar",      key: p => p.topKpis.totalAPagar ?? p.topKpis.neto,        fmt: fmtV },
    { label: "Comisión",           key: p => p.topKpis.comision,                             fmt: fmtV },
    { label: "Impuestos",          key: p => p.topKpis.impuestosTotalExacto ?? p.topKpis.totalImpuestos ?? 0, fmt: fmtV },
    { label: "Tarifa Efectiva",    key: p => p.topKpis.effectiveFee,                         fmt: pct },
    { label: "Inversión DAR",      key: p => p.topKpis.darInversionTotal,                    fmt: fmtV },
    { label: "Cuota RappiAds",     key: p => p.topKpis.cuotaRappiAds ?? 0,                   fmt: fmtV },
    { label: "Compensaciones",     key: p => p.topKpis.compensaciones,                       fmt: fmtV },
    { label: "Órdenes",            key: p => p.ordersTable.length,                            fmt: v => v },
  ];

  // Totals row
  const totals = {
    ventas:    sorted.reduce((s, p) => s + (p.topKpis.ventaBruta ?? 0), 0),
    neto:      sorted.reduce((s, p) => s + (p.topKpis.totalAPagar ?? p.topKpis.neto ?? 0), 0),
    ordenes:   sorted.reduce((s, p) => s + p.ordersTable.length, 0),
  };

  return (
    <div style={{ background: "white", borderRadius: 14, border: "1.5px solid #e2e8f0", overflow: "hidden", marginBottom: 16 }}>
      <div style={{ padding: "12px 16px", background: "linear-gradient(135deg,#fff7ed,#fffbf5)", borderBottom: "1px solid #fed7aa", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 18 }}>📊</span>
          <span style={{ fontWeight: 800, fontSize: 14, color: "#0f172a" }}>Comparación de períodos</span>
          <span style={{ fontSize: 11, color: "#94a3b8" }}>{sorted.length} períodos seleccionados</span>
        </div>
        <div style={{ fontSize: 12, color: "#64748b" }}>
          Total consolidado: <strong style={{ color: "#ff441f" }}>{fmtV(totals.ventas)}</strong> ventas · <strong style={{ color: "#10b981" }}>{fmtV(totals.neto)}</strong> a pagar · <strong>{totals.ordenes}</strong> órdenes
        </div>
      </div>
      <div style={{ overflowX: "auto" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead>
            <tr style={{ background: "#f8fafc" }}>
              <th style={{ padding: "10px 16px", textAlign: "left", fontWeight: 700, color: "#64748b", fontSize: 11, borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap" }}>Indicador</th>
              {sorted.map(p => (
                <th key={p.id} style={{ padding: "10px 14px", textAlign: "right", fontWeight: 700, color: "#0f172a", fontSize: 11, borderBottom: "1px solid #e2e8f0", whiteSpace: "nowrap" }}>
                  {p.resumen.inicio}<br /><span style={{ fontWeight: 500, color: "#94a3b8" }}>→ {p.resumen.fin}</span>
                </th>
              ))}
              <th style={{ padding: "10px 14px", textAlign: "right", fontWeight: 800, color: "#ff441f", fontSize: 11, borderBottom: "1px solid #e2e8f0", background: "#fff7ed", whiteSpace: "nowrap" }}>Δ vs anterior</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((row, ri) => (
              <tr key={ri} style={{ borderBottom: "1px solid #f1f5f9", background: ri % 2 === 0 ? "white" : "#fafafa" }}>
                <td style={{ padding: "9px 16px", fontWeight: 600, color: "#475569", whiteSpace: "nowrap" }}>{row.label}</td>
                {sorted.map((p, pi) => {
                  const val = row.key(p);
                  const prev = pi > 0 ? row.key(sorted[pi - 1]) : null;
                  const delta = prev !== null && typeof val === "number" && typeof prev === "number" ? val - prev : null;
                  const isGood = delta !== null && (row.label === "Total a Pagar" || row.label === "Ventas Brutas" || row.label === "Órdenes" || row.label === "Inversión DAR") ? delta > 0 : delta !== null ? delta < 0 : null;
                  return (
                    <td key={p.id} style={{ padding: "9px 14px", textAlign: "right", fontWeight: pi === sorted.length - 1 ? 800 : 500, color: "#0f172a" }}>
                      {row.fmt(val)}
                      {delta !== null && pi === sorted.length - 1 && (
                        <span style={{ marginLeft: 6, fontSize: 10, fontWeight: 700, color: isGood ? "#10b981" : "#ef4444" }}>
                          {delta > 0 ? "▲" : "▼"} {typeof delta === "number" && Math.abs(delta) > 1 ? row.fmt(Math.abs(delta)) : ""}
                        </span>
                      )}
                    </td>
                  );
                })}
                <td style={{ padding: "9px 14px", textAlign: "right", background: "#fff7ed" }}>
                  {(() => {
                    const last = row.key(sorted[sorted.length - 1]);
                    const first = row.key(sorted[0]);
                    if (typeof last !== "number" || typeof first !== "number" || first === 0) return <span style={{ color: "#94a3b8" }}>—</span>;
                    const chg = ((last - first) / Math.abs(first)) * 100;
                    const isPos = chg > 0;
                    const goodLabel = row.label === "Total a Pagar" || row.label === "Ventas Brutas" || row.label === "Órdenes" || row.label === "Inversión DAR";
                    const color = goodLabel ? (isPos ? "#10b981" : "#ef4444") : (isPos ? "#ef4444" : "#10b981");
                    return <span style={{ fontWeight: 800, color }}>{isPos ? "+" : ""}{chg.toFixed(1)}%</span>;
                  })()}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────────────────────
// § 7. UI COMPONENTS
// ─────────────────────────────────────────────────────────────────────────────

const Tooltip = memo(({ text, children }) => {
  const [show, setShow] = useState(false);
  return (
    <span style={{ position: "relative", display: "inline-flex", alignItems: "center" }} onMouseEnter={() => setShow(true)} onMouseLeave={() => setShow(false)}>
      {children}
      {show && <span style={{ position: "absolute", bottom: "calc(100% + 8px)", left: "50%", transform: "translateX(-50%)", background: "#1e293b", color: "white", borderRadius: 8, padding: "8px 12px", fontSize: 11, lineHeight: 1.5, maxWidth: 280, whiteSpace: "normal", zIndex: 500, boxShadow: "0 8px 24px rgba(0,0,0,0.2)", pointerEvents: "none" }}>{text}</span>}
    </span>
  );
});

const Badge = memo(({ label, style }) => <span style={{ fontSize: 10, fontWeight: 700, padding: "2px 8px", borderRadius: 20, ...style }}>{label}</span>);

const DataTable = memo(({ columns, rows, emptyMsg }) => (
  rows.length === 0
    ? <div style={{ padding: "28px 0", textAlign: "center", color: "#94a3b8", fontSize: 13 }}>{emptyMsg ?? "Sin datos"}</div>
    : <div style={{ overflowX: "auto", borderRadius: 12, border: "1px solid #e2e8f0" }}>
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
          <thead><tr style={{ background: "#0f172a" }}>{columns.map((c, i) => <th key={i} style={{ padding: "9px 12px", textAlign: "left", fontWeight: 700, color: "white", whiteSpace: "nowrap", fontSize: 11 }}>{c}</th>)}</tr></thead>
          <tbody>{rows.map((row, ri) => <tr key={ri} style={{ borderTop: "1px solid #e2e8f0", background: ri % 2 === 0 ? "white" : "#f8fafc" }}>{row.map((cell, ci) => <td key={ci} style={{ padding: "8px 12px", color: "#334155", verticalAlign: "top" }}>{cell}</td>)}</tr>)}</tbody>
        </table>
      </div>
));


// ── KPIGrid — driven by KPI_CONFIG registry ───────────────────────────────────
// Pure display component. Reads KPI_CONFIG for structure, topKpis for values.
// Each card is independently clickable to scroll to its section.
const KPIGrid = memo(({ topKpis, country, onSelectKpi, selectedKpi, compact = false }) => (
  <div style={{ display: "grid", gridTemplateColumns: compact ? "repeat(7,minmax(0,1fr))" : "repeat(auto-fit,minmax(148px,1fr))", gap: compact ? 6 : 10, marginBottom: compact ? 0 : 0 }}>
    {KPI_CONFIG.map(k => {
      const rawVal = topKpis[k.key] ?? 0;
      const isEffective = k.key === "effectiveFee";
      const displayVal = isEffective
        ? fmtPct(rawVal)
        : fmt(rawVal, country);
      const dynamicColor = isEffective
        ? (rawVal > 0.35 ? "#f97316" : "#10b981")
        : k.color;
      const isSelected = selectedKpi === k.key;
      return (
        <div
          key={k.key}
          onClick={() => {
            if (k.scrollTo) {
              const el = document.getElementById(k.scrollTo);
              if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
            }
            onSelectKpi && onSelectKpi(k.key);
          }}
          style={{
            background: "white",
            borderRadius: compact ? 10 : 14,
            padding: compact ? "8px 10px" : "14px 16px",
            border: `1px solid ${isSelected ? dynamicColor : (isEffective && rawVal > 0.35 ? "#fed7aa" : "#e2e8f0")}`,
            cursor: onSelectKpi ? "pointer" : "default",
            transition: "border-color 0.2s, box-shadow 0.2s",
            boxShadow: isSelected ? `0 0 0 3px ${dynamicColor}22` : "none",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", marginBottom: 6 }}>
            <span style={{ fontSize: compact ? 14 : 18 }}>{k.icon}</span>
            <Tooltip text={k.key === "effectiveFee"
              ? "Tarifa real = (Comisión + Impuestos) / Venta Bruta. Naranja si supera el 35%."
              : `${k.label} del período`}>
              <span style={{ fontSize: 11, color: "#94a3b8", cursor: "help" }}>ⓘ</span>
            </Tooltip>
          </div>
          <div style={{ fontSize: compact ? 9 : 10, color: "#64748b", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: compact ? 3 : 5 }}>{k.label}</div>
          <div style={{ fontSize: compact ? 13 : 16, fontWeight: 900, color: dynamicColor, fontVariantNumeric: "tabular-nums" }}>{displayVal}</div>
        </div>
      );
    })}
  </div>
));

// ── KPIPanel — vertical sidebar KPI list ──────────────────────────────────────
// Used in the right sidebar. Each row: icon + label/value + tooltip info button.
const KPIPanel = memo(({ topKpis, country, onSelectKpi, selectedKpi }) => (
  <div>
    <div style={{ fontSize: 10, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: "0.07em", marginBottom: 8, padding: "0 2px" }}>
      KPIs del Período
    </div>
    <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
      {KPI_CONFIG.map((k) => {
        const rawVal = topKpis[k.key] ?? 0;
        const isEffective = k.key === "effectiveFee";
        const displayVal = isEffective ? fmtPct(rawVal) : fmt(rawVal, country);
        const dynamicColor = isEffective ? (rawVal > 0.35 ? "#f97316" : "#10b981") : k.color;
        const isSelected = selectedKpi === k.key;
        const isAlert = isEffective && rawVal > 0.35;
        return (
          <div
            key={k.key}
            onClick={() => { onSelectKpi && onSelectKpi(k.key); }}
            style={{
              background: isSelected ? `${dynamicColor}0e` : isAlert ? "#fff7ed" : "white",
              borderRadius: 10,
              padding: "9px 12px",
              border: `1.5px solid ${isSelected ? dynamicColor : isAlert ? "#fed7aa" : "#e2e8f0"}`,
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              gap: 10,
              transition: "all 0.18s",
              boxShadow: isSelected ? `0 0 0 3px ${dynamicColor}18` : "none",
            }}
          >
            <span style={{ fontSize: 20, flexShrink: 0, lineHeight: 1 }}>{k.icon}</span>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={{ fontSize: 9, color: "#94a3b8", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 2, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                {k.label}
              </div>
              <div style={{ fontSize: 15, fontWeight: 900, color: dynamicColor, fontVariantNumeric: "tabular-nums", lineHeight: 1.1 }}>
                {displayVal}
              </div>
            </div>
            <Tooltip text={isEffective ? "Tarifa real = (Comisión + Impuestos) / Venta Bruta. Naranja si supera el 35%." : `${k.label} del período`}>
              <span style={{ fontSize: 11, color: "#cbd5e1", cursor: "help", flexShrink: 0 }}>ⓘ</span>
            </Tooltip>
          </div>
        );
      })}
    </div>
  </div>
));

// ── ConciliationPill — compact inline conciliation status badge ───────────────
const ConciliationPill = memo(({ data }) => {
  if (!data) return null;
  const isOk = data.status === "OK";
  return (
    <div style={{ display: "inline-flex", alignItems: "center", gap: 6, padding: "4px 12px", background: isOk ? "#f0fdf4" : "#fef2f2", border: `1px solid ${isOk ? "#86efac" : "#fca5a5"}`, borderRadius: 20, fontSize: 11, fontWeight: 700, color: isOk ? "#15803d" : "#dc2626" }}>
      {isOk ? "✅" : "⚠️"} Factura {isOk ? "leída correctamente" : `REVISAR · diff ${data.diff?.toFixed(2)}`}
    </div>
  );
});

// ── TaxRulesPanel — shows TAX_RULES exact column matches found in paidlot ────
// Gives the farmer visibility into which tax columns were matched and their values.
const TaxRulesPanel = memo(({ topKpis, country }) => {
  const [open, setOpen] = useState(false);
  const rules = topKpis.impuestosPorRegla ?? [];
  const matched = rules.filter(r => r.value > 0);
  if (matched.length === 0) return null;
  return (
    <div style={{ marginTop: 12, borderRadius: 10, border: "1px solid #bfdbfe", overflow: "hidden" }}>
      <button onClick={() => setOpen(o => !o)}
        style={{ width: "100%", background: "#eff6ff", padding: "9px 14px", border: "none", cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
        <span style={{ fontSize: 12, fontWeight: 700, color: "#1d4ed8" }}>
          🔍 Detalle por regla fiscal — {country} ({matched.length} columnas encontradas)
        </span>
        <span style={{ fontSize: 11, color: "#3b82f6" }}>{open ? "▲" : "▼"}</span>
      </button>
      {open && (
        <div style={{ background: "white", padding: "10px 14px" }}>
          {matched.map(r => (
            <div key={r.match} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "7px 0", borderBottom: "1px solid #f1f5f9", fontSize: 12 }}>
              <div style={{ flex: 1, marginRight: 12 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                  <span style={{ fontWeight: 700, color: "#1e293b" }}>{r.name}</span>
                  {r.pct && <span style={{ fontSize: 10, fontWeight: 700, background: "#e0f2fe", color: "#0369a1", borderRadius: 20, padding: "1px 7px" }}>{r.pct}</span>}
                </div>
                <div style={{ fontFamily: "monospace", fontSize: 10, color: "#94a3b8", marginTop: 2 }}>{r.match.slice(0, 60)}{r.match.length > 60 ? "…" : ""}</div>
              </div>
              <span style={{ fontWeight: 800, color: "#0369a1", whiteSpace: "nowrap" }}>{fmt(r.value, country)}</span>
            </div>
          ))}
        </div>
      )}
    </div>
  );
});

const DarKpiPanel = memo(({ kpis, country }) => {
  if (!kpis.hasDar) return null;
  return (
    <div style={{ background: "linear-gradient(135deg,#fff7ed,#ffedd5)", border: "1.5px solid #fb923c", borderRadius: 16, padding: "16px 20px", marginBottom: 16 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <span style={{ fontSize: 20 }}>🎯</span>
        <div><div style={{ fontWeight: 800, color: "#c2410c", fontSize: 14 }}>Inversión DAR Rappi — Período activo</div><div style={{ fontSize: 11, color: "#9a3412" }}>Rappi está invirtiendo en este aliado. El neto del aliado NO cambia.</div></div>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit,minmax(140px,1fr))", gap: 10 }}>
        {[
          { label: "Inversión producto", val: fmt(kpis.darInversionProducto ?? kpis.darInversionTotal, country), tip: "Descuento aplicado sobre el precio del producto al usuario." },
          { label: "Compensación comisión", val: fmt(kpis.darComisionTotal, country), tip: "Rappi devuelve la proporción de comisión correspondiente al DAR." },
          { label: "Compensación IVA/IGV", val: fmt(kpis.darIvaTotal, country), tip: "Rappi devuelve también el impuesto correspondiente." },
          { label: "Beneficio total DAR", val: fmt(kpis.darBeneficioTotal, country), tip: "Suma total del valor que Rappi invierte vía DAR." },
          { label: "DAR % sobre ventas", val: fmtPct(kpis.darPctSobreVentas), tip: "Porcentaje que representa la inversión DAR sobre la venta bruta total." },
        ].map(k => (
          <div key={k.label} style={{ background: "white", borderRadius: 10, padding: "10px 12px", border: "1px solid #fed7aa" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 4, marginBottom: 4 }}>
              <span style={{ fontSize: 10, color: "#9a3412", fontWeight: 700, textTransform: "uppercase" }}>{k.label}</span>
              <Tooltip text={k.tip}><span style={{ fontSize: 10, color: "#94a3b8", cursor: "help" }}>ⓘ</span></Tooltip>
            </div>
            <div style={{ fontSize: 15, fontWeight: 800, color: "#c2410c" }}>{k.val}</div>
          </div>
        ))}
      </div>
    </div>
  );
});

// ── AdsAlertBanner — shown when Ads/totalAPagar > 20% ────────────────────────
const AdsAlertBanner = memo(({ kpis, country }) => {
  const baseVentas = kpis.ventaBruta;
  if (!kpis.cuotaRappiAds || baseVentas <= 0) return null;
  const pct = kpis.cuotaRappiAds / baseVentas;
  if (pct <= 0.10) return null;
  return (
    <div style={{ background: "linear-gradient(135deg,#fff7ed,#fef2f2)", border: "2.5px solid #f97316", borderRadius: 16, padding: "16px 20px", marginBottom: 16, display: "flex", gap: 14, alignItems: "flex-start", animation: "fadeIn 0.3s ease" }}>
      <span style={{ fontSize: 28, flexShrink: 0 }}>⚠️</span>
      <div style={{ flex: 1 }}>
        <div style={{ fontWeight: 900, color: "#c2410c", fontSize: 15, marginBottom: 5 }}>
          Alerta de Pauta: La inversión en Ads representa el {(pct * 100).toFixed(1)}% de las ventas facturadas
        </div>
        <div style={{ fontSize: 13, color: "#9a3412", lineHeight: 1.65 }}>
          La <strong>Cuota de RappiAds</strong> ({fmt(kpis.cuotaRappiAds, country)}) supera el 10% de las <strong>Ventas Facturadas</strong> ({fmt(baseVentas, country)}). Esto podría comprometer el flujo de caja del aliado. Evalúa si el retorno de la inversión ADS justifica este nivel de pauta, o ajusta el presupuesto publicitario para el próximo período.
        </div>
      </div>
      <div style={{ background: "#f97316", color: "white", borderRadius: 12, padding: "6px 14px", fontSize: 13, fontWeight: 900, whiteSpace: "nowrap", alignSelf: "center" }}>
        {(pct * 100).toFixed(1)}% ADS
      </div>
    </div>
  );
});

const DarZeroAlert = memo(({ kpis, tienda }) => {
  if (kpis.hasDar) return null;
  return (
    <div style={{ background: "#fef2f2", border: "1.5px solid #fca5a5", borderRadius: 14, padding: "14px 18px", marginBottom: 16, display: "flex", gap: 12 }}>
      <span style={{ fontSize: 22, flexShrink: 0 }}>⚠️</span>
      <div>
        <div style={{ fontWeight: 800, color: "#dc2626", fontSize: 14, marginBottom: 4 }}>Este aliado no tiene inversión DAR activa en este período</div>
        <div style={{ fontSize: 12, color: "#7f1d1d", lineHeight: 1.6 }}><strong>{tienda}</strong> no está recibiendo descuentos de Rappi. Revisar si cumple requisitos de GMV mínimo para campañas DAR, o gestionar activación con el equipo comercial.</div>
      </div>
    </div>
  );
});

const DarOrderTable = memo(({ ordersTable, country, hasDar }) => {
  if (!hasDar) return <div style={{ textAlign: "center", color: "#94a3b8", padding: "28px 0", fontSize: 13 }}>No hay inversión DAR en las órdenes de este paidlot.</div>;
  const ordersWithDar = ordersTable.filter(o => o.darTotal > 0);
  if (!ordersWithDar.length) return <div style={{ textAlign: "center", color: "#94a3b8", padding: "28px 0", fontSize: 13 }}>El DAR está consolidado en el período, sin desglose individual por orden.</div>;
  const rows = ordersWithDar.map(o => [o.fecha, <span style={{ fontFamily: "monospace", fontSize: 11 }}>{o.ordenId}</span>, <span style={{ fontWeight: 700 }}>{fmt(o.ventaBruta, country)}</span>, <span style={{ color: "#f97316", fontWeight: 700 }}>{fmt(o.darInversion, country)}</span>, <span style={{ color: "#64748b" }}>{fmt(o.darComision, country)}</span>, <span style={{ color: "#64748b" }}>{fmt(o.darIva, country)}</span>, <span style={{ fontWeight: 800, color: "#c2410c" }}>{fmt(o.darTotal, country)}</span>, o.darPct > 0 ? <Badge label={fmtPct(o.darPct)} style={{ background: "#fff7ed", color: "#c2410c" }} /> : "—"]);
  return <DataTable columns={["Fecha", "Orden ID", "Venta Bruta", "DAR Producto", "DAR Comisión", "DAR IVA", "DAR Total", "DAR %"]} rows={rows} />;
});

const BarMiniChart = memo(({ groups }) => {
  if (!groups?.length) return null;
  const max = Math.max(...groups.map(g => g.total));
  return (
    <svg viewBox="0 0 120 56" style={{ width: "100%", height: 56 }}>
      {groups.map((g, i) => {
        const cfg = CONFIG.groups.find(c => c.key === g.key);
        const bH = max > 0 ? (g.total / max) * 40 : 0;
        const bW = Math.max(10, (120 - groups.length * 6) / groups.length);
        const x = i * (bW + 6) + 4;
        return <g key={g.key}><rect x={x} y={48 - bH} width={bW} height={bH} rx={2} fill={cfg?.color ?? "#64748b"} opacity={0.85} /><text x={x + bW / 2} y={55} textAnchor="middle" fontSize={7} fill="#64748b">{cfg?.icon ?? "·"}</text></g>;
      })}
    </svg>
  );
});

// ── Education Hub ─────────────────────────────────────────────────────────────
const EDUCATION_TABS = ["DAR", "ADS", "IMPUESTOS", "COMPENSACIONES", "CONCEPTOS", "FÓRMULAS"];

const EducationHub = memo(({ country, topKpis, embedded = false, embeddedTab = "DAR" }) => {
  const [open, setOpen] = useState(false);
  const [tab, setTab] = useState("DAR");
  const dar = DAR_CONFIG[country] ?? DAR_CONFIG["No detectado"];

  // ── helpers de estilo ────────────────────────────────────────────────────
  const card  = (bg, border) => ({ background: bg, borderRadius: 12, border: `1px solid ${border}`, padding: "12px 14px", marginBottom: 10 });
  const pill  = (bg, color)  => ({ display:"inline-block", background: bg, color, borderRadius: 20, padding: "1px 8px", fontSize: 10, fontWeight: 800, marginLeft: 6 });
  const rowFx = { display:"flex", gap: 6, flexWrap:"wrap", marginTop: 4 };
  const tag   = (c) => ({ background: c+"22", color: c, borderRadius: 6, padding: "2px 7px", fontSize: 10, fontWeight: 700 });

  // ── Contenido por tab ────────────────────────────────────────────────────
  const tabContent = {

    // ── DAR ──────────────────────────────────────────────────────────────
    DAR: (
      <div>
        <div style={card("#fff7ed","#fed7aa")}>
          <div style={{ fontWeight:800, color:"#c2410c", marginBottom:6, fontSize:13 }}>¿Qué es DAR y por qué NO reduce tu ingreso?</div>
          <div style={{ fontSize:12, color:"#9a3412", lineHeight:1.7 }}>
            <strong>DAR = Descuento Asumido por Rappi.</strong> Rappi financia descuentos al usuario para generar demanda.
            El aliado siempre factura al precio real de lista; Rappi paga la diferencia y recalcula la comisión sobre la
            venta neta. <strong>El neto transferido es idéntico con o sin DAR.</strong>
          </div>
        </div>

        <div style={{ fontSize:11, fontWeight:700, color:"#475569", textTransform:"uppercase", letterSpacing:"0.04em", margin:"14px 0 6px" }}>Las 3 columnas DAR en tu paidlot</div>
        {[
          { col:"Descuentos por inversión de Rappi DAR", desc:"Inversión en producto. Rappi absorbe el descuento sobre el precio del producto.", color:"#f97316" },
          { col:"Descuento por inversión de Rappi a aplicar sobre Uso y alquiler de plataforma Rappi DAR", desc:"Rappi descuenta su propia comisión proporcional al DAR aplicado.", color:"#f97316" },
          { col:"Descuento por inversión de Rappi a aplicar sobre el IVA Uso y alquiler de plataforma Rappi DAR", desc:"Descuento sobre el IVA de la comisión correspondiente al DAR.", color:"#f97316" },
        ].map(r => (
          <div key={r.col} style={card("#fff7ed","#fed7aa")}>
            <div style={{ fontSize:11, fontWeight:800, color:"#c2410c" }}>{r.col.length>65?r.col.slice(0,62)+"…":r.col}</div>
            <div style={{ fontSize:11, color:"#92400e", marginTop:3 }}>{r.desc}</div>
          </div>
        ))}

        <div style={card("#f0fdf4","#86efac")}>
          <div style={{ fontWeight:700, color:"#166534", fontSize:12, marginBottom:4 }}>💡 Libreto para el aliado</div>
          <div style={{ fontSize:11, color:"#14532d", lineHeight:1.7 }}>
            "El DAR aparece como descuento en el detalle, pero Rappi lo financia íntegramente.
            Tu venta neta no cambia porque la comisión también se recalcula hacia abajo en el mismo monto.
            No hay impacto en tu liquidación."
          </div>
        </div>

        {dar && dar.noteCredit && (
          <div style={{ marginTop:10 }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#475569", textTransform:"uppercase", marginBottom:6 }}>Nota Crédito — {country}</div>
            <div style={card("#f0fdf4","#86efac")}>
              <div style={{ fontSize:11, color:"#166534" }}>{dar.noteCredit}</div>
            </div>
          </div>
        )}
      </div>
    ),

    // ── ADS ──────────────────────────────────────────────────────────────
    ADS: (
      <div>
        <div style={card("#fefce8","#fde68a")}>
          <div style={{ fontWeight:800, color:"#92400e", marginBottom:6, fontSize:13 }}>RappiAds — Publicidad en plataforma</div>
          <div style={{ fontSize:12, color:"#78350f", lineHeight:1.7 }}>
            <strong>Cuota de RappiAds:</strong> costo por campañas de publicidad y visibilidad en la app.
            Se factura en el <strong>período siguiente</strong> al que se pauta (semana vencida).
            Un cobro de ADS en Abril corresponde a la inversión de la última semana de Marzo.
          </div>
        </div>

        <div style={{ fontSize:11, fontWeight:700, color:"#475569", textTransform:"uppercase", letterSpacing:"0.04em", margin:"14px 0 6px" }}>Conceptos relacionados con ADS en el paidlot</div>
        {[
          { col:"Cuota de RappiAds", desc:"Costo principal de campañas Rappi Ads pautadas por el aliado." },
          { col:"Descuento rappi_ads_invoiced_collection", desc:"Descuento por cobro de RappiAds ya facturado anteriormente (evita doble cobro)." },
          { col:"IVA Rappi Ads", desc:"IVA aplicado sobre la cuota de campañas según el país." },
          { col:"ReteIVA Rappi Ads", desc:"Retención de IVA sobre Rappi Ads (Colombia)." },
          { col:"Retefuente Rappi Ads", desc:"Retención en la fuente sobre Rappi Ads (Colombia)." },
          { col:"IVA Campañas", desc:"IVA sobre campañas (Argentina, otros países)." },
          { col:"Percepción Campañas", desc:"Percepción IIBB sobre campañas (Argentina)." },
        ].map(r => (
          <div key={r.col} style={card("#fefce8","#fde68a")}>
            <div style={{ fontSize:11, fontWeight:800, color:"#92400e" }}>{r.col}</div>
            <div style={{ fontSize:11, color:"#78350f", marginTop:3 }}>{r.desc}</div>
          </div>
        ))}

        <div style={card("#fef2f2","#fca5a5")}>
          <div style={{ fontWeight:700, color:"#991b1b", fontSize:12, marginBottom:4 }}>⚠️ Alerta de pauta alta</div>
          <div style={{ fontSize:11, color:"#7f1d1d", lineHeight:1.7 }}>
            Si <strong>Cuota RappiAds ÷ Venta Bruta {">"} 20%</strong> el aliado está invirtiendo una porción elevada de sus ventas
            en publicidad. Evalúa con el aliado si el retorno justifica el nivel de pauta o
            si conviene ajustar el presupuesto publicitario para el próximo período.
          </div>
        </div>

        <div style={card("#f0fdf4","#86efac")}>
          <div style={{ fontWeight:700, color:"#166534", fontSize:12, marginBottom:4 }}>💡 Libreto para el aliado</div>
          <div style={{ fontSize:11, color:"#14532d", lineHeight:1.7 }}>
            "El cobro de ADS que ves en este período corresponde a la inversión pautada
            la semana anterior. Si no reconoces el monto, revisemos juntos el reporte de campañas
            para confirmar la inversión que autorizaste en ese período."
          </div>
        </div>
      </div>
    ),

    // ── IMPUESTOS ─────────────────────────────────────────────────────────
    IMPUESTOS: (
      <div>
        <div style={card("#eff6ff","#bfdbfe")}>
          <div style={{ fontWeight:800, color:"#1e40af", marginBottom:6, fontSize:13 }}>Impuestos en liquidaciones Rappi</div>
          <div style={{ fontSize:12, color:"#1e3a8a", lineHeight:1.7 }}>
            Los impuestos se calculan sobre la <strong>comisión de plataforma</strong> (y otros costos de servicio), no sobre la venta bruta.
            Cada país tiene una combinación distinta de impuestos según su legislación fiscal.
          </div>
        </div>

        {[
          {
            pais:"🇨🇴 Colombia", countries:["Colombia"], bg:"#f0fdf4", border:"#86efac", color:"#166534",
            items:[
              { name:"IVA Plataforma (19%)", formula:"Comisión × 19%", reason:"R3", nota:"Impuesto al valor agregado sobre uso de plataforma." },
              { name:"ReteIVA (15%)", formula:"Comisión × 15%", reason:"Retención IVA", nota:"El aliado retiene el 15% del IVA generado." },
              { name:"Retefuente (var.)", formula:"Comisión × % contrato", reason:"R4", nota:"Retención en la fuente según actividad económica." },
              { name:"ReteICA (0.414%)", formula:"(Com.+MarketplaceFee) × 0.414%", reason:"R5", nota:"Tasa varía por ciudad. Solo Colombia." },
            ]
          },
          {
            pais:"🇦🇷 Argentina", countries:["Argentina"], bg:"#fefce8", border:"#fde68a", color:"#92400e",
            items:[
              { name:"IVA Plataforma (21%)", formula:"Comisión × 21%", reason:"R3", nota:"IVA estándar Argentina." },
              { name:"Percepción IVA (10.5%)", formula:"Comisión × 10.5%", reason:"R3", nota:"Percepción adicional de IVA." },
              { name:"Percepción IIBB CABA (3–5%)", formula:"(Com.+MarketplaceFee) × alícuota", reason:"R145", nota:"Ingresos Brutos Ciudad Autónoma de Buenos Aires." },
              { name:"Percepción IIBB Córdoba (3%)", formula:"(Com.+MarketplaceFee) × 3%", reason:"R145", nota:"Ingresos Brutos provincia de Córdoba." },
              { name:"Ret. Ganancias (6%)", formula:"Comisión × 6%", reason:"Ret.", nota:"Retención del impuesto a las ganancias." },
            ]
          },
          {
            pais:"🇲🇽 México", countries:["México"], bg:"#fdf4ff", border:"#e9d5ff", color:"#6b21a8",
            items:[
              { name:"IVA Plataforma (16%)", formula:"Comisión × 16%", reason:"R3", nota:"IVA estándar México." },
              { name:"ISR Retención (1%)", formula:"Comisión × 1%", reason:"ISR", nota:"Retención del Impuesto Sobre la Renta." },
            ]
          },
          {
            pais:"🇵🇪 Perú", countries:["Perú"], bg:"#fff7ed", border:"#fed7aa", color:"#c2410c",
            items:[
              { name:"IGV Plataforma (18%)", formula:"Comisión × 18%", reason:"R3", nota:"Impuesto General a las Ventas." },
            ]
          },
          {
            pais:"🇨🇱 Chile / 🇪🇨 Ecuador / 🇺🇾 Uruguay", countries:["Chile","Ecuador","Uruguay"], bg:"#f8fafc", border:"#e2e8f0", color:"#475569",
            items:[
              { name:"IVA Plataforma", formula:"Comisión × % local", reason:"R3", nota:"CL: 19% | EC: 15% | UY: 22%." },
            ]
          },
        ].filter(g => g.countries.some(c => c === country)).map(country => (
          <div key={country.pais} style={{ marginBottom:10 }}>
            <div style={{ fontSize:11, fontWeight:700, color:country.color, marginBottom:5, textTransform:"uppercase" }}>{country.pais}</div>
            {country.items.map(item => (
              <div key={item.name} style={card(country.bg, country.border)}>
                <div style={{ display:"flex", justifyContent:"space-between", alignItems:"flex-start" }}>
                  <div style={{ fontSize:11, fontWeight:800, color:country.color }}>{item.name}</div>
                  <span style={pill(country.color+"22", country.color)}>{item.reason}</span>
                </div>
                <div style={{ fontSize:11, color:country.color+"cc", marginTop:3 }}>
                  <strong>Fórmula:</strong> {item.formula}
                </div>
                <div style={{ fontSize:10, color:country.color+"99", marginTop:2 }}>{item.nota}</div>
              </div>
            ))}
          </div>
        ))}

        <div style={card("#fef2f2","#fca5a5")}>
          <div style={{ fontWeight:700, color:"#991b1b", fontSize:12, marginBottom:4 }}>📋 Gasto Bancario (Reason 100)</div>
          <div style={{ fontSize:11, color:"#7f1d1d", lineHeight:1.7 }}>
            <strong>Fórmula:</strong> Valor del Paidlot × % Gasto Bancario<br/>
            Se cobra a nivel de paidlot cuando el valor es positivo. Aplica todos los países.
            Reemplazó a la antigua Reason 11 (por orden).
          </div>
        </div>
      </div>
    ),

    // ── COMPENSACIONES ────────────────────────────────────────────────────
    COMPENSACIONES: (
      <div>
        <div style={card("#faf5ff","#e9d5ff")}>
          <div style={{ fontWeight:800, color:"#6b21a8", marginBottom:6, fontSize:13 }}>¿Qué son las Compensaciones?</div>
          <div style={{ fontSize:12, color:"#581c87", lineHeight:1.7 }}>
            Retribución que Rappi paga al usuario por problemas de la orden atribuibles al aliado
            (producto equivocado, faltante o en mal estado). El valor se cobra al aliado y
            depende del tipo de queja sobre el producto específico.
            <strong> Pueden aparecer días o semanas después de la orden original.</strong>
          </div>
        </div>

        <div style={{ fontSize:11, fontWeight:700, color:"#475569", textTransform:"uppercase", margin:"14px 0 6px" }}>Tipos de compensación (Reason 18)</div>
        {[
          { tipo:"product_poor", desc:"Producto en mal estado o de mala calidad.", condicion:"Condición de contrato 21" },
          { tipo:"product_difference", desc:"Producto diferente al solicitado por el usuario.", condicion:"Condición de contrato 22" },
          { tipo:"product_missing", desc:"Producto faltante en el pedido.", condicion:"Condición de contrato 23" },
          { tipo:"product_not_delivered", desc:"Producto que no fue entregado al usuario.", condicion:"Condición de contrato 24" },
        ].map(r => (
          <div key={r.tipo} style={card("#faf5ff","#e9d5ff")}>
            <div style={{ display:"flex", justifyContent:"space-between" }}>
              <div style={{ fontSize:11, fontWeight:800, color:"#6b21a8" }}>{r.tipo}</div>
              <span style={pill("#e9d5ff","#6b21a8")}>{r.condicion}</span>
            </div>
            <div style={{ fontSize:11, color:"#581c87", marginTop:3 }}>{r.desc}</div>
          </div>
        ))}

        <div style={card("#f0fdf4","#86efac")}>
          <div style={{ fontWeight:700, color:"#166534", fontSize:12, marginBottom:4 }}>Devolución de Compensaciones</div>
          <div style={{ fontSize:11, color:"#14532d", lineHeight:1.7 }}>
            Corrección positiva cuando una compensación fue cobrada de forma incorrecta.
            Aparece como "Devolucion de Compensaciones" en el detalle del paidlot.
          </div>
        </div>

        <div style={card("#fff7ed","#fed7aa")}>
          <div style={{ fontWeight:700, color:"#c2410c", fontSize:12, marginBottom:4 }}>Costo Canceladas (Reason 10)</div>
          <div style={{ fontSize:12, color:"#9a3412", lineHeight:1.7 }}>
            Porcentaje acordado en contrato que se descuenta por órdenes canceladas.
            <strong> Fórmula:</strong> (Venta Bruta + Antojos) × (1 − % cancelación)<br/>
            Si hay Meal Vouchers: Meal Vouchers × (1 − % cancelación).
          </div>
        </div>

        <div style={card("#eff6ff","#bfdbfe")}>
          <div style={{ fontWeight:700, color:"#1e40af", fontSize:12, marginBottom:4 }}>Contracargos (Reason 46)</div>
          <div style={{ fontSize:11, color:"#1e3a8a", lineHeight:1.7 }}>
            Devolución de dinero al cliente por disputa de cargo. Se genera cuando el banco
            o la pasarela de pagos reversa el cobro al usuario. El costo es asumido por el aliado.
          </div>
        </div>

        <div style={card("#fef2f2","#fca5a5")}>
          <div style={{ fontWeight:700, color:"#991b1b", fontSize:12, marginBottom:4 }}>💡 Libreto para el aliado</div>
          <div style={{ fontSize:11, color:"#7f1d1d", lineHeight:1.7 }}>
            "Cada compensación en el paidlot corresponde a una queja de usuario aprobada por Soporte.
            Puedes ver el ID de orden y el tipo de reclamo en el detalle para identificar qué pedido
            la generó. Si consideras que fue incorrecta, puedes escalarla con el ticket del Soporte."
          </div>
        </div>
      </div>
    ),

    // ── CONCEPTOS ─────────────────────────────────────────────────────────
    CONCEPTOS: (
      <div>
        <div style={card("#f8fafc","#e2e8f0")}>
          <div style={{ fontWeight:800, color:"#1e293b", marginBottom:6, fontSize:13 }}>Glosario oficial — Merchant Revenue Playbook</div>
          <div style={{ fontSize:11, color:"#475569" }}>Definiciones extraídas del Playbook interno de Rappi.</div>
        </div>
        {[
          { term:"Venta Bruta", def:"Ventas sin descuento. Valor total de productos antes de cualquier deducción." },
          { term:"Descuento de Producto asumido por el aliado", def:"Descuento que el aliado financia sobre sus propios productos (markdowns, promos del aliado)." },
          { term:"Descuento en créditos", def:"Descuentos pagados con Rappi Créditos asumidos por el aliado (campo vfdPartner en PP)." },
          { term:"Ventas base por Uso y alquiler de plataforma Rappi", def:"Base para cálculo de fee = Venta Bruta − Descuentos del aliado. Informativo." },
          { term:"Uso y alquiler de plataforma Rappi", def:"Comisión Rappi. Costo de uso de plataforma = Ventas base × % comisión del contrato." },
          { term:"Tarifa Transaccional", def:"Fee adicional por pagos con tarjeta de crédito/débito. Reason 6 o R9 según tipo." },
          { term:"Tarifa de Integration", def:"Cargo fijo por orden para aliados con integración de POS. Reason 17." },
          { term:"Servicio de Cargo", def:"Permite solicitar repartidor independiente. Costo = número de envíos × tarifa." },
          { term:"Descuento por Service Fee", def:"Descuento asumido por el aliado de la tarifa de servicio cobrada al usuario." },
          { term:"Descuento por Domicilio gratis", def:"Descuento en costos de envío asumido por el aliado para promociones de domicilio gratis." },
          { term:"Costo de Domicilio - Propinas (marketplace)", def:"Valor de domicilios y propinas en órdenes marketplace." },
          { term:"Meal Vouchers", def:"Tickets de alimentación empresarial. El usuario paga directo al aliado; Rappi descuenta ese valor del pago." },
          { term:"Gasto Bancario", def:"Costo de transferencia bancaria. Reason 100: calculado sobre el valor total del paidlot × % contrato." },
          { term:"Deuda Periodos Anteriores", def:"Saldo pendiente de paidlots anteriores que se recupera en el período actual." },
          { term:"Valor Ajustes Manuales", def:"Ajustes manuales (R15 por orden / R16 por tienda) cargados por el equipo de Operaciones." },
          { term:"Anticipación de cobranza", def:"Cobro por pago diario anticipado. Fee por adelanto de liquidación." },
          { term:"Cuota de préstamo", def:"Cuota de financiamiento Rappi pagada en el período (capital de trabajo, anticipo de ventas, Rappi Bank, etc.)." },
          { term:"Descuento por pago anticipado", def:"Porcentaje cobrado a comercios que aceptan el programa de pago anticipado." },
          { term:"Cashback en Rappi créditos", def:"Beneficio en créditos ofertado al usuario por cada compra, asumido por el aliado." },
          { term:"Impoconsumo / IVA de la venta (informativo)", def:"Impuesto nacional al consumo en Colombia (8%). Solo informativo, no descuenta al aliado.", solo:["Colombia"] },
        ].filter(c => !c.solo || c.solo.includes(country)).map(c => (
          <div key={c.term} style={card("#f8fafc","#e2e8f0")}>
            <div style={{ fontSize:11, fontWeight:800, color:"#1e293b" }}>{c.term}</div>
            <div style={{ fontSize:11, color:"#475569", marginTop:2 }}>{c.def}</div>
          </div>
        ))}
      </div>
    ),

    // ── FÓRMULAS ──────────────────────────────────────────────────────────
    FÓRMULAS: (
      <div>
        <div style={card("#eff6ff","#bfdbfe")}>
          <div style={{ fontWeight:800, color:"#1e40af", marginBottom:6, fontSize:13 }}>Cálculo del Pago Final al Aliado</div>
          <div style={{ fontSize:11, color:"#1e3a8a" }}>Flujo oficial del Playbook de Merchant Revenue — Transaction Reasons</div>
        </div>

        {[
          { step:"1", group:"Ventas y Descuentos", color:"#22c55e", bg:"#dcfce7", border:"#86efac",
            rows:[
              { concept:"Venta Bruta (Reason 1)", formula:"productsTotal − productsMarkup − (productsMarkdown + offersPartner + vfdPartner)", sign:"+" },
              { concept:"Descuento Producto (aliado)", formula:"Descuentos / promos asumidos por el aliado", sign:"−" },
              { concept:"Descuento en créditos", formula:"Rappi Créditos aplicados por el aliado", sign:"−" },
              { concept:"= Ventas Base", formula:"Base para cálculo de comisiones", sign:"=" },
            ]
          },
          { step:"2", group:"Descuentos sobre la Venta", color:"#8b5cf6", bg:"#ede9fe", border:"#c4b5fd",
            rows:[
              { concept:"Compensaciones (R18)", formula:"product_poor + product_difference + product_missing + product_not_delivered", sign:"−" },
              { concept:"Costo Canceladas (R10)", formula:"(VentaBruta + Antojos) × (1 − % cancelación)", sign:"−" },
              { concept:"Domicilio gratis / Propinas", formula:"Costos de envío asumidos por el aliado", sign:"−" },
              { concept:"Meal Vouchers (R45)", formula:"−mealVouchersTotal (usuario pagó directo)", sign:"−" },
            ]
          },
          { step:"3", group:"Plataforma Rappi y Tasas", color:"#ef4444", bg:"#fee2e2", border:"#fca5a5",
            rows:[
              { concept:"Comisión plataforma (R2)", formula:"Ventas Base × % comisión del contrato", sign:"−" },
              { concept:"Tarifa Transaccional (R6/R9)", formula:"(VentaBruta + Antojos) × % tarjetas", sign:"−" },
              { concept:"Tarifa Integration (R17)", formula:"Monto fijo por orden (solo aliados integrados)", sign:"−" },
              { concept:"RappiAds / Campañas", formula:"Inversión pautada semana anterior", sign:"−" },
              { concept:"DAR Plataforma", formula:"Descuento inversión Rappi sobre comisión", sign:"+" },
            ]
          },
          { step:"4", group:"Impuestos", color:"#0ea5e9", bg:"#e0f2fe", border:"#7dd3fc",
            rows:[
              { concept:"IVA Comisión (R3)", formula:"(Comisión + GastoBancario + MarketplaceFee) × % IVA", sign:"−" },
              { concept:"Retefuente (R4)", formula:"(Comisión + GastoBancario + MarketplaceFee) × % Retefuente", sign:"+", solo:["Colombia","Ecuador"] },
              { concept:"ReteICA — CO (R5)", formula:"(Comisión + GastoBancario + MarketplaceFee) × % ReteICA por ciudad", sign:"+", solo:["Colombia"] },
              { concept:"Percepción AR (R145)", formula:"(Comisión + MarketplaceFee) × alícuota IIBB por negociación", sign:"−", solo:["Argentina"] },
              { concept:"Gasto Bancario (R100)", formula:"ValorPaidlot × % gasto bancario del contrato", sign:"−" },
            ]
          },
          { step:"5", group:"Otros Descuentos y Préstamos", color:"#475569", bg:"#f1f5f9", border:"#cbd5e1",
            rows:[
              { concept:"Ajustes Manuales (R15/R16)", formula:"Valores positivos o negativos según corrección", sign:"±" },
              { concept:"Deuda Periodos Anteriores", formula:"Saldo acumulado de períodos no pagados", sign:"−" },
              { concept:"Cashback / Challenge", formula:"Beneficios asumidos por el aliado", sign:"−" },
              { concept:"Cuota de Préstamo", formula:"Valor fijo según plan de financiamiento Rappi", sign:"−" },
            ]
          },
          { step:"✅", group:"Total a Transferir", color:"#f59e0b", bg:"#fefce8", border:"#fde68a",
            rows:[
              { concept:"Valor Final Paidlot", formula:"Ventas Base − Descuentos − Fees − Impuestos − Préstamos", sign:"=" },
            ]
          },
        ].map(g => (
          <div key={g.step} style={{ marginBottom:12 }}>
            <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:5 }}>
              <div style={{ width:22, height:22, borderRadius:"50%", background:g.color, color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:10, fontWeight:900, flexShrink:0 }}>{g.step}</div>
              <div style={{ fontWeight:800, color:g.color, fontSize:12 }}>{g.group}</div>
            </div>
{g.rows.filter(r => !r.solo || r.solo.includes(country)).map(r => (
              <div key={r.concept} style={{ display:"flex", gap:8, alignItems:"flex-start", background:g.bg, border:`1px solid ${g.border}`, borderRadius:9, padding:"7px 11px", marginBottom:4 }}>
                <span style={{ fontWeight:900, color:g.color, fontSize:13, minWidth:14 }}>{r.sign}</span>
                <div>
                  <div style={{ fontSize:11, fontWeight:700, color:g.color }}>{r.concept}</div>
                  <div style={{ fontSize:10, color:g.color+"bb", marginTop:1 }}>{r.formula}</div>
                </div>
              </div>
            ))}
          </div>
        ))}
      </div>
    ),
  };

  // Embedded mode: just the tab content, no wrapper/toggle (used inside KnowledgeCenterModal)
  if (embedded) {
    return <div style={{ padding: 0 }}>{tabContent[embeddedTab] ?? tabContent["DAR"]}</div>;
  }

  return (
    <div style={{ marginTop:16, borderRadius:14, border:"1.5px solid #e2e8f0", overflow:"hidden" }}>
      <button onClick={() => setOpen(o => !o)}
        style={{ width:"100%", background:"#f8fafc", padding:"12px 16px", border:"none", cursor:"pointer", display:"flex", alignItems:"center", justifyContent:"space-between" }}>
        <div style={{ display:"flex", alignItems:"center", gap:8 }}>
          <span style={{ fontSize:16 }}>🎓</span>
          <div style={{ textAlign:"left" }}>
            <div style={{ fontWeight:800, color:"#1e293b", fontSize:13 }}>Hub de Conocimiento Rappi</div>
            <div style={{ fontSize:11, color:"#64748b" }}>DAR · ADS · Impuestos · Compensaciones · Conceptos · Fórmulas — {dar.flag} {country}</div>
          </div>
        </div>
        <span style={{ color:"#64748b", fontSize:14 }}>{open ? "▲" : "▼"}</span>
      </button>
      {open && (
        <div style={{ background:"white" }}>
          <div style={{ display:"flex", borderBottom:"1px solid #e2e8f0", padding:"0 16px", overflowX:"auto" }}>
            {EDUCATION_TABS.map(t => (
              <button key={t} onClick={() => setTab(t)}
                style={{ padding:"9px 14px", border:"none", borderBottom: tab===t ? "2.5px solid #ff441f" : "2.5px solid transparent",
                  background:"none", cursor:"pointer", fontSize:11, fontWeight: tab===t ? 800 : 500,
                  color: tab===t ? "#ff441f" : "#64748b", whiteSpace:"nowrap" }}>{t}</button>
            ))}
          </div>
          <div style={{ padding:18, maxHeight:520, overflowY:"auto" }}>
            {tabContent[tab]}
          </div>
        </div>
      )}
    </div>
  );
});
// § 8. ASISTENTE IA DE DUDAS — Semántica, sección highlight y libreto farmer
// ─────────────────────────────────────────────────────────────────────────────

// ── COUNTRY_TAX_DETAIL — descripción fiscal EXCLUSIVA por país ────────────────
// Garantiza que EducationHub y libretos NO mezclen AFIP/DIAN/SAT entre países.
const COUNTRY_TAX_DETAIL = {
  Argentina: {
    organismo: "AFIP", norma: "RG AFIP 4540/2019", iva: "IVA 21%",
    retenciones: "Percepciones IIBB (CABA, Buenos Aires, Santa Fe, Córdoba, Tucumán, Corrientes) · Retenciones Ganancias",
    nota: "Las percepciones provinciales varían según la jurisdicción donde está registrado el comercio.",
    alerta: false, urlOrganismo: "https://www.afip.gob.ar",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi","Percepcion","Percepcion Cordoba","CABA  ","SANTA FE ","Retencion Ganancias "],
  },
  Colombia: {
    organismo: "DIAN", norma: "Resolución DIAN 165/2023 · E.T. Art. 454", iva: "IVA 19%",
    retenciones: "ReteIVA · ReteICA · Retefuente · Impoconsumo (8% en restaurantes)",
    nota: "El ReteICA varía por municipio. Bogotá difiere de Medellín y Cali.",
    alerta: false, urlOrganismo: "https://www.dian.gov.co",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi","Reteiva Uso y alquiler de plataforma Rappi","ReteICA Uso y alquiler de plataforma Rappi"],
  },
  México: {
    organismo: "SAT", norma: "LISR Art. 25 Fracc. I · CFF Art. 29 · CFDI tipo E", iva: "IVA 16%",
    retenciones: "ISR (retención sobre comisión) · CFDI electrónico obligatorio",
    nota: "⚠️ CRÍTICO: La NC de DAR debe emitirse en el mismo mes fiscal. Si no, el IVA no se ajusta en ese período.",
    alerta: true, urlOrganismo: "https://www.sat.gob.mx",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi","ISR"],
  },
  Chile: {
    organismo: "SII", norma: "Ley de Facturación Electrónica · RUT 76.837.223-3", iva: "IVA 19%",
    retenciones: "Retenciones SII",
    nota: "El descuento DAR reduce la base imponible del IVA de la comisión de Rappi.",
    alerta: false, urlOrganismo: "https://www.sii.cl",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi"],
  },
  Perú: {
    organismo: "SUNAT", norma: "Factura Electrónica SUNAT · RUC 20602985971", iva: "IGV 18%",
    retenciones: "Detracción SPOT 10%",
    nota: "La detracción SPOT se calcula sobre el valor neto de la operación.",
    alerta: false, urlOrganismo: "https://www.sunat.gob.pe",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi"],
  },
  Ecuador: {
    organismo: "SRI", norma: "Ley de Régimen Tributario Interno · RUC 1793010105001", iva: "IVA 15%",
    retenciones: "IRBPNR",
    nota: "La tasa de IVA aumentó de 12% a 15% en 2024.",
    alerta: false, urlOrganismo: "https://www.sri.gob.ec",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi"],
  },
  Uruguay: {
    organismo: "DGI", norma: "e-Factura DGI · RUT 21817500000", iva: "IVA 22%",
    retenciones: "IRAE",
    nota: "El IVA uruguayo es el más alto de la región. El DAR tiene impacto relevante en la base.",
    alerta: false, urlOrganismo: "https://www.dgi.gub.uy",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi"],
  },
  "Costa Rica": {
    organismo: "Ministerio de Hacienda", norma: "Factura Electrónica DGII · Cédula Jurídica 3101768820", iva: "IVA 13%",
    retenciones: "—",
    nota: "Factura electrónica requerida por el Ministerio de Hacienda.",
    alerta: false, urlOrganismo: "https://www.hacienda.go.cr",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi"],
  },
  Nicaragua: {
    organismo: "DGI Nicaragua", norma: "Facturación local DGI", iva: "IVA / IR",
    retenciones: "IR",
    nota: "Consultar con contador local el tratamiento exacto.",
    alerta: false, urlOrganismo: "https://www.dgi.gob.ni",
    columnasClave: ["IVA Uso y alquiler de plataforma Rappi"],
  },
  "No detectado": {
    organismo: "—", norma: "—", iva: "—", retenciones: "—",
    nota: "Selecciona el país para ver el detalle fiscal.",
    alerta: false, urlOrganismo: "#", columnasClave: [],
  },
};

// ── Module-level constants — compiled once, never recreated ──────────────────
const FISCAL_TERM_RE = /ley\s+\d+|resoluci[oó]n\s+\d+|art[íi]culo\s+\d+|rg\s+\d+|decreto\s+\d+|cfdi|lisr|isr|sunat|dian|afip|sat(?!\w)/i;
const FISCAL_KW_RE   = /impuesto|retenci[oó]n|percepci[oó]n|iva|igv|iibb|ingresos brutos|reteica|retefuente|exenci[oó]n|tributar|fiscal|constancia|certificado|alicuota|al[íi]cuota/i;
const TAX_WORDS_RE   = /impuesto|retenci[oó]n|percepci[oó]n|iva|igv|isr|reteica|reteiva|iibb/i;

const DOUBT_EXAMPLES = [
  "¿Por qué me cobraron 200 pesos el 13 de abril?",
  "¿Qué es el descuento DAR en mis órdenes?",
  "¿Por qué hay un cobro de ADS si no tuve campaña?",
  "Explícame la compensación de la orden del 3 de abril",
  "¿Cuánto fue mi comisión de plataforma este período?",
];

const QUICK_QUESTIONS = [
  {
    cat: "Pagos",
    color: "#10b981", bg: "#f0fdf4", border: "#86efac",
    questions: [
      "¿Cuánto fue mi total a pagar este período?",
      "¿Por qué el pago es diferente al que calculé?",
      "¿Cuándo llega el depósito a mi cuenta?",
      "¿Qué es la deuda de períodos anteriores?",
    ],
  },
  {
    cat: "DAR",
    color: "#f97316", bg: "#fff7ed", border: "#fb923c",
    questions: [
      "¿Qué es el DAR y por qué aparece en mi paidlot?",
      "¿El DAR me reduce el pago?",
      "¿Tengo que emitir una nota crédito por el DAR?",
      "¿Cómo sé cuánto invirtió Rappi en DAR para mí?",
    ],
  },
  {
    cat: "Impuestos",
    color: "#0ea5e9", bg: "#f0f9ff", border: "#7dd3fc",
    questions: [
      "¿Por qué hay impuestos en mi liquidación?",
      "¿Qué es la percepción / retención que me cobran?",
      "¿Cómo bajo mi carga impositiva?",
      "¿Los impuestos me los descuenta Rappi o los pago yo?",
    ],
  },
  {
    cat: "Comisión",
    color: "#ef4444", bg: "#fef2f2", border: "#fca5a5",
    questions: [
      "¿Cuánto fue mi comisión de plataforma?",
      "¿Por qué mi tarifa efectiva es tan alta?",
      "¿En qué se diferencia la comisión de los impuestos?",
      "¿Qué incluye el cobro de 'Uso y alquiler de plataforma'?",
    ],
  },
  {
    cat: "Compensaciones",
    color: "#8b5cf6", bg: "#f5f3ff", border: "#c4b5fd",
    questions: [
      "¿Por qué me cobran compensaciones?",
      "¿Puedo reclamar una compensación que creo incorrecta?",
      "¿Cuál es la diferencia entre compensación y cancelación?",
      "¿Qué significa el Reason de la compensación?",
    ],
  },
  {
    cat: "RappiAds",
    color: "#7c3aed", bg: "#faf5ff", border: "#d8b4fe",
    questions: [
      "¿Por qué me cobran ADS si no tuve campaña este período?",
      "¿Cómo calculo el ROI de mi pauta en RappiAds?",
      "¿Cuánto debo invertir en ADS para que sea rentable?",
      "¿Qué es la facturación en semana vencida?",
    ],
  },
];

const CHAT_LS_KEY = (tiendaId) => `rappi_chat_v1_${tiendaId}`;

// Keyword → section mapping. Each entry: keywords that trigger the section,
// the section tab key to highlight, a human label, and the farmer script template.
const SECTION_MAP = [
  {
    keywords: ["impuesto", "iva", "igv", "retencion", "percepcion", "reteiva", "reteica", "isr", "tributario", "fiscal", "cobro fiscal", "caba", "ingresos brutos", "iibb", "sri", "dian", "afip", "sunat", "sat", "deduccion"],
    section: "impuestos",
    label: "Impuestos",
    icon: "🧾",
    color: "#0ea5e9",
    bg: "#e0f2fe",
    farmerScript: (p, country, result) => {
      const taxInfo = COUNTRY_TAX_DETAIL[country] ?? COUNTRY_TAX_DETAIL["No detectado"];
      return `Hola ${p.meta.tienda} 👋

Revisé tu relación de ventas del período ${p.resumen.inicio} al ${p.resumen.fin} (Paidlot ${p.meta.paidlotId}).

Sobre tu consulta de impuestos en ${country}: los valores de la sección de Impuestos corresponden a retenciones y percepciones que Rappi aplica según la normativa de ${taxInfo.organismo} (${taxInfo.norma}). El impuesto aplicado es ${taxInfo.iva}${taxInfo.retenciones !== "—" ? ` junto a: ${taxInfo.retenciones}` : ""}. Estos montos NO son un cobro adicional — el neto que recibes ya los considera.

${result}

${taxInfo.nota ? `Nota importante: ${taxInfo.nota}` : ""}

Para más información oficial, puedes consultar: ${taxInfo.urlOrganismo}

¡Cualquier otra duda, estamos aquí para apoyarte! 🙌`;
    },
  },
  {
    keywords: ["compensacion", "devolucion", "devolver", "reembolso", "pedido equivocado", "producto faltante", "faltante", "cancelad", "cancelacion", "nota credito", "nota de credito", "reintegro", "ajuste a favor", "domicilio", "meal voucher", "repartidor", "descuento sobre venta"],
    section: "descuentosVenta",
    label: "Descuentos sobre la venta",
    icon: "🔄",
    color: "#8b5cf6",
    bg: "#ede9fe",
    farmerScript: (p, country, result) => `Hola ${p.meta.tienda} 👋

Te escribo sobre los descuentos sobre la venta reflejados en tu liquidación del período ${p.resumen.inicio} al ${p.resumen.fin}.

${result}

Esta sección incluye compensaciones por pedidos con problemas, costos de domicilio, meal vouchers y canceladas. Si consideras que algún monto fue aplicado por error, podemos revisarlo juntos. ¡Escríbeme! 🙌`,
  },
  {
    keywords: ["dar", "descuento rappi", "descuento asumido", "inversion rappi", "rappi invierte", "cupon", "promo", "promocion", "descuento del producto", "no cobro descuento", "quien paga el descuento"],
    section: "dar",
    label: "Inversión DAR Rappi",
    icon: "🎯",
    color: "#f97316",
    bg: "#fff7ed",
    farmerScript: (p, country, result) => `Hola ${p.meta.tienda} 👋

Quiero explicarte sobre los descuentos que aparecen en tu liquidación del período ${p.resumen.inicio} al ${p.resumen.fin}.

${result}

Lo más importante: estos descuentos son financiados 100% por Rappi (los llamamos DAR — Descuento Asumido por Rappi). Tu ingreso neto NO se reduce. En el paidlot verás la columna "Descuentos por inversión de Rappi DAR" que refleja esta inversión, y una compensación equivalente en la comisión de plataforma.

¡Cualquier duda adicional, con mucho gusto! 🙌`,
  },
  {
    keywords: ["comision", "porcentaje", "tarifa", "plataforma", "uso plataforma", "alquiler", "fee", "cobro plataforma", "cuota plataforma", "rappiads", "ads", "publicidad", "pauta"],
    section: "plataforma",
    label: "Plataforma Rappi",
    icon: "🏢",
    color: "#ef4444",
    bg: "#fee2e2",
    farmerScript: (p, country, result) => `Hola ${p.meta.tienda} 👋

Sobre la consulta de comisiones y tarifas de plataforma en tu liquidación del ${p.resumen.inicio} al ${p.resumen.fin}:

${result}

La comisión de plataforma se calcula sobre tu venta bruta según el porcentaje acordado en tu contrato. Si tienes inversión DAR activa, verás también un descuento sobre esa comisión que compensa el descuento que Rappi hizo a tus clientes.

Si quieres revisar el porcentaje que aplica a tu cuenta o tienes dudas sobre algún cobro específico, podemos revisarlo juntos. 🙌`,
  },
  {
    keywords: ["orden", "pedido", "venta", "factura", "ingreso", "venta bruta", "neto", "transferencia", "pago", "fecha", "dia", "cuanto me pagan", "cuanto recibi", "recibo"],
    section: "ordenes",
    label: "Órdenes y Ventas",
    icon: "📦",
    color: "#22c55e",
    bg: "#dcfce7",
    farmerScript: (p, country, result) => `Hola ${p.meta.tienda} 👋

Revisé el detalle de órdenes de tu liquidación del período ${p.resumen.inicio} al ${p.resumen.fin} (Paidlot ${p.meta.paidlotId}).

${result}

Recuerda que el "Valor a transferir" de cada orden es el resultado de restar la comisión de Rappi a tu venta bruta. La suma de todos esos valores es lo que recibes en la transferencia del período.

Si ves alguna orden que no reconoces o un monto que no cuadra, compárteme el número de orden y lo revisamos. 🙌`,
  },
  {
    keywords: ["ajuste", "deuda anterior", "deuda periodo", "ajuste manual", "descuento prestamo", "prestamo", "cuota prestamo", "saldo pendiente"],
    section: "ajustes",
    label: "Ajustes y Deudas",
    icon: "⚖️",
    color: "#64748b",
    bg: "#f1f5f9",
    farmerScript: (p, country, result) => `Hola ${p.meta.tienda} 👋

Sobre los ajustes que aparecen en tu liquidación del período ${p.resumen.inicio} al ${p.resumen.fin}:

${result}

Los ajustes manuales son correcciones aplicadas por el equipo financiero de Rappi fuera del ciclo operativo normal. Las deudas de períodos anteriores corresponden a saldos pendientes de liquidaciones previas que se recuperan en el período actual. Ambos impactan directamente tu neto a transferir.

Si tienes dudas sobre el origen de algún ajuste específico, podemos solicitarle al equipo financiero el detalle. ¡Avísame! 🙌`,
  },
];

// Detect which section best matches the query (keyword scoring)
function detectSection(query) {
  const q = query.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  let best = null, bestScore = 0;
  for (const entry of SECTION_MAP) {
    const score = entry.keywords.reduce((s, kw) => {
      const k = kw.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      return s + (q.includes(k) ? (k.length > 5 ? 3 : 1) : 0);
    }, 0);
    if (score > bestScore) { bestScore = score; best = entry; }
  }
  return bestScore >= 1 ? best : null;
}

const DoubtAssistant = memo(({ paidlot, country, allPaidlots, onHighlightSection, onHighlightTab, inline = false, onOpenKnowledge }) => {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [result, setResult] = useState(null);
  const [script, setScript] = useState(null);
  const [detectedSection, setDetectedSection] = useState(null);
  const [loading, setLoading] = useState(false);
  const [copied, setCopied] = useState(false);
  const [localResolved, setLocalResolved] = useState(false);
  const [externalTerm, setExternalTerm] = useState(null);
  const textareaRef = useRef(null);
  const [trend, setTrend] = useState(null);
  const [historyOpen, setHistoryOpen] = useState(false);
  const [quickCat, setQuickCat] = useState(null);

  // ── Persisted chat history per store ────────────────────────────────────────
  const tiendaId = paidlot?.meta?.tiendaId ?? "unknown";
  const [chatHistory, setChatHistory] = useState(() => {
    try { const r = localStorage.getItem(CHAT_LS_KEY(tiendaId)); return r ? JSON.parse(r) : []; } catch { return []; }
  });

  useEffect(() => {
    try { const r = localStorage.getItem(CHAT_LS_KEY(tiendaId)); setChatHistory(r ? JSON.parse(r) : []); } catch { setChatHistory([]); }
  }, [tiendaId]);

  const addToHistory = useCallback((q, a) => {
    setChatHistory(prev => {
      const next = [{ q, a, ts: new Date().toLocaleString("es", { day:"2-digit", month:"short", hour:"2-digit", minute:"2-digit" }), paidlotId: paidlot?.meta?.paidlotId }, ...prev].slice(0, 30);
      try { localStorage.setItem(CHAT_LS_KEY(tiendaId), JSON.stringify(next)); } catch {}
      return next;
    });
  }, [tiendaId, paidlot]);

  const clearHistory = useCallback(() => {
    setChatHistory([]);
    try { localStorage.removeItem(CHAT_LS_KEY(tiendaId)); } catch {}
  }, [tiendaId]);

  // legacy aiLog kept for compat
  const [aiLog, setAiLog] = useState([]);

  // ── buildPreSearch: deterministic resolution BEFORE calling the API ─────────
  // Returns { resolved: true, text, scriptText } if found locally, else null.
  const buildPreSearch = useCallback((q, sec) => {
    if (!paidlot) return null;
    const p = paidlot;

    // 0. ADS semana vencida — intercept before any other check
    if (p.topKpis.rappiAdsCollection > 0 && /ads|publicidad|rappiads|semana|vencid/i.test(q)) {
      const text = "📺 Cobro de RappiAds por semana vencida detectado: " + fmt(p.topKpis.rappiAdsCollection, country) + ".\n" +
        "Este monto NO corresponde al período actual — es la inversión de publicidad pautada en la semana anterior (ej: si el paidlot es de Abril, este cobro es de la última semana de Marzo). Rappi factura ADS en semana vencida por diseño. El aliado puede verificarlo contrastando con el reporte de inversión ADS del período previo.";
      return { resolved: true, text, scriptText: sec ? sec.farmerScript(p, country, text) : null };
    }

    // 1. Weekday detection — filter orders by day of week
    const dayMap = { lunes:1, martes:2, miércoles:3, miercoles:3, jueves:4, viernes:5, sábado:6, sabado:6, domingo:0 };
    const dayMatch = Object.entries(dayMap).find(([d]) => q.includes(d));
    if (dayMatch) {
      const targetDay = dayMatch[1];
      const ordersOnDay = p.ordersTable.filter(o => {
        const d = new Date(o.fecha);
        return !isNaN(d) && d.getDay() === targetDay;
      });
      if (ordersOnDay.length > 0) {
        const text = "Órdenes del " + dayMatch[0] + " (" + ordersOnDay.length + "):\n" +
          ordersOnDay.map(o => "• " + o.fecha + " · ID " + o.ordenId + " — venta bruta " + fmt(o.ventaBruta, country) + ", neto " + fmt(o.neto, country) + ", DAR " + fmt(o.darTotal, country)).join("\n");
        return { resolved: true, text, scriptText: sec ? sec.farmerScript(p, country, text) : null };
      }
    }

    // 2. Amount matching — exact first, then ±2 range fallback
    const amountMatch = q.match(/[\d]{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?/);
    if (amountMatch) {
      const raw = amountMatch[0].replace(/\./g, "").replace(",", ".");
      const target = parseFloat(raw);
      if (target > 0) {
        const allTx = [
          ...p.ordersTable.map(o => ({ id: o.ordenId, date: o.fecha, val: Math.abs(o.neto), type: "orden", extra: "venta bruta " + fmt(o.ventaBruta, country) })),
          ...p.ordersTable.map(o => ({ id: o.ordenId, date: o.fecha, val: Math.abs(o.ventaBruta), type: "venta bruta", extra: "neto " + fmt(o.neto, country) })),
          ...p.compRows.map(c => ({ id: c.orderId, date: c.fecha, val: Math.abs(c.monto), type: "compensación", extra: c.razon })),
          ...p.ajustesRows.map(a => ({ id: a.ordenId, date: a.fecha, val: Math.abs(a.ajuste || a.deuda), type: "ajuste", extra: a.razon })),
          ...(p.topKpis.impuestosPorRegla ?? []).filter(r => r.value > 0).map(r => ({ id: "—", date: "período", val: r.value, type: "impuesto", extra: r.name })),
        ];
        // Exact match (0.5% tolerance)
        const exactTol = Math.max(0.5, target * 0.005);
        const hits = allTx.filter(t => Math.abs(t.val - target) <= exactTol);
        if (hits.length > 0) {
          const text = "Encontré el valor " + fmt(target, country) + " en:\n" +
            hits.map(h => "• " + h.type.toUpperCase() + " · ID " + h.id + " · " + h.date + " — " + h.extra).join("\n");
          return { resolved: true, text, scriptText: sec ? sec.farmerScript(p, country, text) : null };
        }
        // ±2 range fallback — suggest nearest match
        const rangeBand = 2;
        const near = allTx.filter(t => Math.abs(t.val - target) <= rangeBand).sort((a, b) => Math.abs(a.val - target) - Math.abs(b.val - target));
        if (near.length > 0) {
          const best = near[0];
          const text = "No encontré " + fmt(target, country) + " exacto, pero encontré un " + best.type + " de " + fmt(best.val, country) + " (ID " + best.id + " · " + best.date + " — " + best.extra + "). ¿Es este el valor que buscas?";
          return { resolved: true, text, scriptText: sec ? sec.farmerScript(p, country, text) : null };
        }
      }
    }

    // 3. Date matching — list all orders for a mentioned date
    const dateMatch = q.match(/(\d{1,2})[/\-\s](?:de\s)?(\w+)(?:[/\-\s](\d{2,4}))?/i);
    if (dateMatch) {
      const monthNames = { enero:1,febrero:2,marzo:3,abril:4,mayo:5,junio:6,julio:7,agosto:8,septiembre:9,octubre:10,noviembre:11,diciembre:12 };
      const day = dateMatch[1].padStart(2, "0");
      const monthRaw = dateMatch[2]?.toLowerCase();
      const monthNum = monthNames[monthRaw] ? String(monthNames[monthRaw]).padStart(2, "0") : monthRaw;
      const ordersOnDate = p.ordersTable.filter(o => {
        const d = String(o.fecha ?? "");
        return d.includes("-" + monthNum + "-" + day) || d.includes(day + "/" + monthNum);
      });
      if (ordersOnDate.length > 0) {
        const text = "Órdenes del " + day + "/" + monthNum + ":\n" +
          ordersOnDate.map(o => "• ID " + o.ordenId + " — venta bruta " + fmt(o.ventaBruta, country) + ", neto " + fmt(o.neto, country) + ", DAR " + fmt(o.darTotal, country)).join("\n");
        return { resolved: true, text, scriptText: sec ? sec.farmerScript(p, country, text) : null };
      }
    }

    return null;
  }, [paidlot, country]);

  // ── computeTrend: cross-paidlot net trend (called once on open) ─────────────
  const computeTrend = useCallback(() => {
    if (!allPaidlots || allPaidlots.length < 2 || !paidlot) return null;
    const sorted = [...allPaidlots].sort((a, b) => (a.loadedAt ?? 0) - (b.loadedAt ?? 0));
    const idx = sorted.findIndex(pl => pl.id === paidlot.id);
    if (idx < 1) return null;
    const prev = sorted[idx - 1];
    const diff = round2(paidlot.topKpis.neto - prev.topKpis.neto);
    const pct = prev.topKpis.neto > 0 ? round2((diff / prev.topKpis.neto) * 100) : 0;
    return { diff, pct, prevPeriod: `${prev.resumen.inicio}→${prev.resumen.fin}`, prevNeto: prev.topKpis.neto };
  }, [paidlot, allPaidlots]);

  const buildContext = useCallback(() => {
    if (!paidlot) return "";
    const p = paidlot;
    const kpi = p.topKpis;
    const cty = country;
    const darConfig = DAR_CONFIG[country] ?? DAR_CONFIG["No detectado"];

    const ordersText = p.ordersTable.slice(0, 15).map(o =>
      `  ${o.ordenId}|${o.fecha}|vb=${fmt(o.ventaBruta, cty)}|neto=${fmt(o.neto, cty)}|DAR=${fmt(o.darTotal, cty)}`
    ).join("\n") + (p.ordersTable.length > 15 ? `\n  (+${p.ordersTable.length - 15} más)` : "");
    const compsText = p.compRows.slice(0, 10).map(c =>
      `  ${c.orderId}|${c.fecha}|${c.razon}|${fmt(c.monto, cty)}`
    ).join("\n");
    const ajustesText = p.ajustesRows.slice(0, 10).map(a =>
      `  ${a.ordenId}|${a.fecha}|${a.razon}|ajuste=${fmt(a.ajuste, cty)}`
    ).join("\n");
    const extrasText = p.extrasTable.slice(0, 5).map(e =>
      `  ${e.ordenId}|${e.fecha}|${e.tipo}|${fmt(e.ventaBruta, cty)}`
    ).join("\n");

    return `Eres un asesor comercial y contable de Rappi, especializado en ayudar a los aliados (restaurantes y tiendas) a entender y optimizar sus liquidaciones de pago (paidlots). Tu rol es:
- NUNCA sugerir vacíos legales, evasión fiscal, ni formas de reducir pagos al margen de la ley.
- SIEMPRE orientar al aliado hacia soluciones legales: certificados de exención vigentes, correcta categorización fiscal, deducciones legalmente aplicables, revisión de contratos y tarifas con su ejecutivo de cuenta.
- SIEMPRE destacar el valor de RappiAds como herramienta que el aliado puede activar para mejorar su visibilidad y ventas.
- NO decirle al aliado que puede "activar DAR" o "invertir en DAR" — el DAR es una inversión que Rappi decide según sus criterios internos de coinversión, no algo que el aliado activa por voluntad propia. Si el aliado pregunta por DAR, explicar que es un programa de inversión de Rappi y que deben conversarlo con su ejecutivo de cuenta.
- Responder en español, con tono positivo, empático y orientado a soluciones. Máximo 3 párrafos, sin tecnicismos. Usa los datos del paidlot para personalizar la respuesta.

PAIDLOT: ${p.meta.tienda} | ${country} | ${p.resumen.inicio}→${p.resumen.fin}
KPIs: VentaBruta=${fmt(kpi.ventaBruta,cty)} | TotalPagar=${fmt(kpi.totalAPagar??kpi.neto,cty)} | Comision=${fmt(kpi.comision,cty)} | Impuestos=${fmt(kpi.impuestosTotalExacto??kpi.totalImpuestos??0,cty)} | TarifaEfectiva=${fmtPct(kpi.effectiveFee)} | DAR=${kpi.hasDar?"SI "+fmt(kpi.darInversionTotal,cty):"NO"} | ADS=${fmt(kpi.cuotaRappiAds??0,cty)} | ADSvencida=${fmt(kpi.rappiAdsCollection??0,cty)} | Compensaciones=${fmt(kpi.compensaciones,cty)} | Ajustes=${fmt(kpi.ajustesTotal,cty)} | Ordenes=${kpi.ordenes}
DAR: organismo=${darConfig.organismo} | plazoNC=${darConfig.plazoNC}
Impuestos detectados: ${(kpi.impuestosPorRegla??[]).filter(r=>r.value>0).map(r=>`${r.name}=${fmt(r.value,cty)}`).join(", ")||"ninguno"}

ÓRDENES (mostrando ${Math.min(p.ordersTable.length,15)} de ${p.ordersTable.length}):
${ordersText||"Sin órdenes."}

COMPENSACIONES (${p.compRows.length}):
${compsText||"Sin compensaciones."}

AJUSTES (${p.ajustesRows.length}):
${ajustesText||"Sin ajustes."}`;
  }, [paidlot, country]);

  const handleSearch = useCallback(async (queryOverride) => {
    const q = (queryOverride !== undefined ? queryOverride : query).trim();
    if (!q || !paidlot) return;
    setLoading(true);
    setResult(null);
    setScript(null);
    setDetectedSection(null);
    setLocalResolved(false);

    // 1. Semantic section detection (always first)
    const sec = detectSection(q);
    if (sec && onHighlightSection) onHighlightSection(sec.section);
    if (sec && onHighlightTab) onHighlightTab(sec.section);

    // 2. Deterministic pre-search — amount/date matching (no API call needed)
    const local = buildPreSearch(q.toLowerCase(), sec);
    if (local) {
      setLocalResolved(true);
      setResult(local.text);
      if (local.scriptText) setScript(local.scriptText);
      setLoading(false);
      return;
    }
    setLocalResolved(false);

    // Always expose Google search for fiscal/tax queries — broad detection
    const isFiscalQuery = FISCAL_TERM_RE.test(q) || FISCAL_KW_RE.test(q) || sec?.section === "impuestos";
    if (isFiscalQuery) setExternalTerm(q); else setExternalTerm(null);

    // ── REASONS KNOWLEDGE BASE (Playbook of Merchant Revenue) ────────────────
    const REASONS = {
      "145": { country: "Argentina", name: "Percepción IIBB CABA", formula: "(Comisión + Fee Marketplace) × alícuota IIBB", norma: "AGIP / jurisdicción provincial", hint: "Verificar si el aliado tiene certificado de exención IIBB vigente. Si tributa en exceso, puede reclamar como saldo a favor ante AFIP." },
      "165": { country: "Colombia", name: "ReteICA Bogotá", formula: "Venta Bruta × 0.414% (Bogotá Restaurantes)", norma: "Res. DIAN 165/2023 · Estatuto Tributario Art. 368", hint: "Si el aliado opera en otro municipio, la alícuota puede diferir. Verificar registro de industria y comercio." },
      "012": { country: "México", name: "Retención ISR", formula: "Comisión × 1% (personas morales) ó 0.1% (personas físicas)", norma: "LISR Art. 106 / Art. 116", hint: "Revisar si el aliado está dado de alta como persona física o moral ante el SAT. La alícuota varía." },
      "023": { country: "Perú", name: "Detracción SPOT", formula: "Venta Neta × 10%", norma: "SUNAT RS 183-2004 SPOT", hint: "Verificar que el aliado tenga cuenta corriente de detracciones activa en el Banco de la Nación." },
    };

    const qLower = q.toLowerCase();
    const reasonMatch = qLower.match(/reason\s*(\d+)|motivo\s*(\d+)|raz[oó]n\s*(\d+)/i);
    const reasonId = reasonMatch ? (reasonMatch[1] ?? reasonMatch[2] ?? reasonMatch[3]) : null;
    const taxKeywords = /percepci[oó]n|retenci[oó]n|reteica|iibb|ingresos brutos|isr|spot|detrac|alicuota|al[íi]cuota/i;

    if (reasonId && REASONS[reasonId]) {
      const r = REASONS[reasonId];
      const text = "📋 Reason " + reasonId + " — " + r.name + " (" + r.country + ")\n\n" +
        "Fórmula: " + r.formula + "\n" +
        "Norma: " + r.norma + "\n\n" +
        "💡 " + r.hint;
      const playbook_script = "Hola " + paidlot.meta.tienda + " 👋\n\nEl cobro corresponde al Reason " + reasonId + " (" + r.name + "). " +
        "Se calcula como: " + r.formula + ".\n\n" + r.hint + "\n\n¡Quedo atento! 🙌";
      setResult(text);
      setScript(playbook_script);
      setExternalTerm("Reason " + reasonId + " " + r.name + " " + country);
      setAiLog(prev => [...prev, { q, a: text, ts: new Date().toLocaleTimeString() }]); addToHistory(q, text);
      logQueryToSheets({ aliado: paidlot.meta.tienda, pais: country, pregunta: q, respuesta: text });
      setLoading(false);
      return;
    }

    // ── TAX DIAGNOSTIC — alícuota excesiva ───────────────────────────────────
    if (taxKeywords.test(qLower) && paidlot.topKpis.totalImpuestos > 0) {
      const taxRate = paidlot.topKpis.ventaBruta > 0
        ? paidlot.topKpis.totalImpuestos / paidlot.topKpis.ventaBruta
        : 0;
      if (taxRate > 0.05) {
        const diagText = "⚠️ Riesgo Fiscal detectado: Los impuestos del período representan el " +
          (taxRate * 100).toFixed(2) + "% de la venta bruta (" +
          fmt(paidlot.topKpis.totalImpuestos, country) + " sobre " +
          fmt(paidlot.topKpis.ventaBruta, country) + ").\n\n" +
          "Esto supera el umbral del 5%, lo que puede indicar que el aliado está mal categorizado ante el fisco o que sus certificados de exención están vencidos.\n\n" +
          "✅ Acción recomendada: Solicitar al aliado que actualice sus certificados de exención o constancias tributarias ante " +
          (DAR_CONFIG[country]?.organismo ?? "el organismo fiscal") +
          ". El monto en exceso puede recuperarse como saldo a favor.";
        const cert_script = "Hola " + paidlot.meta.tienda + " 👋\n\nRevisando tu factura detectamos un posible Riesgo Fiscal: los impuestos representan el " +
          (taxRate * 100).toFixed(2) + "% de tus ventas, por encima del rango normal.\n\n" +
          "Te recomendamos revisar con tu contador si tus certificados de exención están vigentes — en caso de sobrepago, puedes solicitar el saldo a favor ante " +
          (DAR_CONFIG[country]?.organismo ?? "el organismo fiscal de tu país") + ".\n\n¡Estamos para ayudarte! 🙌";
        setResult(diagText);
        setScript(cert_script);
        setExternalTerm("certificado exención impuestos Rappi aliados " + country);
        setAiLog(prev => [...prev, { q, a: diagText, ts: new Date().toLocaleTimeString() }]); addToHistory(q, diagText);
        logQueryToSheets({ aliado: paidlot.meta.tienda, pais: country, pregunta: q, respuesta: diagText });
        setLoading(false);
        return;
      }
    }

    // 3. AI fallback — Groq (Llama 3.3 70B)
    const ctrl = new AbortController();
    try {
      const context = buildContext();
      const groqKey = import.meta.env.VITE_GROQ_API_KEY ?? "";
      const response = await fetch(
        `https://api.groq.com/openai/v1/chat/completions`,
        {
          method: "POST",
          signal: ctrl.signal,
          headers: {
            "Content-Type": "application/json",
            "Authorization": `Bearer ${groqKey}`,
          },
          body: JSON.stringify({
            model: "llama-3.3-70b-versatile",
            messages: [
              { role: "system", content: context },
              { role: "user", content: q },
            ],
            max_tokens: 1000,
            temperature: 0.3,
          }),
        }
      );
      const data = await response.json();
      // Groq API error (bad key, quota, etc.)
      if (data.error) {
        setResult(`⚠️ Error Groq API: ${data.error.message ?? JSON.stringify(data.error)}\n\nVerifica que la variable VITE_GROQ_API_KEY esté configurada en Render.`);
        setLoading(false);
        return;
      }
      let text = data.choices?.[0]?.message?.content ?? "";

      // Intelligent fallback — never leave farmer with a blank/error
      if (!text || text.length < 20) {
        const topGroup = [...(paidlot.groups ?? [])].sort((a,b) => b.total - a.total)[0];
        if (TAX_WORDS_RE.test(q)) {
          text = "Basado en el Playbook de Merchant Revenue, este cobro suele ser una retención o percepción fiscal aplicada sobre la comisión de plataforma. " +
            "Si el monto es elevado, es probable que la categoría fiscal del aliado deba ser actualizada para reducir la alícuota.\n\n" +
            "✅ Solución recomendada: Revisar el 'Certificado de Exención' en el Portal de Aliados (sección Documentación Fiscal). " +
            "Si el certificado está vencido, el aliado está pagando una tasa general en lugar de su tasa preferencial.";
        } else {
          text = topGroup
            ? "El mayor movimiento del período está en " + topGroup.label + " (" + fmt(topGroup.total, country) + "). " +
              "Para afinar la búsqueda, pregunta por un monto específico (ej: '$ 1.200'), una fecha (ej: '13 de abril') o un concepto como DAR, compensaciones o comisión."
            : "Para encontrar el dato que buscas, formula la pregunta con un monto específico, una fecha o un concepto: DAR, compensaciones, impuestos o comisión de plataforma.";
        }
        setExternalTerm(q);
      }

      setResult(text);
      if (sec) setScript(sec.farmerScript(paidlot, country, text));

      setAiLog(prev => [...prev, { q, a: text, ts: new Date().toLocaleTimeString() }]); addToHistory(q, text);
      logQueryToSheets({ aliado: paidlot.meta.tienda, pais: country, pregunta: q, respuesta: text });

    } catch (err) {
      console.error("[Gemini]", err);
      const topGroup = [...(paidlot.groups ?? [])].sort((a,b) => b.total - a.total)[0];
      const fallback = topGroup
        ? "El análisis automático muestra que el mayor movimiento del período está en " +
          topGroup.label + " (" + fmt(topGroup.total, country) + "). ¿Es a ese concepto al que te refieres? Puedes preguntarme por un monto específico, una fecha o un concepto como DAR, compensaciones o impuestos."
        : "No encontré datos específicos para esa consulta. Prueba preguntando por un monto (ej: '200 pesos'), una fecha (ej: '13 de abril') o un concepto: DAR, compensaciones, impuestos o comisión.";
      setResult(fallback);
      setExternalTerm(q);
    } finally {
      setLoading(false);
    }
  }, [query, buildContext, buildPreSearch, paidlot, country, onHighlightSection, onHighlightTab, addToHistory]);

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) handleSearch();
  };

  const handleCopyScript = () => {
    if (!script) return;
    navigator.clipboard.writeText(script).then(() => {
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    });
  };

  if (!paidlot) return null;

  // ── Shared body: insights + result + script ────────────────────────────────
  const assistantBody = (
    <div style={{ flex: 1, overflowY: "auto", padding: inline ? "0" : "16px 20px" }}>

              {/* Detected section chip */}
              {detectedSection && (
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14, padding: "8px 14px", background: detectedSection.bg, border: `1.5px solid ${detectedSection.color}`, borderRadius: 30 }}>
                  <span style={{ fontSize: 16 }}>{detectedSection.icon}</span>
                  <span style={{ fontSize: 12, fontWeight: 700, color: detectedSection.color }}>
                    Esta consulta corresponde a la sección: <strong>{detectedSection.label}</strong>
                  </span>
                  <span style={{ fontSize: 10, color: detectedSection.color, marginLeft: "auto" }}>↑ resaltado en el dashboard</span>
                </div>
              )}

              {/* Trend banner — shown once when there are multiple paidlots */}
              {!result && !loading && (() => {
                const t = computeTrend();
                if (!t) return null;
                const isUp = t.diff >= 0;
                return (
                  <div style={{ marginBottom: 12, padding: "10px 14px", background: isUp ? "#f0fdf4" : "#fef2f2", borderRadius: 10, border: `1px solid ${isUp ? "#86efac" : "#fca5a5"}` }}>
                    <div style={{ fontSize: 11, fontWeight: 800, color: isUp ? "#15803d" : "#dc2626", marginBottom: 3 }}>
                      {isUp ? "📈" : "📉"} Tendencia vs período anterior ({t.prevPeriod})
                    </div>
                    <div style={{ fontSize: 12, color: isUp ? "#166534" : "#7f1d1d" }}>
                      Neto anterior: <strong>{fmt(t.prevNeto, country)}</strong> →
                      Diferencia: <strong>{isUp ? "+" : ""}{fmt(t.diff, country)}</strong> ({isUp ? "+" : ""}{t.pct}%)
                    </div>
                  </div>
                );
              })()}

              {/* ── PROACTIVE INSIGHTS — shown before first query ─────────── */}
              {!result && !loading && (() => {
                const kpi = paidlot.topKpis;
                const insights = [];

                // 1. ADS semana vencida
                if (kpi.rappiAdsCollection > 0) {
                  insights.push({
                    key: "ads",
                    icon: "📺",
                    bg: "#faf5ff", border: "#d8b4fe", text: "#7c3aed",
                    title: `Cobro ADS semana vencida: ${fmt(kpi.rappiAdsCollection, country)}`,
                    body: "Este monto corresponde a publicidad pautada la semana anterior al período. ¿Quiero redactar la explicación para el aliado?",
                    action: "Redactar explicación ADS para el aliado",
                    query: `Explica al aliado el cobro de ${fmt(kpi.rappiAdsCollection, country)} de RappiAds por semana vencida`,
                  });
                }

                // 2. Tarifa efectiva alta → sugerir notas crédito DAR
                if (kpi.effectiveFee > 0.35) {
                  insights.push({
                    key: "fee",
                    icon: "📐",
                    bg: "#fff7ed", border: "#fed7aa", text: "#c2410c",
                    title: `Tarifa efectiva alta: ${fmtPct(kpi.effectiveFee)} (umbral: 35%)`,
                    body: "La carga real sobre ventas supera el umbral. Si hay DAR activo, verificar que el aliado esté emitiendo las Notas de Crédito correspondientes para reducir la base imponible.",
                    action: "¿Cómo emitir la Nota Crédito DAR?",
                    query: `La tarifa efectiva es ${fmtPct(kpi.effectiveFee)}, ¿cómo aplica la nota crédito DAR en ${country}?`,
                  });
                }

                // 3. DAR activo → recordatorio de NC
                if (kpi.hasDar && kpi.darInversionTotal > 0) {
                  insights.push({
                    key: "dar",
                    icon: "🎯",
                    bg: "#fff7ed", border: "#fb923c", text: "#9a3412",
                    title: `DAR activo: ${fmt(kpi.darInversionTotal, country)} de inversión Rappi`,
                    body: "Rappi está invirtiendo en esta marca. El neto del aliado no cambia, pero puede emitir Nota de Crédito para reflejar el descuento fiscalmente.",
                    action: "Ver guía de Nota Crédito DAR",
                    query: `Explica el DAR de ${fmt(kpi.darInversionTotal, country)} y si el aliado debe emitir nota crédito en ${country}`,
                  });
                }

                // 4. Riesgo Fiscal — impuestos > 5% de venta bruta
                const taxRate = kpi.ventaBruta > 0 ? kpi.totalImpuestos / kpi.ventaBruta : 0;
                if (taxRate > 0.05) {
                  insights.push({
                    key: "riesgo-fiscal",
                    icon: "🧾",
                    bg: "#fef2f2", border: "#fca5a5", text: "#dc2626",
                    title: `Riesgo Fiscal: impuestos representan el ${(taxRate * 100).toFixed(1)}% de las ventas`,
                    body: "Los impuestos superan el umbral del 5%. El aliado podría estar mal categorizado o con certificados de exención vencidos. El sobrepago es recuperable.",
                    action: "Diagnosticar riesgo fiscal",
                    query: `Los impuestos representan el ${(taxRate * 100).toFixed(1)}% de las ventas. ¿Hay riesgo de sobrepago fiscal en ${country}?`,
                  });
                }

                // 5. Alerta de Ads alta
                const adsPct = kpi.ventaBruta > 0 ? kpi.cuotaRappiAds / kpi.ventaBruta : 0;
                if (adsPct > 0.20) {
                  insights.push({
                    key: "ads-alta",
                    icon: "📺",
                    bg: "#fff7ed", border: "#f97316", text: "#c2410c",
                    title: `Ads representa el ${(adsPct * 100).toFixed(1)}% de las ventas facturadas`,
                    body: `La inversión en RappiAds (${fmt(kpi.cuotaRappiAds, country)}) es alta respecto a las ventas facturadas. Evalúa con el aliado el ROI de la pauta.`,
                    action: "Revisar inversión ADS",
                    query: `La cuota de RappiAds es ${fmt(kpi.cuotaRappiAds, country)}, que representa el ${(adsPct * 100).toFixed(1)}% de las ventas facturadas. ¿Cómo analizo el ROI?`,
                  });
                }

                if (!insights.length) return null;
                return (
                  <div style={{ marginBottom: 14 }}>
                    <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>⚡ Insights detectados automáticamente</div>
                    {insights.map(ins => (
                      <div key={ins.key} style={{ marginBottom: 8, padding: "10px 14px", background: ins.bg, border: `1.5px solid ${ins.border}`, borderRadius: 12 }}>
                        <div style={{ fontWeight: 800, color: ins.text, fontSize: 12, marginBottom: 3 }}>{ins.icon} {ins.title}</div>
                        <div style={{ fontSize: 11, color: ins.text, lineHeight: 1.5, marginBottom: 8 }}>{ins.body}</div>
                        <button
                          onClick={() => { setQuery(ins.query); setTimeout(handleSearch, 50); }}
                          style={{ fontSize: 11, fontWeight: 700, padding: "5px 12px", background: "white", border: `1px solid ${ins.border}`, borderRadius: 20, cursor: "pointer", color: ins.text }}
                        >{ins.action} →</button>
                      </div>
                    ))}
                  </div>
                );
              })()}

              {/* ── Historial de conversación por marca ─────────────────── */}
              {chatHistory.length > 0 && (
                <div style={{ marginBottom: 14, borderRadius: 12, border: "1px solid #e2e8f0", overflow: "hidden" }}>
                  <div onClick={() => setHistoryOpen(v => !v)} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "9px 14px", background: "#f8fafc", cursor: "pointer" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
                      <span style={{ fontSize: 14 }}>🕐</span>
                      <span style={{ fontWeight: 700, fontSize: 12, color: "#475569" }}>Historial — {paidlot.meta.tienda}</span>
                      <span style={{ fontSize: 10, background: "#e0f2fe", color: "#0369a1", fontWeight: 700, padding: "1px 7px", borderRadius: 20 }}>{chatHistory.length} consulta{chatHistory.length > 1 ? "s" : ""}</span>
                    </div>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <button onClick={e => { e.stopPropagation(); clearHistory(); }} style={{ fontSize: 10, color: "#ef4444", background: "none", border: "none", cursor: "pointer", fontWeight: 700 }}>Limpiar</button>
                      <span style={{ fontSize: 11, color: "#94a3b8" }}>{historyOpen ? "▲" : "▼"}</span>
                    </div>
                  </div>
                  {historyOpen && (
                    <div style={{ maxHeight: 280, overflowY: "auto" }}>
                      {chatHistory.map((h, i) => (
                        <div key={i} style={{ padding: "10px 14px", borderTop: "1px solid #f1f5f9", background: i % 2 === 0 ? "white" : "#fafafa" }}>
                          <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8, marginBottom: 4 }}>
                            <div style={{ fontSize: 11, fontWeight: 700, color: "#ff441f", flex: 1 }}>❓ {h.q}</div>
                            <span style={{ fontSize: 10, color: "#94a3b8", flexShrink: 0 }}>{h.ts}</span>
                          </div>
                          <div style={{ fontSize: 11, color: "#475569", lineHeight: 1.6 }}>{h.a.slice(0, 180)}{h.a.length > 180 ? "…" : ""}</div>
                          <button onClick={() => { setQuery(h.q); setTimeout(() => handleSearch(h.q), 50); }} style={{ marginTop: 5, fontSize: 10, color: "#1d4ed8", background: "none", border: "none", cursor: "pointer", fontWeight: 600, padding: 0 }}>↩ Repetir consulta</button>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}

              {/* ── Preguntas rápidas por categoría ─────────────────────────── */}
              {!result && !loading && (
                <div style={{ marginBottom: 14 }}>
                  <div style={{ fontSize: 11, fontWeight: 700, color: "#94a3b8", textTransform: "uppercase", letterSpacing: "0.04em", marginBottom: 8 }}>Preguntas rápidas</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 8 }}>
                    {QUICK_QUESTIONS.map(cat => (
                      <button key={cat.cat} onClick={() => setQuickCat(quickCat === cat.cat ? null : cat.cat)}
                        style={{ fontSize: 11, fontWeight: 700, padding: "5px 12px", borderRadius: 20, border: `1.5px solid ${quickCat === cat.cat ? cat.color : cat.border}`, background: quickCat === cat.cat ? cat.bg : "white", color: quickCat === cat.cat ? cat.color : "#64748b", cursor: "pointer", transition: "all 0.15s" }}>
                        {cat.cat}
                      </button>
                    ))}
                  </div>
                  {quickCat && (() => {
                    const catData = QUICK_QUESTIONS.find(c => c.cat === quickCat);
                    if (!catData) return null;
                    return (
                      <div style={{ background: catData.bg, border: `1.5px solid ${catData.border}`, borderRadius: 12, padding: "10px 12px", display: "flex", flexDirection: "column", gap: 5 }}>
                        {catData.questions.map(q => (
                          <button key={q} onClick={() => { setQuery(q); setQuickCat(null); setTimeout(() => handleSearch(q), 50); }}
                            style={{ textAlign: "left", fontSize: 12, padding: "7px 12px", background: "white", border: `1px solid ${catData.border}`, borderRadius: 8, cursor: "pointer", color: "#1e293b", fontWeight: 500, lineHeight: 1.4 }}>
                            → {q}
                          </button>
                        ))}
                      </div>
                    );
                  })()}
                </div>
              )}

              {/* Loading */}
              {loading && (
                <div style={{ display: "flex", alignItems: "center", gap: 10, padding: "16px", background: "#f8fafc", borderRadius: 12, marginBottom: 14 }}>
                  <div style={{ width: 18, height: 18, border: "2px solid #ff441f", borderTopColor: "transparent", borderRadius: "50%", animation: "spin 0.8s linear infinite", flexShrink: 0 }} />
                  <span style={{ fontSize: 13, color: "#64748b" }}>Analizando el paidlot y construyendo la respuesta...</span>
                </div>
              )}

              {/* ── RESULT BLOCK ── */}
              {result && !loading && (
                <div style={{ animation: "fadeIn 0.3s ease" }}>

                  {/* AI Answer */}
                  <div style={{ marginBottom: 16 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
                      <span style={{ fontSize: 14 }}>🔍</span>
                      <span style={{ fontSize: 11, fontWeight: 700, color: "#475569", textTransform: "uppercase" }}>Análisis del paidlot</span>
                      {localResolved && <span style={{ fontSize: 10, fontWeight: 700, background: "#dcfce7", color: "#15803d", padding: "2px 8px", borderRadius: 20 }}>⚡ Local — sin IA</span>}
                    </div>
                    <div style={{ background: "#f8fafc", border: "1px solid #e2e8f0", borderRadius: 12, padding: "14px 16px" }}>
                      {result.split("\n").map((line, i) => (
                        <p key={i} style={{ fontSize: 13, color: "#1e293b", lineHeight: 1.65, margin: "0 0 6px" }}>{line || "\u00A0"}</p>
                      ))}
                    </div>
                  </div>

                  {/* ── FARMER SCRIPT ── */}
                  {script && (
                    <div style={{ marginBottom: 14 }}>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                          <span style={{ fontSize: 14 }}>📋</span>
                          <span style={{ fontSize: 11, fontWeight: 700, color: "#475569", textTransform: "uppercase" }}>Libreto — Mensaje para enviar al aliado</span>
                        </div>
                        <button
                          onClick={handleCopyScript}
                          style={{ fontSize: 11, fontWeight: 700, padding: "5px 12px", borderRadius: 20, border: "none", cursor: "pointer", background: copied ? "#dcfce7" : "#f1f5f9", color: copied ? "#166534" : "#475569", transition: "all 0.2s" }}
                        >
                          {copied ? "✅ Copiado" : "📋 Copiar"}
                        </button>
                      </div>
                      <div style={{ background: "#fffbeb", border: "1.5px solid #fde68a", borderRadius: 12, padding: "14px 16px", position: "relative" }}>
                        <div style={{ position: "absolute", top: 10, right: 12, fontSize: 10, color: "#92400e", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" }}>WhatsApp / Chat</div>
                        {script.split("\n").map((line, i) => (
                          <p key={i} style={{ fontSize: 12, color: "#1e293b", lineHeight: 1.7, margin: "0 0 4px", fontFamily: line.startsWith("Hola") ? "inherit" : "inherit" }}>{line || "\u00A0"}</p>
                        ))}
                      </div>
                      <p style={{ fontSize: 10, color: "#94a3b8", marginTop: 6 }}>💡 Edita el mensaje antes de enviarlo si necesitas personalizarlo para este aliado.</p>
                    </div>
                  )}

                  <div style={{ display: "flex", gap: 8, marginTop: 4, flexWrap: "wrap" }}>
                    <button onClick={() => { setResult(null); setScript(null); setDetectedSection(null); setExternalTerm(null); }} style={{ fontSize: 11, color: "#64748b", background: "none", border: "none", cursor: "pointer", textDecoration: "underline" }}>Hacer otra consulta</button>
                    {externalTerm && (
                      <a
                        href={"https://www.google.com/search?q=" + encodeURIComponent("normativa fiscal Rappi aliados " + country + " " + externalTerm)}
                        target="_blank" rel="noreferrer"
                        style={{ fontSize: 11, fontWeight: 700, color: "#1d4ed8", background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 20, padding: "3px 10px", textDecoration: "none" }}
                      >🔍 Investigar base legal</a>
                    )}
                    {aiLog.length > 0 && (
                      <button
                        onClick={() => exportPDF(paidlot, country, aiLog)}
                        style={{ fontSize: 11, fontWeight: 700, color: "#7c3aed", background: "#faf5ff", border: "1px solid #d8b4fe", borderRadius: 20, padding: "3px 10px", cursor: "pointer" }}
                      >📄 Exportar log de consultas</button>
                    )}
                  </div>
                </div>
              )}
    </div>
  ); // end assistantBody

  const assistantInput = (
    <div style={{ padding: inline ? "12px 0 0" : "14px 20px", borderTop: "1px solid #f1f5f9", flexShrink: 0 }}>
              <div style={{ position: "relative" }}>
                <textarea
                  ref={textareaRef}
                  value={query}
                  onChange={e => setQuery(e.target.value)}
                  onKeyDown={handleKeyDown}
                  placeholder='Ej: "El aliado pregunta por qué le descontaron $200 el 13 de abril en impuestos"...'
                  rows={3}
                  style={{ width: "100%", padding: "12px 50px 12px 14px", borderRadius: 12, border: "1.5px solid #e2e8f0", fontSize: 13, lineHeight: 1.5, resize: "none", outline: "none", fontFamily: "inherit", color: "#0f172a", background: "#ffffff", boxSizing: "border-box", transition: "border-color 0.15s" }}
                  onFocus={e => e.target.style.borderColor = "#ff441f"}
                  onBlur={e => e.target.style.borderColor = "#e2e8f0"}
                />
                <button
                  onClick={handleSearch}
                  disabled={loading || !query.trim()}
                  style={{ position: "absolute", right: 10, bottom: 10, width: 32, height: 32, borderRadius: 8, background: loading || !query.trim() ? "#f1f5f9" : "linear-gradient(135deg,#ff441f,#ff6b47)", color: loading || !query.trim() ? "#94a3b8" : "white", border: "none", cursor: loading || !query.trim() ? "not-allowed" : "pointer", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 16, transition: "all 0.15s" }}
                  title="Enviar (Ctrl+Enter)"
                >
                  {loading ? "⏳" : "→"}
                </button>
              </div>
              <div style={{ marginTop: 6, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <div style={{ fontSize: 10, color: "#94a3b8" }}>Ctrl+Enter · Detecta la sección automáticamente</div>
                <button
                  onClick={() => exportPDF(paidlot, country)}
                  style={{ fontSize: 10, fontWeight: 700, color: "#1d4ed8", background: "#eff6ff", border: "1px solid #bfdbfe", borderRadius: 20, padding: "4px 10px", cursor: "pointer" }}
                  title="Exportar PDF con desglose completo del paidlot"
                >📄 Informe PDF</button>
              </div>
            </div>
  ); // end assistantInput

  // ── Inline mode: just body + input, no floating button or overlay ──────────
  if (inline) {
    return (
      <div style={{ display: "flex", flexDirection: "column", height: "100%", minHeight: 400 }}>
        {assistantBody}
        {assistantInput}
      </div>
    );
  }

  // ── Standalone mode: floating button + modal overlay ──────────────────────
  return (
    <>
      {/* Floating button */}
      <button
        onClick={() => onOpenKnowledge ? onOpenKnowledge("CONSULTAR") : setOpen(true)}
        style={{ position: "fixed", bottom: 28, right: 28, width: 56, height: 56, borderRadius: "50%", background: "linear-gradient(135deg,#ff441f,#ff6b47)", color: "white", border: "none", cursor: "pointer", fontSize: 24, boxShadow: "0 8px 24px rgba(255,68,31,0.4)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", transition: "transform 0.15s" }}
        onMouseEnter={e => e.currentTarget.style.transform = "scale(1.1)"}
        onMouseLeave={e => e.currentTarget.style.transform = "scale(1)"}
        title="Centro de Conocimiento Rappi"
      >💬</button>

      {/* Standalone modal */}
      {open && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "flex-end", justifyContent: "flex-end", zIndex: 300, padding: 24 }}>
          <div style={{ background: "white", borderRadius: 20, width: "100%", maxWidth: 580, maxHeight: "88vh", display: "flex", flexDirection: "column", boxShadow: "0 24px 64px rgba(0,0,0,0.25)", animation: "slideUp 0.2s ease" }}>
            <style>{`@keyframes slideUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}} @keyframes fadeIn{from{opacity:0}to{opacity:1}} @keyframes spin{to{transform:rotate(360deg)}}`}</style>
            <div style={{ padding: "16px 20px", borderBottom: "1px solid #f1f5f9", display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                <div style={{ width: 36, height: 36, borderRadius: 10, background: "linear-gradient(135deg,#ff441f,#ff6b47)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 18 }}>💬</div>
                <div>
                  <div style={{ fontWeight: 800, fontSize: 14, color: "#0f172a" }}>Asistente de Dudas del Farmer</div>
                  <div style={{ fontSize: 11, color: "#64748b" }}>{paidlot.meta.tienda} · {country} · Paidlot {paidlot.meta.paidlotId}</div>
                </div>
              </div>
              <button onClick={() => { setOpen(false); setResult(null); setScript(null); setQuery(""); setDetectedSection(null); }} style={{ background: "none", border: "none", cursor: "pointer", fontSize: 20, color: "#94a3b8" }}>×</button>
            </div>
            {assistantBody}
            {assistantInput}
          </div>
        </div>
      )}
    </>
  );
});

// ─────────────────────────────────────────────────────────────────────────────
// § 9. KNOWLEDGE CENTER MODAL — Hub + Consultar + Objeciones
// ─────────────────────────────────────────────────────────────────────────────

const OBJECTIONS = [
  { id: "dar-reduce-pago", icon: "🎯", title: '"El DAR me está bajando el pago / me perjudica"', tags: ["DAR", "Comisión"],
    script: `Hola [aliado] 👋\n\nEntendemos la preocupación, pero el DAR (Descuento Asumido por Rappi) nunca reduce tu pago neto.\n\nAsí funciona: Rappi financia el descuento al usuario final. Tu factura se emite al precio de lista completo. La comisión se recalcula proporcionalmente hacia abajo para compensar el descuento, de modo que el neto que recibes es idéntico con o sin DAR.\n\nEn tu paidlot verás tres columnas DAR que se cancelan entre sí:\n• Descuento inversión DAR (producto)\n• Descuento sobre comisión DAR\n• Descuento sobre IVA de comisión DAR\n\n✅ El DAR es una inversión de Rappi en demanda para tu tienda. ¡Te ayuda a vender más sin costarte nada! 🙌` },
  { id: "comision-alta", icon: "🏢", title: '"La comisión es muy alta / no entiendo los cobros de plataforma"', tags: ["Comisión", "Plataforma"],
    script: `Hola [aliado] 👋\n\nLa sección "Costos de Plataforma" tiene varias líneas que vale la pena revisar juntos:\n\n1. Uso y alquiler de plataforma Rappi → tarifa base acordada en el contrato.\n2. Tarifas adicionales (transaccional, integration, etc.) → costos de procesamiento de pago.\n3. Cuota RappiAds → inversión publicitaria (solo si tiene pauta activa).\n\nTu tarifa efectiva real (comisión + impuestos ÷ ventas brutas) está en [X]%. El rango saludable para restaurantes es 15%–30%.\n\n✅ Si considerás que la tarifa base no refleja el contrato, podemos revisar el contrato y el detalle del paidlot juntos. ¡Quedo atento! 🙌` },
  { id: "cobro-no-reconozco", icon: "❓", title: '"Hay un cobro que no reconozco en mi paidlot"', tags: ["Ajustes", "Plataforma"],
    script: `Hola [aliado] 👋\n\nEntiendo la inquietud. Revisemos juntos — los cobros que más generan confusión son:\n\n📺 RappiAds semana vencida → es la pauta de la semana ANTERIOR al período, no de la actual. Rappi factura ADS en semana vencida por diseño.\n🔄 Ajustes manuales → devoluciones o cancelaciones aplicadas al período.\n💳 Tarifa transaccional → costo de procesamiento del método de pago del usuario.\n📅 Deuda períodos anteriores → saldo pendiente de períodos previos.\n\n✅ Comparte el monto exacto y la fecha aproximada y lo ubicamos en el detalle del paidlot en segundos. 🙌` },
  { id: "impuestos-incorrectos", icon: "🧾", title: '"Los impuestos están mal / son muy altos"', tags: ["Impuestos", "Fiscal"],
    script: `Hola [aliado] 👋\n\nLos impuestos en el paidlot se aplican conforme a la normativa fiscal vigente. Sin embargo, si el monto parece excesivo, puede haber dos causas:\n\n1. Categoría fiscal desactualizada → si tu actividad económica cambió y no fue actualizado ante el organismo fiscal, puede aplicarse una alícuota mayor.\n2. Certificados de exención vencidos → si tenías exención de algún impuesto (ej. IIBB, ReteICA), al vencer el certificado se retoma la alícuota normal.\n\n✅ Acción: Revisa con tu contador si los certificados están vigentes. El monto pagado en exceso es recuperable como saldo a favor en tu próxima declaración. ¡Te ayudo con el proceso! 🙌` },
  { id: "total-no-cuadra", icon: "✅", title: '"El total a pagar no me cuadra / no coincide con mis cuentas"', tags: ["Total", "Conciliación"],
    script: `Hola [aliado] 👋\n\nLa fórmula del Total a Pagar es:\n\nVenta Bruta\n− DAR (financiado por Rappi, no te afecta)\n− Descuentos sobre la venta (compensaciones, cashback)\n− Comisión de plataforma (y tarifas adicionales)\n− Impuestos del período\n± Ajustes manuales\n± Deuda períodos anteriores\n= Total a Pagar\n\nSi la diferencia es menor a $1, puede ser redondeo. Si es mayor, identifiquemos qué línea está descuadrando.\n\n✅ Comparte tu cálculo propio y el monto de la diferencia y lo revisamos juntos ahora. 🙌` },
  { id: "ads-no-funciona", icon: "📺", title: '"RappiAds no me genera ventas / es muy caro"', tags: ["ADS", "RappiAds"],
    script: `Hola [aliado] 👋\n\nEl ROI de RappiAds depende de varios factores:\n\n⏰ Horario de pauta → ¿está activa en tu hora pico?\n📍 Segmentación → ¿el radio de entrega está correctamente configurado?\n📊 Umbral saludable → la inversión en ADS debería ser entre 5% y 15% de las ventas brutas. Si supera el 20%, el ROI puede volverse negativo.\n\nTu inversión actual es de [MONTO] ([X]% de ventas brutas).\n\n✅ Revisemos el reporte de RappiAds para ver el costo por pedido generado. Si el CPA es mayor al margen del pedido, conviene pausar y reconfigurar la pauta. 🙌` },
  { id: "compensaciones-incorrectas", icon: "🔄", title: '"Las compensaciones / devoluciones están mal"', tags: ["Compensaciones", "Descuentos"],
    script: `Hola [aliado] 👋\n\nLas compensaciones representan órdenes canceladas o con incidencias donde Rappi reembolsó al usuario. Las causas más frecuentes:\n\n🕐 Pedido tardío (tiempo de preparación excedido)\n🚫 Producto agotado no informado oportunamente\n⭐ Problema de calidad reportado por el usuario\n❌ Cancelación iniciada por la tienda\n\nCada compensación tiene un "Reason" (código) que indica el motivo exacto.\n\n✅ Si considerás que alguna fue aplicada incorrectamente, podemos revisar el ID de orden y la razón. Si hay un error, se gestiona la reversión vía soporte. Comparte el ID y lo reviso. 🙌` },
  { id: "pago-tarde", icon: "📅", title: '"¿Por qué no me han pagado? / El pago está atrasado"', tags: ["Pago", "Transferencia"],
    script: `Hola [aliado] 👋\n\nEl paidlot muestra la fecha de pago acordada como [FECHA_PAGO]. Las transferencias siguen este flujo:\n\n1. Cierre del período de facturación.\n2. Generación del paidlot (liquidación).\n3. Transferencia bancaria (1–3 días hábiles según el banco del aliado).\n\nAlgunas demoras frecuentes:\n🏦 Validación bancaria pendiente → confirmar que los datos bancarios estén actualizados en el portal.\n📋 Documentación fiscal incompleta → especialmente para aliados nuevos o con cambios recientes.\n\n✅ Si la fecha de pago ya pasó y no se acreditó, comunícate con soporte de Rappi adjuntando el número de paidlot. ¡Te ayudo a gestionar el caso! 🙌` },
];

// ── ObjecionesPanel — muestra libretos de objeciones con buscador interno ────
const ObjecionesPanel = memo(({ objections, search }) => {
  const [copied, setCopied] = useState(null);
  const [openId, setOpenId] = useState(null);
  const filtered = search
    ? objections.filter(o =>
        o.title.toLowerCase().includes(search.toLowerCase()) ||
        o.script.toLowerCase().includes(search.toLowerCase()) ||
        o.tags.some(t => t.toLowerCase().includes(search.toLowerCase()))
      )
    : objections;
  return (
    <div>
      <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16, padding: "10px 14px", background: "#f8fafc", borderRadius: 10, border: "1px solid #e2e8f0" }}>
        💡 <strong>Libretos listos</strong> para responder objeciones del aliado. Copia el texto, personaliza los datos entre corchetes <code>[así]</code> y envíalo por WhatsApp o chat.
      </div>
      {filtered.length === 0 && (
        <div style={{ textAlign: "center", color: "#94a3b8", padding: "32px 0", fontSize: 13 }}>Sin resultados para "{search}"</div>
      )}
      {filtered.map(obj => {
        const isOpen = openId === obj.id;
        return (
          <div key={obj.id} style={{ marginBottom: 8, borderRadius: 12, border: `1.5px solid ${isOpen ? "#fbbf24" : "#e2e8f0"}`, overflow: "hidden", transition: "border-color 0.2s" }}>
            <button
              onClick={() => setOpenId(isOpen ? null : obj.id)}
              style={{ width: "100%", background: isOpen ? "#fffbeb" : "#f8fafc", padding: "11px 16px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, border: "none", cursor: "pointer", textAlign: "left" }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 10, minWidth: 0 }}>
                <span style={{ fontSize: 20, flexShrink: 0 }}>{obj.icon}</span>
                <div style={{ minWidth: 0 }}>
                  <div style={{ fontWeight: 800, fontSize: 13, color: "#0f172a", marginBottom: 3 }}>{obj.title}</div>
                  <div style={{ display: "flex", gap: 4, flexWrap: "wrap" }}>
                    {obj.tags.map(t => (
                      <span key={t} style={{ background: "#e0f2fe", color: "#0369a1", fontSize: 9, fontWeight: 700, padding: "1px 7px", borderRadius: 20 }}>{t}</span>
                    ))}
                  </div>
                </div>
              </div>
              <span style={{ fontSize: 18, color: "#94a3b8", flexShrink: 0, display: "inline-block", transition: "transform 0.2s", transform: isOpen ? "rotate(180deg)" : "rotate(0deg)" }}>▾</span>
            </button>
            {isOpen && (
              <div style={{ padding: "12px 16px 14px", background: "#fffbeb", borderTop: "1px solid #fde68a" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <span style={{ fontSize: 9, color: "#92400e", fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.05em" }}>WhatsApp / Chat</span>
                  <button
                    onClick={(e) => { e.stopPropagation(); navigator.clipboard.writeText(obj.script); setCopied(obj.id); setTimeout(() => setCopied(null), 2500); }}
                    style={{ fontSize: 11, fontWeight: 700, padding: "5px 12px", borderRadius: 20, cursor: "pointer", background: copied === obj.id ? "#dcfce7" : "white", color: copied === obj.id ? "#166534" : "#475569", border: "1.5px solid #e2e8f0", transition: "all 0.2s" }}
                  >{copied === obj.id ? "✅ Copiado!" : "📋 Copiar libreto"}</button>
                </div>
                {obj.script.split("\n").map((line, i) => (
                  <p key={i} style={{ fontSize: 12, color: "#1e293b", lineHeight: 1.75, margin: "0 0 2px" }}>{line || " "}</p>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
});

// ── KnowledgeCenterModal — 9 tabs: 6 edu + PAGOS + CONSULTAR + OBJECIONES ────
const KC_TABS = [...EDUCATION_TABS, "PAGOS", "CONSULTAR", "OBJECIONES"];
const KC_ICONS = { DAR:"🎯", ADS:"📺", IMPUESTOS:"🧾", COMPENSACIONES:"🔄", CONCEPTOS:"📖", "FÓRMULAS":"🔢", PAGOS:"💳", CONSULTAR:"🔍", OBJECIONES:"⚡" };

const PAGOS_FAQ = [
  { q: "¿Cuándo me pagan? ¿Cuáles son los períodos de pago?", a: "Rappi tiene diferentes frecuencias de pago según el contrato del aliado:\n\n• **Semanal** — el período cierra el domingo a medianoche. El paidlot se genera el lunes y el depósito llega en 1–3 días hábiles.\n• **Quincenal** — cierres el día 15 y el último día del mes.\n• **Mensual** — cierre el último día del mes.\n\nLa fecha exacta de depósito aparece en el campo 'Fecha de Pago' en el encabezado de tu paidlot." },
  { q: "¿Por qué mi pago es diferente al que calculé?", a: "Los factores que más generan diferencias entre el cálculo propio y el paidlot son:\n\n1. **RappiAds en semana vencida** — se descuenta la pauta de la semana ANTERIOR, no la del período actual.\n2. **Compensaciones** — devoluciones por pedidos con incidencias.\n3. **Deuda períodos anteriores** — saldo pendiente de liquidaciones previas.\n4. **Ajustes manuales** — correcciones aplicadas por el equipo de Finanzas.\n5. **Impuestos** — retenciones según el país y categoría fiscal.\n\nRevisa la pestaña 'Conciliación' del paidlot para ver el desglose exacto." },
  { q: "¿Qué hacer si no recibí el pago en la fecha indicada?", a: "Verifica primero:\n\n1. **Datos bancarios** — confirmar que el CBU/IBAN/cuenta registrada en el Portal de Aliados esté actualizada y sin errores tipográficos.\n2. **Documentación fiscal** — algunos países requieren que la factura/comprobante esté enviado antes del depósito.\n3. **Retención bancaria** — algunos bancos tardan 1–2 días hábiles adicionales en acreditar transferencias externas.\n\nSi la fecha de pago ya pasó hace más de 3 días hábiles, contactar a soporte de Rappi con el número de paidlot." },
  { q: "¿Qué es el 'Total a Pagar' y cómo se calcula?", a: "El Total a Pagar es el neto que Rappi deposita al aliado:\n\nVenta Bruta\n− Descuentos de producto asumidos por el aliado\n− Costo de domicilio y propinas\n− Comisión de plataforma\n− Impuestos del período\n− Compensaciones\n− RappiAds y otras tarifas\n± Ajustes manuales\n± Deuda períodos anteriores\n= **Total a Pagar**\n\nEl DAR (Descuento Asumido por Rappi) aparece en el paidlot pero se neutraliza con un descuento equivalente en comisión — NO afecta el neto." },
  { q: "¿Qué es la 'Deuda Períodos Anteriores'?", a: "Es un saldo negativo arrastrado de un paidlot anterior donde el aliado quedó en deuda con Rappi (por ejemplo, si las compensaciones y devoluciones superaron las ventas del período).\n\nEste monto se descuenta del paidlot actual hasta saldar la deuda. El aliado puede ver el detalle de qué período originó la deuda en la columna 'Descripción' del paidlot." },
  { q: "¿Cómo actualizar los datos bancarios para recibir el pago?", a: "Los datos bancarios se actualizan en el **Portal de Aliados (Partners)**:\n\n1. Ingresar a partners.rappi.com\n2. Ir a **Mi negocio → Datos Bancarios**\n3. Actualizar CBU/CCI/CLABE/IBAN según el país\n4. El cambio puede tardar 1–2 períodos en aplicar\n\nImportante: Rappi nunca solicitará datos bancarios por WhatsApp o correo electrónico. Cualquier solicitud de este tipo es un intento de fraude." },
  { q: "¿Qué son los 'Ajustes Manuales' en el paidlot?", a: "Los ajustes manuales son correcciones realizadas por el equipo de Finanzas de Rappi. Pueden ser:\n\n• **Positivos (+)** — devolución de un cobro incorrecto de períodos anteriores, bonificación especial, etc.\n• **Negativos (−)** — corrección de un pago en exceso del período anterior.\n\nCada ajuste tiene una descripción y fecha de referencia. Si el ajuste es significativo y no lo reconoces, solicita el detalle a soporte indicando el número de paidlot." },
  { q: "¿Cómo afectan los pedidos cancelados al pago?", a: "Los pedidos cancelados generan compensaciones en el paidlot según la causa:\n\n• **Cancelado por la tienda** → el aliado asume el costo.\n• **Cancelado por el usuario** después de preparado → Rappi evalúa caso a caso.\n• **Cancelado por el repartidor** o problemas técnicos → generalmente Rappi asume el costo.\n\nCada compensación tiene un Reason Code que indica quién asumió el costo. Puedes verlos en la sección 'Compensaciones' del paidlot." },
];

const PagosPanel = memo(({ search }) => {
  const [open, setOpen] = useState(null);
  const filtered = search ? PAGOS_FAQ.filter(f => f.q.toLowerCase().includes(search.toLowerCase()) || f.a.toLowerCase().includes(search.toLowerCase())) : PAGOS_FAQ;
  return (
    <div>
      <div style={{ fontSize: 12, color: "#64748b", marginBottom: 16, padding: "10px 14px", background: "#f0fdf4", borderRadius: 10, border: "1px solid #bbf7d0" }}>
        💳 <strong>Guía de pagos y liquidaciones</strong> — Respuestas a las preguntas más frecuentes sobre el proceso de pago de Rappi.
      </div>
      {filtered.length === 0 && <div style={{ textAlign: "center", color: "#94a3b8", padding: "32px 0", fontSize: 13 }}>Sin resultados para "{search}"</div>}
      {filtered.map((f, i) => (
        <div key={i} style={{ marginBottom: 10, borderRadius: 12, border: "1.5px solid #e2e8f0", overflow: "hidden" }}>
          <div onClick={() => setOpen(open === i ? null : i)} style={{ background: open === i ? "#f0fdf4" : "#f8fafc", padding: "12px 16px", cursor: "pointer", display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
            <div style={{ fontWeight: 700, fontSize: 13, color: "#0f172a" }}>💳 {f.q}</div>
            <span style={{ fontSize: 13, color: "#94a3b8", flexShrink: 0 }}>{open === i ? "▲" : "▼"}</span>
          </div>
          {open === i && (
            <div style={{ padding: "14px 16px", borderTop: "1px solid #e2e8f0", background: "white" }}>
              {f.a.split("\n").map((line, j) => (
                <p key={j} style={{ fontSize: 12, color: "#374151", lineHeight: 1.75, margin: "0 0 4px" }}>{line || " "}</p>
              ))}
            </div>
          )}
        </div>
      ))}
    </div>
  );
});

const KnowledgeCenterModal = memo(({ country, topKpis, paidlot, allPaidlots, onClose, initialTab = "DAR", onHighlightSection, onHighlightTab }) => {
  const [tab, setTab] = useState(initialTab);
  const [search, setSearch] = useState("");
  const dar = DAR_CONFIG[country] ?? DAR_CONFIG["No detectado"];

  return (
    <div
      style={{ position: "fixed", inset: 0, zIndex: 500, background: "rgba(15,23,42,0.65)", display: "flex", alignItems: "flex-start", justifyContent: "center", padding: "20px 16px", overflowY: "auto" }}
      onClick={e => e.target === e.currentTarget && onClose()}
    >
      <div style={{ background: "white", borderRadius: 20, width: "100%", maxWidth: 980, minHeight: "80vh", maxHeight: "calc(100vh - 40px)", display: "flex", flexDirection: "column", boxShadow: "0 32px 80px rgba(0,0,0,0.35)", animation: "fadeIn 0.2s ease" }}>

        {/* ── Header ── */}
        <div style={{ padding: "18px 24px 0", borderBottom: "1px solid #e2e8f0", flexShrink: 0 }}>
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14, gap: 12 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
              <div style={{ width: 42, height: 42, borderRadius: 12, background: "linear-gradient(135deg,#ff441f,#ff6b47)", display: "flex", alignItems: "center", justifyContent: "center", fontSize: 22 }}>🎓</div>
              <div>
                <div style={{ fontWeight: 900, fontSize: 16, color: "#0f172a", lineHeight: 1.2 }}>Centro de Conocimiento Rappi</div>
                <div style={{ fontSize: 11, color: "#64748b" }}>{dar.flag} {country} · Educación · Consultas · Libretos para objeciones del aliado</div>
              </div>
            </div>
            <div style={{ display: "flex", gap: 8, alignItems: "center", flexShrink: 0 }}>
              <div style={{ position: "relative" }}>
                <input
                  placeholder="Buscar en el hub…"
                  value={search}
                  onChange={e => setSearch(e.target.value)}
                  style={{ padding: "7px 12px 7px 32px", borderRadius: 20, border: "1.5px solid #e2e8f0", fontSize: 12, width: 200, outline: "none", color: "#0f172a", background: "#ffffff" }}
                />
                <span style={{ position: "absolute", left: 10, top: "50%", transform: "translateY(-50%)", fontSize: 13, color: "#94a3b8" }}>🔍</span>
              </div>
              <button onClick={onClose} style={{ background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 10, padding: "7px 13px", cursor: "pointer", fontSize: 13, color: "#64748b", fontWeight: 700 }}>✕</button>
            </div>
          </div>

          {/* Tab strip */}
          <div style={{ display: "flex", gap: 0, overflowX: "auto" }}>
            {KC_TABS.map(t => {
              const isSpecial = t === "CONSULTAR" || t === "OBJECIONES";
              return (
                <button key={t} onClick={() => setTab(t)}
                  style={{ padding: "9px 14px", border: "none", borderBottom: tab === t ? `2.5px solid ${isSpecial ? (t === "OBJECIONES" ? "#f97316" : "#7c3aed") : "#ff441f"}` : "2.5px solid transparent",
                    background: "none", cursor: "pointer", fontSize: 11, fontWeight: tab === t ? 800 : 500,
                    color: tab === t ? (isSpecial ? (t === "OBJECIONES" ? "#f97316" : "#7c3aed") : "#ff441f") : "#64748b",
                    whiteSpace: "nowrap", display: "flex", alignItems: "center", gap: 4 }}>
                  <span>{KC_ICONS[t]}</span>{t}
                </button>
              );
            })}
          </div>
        </div>

        {/* ── Content ── */}
        <div style={{ flex: 1, overflowY: "auto", padding: "20px 24px" }}>
          {/* Education tabs */}
          {EDUCATION_TABS.includes(tab) && (
            <EducationHub country={country} topKpis={topKpis} embedded={true} embeddedTab={tab} />
          )}

          {/* Pagos FAQ */}
          {tab === "PAGOS" && <PagosPanel search={search} />}

          {/* Consultar — full DoubtAssistant inline */}
          {tab === "CONSULTAR" && (
            <div style={{ maxWidth: 680 }}>
              <div style={{ marginBottom: 16, padding: "10px 14px", background: "#faf5ff", border: "1.5px solid #d8b4fe", borderRadius: 12 }}>
                <div style={{ fontWeight: 800, fontSize: 12, color: "#7c3aed", marginBottom: 3 }}>🔍 Asistente de consultas del paidlot</div>
                <div style={{ fontSize: 11, color: "#6d28d9" }}>Describe la duda del aliado con el mayor detalle posible. Puedes mencionar montos, fechas o columnas del paidlot.</div>
              </div>
              <DoubtAssistant
                paidlot={paidlot}
                country={country}
                allPaidlots={allPaidlots}
                onHighlightSection={onHighlightSection}
                onHighlightTab={onHighlightTab}
                inline={true}
              />
            </div>
          )}

          {/* Objeciones */}
          {tab === "OBJECIONES" && (
            <ObjecionesPanel objections={OBJECTIONS} search={search} />
          )}
        </div>
      </div>
    </div>
  );
});

// ─────────────────────────────────────────────────────────────────────────────
// § 10. MAIN APP
// ─────────────────────────────────────────────────────────────────────────────

// Limpieza mensual del historial de chat (día 30 de cada mes)
function runMonthlyChatCleanup() {
  try {
    const today = new Date();
    if (today.getDate() < 30) return; // Solo ejecutar a partir del día 30
    const currentMonth = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, "0")}`;
    const lastCleared = localStorage.getItem("rappi_chat_last_cleared") ?? "";
    if (lastCleared === currentMonth) return; // Ya se limpió este mes
    // Borrar todos los historiales de chat
    const keys = Object.keys(localStorage).filter(k => k.startsWith("rappi_chat_v1_"));
    keys.forEach(k => localStorage.removeItem(k));
    localStorage.setItem("rappi_chat_last_cleared", currentMonth);
  } catch {}
}

export default function RappiPaidlotAuditorPro() {
  const [paidlots, setPaidlots] = useState(() => { runMonthlyChatCleanup(); return loadPaidlots(); });
  const [activeIdx, setActiveIdx] = useState(0);
  const [selected, setSelected] = useState(new Set());
  const [activeGroupKey, setActiveGroupKey] = useState(null);
  const [activeTab, setActiveTab] = useState("ordenes");
  const [loading, setLoading] = useState(false);
  const [errors, setErrors] = useState([]);
  const [countryModal, setCountryModal] = useState(null);
  const [tempCountry, setTempCountry] = useState("");
  // Highlight state driven by DoubtAssistant semantic detection
  const [highlightedGroup, setHighlightedGroup] = useState(null);
  const highlightTimeoutRef = useRef(null);
  const [selectedKpi, setSelectedKpi] = useState(null);
  const [ordersOpen, setOrdersOpen] = useState(false);
  const [adsModal, setAdsModal] = useState(false);
  const [hubOpen, setHubOpen] = useState(true);
  const [knowledgeOpen, setKnowledgeOpen] = useState(false);
  const [knowledgeTab, setKnowledgeTab] = useState("DAR");
  const [pdfLoading, setPdfLoading] = useState(false);

  const handleHighlightSection = useCallback((sectionKey) => {
    const groupMap = { impuestos: "impuestos", compensaciones: "compensaciones", dar: "dar", plataforma: "plataforma", ordenes: "ventas", ajustes: "ajustes" };
    const gKey = groupMap[sectionKey];
    if (gKey) setActiveGroupKey(gKey);
    setHighlightedGroup(gKey ?? null);
    // Scroll to transaction section after a tick
    setTimeout(() => {
      const el = document.getElementById(`section-${sectionKey}`);
      if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
    }, 100);
    clearTimeout(highlightTimeoutRef.current);
    highlightTimeoutRef.current = setTimeout(() => setHighlightedGroup(null), 5000);
  }, []);

  useEffect(() => { savePaidlots(paidlots); }, [paidlots]);

  // handleHardReset
  const handleHardReset = useCallback(() => {
    if (!window.confirm("¿Limpiar toda la auditoría? Se borrarán todos los paidlots cargados.")) return;
    try { localStorage.removeItem(LS_KEY); } catch {}
    setPaidlots([]); setActiveIdx(0); setSelected(new Set()); setActiveGroupKey(null); setActiveTab("ordenes"); setErrors([]);
  }, []);

  const handleFiles = useCallback(async (e) => {
    const files = Array.from(e.target.files ?? []);
    if (!files.length) return;
    e.target.value = "";
    setLoading(true); setErrors([]);
    const newPaidlots = [], newErrors = [];
    for (const file of files) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(new Uint8Array(buf), { type: "array" });
        newPaidlots.push(parseWorkbook(wb, null));
      } catch (err) { newErrors.push(`${file.name}: ${err.message}`); }
    }
    if (newErrors.length) setErrors(newErrors);
    if (newPaidlots.length) {
      const startIdx = paidlots.length;
      setPaidlots(prev => [...prev, ...newPaidlots]);
      setActiveIdx(startIdx + newPaidlots.length - 1);
      setActiveGroupKey(null); setActiveTab("ordenes");
      const ambiguous = newPaidlots.findIndex(p => p.detection.confidence === "low");
      if (ambiguous >= 0) { setTempCountry(newPaidlots[ambiguous].detection.country); setCountryModal({ idx: startIdx + ambiguous }); }
    }
    setLoading(false);
  }, [paidlots]);

  const confirmCountry = useCallback(() => {
    if (!countryModal) return;
    setPaidlots(prev => prev.map((p, i) => i === countryModal.idx ? { ...p, detection: { country: tempCountry, confidence: "manual" } } : p));
    setCountryModal(null);
  }, [countryModal, tempCountry]);

  const removePaidlot = useCallback((idx) => {
    const p = paidlots[idx];
    if (p) setSelected(prev => { const s = new Set(prev); s.delete(p.id); return s; });
    setPaidlots(prev => prev.filter((_, i) => i !== idx));
    setActiveIdx(i => Math.max(0, i >= idx ? i - 1 : i));
  }, [paidlots]);

  const toggleSelect = useCallback((id) => {
    setSelected(prev => { const s = new Set(prev); s.has(id) ? s.delete(id) : s.add(id); return s; });
  }, []);

  const selectedPaidlots = useMemo(() => paidlots.filter(p => selected.has(p.id)), [paidlots, selected]);
  const exportOk = useMemo(() => {
    if (!selectedPaidlots.length) return { ok: false, msg: "Selecciona al menos un paidlot." };
    const ids = new Set(selectedPaidlots.map(p => p.meta.tiendaId));
    if (ids.size > 1) return { ok: false, msg: `IDs de tienda distintos: ${[...ids].join(", ")}` };
    return { ok: true };
  }, [selectedPaidlots]);

  // Duplicate order detection across all paidlots
  const duplicateOrders = useMemo(() => {
    const seen = new Map();
    for (const p of paidlots) {
      for (const o of p.ordersTable) {
        if (!o.ordenId || o.ordenId === "—") continue;
        if (!seen.has(o.ordenId)) seen.set(o.ordenId, []);
        seen.get(o.ordenId).push(p.meta.paidlotId);
      }
    }
    const dupes = new Set();
    seen.forEach((ids, ordenId) => { if (ids.length > 1) dupes.add(ordenId); });
    return dupes;
  }, [paidlots]);

  const activePaidlot = paidlots[activeIdx] ?? null;
  const activeCountry = activePaidlot?.detection?.country ?? "No detectado";
  const activeCfg = CONFIG.countries[activeCountry] ?? CONFIG.countries["No detectado"];
  const activeGroups = activePaidlot?.groups ?? [];
  const currentGroupKey = activeGroupKey ?? activeGroups[0]?.key;

  const handleExportPDF = useCallback(async () => {
    if (!activePaidlot) return;
    const country = activeCountry;
    const cfg = CONFIG.countries[country] ?? CONFIG.countries["No detectado"];
    const fmtV = (v) => { try { return new Intl.NumberFormat(cfg.locale, { style: "currency", currency: cfg.currency, maximumFractionDigits: 0 }).format(v ?? 0); } catch { return String(v ?? 0); } };
    setPdfLoading(true);
    let aiInsights = "";
    try {
      const groqKey = import.meta.env.VITE_GROQ_API_KEY ?? "";
      const periodsData = selectedPaidlots.map(p => {
        const kpi = p.topKpis;
        const tax = kpi.impuestosTotalExacto ?? kpi.totalImpuestos ?? 0;
        return `Período: ${p.resumen.inicio}→${p.resumen.fin} | Ventas: ${fmtV(kpi.ventaBruta)} | Neto: ${fmtV(kpi.totalAPagar ?? kpi.neto)} | Comisión: ${fmtV(kpi.comision)} | Impuestos: ${fmtV(tax)} (${kpi.ventaBruta > 0 ? ((tax/kpi.ventaBruta)*100).toFixed(1) : 0}%) | TarifaEfectiva: ${(kpi.effectiveFee*100).toFixed(1)}% | DAR: ${kpi.hasDar ? "Sí "+fmtV(kpi.darInversionTotal) : "No"} | ADS: ${fmtV(kpi.cuotaRappiAds??0)} | Órdenes: ${p.ordersTable.length}`;
      }).join("\n");
      const prompt = selectedPaidlots.length > 1
        ? `Analiza ${selectedPaidlots.length} períodos de ${selectedPaidlots[0].meta.tienda} en ${country}:\n${periodsData}\n\nDa: 1) Resumen ejecutivo (3 oraciones). 2) 3 insights de cambios entre períodos. 3) 2 acciones prioritarias. En español, sin markdown. Máx 200 palabras.`
        : `Analiza el paidlot de ${selectedPaidlots[0].meta.tienda} en ${country}:\n${periodsData}\n\nDa: 1) Resumen (2 oraciones). 2) 2 insights. 3) 1 acción prioritaria. En español, sin markdown. Máx 120 palabras.`;
      const res = await fetch("https://api.groq.com/openai/v1/chat/completions", {
        method: "POST",
        headers: { "Content-Type": "application/json", "Authorization": `Bearer ${groqKey}` },
        body: JSON.stringify({ model: "llama-3.3-70b-versatile", messages: [{ role: "system", content: "Eres asesor comercial de Rappi. Analiza paidlots de aliados en español, tono positivo. Destaca oportunidades de RappiAds y mejoras legales de impuestos. NUNCA sugieras vacíos legales ni evasión. NUNCA digas que el aliado puede activar o invertir en DAR — DAR es decisión de Rappi. Conciso, sin markdown." }, { role: "user", content: prompt }], max_tokens: 300, temperature: 0.4 }),
      });
      const data = await res.json();
      aiInsights = data.choices?.[0]?.message?.content ?? "";
    } catch { aiInsights = ""; }
    setPdfLoading(false);
    exportPDF(activePaidlot, country, aiInsights, selectedPaidlots);
  }, [activePaidlot, activeCountry, selectedPaidlots]);
  const currentGroupCfg = CONFIG.groups.find(g => g.key === currentGroupKey);
  const currentGroupData = activePaidlot?.groups.find(g => g.key === currentGroupKey);

  const orderRows = useMemo(() => {
    if (!activePaidlot) return [];
    return activePaidlot.ordersTable.map(o => {
      const isDupe = duplicateOrders.has(o.ordenId);
      return [
        o.fecha,
        <span style={{ fontFamily: "monospace", fontSize: 11, display: "inline-flex", alignItems: "center", gap: 4 }}>
          {o.ordenId}
          {isDupe && <Badge label="⚠️ Duplicado" style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5" }} />}
        </span>,
        o.tienda, o.metodoPago,
        o.prime ? <Badge label="★ Prime" style={{ background: "#fef3c7", color: "#92400e" }} /> : "",
        <span style={{ fontWeight: 700 }}>{fmt(o.ventaBruta, activeCountry)}</span>,
        o.darInversion > 0 ? <span style={{ color: "#f97316", fontWeight: 700 }}>{fmt(o.darInversion, activeCountry)}</span> : <span style={{ color: "#cbd5e1" }}>—</span>,
        <span style={{ color: "#ef4444" }}>{fmt(Math.abs(o.comision), activeCountry)}</span>,
        <span style={{ fontWeight: 800, color: "#10b981" }}>{fmt(o.neto, activeCountry)}</span>,
      ];
    });
  }, [activePaidlot, activeCountry, duplicateOrders]);

  const compRowsUI = useMemo(() => {
    if (!activePaidlot) return [];
    return activePaidlot.compRows.map(c => [c.fecha, c.orderId, <Badge label={c.razon.slice(0, 30)} style={{ background: "#fef3c7", color: "#92400e" }} />, fmt(c.monto, activeCountry), <span style={{ fontSize: 11, color: "#64748b" }}>{c.comentario?.slice(0, 70)}</span>]);
  }, [activePaidlot, activeCountry]);

  const extraRows = useMemo(() => {
    if (!activePaidlot) return [];
    return activePaidlot.extrasTable.map(r => [r.fecha, r.ordenId, r.tipo, fmt(r.ventaBruta, activeCountry), fmt(r.neto, activeCountry)]);
  }, [activePaidlot, activeCountry]);

  const ajusteRows = useMemo(() => {
    if (!activePaidlot) return [];
    return activePaidlot.ajustesRows.map(r => [r.fecha, r.ordenId, <Badge label={r.razon.slice(0, 30)} style={{ background: "#f1f5f9", color: "#475569" }} />, <span style={{ color: r.ajuste < 0 ? "#dc2626" : "#10b981", fontWeight: 700 }}>{fmt(r.ajuste, activeCountry)}</span>, <span style={{ color: r.deuda !== 0 ? "#dc2626" : "#94a3b8", fontWeight: 700 }}>{fmt(r.deuda, activeCountry)}</span>, <span style={{ fontSize: 11, color: "#64748b" }}>{r.descripcion.slice(0, 60)}</span>]);
  }, [activePaidlot, activeCountry]);

  const COUNTRY_LIST = Object.keys(CONFIG.countries);

  return (
    <div style={{ minHeight: "100vh", background: "#f1f5f9", fontFamily: "'Segoe UI', system-ui, -apple-system, sans-serif", color: "#0f172a" }}>

      {/* Country modal */}
      {countryModal && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 999 }}>
          <div style={{ background: "white", borderRadius: 20, padding: "28px 32px", maxWidth: 460, width: "90%", boxShadow: "0 24px 64px rgba(0,0,0,0.2)" }}>
            <div style={{ fontSize: 28, marginBottom: 12 }}>🌎</div>
            <h3 style={{ margin: "0 0 8px", fontSize: 18, fontWeight: 800 }}>Confirmar país</h3>
            <p style={{ color: "#64748b", fontSize: 13, lineHeight: 1.6, marginBottom: 16 }}>No encontramos columnas fiscales definitivas. Selecciona el país para que la guía DAR, moneda y notas crédito sean correctas.</p>
            <select value={tempCountry} onChange={e => setTempCountry(e.target.value)} style={{ width: "100%", padding: "10px 14px", borderRadius: 10, border: "1px solid #e2e8f0", fontSize: 14, fontWeight: 600, marginBottom: 20 }}>
              {COUNTRY_LIST.map(c => <option key={c}>{c}</option>)}
            </select>
            <button onClick={confirmCountry} style={{ width: "100%", background: "linear-gradient(135deg,#ff441f,#ff6b47)", color: "white", border: "none", borderRadius: 12, padding: "12px", fontWeight: 800, fontSize: 14, cursor: "pointer" }}>Confirmar</button>
          </div>
        </div>
      )}

      {/* Header */}
      <header style={{ background: "white", borderBottom: "1px solid #e2e8f0", padding: "0 20px", display: "flex", alignItems: "center", justifyContent: "space-between", height: 58, position: "sticky", top: 0, zIndex: 100 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 11 }}>
          <div style={{ width: 34, height: 34, borderRadius: 9, background: "linear-gradient(135deg,#ff441f,#ff6b47)", display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 900, color: "white", fontSize: 15 }}>R</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 800, lineHeight: 1 }}>Rappi Paidlot Auditor Pro</div>
            <div style={{ fontSize: 10, color: "#94a3b8", letterSpacing: "0.06em", textTransform: "uppercase" }}>Auditoría Contable, Inversión DAR & Ads · v4.7</div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {paidlots.length > 0 && (
            <button onClick={handleHardReset} title="Limpiar toda la sesión de auditoría" style={{ background: "#f1f5f9", border: "1px solid #e2e8f0", color: "#64748b", borderRadius: 9, padding: "7px 13px", fontSize: 11, fontWeight: 700, cursor: "pointer" }}>🗑️ Limpiar</button>
          )}
          {selectedPaidlots.length > 0 && (
            exportOk.ok
              ? <button onClick={handleExportPDF} disabled={pdfLoading} style={{ background: pdfLoading ? "#93c5fd" : "#1d4ed8", color: "white", border: "none", borderRadius: 8, padding: "7px 13px", fontSize: 12, fontWeight: 700, cursor: pdfLoading ? "wait" : "pointer", minWidth: 130 }}>
                  {pdfLoading ? "⏳ Generando IA…" : `📄 Informe PDF${selectedPaidlots.length > 1 ? ` (${selectedPaidlots.length} períodos)` : ""}`}
                </button>
              : <Tooltip text={exportOk.msg}><div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 8, padding: "6px 12px", fontSize: 12, color: "#dc2626", fontWeight: 600 }}>🚫 {selectedPaidlots.length} sel.</div></Tooltip>
          )}
          <label style={{ background: "linear-gradient(135deg,#ff441f,#ff6b47)", color: "white", borderRadius: 9, padding: "8px 14px", fontSize: 12, fontWeight: 700, cursor: "pointer" }}>
            📂 Cargar Excel
            <input type="file" accept=".xlsx,.xls" onChange={handleFiles} multiple style={{ display: "none" }} />
          </label>

        </div>
      </header>

      {/* Paidlot tabs */}
      {paidlots.length > 0 && (
        <div style={{ background: "white", borderBottom: "1px solid #e2e8f0", padding: "0 24px", display: "flex", gap: 2, overflowX: "auto", position: "sticky", top: 58, zIndex: 95 }}>
          {paidlots.map((p, i) => {
            const pcfg = CONFIG.countries[p.detection?.country] ?? CONFIG.countries["No detectado"];
            const isActive = i === activeIdx;
            return (
              <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 7, padding: "0 12px", borderBottom: isActive ? "2.5px solid #ff441f" : "2.5px solid transparent", cursor: "pointer", whiteSpace: "nowrap", height: 42, color: isActive ? "#ff441f" : "#64748b", fontWeight: isActive ? 700 : 500, fontSize: 12 }} onClick={() => { setActiveIdx(i); setActiveGroupKey(null); }}>
                <input type="checkbox" checked={selected.has(p.id)} onChange={ev => { ev.stopPropagation(); toggleSelect(p.id); }} onClick={e => e.stopPropagation()} style={{ cursor: "pointer", accentColor: "#ff441f", width: 13, height: 13 }} />
                <span>{pcfg.flag}</span>
                <span>{p.meta.tienda.slice(0, 18)}</span>
                <span style={{ fontSize: 10, color: "#94a3b8" }}>#{p.meta.paidlotId.slice(0, 7)}</span>
                {p.topKpis.hasDar && <span title="DAR activo">🎯</span>}
                {!p.topKpis.hasDar && <span title="Sin DAR">⚪</span>}
                <span onClick={ev => { ev.stopPropagation(); removePaidlot(i); }} style={{ marginLeft: 2, color: "#cbd5e1", cursor: "pointer", fontSize: 14, fontWeight: 700 }}>×</span>
              </div>
            );
          })}
        </div>
      )}

      <div style={{ maxWidth: 1600, margin: "0 auto", padding: "0 24px 32px" }}>
        <style>{`@keyframes fadeIn{from{opacity:0;transform:translateY(4px)}to{opacity:1;transform:translateY(0)}}`}</style>

        {errors.map((err, i) => (
          <div key={i} style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 10, padding: "12px 16px", marginBottom: 12, color: "#dc2626", fontSize: 13, display: "flex", justifyContent: "space-between" }}>
            <span>⚠️ {err}</span>
            <button onClick={() => setErrors(p => p.filter((_, j) => j !== i))} style={{ background: "none", border: "none", cursor: "pointer", color: "#dc2626", fontWeight: 700, fontSize: 16 }}>×</button>
          </div>
        ))}

        {!activePaidlot && !loading && (
          <div style={{ background: "white", borderRadius: 24, padding: "64px 32px", textAlign: "center", border: "2px dashed #e2e8f0" }}>
            <div style={{ fontSize: 52, marginBottom: 14 }}>🎯</div>
            <h2 style={{ fontSize: 22, fontWeight: 800, marginBottom: 8 }}>Rappi Paidlot Auditor Pro</h2>
            <p style={{ color: "#64748b", maxWidth: 520, margin: "0 auto 8px", lineHeight: 1.7, fontSize: 13 }}>Auditoría contable, inversión DAR, análisis de ADS, detección de duplicados, tarifa efectiva real y asistente IA para resolver dudas de aliados.</p>
            <p style={{ color: "#94a3b8", fontSize: 12, marginBottom: 24 }}>DAR = Descuento Asumido por Rappi · Neutralidad económica garantizada</p>
            <label style={{ background: "linear-gradient(135deg,#ff441f,#ff6b47)", color: "white", borderRadius: 14, padding: "13px 28px", fontSize: 14, fontWeight: 700, cursor: "pointer", display: "inline-block" }}>
              📂 Cargar paidlot(s)
              <input type="file" accept=".xlsx,.xls" onChange={handleFiles} multiple style={{ display: "none" }} />
            </label>
          </div>
        )}

        {loading && <div style={{ textAlign: "center", padding: "72px 0", color: "#64748b" }}><div style={{ fontSize: 36, marginBottom: 12 }}>⚙️</div><div style={{ fontWeight: 700 }}>Procesando...</div></div>}

        {activePaidlot && !loading && (() => {
          const p = activePaidlot;
          return (
            <>
              {/* ── STICKY BAR: Identity only ── */}
              <div style={{ position: "sticky", top: 100, zIndex: 89, background: "white", margin: "0 -24px", padding: "8px 24px 10px", borderBottom: "1px solid #e2e8f0", animation: "fadeIn 0.2s ease" }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 8 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontSize: 22 }}>{activeCfg.flag}</span>
                    <div>
                      <div style={{ fontSize: 15, fontWeight: 800, lineHeight: 1.2 }}>{p.meta.tienda}</div>
                      <div style={{ fontSize: 10, color: "#64748b", display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
                        <span>{activeCountry}</span>
                        <span style={{ color: "#cbd5e1" }}>·</span>
                        <span>Paidlot <b>{p.meta.paidlotId}</b></span>
                        <span style={{ color: "#cbd5e1" }}>·</span>
                        <span>{p.resumen.inicio} → {p.resumen.fin}</span>
                        {p.detection.confidence === "high" && <Badge label="auto ✓" style={{ background: "#dcfce7", color: "#166534" }} />}
                        {p.detection.confidence === "manual" && <Badge label="manual" style={{ background: "#fef3c7", color: "#92400e" }} />}
                        <ConciliationPill data={p.conciliationService} />
                      </div>
                    </div>
                  </div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", alignItems: "center" }}>
                    <div style={{ background: "#f8fafc", borderRadius: 7, padding: "4px 10px", fontSize: 10 }}><span style={{ color: "#94a3b8" }}>Pago: </span><span style={{ fontWeight: 700 }}>{p.resumen.fechaPago}</span></div>
                    <button onClick={() => { setTempCountry(activeCountry); setCountryModal({ idx: activeIdx }); }} style={{ background: "#f1f5f9", border: "1px solid #e2e8f0", borderRadius: 7, padding: "4px 10px", fontSize: 10, cursor: "pointer", fontWeight: 600 }}>🌎 País</button>
                  </div>
                </div>
              </div>

              {/* ── Alerts (below sticky bar) ── */}
              <div style={{ marginTop: 12 }}>
                <AutoAlertsBanner kpi={p.topKpis} country={activeCountry} />
                <DarZeroAlert kpis={p.topKpis} tienda={p.meta.tienda} />
                <DarKpiPanel kpis={p.topKpis} country={activeCountry} />
                <AdsAlertBanner kpis={p.topKpis} country={activeCountry} />
              </div>

              {/* ── Comparación de períodos (cuando hay múltiples seleccionados) ── */}
              {selectedPaidlots.length >= 2 && <ComparisonPanel paidlots={selectedPaidlots} country={activeCountry} />}

              {/* ── TWO-COLUMN LAYOUT: sections left + EducationHub right ── */}
              <div style={{ display: "grid", gridTemplateColumns: "minmax(0,1fr) 350px", gap: 20, alignItems: "start", marginTop: 16 }}>
                {/* LEFT column: all detail sections */}
                <div>

              {/* ═══════════════════════════════════════════════════════════
                  DETALLE COMPLETO — Todas las secciones siempre visibles
                  ═══════════════════════════════════════════════════════════ */}

              {/* ── ÓRDENES ── */}
              <div id="section-ordenes" style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: `1.5px solid ${highlightedGroup === "ventas" ? "#22c55e" : "#e2e8f0"}`, marginBottom: 10, transition: "border-color 0.4s, box-shadow 0.4s", boxShadow: highlightedGroup === "ventas" ? "0 0 0 3px #dcfce7" : "none" }}>
                <div onClick={() => setOrdersOpen(o => !o)} style={{ cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 16 }}>📦</span>
                    <h3 style={{ margin: 0, fontSize: 13, fontWeight: 800 }}>Órdenes ({p.ordersTable.length})</h3>
                    {duplicateOrders.size > 0 && <Badge label={`⚠️ ${duplicateOrders.size} duplicada(s)`} style={{ background: "#fef2f2", color: "#dc2626", border: "1px solid #fca5a5" }} />}
                  </div>
                  <span style={{ fontSize: 11, color: "#94a3b8", fontWeight: 600 }}>{ordersOpen ? "▲ Ocultar" : "▼ Ver detalle"}</span>
                </div>
                {ordersOpen && (
                  <div style={{ marginTop: 10, maxHeight: 280, overflowY: "auto", overflowX: "auto", borderRadius: 10 }}>
                    <DataTable columns={["Fecha", "Orden ID", "Tienda", "Pago", "Prime", "Venta Bruta", "DAR 🎯", "Comisión", "Neto"]} rows={orderRows} emptyMsg="No hay órdenes en este período." />
                  </div>
                )}
              </div>

              {/* ── DAR POR ORDEN ── */}
              <div id="section-dar" style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: `1.5px solid ${highlightedGroup === "dar" ? "#f97316" : "#e2e8f0"}`, marginBottom: 10, transition: "border-color 0.4s, box-shadow 0.4s", boxShadow: highlightedGroup === "dar" ? "0 0 0 3px #fff7ed" : "none" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 18 }}>🎯</span>
                    <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Inversión DAR por Orden</h3>
                    {p.topKpis.hasDar && <Badge label="Rappi paga" style={{ background: "#ffedd5", color: "#c2410c" }} />}
                  </div>
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#f97316" }}>{fmt(p.topKpis.darBeneficioTotal, activeCountry)} total período</span>
                </div>
                {!p.topKpis.hasDar
                  ? <div style={{ background: "#fef2f2", border: "1px solid #fecaca", borderRadius: 10, padding: "12px 16px", fontSize: 12, color: "#dc2626" }}>Sin inversión DAR activa en este período.</div>
                  : <div style={{ maxHeight: 300, overflowY: "auto", overflowX: "auto", borderRadius: 12 }}><DarOrderTable ordersTable={p.ordersTable} country={activeCountry} hasDar={p.topKpis.hasDar} /></div>
                }
              </div>

              {/* ── DESCUENTOS SOBRE LA VENTA ── */}
              {(() => {
                const dvGroup = p.groups.find(g => g.key === "descuentosVenta");
                const totalAPagarDV = p.topKpis.totalAPagar ?? p.topKpis.neto;
                const adsAlarm = p.topKpis.cuotaRappiAds > 0 && p.topKpis.ventaBruta > 0 && (p.topKpis.cuotaRappiAds / p.topKpis.ventaBruta) > 0.10;
                const dvTotal = dvGroup ? dvGroup.total : p.topKpis.compensaciones;
                return (
                  <div id="section-descuentosVenta" style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: `1.5px solid ${highlightedGroup === "descuentosVenta" || highlightedGroup === "compensaciones" ? "#8b5cf6" : "#e2e8f0"}`, marginBottom: 10, transition: "border-color 0.4s, box-shadow 0.4s", boxShadow: highlightedGroup === "descuentosVenta" || highlightedGroup === "compensaciones" ? "0 0 0 3px #ede9fe" : "none" }}>
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                        <span style={{ fontSize: 18 }}>🔄</span>
                        <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Descuentos sobre la venta</h3>

                      </div>
                      {dvTotal > 0 && <span style={{ fontSize: 12, fontWeight: 700, color: "#8b5cf6" }}>{fmt(dvTotal, activeCountry)} total</span>}
                    </div>
                    {dvGroup && dvGroup.items.length > 0 && (
                      <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                        {dvGroup.items.map(item => (
                          <div key={item.col} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#faf5ff", borderRadius: 9, padding: "9px 13px", borderLeft: "4px solid #8b5cf6" }}>
                            <span style={{ fontSize: 11, color: "#475569", flex: 1, marginRight: 10 }}>{item.label.length > 65 ? item.label.slice(0, 62) + "…" : item.label}</span>
                            <span style={{ fontSize: 12, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>{fmt(Math.abs(item.value), activeCountry)}</span>
                          </div>
                        ))}
                      </div>
                    )}
                    {p.compRows.length > 0 && (
                      <>
                        <div style={{ fontSize: 12, fontWeight: 700, color: "#64748b", marginBottom: 6, marginTop: dvGroup?.items.length ? 4 : 0 }}>Detalle de compensaciones ({p.compRows.length})</div>
                        <div style={{ maxHeight: 280, overflowY: "auto", overflowX: "auto", borderRadius: 12 }}>
                          <DataTable columns={["Fecha", "Orden ID", "Razón", "Monto", "Comentario"]} rows={compRowsUI} emptyMsg="Sin compensaciones en este período." />
                        </div>
                      </>
                    )}
                    {(!dvGroup || dvGroup.items.length === 0) && p.compRows.length === 0 && (
                      <div style={{ color: "#94a3b8", fontSize: 13 }}>Sin descuentos sobre la venta en este período.</div>
                    )}
                  </div>
                );
              })()}

              {/* ── IMPUESTOS ── */}
              <div id="section-impuestos" style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: `1.5px solid ${highlightedGroup === "impuestos" ? "#0ea5e9" : "#e2e8f0"}`, marginBottom: 10, transition: "border-color 0.4s, box-shadow 0.4s", boxShadow: highlightedGroup === "impuestos" ? "0 0 0 3px #e0f2fe" : "none" }}>
                {(() => {
                  const taxInfo = COUNTRY_TAX_DETAIL[activeCountry] ?? COUNTRY_TAX_DETAIL["No detectado"];
                  const taxGroup = p.groups.find(g => g.key === "impuestos");
                  return (
                    <>
                      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 10 }}>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ fontSize: 18 }}>🧾</span>
                          <div>
                            <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Impuestos — {activeCountry}</h3>
                            <div style={{ fontSize: 11, color: "#64748b" }}>{taxInfo.organismo} · {taxInfo.norma}</div>
                          </div>
                        </div>
                        <span style={{ fontSize: 12, fontWeight: 700, color: "#0ea5e9" }}>{fmt(p.topKpis.totalImpuestos, activeCountry)} total período</span>
                      </div>

                      {/* Alert note only if país has special alert flag */}
                      {taxInfo.alerta && (
                        <div style={{ padding: "8px 14px", background: "#fef2f2", borderRadius: 8, border: "1px solid #fca5a5", marginBottom: 10, fontSize: 11, color: "#dc2626" }}>
                          ⚠️ {taxInfo.nota}
                        </div>
                      )}

                      {/* Line items from paidlot — % inline from TAX_RULES match */}
                      {!taxGroup || taxGroup.items.length === 0
                        ? <div style={{ color: "#94a3b8", fontSize: 13, padding: "8px 0" }}>Sin impuestos registrados en este período.</div>
                        : (
                          <div style={{ display: "flex", flexDirection: "column", gap: 6, marginBottom: 12 }}>
                            {taxGroup.items.map(item => {
                              const rule = (p.topKpis.impuestosPorRegla ?? []).find(r => r.match === item.col);
                              return (
                                <div key={item.col} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#f0f9ff", borderRadius: 9, padding: "9px 13px", borderLeft: "4px solid #0ea5e9" }}>
                                  <div style={{ flex: 1, marginRight: 10, display: "flex", alignItems: "center", gap: 7, flexWrap: "wrap" }}>
                                    <span style={{ fontSize: 11, color: "#475569" }}>{item.label.length > 65 ? item.label.slice(0, 62) + "…" : item.label}</span>
                                    {rule?.pct && (
                                      <span style={{ fontSize: 10, fontWeight: 700, background: "#e0f2fe", color: "#0369a1", borderRadius: 20, padding: "1px 7px", whiteSpace: "nowrap" }}>{rule.pct}</span>
                                    )}
                                  </div>
                                  <span style={{ fontSize: 12, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>{fmt(Math.abs(item.value), activeCountry)}</span>
                                </div>
                              );
                            })}
                          </div>
                        )
                      }

                      {/* Official link — country-specific only */}
                      <a href={taxInfo.urlOrganismo} target="_blank" rel="noreferrer"
                        style={{ display: "inline-flex", alignItems: "center", gap: 6, color: "#2563eb", fontSize: 11, fontWeight: 600, textDecoration: "none", padding: "6px 12px", background: "#eff6ff", borderRadius: 8, border: "1px solid #bfdbfe" }}>
                        ↗ Sitio oficial {taxInfo.organismo}
                      </a>
                      {/* TaxRulesPanel — exact TAX_RULES column matches */}
                      <TaxRulesPanel topKpis={p.topKpis} country={activeCountry} />
                    </>
                  );
                })()}
              </div>

              {/* ── PLATAFORMA RAPPI ── */}
              <div id="section-plataforma" style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: `1.5px solid ${highlightedGroup === "plataforma" ? "#ef4444" : "#e2e8f0"}`, marginBottom: 10, transition: "border-color 0.4s, box-shadow 0.4s", boxShadow: highlightedGroup === "plataforma" ? "0 0 0 3px #fee2e2" : "none" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 18 }}>🏢</span>
                    <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Costos de Plataforma Rappi</h3>
                  </div>
                  <span style={{ fontSize: 12, fontWeight: 700, color: "#ef4444" }}>{fmt(p.topKpis.comision, activeCountry)} comisión base</span>
                </div>
                {(() => {
                  const platGroup = p.groups.find(g => g.key === "plataforma");
                  if (!platGroup || platGroup.items.length === 0) return <div style={{ color: "#94a3b8", fontSize: 13 }}>Sin costos de plataforma en este período.</div>;
                  return (
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {platGroup.items.map(item => (
                        <div key={item.col} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#fff5f5", borderRadius: 9, padding: "9px 13px", borderLeft: "4px solid #ef4444" }}>
                          <span style={{ fontSize: 11, color: "#475569", flex: 1, marginRight: 10 }}>{item.label.length > 65 ? item.label.slice(0, 62) + "…" : item.label}</span>
                          <span style={{ fontSize: 12, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>{fmt(Math.abs(item.value), activeCountry)}</span>
                        </div>
                      ))}
                      {p.topKpis.cuotaRappiAds > 0 && (() => {
                        const baseVentas = p.topKpis.ventaBruta;
                        const adsPct = baseVentas > 0 ? p.topKpis.cuotaRappiAds / baseVentas : 0;
                        const alertLevel = adsPct > 0.20 ? "red" : adsPct > 0.10 ? "orange" : "purple";
                        const colors = { red: ["#fef2f2","#fca5a5","#dc2626"], orange: ["#fff7ed","#fed7aa","#c2410c"], purple: ["#faf5ff","#d8b4fe","#5b21b6"] };
                        const [bg, border, text] = colors[alertLevel];
                        return (
                          <div style={{ marginTop: 8, padding: "10px 14px", background: bg, borderRadius: 10, border: `1.5px solid ${border}`, fontSize: 12, color: text, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                            <div>
                              <strong>📺 RappiAds:</strong> Cuota {fmt(p.topKpis.cuotaRappiAds, activeCountry)}
                              {baseVentas > 0 && <span style={{ marginLeft: 8, fontSize: 11, fontWeight: 700, background: "white", padding: "1px 8px", borderRadius: 20, border: `1px solid ${border}` }}>{(adsPct*100).toFixed(1)}% de ventas</span>}
                              {alertLevel !== "purple" && <span style={{ marginLeft: 6, fontSize: 10, color: text }}> — inversión alta</span>}
                            </div>
                            <button onClick={() => setAdsModal(true)} style={{ fontSize: 11, fontWeight: 700, padding: "4px 12px", background: "white", border: `1px solid ${border}`, borderRadius: 20, cursor: "pointer", color: text, whiteSpace: "nowrap" }}>Ver análisis →</button>
                          </div>
                        );
                      })()}
                      {p.topKpis.rappiAdsCollection > 0 && (
                        <div style={{ marginTop: 8, padding: "12px 14px", background: "#fffbeb", borderRadius: 10, border: "1.5px solid #fde68a", display: "flex", gap: 10, alignItems: "flex-start" }}>
                          <span style={{ fontSize: 18, flexShrink: 0 }}>📺</span>
                          <div>
                            <div style={{ fontSize: 12, fontWeight: 800, color: "#92400e", marginBottom: 3 }}>
                              RappiAds — Cobro por Semana Vencida: {fmt(p.topKpis.rappiAdsCollection, activeCountry)}
                            </div>
                            <div style={{ fontSize: 11, color: "#78350f", lineHeight: 1.6 }}>
                              Este monto corresponde a la inversión de ADS pautada en la <strong>semana anterior</strong> al período actual. Rappi factura ADS en semana vencida: un cobro que aparece en Abril corresponde legalmente a la última semana de Marzo. Es correcto y esperable. Para contrastar, solicita al aliado el reporte de inversión ADS del período previo.
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  );
                })()}
              </div>

              {/* ── EXTRA SERVICES ── */}
              {p.extrasTable.length > 0 && (
                <div style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: "1px solid #e2e8f0", marginBottom: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                    <span style={{ fontSize: 18 }}>⚡</span>
                    <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Extra Services ({p.extrasTable.length})</h3>
                  </div>
                  <DataTable columns={["Fecha", "Orden ID", "Tipo", "Venta Bruta", "Neto"]} rows={extraRows} emptyMsg="Sin extra services." />
                </div>
              )}

              {/* ── AJUSTES Y DEUDAS ── */}
              <div id="section-ajustes" style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: `1.5px solid ${highlightedGroup === "ajustes" ? "#64748b" : "#e2e8f0"}`, marginBottom: 10, transition: "border-color 0.4s, box-shadow 0.4s", boxShadow: highlightedGroup === "ajustes" ? "0 0 0 3px #f1f5f9" : "none" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 14 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{ fontSize: 18 }}>⚖️</span>
                    <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Ajustes y Deudas ({p.ajustesRows.length})</h3>
                  </div>
                  {p.topKpis.ajustesTotal > 0 && <Badge label={`⚠️ ${fmt(p.topKpis.ajustesTotal, activeCountry)}`} style={{ background: "#fffbeb", color: "#92400e", border: "1px solid #fde68a" }} />}
                </div>
                {p.ajustesRows.length === 0
                  ? <div style={{ color: "#94a3b8", fontSize: 13 }}>Sin ajustes manuales ni deudas anteriores en este período.</div>
                  : (
                    <>
                      <div style={{ background: "#fffbeb", border: "1px solid #fde68a", borderRadius: 10, padding: "10px 14px", marginBottom: 12, fontSize: 12, color: "#92400e" }}>
                        ⚖️ Los <strong>Ajustes Manuales</strong> son correcciones fuera del ciclo operativo. Las <strong>Deudas Anteriores</strong> son saldos pendientes de liquidaciones previas. Ambos afectan el neto final.
                      </div>
                      <div style={{ maxHeight: 280, overflowY: "auto", overflowX: "auto", borderRadius: 12 }}>
                        <DataTable columns={["Fecha", "Orden ID", "Razón", "Ajuste", "Deuda Ant.", "Descripción"]} rows={ajusteRows} emptyMsg="Sin ajustes." />
                      </div>
                    </>
                  )
                }
              </div>

              {/* ── OTROS DESCUENTOS (si hay) ── */}
              {(() => {
                const otrosGroup = p.groups.find(g => g.key === "otrosDescuentos");
                if (!otrosGroup || otrosGroup.items.length === 0) return null;
                return (
                  <div style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: "1px solid #e2e8f0", marginBottom: 10 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                      <span style={{ fontSize: 18 }}>📋</span>
                      <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Otros Descuentos</h3>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {otrosGroup.items.map(item => (
                        <div key={item.col} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#f8fafc", borderRadius: 9, padding: "9px 13px", borderLeft: "4px solid #475569" }}>
                          <span style={{ fontSize: 11, color: "#475569", flex: 1, marginRight: 10 }}>{item.label.length > 65 ? item.label.slice(0, 62) + "…" : item.label}</span>
                          <span style={{ fontSize: 12, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>{fmt(Math.abs(item.value), activeCountry)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })()}

              {/* ── PRÉSTAMOS (si hay) ── */}
              {(() => {
                const prestGroup = p.groups.find(g => g.key === "prestamos");
                if (!prestGroup || prestGroup.items.length === 0) return null;
                return (
                  <div style={{ background: "white", borderRadius: 14, padding: "14px 16px", border: "1px solid #e2e8f0", marginBottom: 10 }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 14 }}>
                      <span style={{ fontSize: 18 }}>🏦</span>
                      <h3 style={{ margin: 0, fontSize: 14, fontWeight: 800 }}>Préstamos</h3>
                    </div>
                    <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                      {prestGroup.items.map(item => (
                        <div key={item.col} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", background: "#f1f5f9", borderRadius: 9, padding: "9px 13px", borderLeft: "4px solid #0f172a" }}>
                          <span style={{ fontSize: 11, color: "#475569", flex: 1, marginRight: 10 }}>{item.label}</span>
                          <span style={{ fontSize: 12, fontWeight: 800, color: "#0f172a", whiteSpace: "nowrap" }}>{fmt(Math.abs(item.value), activeCountry)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                );
              })()}

                </div>{/* end LEFT column */}

                {/* RIGHT column: KPI sidebar + EducationHub */}
                <div style={{ position: "sticky", top: 158, maxHeight: "calc(100vh - 170px)", overflowY: "auto", paddingBottom: 16, display: "flex", flexDirection: "column", gap: 12 }}>

                  {/* ── KPI Panel ── */}
                  <div style={{ background: "white", borderRadius: 14, padding: "14px 14px 12px", border: "1px solid #e2e8f0", boxShadow: "0 1px 4px rgba(0,0,0,0.04)" }}>
                    <KPIPanel
                      topKpis={activePaidlot.topKpis}
                      country={activeCountry}
                      selectedKpi={selectedKpi}
                      onSelectKpi={(key) => {
                        setSelectedKpi(key);
                        const sectionMap = {
                          ventaBruta: "ordenes",
                          darInversionTotal: "dar",
                          descuentosVenta: "descuentosVenta",
                          comision: "plataforma",
                          impuestosTotalExacto: "impuestos",
                          totalAPagar: "ordenes",
                          otrosDescuentos: "ajustes",
                          prestamos: "ajustes",
                        };
                        if (sectionMap[key]) handleHighlightSection(sectionMap[key]);
                        const el = document.getElementById(`section-${sectionMap[key] ?? key}`);
                        if (el) el.scrollIntoView({ behavior: "smooth", block: "start" });
                      }}
                    />
                  </div>

                  {/* ── Knowledge Center open button ── */}
                  <div style={{ display: "flex", flexDirection: "column", gap: 6 }}>
                    <button
                      onClick={() => { setKnowledgeTab("DAR"); setKnowledgeOpen(true); }}
                      style={{ width: "100%", background: "linear-gradient(135deg,#fff7ed,#fff)", border: "1.5px solid #fed7aa", borderRadius: 12, padding: "12px 16px", display: "flex", alignItems: "center", gap: 10, cursor: "pointer", textAlign: "left" }}
                    >
                      <span style={{ fontSize: 22, flexShrink: 0 }}>🎓</span>
                      <div>
                        <div style={{ fontSize: 12, fontWeight: 800, color: "#c2410c" }}>Hub de Conocimiento</div>
                        <div style={{ fontSize: 10, color: "#92400e" }}>DAR · ADS · Impuestos · Conceptos · Fórmulas</div>
                      </div>
                      <span style={{ marginLeft: "auto", fontSize: 14, color: "#f97316" }}>→</span>
                    </button>
                    <button
                      onClick={() => { setKnowledgeTab("OBJECIONES"); setKnowledgeOpen(true); }}
                      style={{ width: "100%", background: "linear-gradient(135deg,#fff7f5,#fff)", border: "1.5px solid #fca5a5", borderRadius: 12, padding: "12px 16px", display: "flex", alignItems: "center", gap: 10, cursor: "pointer", textAlign: "left" }}
                    >
                      <span style={{ fontSize: 22, flexShrink: 0 }}>⚡</span>
                      <div>
                        <div style={{ fontSize: 12, fontWeight: 800, color: "#dc2626" }}>Libretos de Objeciones</div>
                        <div style={{ fontSize: 10, color: "#b91c1c" }}>Respuestas listas para el aliado</div>
                      </div>
                      <span style={{ marginLeft: "auto", fontSize: 14, color: "#ef4444" }}>→</span>
                    </button>
                    <button
                      onClick={() => { setKnowledgeTab("CONSULTAR"); setKnowledgeOpen(true); }}
                      style={{ width: "100%", background: "linear-gradient(135deg,#faf5ff,#fff)", border: "1.5px solid #d8b4fe", borderRadius: 12, padding: "12px 16px", display: "flex", alignItems: "center", gap: 10, cursor: "pointer", textAlign: "left" }}
                    >
                      <span style={{ fontSize: 22, flexShrink: 0 }}>🔍</span>
                      <div>
                        <div style={{ fontSize: 12, fontWeight: 800, color: "#7c3aed" }}>Consultar al Asistente</div>
                        <div style={{ fontSize: 10, color: "#6d28d9" }}>Dudas del aliado con IA + datos del paidlot</div>
                      </div>
                      <span style={{ marginLeft: "auto", fontSize: 14, color: "#7c3aed" }}>→</span>
                    </button>
                  </div>

                </div>
              </div>{/* end TWO-COLUMN GRID */}

            </>
          );
        })()}
      </div>

      {/* RappiAds analysis modal */}
      {adsModal && activePaidlot && (() => {
        const kpi = activePaidlot.topKpis;
        const baseVentas = kpi.ventaBruta;
        const adsPct = baseVentas > 0 ? kpi.cuotaRappiAds / baseVentas : 0;
        const isHigh = adsPct > 0.20;
        const isMed = adsPct > 0.10;
        const color = isHigh ? "#dc2626" : isMed ? "#c2410c" : "#5b21b6";
        const bg = isHigh ? "#fef2f2" : isMed ? "#fff7ed" : "#faf5ff";
        return (
          <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 400, padding: 24 }} onClick={() => setAdsModal(false)}>
            <div style={{ background: "white", borderRadius: 20, width: "100%", maxWidth: 480, padding: "28px 32px", boxShadow: "0 24px 64px rgba(0,0,0,0.2)", animation: "fadeIn 0.2s ease" }} onClick={e => e.stopPropagation()}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 18 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <span style={{ fontSize: 28 }}>📺</span>
                  <div>
                    <div style={{ fontWeight: 900, fontSize: 16, color: "#0f172a" }}>Análisis de Inversión ADS</div>
                    <div style={{ fontSize: 11, color: "#64748b" }}>{activePaidlot.meta.tienda} · {activeCountry}</div>
                  </div>
                </div>
                <button onClick={() => setAdsModal(false)} style={{ background: "none", border: "none", fontSize: 22, cursor: "pointer", color: "#94a3b8" }}>×</button>
              </div>

              <div style={{ background: bg, border: `2px solid ${color}`, borderRadius: 14, padding: "16px 20px", marginBottom: 18, textAlign: "center" }}>
                <div style={{ fontSize: 36, fontWeight: 900, color }}>{(adsPct * 100).toFixed(1)}%</div>
                <div style={{ fontSize: 12, color, fontWeight: 700, marginTop: 4 }}>de las ventas facturadas se destina a RappiAds</div>
                <div style={{ fontSize: 11, color: "#64748b", marginTop: 6 }}>
                  {fmt(kpi.cuotaRappiAds, activeCountry)} de {fmt(baseVentas, activeCountry)} ventas facturadas
                </div>
              </div>

              <div style={{ display: "flex", flexDirection: "column", gap: 10, marginBottom: 18 }}>
                {[
                  { label: "Cuota RappiAds", val: fmt(kpi.cuotaRappiAds, activeCountry), color: "#7c3aed" },
                  { label: "Ventas Facturadas", val: fmt(baseVentas, activeCountry), color: "#10b981" },
                  { label: "% invertido en ADS", val: `${(adsPct*100).toFixed(2)}%`, color },
                  ...(kpi.rappiAdsCollection > 0 ? [{ label: "Cobro ADS semana vencida", val: fmt(kpi.rappiAdsCollection, activeCountry), color: "#f59e0b" }] : []),
                ].map(row => (
                  <div key={row.label} style={{ display: "flex", justifyContent: "space-between", padding: "8px 12px", background: "#f8fafc", borderRadius: 8 }}>
                    <span style={{ fontSize: 12, color: "#475569" }}>{row.label}</span>
                    <span style={{ fontSize: 13, fontWeight: 800, color: row.color }}>{row.val}</span>
                  </div>
                ))}
              </div>

              <div style={{ padding: "12px 16px", background: isHigh ? "#fef2f2" : isMed ? "#fffbeb" : "#f0fdf4", borderRadius: 10, border: `1px solid ${isHigh ? "#fca5a5" : isMed ? "#fde68a" : "#86efac"}`, fontSize: 12, color: isHigh ? "#7f1d1d" : isMed ? "#78350f" : "#166534", lineHeight: 1.6 }}>
                {isHigh
                  ? "⚠️ La inversión en ADS supera el 20% del neto. Esto compromete seriamente el flujo de caja del aliado. Evalúa urgentemente el ROI de la pauta y considera reducir el presupuesto."
                  : isMed
                  ? "🟡 La inversión en ADS es moderada-alta (10–20% del neto). Verifica con el aliado que las campañas estén generando retorno visible en pedidos y GMV."
                  : "✅ La inversión en ADS está en rango saludable (<10% del neto). El aliado mantiene un balance adecuado entre publicidad y flujo de caja."}
              </div>
            </div>
          </div>
        );
      })()}

      {/* Floating doubt assistant — clicking opens KnowledgeCenter to CONSULTAR */}
      <DoubtAssistant
        paidlot={activePaidlot}
        country={activeCountry}
        allPaidlots={paidlots}
        onHighlightSection={handleHighlightSection}
        onHighlightTab={handleHighlightSection}
        onOpenKnowledge={(tab) => { setKnowledgeTab(tab ?? "CONSULTAR"); setKnowledgeOpen(true); }}
      />

      {/* Knowledge Center Modal */}
      {knowledgeOpen && activePaidlot && (
        <KnowledgeCenterModal
          country={activeCountry}
          topKpis={activePaidlot.topKpis}
          paidlot={activePaidlot}
          allPaidlots={paidlots}
          initialTab={knowledgeTab}
          onClose={() => setKnowledgeOpen(false)}
          onHighlightSection={handleHighlightSection}
          onHighlightTab={handleHighlightSection}
        />
      )}
    </div>
  );
}
