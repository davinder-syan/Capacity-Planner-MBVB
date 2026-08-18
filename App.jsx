import React, { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

/* ============================================================
   CONSTANTS
   ============================================================ */

const BRAND_NAME = "MB & VB Group Capacity Planner";
const BRAND_SHORT = "MB & VB Group";

const DEFAULT_CATEGORIES = [
  "T-Shirt", "Polo", "Jersey", "Knitted Tops", "Woven Shirt", "Blouse",
  "Dress", "Jacket", "Outerwear", "Heavy Garment", "Denim",
  "Scrubs", "Pants", "Knitted Shirt", "Jogger", "Hoodie"
];
const DEFAULT_SEASONS = ["SS26", "AW26", "SS27", "AW27", "SS28"];
const DEFAULT_CUSTOMERS = ["Decathlon", "JD Sports", "LT Apparel", "LC Waikiki", "Costco", "Fashion Option", "Fabletics", "Uniteks", "Asmara", "Inditex"];
const DEFAULT_BRANDS = ["Domyos", "Kiprun", "Wedze", "Quechua", "Solognac", "Inesis", "Soneti London", "Hollister", "32 Degree", "H&M", "Zara"];

const ADDITIONAL_PROCESS_ALL = ["Printing", "Embroidery", "Washing", "Garment Dye", "Special Finishing", "Other"];
const ADDITIONAL_PROCESS_TRACKED = ["Printing", "Embroidery", "Washing", "Garment Dye"];
const WASH_TYPES = ["Normal Wash", "Garment Wash", "Enzyme Wash", "Stone Wash", "Acid Wash", "Bleach Wash", "Softener Wash", "Silicone Wash", "Bio Wash", "Bio-Polish", "Pigment Wash", "Vintage Wash", "Towel Wash", "Sand Wash", "Special / Other"];
const RAW_MATERIAL_OPTIONS = ["All Raw Material Available", "Partially Available", "Not Available"];
const ORDER_TYPES = ["Projection", "Confirmed Order"];
const MM_HEADS = ["Naresh Nagda", "Davinder Singh", "Jaidev Tegginamani", "Jaitin Naswa"];

const LOCATIONS = ["Ismailia", "Port Said", "Suez", "Alexanderia"];
const LOCATION_LABELS = {
  "Ismailia": "Ismailia Factories", "Port Said": "Port Said Factories",
  "Suez": "Suez Factories", "Alexanderia": "Alexanderia Factories"
};

const DEFAULT_SETTINGS = { urgentThresholdDays: 30, normalPlanningDays: 90, minLeadTimeDays: 30, projectionAdvanceMinDays: 120, projectionAdvanceMaxDays: 180, ppcPassword: "1234" };
const DEFAULT_COUNTERS = { nextItemNumber: 100, nextProjectionNumber: 100 };

const STORAGE_KEY = "mbvb-group-capacity-planner-v8";

const NAV = [
  { id: "book", label: "Capacity booking", icon: "plus", open: true },
  { id: "dashboard", label: "Dashboard", icon: "grid" },
  { id: "factories", label: "Factory planning", icon: "building" },
  { id: "lines", label: "Production lines", icon: "rows" },
  { id: "orders", label: "Style / order master", icon: "list" },
  { id: "loading", label: "Line loading board", icon: "gauge" },
  { id: "schedule", label: "Production planning", icon: "calendar" },
  { id: "risk", label: "Capacity risk & reminders", icon: "alert" },
  { id: "process", label: "Additional processes", icon: "spark" },
  { id: "reports", label: "Reports", icon: "doc" },
  { id: "settings", label: "Settings", icon: "cog" }
];

/* ============================================================
   DATE HELPERS  (dates stored internally as 'YYYY-MM-DD'; always DISPLAYED as DD/MM/YY)
   ============================================================ */

function parseDate(s) { return new Date(s + "T00:00:00"); }
function fmt(d) {
  const y = d.getFullYear(), m = String(d.getMonth() + 1).padStart(2, "0"), day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}
function addDays(d, n) { const nd = new Date(d); nd.setDate(nd.getDate() + n); return nd; }
function isWorkingDay(date, workingDaysPerWeek) {
  const dow = date.getDay();
  if (workingDaysPerWeek >= 7) return true;
  if (workingDaysPerWeek === 6) return dow !== 0;
  if (workingDaysPerWeek === 5) return dow !== 0 && dow !== 6;
  return dow !== 0;
}
function countWorkingDays(startStr, endStr, workingDaysPerWeek) {
  if (!startStr || !endStr) return 0;
  let start = parseDate(startStr), end = parseDate(endStr);
  if (end < start) return 0;
  let n = 0, d = new Date(start), guard = 0;
  while (d <= end && guard < 4000) {
    if (isWorkingDay(d, workingDaysPerWeek)) n++;
    d = addDays(d, 1); guard++;
  }
  return n;
}
function daysBetween(aStr, bStr) {
  if (!aStr || !bStr) return null;
  return Math.round((parseDate(bStr) - parseDate(aStr)) / 86400000);
}
function niceDate(s) {
  if (!s) return "—";
  const d = parseDate(s);
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = String(d.getFullYear()).slice(-2);
  return `${dd}/${mm}/${yy}`;
}
function todayStr() { return fmt(new Date()); }
function parseDMY(str) {
  const m = (str || "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/);
  if (!m) return null;
  let [, d, mo, y] = m;
  d = d.padStart(2, "0"); mo = mo.padStart(2, "0");
  if (y.length === 2) y = "20" + y;
  const day = Number(d), month = Number(mo);
  if (month < 1 || month > 12 || day < 1 || day > 31) return null;
  return `${y}-${mo}-${d}`;
}
function maxDateStr(a, b) { return a >= b ? a : b; }
function fmtNum(n) {
  if (n === null || n === undefined || Number.isNaN(n)) return "—";
  return Math.round(n).toLocaleString();
}
function isValidEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test((s || "").trim()); }

/* ============================================================
   SAMPLE DATA
   ============================================================ */

function makeMachines(overlock, flatlock, single, double, cover, bartack, buttonhole, buttonAttach) {
  return {
    "Overlock": { qty: overlock }, "Flatlock": { qty: flatlock },
    "Single Needle": { qty: single }, "Double Needle": { qty: double },
    "Coverstitch": { qty: cover }, "Bartack": { qty: bartack },
    "Buttonhole": { qty: buttonhole }, "Button Attach": { qty: buttonAttach }
  };
}

function genRegularLines(factory, count, hallId, startIndex) {
  const out = [];
  for (let i = 0; i < count; i++) {
    const n = startIndex + i;
    out.push({
      id: `${factory.id}-L${n}`, factoryId: factory.id, hallId, lineNumber: n,
      lineName: `Line ${String(n).padStart(2, "0")}`, lineType: "Sewing",
      operators: 28 + (n % 6) * 3, helpers: 6 + (n % 3),
      workingHoursPerDay: factory.workingHoursPerDay, efficiency: 66 + (n % 9),
      overrideMinutes: null, isDenim: false, plannedPcsPerDay: null,
      machines: makeMachines(9 + (n % 4), 4 + (n % 3), 9 + (n % 3), 3 + (n % 2), 2 + (n % 2), 1 + (n % 2), 2, 1),
      active: true
    });
  }
  return out;
}
function genDenimLines(factory, count, hallId) {
  const out = [];
  for (let i = 1; i <= count; i++) {
    out.push({
      id: `${factory.id}-D${i}`, factoryId: factory.id, hallId, lineNumber: 100 + i,
      lineName: `Denim Line ${String(i).padStart(2, "0")}`, lineType: "Denim Sewing",
      operators: 36, helpers: 8, workingHoursPerDay: factory.workingHoursPerDay, efficiency: 70,
      overrideMinutes: null, isDenim: true, plannedPcsPerDay: null,
      machines: makeMachines(14, 6, 12, 8, 3, 2, 2, 2), active: true
    });
  }
  return out;
}

// Builds a line using the real line code + machine count from the factory's actual line sheet
// (e.g. "1A", "7B", "Smart", "9 ROSE"), rather than a generic generated name.
function mkRealLine(factory, hallId, lineNumber, code, machineCount) {
  const mc = machineCount || 20;
  return {
    id: `${factory.id}-L${String(code).replace(/\s+/g, "")}`, factoryId: factory.id, hallId, lineNumber,
    lineName: `Line ${code}`, lineType: "Sewing",
    operators: mc, helpers: Math.max(2, Math.round(mc * 0.2)),
    workingHoursPerDay: factory.workingHoursPerDay, efficiency: 70,
    overrideMinutes: null, isDenim: false, plannedPcsPerDay: null,
    machines: makeMachines(Math.round(mc * 0.30), Math.round(mc * 0.15), Math.round(mc * 0.28), Math.round(mc * 0.10), Math.round(mc * 0.07), Math.round(mc * 0.04), Math.round(mc * 0.04), Math.round(mc * 0.02)),
    // Reference-only backlog info from the factory's own line sheet — customer/brand currently
    // running, begin/end of the current run, day gap, and the date the line is booked through.
    // This is informational display data only; it never feeds capacity calculations.
    currentBooking: null,
    active: true
  };
}
// Attaches real customer/brand + booked-till reference info (from the factory line sheet) onto
// already-generated lines, matched by line code.
function attachBookingInfo(lines, factoryId, infoByCode) {
  return lines.map(l => {
    if (l.factoryId !== factoryId) return l;
    const code = l.lineName.replace(/^Line /, "");
    const info = infoByCode[code];
    return info ? { ...l, currentBooking: info } : l;
  });
}

function buildSampleData() {
  const factories = [
    { id: "F1", name: "Embee 1", code: "EMB1", location: "Ismailia", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 75, active: true },
    { id: "F2", name: "Embee 2", code: "EMB2", location: "Ismailia", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 73, active: true },
    { id: "F3", name: "Embee 5", code: "EMB5", location: "Ismailia", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 72, active: true },
    { id: "F4", name: "Embee 9", code: "EMB9", location: "Ismailia", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 74, active: true },
    { id: "F5", name: "Embee 11", code: "EMB11", location: "Ismailia", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 70, active: true },
    { id: "F6", name: "Plaza Port Said", code: "PLZ", location: "Port Said", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 71, active: true },
    { id: "F7", name: "Embee Suez", code: "EMBS", location: "Suez", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 70, active: true },
    { id: "F8", name: "Cannon 1", code: "CAN1", location: "Alexanderia", workingDaysPerWeek: 6, workingHoursPerDay: 8, shiftPattern: "Single shift (8h)", defaultEfficiency: 68, active: true },
    { id: "F9", name: "Globe", code: "GLB", location: "Alexanderia", workingDaysPerWeek: 6, workingHoursPerDay: 8, shiftPattern: "Single shift (8h)", defaultEfficiency: 70, active: true },
    { id: "F10", name: "Globe 5", code: "GLB5", location: "Alexanderia", workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 72, active: true }
  ];
  const byId = Object.fromEntries(factories.map(f => [f.id, f]));

  const halls = [];
  factories.forEach(f => {
    if (f.id === "F5") { for (let i = 1; i <= 4; i++) halls.push({ id: `${f.id}-H${i}`, factoryId: f.id, name: `Hall ${i}` }); }
    else { halls.push({ id: `${f.id}-H1`, factoryId: f.id, name: "Hall 1" }); }
  });

  // Real line-by-line structure for Ismailia (Embee 1/2/5/9/11), from the 12th August factory-wise sheet.
  // Other locations keep their existing generated structure, untouched.
  const F1_LINES = [["1A", 35], ["1B", 35], ["2", 35], ["3", 35], ["4", 35], ["5", 35], ["6", 35], ["7A", 21], ["7B", 23], ["8A", 21], ["8B", 23], ["Smart", 18]];
  const F2_LINES = [["1", 30], ["2", 30], ["3", 30], ["4", 30], ["5", 30], ["6", 30], ["7", 30], ["8", 30], ["9", 30]];
  const F3_LINES = [["1", 30], ["2", 30], ["3", 30], ["4", 30], ["5", 30], ["6", 30], ["7", 30], ["8", 30], ["9", 28], ["10", 28], ["11", 17]]; // Embee 5
  const F4_LINES = [["1A", 16], ["1B", 16], ["2A", 16], ["2B", 16], ["3A", 16], ["3B", 16], ["4A", 16], ["4B", 16], ["5A", 16], ["5B", 16]]; // Embee 9
  const F5_LINES = [["1", 25], ["2", 25], ["3", 25], ["4", 25], ["5", 25], ["6", 25], ["7", 25], ["8", 25], ["9", 25], ["9 ROSE", 25], ["10", 45], ["11", 45], ["12", 45], ["13", 45], ["14", 25], ["15", 45], ["16", 45], ["17", 45], ["18", 45], ["19", 45], ["21", 45], ["22", 45], ["23", 45], ["24", 45], ["25", 45]]; // Embee 11, split across Halls 1-3

  const f1Lines = attachBookingInfo(F1_LINES.map(([code, mc], i) => mkRealLine(byId.F1, "F1-H1", i + 1, code, mc)), "F1", {
    "1A": { customerBrand: "32 Degree / Sonneti / ClassRoom", beginDate: "2026-09-03", endDate: "2026-09-13", gapDays: 9, bookedTill: "2026-10-25" },
    "1B": { customerBrand: "Sonneti (merge with Line 2)", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-10-13" },
    "2": { customerBrand: "Sonneti / ClassRoom", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-12-07" },
    "3": { customerBrand: "Sonneti / 32 Degree", beginDate: "2026-09-03", endDate: "2026-09-13", gapDays: 8, bookedTill: "2026-10-15" },
    "4": { customerBrand: "Sonneti / Supply & Demand", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-10-19" },
    "5": { customerBrand: "Sonneti / Zara", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-10-20" },
    "6": { customerBrand: "Sonneti / H&M", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-10-17" },
    "7A": { customerBrand: "Sonneti (line merged with 6 & 5)", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-10-19" },
    "7B": { customerBrand: "Sonneti", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-09-03" },
    "8A": { customerBrand: "Sonneti", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-08-30" },
    "8B": { customerBrand: "Sonneti", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-09-03" },
    "Smart": { customerBrand: "Sonneti / H&M", beginDate: null, endDate: null, gapDays: null, bookedTill: "2026-08-27" }
  });
  const f2Lines = F2_LINES.map(([code, mc], i) => mkRealLine(byId.F2, "F2-H1", i + 1, code, mc));
  const f3Lines = F3_LINES.map(([code, mc], i) => mkRealLine(byId.F3, "F3-H1", i + 1, code, mc));
  const f4Lines = F4_LINES.map(([code, mc], i) => mkRealLine(byId.F4, "F4-H1", i + 1, code, mc));
  const f5Hall1 = F5_LINES.slice(0, 9).map(([code, mc], i) => mkRealLine(byId.F5, "F5-H1", i + 1, code, mc));
  const f5Hall2 = F5_LINES.slice(9, 17).map(([code, mc], i) => mkRealLine(byId.F5, "F5-H2", i + 10, code, mc));
  const f5Hall3 = F5_LINES.slice(17, 25).map(([code, mc], i) => mkRealLine(byId.F5, "F5-H3", i + 18, code, mc));

  const lines = [
    ...f1Lines, ...f2Lines, ...f3Lines, ...f4Lines,
    ...f5Hall1, ...f5Hall2, ...f5Hall3,
    ...genDenimLines(byId.F5, 4, "F5-H4"),
    ...genRegularLines(byId.F6, 20, "F6-H1", 1),
    ...genRegularLines(byId.F7, 15, "F7-H1", 1),
    ...genRegularLines(byId.F8, 2, "F8-H1", 1),
    ...genRegularLines(byId.F9, 25, "F9-H1", 1),
    ...genRegularLines(byId.F10, 18, "F10-H1", 1)
  ];
  const l1 = lines.find(l => l.id === "F1-L2"); if (l1) l1.overrideMinutes = 14400;
  const l2 = lines.find(l => l.id === "F3-L1"); if (l2) l2.overrideMinutes = 12000;
  const l3 = lines.find(l => l.id === "F8-L1"); if (l3) l3.overrideMinutes = 13000;
  const l4 = lines.find(l => l.id === "F5-D1"); if (l4) l4.overrideMinutes = 9600;
  lines.filter(l => l.factoryId === "F1" || l.factoryId === "F9").forEach(l => { l.plannedPcsPerDay = 800 + (l.lineNumber % 5) * 40; });

  const additionalProcesses = [
    { id: "SP1", name: "Printing", capacityPerDay: 10000, bookedPerDay: 8000 },
    { id: "SP2", name: "Embroidery", capacityPerDay: 12000, bookedPerDay: 5000 },
    { id: "SP3", name: "Washing", capacityPerDay: 20000, bookedPerDay: 9000 },
    { id: "SP4", name: "Garment Dye", capacityPerDay: 8000, bookedPerDay: 3000 }
  ];

  const settings = { ...DEFAULT_SETTINGS };
  const counters = { nextItemNumber: 109, nextProjectionNumber: 101 };
  const savedEmailContacts = ["planning@mbvbgroup.com", "merchandising@mbvbgroup.com"];

  function mkBooking(o) {
    const eff = o.budgetedEfficiency || 100;
    const requiredMinutes = Math.round((o.qty * o.sam) / (eff / 100));
    const dailyMinutes = o.dailyMinutes || Math.round(requiredMinutes / Math.max(1, countWorkingDays(o.startDate, o.completionDate, 6)));
    return {
      id: o.id, itemNumber: o.itemNumber, referenceNumber: o.itemNumber, orderType: o.orderType || "Confirmed Order",
      customer: o.customer, brand: o.brand, styleNumber: o.styleNumber, description: o.description,
      category: o.category, gender: o.gender || "Unisex", season: o.season || "SS26",
      mmHead: o.mmHead || MM_HEADS[0], infoAlertEmails: o.infoAlertEmails || [],
      qty: o.qty, sam: o.sam, budgetedEfficiency: eff, fob: o.fob || 0, priority: o.priority || "Normal",
      confirmed: o.confirmed !== undefined ? o.confirmed : true, bookingDate: o.startDate,
      pcd: o.pcd, deliveryDate: o.deliveryDate, startDate: o.startDate, completionDate: o.completionDate,
      factoryId: o.factoryId, hallId: o.hallId, lineId: o.lineId,
      additionalProcesses: o.additionalProcesses || [], additionalProcessRemark: o.additionalProcessRemark || null,
      embroideryStitches: o.embroideryStitches || null, embroideryCapacityConfirmed: o.embroideryCapacityConfirmed || null,
      washType: o.washType || null, rawMaterialAvailability: o.rawMaterialAvailability || "",
      requiredMinutes, dailyMinutes, status: o.status, capacityStatus: o.capacityStatus,
      specialProcessStatus: o.specialProcessStatus || (o.additionalProcesses && o.additionalProcesses.length ? "AVAILABLE" : "N/A"),
      pcdStatus: o.pcdStatus, deliveryStatus: o.deliveryStatus, buffer: o.buffer, bottleneck: o.bottleneck || null,
      plannedQty: o.plannedQty, remainingQty: o.remainingQty || 0, overrideUsed: o.overrideUsed || false, remarks: o.remarks || "",
      createdAt: o.startDate
    };
  }

  const bookings = [
    mkBooking({ id: "B1", itemNumber: "C100", orderType: "Confirmed Order", mmHead: "Naresh Nagda", customer: "Northline Apparel", brand: "Northline", styleNumber: "ABC123", description: "Crew neck tee", category: "T-Shirt", qty: 3000, sam: 10, fob: 4.2, startDate: "2026-08-01", pcd: "2026-11-25", deliveryDate: "2026-12-26", completionDate: "2026-08-15", factoryId: "F1", hallId: "F1-H1", lineId: "F1-L2", status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 3000 }),
    mkBooking({ id: "B2", itemNumber: "C101", orderType: "Confirmed Order", mmHead: "Davinder Singh", customer: "Harlow Kids", brand: "Harlow", styleNumber: "XYZ456", description: "Polo shirt", category: "Polo", qty: 2200, sam: 12, fob: 5.1, startDate: "2026-08-03", pcd: "2026-12-01", deliveryDate: "2027-01-01", completionDate: "2026-08-18", factoryId: "F3", hallId: "F3-H1", lineId: "F3-L1", status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 2200 }),
    mkBooking({ id: "B3", itemNumber: "C102", orderType: "Confirmed Order", mmHead: "Naresh Nagda", customer: "Rivermill Co", brand: "Rivermill", styleNumber: "DEF789", description: "Zip jacket", category: "Jacket", qty: 1800, sam: 28, fob: 11.5, startDate: "2026-08-01", pcd: "2026-08-20", deliveryDate: "2026-09-20", completionDate: "2026-08-19", factoryId: "F8", hallId: "F8-H1", lineId: "F8-L1", rawMaterialAvailability: "All Raw Material Available", status: "AT RISK", capacityStatus: "PARTIAL", pcdStatus: "AT RISK", deliveryStatus: "AT RISK", buffer: 31, plannedQty: 1800, bottleneck: "Urgent booking (PCD 19 days from booking) — Sewing capacity near full on Cannon 1 Line 01" }),
    mkBooking({ id: "B4", itemNumber: "C103", orderType: "Confirmed Order", mmHead: "Jaidev Tegginamani", customer: "Solvay Textiles", brand: "Solvay", styleNumber: "PRT220", description: "Printed graphic tee", category: "T-Shirt", qty: 1400, sam: 9, fob: 3.6, startDate: "2026-08-05", pcd: "2026-10-24", deliveryDate: "2026-11-24", completionDate: "2026-08-20", factoryId: "F2", hallId: "F2-H1", lineId: "F2-L3", additionalProcesses: ["Printing"], status: "AT RISK", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 1400, specialProcessStatus: "SHORTAGE", bottleneck: "Printing capacity — only 2,000 pcs/day free" }),
    mkBooking({ id: "B5", itemNumber: "C104", orderType: "Confirmed Order", mmHead: "Jaitin Naswa", customer: "Maren Studio", brand: "Maren", styleNumber: "BLZ301", description: "Tailored blouse", category: "Blouse", qty: 2600, sam: 18, fob: 8.4, startDate: "2026-08-10", pcd: "2026-09-05", deliveryDate: "2026-10-06", completionDate: "2026-09-08", factoryId: "F6", hallId: "F6-H1", lineId: "F6-L2", rawMaterialAvailability: "Partially Available", status: "RAW MATERIAL HOLD", capacityStatus: "PARTIAL", pcdStatus: "NOT ACHIEVABLE", deliveryStatus: "AT RISK", buffer: 31, plannedQty: 0, remainingQty: 2600, bottleneck: "This booking has a PCD of less than 30 days from the booking date. All Raw Material must be available to proceed with this booking." }),
    mkBooking({ id: "B6", itemNumber: "P100", orderType: "Projection", mmHead: "Davinder Singh", customer: "Coastal Kids", brand: "Coastal", styleNumber: "JRS112", description: "Jersey dress", category: "Dress", qty: 1500, sam: 20, fob: 6.2, startDate: "2026-08-12", pcd: "2026-12-01", deliveryDate: "2027-01-01", completionDate: "2026-08-30", factoryId: "F9", hallId: "F9-H1", lineId: "F9-L1", confirmed: false, status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 1500, infoAlertEmails: ["planning@mbvbgroup.com"] }),
    mkBooking({ id: "B7", itemNumber: "C105", orderType: "Confirmed Order", mmHead: "Naresh Nagda", customer: "Globe Retail", brand: "Globe", styleNumber: "OTW450", description: "Padded outerwear", category: "Outerwear", qty: 1600, sam: 32, fob: 14.0, startDate: "2026-08-04", pcd: "2026-11-29", deliveryDate: "2026-12-30", completionDate: "2026-08-25", factoryId: "F8", hallId: "F8-H1", lineId: "F8-L2", status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 1600 }),
    mkBooking({ id: "B9", itemNumber: "C106", orderType: "Confirmed Order", mmHead: "Jaidev Tegginamani", customer: "Vantage Sport", brand: "Vantage", styleNumber: "EMB700", description: "Embroidered polo", category: "Polo", qty: 2000, sam: 11, fob: 5.9, startDate: "2026-08-06", pcd: "2026-08-30", deliveryDate: "2026-09-30", completionDate: "2026-08-28", factoryId: "F2", hallId: "F2-H1", lineId: "F2-L2", additionalProcesses: ["Embroidery", "Washing"], embroideryStitches: 18000, embroideryCapacityConfirmed: "Pending", washType: "Enzyme Wash", rawMaterialAvailability: "All Raw Material Available", status: "AT RISK", capacityStatus: "AVAILABLE", pcdStatus: "AT RISK", deliveryStatus: "AT RISK", buffer: 31, plannedQty: 2000, specialProcessStatus: "SHORTAGE", bottleneck: "Embroidery capacity constrained; planner to confirm space (stitch count is informational only)" }),
    mkBooking({ id: "B10", itemNumber: "C107", orderType: "Confirmed Order", mmHead: "Jaitin Naswa", customer: "Suez Trading Co", brand: "Suez Line", styleNumber: "WSH210", description: "Oxford woven shirt", category: "Woven Shirt", qty: 1900, sam: 22, fob: 9.3, startDate: "2026-08-08", pcd: "2026-12-02", deliveryDate: "2027-01-02", completionDate: "2026-08-30", factoryId: "F7", hallId: "F7-H1", lineId: "F7-L1", status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 1900 }),
    mkBooking({ id: "B11", itemNumber: "C108", orderType: "Confirmed Order", mmHead: "Davinder Singh", customer: "Denver Denim Co", brand: "Denver", styleNumber: "DNM801", description: "5-pocket denim jean", category: "Denim", qty: 2400, sam: 30, fob: 13.5, startDate: "2026-08-05", pcd: "2026-12-01", deliveryDate: "2027-01-01", completionDate: "2026-08-25", factoryId: "F5", hallId: "F5-H4", lineId: "F5-D1", status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 31, plannedQty: 2400 }),
    mkBooking({ id: "B12", itemNumber: "P101", orderType: "Projection", mmHead: "Jaitin Naswa", customer: "Atlas Retail Group", brand: "Atlas", styleNumber: "PRJ900", description: "Fleece crew — FW26 program", category: "Knitted Tops", qty: 2200, sam: 14, fob: 6.8, startDate: "2026-08-01", pcd: "2026-11-10", deliveryDate: "2026-12-15", completionDate: "2026-08-20", factoryId: "F10", hallId: "F10-H1", lineId: "F10-L3", confirmed: false, status: "ACCEPTED", capacityStatus: "AVAILABLE", pcdStatus: "ACHIEVABLE", deliveryStatus: "ACHIEVABLE", buffer: 35, plannedQty: 2200, infoAlertEmails: ["merchandising@mbvbgroup.com"] })
  ];

  return { factories, halls, lines, bookings, additionalProcesses, settings, counters, categories: [...DEFAULT_CATEGORIES], seasons: [...DEFAULT_SEASONS], customers: [...DEFAULT_CUSTOMERS], brands: [...DEFAULT_BRANDS], savedEmailContacts };
}

/* ============================================================
   CORE CAPACITY ENGINE
   ============================================================ */

function getLineDailyMinutes(line, factory) {
  if (line.overrideMinutes) return line.overrideMinutes;
  return (factory ? factory.workingHoursPerDay : 7) * 60;
}
function dailyBookedOnLine(lineId, dateStr, bookings, excludeId) {
  return bookings.filter(b => b.lineId === lineId && b.id !== excludeId && !isBlockedStatus(b.status) && b.startDate <= dateStr && b.completionDate >= dateStr).reduce((s, b) => s + (b.dailyMinutes || 0), 0);
}
function isBlockedStatus(status) { return status === "REJECTED" || status === "RAW MATERIAL HOLD"; }

function findEarliestStart({ line, factory, bookings, fromDate, excludeId }) {
  let day = new Date(fromDate), guard = 0;
  const dailyCap = getLineDailyMinutes(line, factory);
  while (guard < 900) {
    if (isWorkingDay(day, factory.workingDaysPerWeek)) {
      const booked = dailyBookedOnLine(line.id, fmt(day), bookings, excludeId);
      if (dailyCap - booked > 0) return fmt(day);
    }
    day = addDays(day, 1); guard++;
  }
  return fmt(day);
}

function simulateLine({ line, factory, bookings, requiredMinutes, startDate, pcdDate, deliveryDate, excludeId }) {
  let cum = 0, completionDate = null, cumAtPCD = null, cumAtDelivery = null;
  const dailyCap = getLineDailyMinutes(line, factory);
  let day = new Date(startDate);
  const laterTarget = new Date(Math.max(pcdDate.getTime(), deliveryDate.getTime()));
  const horizonEnd = addDays(laterTarget, 150);
  let guard = 0;
  while (guard < 3000) {
    if (isWorkingDay(day, factory.workingDaysPerWeek)) {
      const dStr = fmt(day);
      const dailyBooked = dailyBookedOnLine(line.id, dStr, bookings, excludeId);
      const dailyAvail = Math.max(0, dailyCap - dailyBooked);
      cum += dailyAvail;
      if (completionDate === null && cum >= requiredMinutes) completionDate = fmt(day);
    }
    if (cumAtPCD === null && day >= pcdDate) cumAtPCD = cum;
    if (cumAtDelivery === null && day >= deliveryDate) cumAtDelivery = cum;
    if (completionDate !== null && cumAtPCD !== null && cumAtDelivery !== null) break;
    if (day > horizonEnd) break;
    day = addDays(day, 1); guard++;
  }
  if (cumAtPCD === null) cumAtPCD = cum;
  if (cumAtDelivery === null) cumAtDelivery = cum;
  return { completionDate, cumAtPCD, cumAtDelivery, dailyCap };
}

function pcdRules(order, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const bookingDate = order.bookingDate || todayStr();
  const daysToP = daysBetween(bookingDate, order.pcd);
  const minPCD = fmt(addDays(parseDate(bookingDate), s.urgentThresholdDays || 30));
  const normalPCD = fmt(addDays(parseDate(bookingDate), s.normalPlanningDays || 90));
  let tier = null;
  if (daysToP !== null) {
    if (daysToP >= (s.normalPlanningDays || 90)) tier = "NORMAL";
    else if (daysToP >= (s.urgentThresholdDays || 30)) tier = "SHORT";
    else tier = "URGENT";
  }
  return { daysToP, minPCD, normalPCD, tier };
}
function rawMaterialGate(order, settings) {
  const { tier, minPCD } = pcdRules(order, settings);
  if (tier !== "URGENT") {
    const info = tier === "SHORT" ? "This booking is within the normal 90-day planning window. Please ensure Raw Material planning is aligned with the production schedule." : null;
    return { blocked: false, tier, info, minPCD };
  }
  if (order.rawMaterialAvailability === "All Raw Material Available") return { blocked: false, tier: "URGENT", urgent: true, minPCD };
  const threshold = (settings && settings.urgentThresholdDays) || 30;
  const message = `This booking has a PCD of less than ${threshold} days from the booking date. All Raw Material must be available to proceed with this booking. Earliest permissible standard PCD: ${niceDate(minPCD)}.`;
  return { blocked: true, tier: "URGENT", message, minPCD };
}
function leadTimeGate(order, settings) {
  const s = settings || DEFAULT_SETTINGS;
  const minLead = s.minLeadTimeDays || 30;
  const lead = daysBetween(order.pcd, order.deliveryDate);
  if (lead === null) return { blocked: false, lead: null, minLead };
  if (lead < minLead) return { blocked: true, lead, minLead, message: `The standard production lead time is less than ${minLead} days. Please revise the Delivery / Ex-Factory Date or PCD to allow a minimum ${minLead}-day production lead time. Current lead time: ${lead} day${lead === 1 ? "" : "s"}.` };
  return { blocked: false, lead, minLead };
}
// Projection → Confirmed Order conversion reminder tier, per the 120/105/90-day schedule.
function projectionReminderTier(booking, settings) {
  if (booking.confirmed) return { tier: "N/A", daysToPCD: null };
  const s = settings || DEFAULT_SETTINGS;
  const daysToPCD = daysBetween(todayStr(), booking.pcd);
  if (daysToPCD === null) return { tier: "N/A", daysToPCD: null };
  if (daysToPCD < (s.normalPlanningDays || 90)) return { tier: "ESCALATION", daysToPCD };
  if (daysToPCD < 105) return { tier: "CONFIRMATION REQUIRED", daysToPCD };
  if (daysToPCD < (s.projectionAdvanceMinDays || 120)) return { tier: "FOLLOW-UP", daysToPCD };
  return { tier: "OK", daysToPCD };
}

function checkAndPlan({ order, line, factory, bookings, additionalProcesses, settings, excludeId }) {
  const eff = Number(order.budgetedEfficiency) > 0 ? Number(order.budgetedEfficiency) : 100;
  const requiredMinutes = Math.round((order.qty * order.sam) / (eff / 100));
  const baseline = order.bookingDate ? maxDateStr(order.bookingDate, todayStr()) : todayStr();
  const estimatedStart = findEarliestStart({ line, factory, bookings, fromDate: parseDate(baseline), excludeId });
  const startDate = parseDate(estimatedStart), pcdDate = parseDate(order.pcd), deliveryDate = parseDate(order.deliveryDate);
  const sim = simulateLine({ line, factory, bookings, requiredMinutes, startDate, pcdDate, deliveryDate, excludeId });

  const denimEligible = !line.isDenim || order.category === "Denim";
  const rawGate = rawMaterialGate(order, settings);
  const leadGate = leadTimeGate(order, settings);

  const selectedProcesses = (order.additionalProcesses || []).filter(p => p && p !== "None");
  const specialProcessIssues = [];
  selectedProcesses.forEach(name => {
    const sp = additionalProcesses.find(s => s.name === name);
    if (sp) {
      const availDaily = Math.max(0, sp.capacityPerDay - sp.bookedPerDay);
      const wd = countWorkingDays(estimatedStart, order.pcd, factory.workingDaysPerWeek);
      const availTotal = availDaily * wd;
      if (order.qty > availTotal) specialProcessIssues.push({ name, required: order.qty, available: availTotal });
    }
  });

  const capBefore = Math.max(0, getLineDailyMinutes(line, factory) - dailyBookedOnLine(line.id, estimatedStart, bookings, excludeId));

  const plannedQtyAtPCD = Math.max(0, Math.min(order.qty, Math.floor((sim.cumAtPCD * (eff / 100)) / order.sam)));
  const capacityStatus = plannedQtyAtPCD >= order.qty ? "AVAILABLE" : (plannedQtyAtPCD > 0 ? "PARTIAL" : "INSUFFICIENT");
  const pcdStatus = plannedQtyAtPCD >= order.qty ? "ACHIEVABLE" : (plannedQtyAtPCD > 0 ? "AT RISK" : "NOT ACHIEVABLE");

  let deliveryStatus, buffer;
  if (sim.completionDate) {
    buffer = daysBetween(order.deliveryDate, sim.completionDate) * -1;
    deliveryStatus = sim.completionDate <= order.deliveryDate ? "ACHIEVABLE" : "AT RISK";
  } else { buffer = null; deliveryStatus = "NOT ACHIEVABLE"; }

  const specialProcessStatus = selectedProcesses.length ? (specialProcessIssues.length ? "SHORTAGE" : "AVAILABLE") : "N/A";

  const bottlenecks = [];
  if (rawGate.blocked) bottlenecks.push(rawGate.message);
  if (leadGate.blocked) bottlenecks.push(leadGate.message);
  if (!denimEligible) bottlenecks.push("Denim line — reserved for Denim category orders only");
  if (capacityStatus !== "AVAILABLE") bottlenecks.push("Sewing capacity");
  if (specialProcessStatus === "SHORTAGE") bottlenecks.push(specialProcessIssues.map(i => `${i.name} capacity shortage (required ${fmtNum(i.required)} pcs, available ${fmtNum(i.available)} pcs before PCD)`).join("; "));

  const hardBlock = rawGate.blocked || leadGate.blocked || !denimEligible || capacityStatus === "INSUFFICIENT";
  const fullyClean = !rawGate.blocked && !leadGate.blocked && denimEligible && capacityStatus === "AVAILABLE" && specialProcessStatus !== "SHORTAGE" && pcdStatus === "ACHIEVABLE" && deliveryStatus === "ACHIEVABLE";
  const overall = fullyClean ? "GREEN" : hardBlock ? "RED" : "AMBER";

  let decision;
  if (rawGate.blocked) decision = "RAW MATERIAL HOLD";
  else if (leadGate.blocked) decision = "LEAD TIME HOLD";
  else if (overall === "GREEN") decision = "ACCEPTED";
  else if (!denimEligible || capacityStatus === "INSUFFICIENT") decision = "REJECTED";
  else if (capacityStatus === "PARTIAL") decision = "PARTIALLY ACCEPTED";
  else decision = "ACCEPTED WITH WARNING";

  const additionalDaysNeeded = pcdStatus !== "ACHIEVABLE" && sim.dailyCap > 0 ? Math.ceil((requiredMinutes - sim.cumAtPCD) / sim.dailyCap) : 0;

  return {
    requiredMinutes, dailyCap: sim.dailyCap, capBefore, capAfter: Math.max(0, capBefore - Math.min(requiredMinutes, sim.dailyCap)),
    estimatedStart, completionDate: sim.completionDate, cumAtPCD: sim.cumAtPCD, cumAtDelivery: sim.cumAtDelivery,
    denimEligible, rawGate, leadGate, specialProcessIssues, selectedProcesses,
    plannedQtyAtPCD, capacityStatus, pcdStatus, deliveryStatus, buffer,
    specialProcessStatus, bottlenecks, overall, decision, additionalDaysNeeded
  };
}

function scoreLine({ order, line, factory, bookings, additionalProcesses, settings }) {
  const result = checkAndPlan({ order, line, factory, bookings, additionalProcesses, settings });
  if (!result.denimEligible || result.rawGate.blocked || result.leadGate.blocked) return { line, result, score: 0 };
  let score = 0;
  score += Math.max(0, Math.min(65, 65 * (result.cumAtPCD / Math.max(1, result.requiredMinutes))));
  score += result.specialProcessIssues.length === 0 ? 35 : Math.max(0, 35 - result.specialProcessIssues.length * 10);
  return { line, result, score: Math.round(score) };
}

// Sequential, persistent, never-reused numbering: C-series for Confirmed Orders, P-series for Projections.
function issueDocNumber(orderType, counters) {
  const isProjection = orderType === "Projection";
  const key = isProjection ? "nextProjectionNumber" : "nextItemNumber";
  const n = counters[key];
  const code = isProjection ? `P${n}` : `C${n}`;
  return { code, nextCounters: { ...counters, [key]: n + 1 } };
}

function groupCapacityTotals(lines, factories, bookings) {
  const factoriesById = Object.fromEntries(factories.map(f => [f.id, f]));
  const regular = lines.filter(l => l.active && !l.isDenim);
  const denim = lines.filter(l => l.active && l.isDenim);
  const today = todayStr();
  const capFor = (list) => list.reduce((s, l) => s + getLineDailyMinutes(l, factoriesById[l.factoryId]), 0);
  const bookedFor = (list) => list.reduce((s, l) => s + dailyBookedOnLine(l.id, today, bookings), 0);
  const regularCap = capFor(regular), regularBooked = bookedFor(regular);
  const denimCap = capFor(denim), denimBooked = bookedFor(denim);
  return {
    regularCap, regularBooked, regularRemaining: Math.max(0, regularCap - regularBooked), regularUtil: regularCap > 0 ? (regularBooked / regularCap) * 100 : 0,
    denimCap, denimBooked, denimRemaining: Math.max(0, denimCap - denimBooked), denimUtil: denimCap > 0 ? (denimBooked / denimCap) * 100 : 0,
    regularLineCount: regular.length, denimLineCount: denim.length
  };
}

const MANDATORY_FIELDS = [
  { key: "orderType", label: "Order Type", check: (o) => !!o.orderType },
  { key: "customer", label: "Customer", check: (o) => o.customer.trim().length > 0 },
  { key: "brand", label: "Brand", check: (o) => o.brand.trim().length > 0 },
  { key: "styleNumber", label: "Style Number", check: (o) => o.styleNumber.trim().length > 0 },
  { key: "sam", label: "SAM", check: (o) => Number(o.sam) > 0 },
  { key: "mmHead", label: "Marketing & Merchandising Head", check: (o) => !!o.mmHead },
  { key: "pcd", label: "PCD", check: (o) => !!o.pcd }
];
function missingMandatoryFields(order) { return MANDATORY_FIELDS.filter(f => !f.check(order)).map(f => f.label); }

/* ============================================================
   UI PRIMITIVES  (soft blue / grey / white theme)
   ============================================================ */

const COLORS = {
  bg: "#F1F4F8", panel: "#FFFFFF", ink: "#1B2430", inkSoft: "#57616F", inkFaint: "#8892A0",
  line: "#E2E7EE", lineStrong: "#CBD3DE",
  accent: "#3E72B0", accentSoft: "#E7EFF8", accentInk: "#234876",
  success: "#1F8A52", successSoft: "#E1F4E9",
  warning: "#C9820A", warningSoft: "#FCF0D8",
  danger: "#D93A2E", dangerSoft: "#FCE7E5",
  info: "#3E72B0", infoSoft: "#E7EFF8",
  denim: "#4C5E82", denimSoft: "#E7EAF2",
  sidebar: "#2E5C8A"
};

function StatusPill({ status }) {
  const map = {
    "AVAILABLE": "success", "ACHIEVABLE": "success", "ACCEPTED": "success", "GREEN": "success", "ORDER CONFIRMED": "success", "CONFIRMED": "success", "OK": "success", "SENT": "success",
    "AT RISK": "warning", "PARTIAL": "warning", "PARTIALLY ACCEPTED": "warning", "AMBER": "warning", "ACCEPTED WITH WARNING": "warning", "SHORTAGE": "warning", "PROJECTION": "warning", "FOLLOW-UP": "warning", "CONFIRMATION REQUIRED": "warning", "PENDING": "warning",
    "NOT ACHIEVABLE": "danger", "INSUFFICIENT": "danger", "REJECTED": "danger", "RED": "danger", "RAW MATERIAL HOLD": "danger", "LEAD TIME HOLD": "danger", "ESCALATION": "danger", "FAILED": "danger",
    "N/A": "muted"
  };
  const kind = map[status] || "muted";
  const styles = { success: { bg: COLORS.successSoft, fg: COLORS.success }, warning: { bg: COLORS.warningSoft, fg: COLORS.warning }, danger: { bg: COLORS.dangerSoft, fg: COLORS.danger }, muted: { bg: "#EEF1F5", fg: COLORS.inkFaint } }[kind];
  return <span style={{ display: "inline-flex", alignItems: "center", gap: 5, padding: "3px 9px", borderRadius: 3, background: styles.bg, color: styles.fg, fontSize: 11.5, fontWeight: 600, letterSpacing: 0.3, textTransform: "uppercase", fontFamily: "'IBM Plex Mono', monospace", whiteSpace: "nowrap" }}>{status}</span>;
}
function DenimTag() { return <span style={{ display: "inline-flex", alignItems: "center", padding: "3px 8px", borderRadius: 3, background: COLORS.denimSoft, color: COLORS.denim, fontSize: 10.5, fontWeight: 700, letterSpacing: 0.5, fontFamily: "'IBM Plex Mono', monospace" }}>DENIM</span>; }
function Card({ title, action, children, style }) {
  return (
    <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 10, boxShadow: "0 1px 2px rgba(27,36,48,0.04), 0 1px 8px rgba(27,36,48,0.04)", padding: "20px 22px", ...style }}>
      {(title || action) && (
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 16 }}>
          {title && <h3 style={{ margin: 0, fontSize: 13.5, fontWeight: 700, color: COLORS.ink, fontFamily: "'Space Grotesk', sans-serif", letterSpacing: 0.2 }}>{title}</h3>}
          {action}
        </div>
      )}
      {children}
    </div>
  );
}
function Field({ label, required, children, hint }) {
  return (
    <label style={{ display: "block", marginBottom: 12 }}>
      <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.inkSoft, marginBottom: 5, textTransform: "uppercase", letterSpacing: 0.4 }}>{label}{required && <span style={{ color: COLORS.danger }}> *</span>}</div>
      {children}
      {hint && <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 4 }}>{hint}</div>}
    </label>
  );
}
const inputStyle = { width: "100%", padding: "8px 11px", borderRadius: 7, border: `1px solid ${COLORS.lineStrong}`, fontSize: 13.5, background: "#FCFDFE", color: COLORS.ink, fontFamily: "'Inter', sans-serif", boxSizing: "border-box", transition: "border-color 0.15s, box-shadow 0.15s" };
function TextInput(props) { return <input {...props} className={`app-input ${props.className || ""}`} style={{ ...inputStyle, ...(props.style || {}) }} />; }
function Select({ children, ...props }) { return <select {...props} className={`app-input ${props.className || ""}`} style={{ ...inputStyle, ...(props.style || {}) }}>{children}</select>; }
// Displays and accepts dates as DD/MM/YY (the platform-wide format). Internally still stores
// ISO 'YYYY-MM-DD' so all date math elsewhere keeps working. The small calendar icon opens a
// native date picker (via a fully transparent overlaid <input type="date">) for convenience —
// native pickers can't be reformatted, so typed/displayed text is the source of truth for format.
function DateField({ value, onChange, style }) {
  const [text, setText] = useState(value ? niceDate(value) : "");
  useEffect(() => { setText(value ? niceDate(value) : ""); }, [value]);
  const commit = () => {
    if (!text.trim()) { onChange(""); return; }
    const iso = parseDMY(text);
    if (iso) onChange(iso); else setText(value ? niceDate(value) : "");
  };
  return (
    <div style={{ position: "relative" }}>
      <TextInput value={text} onChange={e => setText(e.target.value)} onBlur={commit}
        onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); commit(); } }}
        placeholder="DD/MM/YY" style={{ paddingRight: 34, ...(style || {}) }} />
      <input type="date" value={value || ""} onChange={e => onChange(e.target.value)}
        title="Pick a date" style={{ position: "absolute", right: 6, top: "50%", transform: "translateY(-50%)", width: 24, height: 24, opacity: 0, cursor: "pointer", border: "none", background: "transparent" }} />
      <span style={{ position: "absolute", right: 9, top: "50%", transform: "translateY(-50%)", pointerEvents: "none", color: COLORS.inkFaint }}><Icon name="calendar" size={14} /></span>
    </div>
  );
}
function Btn({ children, onClick, variant = "default", disabled, style, type = "button" }) {
  const variants = {
    default: { bg: "#FFFFFF", fg: COLORS.ink, border: COLORS.lineStrong, shadow: "none" },
    primary: { bg: COLORS.accent, fg: "#FFFFFF", border: COLORS.accent, shadow: "0 1px 2px rgba(62,114,176,0.25)" },
    success: { bg: COLORS.success, fg: "#FFFFFF", border: COLORS.success, shadow: "0 1px 2px rgba(31,138,82,0.25)" },
    ghost: { bg: "transparent", fg: COLORS.inkSoft, border: "transparent", shadow: "none" },
    danger: { bg: "#FFFFFF", fg: COLORS.danger, border: COLORS.danger, shadow: "none" }
  };
  const v = variants[variant];
  return <button type={type} onClick={onClick} disabled={disabled} className={`app-btn app-btn-${variant}`} style={{ padding: "8px 15px", borderRadius: 7, border: `1px solid ${v.border}`, background: v.bg, color: v.fg, fontSize: 12.5, fontWeight: 600, cursor: disabled ? "not-allowed" : "pointer", opacity: disabled ? 0.5 : 1, fontFamily: "'Inter', sans-serif", letterSpacing: 0.2, boxShadow: v.shadow, transition: "filter 0.15s, box-shadow 0.15s", ...style }}>{children}</button>;
}
function Toggle({ active, onClick, children }) {
  return <button type="button" onClick={onClick} style={{ padding: "5px 11px", borderRadius: 3, fontSize: 11.5, fontWeight: 500, cursor: "pointer", border: `1px solid ${active ? COLORS.info : COLORS.lineStrong}`, background: active ? COLORS.infoSoft : "#fff", color: active ? COLORS.info : COLORS.inkSoft }}>{children}</button>;
}
function UtilBar({ pct }) {
  const p = Math.max(0, Math.min(150, pct));
  const color = p >= 95 ? COLORS.danger : p >= 80 ? COLORS.warning : COLORS.success;
  return <div style={{ height: 7, background: "#E9EEF5", borderRadius: 4, overflow: "hidden" }}><div style={{ height: "100%", width: `${Math.min(100, p)}%`, background: color, borderRadius: 4 }} /></div>;
}
function Icon({ name, size = 16 }) {
  const paths = {
    grid: "M4 4h6v6H4V4zm10 0h6v6h-6V4zM4 14h6v6H4v-6zm10 0h6v6h-6v-6z",
    building: "M4 21V4h9v6h7v11H4zm2-2h2v-2H6v2zm4 0h2v-2h-2v2zm-4-5h2v-2H6v2zm4 0h2v-2h-2v2zm4 5h2v-2h-2v2zm-4-9h2V8h-2v2zm4 0h2V8h-2v2zM6 10h2V8H6v2z",
    rows: "M3 5h18v3H3V5zm0 5.5h18v3H3v-3zM3 16h18v3H3v-3z", plus: "M12 4v16m-8-8h16",
    list: "M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01",
    gauge: "M12 21a9 9 0 100-18 9 9 0 000 18zm0-9l4-4M12 3v2M3 12h2M19 12h2",
    calendar: "M7 3v3M17 3v3M4 8h16M5 6h14a1 1 0 011 1v12a1 1 0 01-1 1H5a1 1 0 01-1-1V7a1 1 0 011-1z",
    alert: "M12 3l10 18H2L12 3zm0 7v4m0 3h.01", spark: "M12 2l2 6 6 2-6 2-2 6-2-6-6-2 6-2z",
    doc: "M6 2h9l5 5v15H6V2zM14 2v6h6",
    cog: "M12 15a3 3 0 100-6 3 3 0 000 6zm8-3a8 8 0 01-.2 1.7l2 1.6-2 3.4-2.3-.9a8 8 0 01-2.9 1.7L14 22h-4l-.6-2.5a8 8 0 01-2.9-1.7l-2.3.9-2-3.4 2-1.6A8 8 0 014 12a8 8 0 01.2-1.7l-2-1.6 2-3.4 2.3.9a8 8 0 012.9-1.7L10 2h4l.6 2.5a8 8 0 012.9 1.7l2.3-.9 2 3.4-2 1.6c.13.55.2 1.12.2 1.7z",
    lock: "M6 11V8a6 6 0 1112 0v3m-14 0h16v9H4v-9z"
  };
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"><path d={paths[name] || paths.grid} /></svg>;
}
function TagListEditor({ items, onChange, placeholder }) {
  const [val, setVal] = useState("");
  const add = () => { const v = val.trim(); if (v && !items.includes(v)) { onChange([...items, v]); setVal(""); } };
  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
        {items.map(it => (<span key={it} style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, padding: "4px 8px", background: COLORS.accentSoft, color: COLORS.accentInk, borderRadius: 3, fontWeight: 600 }}>{it}<span onClick={() => onChange(items.filter(x => x !== it))} style={{ cursor: "pointer", opacity: 0.7 }}>✕</span></span>))}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <TextInput value={val} onChange={e => setVal(e.target.value)} placeholder={placeholder} onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); add(); } }} />
        <Btn onClick={add}>Add</Btn>
      </div>
    </div>
  );
}
function EmailListEditor({ items, onChange, placeholder }) {
  const [val, setVal] = useState("");
  const [err, setErr] = useState("");
  const add = () => {
    const v = val.trim();
    if (!v) return;
    if (!isValidEmail(v)) { setErr(`"${v}" doesn't look like a valid email address.`); return; }
    if (!items.includes(v)) onChange([...items, v]);
    setVal(""); setErr("");
  };
  return (
    <div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
        {items.map(it => (<span key={it} className="mono" style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11.5, padding: "4px 8px", background: COLORS.infoSoft, color: COLORS.info, borderRadius: 3, fontWeight: 600 }}>{it}<span onClick={() => onChange(items.filter(x => x !== it))} style={{ cursor: "pointer", opacity: 0.7 }}>✕</span></span>))}
      </div>
      <div style={{ display: "flex", gap: 8 }}>
        <TextInput value={val} onChange={e => { setVal(e.target.value); setErr(""); }} placeholder={placeholder} onKeyDown={e => { if (e.key === "Enter") { e.preventDefault(); add(); } }} style={err ? { border: `1px solid ${COLORS.danger}` } : {}} />
        <Btn onClick={add}>Add</Btn>
      </div>
      {err && <div style={{ fontSize: 11.5, color: COLORS.danger, marginTop: 4 }}>{err}</div>}
    </div>
  );
}
function ConfirmDialog({ message, onYes, onCancel }) {
  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,36,48,0.4)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}>
      <div style={{ background: "#fff", borderRadius: 12, padding: "24px 26px", maxWidth: 380, width: "100%", boxShadow: "0 16px 48px rgba(27,36,48,0.28)" }}>
        <div style={{ fontSize: 14, color: COLORS.ink, marginBottom: 18, lineHeight: 1.6 }}>{message}</div>
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <Btn onClick={onCancel}>Cancel</Btn>
          <Btn variant="danger" onClick={onYes}>Yes, Reset</Btn>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */

export default function App() {
  const [data, setData] = useState(null);
  const [loaded, setLoaded] = useState(false);
  // Main application username/password login is disabled for now (removed on request).
  // The PPC password gate below (Level 2, `unlocked`) is untouched and still protects every
  // tab except Capacity Booking. `currentUser` has no login to source it from right now, so
  // it stays blank — booking records just show "—" for creator until login is reinstated.
  const [currentUser] = useState("");
  const [view, setView] = useState("book");
  const [unlocked, setUnlocked] = useState(false);
  const [pendingView, setPendingView] = useState(null);
  const [pwError, setPwError] = useState("");
  const [pwInput, setPwInput] = useState("");
  const [selectedFactoryId, setSelectedFactoryId] = useState(null);
  const [orderFilter, setOrderFilter] = useState({ q: "", status: "ALL" });
  const [toast, setToast] = useState(null);
  const [dashLocation, setDashLocation] = useState("ALL");
  const [dashFactoryId, setDashFactoryId] = useState("ALL");
  const [printJob, setPrintJob] = useState(null);

  useEffect(() => { try { document.title = BRAND_NAME; } catch (e) { } }, []);

  useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, false);
        if (mounted && res && res.value) {
          const parsed = JSON.parse(res.value);
          if (!parsed.settings) parsed.settings = { ...DEFAULT_SETTINGS };
          parsed.settings = { ...DEFAULT_SETTINGS, ...parsed.settings };
          if (parsed.settings.urgentThresholdDays === 60) parsed.settings.urgentThresholdDays = 30; // one-time: old default → new default
          if (!parsed.counters) parsed.counters = { ...DEFAULT_COUNTERS };
          if (!parsed.categories) parsed.categories = [...DEFAULT_CATEGORIES];
          if (!parsed.seasons) parsed.seasons = [...DEFAULT_SEASONS];
          if (!parsed.customers) parsed.customers = [...DEFAULT_CUSTOMERS];
          // One-time, idempotent customer-list update: remove H&M/Hollister, add the three new
          // customers, but only touch those specific entries — leaves any other customisation intact.
          parsed.customers = parsed.customers.filter(c => c !== "H&M" && c !== "Hollister");
          ["Uniteks", "Asmara", "Inditex"].forEach(c => { if (!parsed.customers.includes(c)) parsed.customers.push(c); });
          if (!parsed.brands) parsed.brands = [...DEFAULT_BRANDS];
          if (!parsed.savedEmailContacts) parsed.savedEmailContacts = [];
          if (!parsed.halls) parsed.halls = [];
          if (!parsed.additionalProcesses) parsed.additionalProcesses = [];
          setData(parsed);
        } else if (mounted) { setData(buildSampleData()); }
      } catch (e) { if (mounted) setData(buildSampleData()); }
      if (mounted) setLoaded(true);
    })();
    return () => { mounted = false; };
  }, []);

  const persist = useCallback(async (next) => {
    setData(next);
    try { await window.storage.set(STORAGE_KEY, JSON.stringify(next), false); } catch (e) { }
  }, []);

  const showToast = (msg, kind = "success") => { setToast({ msg, kind }); setTimeout(() => setToast(null), 3600); };

  const requestView = (id) => {
    const navItem = NAV.find(n => n.id === id);
    if (navItem && navItem.open) { setView(id); return; }
    if (unlocked) { setView(id); return; }
    setPendingView(id); setPwError(""); setPwInput("");
  };

  // Re-locks the PPC-protected tabs and returns to Capacity Booking — the PPC equivalent of
  // logging out, now that there's no main login session to end.
  const lockPPC = () => {
    setUnlocked(false); setPendingView(null); setPwError(""); setPwInput(""); setView("book");
  };

  if (!loaded || !data) return <div style={{ minHeight: 400, display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "'Inter', sans-serif", color: COLORS.inkSoft }}>Loading {BRAND_SHORT} capacity data…</div>;

  const { factories, halls, lines, bookings, additionalProcesses, categories, seasons, customers, brands, savedEmailContacts } = data;
  const settings = { ...DEFAULT_SETTINGS, ...(data.settings || {}) };

  return (
    <>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600;700&display=swap');
        html, body { height: 100%; width: 100%; margin: 0; padding: 0; }
        #root, #app, body > div { height: 100%; width: 100%; max-width: none !important; margin: 0 !important; padding: 0 !important; }
        .print-only { display: none; }
        @media print { .app-shell { display: none !important; } .print-only { display: block !important; } }
        @media (max-width: 980px) { .book-grid { grid-template-columns: 1fr !important; } .book-grid > div:last-child { position: static !important; } }
      `}</style>
      <div className="app-shell" style={{ width: "100%", maxWidth: "none", minHeight: "100vh", fontFamily: "'Inter', sans-serif", background: COLORS.bg, color: COLORS.ink, display: "flex", border: "none", borderRadius: 0, position: "relative" }}>
        <style>{`
          * { box-sizing: border-box; }
          table { border-collapse: collapse; width: 100%; }
          th, td { text-align: left; padding: 9px 10px; font-size: 12.5px; }
          th { color: ${COLORS.inkFaint}; font-weight: 700; text-transform: uppercase; font-size: 10.5px; letter-spacing: 0.4px; background: #F7F9FC; border-bottom: 1px solid ${COLORS.line}; }
          tbody tr { border-bottom: 1px solid ${COLORS.line}; }
          tbody tr:hover { background: #F7FAFD; }
          .mono { font-family: 'IBM Plex Mono', monospace; }
          ::-webkit-scrollbar { width: 8px; height: 8px; }
          ::-webkit-scrollbar-thumb { background: ${COLORS.lineStrong}; border-radius: 4px; }
          .app-btn:hover:not(:disabled) { filter: brightness(0.97); }
          .app-btn-primary:hover:not(:disabled), .app-btn-success:hover:not(:disabled) { filter: brightness(1.08); }
          .app-btn-default:hover:not(:disabled) { background: #F7F9FC; }
          .app-btn-danger:hover:not(:disabled), .app-btn-ghost:hover:not(:disabled) { background: ${COLORS.bg}; }
          .app-input:focus { outline: none; border-color: ${COLORS.accent} !important; box-shadow: 0 0 0 3px ${COLORS.accentSoft}; }
        `}</style>

        <Sidebar view={view} setView={requestView} unlocked={unlocked} onLogout={lockPPC} />

        <div style={{ flex: "1 1 0%", minWidth: 0, width: "100%", display: "flex", flexDirection: "column" }}>
          <TopBar bookings={bookings} view={view} />
          <div style={{ flex: 1, overflowY: "auto", padding: "22px 26px", background: COLORS.bg }}>
            {view === "book" && <BookOrder factories={factories} halls={halls} lines={lines} bookings={bookings} additionalProcesses={additionalProcesses} categories={categories} seasons={seasons} customers={customers} brands={brands} savedEmailContacts={savedEmailContacts} settings={settings} data={data} persist={persist} showToast={showToast} setView={requestView} currentUser={currentUser} />}
            {view === "dashboard" && unlocked && <Dashboard factories={factories} halls={halls} lines={lines} bookings={bookings} additionalProcesses={additionalProcesses} settings={settings} setView={requestView} location={dashLocation} setLocation={setDashLocation} factoryId={dashFactoryId} setFactoryId={setDashFactoryId} />}
            {view === "factories" && unlocked && <FactorySetup factories={factories} halls={halls} lines={lines} data={data} persist={persist} showToast={showToast} />}
            {view === "lines" && unlocked && <LineSetup factories={factories} halls={halls} lines={lines} data={data} persist={persist} selectedFactoryId={selectedFactoryId} setSelectedFactoryId={setSelectedFactoryId} showToast={showToast} />}
            {view === "orders" && unlocked && <OrdersMaster factories={factories} lines={lines} bookings={bookings} filter={orderFilter} setFilter={setOrderFilter} data={data} persist={persist} showToast={showToast} />}
            {view === "loading" && unlocked && <LineLoadingBoard factories={factories} halls={halls} lines={lines} bookings={bookings} />}
            {view === "schedule" && unlocked && <ProductionSchedule factories={factories} lines={lines} bookings={bookings} />}
            {view === "risk" && unlocked && <CapacityRisk factories={factories} lines={lines} bookings={bookings} additionalProcesses={additionalProcesses} settings={settings} savedEmailContacts={savedEmailContacts} />}
            {view === "process" && unlocked && <AdditionalProcessesView additionalProcesses={additionalProcesses} bookings={bookings} data={data} persist={persist} showToast={showToast} />}
            {view === "reports" && unlocked && <Reports factories={factories} halls={halls} lines={lines} bookings={bookings} onPrint={setPrintJob} />}
            {view === "settings" && unlocked && <Settings data={data} persist={persist} showToast={showToast} />}
          </div>
        </div>

        {toast && <div style={{ position: "absolute", bottom: 22, right: 22, background: toast.kind === "danger" ? COLORS.danger : toast.kind === "success" ? COLORS.success : COLORS.accent, color: "#fff", padding: "11px 18px", borderRadius: 9, fontSize: 13, fontWeight: 600, boxShadow: "0 8px 24px rgba(27,36,48,0.22)" }}>{toast.msg}</div>}

        {pendingView && (
          <div style={{ position: "fixed", inset: 0, background: "rgba(30,41,59,0.35)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 60 }}>
            <div style={{ background: "#fff", borderRadius: 12, padding: "26px 28px", maxWidth: 380, width: "100%", boxShadow: "0 16px 48px rgba(27,36,48,0.28)" }}>
              <div style={{ width: 40, height: 40, borderRadius: 10, background: COLORS.accentSoft, display: "flex", alignItems: "center", justifyContent: "center", color: COLORS.accentInk, marginBottom: 14 }}><Icon name="lock" size={19} /></div>
              <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, marginBottom: 6 }}>PPC Password Required</div>
              <div style={{ fontSize: 12.5, color: COLORS.inkSoft, marginBottom: 16, lineHeight: 1.6 }}>This section is restricted to authorised PPC users. Please enter the PPC password.</div>
              <TextInput type="password" autoFocus placeholder="Password" value={pwInput} onChange={e => { setPwInput(e.target.value); setPwError(""); }} onKeyDown={e => {
                if (e.key === "Enter") {
                  if (pwInput === settings.ppcPassword) { setUnlocked(true); setView(pendingView); setPendingView(null); setPwError(""); setPwInput(""); }
                  else setPwError("Incorrect PPC password. Access denied.");
                }
              }} />
              {pwError && <div style={{ fontSize: 11.5, color: COLORS.danger, marginTop: 6, fontWeight: 600 }}>{pwError}</div>}
              <div style={{ display: "flex", gap: 8, marginTop: 14, justifyContent: "flex-end" }}>
                <Btn onClick={() => { setPendingView(null); setPwError(""); setPwInput(""); }}>Cancel</Btn>
                <Btn variant="primary" onClick={() => {
                  if (pwInput === settings.ppcPassword) { setUnlocked(true); setView(pendingView); setPendingView(null); setPwError(""); setPwInput(""); }
                  else setPwError("Incorrect PPC password. Access denied.");
                }}>Unlock</Btn>
              </div>
            </div>
          </div>
        )}
      </div>

      {printJob && <PrintReport job={printJob} onDone={() => setPrintJob(null)} />}
    </>
  );
}

/* ============================================================
   SIDEBAR + TOPBAR
   ============================================================ */

function Sidebar({ view, setView, unlocked, onLogout }) {
  return (
    <div style={{ width: 240, background: COLORS.sidebar, color: "#E4E4E2", flexShrink: 0, display: "flex", flexDirection: "column" }}>
      <style>{`
        .nav-item { display: flex; align-items: center; justify-content: space-between; gap: 10px; width: 100%; padding: 10px 12px; margin-bottom: 3px; border: none; border-radius: 8px; cursor: pointer; font-size: 12.8px; font-weight: 500; text-align: left; font-family: 'Inter', sans-serif; transition: background 0.15s; }
        .nav-item:hover { background: rgba(255,255,255,0.10); }
        .nav-item-active { background: rgba(255,255,255,0.18) !important; color: #fff !important; font-weight: 600; }
        .nav-item-primary { background: rgba(255,255,255,0.94); color: ${COLORS.accentInk}; font-weight: 700; box-shadow: 0 1px 3px rgba(0,0,0,0.2); }
        .nav-item-primary:hover { background: #fff; }
      `}</style>
      <div style={{ padding: "20px 20px 16px", borderBottom: "1px solid rgba(255,255,255,0.15)" }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 15.5, color: "#fff", letterSpacing: 0.2, lineHeight: 1.3 }}>MB & VB Group</div>
        <div style={{ fontSize: 11, color: "#C6D2E0", marginTop: 2, letterSpacing: 0.3 }}>Capacity Planner</div>
      </div>
      <div style={{ padding: "14px 12px 6px" }}>
        <button onClick={() => setView("book")} className={`nav-item nav-item-primary`}>
          <span style={{ display: "flex", alignItems: "center", gap: 10 }}><Icon name="plus" size={16} />Capacity Booking</span>
        </button>
      </div>
      <div style={{ flex: 1, overflowY: "auto", padding: "6px 12px 12px" }}>
        {NAV.filter(n => n.id !== "book").map(n => (
          <button key={n.id} onClick={() => setView(n.id)} className={`nav-item ${view === n.id ? "nav-item-active" : ""}`} style={{ background: view === n.id ? undefined : "transparent", color: view === n.id ? undefined : "#CDD8E6" }}>
            <span style={{ display: "flex", alignItems: "center", gap: 10 }}><Icon name={n.icon} size={15} />{n.label}</span>
            {!n.open && !unlocked && <Icon name="lock" size={12} />}
          </button>
        ))}
      </div>
      <button onClick={onLogout} style={{ margin: "0 12px 12px", padding: "9px 12px", background: "rgba(255,255,255,0.08)", color: "#fff", border: "none", borderRadius: 8, cursor: "pointer", fontSize: 12.5, fontWeight: 600, textAlign: "left", fontFamily: "'Inter', sans-serif" }}>Lock PPC sections</button>
      <div style={{ padding: "12px 20px", fontSize: 10.5, color: "#AEBBCC", borderTop: "1px solid rgba(255,255,255,0.15)" }}>Ismailia · Port Said · Suez · Alexanderia</div>
    </div>
  );
}
function TopBar({ bookings, view }) {
  const label = (NAV.find(n => n.id === view) || {}).label || "";
  const atRisk = bookings.filter(b => b.pcdStatus === "AT RISK" || b.pcdStatus === "NOT ACHIEVABLE" || b.deliveryStatus === "AT RISK" || b.deliveryStatus === "NOT ACHIEVABLE" || b.status === "RAW MATERIAL HOLD").length;
  return (
    <div style={{ padding: "16px 28px", background: COLORS.panel, borderBottom: `1px solid ${COLORS.line}`, boxShadow: "0 1px 3px rgba(27,36,48,0.03)", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
      <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", letterSpacing: 0.6, fontWeight: 700, marginBottom: 3 }}>{BRAND_NAME}</div><h2 style={{ margin: 0, fontFamily: "'Space Grotesk', sans-serif", fontSize: 18, fontWeight: 700 }}>{label}</h2></div>
      {atRisk > 0 && <div style={{ display: "flex", alignItems: "center", gap: 6, fontSize: 12, color: COLORS.danger, fontWeight: 700, background: COLORS.dangerSoft, padding: "6px 12px", borderRadius: 20 }}><Icon name="alert" size={14} /> {atRisk} order{atRisk > 1 ? "s" : ""} need attention</div>}
    </div>
  );
}

/* ============================================================
   DASHBOARD
   ============================================================ */

function Dashboard({ factories, halls, lines, bookings, additionalProcesses, settings, setView, location, setLocation, factoryId, setFactoryId }) {
  const factoriesById = Object.fromEntries(factories.map(f => [f.id, f]));
  const locFactories = location === "ALL" ? factories : factories.filter(f => f.location === location);
  const scopeFactoryIds = factoryId === "ALL" ? locFactories.map(f => f.id) : [factoryId];
  const scopedFactories = factories.filter(f => scopeFactoryIds.includes(f.id));
  const scopedLines = lines.filter(l => scopeFactoryIds.includes(l.factoryId));
  const onLocationChange = (loc) => { setLocation(loc); setFactoryId("ALL"); };

  const active = bookings.filter(b => !isBlockedStatus(b.status) && scopeFactoryIds.includes(b.factoryId));
  const today = todayStr();
  const regularLines = scopedLines.filter(l => l.active && !l.isDenim);
  const denimLines = scopedLines.filter(l => l.active && l.isDenim);
  const totalCapacityDay = regularLines.reduce((s, l) => s + getLineDailyMinutes(l, factoriesById[l.factoryId]), 0);
  const denimCapacityDay = denimLines.reduce((s, l) => s + getLineDailyMinutes(l, factoriesById[l.factoryId]), 0);
  const bookedDay = active.reduce((s, b) => (b.startDate <= today && b.completionDate >= today) ? s + (b.dailyMinutes || 0) : s, 0);
  const denimBookedDay = active.filter(b => denimLines.some(l => l.id === b.lineId)).reduce((s, b) => (b.startDate <= today && b.completionDate >= today) ? s + (b.dailyMinutes || 0) : s, 0);
  const utilization = totalCapacityDay > 0 ? (bookedDay / totalCapacityDay) * 100 : 0;

  const totalQty = active.reduce((s, b) => s + b.qty, 0);
  const totalFOB = active.reduce((s, b) => s + b.qty * (b.fob || 0), 0);
  const projectionCount = bookings.filter(b => scopeFactoryIds.includes(b.factoryId) && b.orderType === "Projection").length;
  const confirmedCount = bookings.filter(b => scopeFactoryIds.includes(b.factoryId) && b.orderType === "Confirmed Order").length;
  const within60 = bookings.filter(b => scopeFactoryIds.includes(b.factoryId) && daysBetween(today, b.pcd) !== null && daysBetween(today, b.pcd) <= (settings.urgentThresholdDays || 30)).length;
  const awaitingRawMaterial = bookings.filter(b => scopeFactoryIds.includes(b.factoryId) && b.status === "RAW MATERIAL HOLD").length;
  const pcdRisk = active.filter(b => b.pcdStatus === "AT RISK" || b.pcdStatus === "NOT ACHIEVABLE").length;
  const deliveryRisk = active.filter(b => b.deliveryStatus === "AT RISK" || b.deliveryStatus === "NOT ACHIEVABLE").length;
  const capacityShortage = active.filter(b => b.capacityStatus !== "AVAILABLE").length;
  const overrideCount = active.filter(b => b.overrideUsed).length;
  const projectionsNeedingAction = bookings.filter(b => scopeFactoryIds.includes(b.factoryId) && !b.confirmed && projectionReminderTier(b, settings).tier !== "OK").length;

  const byFactory = scopedFactories.map(f => {
    const fLinesRegular = lines.filter(l => l.factoryId === f.id && l.active && !l.isDenim);
    const fLinesDenim = lines.filter(l => l.factoryId === f.id && l.active && l.isDenim);
    const cap = fLinesRegular.reduce((s, l) => s + getLineDailyMinutes(l, f), 0);
    const denimCap = fLinesDenim.reduce((s, l) => s + getLineDailyMinutes(l, f), 0);
    const booked = active.filter(b => b.factoryId === f.id && !fLinesDenim.some(l => l.id === b.lineId)).reduce((s, b) => (b.startDate <= today && b.completionDate >= today) ? s + (b.dailyMinutes || 0) : s, 0);
    return { factory: f, cap, booked, remaining: Math.max(0, cap - booked), util: cap > 0 ? (booked / cap) * 100 : 0, denimCap, lineCount: fLinesRegular.length, denimLineCount: fLinesDenim.length };
  });
  const byHall = halls.filter(h => scopeFactoryIds.includes(h.factoryId)).map(h => {
    const hLines = lines.filter(l => l.hallId === h.id && l.active);
    const cap = hLines.reduce((s, l) => s + getLineDailyMinutes(l, factoriesById[h.factoryId]), 0);
    const booked = hLines.reduce((s, l) => s + dailyBookedOnLine(l.id, today, bookings), 0);
    return { hall: h, factory: factoriesById[h.factoryId], cap, booked, lineCount: hLines.length, util: cap > 0 ? (booked / cap) * 100 : 0 };
  }).filter(h => h.lineCount > 0);

  const atRiskOrders = active.filter(b => b.pcdStatus === "AT RISK" || b.pcdStatus === "NOT ACHIEVABLE" || b.deliveryStatus === "AT RISK" || b.deliveryStatus === "NOT ACHIEVABLE" || b.status === "PARTIALLY ACCEPTED");
  const holdOrders = bookings.filter(b => scopeFactoryIds.includes(b.factoryId) && b.status === "RAW MATERIAL HOLD");

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 14 }}>
        <div><div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18 }}>{BRAND_NAME}</div><div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 2 }}>Consolidated capacity across every facility</div></div>
        <div style={{ display: "flex", gap: 12 }}>
          <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600, marginBottom: 4, letterSpacing: 0.4 }}>Select location</div><Select value={location} onChange={e => onLocationChange(e.target.value)} style={{ minWidth: 190 }}><option value="ALL">All Locations</option>{LOCATIONS.map(l => <option key={l} value={l}>{LOCATION_LABELS[l]}</option>)}</Select></div>
          <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600, marginBottom: 4, letterSpacing: 0.4 }}>Select factory</div><Select value={factoryId} onChange={e => setFactoryId(e.target.value)} style={{ minWidth: 190 }}><option value="ALL">All Factories</option>{locFactories.map(f => <option key={f.id} value={f.id}>{f.name}</option>)}</Select></div>
        </div>
      </div>

      {holdOrders.length > 0 && <div style={{ background: COLORS.dangerSoft, border: `1px solid ${COLORS.danger}44`, borderRadius: 6, padding: "12px 16px", display: "flex", alignItems: "center", gap: 10 }}><Icon name="alert" size={16} /><div style={{ fontSize: 12.5, color: COLORS.danger, fontWeight: 600 }}>{holdOrders.length} booking{holdOrders.length > 1 ? "s" : ""} on RAW MATERIAL HOLD — PCD within {settings.urgentThresholdDays || 30} days without confirmed raw material.</div></div>}
      {projectionsNeedingAction > 0 && <div style={{ background: COLORS.warningSoft, border: `1px solid ${COLORS.warning}44`, borderRadius: 6, padding: "12px 16px", display: "flex", alignItems: "center", gap: 10 }}><Icon name="alert" size={16} /><div style={{ fontSize: 12.5, color: COLORS.warning, fontWeight: 600 }}>{projectionsNeedingAction} projection{projectionsNeedingAction > 1 ? "s" : ""} approaching the 90-day confirmation deadline — see Capacity Risk & Reminders.</div></div>}

      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        <Metric label="Total available capacity" value={fmtNum(totalCapacityDay)} sub="minutes/day" />
        <Metric label="Booked capacity" value={fmtNum(bookedDay)} sub="minutes/day" />
        <Metric label="Remaining capacity" value={fmtNum(Math.max(0, totalCapacityDay - bookedDay))} sub="minutes/day" tone="success" />
        <Metric label="Capacity utilization" value={`${utilization.toFixed(1)}%`} tone={utilization >= 95 ? "danger" : utilization >= 80 ? "warning" : "success"} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        <Metric label="Total styles booked" value={active.length} />
        <Metric label="Projection orders" value={projectionCount} tone="warning" />
        <Metric label="Confirmed orders" value={confirmedCount} tone="success" />
        <Metric label="Total FOB value" value={fmtNum(totalFOB)} sub="Σ qty × FOB" />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        <Metric label={`PCD within ${settings.urgentThresholdDays || 30} days`} value={within60} tone={within60 > 0 ? "warning" : "success"} />
        <Metric label="Awaiting raw material" value={awaitingRawMaterial} tone={awaitingRawMaterial > 0 ? "danger" : "success"} />
        <Metric label="Capacity overload alerts" value={capacityShortage} tone={capacityShortage > 0 ? "danger" : "success"} />
        <Metric label="Projections needing action" value={projectionsNeedingAction} tone={projectionsNeedingAction > 0 ? "warning" : "success"} />
      </div>

      {denimCapacityDay > 0 && (
        <Card title="Denim capacity (Embee 11 Hall 4 — 4 dedicated lines)">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
            <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600 }}>Total</div><div className="mono" style={{ fontSize: 17, fontWeight: 600 }}>{fmtNum(denimCapacityDay)} min</div></div>
            <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600 }}>Booked</div><div className="mono" style={{ fontSize: 17, fontWeight: 600 }}>{fmtNum(denimBookedDay)} min</div></div>
            <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600 }}>Remaining</div><div className="mono" style={{ fontSize: 17, fontWeight: 600, color: COLORS.success }}>{fmtNum(Math.max(0, denimCapacityDay - denimBookedDay))} min</div></div>
            <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600 }}>Utilization</div><div className="mono" style={{ fontSize: 17, fontWeight: 600 }}>{denimCapacityDay ? ((denimBookedDay / denimCapacityDay) * 100).toFixed(0) : 0}%</div></div>
          </div>
        </Card>
      )}

      <Card title="Factory-wise capacity">
        {byFactory.length === 0 ? <Empty text="No factories in this selection." /> : (
          <table><thead><tr><th>Factory</th><th>Location</th><th>Lines</th><th>Total capacity</th><th>Booked</th><th>Remaining</th><th>Utilization</th></tr></thead>
            <tbody>{byFactory.map(({ factory, cap, booked, remaining, util, denimLineCount, lineCount }) => (
              <tr key={factory.id}><td style={{ fontWeight: 600 }}>{factory.name}</td><td style={{ color: COLORS.inkSoft }}>{factory.location}</td>
                <td className="mono">{lineCount}{denimLineCount > 0 && <span style={{ color: COLORS.denim }}> (+{denimLineCount} denim)</span>}</td>
                <td className="mono">{fmtNum(cap)}</td><td className="mono">{fmtNum(booked)}</td><td className="mono" style={{ color: COLORS.success }}>{fmtNum(remaining)}</td>
                <td style={{ width: 130 }}><UtilBar pct={util} /><span style={{ fontSize: 11, color: COLORS.inkFaint }}>{util.toFixed(0)}%</span></td></tr>
            ))}</tbody>
          </table>
        )}
      </Card>

      <Card title="Hall-wise capacity">
        {byHall.length === 0 ? <Empty text="No halls in this selection." /> : (
          <table><thead><tr><th>Factory</th><th>Hall</th><th>Lines</th><th>Total capacity</th><th>Booked</th><th>Utilization</th></tr></thead>
            <tbody>{byHall.map(({ hall, factory, cap, booked, lineCount, util }) => (
              <tr key={hall.id}><td>{factory?.name}</td><td style={{ fontWeight: 600 }}>{hall.name}</td><td className="mono">{lineCount}</td><td className="mono">{fmtNum(cap)}</td><td className="mono">{fmtNum(booked)}</td><td style={{ width: 130 }}><UtilBar pct={util} /></td></tr>
            ))}</tbody>
          </table>
        )}
      </Card>

      <Card title="Orders at risk — bottleneck analysis" action={<Btn variant="ghost" onClick={() => setView("risk")}>Open full risk dashboard →</Btn>}>
        {atRiskOrders.length === 0 && holdOrders.length === 0 ? <Empty text="No orders currently at risk in this selection." /> : (
          <table><thead><tr><th>Item</th><th>Style</th><th>Customer</th><th>PCD</th><th>Delivery</th><th>Bottleneck</th></tr></thead>
            <tbody>
              {holdOrders.map(b => <tr key={b.id}><td className="mono">{b.itemNumber}</td><td>{b.styleNumber}</td><td>{b.customer}</td><td><StatusPill status="RAW MATERIAL HOLD" /></td><td>—</td><td style={{ color: COLORS.danger, fontSize: 12 }}>{b.bottleneck}</td></tr>)}
              {atRiskOrders.map(b => <tr key={b.id}><td className="mono">{b.itemNumber}</td><td>{b.styleNumber}</td><td>{b.customer}</td><td><StatusPill status={b.pcdStatus} /></td><td><StatusPill status={b.deliveryStatus} /></td><td style={{ color: COLORS.inkSoft, fontSize: 12 }}>{b.bottleneck || "—"}</td></tr>)}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}
function Metric({ label, value, sub, tone }) {
  const color = tone === "danger" ? COLORS.danger : tone === "warning" ? COLORS.warning : tone === "success" ? COLORS.success : COLORS.ink;
  return (
    <div style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderRadius: 6, padding: "12px 14px" }}>
      <div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", letterSpacing: 0.4, fontWeight: 600, marginBottom: 6 }}>{label}</div>
      <div className="mono" style={{ fontSize: 21, fontWeight: 600, color }}>{value}</div>
      {sub && <div style={{ fontSize: 11, color: COLORS.inkFaint, marginTop: 2 }}>{sub}</div>}
    </div>
  );
}
function Empty({ text }) { return <div style={{ padding: "24px 0", textAlign: "center", color: COLORS.inkFaint, fontSize: 13 }}>{text}</div>; }

/* ============================================================
   FACTORY SETUP  (incl. Hall management)
   ============================================================ */

function emptyFactory() { return { id: "F" + Date.now(), name: "", code: "", location: LOCATIONS[0], workingDaysPerWeek: 6, workingHoursPerDay: 7, shiftPattern: "Single shift (7h)", defaultEfficiency: 70, active: true }; }

function FactorySetup({ factories, halls, lines, data, persist, showToast }) {
  const [editing, setEditing] = useState(null);
  const [newHallName, setNewHallName] = useState({});

  const save = (factory) => {
    const exists = factories.some(f => f.id === factory.id);
    let nextFactories = exists ? factories.map(f => f.id === factory.id ? factory : f) : [...factories, factory];
    let nextHalls = halls;
    if (!exists && !halls.some(h => h.factoryId === factory.id)) nextHalls = [...halls, { id: `${factory.id}-H1`, factoryId: factory.id, name: "Hall 1" }];
    persist({ ...data, factories: nextFactories, halls: nextHalls });
    setEditing(null); showToast(exists ? "Factory updated" : "Factory created");
  };
  const addHall = (factoryId) => {
    const name = (newHallName[factoryId] || "").trim();
    if (!name) return;
    persist({ ...data, halls: [...halls, { id: `${factoryId}-H${Date.now()}`, factoryId, name }] });
    setNewHallName({ ...newHallName, [factoryId]: "" });
  };
  const groups = LOCATIONS.map(loc => ({ loc, list: factories.filter(f => f.location === loc) })).filter(g => g.list.length > 0);

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 22 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16 }}>{BRAND_NAME}</div>
        <Btn variant="primary" onClick={() => setEditing(emptyFactory())}>+ New factory</Btn>
      </div>
      {editing && <FactoryForm factory={editing} onCancel={() => setEditing(null)} onSave={save} />}
      {groups.map(({ loc, list }) => (
        <div key={loc}>
          <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, fontSize: 14, marginBottom: 10, color: COLORS.accentInk }}>{LOCATION_LABELS[loc]}</div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14 }}>
            {list.map(f => {
              const fLines = lines.filter(l => l.factoryId === f.id);
              const regular = fLines.filter(l => l.active && !l.isDenim), denim = fLines.filter(l => l.active && l.isDenim);
              const cap = regular.reduce((s, l) => s + getLineDailyMinutes(l, f), 0);
              const denimCap = denim.reduce((s, l) => s + getLineDailyMinutes(l, f), 0);
              const fHalls = halls.filter(h => h.factoryId === f.id);
              return (
                <Card key={f.id}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                    <div><div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, fontSize: 15 }}>{f.name}</div><div className="mono" style={{ fontSize: 11.5, color: COLORS.inkFaint }}>{f.code} · {f.location}</div></div>
                    <StatusPill status={f.active ? "AVAILABLE" : "N/A"} />
                  </div>
                  <div style={{ marginTop: 12, display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, fontSize: 12 }}>
                    <div><span style={{ color: COLORS.inkFaint }}>Lines</span><div className="mono" style={{ fontWeight: 600 }}>{regular.length}{denim.length > 0 && <span style={{ color: COLORS.denim }}> +{denim.length}D</span>}</div></div>
                    <div><span style={{ color: COLORS.inkFaint }}>Working hours/day</span><div className="mono" style={{ fontWeight: 600 }}>{f.workingHoursPerDay}h</div></div>
                  </div>
                  <div style={{ marginTop: 10 }}>
                    <div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600, marginBottom: 4 }}>Halls</div>
                    <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{fHalls.map(h => <span key={h.id} style={{ fontSize: 11, padding: "3px 8px", background: COLORS.accentSoft, color: COLORS.accentInk, borderRadius: 3, fontWeight: 600 }}>{h.name}</span>)}</div>
                    <div style={{ display: "flex", gap: 6, marginTop: 6 }}>
                      <TextInput placeholder="New hall name" value={newHallName[f.id] || ""} onChange={e => setNewHallName({ ...newHallName, [f.id]: e.target.value })} style={{ fontSize: 11.5, padding: "5px 8px" }} />
                      <Btn variant="ghost" onClick={() => addHall(f.id)}>+ Add</Btn>
                    </div>
                  </div>
                  <div style={{ marginTop: 12, paddingTop: 10, borderTop: `1px solid ${COLORS.line}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                    <div><div className="mono" style={{ fontSize: 13, fontWeight: 600, color: COLORS.accentInk }}>{fmtNum(cap)} min/day</div>{denimCap > 0 && <div className="mono" style={{ fontSize: 11, color: COLORS.denim }}>+{fmtNum(denimCap)} denim</div>}</div>
                    <Btn variant="ghost" onClick={() => setEditing(f)}>Edit</Btn>
                  </div>
                </Card>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}
function FactoryForm({ factory, onCancel, onSave }) {
  const [f, setF] = useState(factory);
  const set = (k, v) => setF({ ...f, [k]: v });
  return (
    <Card title={factory.name ? "Edit factory" : "New factory"}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
        <Field label="Factory name"><TextInput value={f.name} onChange={e => set("name", e.target.value)} placeholder="e.g. Embee 1" /></Field>
        <Field label="Factory code"><TextInput value={f.code} onChange={e => set("code", e.target.value)} placeholder="EMB1" /></Field>
        <Field label="Location"><Select value={f.location} onChange={e => set("location", e.target.value)}>{LOCATIONS.map(l => <option key={l} value={l}>{l}</option>)}</Select></Field>
        <Field label="Shift pattern"><TextInput value={f.shiftPattern} onChange={e => set("shiftPattern", e.target.value)} /></Field>
        <Field label="Working days/week"><TextInput type="number" value={f.workingDaysPerWeek} onChange={e => set("workingDaysPerWeek", Number(e.target.value))} /></Field>
        <Field label="Working hours/day" hint="Drives available minutes/day = hours × 60 per line."><TextInput type="number" value={f.workingHoursPerDay} onChange={e => set("workingHoursPerDay", Number(e.target.value))} /></Field>
        <Field label="Default efficiency % (informational)"><TextInput type="number" value={f.defaultEfficiency} onChange={e => set("defaultEfficiency", Number(e.target.value))} /></Field>
        <Field label="Status"><Select value={f.active ? "1" : "0"} onChange={e => set("active", e.target.value === "1")}><option value="1">Active</option><option value="0">Inactive</option></Select></Field>
      </div>
      <div style={{ display: "flex", gap: 8, marginTop: 6 }}><Btn variant="primary" onClick={() => f.name && onSave(f)} disabled={!f.name}>Save factory</Btn><Btn onClick={onCancel}>Cancel</Btn></div>
    </Card>
  );
}

/* ============================================================
   LINE SETUP
   ============================================================ */

function emptyLine(factoryId, hallId, nextNum) {
  return { id: "L" + Date.now(), factoryId, hallId, lineNumber: nextNum, lineName: `Line ${String(nextNum).padStart(2, "0")}`, lineType: "Sewing", operators: 30, helpers: 6, workingHoursPerDay: 7, efficiency: 70, overrideMinutes: null, isDenim: false, plannedPcsPerDay: null, machines: makeMachines(8, 4, 8, 3, 2, 1, 1, 1), active: true };
}

function LineSetup({ factories, halls, lines, data, persist, selectedFactoryId, setSelectedFactoryId, showToast }) {
  const fid = selectedFactoryId || factories[0]?.id;
  const [editing, setEditing] = useState(null);
  const [hallFilter, setHallFilter] = useState("ALL");
  const [filterText, setFilterText] = useState("");
  const factory = factories.find(f => f.id === fid);
  const factoryHalls = halls.filter(h => h.factoryId === fid);
  const factoryLines = lines.filter(l => l.factoryId === fid && (hallFilter === "ALL" || l.hallId === hallFilter) && l.lineName.toLowerCase().includes(filterText.toLowerCase()));

  const save = (line) => {
    const exists = lines.some(l => l.id === line.id);
    persist({ ...data, lines: exists ? lines.map(l => l.id === line.id ? line : l) : [...lines, line] });
    setEditing(null); showToast(exists ? "Line updated" : "Line created");
  };
  const updatePlannedPcs = (lineId, value) => persist({ ...data, lines: lines.map(l => l.id === lineId ? { ...l, plannedPcsPerDay: value === "" ? null : Number(value) } : l) });
  const totalFactoryLines = lines.filter(l => l.factoryId === fid);
  const totalRegular = totalFactoryLines.filter(l => !l.isDenim).length;
  const totalDenim = totalFactoryLines.filter(l => l.isDenim).length;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-end", flexWrap: "wrap", gap: 10 }}>
        <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600, marginBottom: 4, letterSpacing: 0.4 }}>Factory</div>
          <Select value={fid} onChange={e => { setSelectedFactoryId(e.target.value); setFilterText(""); setHallFilter("ALL"); }} style={{ minWidth: 220 }}>
            {LOCATIONS.map(loc => { const list = factories.filter(f => f.location === loc); if (list.length === 0) return null; return <optgroup key={loc} label={LOCATION_LABELS[loc]}>{list.map(f => <option key={f.id} value={f.id}>{f.name}</option>)}</optgroup>; })}
          </Select>
        </div>
        <div><div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600, marginBottom: 4, letterSpacing: 0.4 }}>Hall</div>
          <Select value={hallFilter} onChange={e => setHallFilter(e.target.value)} style={{ minWidth: 160 }}><option value="ALL">All halls</option>{factoryHalls.map(h => <option key={h.id} value={h.id}>{h.name}</option>)}</Select>
        </div>
        <TextInput placeholder="Filter lines…" value={filterText} onChange={e => setFilterText(e.target.value)} style={{ maxWidth: 180 }} />
        <div style={{ background: COLORS.accentSoft, borderRadius: 5, padding: "8px 14px" }}>
          <div style={{ fontSize: 10, color: COLORS.accentInk, textTransform: "uppercase", fontWeight: 600, letterSpacing: 0.4 }}>Total lines — {factory?.name}</div>
          <div className="mono" style={{ fontSize: 17, fontWeight: 700, color: COLORS.accentInk }}>{totalRegular}{totalDenim > 0 && <span style={{ color: COLORS.denim }}> +{totalDenim} Denim</span>}</div>
        </div>
        <Btn variant="primary" onClick={() => setEditing(emptyLine(fid, factoryHalls[0]?.id, lines.filter(l => l.factoryId === fid && !l.isDenim).length + 1))}>+ New line</Btn>
      </div>
      <div style={{ fontSize: 12, color: COLORS.inkFaint }}>Lines are general-purpose — any style can be planned on any line, subject to capacity and dates. Only lines tagged <DenimTag /> are reserved for Denim-category orders (max 4, Embee 11 Hall 4).</div>
      {editing && <LineForm line={editing} factory={factory} halls={factoryHalls} onCancel={() => setEditing(null)} onSave={save} />}
      <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
        {factoryLines.map(l => {
          const cap = getLineDailyMinutes(l, factory);
          const hall = halls.find(h => h.id === l.hallId);
          return (
            <Card key={l.id}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 16 }}>
                <div style={{ flex: 1 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, fontSize: 14.5 }}>{l.lineName}</span>
                    {l.isDenim && <DenimTag />}<StatusPill status={l.active ? "AVAILABLE" : "N/A"} />
                    {l.overrideMinutes && <span style={{ fontSize: 10.5, color: COLORS.accentInk, fontWeight: 600 }}>OVERRIDE</span>}
                    <span style={{ fontSize: 11, color: COLORS.inkFaint }}>{hall?.name}</span>
                  </div>
                  <div style={{ fontSize: 12, color: COLORS.inkSoft, marginTop: 6 }}>{l.operators} operators · {l.helpers} helpers · {l.workingHoursPerDay}h/day</div>
                  <div style={{ marginTop: 8, display: "flex", alignItems: "center", gap: 8 }}><span style={{ fontSize: 11, color: COLORS.inkFaint }}>IE planned pcs/day:</span><TextInput type="number" value={l.plannedPcsPerDay ?? ""} onChange={e => updatePlannedPcs(l.id, e.target.value)} placeholder="not set" style={{ width: 100, padding: "3px 6px", fontSize: 11.5 }} /></div>
                  {l.currentBooking && (
                    <div style={{ marginTop: 8, background: "#FAFAF9", border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "6px 10px", fontSize: 11.5 }}>
                      <span style={{ color: COLORS.inkFaint }}>Currently:</span> {l.currentBooking.customerBrand}
                      {l.currentBooking.beginDate && <span> · <span style={{ color: COLORS.inkFaint }}>Run:</span> {niceDate(l.currentBooking.beginDate)} → {niceDate(l.currentBooking.endDate)} ({l.currentBooking.gapDays}d gap)</span>}
                      {l.currentBooking.bookedTill && <span> · <span style={{ color: COLORS.inkFaint }}>Booked till:</span> <span className="mono" style={{ fontWeight: 600 }}>{niceDate(l.currentBooking.bookedTill)}</span></span>}
                    </div>
                  )}
                </div>
                <div style={{ textAlign: "right", flexShrink: 0 }}>
                  <div style={{ fontSize: 10.5, color: COLORS.inkFaint, textTransform: "uppercase", fontWeight: 600 }}>Available / day</div>
                  <div className="mono" style={{ fontSize: 19, fontWeight: 600, color: l.isDenim ? COLORS.denim : COLORS.accentInk }}>{fmtNum(cap)}</div>
                  <div style={{ fontSize: 10.5, color: COLORS.inkFaint }}>minutes</div>
                  <Btn variant="ghost" onClick={() => setEditing(l)} style={{ marginTop: 8 }}>Edit</Btn>
                </div>
              </div>
            </Card>
          );
        })}
        {factoryLines.length === 0 && <Empty text="No lines match this filter." />}
      </div>
    </div>
  );
}
function LineForm({ line, factory, halls, onCancel, onSave }) {
  const [l, setL] = useState(JSON.parse(JSON.stringify(line)));
  const set = (k, v) => setL({ ...l, [k]: v });
  const defaultMinutes = (factory?.workingHoursPerDay || 7) * 60;
  return (
    <Card title={`${line.lineName || "New line"} — ${factory?.name || ""}`}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
        <Field label="Line number"><TextInput type="number" value={l.lineNumber} onChange={e => set("lineNumber", Number(e.target.value))} /></Field>
        <Field label="Line name"><TextInput value={l.lineName} onChange={e => set("lineName", e.target.value)} /></Field>
        <Field label="Hall"><Select value={l.hallId || ""} onChange={e => set("hallId", e.target.value)}>{halls.map(h => <option key={h.id} value={h.id}>{h.name}</option>)}</Select></Field>
        <Field label="Dedicated Denim line"><Select value={l.isDenim ? "1" : "0"} onChange={e => set("isDenim", e.target.value === "1")}><option value="0">No — general line</option><option value="1">Yes — Denim only</option></Select></Field>
        <Field label="Status"><Select value={l.active ? "1" : "0"} onChange={e => set("active", e.target.value === "1")}><option value="1">Active</option><option value="0">Inactive</option></Select></Field>
        <Field label="Operators"><TextInput type="number" value={l.operators} onChange={e => set("operators", Number(e.target.value))} /></Field>
        <Field label="Helpers"><TextInput type="number" value={l.helpers} onChange={e => set("helpers", Number(e.target.value))} /></Field>
        <Field label="IE planned pcs/day"><TextInput type="number" value={l.plannedPcsPerDay ?? ""} onChange={e => set("plannedPcsPerDay", e.target.value ? Number(e.target.value) : null)} /></Field>
      </div>
      <div style={{ background: l.isDenim ? COLORS.denimSoft : COLORS.accentSoft, borderRadius: 5, padding: "12px 14px", margin: "6px 0 16px" }}>
        <div style={{ fontSize: 12.5, color: l.isDenim ? COLORS.denim : COLORS.accentInk }}>Formula default (factory working hours × 60): <span className="mono" style={{ fontWeight: 700 }}>{fmtNum(defaultMinutes)}</span> minutes/day</div>
      </div>
      <Field label="Manual override (minutes/day) — leave blank to use the formula default above"><TextInput type="number" value={l.overrideMinutes || ""} onChange={e => set("overrideMinutes", e.target.value ? Number(e.target.value) : null)} placeholder={String(defaultMinutes)} /></Field>
      <div style={{ display: "flex", gap: 8, marginTop: 8 }}><Btn variant="primary" onClick={() => onSave(l)}>Save line</Btn><Btn onClick={onCancel}>Cancel</Btn></div>
    </Card>
  );
}

/* ============================================================
   CAPACITY BOOKING
   ============================================================ */

function emptyOrder(categories, seasons) {
  const today = todayStr();
  return {
    orderType: "Confirmed Order",
    customer: "", brand: "", styleNumber: "", description: "", category: categories[0], gender: "Unisex", season: seasons[0],
    mmHead: "", infoAlertEmails: [],
    qty: 1000, sam: 12, budgetedEfficiency: 100, fob: "", priority: "Normal", bookingDate: today,
    pcd: fmt(addDays(new Date(), 90)), deliveryDate: fmt(addDays(new Date(), 121)), possibleOCD: "",
    rawMaterialAvailability: "",
    factoryId: "", hallId: "", lineId: "", additionalProcesses: [], additionalProcessRemark: "",
    embroideryStitches: "", embroideryCapacityConfirmed: "Pending", washType: "", washTypeOther: "",
    overrideAuthorized: false, remarks: ""
  };
}

function BookOrder({ factories, halls, lines, bookings, additionalProcesses, categories, seasons, customers, brands, savedEmailContacts, settings, data, persist, showToast, setView, currentUser }) {
  const [order, setOrder] = useState(() => ({ ...emptyOrder(categories, seasons), factoryId: factories[0]?.id || "" }));
  const [result, setResult] = useState(null);
  const [rankedLines, setRankedLines] = useState([]);
  const [checkedLineId, setCheckedLineId] = useState(null);
  const [confirmedBooking, setConfirmedBooking] = useState(null);
  const [validationMsg, setValidationMsg] = useState("");
  const [confirmingReset, setConfirmingReset] = useState(false);

  const set = (k, v) => { setOrder({ ...order, [k]: v }); setResult(null); setValidationMsg(""); };
  const toggleProcess = (name) => {
    const has = order.additionalProcesses.includes(name);
    const next = has ? order.additionalProcesses.filter(p => p !== name) : [...order.additionalProcesses, name];
    set("additionalProcesses", next);
  };

  const factory = factories.find(f => f.id === order.factoryId);
  const factoryHalls = halls.filter(h => h.factoryId === order.factoryId);
  const factoryLines = lines.filter(l => l.factoryId === order.factoryId && l.active && (!order.hallId || l.hallId === order.hallId) && (!l.isDenim || order.category === "Denim"));

  const referenceLabel = order.orderType === "Projection" ? "Projection Number" : "Item Number";
  const referencePrefix = order.orderType === "Projection" ? "P" : "C";
  const pcdInfo = pcdRules(order, settings);
  const rawGate = rawMaterialGate(order, settings);
  const leadGate = leadTimeGate(order, settings);
  const projAdvance = order.orderType === "Projection" && pcdInfo.daysToP !== null
    ? (pcdInfo.daysToP > (settings.projectionAdvanceMaxDays || 180) ? "far" : pcdInfo.daysToP >= (settings.projectionAdvanceMinDays || 120) ? "ok" : "soon")
    : null;

  const runCheck = (lineId) => {
    const line = lines.find(l => l.id === lineId);
    if (!line || !factory) return;
    const r = checkAndPlan({ order, line, factory, bookings, additionalProcesses, settings });
    setResult(r); setCheckedLineId(lineId);
  };
  const runRecommend = () => {
    if (!factory) return;
    const scored = factoryLines.map(line => scoreLine({ order, line, factory, bookings, additionalProcesses, settings })).filter(s => s.score > 0 || factoryLines.length === 1).sort((a, b) => b.score - a.score);
    setRankedLines(scored);
    if (scored[0]) runCheck(scored[0].line.id);
  };

  const saveInformation = () => {
    const missing = missingMandatoryFields(order);
    if (missing.length > 0) { setValidationMsg(`Please complete all mandatory fields marked with * before saving the booking. Missing: ${missing.join(", ")}.`); return; }
    setValidationMsg("");
    order.lineId ? runCheck(order.lineId) : runRecommend();
  };

  const doConfirm = (qtyToBook, overrideAuthorized) => {
    if (!checkedLineId || !result) return;
    const missing = missingMandatoryFields(order);
    if (missing.length > 0) { showToast(`Missing mandatory fields: ${missing.join(", ")}.`, "danger"); return; }
    const freshRaw = rawMaterialGate(order, settings);
    if (freshRaw.blocked) { showToast(freshRaw.message, "danger"); return; }
    const freshLead = leadTimeGate(order, settings);
    if (freshLead.blocked) { showToast(freshLead.message, "danger"); return; }
    if (result.capacityStatus === "INSUFFICIENT" && !overrideAuthorized) { showToast("Capacity not available — authorize override or adjust the order.", "danger"); return; }

    const line = lines.find(l => l.id === checkedLineId);
    const counters = data.counters || DEFAULT_COUNTERS;
    const { code: itemNumber, nextCounters } = issueDocNumber(order.orderType, counters);
    const eff = Number(order.budgetedEfficiency) > 0 ? Number(order.budgetedEfficiency) : 100;
    const requiredMinutes = Math.round((qtyToBook * order.sam) / (eff / 100));
    const baseline = order.bookingDate ? maxDateStr(order.bookingDate, todayStr()) : todayStr();
    const estimatedStart = findEarliestStart({ line, factory, bookings, fromDate: parseDate(baseline) });
    const startDate = parseDate(estimatedStart), pcdDate = parseDate(order.pcd), deliveryDate = parseDate(order.deliveryDate);
    const capBefore = Math.max(0, getLineDailyMinutes(line, factory) - dailyBookedOnLine(line.id, estimatedStart, bookings));
    const sim = simulateLine({ line, factory, bookings, requiredMinutes, startDate, pcdDate, deliveryDate });
    const wdUsed = Math.max(1, countWorkingDays(estimatedStart, sim.completionDate || order.pcd, factory.workingDaysPerWeek));
    const dailyMinutes = Math.round(requiredMinutes / wdUsed);
    const capAfter = Math.max(0, capBefore - dailyMinutes);
    const remainingQty = order.qty - qtyToBook;

    const booking = {
      id: "B" + Date.now(), itemNumber, referenceNumber: itemNumber, orderType: order.orderType,
      customer: order.customer, brand: order.brand, styleNumber: order.styleNumber, description: order.description,
      category: order.category, gender: order.gender, season: order.season, mmHead: order.mmHead, infoAlertEmails: order.infoAlertEmails,
      qty: qtyToBook, sam: order.sam, budgetedEfficiency: eff, fob: Number(order.fob) || 0, priority: order.priority,
      confirmed: order.orderType === "Confirmed Order", bookingDate: order.bookingDate,
      pcd: order.pcd, deliveryDate: order.deliveryDate, possibleOCD: order.orderType === "Projection" ? (order.possibleOCD || null) : null, startDate: estimatedStart, completionDate: sim.completionDate || order.pcd,
      factoryId: order.factoryId, hallId: order.hallId || factoryLines[0]?.hallId, lineId: checkedLineId,
      additionalProcesses: order.additionalProcesses, additionalProcessRemark: order.additionalProcesses.includes("Other") ? order.additionalProcessRemark : null,
      embroideryStitches: order.additionalProcesses.includes("Embroidery") ? Number(order.embroideryStitches) || null : null,
      embroideryCapacityConfirmed: order.additionalProcesses.includes("Embroidery") ? order.embroideryCapacityConfirmed : null,
      washType: order.additionalProcesses.includes("Washing") ? (order.washType === "Special / Other" ? order.washTypeOther : order.washType) : null,
      rawMaterialAvailability: order.rawMaterialAvailability,
      requiredMinutes, dailyMinutes, status: overrideAuthorized ? "ACCEPTED (OVERRIDE)" : (remainingQty > 0 ? "PARTIALLY ACCEPTED" : result.decision),
      capacityStatus: result.capacityStatus, specialProcessStatus: result.specialProcessStatus,
      pcdStatus: result.pcdStatus, deliveryStatus: result.deliveryStatus, buffer: result.buffer,
      bottleneck: result.bottlenecks.join("; ") || null, plannedQty: qtyToBook, remainingQty: Math.max(0, remainingQty),
      overrideUsed: !!overrideAuthorized, remarks: order.remarks, createdAt: todayStr(),
      createdBy: currentUser || "—", createdAtTs: new Date().toISOString(),
      emailStatus: (order.infoAlertEmails && order.infoAlertEmails.length) ? "Pending" : "N/A"
    };
    const next = { ...data, bookings: [...bookings, booking], counters: nextCounters };
    persist(next);
    showToast(`${itemNumber} booked — ${fmtNum(qtyToBook)} pcs on ${line.lineName}. Capacity ${fmtNum(capBefore)} → ${fmtNum(capAfter)} min/day.`);
    setConfirmedBooking({ booking, factory, hall: halls.find(h => h.id === booking.hallId), line, capBefore, capAfter });
    setResult(null); setRankedLines([]);
    setOrder({ ...emptyOrder(categories, seasons), factoryId: order.factoryId });
  };

  const doReset = () => {
    setOrder({ ...emptyOrder(categories, seasons), factoryId: factories[0]?.id || "" });
    setResult(null); setRankedLines([]); setCheckedLineId(null); setValidationMsg(""); setConfirmingReset(false);
    showToast("Form reset");
  };

  const hasSidePanel = rankedLines.length > 0 || !!result;

  return (
    <>
    <div className="book-grid" style={{ display: "grid", gridTemplateColumns: hasSidePanel ? "minmax(0, 1.5fr) minmax(320px, 1fr)" : "1fr", gap: 20, alignItems: "start", width: "100%" }}>
      <div style={{ display: "flex", flexDirection: "column", gap: 16, minWidth: 0 }}>
        <Card title="Order information">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
            <Field label="Order Type" required><Select value={order.orderType} onChange={e => set("orderType", e.target.value)}>{ORDER_TYPES.map(t => <option key={t} value={t}>{t}</option>)}</Select></Field>
            <Field label={referenceLabel}>
              <div style={{ padding: "8px 10px", borderRadius: 4, border: `1px dashed ${COLORS.lineStrong}`, background: "#FAFAF9", fontSize: 12.5, color: COLORS.inkSoft }}>
                Auto-generated as <span className="mono" style={{ fontWeight: 700, color: COLORS.accentInk }}>{referencePrefix}###</span> on save
              </div>
            </Field>
            <Field label="Customer" required><Select value={order.customer} onChange={e => set("customer", e.target.value)}><option value="">Select…</option>{customers.map(c => <option key={c} value={c}>{c}</option>)}</Select></Field>
            <Field label="Brand" required><Select value={order.brand} onChange={e => set("brand", e.target.value)}><option value="">Select…</option>{brands.map(b => <option key={b} value={b}>{b}</option>)}</Select></Field>
            <Field label="Style number" required><TextInput value={order.styleNumber} onChange={e => set("styleNumber", e.target.value)} /></Field>
            <Field label="Season"><Select value={order.season} onChange={e => set("season", e.target.value)}>{seasons.map(s => <option key={s} value={s}>{s}</option>)}</Select></Field>
            <Field label="Marketing & Merchandising Head" required>
              <Select value={order.mmHead} onChange={e => set("mmHead", e.target.value)}><option value="">Select…</option>{MM_HEADS.map(h => <option key={h} value={h}>{h}</option>)}</Select>
            </Field>
          </div>
          <Field label="Information Alert Email ID(s)" hint="Recipients for projection/PCD/raw-material reminders on this style.">
            <EmailListEditor items={order.infoAlertEmails} onChange={v => set("infoAlertEmails", v)} placeholder="name@company.com" />
          </Field>
        </Card>

        <Card title="Product & capacity">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14 }}>
            <Field label="Product category" hint="Informational — Denim lines are the only exception."><Select value={order.category} onChange={e => set("category", e.target.value)}>{categories.map(c => <option key={c} value={c}>{c}</option>)}</Select></Field>
            <Field label="Gender"><Select value={order.gender} onChange={e => set("gender", e.target.value)}><option>Unisex</option><option>Men</option><option>Women</option><option>Kids</option></Select></Field>
            <Field label="Priority"><Select value={order.priority} onChange={e => set("priority", e.target.value)}><option>Low</option><option>Normal</option><option>High</option></Select></Field>
            <Field label="Order quantity (pcs)"><TextInput type="number" value={order.qty} onChange={e => set("qty", Number(e.target.value))} /></Field>
            <Field label="SAM (minutes)" required><TextInput type="number" step="0.01" value={order.sam} onChange={e => set("sam", Number(e.target.value))} /></Field>
            <Field label="Budgeted Efficiency %" hint="100% = no adjustment"><TextInput type="number" value={order.budgetedEfficiency} onChange={e => set("budgetedEfficiency", Number(e.target.value))} /></Field>
            <Field label="FOB" hint="Per-piece value"><TextInput type="number" step="0.01" value={order.fob} onChange={e => set("fob", e.target.value)} placeholder="0.00" /></Field>
            <Field label="Required production minutes" hint="Qty × SAM ÷ Budgeted Efficiency%">
              <div className="mono" style={{ padding: "8px 10px", fontWeight: 700 }}>{fmtNum((order.qty * order.sam) / ((Number(order.budgetedEfficiency) || 100) / 100))} min</div>
            </Field>
          </div>
          {Number(order.fob) > 0 && <div style={{ fontSize: 12, color: COLORS.inkSoft }}>Total FOB value: <span className="mono" style={{ fontWeight: 700, color: COLORS.accentInk }}>{fmtNum(order.qty * Number(order.fob))}</span></div>}
        </Card>

        <Card title="Planning dates & raw material">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14 }}>
            <Field label="Booking date"><DateField value={order.bookingDate} onChange={v => set("bookingDate", v)} /></Field>
            <Field label="PCD" required><DateField value={order.pcd} onChange={v => set("pcd", v)} /></Field>
            <Field label="Delivery / Ex-Factory date"><DateField value={order.deliveryDate} onChange={v => set("deliveryDate", v)} /></Field>
          </div>

          {order.orderType === "Projection" && (
            <Field label="Possible OCD" hint="Order Confirmation Date — anticipated date this Projection may be confirmed.">
              <DateField value={order.possibleOCD} onChange={v => set("possibleOCD", v)} />
            </Field>
          )}

          {order.orderType === "Projection" && projAdvance === "far" && (
            <div style={{ background: COLORS.accentSoft, borderRadius: 5, padding: "10px 12px", marginBottom: 10, fontSize: 12.5, color: COLORS.accentInk }}>
              This Projection's PCD is beyond the typical 120–180 day advance planning window ({pcdInfo.daysToP} days out).
            </div>
          )}
          {order.orderType === "Projection" && projAdvance === "soon" && pcdInfo.tier !== "URGENT" && (
            <div style={{ background: COLORS.warningSoft, borderRadius: 5, padding: "10px 12px", marginBottom: 10, fontSize: 12.5, color: COLORS.warning, fontWeight: 600 }}>
              This Projection has a PCD of less than 90 days. The Projection should be converted to Confirmed Order immediately.
            </div>
          )}

          {pcdInfo.tier === "URGENT" && (
            <div style={{ background: rawGate.blocked ? COLORS.dangerSoft : COLORS.successSoft, border: `1px solid ${(rawGate.blocked ? COLORS.danger : COLORS.success)}44`, borderRadius: 5, padding: "10px 12px", marginBottom: 10 }}>
              <div style={{ fontSize: 12.5, color: rawGate.blocked ? COLORS.danger : COLORS.success, fontWeight: 700, marginBottom: 4 }}>
                {rawGate.blocked ? `Urgent ${order.orderType} — Raw Material must be fully available` : `Urgent ${order.orderType} — proceeding under Raw Material exception`}
              </div>
              <div style={{ fontSize: 12, color: rawGate.blocked ? COLORS.danger : COLORS.success }}>
                {rawGate.blocked ? rawGate.message : `PCD is less than ${settings.urgentThresholdDays || 30} days from the booking date; earliest permissible standard PCD is ${niceDate(pcdInfo.minPCD)}.`}
              </div>
              <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                <Btn variant="ghost" onClick={() => set("pcd", pcdInfo.minPCD)}>Set PCD to earliest permissible ({niceDate(pcdInfo.minPCD)})</Btn>
                <Btn variant="ghost" onClick={() => set("pcd", pcdInfo.normalPCD)}>Set PCD to standard 90-day window ({niceDate(pcdInfo.normalPCD)})</Btn>
              </div>
            </div>
          )}
          {pcdInfo.tier === "SHORT" && <div style={{ background: COLORS.accentSoft, borderRadius: 5, padding: "10px 12px", marginBottom: 10, fontSize: 12.5, color: COLORS.accentInk }}>This booking is within the normal 90-day planning window. Please ensure Raw Material planning is aligned with the production schedule.</div>}

          <Field label={`Raw Material Availability${pcdInfo.tier === "URGENT" ? ` (mandatory — PCD is under ${settings.urgentThresholdDays || 30} days)` : " (optional)"}`}>
            <Select value={order.rawMaterialAvailability} onChange={e => set("rawMaterialAvailability", e.target.value)} style={pcdInfo.tier === "URGENT" && !rawGate.urgent ? { border: `1px solid ${COLORS.danger}` } : {}}>
              <option value="">— Not specified —</option>{RAW_MATERIAL_OPTIONS.map(o => <option key={o} value={o}>{o}</option>)}
            </Select>
          </Field>

          {leadGate.lead !== null && (
            <div style={{ background: leadGate.blocked ? COLORS.dangerSoft : "#FAFAF9", border: `1px solid ${leadGate.blocked ? COLORS.danger : COLORS.line}`, borderRadius: 5, padding: "10px 12px", marginTop: 4 }}>
              <div style={{ fontSize: 12, color: leadGate.blocked ? COLORS.danger : COLORS.inkSoft, fontWeight: leadGate.blocked ? 700 : 500 }}>Production lead time (PCD → Delivery): <span className="mono">{leadGate.lead} day{leadGate.lead === 1 ? "" : "s"}</span> · Required minimum: <span className="mono">{leadGate.minLead} days</span></div>
              {leadGate.blocked && <div style={{ fontSize: 12, color: COLORS.danger, marginTop: 4 }}>{leadGate.message}</div>}
            </div>
          )}
          <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 10 }}>Production Start / End Date are calculated automatically once you Save Information.</div>
        </Card>

        <Card title="Production requirements">
          <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14 }}>
            <Field label="Factory"><Select value={order.factoryId} onChange={e => { setOrder({ ...order, factoryId: e.target.value, hallId: "", lineId: "" }); setResult(null); setRankedLines([]); }}>{LOCATIONS.map(loc => { const list = factories.filter(f => f.location === loc); if (list.length === 0) return null; return <optgroup key={loc} label={LOCATION_LABELS[loc]}>{list.map(f => <option key={f.id} value={f.id}>{f.name}</option>)}</optgroup>; })}</Select></Field>
            <Field label="Hall"><Select value={order.hallId} onChange={e => set("hallId", e.target.value)}><option value="">Any hall</option>{factoryHalls.map(h => <option key={h.id} value={h.id}>{h.name}</option>)}</Select></Field>
            <Field label="Preferred line (optional)"><Select value={order.lineId} onChange={e => set("lineId", e.target.value)}><option value="">Auto-recommend</option>{factoryLines.map(l => <option key={l.id} value={l.id}>{l.lineName}{l.isDenim ? " — DENIM" : ""}</option>)}</Select></Field>
          </div>
          <Field label="Additional processes (select any that apply)"><div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>{ADDITIONAL_PROCESS_ALL.map(p => <Toggle key={p} active={order.additionalProcesses.includes(p)} onClick={() => toggleProcess(p)}>{p}</Toggle>)}</div></Field>
          {order.additionalProcesses.includes("Other") && <Field label="Remark" hint="Details about the Other additional process."><TextInput value={order.additionalProcessRemark} onChange={e => set("additionalProcessRemark", e.target.value)} placeholder="e.g. Special laser finish required on front panel" /></Field>}
          {order.additionalProcesses.includes("Embroidery") && (
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
              <Field label="No. of Stitches (Embroidery)" hint="Informational only — does not block the booking."><TextInput type="number" value={order.embroideryStitches} onChange={e => set("embroideryStitches", e.target.value)} placeholder="e.g. 25000" /></Field>
              <Field label="Embroidery Capacity Confirmed" hint="Planner-controlled — your own assessment of embroidery space."><Select value={order.embroideryCapacityConfirmed} onChange={e => set("embroideryCapacityConfirmed", e.target.value)}><option>Yes</option><option>No</option><option>Pending</option></Select></Field>
            </div>
          )}
          {order.additionalProcesses.includes("Washing") && (
            <div style={{ display: "grid", gridTemplateColumns: order.washType === "Special / Other" ? "1fr 1fr" : "1fr", gap: 14 }}>
              <Field label="Type of Wash"><Select value={order.washType} onChange={e => set("washType", e.target.value)}><option value="">Select wash type</option>{WASH_TYPES.map(w => <option key={w} value={w}>{w}</option>)}</Select></Field>
              {order.washType === "Special / Other" && <Field label="Specify wash type"><TextInput value={order.washTypeOther} onChange={e => set("washTypeOther", e.target.value)} /></Field>}
            </div>
          )}
          <Field label="Remarks"><TextInput value={order.remarks} onChange={e => set("remarks", e.target.value)} placeholder="Optional — included in the confirmation email" /></Field>
        </Card>

        {validationMsg && <div style={{ background: COLORS.dangerSoft, border: `1px solid ${COLORS.danger}55`, borderRadius: 5, padding: "10px 14px", color: COLORS.danger, fontSize: 12.5, fontWeight: 600 }}>{validationMsg}</div>}

        <div style={{ display: "flex", gap: 10 }}>
          <Btn variant="primary" onClick={saveInformation}>Save Information</Btn>
          <Btn variant="danger" onClick={() => setConfirmingReset(true)}>Reset</Btn>
        </div>
        {confirmingReset && <ConfirmDialog message="Are you sure you want to reset all entered information? This only clears the unsaved form — previously confirmed bookings are not affected." onYes={doReset} onCancel={() => setConfirmingReset(false)} />}
      </div>

      {hasSidePanel && (
        <div style={{ position: "sticky", top: 0, display: "flex", flexDirection: "column", gap: 14, minWidth: 0 }}>
          {rankedLines.length > 0 && (
            <Card title="Line recommendation">
              {rankedLines.length === 0 ? <Empty text="No eligible lines." /> : rankedLines.slice(0, 5).map(({ line, result: r, score }, i) => (
                <div key={line.id} onClick={() => runCheck(line.id)} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 10px", borderRadius: 4, marginBottom: 6, cursor: "pointer", background: checkedLineId === line.id ? COLORS.accentSoft : "#FAFAF9", border: `1px solid ${checkedLineId === line.id ? COLORS.accent : COLORS.line}` }}>
                  <div><div style={{ fontSize: 12.5, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>{i === 0 ? "★ " : ""}{line.lineName} {line.isDenim && <DenimTag />}</div><div style={{ fontSize: 11, color: COLORS.inkFaint }}>Est. start {niceDate(r.estimatedStart)}</div></div>
                  <div className="mono" style={{ fontWeight: 700, color: COLORS.accentInk }}>{score}%</div>
                </div>
              ))}
            </Card>
          )}
          {result && (
            <CheckPlanResult order={order} result={result} line={lines.find(l => l.id === checkedLineId)} referenceLabel={referenceLabel} referencePrefix={referencePrefix}
              onConfirmFull={() => doConfirm(order.qty, false)} onConfirmPartial={() => doConfirm(result.plannedQtyAtPCD, false)} onOverride={() => doConfirm(order.qty, true)} />
          )}
        </div>
      )}
    </div>
    {confirmedBooking && (
      <BookingConfirmedModal info={confirmedBooking} savedEmailContacts={savedEmailContacts} data={data} persist={persist} showToast={showToast}
        onClose={() => setConfirmedBooking(null)}
        onDone={() => { setConfirmedBooking(null); setView("orders"); }} />
    )}
    </>
  );
}

function CheckPlanResult({ order, result, line, referenceLabel, referencePrefix, onConfirmFull, onConfirmPartial, onOverride }) {
  const toneBg = result.overall === "GREEN" ? COLORS.successSoft : result.overall === "AMBER" ? COLORS.warningSoft : COLORS.dangerSoft;
  const toneFg = result.overall === "GREEN" ? COLORS.success : result.overall === "AMBER" ? COLORS.warning : COLORS.danger;
  const blocked = result.rawGate.blocked || result.leadGate.blocked;
  const heading = blocked ? "Booking cannot be accepted" : result.overall === "GREEN" ? "Successfully planned" : result.plannedQtyAtPCD > 0 ? "Partial capacity available" : "Capacity not available";
  const [overrideChecked, setOverrideChecked] = useState(false);

  return (
    <Card style={{ border: `1.5px solid ${toneFg}55` }}>
      <div style={{ background: toneBg, margin: "-18px -20px 16px", padding: "14px 20px", borderRadius: "6px 6px 0 0" }}>
        <div style={{ fontSize: 10.5, color: toneFg, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5 }}>{result.decision}</div>
        <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16, color: toneFg, marginTop: 2 }}>{heading}</div>
      </div>
      {result.rawGate.blocked && <div style={{ background: COLORS.dangerSoft, border: `1px solid ${COLORS.danger}55`, borderRadius: 5, padding: "12px 14px", marginBottom: 10, color: COLORS.danger, fontSize: 12.5, fontWeight: 600 }}>{result.rawGate.message}</div>}
      {result.leadGate.blocked && <div style={{ background: COLORS.dangerSoft, border: `1px solid ${COLORS.danger}55`, borderRadius: 5, padding: "12px 14px", marginBottom: 14, color: COLORS.danger, fontSize: 12.5, fontWeight: 600 }}>{result.leadGate.message}</div>}

      <div style={{ fontSize: 12.5, lineHeight: 1.9 }}>
        <Row label={referenceLabel} value={<span className="mono">{referencePrefix}### (auto, on save)</span>} />
        <Row label="Style" value={`${order.styleNumber} — ${order.description || order.category}`} />
        <Row label="Requested qty" value={`${fmtNum(order.qty)} pcs`} />
        <Row label="Line" value={line ? `${line.lineName}${line.isDenim ? " (Denim)" : ""}` : "—"} />
        <Row label="Required production (adj. for efficiency)" value={`${fmtNum(result.requiredMinutes)} min`} />
        <Row label="Available capacity before booking" value={`${fmtNum(result.capBefore)} min/day`} />
        <Row label="Available capacity after booking" value={`${fmtNum(result.capAfter)} min/day`} />
        <Row label="Max planable qty" value={`${fmtNum(result.plannedQtyAtPCD)} pcs`} strong />
        <Row label="Estimated production start" value={niceDate(result.estimatedStart)} />
        <Row label="Estimated production end" value={niceDate(result.completionDate)} />
        <Row label="PCD" value={niceDate(order.pcd)} />
        {order.orderType === "Projection" && order.possibleOCD && <Row label="Possible OCD" value={niceDate(order.possibleOCD)} />}
        <Row label="Delivery / Ex-Factory" value={niceDate(order.deliveryDate)} />
        <Row label="Buffer days" value={result.buffer === null ? "—" : `${result.buffer >= 0 ? result.buffer : "–" + Math.abs(result.buffer)} days`} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, margin: "14px 0" }}>
        <MiniStatus label="Capacity" status={result.capacityStatus} />
        <MiniStatus label="Additional process" status={result.specialProcessStatus} />
        <MiniStatus label="PCD" status={result.pcdStatus} />
        <MiniStatus label="Delivery" status={result.deliveryStatus} />
      </div>

      {result.bottlenecks.length > 0 && !blocked && (
        <div style={{ background: COLORS.warningSoft, borderRadius: 4, padding: "10px 12px", marginBottom: 14 }}>
          <div style={{ fontSize: 11, fontWeight: 700, color: COLORS.warning, textTransform: "uppercase", marginBottom: 4 }}>Constraints</div>
          {result.bottlenecks.map((b, i) => <div key={i} style={{ fontSize: 12, color: COLORS.accentInk, marginBottom: 2 }}>• {b}</div>)}
          {result.additionalDaysNeeded > 0 && <div style={{ fontSize: 12, color: COLORS.accentInk, marginTop: 4 }}>Additional capacity required: approximately {result.additionalDaysNeeded} production day{result.additionalDaysNeeded > 1 ? "s" : ""}.</div>}
        </div>
      )}

      {!blocked && (
        <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
          {(result.decision === "ACCEPTED" || result.decision === "ACCEPTED WITH WARNING") && <Btn variant="success" onClick={onConfirmFull}>Confirm booking — {fmtNum(order.qty)} pcs</Btn>}
          {result.plannedQtyAtPCD > 0 && result.plannedQtyAtPCD < order.qty && result.denimEligible && <Btn variant="success" onClick={onConfirmPartial}>Plan {fmtNum(result.plannedQtyAtPCD)} pcs now (remaining {fmtNum(order.qty - result.plannedQtyAtPCD)} pcs)</Btn>}
          {result.capacityStatus === "INSUFFICIENT" && result.denimEligible && (
            <div style={{ background: COLORS.dangerSoft, borderRadius: 5, padding: "10px 12px" }}>
              <div style={{ fontSize: 12, color: COLORS.danger, fontWeight: 700, marginBottom: 6 }}>Insufficient production capacity available for the selected production period. Required: {fmtNum(result.requiredMinutes)} min. Available: {fmtNum(result.cumAtPCD)} min.</div>
              <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: COLORS.ink, marginBottom: 8 }}><input type="checkbox" checked={overrideChecked} onChange={e => setOverrideChecked(e.target.checked)} />I am authorized to override the capacity limit</label>
              <Btn variant="danger" disabled={!overrideChecked} onClick={onOverride}>Force confirm with override — {fmtNum(order.qty)} pcs</Btn>
            </div>
          )}
        </div>
      )}
    </Card>
  );
}
function Row({ label, value, strong }) { return <div style={{ display: "flex", justifyContent: "space-between", borderBottom: `1px dashed ${COLORS.line}`, padding: "3px 0" }}><span style={{ color: COLORS.inkSoft }}>{label}</span><span className="mono" style={{ fontWeight: strong ? 700 : 500 }}>{value}</span></div>; }
function MiniStatus({ label, status }) { return <div style={{ background: "#FAFAF9", border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "7px 9px" }}><div style={{ fontSize: 10, color: COLORS.inkFaint, textTransform: "uppercase", marginBottom: 4 }}>{label}</div><StatusPill status={status} /></div>; }

/* ============================================================
   BOOKING CONFIRMATION EMAIL
   ============================================================ */

function buildEmailContent(info) {
  const { booking: b, factory, hall, line, capBefore, capAfter } = info;
  const subject = `New Capacity Booking Created – ${b.itemNumber} – ${b.styleNumber}`;
  const lines = [
    `A new Capacity Booking has been created.`, ``,
    `${b.orderType === "Projection" ? "Projection Number" : "Item Number"}: ${b.referenceNumber}`,
    `Style Number: ${b.styleNumber}`, `Customer / Brand: ${b.customer} / ${b.brand}`,
    `Factory: ${factory.name}`, `Hall: ${hall?.name || "—"}`, `Line: ${line.lineName}${line.isDenim ? " (Denim)" : ""}`,
    `Order Type: ${b.orderType}`, `Season: ${b.season}`, `Product Category: ${b.category}`, `Marketing & Merchandising Head: ${b.mmHead}`,
    `Quantity: ${fmtNum(b.qty)} pcs`, `SAM: ${b.sam} min`, `Budgeted Efficiency: ${b.budgetedEfficiency}%`,
    b.fob ? `FOB: ${b.fob} (Total: ${fmtNum(b.qty * b.fob)})` : null,
    `PCD: ${niceDate(b.pcd)}`,
    b.orderType === "Projection" && b.possibleOCD ? `Possible OCD (Order Confirmation Date): ${niceDate(b.possibleOCD)}` : null,
    `Production Start Date: ${niceDate(b.startDate)}`, `Production End Date: ${niceDate(b.completionDate)}`,
    `Required Minutes / Capacity Booked: ${fmtNum(b.requiredMinutes)} min`, `Available Capacity Before Booking: ${fmtNum(capBefore)} min/day`, `Available Capacity After Booking: ${fmtNum(capAfter)} min/day`,
    `Raw Material Availability: ${b.rawMaterialAvailability || "Not specified"}`,
    `Additional Process: ${(b.additionalProcesses || []).join(", ") || "None"}`,
    b.additionalProcessRemark ? `Additional Process Remark: ${b.additionalProcessRemark}` : null,
    b.remarks ? `Remarks: ${b.remarks}` : null,
    ``, `Booked by: ${b.createdBy || "—"}`, `Booking Date/Time: ${b.createdAtTs ? new Date(b.createdAtTs).toLocaleString() : niceDate(b.bookingDate)}`
  ].filter(Boolean);
  return { subject, body: lines.join("\n") };
}

function BookingConfirmedModal({ info, savedEmailContacts, data, persist, showToast, onClose, onDone }) {
  const [selected, setSelected] = useState(info.booking.infoAlertEmails || []);
  const [newEmail, setNewEmail] = useState("");
  const [autoTriggered, setAutoTriggered] = useState(false);
  const { subject, body } = buildEmailContent(info);
  const currentStatus = (data.bookings.find(b => b.id === info.booking.id) || {}).emailStatus || info.booking.emailStatus;
  const b = info.booking;

  const toggle = (email) => setSelected(selected.includes(email) ? selected.filter(e => e !== email) : [...selected, email]);
  const addContact = () => {
    const e = newEmail.trim();
    if (!e) return;
    if (!data.savedEmailContacts.includes(e)) persist({ ...data, savedEmailContacts: [...data.savedEmailContacts, e] });
    if (!selected.includes(e)) setSelected([...selected, e]);
    setNewEmail("");
  };
  const setEmailStatus = (status) => persist({ ...data, bookings: data.bookings.map(x => x.id === b.id ? { ...x, emailStatus: status } : x) });
  const openMail = () => {
    if (selected.length === 0) { showToast("Select at least one recipient", "danger"); return; }
    window.location.href = `mailto:${encodeURIComponent(selected.join(","))}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
    setEmailStatus("Pending");
  };
  const copyContent = async () => {
    try { await navigator.clipboard.writeText(`To: ${selected.join(", ") || "(select recipients)"}\nSubject: ${subject}\n\n${body}`); showToast("Email content copied"); }
    catch (e) { showToast("Couldn't copy — select and copy manually", "danger"); }
  };

  // Automatically opens the mail client the moment this pop-up appears, if recipients were
  // already entered on the booking — the closest a browser-only app can honestly get to "automatic."
  useEffect(() => {
    if (!autoTriggered && selected.length > 0) { openMail(); setAutoTriggered(true); }
  }, []);

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(27,36,48,0.45)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 80, padding: 20 }}>
      <div style={{ background: COLORS.bg, borderRadius: 14, maxWidth: 760, width: "100%", maxHeight: "90vh", overflowY: "auto", boxShadow: "0 20px 60px rgba(20,30,50,0.35)", padding: 26 }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
          <div style={{ background: COLORS.successSoft, border: `1px solid ${COLORS.success}55`, borderRadius: 10, padding: "14px 18px", flex: 1, marginRight: 12 }}>
            <div style={{ fontSize: 11, color: COLORS.success, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5 }}>Booking Confirmed</div>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 18, color: COLORS.success, marginTop: 2 }}>
              Item Number: <span className="mono">{b.itemNumber}</span>
            </div>
            <div style={{ fontSize: 12.5, color: COLORS.success, marginTop: 2 }}>{fmtNum(b.qty)} pcs — {b.styleNumber} — booked on {info.line.lineName}</div>
          </div>
          <button onClick={onClose} style={{ background: "transparent", border: "none", cursor: "pointer", fontSize: 20, color: COLORS.inkFaint, lineHeight: 1 }} title="Close">✕</button>
        </div>

        <Card title="Booking details" style={{ marginTop: 16 }}>
          <div style={{ fontSize: 12.5, lineHeight: 1.9 }}>
            <Row label="Order Type" value={b.orderType} />
            <Row label="Customer / Brand" value={`${b.customer} / ${b.brand}`} />
            <Row label="Style Number" value={b.styleNumber} />
            <Row label="Factory / Hall / Line" value={`${info.factory.name} / ${info.hall?.name || "—"} / ${info.line.lineName}`} />
            <Row label="Season / Category" value={`${b.season} / ${b.category}`} />
            <Row label="Marketing & Merchandising Head" value={b.mmHead} />
            <Row label="Quantity" value={`${fmtNum(b.qty)} pcs`} />
            <Row label="SAM" value={`${b.sam} min`} />
            <Row label="Required Minutes" value={fmtNum(b.requiredMinutes)} />
            <Row label="PCD" value={niceDate(b.pcd)} />
            {b.orderType === "Projection" && b.possibleOCD && <Row label="Possible OCD" value={niceDate(b.possibleOCD)} />}
            <Row label="Delivery / Ex-Factory" value={niceDate(b.deliveryDate)} />
            <Row label="Production Start / End" value={`${niceDate(b.startDate)} → ${niceDate(b.completionDate)}`} />
          </div>
        </Card>

        <Card title="Booking confirmation email" action={currentStatus && currentStatus !== "N/A" ? <StatusPill status={currentStatus.toUpperCase()} /> : null} style={{ marginTop: 16 }}>
          {autoTriggered && selected.length > 0 && (
            <div style={{ background: COLORS.accentSoft, borderRadius: 5, padding: "10px 12px", marginBottom: 14, fontSize: 12, color: COLORS.accentInk }}>
              Your email app should have opened with this message pre-filled — send it from there. This app has no backend, so it can't silently deliver mail on its own; opening your real mail client is the most genuine send available here.
            </div>
          )}
          <div style={{ fontSize: 11, color: COLORS.inkFaint, marginBottom: 8, textTransform: "uppercase", fontWeight: 600 }}>Recipients</div>
          <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>{savedEmailContacts.map(e => <Toggle key={e} active={selected.includes(e)} onClick={() => toggle(e)}>{e}</Toggle>)}</div>
          <div style={{ display: "flex", gap: 8, marginBottom: 16 }}><TextInput value={newEmail} onChange={ev => setNewEmail(ev.target.value)} placeholder="add-new@company.com" onKeyDown={ev => { if (ev.key === "Enter") { ev.preventDefault(); addContact(); } }} /><Btn onClick={addContact}>Add & select</Btn></div>
          <div style={{ fontSize: 11, color: COLORS.inkFaint, marginBottom: 6, textTransform: "uppercase", fontWeight: 600 }}>Preview</div>
          <div style={{ background: "#FAFAF9", border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: 12, fontSize: 12, whiteSpace: "pre-wrap", fontFamily: "'IBM Plex Mono', monospace", maxHeight: 220, overflowY: "auto" }}>Subject: {subject}{"\n\n"}{body}</div>
          <div style={{ display: "flex", gap: 10, marginTop: 14, flexWrap: "wrap" }}>
            <Btn variant="primary" onClick={openMail}>Open in email client again</Btn>
            <Btn onClick={copyContent}>Copy email content</Btn>
            <Btn variant="success" onClick={() => { setEmailStatus("Sent"); showToast("Marked as sent"); }}>Mark as Sent</Btn>
            <Btn variant="danger" onClick={() => { setEmailStatus("Failed"); showToast("Marked as failed — booking is still saved", "danger"); }}>Mark as Failed</Btn>
          </div>
          <div style={{ fontSize: 11, color: COLORS.inkFaint, marginTop: 10 }}>
            This composes the email and opens it in your own mail app (or copies it) — it does not send silently from a server. The booking itself is already saved regardless of what happens with the email.
          </div>
        </Card>

        <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
          <Btn onClick={onClose}>Close — stay on Capacity Booking</Btn>
          <Btn variant="primary" onClick={onDone}>Go to Style / Order Master</Btn>
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   ORDERS / STYLE MASTER
   ============================================================ */

function OrdersMaster({ factories, lines, bookings, filter, setFilter, data, persist, showToast }) {
  const filtered = bookings.filter(b => {
    const q = filter.q.toLowerCase();
    const matchQ = !q || b.itemNumber.toLowerCase().includes(q) || b.styleNumber.toLowerCase().includes(q) || b.customer.toLowerCase().includes(q) || (b.mmHead || "").toLowerCase().includes(q);
    const matchStatus = filter.status === "ALL" || b.status === filter.status;
    return matchQ && matchStatus;
  });
  const remove = (id) => { persist({ ...data, bookings: bookings.filter(b => b.id !== id) }); showToast("Booking removed"); };
  const confirmOrder = (id) => { persist({ ...data, bookings: bookings.map(b => b.id === id ? { ...b, confirmed: true, orderType: "Confirmed Order" } : b) }); showToast("Order confirmed"); };
  const statuses = ["ALL", ...Array.from(new Set(bookings.map(b => b.status)))];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ display: "flex", gap: 10 }}>
        <TextInput placeholder="Search item, style, customer, or M&M Head…" value={filter.q} onChange={e => setFilter({ ...filter, q: e.target.value })} style={{ maxWidth: 340 }} />
        <Select value={filter.status} onChange={e => setFilter({ ...filter, status: e.target.value })} style={{ maxWidth: 220 }}>{statuses.map(s => <option key={s} value={s}>{s}</option>)}</Select>
      </div>
      <Card>
        <table>
          <thead><tr><th>Item</th><th>Order type</th><th>Style</th><th>Customer</th><th>M&M Head</th><th>Qty</th><th>Factory / line</th><th>Confirmed?</th><th>PCD</th><th>Status</th><th>Email</th><th></th></tr></thead>
          <tbody>{filtered.map(b => {
            const factory = factories.find(f => f.id === b.factoryId), line = lines.find(l => l.id === b.lineId);
            return (
              <tr key={b.id}>
                <td className="mono">{b.itemNumber}</td>
                <td><StatusPill status={b.orderType === "Projection" ? "PROJECTION" : "CONFIRMED"} /></td>
                <td>{b.styleNumber}</td><td>{b.customer}</td><td style={{ fontSize: 11.5 }}>{b.mmHead}</td>
                <td className="mono">{fmtNum(b.qty)}{b.remainingQty > 0 && <span style={{ color: COLORS.warning }}> (+{fmtNum(b.remainingQty)} pending)</span>}</td>
                <td style={{ fontSize: 11.5 }}>{factory?.name}<br /><span style={{ color: COLORS.inkFaint }}>{line?.lineName}{line?.isDenim ? " (Denim)" : ""}</span></td>
                <td>{b.confirmed ? <StatusPill status="ORDER CONFIRMED" /> : <><StatusPill status="PROJECTION" /><div style={{ marginTop: 4 }}><Btn variant="ghost" onClick={() => confirmOrder(b.id)}>Confirm order</Btn></div></>}</td>
                <td>{niceDate(b.pcd)}</td><td><StatusPill status={b.status} /></td><td>{b.emailStatus && b.emailStatus !== "N/A" ? <StatusPill status={b.emailStatus.toUpperCase()} /> : <span style={{ color: COLORS.inkFaint, fontSize: 11 }}>N/A</span>}</td><td><Btn variant="ghost" onClick={() => remove(b.id)}>Remove</Btn></td>
              </tr>
            );
          })}</tbody>
        </table>
        {filtered.length === 0 && <Empty text="No orders match this filter." />}
      </Card>
    </div>
  );
}

/* ============================================================
   LINE LOADING BOARD
   ============================================================ */

function LineLoadingBoard({ factories, halls, lines, bookings }) {
  const today = todayStr();
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 20 }}>
      {factories.map(f => {
        const fHalls = halls.filter(h => h.factoryId === f.id);
        return (
          <div key={f.id}>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, fontSize: 14, marginBottom: 10 }}>{f.name}</div>
            {fHalls.map(h => {
              const hLines = lines.filter(l => l.hallId === h.id);
              if (hLines.length === 0) return null;
              return (
                <div key={h.id} style={{ marginBottom: 14 }}>
                  {fHalls.length > 1 && <div style={{ fontSize: 12, fontWeight: 600, color: COLORS.accentInk, marginBottom: 8 }}>{h.name}</div>}
                  <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(240px, 1fr))", gap: 12 }}>
                    {hLines.map(l => {
                      const cap = getLineDailyMinutes(l, f);
                      const todaysBookings = bookings.filter(b => b.lineId === l.id && !isBlockedStatus(b.status) && b.startDate <= today && b.completionDate >= today);
                      const booked = todaysBookings.reduce((s, b) => s + (b.dailyMinutes || 0), 0);
                      const util = cap > 0 ? (booked / cap) * 100 : 0;
                      const status = util >= 100 ? "Fully booked" : util >= 80 ? "High utilization" : "Available";
                      const color = util >= 100 ? COLORS.danger : util >= 80 ? COLORS.warning : COLORS.success;
                      return (
                        <div key={l.id} style={{ background: COLORS.panel, border: `1px solid ${COLORS.line}`, borderLeft: `3px solid ${l.isDenim ? COLORS.denim : color}`, borderRadius: 4, padding: "12px 14px" }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}><div style={{ fontWeight: 600, fontSize: 13 }}>{l.lineName}</div>{l.isDenim && <DenimTag />}</div>
                          <div style={{ fontSize: 11, color: COLORS.inkFaint, margin: "6px 0" }}>{status}</div>
                          <div className="mono" style={{ fontSize: 12, marginBottom: 4 }}><span style={{ color: COLORS.inkSoft }}>Capacity:</span> {fmtNum(cap)} min/day</div>
                          <div className="mono" style={{ fontSize: 12, marginBottom: 4 }}><span style={{ color: COLORS.inkSoft }}>Booked:</span> {fmtNum(booked)} min/day</div>
                          <div className="mono" style={{ fontSize: 12, marginBottom: 8 }}><span style={{ color: COLORS.inkSoft }}>Available:</span> {fmtNum(Math.max(0, cap - booked))} min/day</div>
                          <UtilBar pct={util} />
                          <div style={{ textAlign: "right", fontSize: 11, marginTop: 4, fontWeight: 600, color }}>{util.toFixed(0)}%</div>
                          {todaysBookings.length > 0 && <div style={{ marginTop: 10, paddingTop: 8, borderTop: `1px dashed ${COLORS.line}` }}>{todaysBookings.map(b => <div key={b.id} style={{ display: "flex", justifyContent: "space-between", fontSize: 10.5, color: COLORS.inkFaint, marginBottom: 2 }}><span className="mono">{b.itemNumber}</span><span>{fmtNum(b.dailyMinutes)} min</span></div>)}</div>}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        );
      })}
    </div>
  );
}

/* ============================================================
   PRODUCTION PLANNING (simplified Gantt)
   ============================================================ */

function ProductionSchedule({ factories, lines, bookings }) {
  const active = bookings.filter(b => !isBlockedStatus(b.status));
  if (active.length === 0) return <Empty text="No production booked yet." />;
  const allStarts = active.map(b => parseDate(b.startDate).getTime());
  const allEnds = active.map(b => parseDate(b.completionDate || b.pcd).getTime());
  const minDate = new Date(Math.min(...allStarts));
  const maxDate = new Date(Math.max(...allEnds, ...active.map(b => parseDate(b.deliveryDate).getTime())));
  const totalDays = Math.max(1, Math.round((maxDate - minDate) / 86400000));
  const pctFor = (dateStr) => Math.max(0, Math.min(100, ((parseDate(dateStr) - minDate) / 86400000) / totalDays * 100));

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 22 }}>
      {factories.map(f => {
        const fLines = lines.filter(l => l.factoryId === f.id);
        if (!fLines.some(l => active.some(b => b.lineId === l.id))) return null;
        return (
          <div key={f.id}>
            <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 600, fontSize: 14, marginBottom: 10 }}>{f.name}</div>
            <Card>
              {fLines.map(l => {
                const lb = active.filter(b => b.lineId === l.id).sort((a, b) => a.startDate.localeCompare(b.startDate));
                if (lb.length === 0) return null;
                return (
                  <div key={l.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "8px 0", borderBottom: `1px solid ${COLORS.line}` }}>
                    <div style={{ width: 140, fontSize: 12, fontWeight: 600, flexShrink: 0 }}>{l.lineName}{l.isDenim ? " (D)" : ""}</div>
                    <div style={{ flex: 1, position: "relative", height: 26, background: "#EDEDEB", borderRadius: 3 }}>
                      {lb.map(b => {
                        const left = pctFor(b.startDate), right = pctFor(b.completionDate || b.pcd), width = Math.max(1.5, right - left);
                        const color = b.pcdStatus === "NOT ACHIEVABLE" || b.deliveryStatus === "NOT ACHIEVABLE" ? COLORS.danger : b.pcdStatus === "AT RISK" || b.deliveryStatus === "AT RISK" ? COLORS.warning : COLORS.accent;
                        return (
                          <div key={b.id} title={`${b.itemNumber} · ${b.styleNumber} · ${niceDate(b.startDate)} → ${niceDate(b.completionDate)}${!b.confirmed ? " (Projection)" : ""}`} style={{ position: "absolute", left: `${left}%`, width: `${width}%`, top: 3, bottom: 3, background: color, opacity: b.confirmed === false ? 0.55 : 1, border: b.confirmed === false ? `1px dashed ${COLORS.ink}55` : "none", borderRadius: 3, display: "flex", alignItems: "center", overflow: "hidden", padding: "0 6px" }}>
                            <span style={{ fontSize: 10, color: "#fff", fontWeight: 600, whiteSpace: "nowrap" }}>{b.styleNumber}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </Card>
          </div>
        );
      })}
      <div style={{ fontSize: 11.5, color: COLORS.inkFaint }}>Timeline: {niceDate(fmt(minDate))} → {niceDate(fmt(maxDate))}. Dashed/faded bars are unconfirmed Projections.</div>
    </div>
  );
}

/* ============================================================
   CAPACITY RISK & PROJECTION REMINDERS
   ============================================================ */

function CapacityRisk({ factories, lines, bookings, additionalProcesses, settings, savedEmailContacts }) {
  const active = bookings.filter(b => !isBlockedStatus(b.status));
  const risky = active.filter(b => b.pcdStatus !== "ACHIEVABLE" || b.deliveryStatus !== "ACHIEVABLE" || b.specialProcessStatus === "SHORTAGE" || b.status === "PARTIALLY ACCEPTED");
  const holds = bookings.filter(b => b.status === "RAW MATERIAL HOLD");
  const projections = bookings.filter(b => !b.confirmed).map(b => ({ booking: b, ...projectionReminderTier(b, settings) })).sort((a, b) => (a.daysToPCD ?? 9999) - (b.daysToPCD ?? 9999));

  const sendReminder = (booking) => {
    const recipients = (booking.infoAlertEmails && booking.infoAlertEmails.length ? booking.infoAlertEmails : savedEmailContacts);
    if (!recipients.length) { window.alert("No reminder recipients configured for this projection — add Information Alert Email ID(s) on the booking or in Settings."); return; }
    const subject = `Action needed: convert ${booking.itemNumber} to Confirmed Order`;
    const body = [
      `This Projection needs to be changed to Confirmed Order before ${settings.normalPlanningDays || 90} days from PCD. Please review and confirm the order status.`, "",
      `Projection Number: ${booking.itemNumber}`, `Customer: ${booking.customer}`, `Brand: ${booking.brand}`, `Style Number: ${booking.styleNumber}`,
      `Marketing & Merchandising Head: ${booking.mmHead}`, `PCD: ${niceDate(booking.pcd)}`, `Current Date: ${niceDate(todayStr())}`,
      `Days Remaining to PCD: ${daysBetween(todayStr(), booking.pcd)}`, `Current Status: Projection (unconfirmed)`, `Required Action: Convert to Confirmed Order`
    ].join("\n");
    window.location.href = `mailto:${encodeURIComponent(recipients.join(","))}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  };

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 18 }}>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12 }}>
        <Metric label="Orders at risk" value={risky.length} tone={risky.length > 0 ? "warning" : "success"} />
        <Metric label="Raw material holds" value={holds.length} tone={holds.length > 0 ? "danger" : "success"} />
        <Metric label="Open projections" value={projections.length} />
        <Metric label="Projections needing action" value={projections.filter(p => p.tier !== "OK").length} tone={projections.some(p => p.tier !== "OK") ? "warning" : "success"} />
      </div>

      <Card title="Projection → Confirmed Order reminders">
        <p style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 0 }}>Standard schedule: 120 days = initial reminder, 105 days = follow-up, 90 days = confirmation required, under 90 days = escalation.</p>
        {projections.length === 0 ? <Empty text="No open projections." /> : (
          <table><thead><tr><th>Projection No.</th><th>Customer</th><th>Style</th><th>M&M Head</th><th>PCD</th><th>Days Remaining</th><th>Status</th><th></th></tr></thead>
            <tbody>{projections.map(({ booking, tier, daysToPCD }) => (
              <tr key={booking.id}>
                <td className="mono">{booking.itemNumber}</td><td>{booking.customer}</td><td>{booking.styleNumber}</td><td style={{ fontSize: 11.5 }}>{booking.mmHead}</td>
                <td>{niceDate(booking.pcd)}</td><td className="mono">{daysToPCD}</td><td><StatusPill status={tier} /></td>
                <td>{tier !== "OK" && <Btn variant="ghost" onClick={() => sendReminder(booking)}>Send reminder</Btn>}</td>
              </tr>
            ))}</tbody>
          </table>
        )}
      </Card>

      <Card title="Raw material holds">
        {holds.length === 0 ? <Empty text="No bookings on raw-material hold." /> : (
          <table><thead><tr><th>Item</th><th>Style</th><th>Customer</th><th>PCD</th><th>Reason</th></tr></thead>
            <tbody>{holds.map(b => <tr key={b.id}><td className="mono">{b.itemNumber}</td><td>{b.styleNumber}</td><td>{b.customer}</td><td>{niceDate(b.pcd)}</td><td style={{ color: COLORS.danger, fontSize: 12 }}>{b.bottleneck}</td></tr>)}</tbody>
          </table>
        )}
      </Card>
      <Card title="At-risk orders — root cause">
        {risky.length === 0 ? <Empty text="No orders currently at risk." /> : (
          <table><thead><tr><th>Item</th><th>Style</th><th>Customer</th><th>Line</th><th>Capacity</th><th>PCD</th><th>Delivery</th><th>Main bottleneck</th></tr></thead>
            <tbody>{risky.map(b => { const line = lines.find(l => l.id === b.lineId); return (
              <tr key={b.id}><td className="mono">{b.itemNumber}</td><td>{b.styleNumber}</td><td>{b.customer}</td><td>{line?.lineName}{line?.isDenim ? " (D)" : ""}</td><td><StatusPill status={b.capacityStatus} /></td><td><StatusPill status={b.pcdStatus} /></td><td><StatusPill status={b.deliveryStatus} /></td><td style={{ fontSize: 11.5, color: COLORS.inkSoft, maxWidth: 260 }}>{b.bottleneck || "—"}</td></tr>
            ); })}</tbody>
          </table>
        )}
      </Card>
      <Card title="Additional process capacity">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(200px,1fr))", gap: 12 }}>
          {additionalProcesses.map(sp => { const avail = Math.max(0, sp.capacityPerDay - sp.bookedPerDay); const util = (sp.bookedPerDay / sp.capacityPerDay) * 100; return (
            <div key={sp.id} style={{ border: `1px solid ${COLORS.line}`, borderRadius: 4, padding: "10px 12px" }}>
              <div style={{ fontWeight: 600, fontSize: 12.5, marginBottom: 6 }}>{sp.name}</div>
              <div className="mono" style={{ fontSize: 11.5, color: COLORS.inkSoft }}>{fmtNum(sp.bookedPerDay)} / {fmtNum(sp.capacityPerDay)} pcs/day</div>
              <div style={{ margin: "6px 0" }}><UtilBar pct={util} /></div>
              <div className="mono" style={{ fontSize: 11.5, color: avail > 0 ? COLORS.success : COLORS.danger }}>{fmtNum(avail)} pcs/day free</div>
            </div>
          ); })}
        </div>
      </Card>
    </div>
  );
}

/* ============================================================
   ADDITIONAL PROCESSES
   ============================================================ */

function AdditionalProcessesView({ additionalProcesses, bookings, data, persist, showToast }) {
  const update = (id, field, value) => persist({ ...data, additionalProcesses: additionalProcesses.map(sp => sp.id === id ? { ...sp, [field]: Number(value) || 0 } : sp) });
  const withProcesses = bookings.filter(b => (b.additionalProcesses || []).length > 0);
  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <Card title="Additional process capacity setup">
        <table><thead><tr><th>Process</th><th>Capacity /day (pcs)</th><th>Booked /day (pcs)</th><th>Available /day</th><th>Utilization</th></tr></thead>
          <tbody>{additionalProcesses.map(sp => { const avail = Math.max(0, sp.capacityPerDay - sp.bookedPerDay); const util = (sp.bookedPerDay / sp.capacityPerDay) * 100; return (
            <tr key={sp.id}><td style={{ fontWeight: 600 }}>{sp.name}</td>
              <td><TextInput type="number" value={sp.capacityPerDay} onChange={e => update(sp.id, "capacityPerDay", e.target.value)} style={{ width: 110 }} /></td>
              <td><TextInput type="number" value={sp.bookedPerDay} onChange={e => update(sp.id, "bookedPerDay", e.target.value)} style={{ width: 110 }} /></td>
              <td className="mono">{fmtNum(avail)}</td><td style={{ width: 140 }}><UtilBar pct={util} /></td></tr>
          ); })}</tbody>
        </table>
        <div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 10 }}>Printing, Embroidery, Washing, and Garment Dye have tracked daily capacity. Special Finishing and Other can be selected but aren't capacity-checked yet.</div>
      </Card>
      <Card title="Orders using additional processes">
        <table><thead><tr><th>Item</th><th>Style</th><th>Processes</th><th>Details</th><th>Qty</th><th>Status</th></tr></thead>
          <tbody>{withProcesses.map(b => (
            <tr key={b.id}><td className="mono">{b.itemNumber}</td><td>{b.styleNumber}</td><td style={{ fontSize: 11.5 }}>{b.additionalProcesses.join(", ")}</td>
              <td style={{ fontSize: 11 }}>{b.embroideryStitches ? `${fmtNum(b.embroideryStitches)} stitches` : ""}{b.washType ? ` ${b.washType}` : ""}{b.additionalProcessRemark ? ` — ${b.additionalProcessRemark}` : ""}</td>
              <td className="mono">{fmtNum(b.qty)}</td><td><StatusPill status={b.status} /></td></tr>
          ))}</tbody>
        </table>
        {withProcesses.length === 0 && <Empty text="No orders currently use additional processes." />}
      </Card>
    </div>
  );
}

/* ============================================================
   REPORTS
   ============================================================ */

function toCSV(rows, headers) {
  const esc = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
  const out = [headers.map(esc).join(",")];
  rows.forEach(r => out.push(headers.map(h => esc(r[h])).join(",")));
  return out.join("\n");
}
function downloadFile(filename, content, mime) {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a"); a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function exportCSV(name, rows, headers) { downloadFile(`mbvb-${name}.csv`, `${BRAND_NAME}\r\n\r\n${toCSV(rows, headers)}`, "text/csv;charset=utf-8;"); }
function exportExcel(name, rows) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "Report");
  XLSX.writeFile(wb, `mbvb-${name}.xlsx`);
}
function ReportBlock({ name, desc, headers, rows, onPrint }) {
  return (
    <Card>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", flexWrap: "wrap", gap: 10 }}>
        <div><div style={{ fontWeight: 600, fontSize: 13 }}>{name}</div><div style={{ fontSize: 11.5, color: COLORS.inkFaint, marginTop: 3 }}>{desc}</div></div>
        <div style={{ display: "flex", gap: 6 }}>
          <Btn onClick={() => exportCSV(name.toLowerCase().replace(/\s+/g, "-"), rows, headers)}>CSV</Btn>
          <Btn onClick={() => exportExcel(name.toLowerCase().replace(/\s+/g, "-"), rows)}>Excel</Btn>
          <Btn onClick={() => onPrint({ title: name, headers, rows })}>PDF</Btn>
        </div>
      </div>
    </Card>
  );
}

function Reports({ factories, halls, lines, bookings, onPrint }) {
  const factoriesById = Object.fromEntries(factories.map(f => [f.id, f]));
  const hallsById = Object.fromEntries(halls.map(h => [h.id, h]));

  const [mhFilter, setMhFilter] = useState({ head: "ALL", customer: "", brand: "", season: "ALL", factory: "ALL", orderType: "ALL", from: "", to: "" });
  const mhFiltered = bookings.filter(b => {
    if (mhFilter.head !== "ALL" && b.mmHead !== mhFilter.head) return false;
    if (mhFilter.customer && !b.customer.toLowerCase().includes(mhFilter.customer.toLowerCase())) return false;
    if (mhFilter.brand && !b.brand.toLowerCase().includes(mhFilter.brand.toLowerCase())) return false;
    if (mhFilter.season !== "ALL" && b.season !== mhFilter.season) return false;
    if (mhFilter.factory !== "ALL" && b.factoryId !== mhFilter.factory) return false;
    if (mhFilter.orderType !== "ALL" && b.orderType !== mhFilter.orderType) return false;
    if (mhFilter.from && b.pcd < mhFilter.from) return false;
    if (mhFilter.to && b.pcd > mhFilter.to) return false;
    return true;
  });
  const mmHeadRows = MM_HEADS.map(head => {
    const rows = mhFiltered.filter(b => b.mmHead === head);
    const active = rows.filter(b => !isBlockedStatus(b.status));
    const cap = active.reduce((s, b) => s + (b.dailyMinutes || 0), 0);
    const groupCap = lines.filter(l => l.active).reduce((s, l) => s + getLineDailyMinutes(l, factoriesById[l.factoryId]), 0);
    return {
      "Marketing & Merchandising Head": head, "No. of Styles": rows.length, "No. of Customers": new Set(rows.map(b => b.customer)).size,
      "Total Quantity": rows.reduce((s, b) => s + b.qty, 0), "Total SAM Minutes": rows.reduce((s, b) => s + b.qty * b.sam, 0),
      "Total Required Production Minutes": rows.reduce((s, b) => s + b.requiredMinutes, 0), "Total FOB Value": rows.reduce((s, b) => s + b.qty * (b.fob || 0), 0),
      "Projection Count": rows.filter(b => b.orderType === "Projection").length, "Confirmed Order Count": rows.filter(b => b.orderType === "Confirmed Order").length,
      "Capacity Utilisation %": groupCap ? +(cap / groupCap * 100).toFixed(1) : 0
    };
  });

  const seasons = Array.from(new Set(bookings.map(b => b.season)));

  const capacityRows = () => bookings.map(b => {
    const f = factoriesById[b.factoryId], line = lines.find(l => l.id === b.lineId);
    const cap = line ? getLineDailyMinutes(line, f) : 0;
    return { Factory: f?.name, Hall: hallsById[b.hallId]?.name, Line: line?.lineName, Season: b.season, "Style Number": b.styleNumber, "Reference No.": b.referenceNumber, "Order Type": b.orderType, "M&M Head": b.mmHead, "Product Category": b.category, Quantity: b.qty, SAM: b.sam, "Budgeted Efficiency %": b.budgetedEfficiency, FOB: b.fob, "Total FOB Value": b.qty * (b.fob || 0), "Required Minutes": b.requiredMinutes, "Available Minutes": Math.round(cap), "Booked Minutes": b.dailyMinutes, "Remaining Minutes": Math.max(0, Math.round(cap - b.dailyMinutes)), "Capacity Utilisation %": cap ? Math.round((b.dailyMinutes / cap) * 100) : 0, PCD: niceDate(b.pcd), "Production Start Date": niceDate(b.startDate), "Production End Date": niceDate(b.completionDate), Status: b.status };
  });
  const bookingRows = () => bookings.map(b => ({ "Booking Date": niceDate(b.bookingDate), Factory: factoriesById[b.factoryId]?.name, Hall: hallsById[b.hallId]?.name, Line: lines.find(l => l.id === b.lineId)?.lineName, Style: b.styleNumber, "Reference No.": b.referenceNumber, "Order Type": b.orderType, "M&M Head": b.mmHead, Season: b.season, "Product Category": b.category, Quantity: b.qty, SAM: b.sam, PCD: niceDate(b.pcd), "Raw Material Availability": b.rawMaterialAvailability || "Not specified", "Additional Process": (b.additionalProcesses || []).join(", ") || "None", "Booking Status": b.status, "Booked By": b.createdBy || "—", "Email Status": b.emailStatus || "N/A" }));
  const ieRows = () => lines.filter(l => l.plannedPcsPerDay).map(l => {
    const f = factoriesById[l.factoryId];
    const todaysB = bookings.filter(b => b.lineId === l.id && !isBlockedStatus(b.status) && b.startDate <= todayStr() && b.completionDate >= todayStr());
    const bookedMin = todaysB.reduce((s, b) => s + (b.dailyMinutes || 0), 0);
    const cap = getLineDailyMinutes(l, f);
    return { Factory: f?.name, Hall: hallsById[l.hallId]?.name, Line: l.lineName, Date: niceDate(todayStr()), "Planned Pcs Per Day": l.plannedPcsPerDay, "Planned Quantity": todaysB[0]?.qty || l.plannedPcsPerDay, "Actual Quantity": "", Variance: "", "Capacity Utilisation %": cap ? Math.round((bookedMin / cap) * 100) : 0 };
  });

  const reports = [
    { name: "Capacity Report", desc: "Full capacity detail per booked style", build: capacityRows },
    { name: "Booking Report", desc: "Booking log with order type, raw material, and process detail", build: bookingRows },
    { name: "Daily Production / IE Report", desc: "IE planned pcs/day vs booked utilisation, per line", build: ieRows },
    { name: "PCD Risk Report", desc: "Orders where PCD is at risk or not achievable", build: () => bookings.filter(b => b.pcdStatus !== "ACHIEVABLE").map(b => ({ Item: b.itemNumber, Style: b.styleNumber, PCD: niceDate(b.pcd), Status: b.pcdStatus, Bottleneck: b.bottleneck })) },
    { name: "Delivery Risk Report", desc: "Orders where delivery is at risk or not achievable", build: () => bookings.filter(b => b.deliveryStatus !== "ACHIEVABLE").map(b => ({ Item: b.itemNumber, Style: b.styleNumber, Delivery: niceDate(b.deliveryDate), Status: b.deliveryStatus, Buffer: b.buffer })) },
    { name: "Customer Capacity Report", desc: "Planned quantity and FOB value by customer", build: () => { const by = {}; bookings.forEach(b => { if (!by[b.customer]) by[b.customer] = { Customer: b.customer, Qty: 0, "FOB Value": 0 }; by[b.customer].Qty += b.qty; by[b.customer]["FOB Value"] += b.qty * (b.fob || 0); }); return Object.values(by); } }
  ];

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div><div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16 }}>{BRAND_NAME}</div><div style={{ fontSize: 12, color: COLORS.inkFaint, marginTop: 2 }}>All major reports export as CSV, Excel, and PDF, with dates shown as DD/MM/YY.</div></div>

      <Card title="Marketing & Merchandising Head report">
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 10, marginBottom: 14 }}>
          <Select value={mhFilter.head} onChange={e => setMhFilter({ ...mhFilter, head: e.target.value })}><option value="ALL">All Heads</option>{MM_HEADS.map(h => <option key={h} value={h}>{h}</option>)}</Select>
          <TextInput placeholder="Customer / Buyer" value={mhFilter.customer} onChange={e => setMhFilter({ ...mhFilter, customer: e.target.value })} />
          <TextInput placeholder="Brand" value={mhFilter.brand} onChange={e => setMhFilter({ ...mhFilter, brand: e.target.value })} />
          <Select value={mhFilter.season} onChange={e => setMhFilter({ ...mhFilter, season: e.target.value })}><option value="ALL">All Seasons</option>{seasons.map(s => <option key={s} value={s}>{s}</option>)}</Select>
          <Select value={mhFilter.factory} onChange={e => setMhFilter({ ...mhFilter, factory: e.target.value })}><option value="ALL">All Factories</option>{factories.map(f => <option key={f.id} value={f.id}>{f.name}</option>)}</Select>
          <Select value={mhFilter.orderType} onChange={e => setMhFilter({ ...mhFilter, orderType: e.target.value })}><option value="ALL">All Order Types</option>{ORDER_TYPES.map(t => <option key={t} value={t}>{t}</option>)}</Select>
          <div><div style={{ fontSize: 10, color: COLORS.inkFaint, marginBottom: 3 }}>PCD from</div><DateField value={mhFilter.from} onChange={v => setMhFilter({ ...mhFilter, from: v })} /></div>
          <div><div style={{ fontSize: 10, color: COLORS.inkFaint, marginBottom: 3 }}>PCD to</div><DateField value={mhFilter.to} onChange={v => setMhFilter({ ...mhFilter, to: v })} /></div>
        </div>
        <table>
          <thead><tr><th>Head</th><th>Styles</th><th>Customers</th><th>Qty</th><th>SAM Min</th><th>Req. Min</th><th>FOB Value</th><th>Projections</th><th>Confirmed</th><th>Utilisation</th></tr></thead>
          <tbody>{mmHeadRows.map(r => (
            <tr key={r["Marketing & Merchandising Head"]}>
              <td style={{ fontWeight: 600 }}>{r["Marketing & Merchandising Head"]}</td><td className="mono">{r["No. of Styles"]}</td><td className="mono">{r["No. of Customers"]}</td>
              <td className="mono">{fmtNum(r["Total Quantity"])}</td><td className="mono">{fmtNum(r["Total SAM Minutes"])}</td><td className="mono">{fmtNum(r["Total Required Production Minutes"])}</td>
              <td className="mono">{fmtNum(r["Total FOB Value"])}</td><td className="mono">{r["Projection Count"]}</td><td className="mono">{r["Confirmed Order Count"]}</td><td className="mono">{r["Capacity Utilisation %"]}%</td>
            </tr>
          ))}</tbody>
        </table>
        <div style={{ display: "flex", gap: 6, marginTop: 12 }}>
          <Btn onClick={() => exportCSV("mm-head-report", mmHeadRows, Object.keys(mmHeadRows[0] || {}))}>CSV</Btn>
          <Btn onClick={() => exportExcel("mm-head-report", mmHeadRows)}>Excel</Btn>
          <Btn onClick={() => onPrint({ title: "Marketing & Merchandising Head Report", headers: Object.keys(mmHeadRows[0] || {}), rows: mmHeadRows })}>PDF</Btn>
        </div>
      </Card>

      {reports.map(r => { const rows = r.build(); const headers = rows.length ? Object.keys(rows[0]) : []; return <ReportBlock key={r.name} name={r.name} desc={r.desc} headers={headers} rows={rows} onPrint={onPrint} />; })}
    </div>
  );
}

function PrintReport({ job, onDone }) {
  useEffect(() => {
    const t = setTimeout(() => window.print(), 150);
    const handler = () => onDone();
    window.addEventListener("afterprint", handler);
    return () => { clearTimeout(t); window.removeEventListener("afterprint", handler); };
  }, []);
  const { title, headers, rows } = job;
  return (
    <div className="print-only" style={{ padding: 24, fontFamily: "Arial, sans-serif", color: "#111" }}>
      <div style={{ borderBottom: "3px solid #2D6CDF", paddingBottom: 10, marginBottom: 16 }}>
        <div style={{ fontSize: 20, fontWeight: 700 }}>{BRAND_NAME}</div>
        <div style={{ fontSize: 12, color: "#555" }}>{title} — generated {niceDate(todayStr())}</div>
      </div>
      <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 10 }}>
        <thead><tr style={{ background: "#eee" }}>{headers.map(h => <th key={h} style={{ border: "1px solid #ccc", padding: "4px 6px", textAlign: "left" }}>{h}</th>)}</tr></thead>
        <tbody>{rows.map((r, i) => <tr key={i}>{headers.map(h => <td key={h} style={{ border: "1px solid #ccc", padding: "4px 6px" }}>{typeof r[h] === "number" ? fmtNum(r[h]) : r[h]}</td>)}</tr>)}</tbody>
      </table>
    </div>
  );
}

/* ============================================================
   SETTINGS
   ============================================================ */

function Settings({ data, persist, showToast }) {
  const fileRef = useRef(null);
  const [importError, setImportError] = useState(null);
  const [newPw, setNewPw] = useState(""); const [newPw2, setNewPw2] = useState("");
  const settings = { ...DEFAULT_SETTINGS, ...(data.settings || {}) };

  const updateThreshold = (key, v) => persist({ ...data, settings: { ...settings, [key]: Number(v) || settings[key] } });
  const updateCategories = (list) => persist({ ...data, categories: list });
  const updateSeasons = (list) => persist({ ...data, seasons: list });
  const updateCustomers = (list) => persist({ ...data, customers: list });
  const updateBrands = (list) => persist({ ...data, brands: list });
  const updateContacts = (list) => persist({ ...data, savedEmailContacts: list });
  const reset = () => { persist(buildSampleData()); showToast("Sample data restored"); };
  const clear = () => { persist({ ...data, factories: [], halls: [], lines: [], bookings: [] }); showToast("All bookings cleared", "danger"); };
  const changePassword = () => {
    if (newPw.length < 4) { showToast("Password must be at least 4 characters", "danger"); return; }
    if (newPw !== newPw2) { showToast("Passwords don't match", "danger"); return; }
    persist({ ...data, settings: { ...settings, ppcPassword: newPw } });
    setNewPw(""); setNewPw2(""); showToast("Password updated");
  };

  const exportBackup = () => {
    const payload = { exportedAt: new Date().toISOString(), system: BRAND_NAME, version: 6, data };
    downloadFile(`mbvb-capacity-planner-backup-${todayStr()}.json`, JSON.stringify(payload, null, 2), "application/json;charset=utf-8;");
    showToast("Backup file downloaded");
  };
  const triggerImport = () => { setImportError(null); fileRef.current && fileRef.current.click(); };
  const handleImportFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
      try {
        const parsed = JSON.parse(reader.result);
        const incoming = parsed.data || parsed;
        if (!incoming.factories || !incoming.lines || !incoming.bookings) throw new Error("File doesn't look like a capacity planner backup.");
        persist({
          factories: incoming.factories, halls: incoming.halls || [], lines: incoming.lines, bookings: incoming.bookings,
          additionalProcesses: incoming.additionalProcesses || data.additionalProcesses, settings: { ...DEFAULT_SETTINGS, ...(incoming.settings || settings) },
          counters: incoming.counters || data.counters || DEFAULT_COUNTERS,
          categories: incoming.categories || data.categories, seasons: incoming.seasons || data.seasons, customers: incoming.customers || data.customers, brands: incoming.brands || data.brands,
          savedEmailContacts: incoming.savedEmailContacts || data.savedEmailContacts
        });
        setImportError(null); showToast("Backup restored");
      } catch (err) { setImportError(err.message || "Couldn't read that file."); }
    };
    reader.readAsText(file); e.target.value = "";
  };
  const counters = data.counters || DEFAULT_COUNTERS;

  return (
    <div style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontFamily: "'Space Grotesk', sans-serif", fontWeight: 700, fontSize: 16 }}>{BRAND_NAME}</div>

      <Card title="PPC access password">
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 0 }}>Every section except Capacity Booking requires this password. It's a client-side gate suitable for keeping casual users out — not a substitute for real account-level security, since anyone with access to this browser's data could inspect it.</p>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr auto", gap: 10, alignItems: "end" }}>
          <Field label="New password"><TextInput type="password" value={newPw} onChange={e => setNewPw(e.target.value)} /></Field>
          <Field label="Confirm new password"><TextInput type="password" value={newPw2} onChange={e => setNewPw2(e.target.value)} /></Field>
          <Btn variant="primary" onClick={changePassword} style={{ marginBottom: 12 }}>Update password</Btn>
        </div>
      </Card>

      <Card title="Booking date rules">
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 0 }}>Enforced at confirmation time and cannot be bypassed from the booking form.</p>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3, 1fr)", gap: 14 }}>
          <Field label="Normal planning window (days)" hint="PCD ≥ this: no raw material selection required."><TextInput type="number" value={settings.normalPlanningDays} onChange={e => updateThreshold("normalPlanningDays", e.target.value)} /></Field>
          <Field label="Urgent threshold (days)" hint="PCD below this: All Raw Material Available is mandatory."><TextInput type="number" value={settings.urgentThresholdDays} onChange={e => updateThreshold("urgentThresholdDays", e.target.value)} /></Field>
          <Field label="Minimum PCD → Delivery lead time (days)"><TextInput type="number" value={settings.minLeadTimeDays} onChange={e => updateThreshold("minLeadTimeDays", e.target.value)} /></Field>
          <Field label="Projection advance window — min days"><TextInput type="number" value={settings.projectionAdvanceMinDays} onChange={e => updateThreshold("projectionAdvanceMinDays", e.target.value)} /></Field>
          <Field label="Projection advance window — max days"><TextInput type="number" value={settings.projectionAdvanceMaxDays} onChange={e => updateThreshold("projectionAdvanceMaxDays", e.target.value)} /></Field>
        </div>
      </Card>

      <Card title="Automatic numbering">
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 0 }}>Item Numbers (C-series, Confirmed Orders) and Projection Numbers (P-series) are generated automatically and sequentially on save — never entered manually, never reused.</p>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14 }}>
          <div style={{ background: COLORS.accentSoft, borderRadius: 5, padding: "10px 12px" }}><div style={{ fontSize: 11, color: COLORS.accentInk, textTransform: "uppercase", fontWeight: 600 }}>Next Item Number</div><div className="mono" style={{ fontSize: 16, fontWeight: 700, color: COLORS.accentInk }}>C{counters.nextItemNumber}</div></div>
          <div style={{ background: COLORS.denimSoft, borderRadius: 5, padding: "10px 12px" }}><div style={{ fontSize: 11, color: COLORS.denim, textTransform: "uppercase", fontWeight: 600 }}>Next Projection Number</div><div className="mono" style={{ fontSize: 16, fontWeight: 700, color: COLORS.denim }}>P{counters.nextProjectionNumber}</div></div>
        </div>
      </Card>

      <Card title="Working hours by factory">
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, marginTop: 0 }}>Available Minutes/Day = Working Hours × 60 per line. Standard is 7h; Globe and Cannon 1 run 8h.</p>
        <table><thead><tr><th>Factory</th><th>Hours/day</th><th>Minutes/line/day</th></tr></thead>
          <tbody>{data.factories.map(f => <tr key={f.id}><td>{f.name}</td><td className="mono">{f.workingHoursPerDay}h</td><td className="mono">{fmtNum(f.workingHoursPerDay * 60)}</td></tr>)}</tbody>
        </table>
      </Card>

      <Card title="Seasons"><TagListEditor items={data.seasons} onChange={updateSeasons} placeholder="Add a season e.g. SS29" /></Card>
      <Card title="Product categories"><TagListEditor items={data.categories} onChange={updateCategories} placeholder="Add a category" /></Card>
      <Card title="Customers"><TagListEditor items={data.customers} onChange={updateCustomers} placeholder="Add a customer" /></Card>
      <Card title="Brands"><TagListEditor items={data.brands} onChange={updateBrands} placeholder="Add a brand" /></Card>
      <Card title="Saved email contacts"><TagListEditor items={data.savedEmailContacts} onChange={updateContacts} placeholder="name@company.com" /></Card>

      <Card title="Backup & retrieve your data">
        <p style={{ fontSize: 13, color: COLORS.inkSoft, marginTop: 0, lineHeight: 1.7 }}>Everything saves automatically. Download a backup file for anything you need to retrieve with certainty later.</p>
        <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
          <Btn variant="primary" onClick={exportBackup}>Download backup (.json)</Btn>
          <Btn onClick={triggerImport}>Restore from backup</Btn>
          <input ref={fileRef} type="file" accept="application/json" onChange={handleImportFile} style={{ display: "none" }} />
        </div>
        {importError && <div style={{ marginTop: 10, fontSize: 12.5, color: COLORS.danger }}>{importError}</div>}
      </Card>

      <Card title="Data management"><div style={{ display: "flex", gap: 10 }}><Btn onClick={reset}>Restore sample data</Btn><Btn variant="danger" onClick={clear}>Clear all factories & bookings</Btn></div></Card>

      <Card title={`About ${BRAND_NAME}`}>
        <p style={{ fontSize: 12.5, color: COLORS.inkSoft, lineHeight: 1.8, margin: 0 }}>
          Item Numbers (C-series) and Projection Numbers (P-series) are issued automatically and sequentially on save, never manually, never reused. SAM replaces SMV throughout. Required production minutes
          adjust for Budgeted Efficiency % (Qty × SAM ÷ Efficiency%). A booking needs PCD ≥ {settings.normalPlanningDays} days for a fully open planning window; between {settings.urgentThresholdDays} and {settings.normalPlanningDays} days
          is informational only; below {settings.urgentThresholdDays} days requires "All Raw Material Available" or the booking is held. PCD to Delivery must be at least {settings.minLeadTimeDays} days. Projections should convert
          to Confirmed Orders before {settings.normalPlanningDays} days from PCD — reminders escalate at 120/105/90 days. Every section except Capacity Booking is behind the PPC password.
        </p>
      </Card>
    </div>
  );
}
