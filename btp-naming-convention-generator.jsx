import { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { Copy, Check, Download, Plus, Trash2, ChevronDown, ChevronRight, Sparkles } from "lucide-react";

/* ----------------------------------------------------------------------------
   CONVENTION CATALOG
   Encoded from "HSE IFMS BTP Naming Convention" V1.0 — all four tabs.
   Prefix tokens: {PU} upper-joined · {PT} title-spaced · {PL} lower-joined · {PS} short
   Everything in <angle brackets> is a variable (env-derived or user-supplied).
---------------------------------------------------------------------------- */
const CATALOG = [
  // ---- BTP Admin --------------------------------------------------------
  ["BTP Admin", "Subaccounts & Connectivity", "Subaccount (BTP Application)", "{PT} <Env> <AppName>", "Display name; Title Case with spaces."],
  ["BTP Admin", "Subaccounts & Connectivity", "Service Instance", "<appname>-<Env>-instance", "One instance per app, per environment."],
  ["BTP Admin", "Subaccounts & Connectivity", "Cloud Connector Connection", "{PS}<env><backend application name>.<system id>", "Backend connection per environment."],

  // ---- BTP Development ---------------------------------------------------
  ["BTP Development", "Platform & Connectivity", "S/4HANA System (SID)", "<SID>", "Sandbox S41 · Dev S42 · Test S43 · Pre-Prod S44 · Prod S49."],
  ["BTP Development", "Platform & Connectivity", "BTP Tenant ID (TID)", "<TID>", "DEV · QA · PRD."],
  ["BTP Development", "Platform & Connectivity", "Cloud Foundry Space", "<tid>", "Space = environment; one per subaccount."],
  ["BTP Development", "Platform & Connectivity", "Destination (BTP)", "{PU}_<TID>_<SID>_<Backend>", "No env in the name — the destination already lives in an env-specific subaccount. Referenced by code."],
  ["BTP Development", "Platform & Connectivity", "MCP Server Destination (Joule)", "{PU}_<TID>_MCP_<Tool>", "Exposes an external tool to a Joule agent."],
  ["BTP Development", "Platform & Connectivity", "Role Collection", "{PU}_<UseCase>_<Role>", "Maps to a business role."],
  ["BTP Development", "Platform & Connectivity", "Git Repository", "{PL}-<tid>-<usecase>-<component>", "Lowercase, hyphenated; one repo per component."],
  ["BTP Development", "Platform & Connectivity", "MTA Archive / mta.yaml ID", "{PL}.<usecase>.<component>", "Dot-separated reverse-style ID."],

  ["BTP Development", "SAP Build Process Automation", "Business Process Project", "{PT} <UseCase> <Purpose>", "Display name; Title Case with spaces."],
  ["BTP Development", "SAP Build Process Automation", "Process", "<Verb><Object>", "Start with a verb; one responsibility."],
  ["BTP Development", "SAP Build Process Automation", "Automation (bot)", "Auto_<Verb><Object>", "RPA automation artifact."],
  ["BTP Development", "SAP Build Process Automation", "Action Step", "Act_<Backend>_<Operation>", "External API call artifact."],
  ["BTP Development", "SAP Build Process Automation", "Decision Step", "Dec_<Rule>", "Holds branching business rules."],
  ["BTP Development", "SAP Build Process Automation", "Mail / Notification Step", "Mail_<Purpose>", "Outbound email / notification (use Notif_ for notifications)."],
  ["BTP Development", "SAP Build Process Automation", "Condition (gateway)", "Is<Condition>", "Boolean branch label."],
  ["BTP Development", "SAP Build Process Automation", "Form", "Frm_<Purpose>", "Interactive form artifact."],
  ["BTP Development", "SAP Build Process Automation", "Data Type", "Dt_<Entity>", "Structured data type."],
  ["BTP Development", "SAP Build Process Automation", "Process Variable", "camelCase", "No prefix; innermost scope."],
  ["BTP Development", "SAP Build Process Automation", "Environment Variable", "UPPER_CASE", "Never hardcode the value in logic."],
  ["BTP Development", "SAP Build Process Automation", "Rule", "<Rule>Rule", "Individual rule expression."],

  ["BTP Development", "SAP Joule / Joule Studio", "Joule Studio Project", "{PT} <UseCase> Joule", "SAP Build project of type 'Joule agent and skill'."],
  ["BTP Development", "SAP Joule / Joule Studio", "Action Project", "<UseCase>_<Backend>_Actions", "Defines which backend endpoints Joule may call."],
  ["BTP Development", "SAP Joule / Joule Studio", "Action (endpoint)", "<verb><Object>", "One action per endpoint / operation."],
  ["BTP Development", "SAP Joule / Joule Studio", "Joule Skill", "<UseCase>_<Verb><Object>", "Skill descriptions must be precise and distinct, or Joule triggers unpredictably."],
  ["BTP Development", "SAP Joule / Joule Studio", "Joule Agent", "{PT} <UseCase> Agent", "Orchestrates skills; plans and reasons."],
  ["BTP Development", "SAP Joule / Joule Studio", "Data Type (Joule)", "Dt_<Entity>", "Shared input / output structure."],
  ["BTP Development", "SAP Joule / Joule Studio", "Capability / DTA File", "<usecase>-<capability>", "Design-time capability file (CLI only)."],
  ["BTP Development", "SAP Joule / Joule Studio", "Document Grounding Tool", "<UseCase>_Grounding_<Source>", "Grounds agent answers on a document source."],

  ["BTP Development", "SAP Build Apps (low-code UI)", "Build Apps Project / App", "{PT} <UseCase> App", "Display name; Title Case."],
  ["BTP Development", "SAP Build Apps (low-code UI)", "Page", "<Purpose>Page", "PascalCase; describes the screen."],
  ["BTP Development", "SAP Build Apps (low-code UI)", "App / Page Variable", "camelCase", "Session / page scope."],
  ["BTP Development", "SAP Build Apps (low-code UI)", "Data Resource (OData/REST)", "<Backend>_<Entity>", "Bound data source."],

  ["BTP Development", "SAP HANA Cloud", "HANA Cloud Instance", "{PL}-hana-<env>", "Environment-bound instance."],
  ["BTP Development", "SAP HANA Cloud", "HDI Container", "{PL}-<usecase>-hdi", "One container per deployable schema."],
  ["BTP Development", "SAP HANA Cloud", "Schema", "{PU}_<USECASE>", "Upper case; groups objects."],
  ["BTP Development", "SAP HANA Cloud", "Table", "<AREA>_<ENTITY>", "Upper case."],
  ["BTP Development", "SAP HANA Cloud", "Calculation / DB View", "CV_<Purpose>", "Prefix CV_ for views."],
  ["BTP Development", "SAP HANA Cloud", "Fuzzy Artifact", "<ENTITY>_FUZZY", "Marks a fuzzy-search object."],
  ["BTP Development", "SAP HANA Cloud", "Vector Artifact", "<ENTITY>_VECTOR", "Marks a vector-search object."],

  // ---- CAP & RAP ---------------------------------------------------------
  ["CAP & RAP", "CAP · Project & Data Model", "CAP Project / MTA ID", "{PL}-<usecase>", "Root project folder / mta.yaml id."],
  ["CAP & RAP", "CAP · Project & Data Model", "CDS Namespace", "{PL}.<usecase>", "Declared at the top of the .cds file."],
  ["CAP & RAP", "CAP · Project & Data Model", "Data Model File", "db/schema.cds", "Persistent entities."],
  ["CAP & RAP", "CAP · Project & Data Model", "Entity", "<Entity>", "PascalCase, singular."],
  ["CAP & RAP", "CAP · Project & Data Model", "Entity Set (exposed)", "<Entity>s", "Plural collection name."],
  ["CAP & RAP", "CAP · Project & Data Model", "Element / Field", "camelCase", "Consistent across the model."],
  ["CAP & RAP", "CAP · Project & Data Model", "Association / Composition", "camelCase", "Managed associations (items / toSupplier)."],

  ["CAP & RAP", "CAP · Service Layer", "Service Definition File", "srv/<usecase>-service.cds", "One file per service."],
  ["CAP & RAP", "CAP · Service Layer", "CDS Service", "<UseCase>Service", "PascalCase + 'Service' suffix."],
  ["CAP & RAP", "CAP · Service Layer", "Projection", "<Entity>s", "Exposed view of an entity (plural business name)."],
  ["CAP & RAP", "CAP · Service Layer", "Action", "camelCase", "Entity / service operation."],
  ["CAP & RAP", "CAP · Service Layer", "Function", "camelCase", "Read-only operation."],
  ["CAP & RAP", "CAP · Service Layer", "Event", "camelCase", "Emitted domain event."],
  ["CAP & RAP", "CAP · Service Layer", "Custom Handler", "srv/<usecase>-service.js", "Node / TypeScript logic."],
  ["CAP & RAP", "CAP · Service Layer", "Fiori Annotations", "app/<usecase>/annotations.cds", "UI annotations, kept separate."],

  ["CAP & RAP", "CAP · Runtime & Deploy", "Exposed OData Path", "/odata/v4/<Service>", "Default OData V4 route."],
  ["CAP & RAP", "CAP · Runtime & Deploy", "CF App (service module)", "{PL}-<usecase>-srv", "Deployed srv module."],
  ["CAP & RAP", "CAP · Runtime & Deploy", "HDI / DB Module", "{PL}-<usecase>-db", "Deployed db module."],

  ["CAP & RAP", "RAP · Data Model (CDS)", "Database Table", "Z<Bo>", "Persistent transparent table."],
  ["CAP & RAP", "RAP · Data Model (CDS)", "Root / Interface View", "ZR_<Bo>", "Root view of the BO; add TP suffix if transactional."],
  ["CAP & RAP", "RAP · Data Model (CDS)", "Basic Interface View (VDM)", "ZI_<Bo>", "Reusable read model."],
  ["CAP & RAP", "RAP · Data Model (CDS)", "Extension Include View", "ZE_<Bo>", "Extension include."],

  ["CAP & RAP", "RAP · Business Service", "Service Definition", "ZUI_<Bo>", "Container of entities to expose (ZSD_ for non-UI)."],
  ["CAP & RAP", "RAP · Business Service", "Service Binding (OData V4, UI)", "ZUI_<Bo>_O4", "Suffix O4 = OData V4."],
  ["CAP & RAP", "RAP · Business Service", "Service Binding (Web API)", "ZAPI_<Bo>_O4", "For external / API consumption."],

  // ---- UiPath ------------------------------------------------------------
  ["UiPath", "UiPath", "Tenant", "{PU}_UiPath_<TID>", "Normalized order (legacy sheet used the reversed IFMSHSE_ form). TID = DEV, QA, PRD."],
  ["UiPath", "UiPath", "Process Package — Dispatcher", "<ShortProcessName>Automation_Dispatcher", "Queue-loading logic."],
  ["UiPath", "UiPath", "Process Package — Performer", "<ShortProcessName>Automation_Performer", "Transaction-processing logic."],
  ["UiPath", "UiPath", "Process Name — Dispatcher", "<Process Name> Automation Dispatcher", "Readable name in Orchestrator."],
  ["UiPath", "UiPath", "Process Name — Performer", "<Process Name> Automation Performer", "Readable name in Orchestrator."],
  ["UiPath", "UiPath", "Workflow File", "<Verb><Object>.xaml", "Name describes what the workflow does."],
  ["UiPath", "UiPath", "Dispatcher Workflow", "<Verb><Object>.xaml", "Reads input and adds items to the queue."],
  ["UiPath", "UiPath", "Performer Workflow", "<Verb><Object>.xaml", "Processes one transaction."],
  ["UiPath", "UiPath", "Variable", "camelCase", "Meaningful name; avoid temp / data / value."],
  ["UiPath", "UiPath", "DataTable Variable", "dt_<MeaningfulName>", "Must start with dt_."],
  ["UiPath", "UiPath", "In Argument", "in_<MeaningfulName>", "Values received by a workflow."],
  ["UiPath", "UiPath", "Out Argument", "out_<MeaningfulName>", "Values returned from a workflow."],
  ["UiPath", "UiPath", "In/Out Argument", "io_<MeaningfulName>", "Passed in and updated."],
  ["UiPath", "UiPath", "DataTable Argument", "<direction>_dt_<MeaningfulName>", "Direction prefix + dt_."],
  ["UiPath", "UiPath", "Queue", "<ShortProcessName>Queue", "Short, meaningful queue name."],
  ["UiPath", "UiPath", "Credential Asset", "<ShortProcessName>_<Purpose>", "Purpose-based asset name."],
  ["UiPath", "UiPath", "Config Asset", "<ShortProcessName>_<Purpose>", "Process-level configuration asset."],
  ["UiPath", "UiPath", "URL Asset", "<ShortProcessName>_<Purpose>", "Application / portal URL asset."],
].map(([tab, family, name, pattern, notes]) => ({
  id: `${tab}::${family}::${name}`, tab, family, name, pattern, notes,
}));

/* ---- Variable metadata ----------------------------------------------- */
const ENV_KEYS = { env: "env", sid: "sid", tid: "tid", systemid: "sid" };
const DEFAULTS = {
  usecase: "Procurement", appname: "Build Process Automation", area: "MAT", backend: "S4",
  entity: "PurchaseRequisition", verb: "Create", object: "PurchaseRequisition", purpose: "PRApproval",
  operation: "CreatePR", rule: "ApprovalThreshold", condition: "AboveThreshold", tool: "VendorCatalog",
  role: "Buyer", component: "bpa", capability: "pr-create", source: "Policies", bo: "IfmsPR",
  service: "ProcurementService", shortprocessname: "InvoiceValidation", processname: "Invoice Validation",
  meaningfulname: "InvoiceData", backendapplicationname: "s4", direction: "in",
};
const LABELS = {
  usecase: "Use Case", appname: "App Name", backendapplicationname: "Backend App Name",
  shortprocessname: "Short Process Name", processname: "Process Name", meaningfulname: "Meaningful Name",
  bo: "Business Object (BO)", service: "CDS Service",
};
const prettyLabel = (k) => LABELS[k] || k.replace(/(^|\s)\S/g, (c) => c.toUpperCase());

/* ---- Engine ---------------------------------------------------------- */
const normKey = (inner) => inner.toLowerCase().replace(/\s+/g, "");
const placeholderKeys = (pattern) => [...pattern.matchAll(/<([^>]+)>/g)].map((m) => normKey(m[1]));
const hasEnv = (pattern) => placeholderKeys(pattern).some((k) => k in ENV_KEYS);

function applyCase(inner, value) {
  if (value == null || value === "") return value ?? "";
  const isUpper = inner === inner.toUpperCase() && inner !== inner.toLowerCase();
  const isLower = inner === inner.toLowerCase();
  if (isUpper) return value.toUpperCase().replace(/\s+/g, "_");
  if (isLower) return value.toLowerCase().replace(/\s+/g, "-");
  return value; // mixed / Pascal → keep as entered (preserves display-name spacing)
}

function substitute(pattern, prefix, vars, envRow, keepPlaceholders) {
  let s = pattern
    .replaceAll("{PU}", prefix.pu).replaceAll("{PT}", prefix.pt)
    .replaceAll("{PL}", prefix.pl).replaceAll("{PS}", prefix.ps);
  if (keepPlaceholders) return s;
  return s.replace(/<([^>]+)>/g, (m, inner) => {
    const key = normKey(inner);
    let val;
    if (envRow && key in ENV_KEYS) val = envRow[ENV_KEYS[key]];
    else val = vars[key] ?? DEFAULTS[key];
    if (val == null) return m; // unknown, unresolvable → leave placeholder visible
    return applyCase(inner, val);
  });
}

/* ---- Copy helper (works inside sandboxed iframe) --------------------- */
function copyText(t) {
  try {
    if (navigator.clipboard && navigator.clipboard.writeText) { navigator.clipboard.writeText(t); return; }
  } catch (e) { /* fall through */ }
  const ta = document.createElement("textarea");
  ta.value = t; ta.style.position = "fixed"; ta.style.opacity = "0";
  document.body.appendChild(ta); ta.focus(); ta.select();
  try { document.execCommand("copy"); } catch (e) { /* noop */ }
  document.body.removeChild(ta);
}

/* ====================================================================== */
export default function App() {
  const [projectInput, setProjectInput] = useState("HSE IFMS");
  const [shortForm, setShortForm] = useState("");
  const [envRows, setEnvRows] = useState([
    { env: "Dev", sid: "S42", tid: "DEV" },
    { env: "Test", sid: "S43", tid: "QA" },
    { env: "Pre-Prod", sid: "S44", tid: "PRD" },
    { env: "Prod", sid: "S49", tid: "PRD" },
  ]);
  const [selected, setSelected] = useState(() => new Set());
  const [mode, setMode] = useState("scaffold");
  const [varValues, setVarValues] = useState({});
  const [collapsed, setCollapsed] = useState(() => new Set());
  const [copied, setCopied] = useState("");

  /* --- prefix engine --- */
  const prefix = useMemo(() => {
    const raw = projectInput.trim();
    const tokens = raw
      .replace(/([a-z0-9])([A-Z])/g, "$1 $2")
      .split(/[\s\-_]+/).filter(Boolean);
    const joined = tokens.join("");
    const derivedShort = tokens.length ? tokens[tokens.length - 1].toLowerCase() : "";
    return {
      pt: raw || "PROJECT",
      pu: (joined || "PROJECT").toUpperCase(),
      pl: (joined || "project").toLowerCase(),
      ps: (shortForm.trim() || derivedShort || "prj"),
      derivedShort,
      tokens,
    };
  }, [projectInput, shortForm]);

  /* --- grouped catalog --- */
  const grouped = useMemo(() => {
    const tabs = [];
    for (const a of CATALOG) {
      let t = tabs.find((x) => x.tab === a.tab);
      if (!t) { t = { tab: a.tab, families: [] }; tabs.push(t); }
      let f = t.families.find((x) => x.family === a.family);
      if (!f) { f = { family: a.family, items: [] }; t.families.push(f); }
      f.items.push(a);
    }
    return tabs;
  }, []);

  /* --- required user variables for current selection --- */
  const requiredKeys = useMemo(() => {
    const keys = new Set();
    for (const a of CATALOG) {
      if (!selected.has(a.id)) continue;
      for (const k of placeholderKeys(a.pattern)) if (!(k in ENV_KEYS)) keys.add(k);
    }
    return [...keys].sort();
  }, [selected]);

  /* --- generated output --- */
  const generated = useMemo(() => {
    const vars = mode === "resolve" ? varValues : {};
    return CATALOG.filter((a) => selected.has(a.id)).map((a) => {
      const convention = substitute(a.pattern, prefix, {}, null, true);
      let rows;
      if (hasEnv(a.pattern)) {
        const map = new Map();
        for (const er of envRows) {
          const res = substitute(a.pattern, prefix, vars, er, false);
          if (!map.has(res)) map.set(res, []);
          map.get(res).push(er.env);
        }
        rows = [...map.entries()].map(([result, envs]) => ({ env: envs.join(", "), result }));
      } else {
        rows = [{ env: "", result: substitute(a.pattern, prefix, vars, null, false) }];
      }
      return { ...a, convention, rows };
    });
  }, [selected, prefix, envRows, mode, varValues]);

  const generatedByTab = useMemo(() => {
    const tabs = [];
    for (const g of generated) {
      let t = tabs.find((x) => x.tab === g.tab);
      if (!t) { t = { tab: g.tab, families: [] }; tabs.push(t); }
      let f = t.families.find((x) => x.family === g.family);
      if (!f) { f = { family: g.family, items: [] }; t.families.push(f); }
      f.items.push(g);
    }
    return tabs;
  }, [generated]);

  /* --- selection helpers --- */
  const toggleArtifact = (id) => {
    setSelected((prev) => { const n = new Set(prev); n.has(id) ? n.delete(id) : n.add(id); return n; });
  };
  const familyState = (items) => {
    const on = items.filter((i) => selected.has(i.id)).length;
    return on === 0 ? "none" : on === items.length ? "all" : "some";
  };
  const toggleFamily = (items) => {
    const st = familyState(items);
    setSelected((prev) => {
      const n = new Set(prev);
      if (st === "all") items.forEach((i) => n.delete(i.id));
      else items.forEach((i) => n.add(i.id));
      return n;
    });
  };
  const selectAll = () => setSelected(new Set(CATALOG.map((a) => a.id)));
  const clearAll = () => setSelected(new Set());
  const doToggleCollapse = (key) =>
    setCollapsed((prev) => { const n = new Set(prev); n.has(key) ? n.delete(key) : n.add(key); return n; });

  const flash = (id) => { setCopied(id); setTimeout(() => setCopied(""), 1100); };

  /* --- xlsx export --- */
  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    const cover = [
      ["SAP BTP Naming Convention"],
      ["Generated", new Date().toLocaleString()],
      ["Catalog", "HSE IFMS BTP Naming Convention (V1.0)"],
      ["Mode", mode === "resolve" ? "Resolved names" : "Scaffold (with examples)"],
      [],
      ["Project prefix forms"],
      ["Title (display)", prefix.pt],
      ["Upper (joined)", prefix.pu],
      ["Lower (joined)", prefix.pl],
      ["Short", prefix.ps],
      [],
      ["Environment map"],
      ["Environment", "SID", "TID"],
      ...envRows.map((r) => [r.env, r.sid, r.tid]),
    ];
    const cws = XLSX.utils.aoa_to_sheet(cover);
    cws["!cols"] = [{ wch: 20 }, { wch: 42 }, { wch: 14 }];
    XLSX.utils.book_append_sheet(wb, cws, "Project");

    for (const t of generatedByTab) {
      const aoa = [["Family / Group", "Artifact", "Environment", "Naming Convention", "Result", "Notes"]];
      for (const f of t.families) {
        for (const it of f.items) {
          it.rows.forEach((r, idx) => {
            aoa.push([
              idx === 0 ? f.family : "",
              idx === 0 ? it.name : "",
              r.env,
              idx === 0 ? it.convention : "",
              r.result,
              idx === 0 ? it.notes : "",
            ]);
          });
        }
      }
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 26 }, { wch: 30 }, { wch: 16 }, { wch: 34 }, { wch: 40 }, { wch: 44 }];
      const safe = t.tab.replace(/[\\/*?:\[\]]/g, "").slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, safe);
    }

    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${prefix.pu}_BTP_Naming_Convention.xlsx`;
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  };

  const copyAll = () => {
    const lines = [];
    for (const g of generated) g.rows.forEach((r) => lines.push(r.result));
    copyText(lines.join("\n")); flash("__all__");
  };

  const totalNames = generated.reduce((s, g) => s + g.rows.length, 0);

  /* --------------------------------------------------------------------- */
  return (
    <div className="min-h-screen bg-slate-100 text-slate-800 font-sans">
      <div className="max-w-6xl mx-auto px-4 py-6">
        {/* Header */}
        <header className="mb-6">
          <div className="flex items-center gap-2 text-indigo-700">
            <div className="h-7 w-7 rounded-md bg-indigo-600 text-white grid place-items-center font-bold text-sm">N</div>
            <h1 className="text-xl font-semibold tracking-tight text-slate-900">BTP Naming Convention Generator</h1>
          </div>
          <p className="text-sm text-slate-500 mt-1">
            Pick a project prefix and the artifacts you need — get a governed naming sheet.
            Catalog: HSE IFMS V1.0.
          </p>
        </header>

        <div className="grid grid-cols-1 lg:grid-cols-3 gap-5">
          {/* LEFT: setup + selection */}
          <div className="lg:col-span-1 space-y-5">
            {/* Project */}
            <section className="bg-white rounded-xl border border-slate-200 p-4">
              <h2 className="text-xs font-semibold uppercase tracking-wide text-slate-400 mb-3">Project</h2>
              <label className="block text-sm font-medium text-slate-700 mb-1">Project name or prefix</label>
              <input
                value={projectInput}
                onChange={(e) => setProjectInput(e.target.value)}
                placeholder="e.g. HSE IFMS"
                className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500"
              />
              <div className="grid grid-cols-2 gap-2 mt-3">
                <PrefixChip label="Upper" value={prefix.pu} />
                <PrefixChip label="Title" value={prefix.pt} />
                <PrefixChip label="Lower" value={prefix.pl} />
                <div>
                  <div className="text-[10px] uppercase tracking-wide text-slate-400 mb-0.5">Short</div>
                  <input
                    value={shortForm}
                    onChange={(e) => setShortForm(e.target.value)}
                    placeholder={prefix.derivedShort || "short"}
                    className="w-full rounded-md border border-slate-300 px-2 py-1 text-sm font-mono text-slate-800 focus:outline-none focus:ring-2 focus:ring-indigo-500"
                  />
                </div>
              </div>
              <p className="text-[11px] text-slate-400 mt-2">
                Short defaults to the last word, lowercased. Override it if your team uses a set abbreviation.
              </p>
            </section>

            {/* Environments */}
            <section className="bg-white rounded-xl border border-slate-200 p-4">
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-xs font-semibold uppercase tracking-wide text-slate-400">Environments</h2>
                <button
                  onClick={() => setEnvRows((r) => [...r, { env: "New", sid: "S00", tid: "DEV" }])}
                  className="inline-flex items-center gap-1 text-xs text-indigo-600 hover:text-indigo-800"
                >
                  <Plus size={14} /> Add
                </button>
              </div>
              <div className="grid grid-cols-12 gap-2 text-[10px] uppercase tracking-wide text-slate-400 px-1 mb-1">
                <div className="col-span-5">Env</div><div className="col-span-3">SID</div><div className="col-span-3">TID</div><div className="col-span-1" />
              </div>
              <div className="space-y-1.5">
                {envRows.map((r, i) => (
                  <div key={i} className="grid grid-cols-12 gap-2 items-center">
                    <input className="col-span-5 rounded-md border border-slate-300 px-2 py-1 text-sm" value={r.env}
                      onChange={(e) => setEnvRows((rows) => rows.map((x, j) => j === i ? { ...x, env: e.target.value } : x))} />
                    <input className="col-span-3 rounded-md border border-slate-300 px-2 py-1 text-sm font-mono" value={r.sid}
                      onChange={(e) => setEnvRows((rows) => rows.map((x, j) => j === i ? { ...x, sid: e.target.value } : x))} />
                    <input className="col-span-3 rounded-md border border-slate-300 px-2 py-1 text-sm font-mono" value={r.tid}
                      onChange={(e) => setEnvRows((rows) => rows.map((x, j) => j === i ? { ...x, tid: e.target.value } : x))} />
                    <button className="col-span-1 grid place-items-center text-slate-400 hover:text-rose-600"
                      onClick={() => setEnvRows((rows) => rows.filter((_, j) => j !== i))}>
                      <Trash2 size={14} />
                    </button>
                  </div>
                ))}
              </div>
            </section>

            {/* Selection */}
            <section className="bg-white rounded-xl border border-slate-200 p-4">
              <div className="flex items-center justify-between mb-3">
                <h2 className="text-xs font-semibold uppercase tracking-wide text-slate-400">Artifacts</h2>
                <div className="flex gap-3 text-xs">
                  <button onClick={selectAll} className="text-indigo-600 hover:text-indigo-800">All</button>
                  <button onClick={clearAll} className="text-slate-400 hover:text-slate-600">Clear</button>
                </div>
              </div>
              <div className="space-y-3 max-h-[28rem] overflow-auto pr-1">
                {grouped.map((t) => (
                  <div key={t.tab}>
                    <div className="text-[11px] font-semibold text-slate-500 uppercase tracking-wide mb-1">{t.tab}</div>
                    {t.families.map((f) => {
                      const st = familyState(f.items);
                      const ck = `${t.tab}/${f.family}`;
                      const isColl = collapsed.has(ck);
                      return (
                        <div key={f.family} className="mb-1.5 rounded-lg border border-slate-150 bg-slate-50">
                          <div className="flex items-center gap-2 px-2 py-1.5">
                            <TriBox state={st} onClick={() => toggleFamily(f.items)} />
                            <button className="flex-1 text-left text-sm font-medium text-slate-700 truncate" onClick={() => toggleFamily(f.items)}>
                              {f.family}
                            </button>
                            <span className="text-[10px] text-slate-400">
                              {f.items.filter((i) => selected.has(i.id)).length}/{f.items.length}
                            </span>
                            <button onClick={() => doToggleCollapse(ck)} className="text-slate-400 hover:text-slate-600">
                              {isColl ? <ChevronRight size={15} /> : <ChevronDown size={15} />}
                            </button>
                          </div>
                          {!isColl && (
                            <div className="px-2 pb-2 space-y-0.5">
                              {f.items.map((it) => (
                                <label key={it.id} className="flex items-center gap-2 pl-5 py-0.5 cursor-pointer group">
                                  <input type="checkbox" checked={selected.has(it.id)} onChange={() => toggleArtifact(it.id)}
                                    className="h-3.5 w-3.5 rounded border-slate-300 text-indigo-600 focus:ring-indigo-500" />
                                  <span className="text-[13px] text-slate-600 group-hover:text-slate-900 truncate">{it.name}</span>
                                </label>
                              ))}
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            </section>
          </div>

          {/* RIGHT: mode + output */}
          <div className="lg:col-span-2 space-y-5">
            {/* Toolbar */}
            <div className="bg-white rounded-xl border border-slate-200 p-3 flex flex-wrap items-center gap-3 sticky top-3 z-10">
              <div className="inline-flex rounded-lg bg-slate-100 p-0.5">
                <ModeBtn active={mode === "scaffold"} onClick={() => setMode("scaffold")}>Scaffold</ModeBtn>
                <ModeBtn active={mode === "resolve"} onClick={() => setMode("resolve")}>Resolve</ModeBtn>
              </div>
              <div className="text-sm text-slate-500">
                {selected.size} artifact{selected.size === 1 ? "" : "s"} · {totalNames} name{totalNames === 1 ? "" : "s"}
              </div>
              <div className="ml-auto flex gap-2">
                <button onClick={copyAll} disabled={!totalNames}
                  className="inline-flex items-center gap-1.5 rounded-lg border border-slate-300 px-3 py-1.5 text-sm text-slate-700 hover:bg-slate-50 disabled:opacity-40">
                  {copied === "__all__" ? <Check size={15} className="text-emerald-600" /> : <Copy size={15} />} Copy all
                </button>
                <button onClick={exportXlsx} disabled={!totalNames}
                  className="inline-flex items-center gap-1.5 rounded-lg bg-indigo-600 px-3 py-1.5 text-sm text-white hover:bg-indigo-700 disabled:opacity-40">
                  <Download size={15} /> Download .xlsx
                </button>
              </div>
            </div>

            {/* Resolve inputs */}
            {mode === "resolve" && selected.size > 0 && (
              <section className="bg-white rounded-xl border border-slate-200 p-4">
                <div className="flex items-center gap-1.5 mb-3 text-slate-500">
                  <Sparkles size={15} className="text-indigo-500" />
                  <h2 className="text-xs font-semibold uppercase tracking-wide">Values for your selection</h2>
                </div>
                {requiredKeys.length === 0 ? (
                  <p className="text-sm text-slate-400">Your selection needs no extra values — names are fully derived from the prefix and environments.</p>
                ) : (
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-3">
                    {requiredKeys.map((k) => (
                      <div key={k}>
                        <label className="block text-[11px] font-medium text-slate-500 mb-0.5">{prettyLabel(k)}</label>
                        <input
                          value={varValues[k] ?? DEFAULTS[k] ?? ""}
                          onChange={(e) => setVarValues((v) => ({ ...v, [k]: e.target.value }))}
                          className="w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-indigo-500"
                        />
                      </div>
                    ))}
                  </div>
                )}
              </section>
            )}

            {/* Output */}
            {selected.size === 0 ? (
              <div className="bg-white rounded-xl border border-dashed border-slate-300 p-10 text-center text-slate-400">
                Select artifacts on the left to generate the naming sheet.
              </div>
            ) : (
              generatedByTab.map((t) => (
                <section key={t.tab} className="bg-white rounded-xl border border-slate-200 overflow-hidden">
                  <div className="px-4 py-2.5 bg-slate-800 text-white text-sm font-semibold">{t.tab}</div>
                  {t.families.map((f) => (
                    <div key={f.family}>
                      <div className="px-4 py-1.5 bg-slate-50 border-y border-slate-150 text-[11px] font-semibold uppercase tracking-wide text-slate-500">
                        {f.family}
                      </div>
                      <div className="divide-y divide-slate-100">
                        {f.items.map((it) => (
                          <div key={it.id} className="px-4 py-3">
                            <div className="flex items-baseline justify-between gap-3">
                              <div className="text-sm font-medium text-slate-800">{it.name}</div>
                              <code className="text-[11px] text-slate-400 font-mono truncate">{it.convention}</code>
                            </div>
                            <div className="mt-1.5 space-y-1">
                              {it.rows.map((r, i) => (
                                <div key={i} className="flex items-center gap-2">
                                  {r.env && (
                                    <span className="shrink-0 text-[10px] font-medium text-indigo-700 bg-indigo-50 rounded px-1.5 py-0.5 w-24 text-center truncate">
                                      {r.env}
                                    </span>
                                  )}
                                  <code className="flex-1 text-[13px] font-mono text-slate-900 bg-slate-50 rounded px-2 py-1 truncate">
                                    {r.result}
                                  </code>
                                  <button
                                    onClick={() => { copyText(r.result); flash(it.id + i); }}
                                    className="shrink-0 grid place-items-center h-6 w-6 rounded text-slate-400 hover:text-indigo-600 hover:bg-slate-100"
                                  >
                                    {copied === it.id + i ? <Check size={14} className="text-emerald-600" /> : <Copy size={14} />}
                                  </button>
                                </div>
                              ))}
                            </div>
                            {it.notes && <p className="mt-1 text-[11px] text-slate-400">{it.notes}</p>}
                          </div>
                        ))}
                      </div>
                    </div>
                  ))}
                </section>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---- small components ------------------------------------------------ */
function PrefixChip({ label, value }) {
  return (
    <div>
      <div className="text-[10px] uppercase tracking-wide text-slate-400 mb-0.5">{label}</div>
      <div className="rounded-md bg-slate-50 border border-slate-200 px-2 py-1 text-sm font-mono text-slate-800 truncate">{value}</div>
    </div>
  );
}
function TriBox({ state, onClick }) {
  const base = "h-4 w-4 rounded border grid place-items-center cursor-pointer shrink-0";
  if (state === "all") return <div onClick={onClick} className={base + " bg-indigo-600 border-indigo-600 text-white"}><Check size={11} /></div>;
  if (state === "some") return <div onClick={onClick} className={base + " bg-indigo-100 border-indigo-400"}><div className="h-1.5 w-1.5 bg-indigo-600 rounded-sm" /></div>;
  return <div onClick={onClick} className={base + " bg-white border-slate-300"} />;
}
function ModeBtn({ active, onClick, children }) {
  return (
    <button onClick={onClick}
      className={"px-3 py-1.5 text-sm rounded-md font-medium " + (active ? "bg-white text-indigo-700 shadow-sm" : "text-slate-500 hover:text-slate-700")}>
      {children}
    </button>
  );
}
