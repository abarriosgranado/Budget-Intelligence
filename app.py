from __future__ import annotations

import json
from dataclasses import dataclass
from io import BytesIO
from datetime import datetime
from pathlib import Path
from typing import Iterable

import altair as alt
import pandas as pd
import streamlit as st
from xlsxwriter.utility import xl_col_to_name


MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
MONTH_TO_NUMBER = {month: index + 1 for index, month in enumerate(MONTHS)}
MONTH_NUMBER_TO_NAME = {number: month for month, number in MONTH_TO_NUMBER.items()}
LANGUAGE_OPTIONS = ["Español", "English"]
TRANSLATIONS = {
    "Español": {
        "language": "Idioma",
        "nav_label": "Panel de navegacion",
        "step": "Paso",
        "settings": "Settings",
        "admin_setup": "Admin setup",
        "evidence": "Evidencia",
        "budget_actuals": "Budget y actuals",
        "workspaces": "Workspaces",
        "by_department": "Por departamento",
        "proposal": "Propuesta",
        "challenge": "Challenge",
        "submitted_review": "Revision submit",
        "new_budget": "Nuevo budget",
        "challenge_title": "4. Challenge Budget Intelligence",
        "proposal_title": "5. Propuesta",
        "challenge_intro": "Budget Intelligence revisa los workspaces enviados y reta cada propuesta contra el budget aprobado anterior y los actuals YTD.",
        "no_submitted_departments": "Todavia no hay departamentos enviados. En Workspaces, pulsa Submit para que entren en esta revision.",
        "submitted_departments": "Departamentos enviados",
        "requested_plan": "Plan solicitado",
        "delta_vs_budget": "Delta vs budget anterior",
        "delta_vs_run_rate": "Delta vs run-rate",
        "challenge_questions": "Preguntas de reto",
        "review_status": "Estado",
        "ready_for_review": "Listo para review",
        "add_answer": "Añadir respuesta",
        "approve": "Aprobar",
        "modify": "Modificar",
        "challenge_answer_label": "Respuesta del departamento",
        "save_answer": "Guardar respuesta",
        "challenge_approved": "Aprobado",
        "challenge_needs_changes": "Enviado a modificar en Step 3",
        "challenge_responses_label": "Respuestas al challenge:",
        "load_evidence_title": "2. Evidencia",
        "load_evidence_intro": "Carga un archivo combinado de budget vs actuals o separa el budget aprobado y los reales YTD.",
        "combined_file": "Archivo combinado budget vs actuals",
        "separate_files": "Budget y actuals en archivos separados",
        "previous_budget_file": "Budget anual aprobado del año pasado",
        "actuals_file": "Actuals YTD por departamento / centro de coste",
        "normalized": "Evidencia normalizada.",
        "save_evidence": "Guardar evidencia",
        "last_saved_evidence": "Ultima evidencia guardada",
        "budget_rows": "Filas budget",
        "actual_rows": "Filas actuals",
        "departments": "Departamentos",
        "ytd_actuals": "Actuals YTD",
        "workspaces_title": "3. Workspaces por departamento",
        "open_workspace": "Workspace abierto",
        "change_department": "Cambiar departamento",
        "choose_department": "Dinos a que departamento perteneces. Abriremos su workspace y las preguntas quedaran dentro de ese contexto.",
        "choose_department_continue": "Elige tu departamento para continuar.",
        "open_department_workspace": "Abrir workspace de",
        "current_approved_budget": "Budget aprobado actual",
        "workspace_of": "Workspace de",
        "strategic_interview": "Entrevista estratégica",
        "strategic_caption": "Capturamos clientes y proyectos con campos concretos para construir una propuesta defendible.",
        "clients_projects": "Clientes y proyectos",
        "number_clients_projects": "Numero de clientes / proyectos",
        "client_project": "Cliente / proyecto",
        "client_project_name": "nombre",
        "client_project_cost": "coste estimado",
        "client_project_placeholder": "Ej. Cliente enterprise - lanzamiento producto",
        "go_live": "Go-live / fecha de impacto",
        "billing_month": "Mes de facturacion",
        "cost_type": "Tipo de coste",
        "personnel": "Personal necesario",
        "current_headcount": "Cuantas personas hay hoy en",
        "add_headcount": "Cuantas personas quieres anadir en",
        "remove_headcount": "Cuantas personas quieres quitar en",
        "hire": "Alta",
        "leaver": "Baja",
        "role": "rol",
        "gross_annual": "Gross anual",
        "gross_removed": "Gross anual que se elimina",
        "extra_company_cost": "Coste empresa adicional estimado",
        "company_saving": "Ahorro empresa estimado",
        "gross_help": "Salario bruto anual. La app calcula el coste empresa anadiendo Seguridad Social patronal estimada en Espana.",
        "effective_month": "Mes efectivo",
        "type": "Tipo",
        "number_of": "Numero de",
        "name": "nombre",
        "estimated_cost": "coste estimado",
        "frequency": "Frecuencia",
        "linked_to": "Asociado a",
        "not_linked": "No vinculado",
        "project_prefix": "Cliente/proyecto",
        "event_prefix": "Evento",
        "actions": "Acciones",
        "saved": "Guardado",
        "submitted": "Enviado",
        "last_submit": "Ultimo submit",
        "last_save": "Ultimo save",
        "approved_previous_budget": "Budget anual aprobado anterior",
        "actuals_ytd_until": "Actuals YTD hasta",
        "new_budget_proposal": "Propuesta budget nuevo año",
        "approved_budget": "Budget aprobado",
        "run_rate_forecast": "Run-rate forecast",
        "run_rate_forecast_explainer": "Run-rate forecast: proyeccion de cierre de año = (actuals acumulados hasta {month} ÷ meses transcurridos) × 12. Extrapola al resto del año el ritmo observado hasta ahora.",
        "proposed_budget": "Budget propuesto",
        "change_vs_approved": "Cambio vs aprobado",
        "line_detail": "Detalle por linea",
        "download_csv": "Descargar CSV",
        "download_excel": "Descargar Excel",
        "fpna_review_title": "Revision de FP&A",
        "fpna_review_intro": "FP&A revisa la desviacion y el comentario de cada departamento y decide si tienen sentido antes de dar el visto bueno final.",
        "fpna_comment_label": "Comentario de FP&A (opcional, se envia al departamento si se rechaza)",
        "fpna_approval": "FP&A Approval",
        "fpna_rejection": "FP&A Rejection",
        "fpna_approved_caption": "Aprobado por FP&A",
        "fpna_rejected_caption": "Rechazado por FP&A -- enviado a modificar en Step 3",
        "fpna_rejected_notice": "FP&A ha rechazado esta propuesta y pide revisarla.",
        "settings_title": "1. Settings",
        "settings_caption": "Configuracion maestra para controlar que aparece en Workspaces y como se clasifican automaticamente los datos cargados.",
        "visible_departments": "Departamentos visibles en Workspaces",
        "cost_centers": "Centros de coste",
        "accounts": "Cuentas",
        "categories": "Categorias",
        "extra_dimensions": "Dimensiones adicionales",
        "dimension_names": "Nombres de dimensiones",
        "dimension_values": "Valores permitidos",
        "smart_mapping": "Mapeo inteligente",
        "smart_mapping_caption": "Cada keyword se busca en categoria, cuenta y centro de coste. Si coincide, Budget Intelligence rellena las dimensiones indicadas.",
        "corporate_targets": "Objetivos corporativos (top-down)",
        "corporate_targets_caption": "Metricas top-down del ciclo (ventas, margen, EBITDA, plantilla, cash...). Se usan en el Step 4 (Challenge) para preguntar si lo que pide cada departamento ayuda o contradice estos objetivos.",
        "save_settings": "Save settings",
        "settings_saved": "Settings guardados.",
        "one_off": "Una vez",
        "monthly": "Mensual",
        "bimonthly": "Bimensual",
        "quarterly": "Trimestral",
        "semiannual": "Semestral",
        "annual": "Anual",
        "recurring": "Recurrente",
        "mixed": "Mixto",
        "permanent": "Fijo",
        "temporary": "Temporal",
        "freelance": "Freelance",
        "agency": "Agencia",
    },
    "English": {
        "language": "Language",
        "nav_label": "Navigation panel",
        "step": "Step",
        "settings": "Settings",
        "admin_setup": "Admin setup",
        "evidence": "Evidence",
        "budget_actuals": "Budget and actuals",
        "workspaces": "Workspaces",
        "by_department": "By department",
        "proposal": "Proposal",
        "challenge": "Challenge",
        "submitted_review": "Submitted review",
        "new_budget": "New budget",
        "challenge_title": "4. Budget Intelligence Challenge",
        "proposal_title": "5. Proposal",
        "challenge_intro": "Budget Intelligence reviews submitted workspaces and challenges each proposal against the previous approved budget and YTD actuals.",
        "no_submitted_departments": "No departments have been submitted yet. In Workspaces, press Submit so they enter this review.",
        "submitted_departments": "Submitted departments",
        "requested_plan": "Requested plan",
        "delta_vs_budget": "Delta vs previous budget",
        "delta_vs_run_rate": "Delta vs run-rate",
        "challenge_questions": "Challenge questions",
        "review_status": "Status",
        "ready_for_review": "Ready for review",
        "add_answer": "Add answer",
        "approve": "Approve",
        "modify": "Modify",
        "challenge_answer_label": "Department response",
        "save_answer": "Save answer",
        "challenge_approved": "Approved",
        "challenge_needs_changes": "Sent back to Step 3 for changes",
        "challenge_responses_label": "Challenge responses:",
        "load_evidence_title": "2. Evidence",
        "load_evidence_intro": "Upload a combined budget vs actuals file or separate the approved budget and YTD actuals.",
        "combined_file": "Combined budget vs actuals file",
        "separate_files": "Separate budget and actuals files",
        "previous_budget_file": "Last year's approved annual budget",
        "actuals_file": "YTD actuals by department / cost center",
        "normalized": "Evidence normalized.",
        "save_evidence": "Save evidence",
        "last_saved_evidence": "Last saved evidence",
        "budget_rows": "Budget rows",
        "actual_rows": "Actual rows",
        "departments": "Departments",
        "ytd_actuals": "YTD actuals",
        "workspaces_title": "3. Department workspaces",
        "open_workspace": "Open workspace",
        "change_department": "Change department",
        "choose_department": "Tell us which department you belong to. We will open its workspace and keep the questions in that context.",
        "choose_department_continue": "Choose your department to continue.",
        "open_department_workspace": "Open workspace for",
        "current_approved_budget": "Current approved budget",
        "workspace_of": "Workspace for",
        "strategic_interview": "Strategic interview",
        "strategic_caption": "We capture clients and projects with concrete fields to build a defensible proposal.",
        "clients_projects": "Clients and projects",
        "number_clients_projects": "Number of clients / projects",
        "client_project": "Client / project",
        "client_project_name": "name",
        "client_project_cost": "estimated cost",
        "client_project_placeholder": "E.g. Enterprise client - product launch",
        "go_live": "Go-live / impact date",
        "billing_month": "Billing month",
        "cost_type": "Cost type",
        "personnel": "Required personnel",
        "current_headcount": "How many people are currently in",
        "add_headcount": "How many people do you want to add in",
        "remove_headcount": "How many people do you want to remove in",
        "hire": "Hire",
        "leaver": "Reduction",
        "role": "role",
        "gross_annual": "Annual gross",
        "gross_removed": "Annual gross to remove",
        "extra_company_cost": "Estimated additional company cost",
        "company_saving": "Estimated company saving",
        "gross_help": "Annual gross salary. The app calculates company cost by adding estimated Spanish employer social security.",
        "effective_month": "Effective month",
        "type": "Type",
        "number_of": "Number of",
        "name": "name",
        "estimated_cost": "estimated cost",
        "frequency": "Frequency",
        "linked_to": "Linked to",
        "not_linked": "Not linked",
        "project_prefix": "Client/project",
        "event_prefix": "Event",
        "actions": "Actions",
        "saved": "Saved",
        "submitted": "Submitted",
        "last_submit": "Last submit",
        "last_save": "Last save",
        "approved_previous_budget": "Previous approved annual budget",
        "actuals_ytd_until": "YTD actuals through",
        "new_budget_proposal": "New year budget proposal",
        "approved_budget": "Approved budget",
        "run_rate_forecast": "Run-rate forecast",
        "run_rate_forecast_explainer": "Run-rate forecast: year-end projection = (YTD actuals through {month} ÷ months elapsed) × 12. It extrapolates the pace observed so far to the rest of the year.",
        "proposed_budget": "Proposed budget",
        "change_vs_approved": "Change vs approved",
        "line_detail": "Line detail",
        "download_csv": "Download CSV",
        "download_excel": "Download Excel",
        "fpna_review_title": "FP&A Review",
        "fpna_review_intro": "FP&A reviews each department's variance and comment and decides whether they make sense before giving final sign-off.",
        "fpna_comment_label": "FP&A comment (optional, sent to the department if rejected)",
        "fpna_approval": "FP&A Approval",
        "fpna_rejection": "FP&A Rejection",
        "fpna_approved_caption": "Approved by FP&A",
        "fpna_rejected_caption": "Rejected by FP&A -- sent back to Step 3 for changes",
        "fpna_rejected_notice": "FP&A rejected this proposal and asked for it to be revisited.",
        "settings_title": "1. Settings",
        "settings_caption": "Master setup to control what appears in Workspaces and how uploaded data is classified automatically.",
        "visible_departments": "Departments visible in Workspaces",
        "cost_centers": "Cost centers",
        "accounts": "Accounts",
        "categories": "Categories",
        "extra_dimensions": "Additional dimensions",
        "dimension_names": "Dimension names",
        "dimension_values": "Allowed values",
        "smart_mapping": "Smart mapping",
        "smart_mapping_caption": "Each keyword is searched in category, account, and cost center. If it matches, Budget Intelligence fills the indicated dimensions.",
        "corporate_targets": "Corporate targets (top-down)",
        "corporate_targets_caption": "Top-down metrics for the cycle (sales, margin, EBITDA, headcount, cash...). Used in Step 4 (Challenge) to ask whether each department's request supports or contradicts these company-wide goals.",
        "save_settings": "Save settings",
        "settings_saved": "Settings saved.",
        "one_off": "One-off",
        "monthly": "Monthly",
        "bimonthly": "Bimonthly",
        "quarterly": "Quarterly",
        "semiannual": "Semiannual",
        "annual": "Annual",
        "recurring": "Recurring",
        "mixed": "Mixed",
        "permanent": "Permanent",
        "temporary": "Temporary",
        "freelance": "Freelance",
        "agency": "Agency",
    },
}
BUSINESS_TRANSLATIONS_EN = {
    "Eventos": "Events",
    "Gastos de viaje": "Travel expenses",
    "Software y herramientas": "Software and tools",
    "Merchandising y materiales comerciales": "Merchandising and commercial materials",
    "Eventos comerciales": "Sales events",
    "Viajes comerciales": "Sales travel",
    "Comisiones / incentivos": "Commissions / incentives",
    "Viajes clientes": "Customer travel",
    "Research y discovery": "Research and discovery",
    "Cloud e infraestructura": "Cloud and infrastructure",
    "Seguridad y compliance": "Security and compliance",
    "Servicios externos": "External services",
    "Sistemas internos": "Internal systems",
    "Procesos y automatizacion": "Processes and automation",
    "Formacion y desarrollo": "Training and development",
    "Seleccion y employer branding": "Recruiting and employer branding",
    "Bienestar y engagement": "Wellbeing and engagement",
    "Auditoria y compliance": "Audit and compliance",
    "Banca y treasury": "Banking and treasury",
    "Contratos y legal ops": "Contracts and legal ops",
    "Compras y proveedores": "Purchasing and vendors",
    "Oficina y workplace": "Office and workplace",
    "Data platforms": "Data platforms",
    "Analytics y reporting": "Analytics and reporting",
    "Others": "Others",
    "Tipo de evento": "Event type",
    "Motivo": "Reason",
    "Tipo de herramienta": "Tool type",
    "Tipo de material": "Material type",
    "Tipo de gasto": "Expense type",
    "Tipo": "Type",
    "evento": "event",
    "viaje": "trip",
    "herramienta": "tool",
    "material": "material",
    "gasto": "expense",
    "servicio": "service",
    "programa": "program",
    "estudio": "study",
    "sistema": "system",
    "proceso": "process",
    "iniciativa": "initiative",
    "proveedor": "vendor",
    "licencia": "license",
}
DEFAULT_DIMENSIONS = ["Department", "Cost Center", "Account", "Category"]
DEFAULT_MEASURE = "Amount"
APP_DIR = Path(__file__).resolve().parent
KAGGLE_SAMPLE_PATH = APP_DIR / "kaggle_budget_vs_actual.csv"
SAVED_EVIDENCE_DIR = APP_DIR / "data" / "saved_evidence"
ADMIN_SETTINGS_PATH = SAVED_EVIDENCE_DIR / "admin_settings.json"
SPANISH_EMPLOYER_SS_RATE = 0.3065
STANDARD_DEPARTMENTS = [
    "Marketing",
    "Sales",
    "Customer Success",
    "Product",
    "Engineering",
    "IT",
    "Operations",
    "Finance",
    "People / RH",
    "Legal",
    "Procurement",
    "Facilities",
    "Data & Analytics",
    "General & Admin",
]
CLIENT_PROJECT_DEPARTMENTS = {
    "Marketing",
    "Sales",
    "Customer Success",
    "Product",
    "Engineering",
    "Product & Engineering",
    "IT",
    "Operations",
    "Data & Analytics",
}
FALLBACK_DEPARTMENT_BY_KEYWORD = {
    "development": "Product & Engineering",
    "engineering": "Product & Engineering",
    "product": "Product",
    "operat": "Operations",
    "logistics": "Operations",
    "marketing": "Marketing",
    "campaign": "Marketing",
    "brand": "Marketing",
    "media": "Marketing",
    "sales": "Sales",
    "travelling": "Sales",
    "travel": "Sales",
    "customer": "Customer Success",
    "support": "Customer Success",
    "success": "Customer Success",
    "training": "People",
    "people": "People / RH",
    "hr": "People / RH",
    "rh": "People / RH",
    "recruit": "People / RH",
    "talent": "People / RH",
    "maintenance": "IT & Facilities",
    "software": "IT & Facilities",
    "cloud": "IT & Facilities",
    "security": "IT",
    "finance": "Finance",
    "audit": "Finance",
    "legal": "Legal",
    "procurement": "Procurement",
    "vendor": "Procurement",
    "facility": "Facilities",
    "office": "Facilities",
    "data": "Data & Analytics",
    "analytics": "Data & Analytics",
}
DEFAULT_ADMIN_SETTINGS = {
    "departments": STANDARD_DEPARTMENTS,
    "cost_centers": ["Unassigned"],
    "accounts": ["Unspecified"],
    "categories": [
        "Projects",
        "Personnel",
        "Travel",
        "Software",
        "Events",
        "Merchandising",
        "External services",
        "Training",
        "Compliance",
        "Others",
    ],
    "extra_dimensions": [
        {"dimension": "Region", "values": "Spain, EMEA, US"},
        {"dimension": "Business Unit", "values": "Core, Growth, Corporate"},
    ],
    "smart_mapping": [
        {"keyword": keyword, "department": department, "cost_center": "", "account": "", "category": ""}
        for keyword, department in FALLBACK_DEPARTMENT_BY_KEYWORD.items()
    ],
    # Targets top-down de compañia para el ciclo (ventas, margen, EBITDA,
    # plantilla, cash...). Es una tabla libre (nombre + valor + unidad),
    # no una lista cerrada de metricas -- se usan en Step 4 (Challenge)
    # para preguntar si lo que pide cada departamento ayuda o contradice
    # estos objetivos globales.
    "corporate_targets": [
        {"target": "Sales volume", "value": "", "unit": "EUR", "note": ""},
        {"target": "Gross margin", "value": "", "unit": "%", "note": ""},
        {"target": "EBITDA", "value": "", "unit": "EUR", "note": ""},
        {"target": "Headcount cap", "value": "", "unit": "FTEs", "note": ""},
    ],
}
DEFAULT_ROLE_OPTIONS = ["Department Lead", "Manager", "Specialist", "Analyst", "Coordinator", "External contractor"]
DEPARTMENT_ROLE_OPTIONS = {
    "Marketing": [
        "Demand Generation Manager",
        "Performance Marketing Specialist",
        "Content Marketing Manager",
        "Product Marketing Manager",
        "Marketing Operations",
        "Events Manager",
        "Brand / Comms",
        "Marketing Analyst",
    ],
    "Sales": ["Account Executive", "Sales Development Rep", "Sales Manager", "Sales Operations", "Partner Manager"],
    "Customer Success": ["Customer Success Manager", "Implementation Manager", "Support Specialist", "Renewals Manager"],
    "Product": ["Product Manager", "Product Designer", "UX Researcher", "Product Operations"],
    "Engineering": ["Software Engineer", "QA Engineer", "DevOps Engineer", "Engineering Manager", "Data Engineer"],
    "IT": ["IT Support", "Systems Administrator", "Security Specialist", "Infrastructure Engineer"],
    "Operations": ["Operations Manager", "Process Analyst", "Logistics Coordinator", "Quality Specialist"],
    "Finance": ["FP&A Analyst", "Accountant", "Controller", "Treasury Analyst", "Billing Specialist"],
    "People / RH": ["Recruiter", "People Partner", "People Operations", "Learning & Development", "Compensation Analyst"],
    "Legal": ["Legal Counsel", "Paralegal", "Compliance Specialist", "Contract Manager"],
    "Procurement": ["Procurement Manager", "Category Manager", "Vendor Manager", "Purchasing Specialist"],
    "Facilities": ["Facilities Manager", "Office Manager", "Workplace Coordinator", "Maintenance Specialist"],
    "Data & Analytics": ["Data Analyst", "Analytics Engineer", "BI Developer", "Data Scientist", "Data Governance Lead"],
}
DEPARTMENT_EXPENSE_SECTIONS = {
    "Marketing": [
        {
            "title": "Eventos",
            "storage_key": "events",
            "item_label": "evento",
            "category_label": "Tipo de evento",
            "category_options": ["Feria / trade show", "Evento propio", "Sponsorship", "Webinar", "Customer event", "Partner event"],
            "default_count": 3,
        },
        {
            "title": "Gastos de viaje",
            "storage_key": "travel",
            "item_label": "viaje",
            "category_label": "Motivo",
            "category_options": ["Cliente / prospect", "Evento / feria", "Partner", "Kick-off interno", "Workshop producto", "Research mercado"],
            "link_to_events": True,
        },
        {
            "title": "Software y herramientas",
            "storage_key": "software",
            "item_label": "herramienta",
            "category_label": "Tipo de herramienta",
            "category_options": ["Marketing automation", "CRM / sales enablement", "Analytics / attribution", "SEO / content", "Design / creative", "Webinars / events", "Social media", "Data enrichment"],
        },
        {
            "title": "Merchandising y materiales comerciales",
            "storage_key": "merchandising",
            "item_label": "material",
            "category_label": "Tipo de material",
            "category_options": ["Swag", "Brochures", "Sales decks", "One-pagers", "Samples", "Packaging", "Customer gifts"],
        },
        {
            "title": "Others",
            "storage_key": "others",
            "item_label": "gasto",
            "category_label": "Tipo de gasto",
            "category_options": ["Agencias / freelancers", "Contenido / produccion", "PR / comunicaciones", "Research", "Formacion", "Licencias menores", "Otro"],
        },
    ],
    "Sales": [
        {"title": "Eventos comerciales", "storage_key": "events", "item_label": "evento", "category_label": "Tipo de evento", "category_options": ["Customer event", "Partner event", "Trade show", "Executive dinner", "Webinar"], "default_count": 3},
        {"title": "Viajes comerciales", "storage_key": "travel", "item_label": "viaje", "category_label": "Motivo", "category_options": ["Cliente / prospect", "QBR", "Partner", "Negociacion", "Kick-off"], "link_to_events": True},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["CRM", "Sales engagement", "CPQ", "Data enrichment", "Forecasting", "Call recording"]},
        {"title": "Comisiones / incentivos", "storage_key": "incentives", "item_label": "incentivo", "category_label": "Tipo", "category_options": ["Comisiones", "SPIFF", "Bonus acelerador", "Partner incentive"]},
    ],
    "Customer Success": [
        {"title": "Customer programs", "storage_key": "customer_programs", "item_label": "programa", "category_label": "Tipo", "category_options": ["Onboarding", "Training clientes", "Community", "Advocacy", "Retention campaign"]},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["CS platform", "Support desk", "NPS / surveys", "Knowledge base", "Product analytics"]},
        {"title": "Viajes clientes", "storage_key": "travel", "item_label": "viaje", "category_label": "Motivo", "category_options": ["QBR", "Onsite training", "Escalacion", "Executive meeting"]},
    ],
    "Product": [
        {"title": "Research y discovery", "storage_key": "research", "item_label": "estudio", "category_label": "Tipo", "category_options": ["User research", "Market research", "Prototype testing", "Benchmark competitivo"]},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["Roadmap", "Design", "Analytics", "Experimentation", "User feedback"]},
        {"title": "Servicios externos", "storage_key": "external_services", "item_label": "servicio", "category_label": "Tipo", "category_options": ["UX agency", "Research vendor", "Product consultant", "Localization"]},
    ],
    "Engineering": [
        {"title": "Cloud e infraestructura", "storage_key": "cloud", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Cloud compute", "Storage", "Monitoring", "CI/CD", "Environments"]},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["Developer tools", "Testing", "Security scanning", "Observability", "AI coding tools"]},
        {"title": "Servicios externos", "storage_key": "external_services", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Contractors", "Architecture review", "Pen test", "Implementation partner"]},
    ],
    "IT": [
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["Identity", "Device management", "Helpdesk", "Security", "Backup", "Collaboration"]},
        {"title": "Hardware y equipos", "storage_key": "hardware", "item_label": "equipo", "category_label": "Tipo", "category_options": ["Laptops", "Monitors", "Network", "Phones", "Peripherals"]},
        {"title": "Seguridad y compliance", "storage_key": "security", "item_label": "iniciativa", "category_label": "Tipo", "category_options": ["Audit", "Pen test", "Awareness", "Endpoint security", "Compliance tooling"]},
    ],
    "Operations": [
        {"title": "Servicios operativos", "storage_key": "services", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Logistics", "Warehousing", "BPO", "Quality", "Process improvement"]},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["ERP", "Workflow", "Scheduling", "Inventory", "Automation"]},
        {"title": "Viajes operativos", "storage_key": "travel", "item_label": "viaje", "category_label": "Motivo", "category_options": ["Site visit", "Vendor visit", "Process audit", "Training"]},
    ],
    "Finance": [
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["ERP / accounting", "FP&A", "Close management", "Treasury", "Billing", "Tax"]},
        {"title": "Auditoria, tax y advisory", "storage_key": "advisory", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Audit", "Tax", "Transfer pricing", "Accounting advisory", "Valuation"]},
        {"title": "Compliance y reporting", "storage_key": "compliance", "item_label": "iniciativa", "category_label": "Tipo", "category_options": ["Statutory reporting", "Internal controls", "Policy update", "Regulatory filing"]},
    ],
    "People / RH": [
        {"title": "Recruiting", "storage_key": "recruiting", "item_label": "gasto", "category_label": "Tipo", "category_options": ["Job boards", "Recruiting agency", "Assessment tools", "Employer branding", "Referral bonus"]},
        {"title": "Formacion y desarrollo", "storage_key": "learning", "item_label": "programa", "category_label": "Tipo", "category_options": ["Leadership", "Technical training", "Compliance training", "Coaching", "Certifications"]},
        {"title": "Software HR", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["HRIS", "ATS", "Performance", "Engagement survey", "Payroll"]},
    ],
    "Legal": [
        {"title": "Servicios legales externos", "storage_key": "external_legal", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Corporate", "Commercial contracts", "Employment", "IP", "Litigation", "Privacy"]},
        {"title": "Compliance", "storage_key": "compliance", "item_label": "iniciativa", "category_label": "Tipo", "category_options": ["Privacy", "Regulatory", "Policy update", "Training", "Audit"]},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["Contract management", "eSignature", "Legal research", "Entity management"]},
    ],
    "Procurement": [
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["Procurement suite", "Vendor risk", "Contract repository", "Spend analytics"]},
        {"title": "Vendor management", "storage_key": "vendors", "item_label": "iniciativa", "category_label": "Tipo", "category_options": ["Supplier audit", "RFP support", "Savings program", "Risk review"]},
        {"title": "Servicios externos", "storage_key": "external_services", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Consulting", "Benchmarking", "Legal support", "Implementation"]},
    ],
    "Facilities": [
        {"title": "Oficina y mantenimiento", "storage_key": "office", "item_label": "gasto", "category_label": "Tipo", "category_options": ["Rent", "Maintenance", "Cleaning", "Utilities", "Furniture"]},
        {"title": "Equipamiento", "storage_key": "equipment", "item_label": "equipo", "category_label": "Tipo", "category_options": ["Desks", "Chairs", "Meeting rooms", "Security access", "Kitchen"]},
        {"title": "Servicios externos", "storage_key": "external_services", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Facilities provider", "Repairs", "Office move", "Health and safety"]},
    ],
    "Data & Analytics": [
        {"title": "Data platforms", "storage_key": "platforms", "item_label": "plataforma", "category_label": "Tipo", "category_options": ["Warehouse", "BI", "ETL / ELT", "Reverse ETL", "Data quality", "Governance"]},
        {"title": "Software y herramientas", "storage_key": "software", "item_label": "herramienta", "category_label": "Tipo de herramienta", "category_options": ["BI", "Notebook", "Data catalog", "Monitoring", "Experimentation"]},
        {"title": "Servicios externos", "storage_key": "external_services", "item_label": "servicio", "category_label": "Tipo", "category_options": ["Data consulting", "Implementation", "Training", "Architecture review"]},
    ],
}


@dataclass(frozen=True)
class ColumnMap:
    department: str | None
    cost_center: str | None
    account: str | None
    category: str | None
    month: str | None
    amount: str | None


@dataclass(frozen=True)
class CombinedColumnMap:
    department: str | None
    cost_center: str | None
    account: str | None
    month: str | None
    scenario: str | None
    amount: str | None
    budget_amount: str | None
    actual_amount: str | None
    value_columns: tuple[str, ...]


st.set_page_config(page_title="Budget Intelligence", page_icon=":material/psychology:", layout="wide")


def main() -> None:
    inject_styles()
    initialise_state()

    render_header()
    active_step = render_navigation()
    if active_step == "evidence":
        render_load_evidence()
    elif active_step == "workspaces":
        render_department_interview()
    elif active_step == "challenge":
        render_budget_challenge()
    elif active_step == "proposal":
        render_proposed_budget()
    else:
        render_admin_settings()


def inject_styles() -> None:
    st.markdown(
        """
        <style>
        :root {
            --finance-ink: #1f2937;
            --finance-muted: #667085;
            --finance-soft: #f5f7fb;
            --finance-panel: #ffffff;
            --finance-line: #d9e1ec;
            --finance-accent: #0f766e;
            --finance-accent-strong: #115e59;
            --finance-coral: #e94f4f;
            --finance-warm: #b7791f;
            --finance-blue: #2563eb;
            --finance-shadow: 0 18px 45px rgba(31, 41, 55, 0.08);
        }
        html, body, [class*="css"] {
            font-family: Inter, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        }
        .stApp {
            background:
                linear-gradient(180deg, #f7f9fc 0%, #ffffff 36%),
                #ffffff;
            color: var(--finance-ink);
        }
        .block-container {
            padding-top: 2.4rem;
            padding-bottom: 4rem;
            max-width: 1220px;
        }
        section[data-testid="stSidebar"] {
            display: none;
        }
        button[kind="header"] {
            display: none;
        }
        .budget-hero {
            position: relative;
            padding: 1.15rem 0 1.35rem;
            border-bottom: 1px solid rgba(217, 225, 236, 0.9);
            margin-bottom: 1.15rem;
        }
        .budget-hero::before {
            content: "";
            position: absolute;
            left: 0;
            top: 0;
            width: 58px;
            height: 4px;
            border-radius: 999px;
            background: linear-gradient(90deg, var(--finance-accent), var(--finance-blue), var(--finance-coral));
        }
        .budget-kicker {
            margin-top: 0.8rem;
            color: var(--finance-accent-strong);
            font-size: 0.78rem;
            font-weight: 800;
            letter-spacing: 0.08em;
            text-transform: uppercase;
        }
        .budget-hero h1 {
            margin: 0.25rem 0 0.45rem;
            color: var(--finance-ink);
            font-size: 3rem;
            line-height: 1;
            letter-spacing: 0;
        }
        h2, h3, h4 {
            color: var(--finance-ink);
            letter-spacing: 0;
        }
        .main-nav {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 1rem;
            margin: 0.6rem 0 2rem;
            padding: 0.85rem;
            border: 1px solid var(--finance-line);
            border-radius: 8px;
            background: #ffffff;
            box-shadow: 0 16px 38px rgba(31, 41, 55, 0.08);
        }
        .main-nav a {
            display: flex;
            flex-direction: column;
            justify-content: center;
            min-height: 6.4rem;
            padding: 1.1rem 1.25rem;
            border: 1px solid #dde5ef;
            border-radius: 8px;
            background: #f7f9fc;
            color: #344054;
            text-decoration: none;
            box-shadow: inset 0 -4px 0 transparent;
        }
        .main-nav a:hover {
            border-color: var(--finance-accent);
            background: rgba(15, 118, 110, 0.08);
            color: var(--finance-accent-strong);
        }
        .main-nav a.active {
            border-color: var(--finance-accent);
            background: var(--finance-accent);
            color: #ffffff;
            box-shadow: 0 14px 28px rgba(15, 118, 110, 0.24);
        }
        .main-nav .step-number {
            font-size: 0.82rem;
            font-weight: 850;
            text-transform: uppercase;
            opacity: 0.82;
            margin-bottom: 0.4rem;
        }
        .main-nav .step-title {
            font-size: 1.25rem;
            line-height: 1.2;
            font-weight: 850;
            letter-spacing: 0;
        }
        .main-nav .step-caption {
            margin-top: 0.35rem;
            font-size: 0.86rem;
            line-height: 1.3;
            opacity: 0.78;
        }
        .main-nav-label {
            color: var(--finance-muted);
            font-size: 0.82rem;
            font-weight: 850;
            letter-spacing: 0.08em;
            margin: 0.3rem 0 0.55rem;
            text-transform: uppercase;
        }
        div[class*="st-key-nav_"] button {
            height: 8.4rem !important;
            min-height: 8.4rem !important;
            max-height: 8.4rem !important;
            padding: 1.15rem 1.1rem !important;
            border-radius: 8px !important;
            font-size: 1.15rem !important;
            font-weight: 850 !important;
            justify-content: flex-start !important;
            align-items: flex-start !important;
            text-align: left !important;
            box-shadow: 0 14px 30px rgba(31, 41, 55, 0.08) !important;
            overflow: hidden !important;
        }
        div[class*="st-key-nav_"] button p {
            font-size: 1.15rem !important;
            font-weight: 850 !important;
            line-height: 1.28 !important;
            white-space: pre-line !important;
            margin: 0 !important;
        }
        div[class*="st-key-nav_"] button[kind="primary"] {
            background: var(--finance-accent) !important;
            border-color: var(--finance-accent) !important;
            color: #ffffff !important;
        }
        div[class*="st-key-nav_"] button[kind="primary"] p {
            color: #ffffff !important;
        }
        div[class*="st-key-change_department"] button {
            min-height: 4.2rem !important;
            width: 100% !important;
            padding: 0.9rem 1.2rem !important;
            font-size: 1rem !important;
            font-weight: 850 !important;
            border-color: var(--finance-accent) !important;
            color: var(--finance-accent-strong) !important;
            background: #ffffff !important;
            box-shadow: 0 12px 24px rgba(15, 118, 110, 0.12) !important;
        }
        div[class*="st-key-change_department"] button:hover {
            background: rgba(15, 118, 110, 0.08) !important;
        }
        div[data-testid="stTabs"] [data-baseweb="tab-list"] {
            display: grid !important;
            grid-template-columns: repeat(3, minmax(0, 1fr)) !important;
            gap: 1rem !important;
            width: 100% !important;
            border: 1px solid var(--finance-line) !important;
            border-radius: 8px !important;
            padding: 0.8rem !important;
            background: #ffffff !important;
            box-shadow: 0 14px 34px rgba(31, 41, 55, 0.08) !important;
            margin: 0.45rem 0 1.8rem !important;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"] {
            width: 100% !important;
            height: 5.6rem !important;
            min-height: 5.6rem !important;
            border-radius: 8px !important;
            color: #526071 !important;
            padding: 0 1.45rem !important;
            background: #f7f9fc !important;
            border: 1px solid #e1e7f0 !important;
            justify-content: flex-start !important;
            align-items: center !important;
            box-shadow: inset 0 -3px 0 transparent !important;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"] p {
            font-size: 1.18rem !important;
            font-weight: 850 !important;
            line-height: 1.2 !important;
            letter-spacing: 0 !important;
        }
        div[data-testid="stTabs"] [aria-selected="true"] {
            color: #ffffff !important;
            background: var(--finance-accent) !important;
            border-color: var(--finance-accent) !important;
            box-shadow: 0 14px 28px rgba(15, 118, 110, 0.24) !important;
        }
        div[data-testid="stTabs"] [data-baseweb="tab"]:hover {
            color: var(--finance-accent-strong) !important;
            background: rgba(15, 118, 110, 0.08) !important;
        }
        div[data-testid="stTabs"] [aria-selected="true"]:hover {
            color: #ffffff !important;
            background: var(--finance-accent-strong) !important;
        }
        div[data-testid="stTabs"] [aria-selected="true"] p {
            color: #ffffff !important;
        }
        div[data-testid="stRadio"] {
            padding: 0.55rem 0 0.35rem;
        }
        div[data-testid="stMetric"] {
            border: 1px solid var(--finance-line);
            border-radius: 8px;
            padding: 1rem 1rem 0.9rem;
            background: var(--finance-panel);
            box-shadow: 0 8px 22px rgba(31, 41, 55, 0.045);
        }
        div[data-testid="stMetric"] label {
            color: var(--finance-muted);
            font-weight: 750;
        }
        div[data-testid="stMetric"] [data-testid="stMetricValue"] {
            color: #252b37;
            font-size: 2rem;
            letter-spacing: 0;
        }
        .budget-status {
            border: 1px solid var(--finance-line);
            border-radius: 8px;
            padding: 0.8rem 1rem;
            background: #f8fafc;
        }
        .question-card {
            border: 1px solid var(--finance-line);
            border-left: 5px solid var(--finance-accent);
            border-radius: 8px;
            padding: 1rem 1.05rem;
            background: #ffffff;
            margin: 0.85rem 0 0.55rem;
            box-shadow: 0 10px 26px rgba(31, 41, 55, 0.055);
        }
        .question-card strong {
            display: block;
            color: var(--finance-ink);
            margin-bottom: 0.45rem;
            font-size: 1.02rem;
            line-height: 1.35;
        }
        .question-card span {
            color: var(--finance-muted);
            font-size: 0.92rem;
            line-height: 1.45;
        }
        div[data-testid="stFileUploader"] section {
            border: 1px dashed #b8c2d3;
            border-radius: 8px;
            background: #f3f6fa;
            padding: 1rem;
        }
        div[data-testid="stFileUploader"] button {
            border-radius: 8px;
            border-color: #cfd7e4;
        }
        div[data-testid="stDataFrame"] {
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 10px 30px rgba(31, 41, 55, 0.05);
        }
        div[data-testid="stExpander"] {
            border: 1px solid var(--finance-line);
            border-radius: 8px;
            overflow: hidden;
            box-shadow: 0 8px 24px rgba(31, 41, 55, 0.04);
        }
        div[data-testid="stTextArea"] textarea,
        div[data-testid="stSelectbox"] div[data-baseweb="select"] > div,
        div[data-testid="stNumberInput"] input {
            border-radius: 8px;
            border-color: #cad3df;
        }
        div[data-testid="stButton"] button {
            border-radius: 8px;
            min-height: 3.05rem;
            font-weight: 700;
            border-color: #cbd5e1;
            color: #253044;
            background: #ffffff;
            box-shadow: 0 7px 16px rgba(31, 41, 55, 0.05);
            transition: all 0.15s ease;
        }
        div[data-testid="stButton"] button p {
            white-space: pre-line;
            line-height: 1.28;
        }
        div[data-testid="stButton"] button:hover {
            border-color: var(--finance-accent);
            color: var(--finance-accent-strong);
            transform: translateY(-1px);
            box-shadow: 0 11px 22px rgba(15, 118, 110, 0.12);
        }
        div[data-testid="stButton"] button[kind="primary"] {
            background: var(--finance-accent);
            border-color: var(--finance-accent);
            color: white;
        }
        div[data-testid="stButton"] button[kind="primary"]:hover {
            background: var(--finance-accent-strong);
            border-color: var(--finance-accent-strong);
            color: white;
        }
        @media (max-width: 700px) {
            .block-container {
                padding-top: 1.2rem;
            }
            .budget-hero h1 {
                font-size: 2.2rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def initialise_state() -> None:
    st.session_state.setdefault("budget_year", 2027)
    st.session_state.setdefault("ytd_month", "Aug")
    st.session_state.setdefault("target_margin", 18.0)
    st.session_state.setdefault("growth_cap", 12.0)
    st.session_state.setdefault("department_answers", {})
    st.session_state.setdefault("active_department_workspace", None)
    st.session_state.setdefault("active_step", "settings")
    st.session_state.setdefault("admin_settings", load_admin_settings())
    st.session_state.setdefault("language", "Español")
    load_saved_evidence_if_available()


def t(key: str) -> str:
    language = st.session_state.get("language", "Español")
    return TRANSLATIONS.get(language, TRANSLATIONS["Español"]).get(key, key)


def business_label(text: str) -> str:
    if st.session_state.get("language") == "English":
        return BUSINESS_TRANSLATIONS_EN.get(text, text)
    return text


def department_needs_clients_projects(department: str) -> bool:
    if department in CLIENT_PROJECT_DEPARTMENTS:
        return True
    lowered = department.lower()
    return any(
        token in lowered
        for token in ["marketing", "sales", "customer", "product", "engineering", "it", "operations", "data"]
    )


def frequency_options() -> list[str]:
    return [t("one_off"), t("monthly"), t("bimonthly"), t("quarterly"), t("semiannual"), t("annual")]


def cost_type_options() -> list[str]:
    return [t("one_off"), t("recurring"), t("mixed")]


def contract_type_options() -> list[str]:
    return [t("permanent"), t("temporary"), t("freelance"), t("agency")]


def render_header() -> None:
    header_cols = st.columns([4, 1])
    with header_cols[0]:
        st.markdown(
            """
            <section class="budget-hero">
                <div class="budget-kicker">FP&A planning workspace</div>
                <h1>Budget Intelligence</h1>
            </section>
            """,
            unsafe_allow_html=True,
        )
    with header_cols[1]:
        st.radio(t("language"), LANGUAGE_OPTIONS, horizontal=True, key="language")


def render_navigation() -> str:
    active_step = st.session_state.get("active_step", "settings")
    if active_step not in {"evidence", "workspaces", "challenge", "proposal", "settings"}:
        active_step = "settings"

    steps = [
        ("settings", "1", t("settings"), t("admin_setup")),
        ("evidence", "2", t("evidence"), t("budget_actuals")),
        ("workspaces", "3", t("workspaces"), t("by_department")),
        ("challenge", "4", t("challenge"), t("submitted_review")),
        ("proposal", "5", t("proposal"), t("new_budget")),
    ]
    st.markdown(f'<div class="main-nav-label">{t("nav_label")}</div>', unsafe_allow_html=True)
    columns = st.columns(5)
    for column, (key, number, title, caption) in zip(columns, steps):
        with column:
            button_type = "primary" if key == active_step else "secondary"
            if st.button(
                f"{t('step')} {number}\n\n{title}\n{caption}",
                key=f"nav_{key}",
                type=button_type,
                use_container_width=True,
            ):
                st.session_state.active_step = key
                st.rerun()
    return st.session_state.active_step


def render_load_evidence() -> None:
    st.subheader(t("load_evidence_title"))
    st.write(t("load_evidence_intro"))

    source_mode = st.radio(
        "Input structure",
        [t("combined_file"), t("separate_files")],
        horizontal=True,
        label_visibility="collapsed",
    )

    if source_mode == t("combined_file"):
        combined_file = st.file_uploader(t("combined_file"), type=["csv", "xlsx"], key="combined_budget_actual_file")
        if combined_file:
            combined_raw = read_uploaded_table(combined_file)
            st.session_state.combined_source_raw = combined_raw
            combined_map = guess_combined_columns(combined_raw)
            budget, actuals = normalize_combined_table(combined_raw, combined_map)
            st.session_state.approved_budget = budget
            st.session_state.actuals_ytd = actuals
            st.success(t("normalized"))
    else:
        budget_file = st.file_uploader(t("previous_budget_file"), type=["csv", "xlsx"], key="approved_budget_file")
        actuals_file = st.file_uploader(t("actuals_file"), type=["csv", "xlsx"], key="actuals_ytd_file")

        if budget_file and actuals_file:
            budget_raw = read_uploaded_table(budget_file)
            actuals_raw = read_uploaded_table(actuals_file)
            st.session_state.approved_budget_raw = budget_raw
            st.session_state.actuals_ytd_raw = actuals_raw

            with st.expander("Column mapping", expanded=True):
                st.write("Mapea las columnas una vez. El motor las normaliza al mismo contrato interno.")
                budget_map = render_column_mapping(budget_raw, "budget")
                actuals_map = render_column_mapping(actuals_raw, "actuals")

            if st.button("Normalizar y analizar", type="primary"):
                st.session_state.approved_budget = normalize_table(budget_raw, budget_map, annual_source=True)
                st.session_state.actuals_ytd = normalize_table(actuals_raw, actuals_map, annual_source=False)
                st.success(t("normalized"))

    if evidence_loaded():
        budget = st.session_state.approved_budget
        actuals = st.session_state.actuals_ytd
        save_cols = st.columns([1, 3])
        with save_cols[0]:
            if st.button(t("save_evidence"), type="primary"):
                saved_at = save_current_evidence()
                st.session_state.last_evidence_saved_at = saved_at
                st.success(f"{t('saved')}: {saved_at}.")
        with save_cols[1]:
            if st.session_state.get("last_evidence_saved_at"):
                st.caption(f"{t('last_saved_evidence')}: {st.session_state.last_evidence_saved_at}")

        cols = st.columns(4)
        cols[0].metric(t("budget_rows"), f"{len(budget):,}")
        cols[1].metric(t("actual_rows"), f"{len(actuals):,}")
        cols[2].metric(t("departments"), budget["Department"].nunique())
        cols[3].metric(t("ytd_actuals"), money(actuals["Amount"].sum()))
        st.dataframe(build_gap_summary().head(20), use_container_width=True, hide_index=True, key="load_evidence_gap_preview")


def render_department_interview() -> None:
    st.subheader(t("workspaces_title"))
    if not require_evidence():
        return

    summary = build_gap_summary()
    departments = build_department_workspace_list(summary)
    selected = render_department_workspace_entry(departments)
    if not selected:
        return

    st.markdown(f"### {t('workspace_of')} {selected}")

    if st.session_state.get(f"fpna_status::{selected}") == "rejected":
        notice = t("fpna_rejected_notice")
        fpna_comment = st.session_state.get(f"fpna_comment::{selected}", "")
        if fpna_comment:
            notice = f"{notice} {fpna_comment}"
        st.error(notice)

    answers = st.session_state.department_answers.setdefault(selected, {})
    answers["_department"] = selected
    render_department_budget_snapshot(selected, summary, answers)

    st.markdown(f"#### {t('strategic_interview')}")
    section_number = 1
    if department_needs_clients_projects(selected):
        st.caption(t("strategic_caption"))
        render_department_project_intake(selected, answers)
        section_number = 2
    else:
        answers["_project_count"] = 0
    render_department_headcount_intake(selected, answers, section_number=section_number)
    render_department_configured_expenses(selected, answers, starting_section=section_number + 1)

    render_department_workspace_actions(selected, answers)


def render_department_project_intake(department: str, answers: dict) -> None:
    st.markdown(f"##### 1. {t('clients_projects')}")
    project_count = st.number_input(
        t("number_clients_projects"),
        min_value=1,
        max_value=12,
        value=int(answers.get("_project_count", 3)),
        step=1,
        key=f"project_count_{department}",
    )
    answers["_project_count"] = int(project_count)

    for project_number in range(1, int(project_count) + 1):
        project_key = f"_project_{project_number}"
        project = answers.setdefault(project_key, {})
        if not isinstance(project, dict):
            project = {}
            answers[project_key] = project
        with st.expander(f"{t('client_project')} {project_number}", expanded=project_number == 1):
            name_col, cost_col = st.columns([2, 1])
            with name_col:
                project["name"] = st.text_input(
                    f"{t('client_project')} {project_number}: {t('client_project_name')}",
                    value=project.get("name", ""),
                    key=f"{department}_{project_key}_name",
                    placeholder=t("client_project_placeholder"),
                )
            with cost_col:
                project["cost"] = st.number_input(
                    f"{t('client_project')} {project_number}: {t('client_project_cost')}",
                    min_value=0.0,
                    value=float(project.get("cost", 0.0)),
                    step=1000.0,
                    key=f"{department}_{project_key}_cost",
                )

            timing_cols = st.columns(3)
            with timing_cols[0]:
                project["go_live"] = st.selectbox(
                    t("go_live"),
                    MONTHS,
                    index=select_index(MONTHS, project.get("go_live")),
                    key=f"{department}_{project_key}_go_live",
                )
            with timing_cols[1]:
                project["billing_month"] = st.selectbox(
                    t("billing_month"),
                    MONTHS,
                    index=select_index(MONTHS, project.get("billing_month")),
                    key=f"{department}_{project_key}_billing",
                )
            with timing_cols[2]:
                cost_types = cost_type_options()
                project["cost_type"] = st.selectbox(
                    t("cost_type"),
                    cost_types,
                    index=select_index(cost_types, project.get("cost_type")),
                    key=f"{department}_{project_key}_cost_type",
                )


def render_department_budget_snapshot(department: str, summary: pd.DataFrame, answers: dict) -> None:
    department_summary = summary[summary["Department"] == department]
    approved_budget = float(department_summary["Approved Budget"].sum())
    ytd_actuals = float(department_summary["YTD Actuals"].sum())
    proposed_budget = department_proposed_budget(department, summary, answers)
    ytd_month = detected_ytd_month(department)

    cols = st.columns(3)
    cols[0].metric(t("approved_previous_budget"), money(approved_budget))
    cols[1].metric(f"{t('actuals_ytd_until')} {ytd_month}", money(ytd_actuals))
    cols[2].metric(t("new_budget_proposal"), money(proposed_budget))


def render_department_headcount_intake(department: str, answers: dict, section_number: int) -> None:
    st.markdown(f"##### {section_number}. {t('personnel')}")
    dept_slug = department_slug(department)
    answers[f"_{dept_slug}_current_headcount"] = st.number_input(
        f"{t('current_headcount')} {department}",
        min_value=0.0,
        max_value=500.0,
        value=float(answers.get(f"_{dept_slug}_current_headcount", 0.0)),
        step=0.5,
        key=f"current_headcount_{department}",
    )

    movement_cols = st.columns(2)
    with movement_cols[0]:
        add_count = st.number_input(
            f"{t('add_headcount')} {department}",
            min_value=0,
            max_value=50,
            value=int(answers.get(f"_{dept_slug}_headcount_count", 0)),
            step=1,
            key=f"headcount_count_{department}",
        )
    with movement_cols[1]:
        remove_count = st.number_input(
            f"{t('remove_headcount')} {department}",
            min_value=0,
            max_value=50,
            value=int(answers.get(f"_{dept_slug}_headcount_reduction_count", 0)),
            step=1,
            key=f"headcount_reduction_count_{department}",
        )
    answers[f"_{dept_slug}_headcount_count"] = int(add_count)
    answers[f"_{dept_slug}_headcount_reduction_count"] = int(remove_count)

    render_headcount_movement_rows(
        department=department,
        answers=answers,
        count=int(add_count),
        key_prefix="role",
        title_prefix=t("hire"),
        cost_label=t("gross_annual"),
        caption_action=t("extra_company_cost"),
    )
    render_headcount_movement_rows(
        department=department,
        answers=answers,
        count=int(remove_count),
        key_prefix="reduction_role",
        title_prefix=t("leaver"),
        cost_label=t("gross_removed"),
        caption_action=t("company_saving"),
    )


def render_headcount_movement_rows(
    department: str,
    answers: dict,
    count: int,
    key_prefix: str,
    title_prefix: str,
    cost_label: str,
    caption_action: str,
) -> None:
    dept_slug = department_slug(department)
    for role_number in range(1, count + 1):
        role_key = f"_{dept_slug}_{key_prefix}_{role_number}"
        role = answers.setdefault(role_key, {})
        if not isinstance(role, dict):
            role = {}
            answers[role_key] = role
        with st.expander(f"{title_prefix} {role_number}", expanded=role_number == 1):
            role_cols = st.columns([2, 1, 1])
            with role_cols[0]:
                role_options = department_role_options(department)
                role["role"] = st.selectbox(
                    f"{title_prefix} {role_number}: {t('role')}",
                    role_options,
                    index=select_index(role_options, role.get("role")),
                    key=f"{department}_{role_key}_role",
                )
            with role_cols[1]:
                role["fte"] = st.number_input(
                    "FTEs",
                    min_value=0.0,
                    max_value=10.0,
                    value=float(role.get("fte", 1.0)),
                    step=0.5,
                    key=f"{department}_{role_key}_fte",
                )
            with role_cols[2]:
                role["gross_salary"] = st.number_input(
                    cost_label,
                    help=t("gross_help"),
                    min_value=0.0,
                    value=float(role.get("gross_salary", role.get("annual_cost", 0.0))),
                    step=5000.0,
                    key=f"{department}_{role_key}_gross_salary",
                )

            timing_cols = st.columns(2)
            with timing_cols[0]:
                role["start_month"] = st.selectbox(
                    t("effective_month"),
                    MONTHS,
                    index=select_index(MONTHS, role.get("start_month")),
                    key=f"{department}_{role_key}_start_month",
                )
            with timing_cols[1]:
                contract_types = contract_type_options()
                role["contract_type"] = st.selectbox(
                    t("type"),
                    contract_types,
                    index=select_index(contract_types, role.get("contract_type")),
                    key=f"{department}_{role_key}_contract_type",
                )
            employer_cost = employer_social_security_cost(float(role.get("gross_salary", 0.0)), float(role.get("fte", 0.0)))
            st.caption(
                f"{caption_action}: {money(employer_cost)} "
                f"(gross + {SPANISH_EMPLOYER_SS_RATE:.2%} Seguridad Social patronal base)."
            )

def render_department_configured_expenses(department: str, answers: dict, starting_section: int) -> None:
    for offset, section in enumerate(department_expense_sections(department)):
        render_department_named_expense_intake(
            department=department,
            answers=answers,
            section_number=starting_section + offset,
            title=business_label(section["title"]),
            storage_key=section["storage_key"],
            item_label=business_label(section["item_label"]),
            category_label=business_label(section["category_label"]),
            category_options=[business_label(option) for option in section["category_options"]],
            default_count=int(section.get("default_count", 0)),
            link_to_events=bool(section.get("link_to_events", False)),
        )


def render_department_named_expense_intake(
    department: str,
    answers: dict,
    section_number: int,
    title: str,
    storage_key: str,
    item_label: str,
    category_label: str,
    category_options: list[str],
    default_count: int = 0,
    link_to_events: bool = False,
) -> None:
    st.markdown(f"##### {section_number}. {title}")
    scoped_key = f"{department_slug(department)}_{storage_key}"
    count_key = f"_{scoped_key}_count"
    expense_count = st.number_input(
        f"{t('number_of')} {item_label}s",
        min_value=0,
        max_value=20,
        value=int(answers.get(count_key, default_count)),
        step=1,
        key=f"{storage_key}_count_{department}",
    )
    answers[count_key] = int(expense_count)

    for expense_number in range(1, int(expense_count) + 1):
        expense_key = f"_{scoped_key}_{expense_number}"
        expense = answers.setdefault(expense_key, {})
        if not isinstance(expense, dict):
            expense = {}
            answers[expense_key] = expense
        with st.expander(f"{title} {expense_number}", expanded=expense_number == 1):
            first_row = st.columns([2, 1])
            with first_row[0]:
                expense["name"] = st.text_input(
                    f"{item_label.title()} {expense_number}: {t('name')}",
                    value=expense.get("name", ""),
                    key=f"{department}_{expense_key}_name",
                    placeholder=f"Ej. {title}",
                )
            with first_row[1]:
                expense["cost"] = st.number_input(
                    f"{item_label.title()} {expense_number}: {t('estimated_cost')}",
                    min_value=0.0,
                    value=float(expense.get("cost", 0.0)),
                    step=1000.0,
                    key=f"{department}_{expense_key}_cost",
                )

            second_row = st.columns(3)
            with second_row[0]:
                expense["category"] = st.selectbox(
                    category_label,
                    category_options,
                    index=select_index(category_options, expense.get("category")),
                    key=f"{department}_{expense_key}_category",
                )
            with second_row[1]:
                expense["billing_month"] = st.selectbox(
                    t("billing_month"),
                    MONTHS,
                    index=select_index(MONTHS, expense.get("billing_month")),
                    key=f"{department}_{expense_key}_billing_month",
                )
            with second_row[2]:
                frequencies = frequency_options()
                expense["frequency"] = st.selectbox(
                    t("frequency"),
                    frequencies,
                    index=select_index(frequencies, expense.get("frequency")),
                    key=f"{department}_{expense_key}_frequency",
                )
            if link_to_events:
                link_options = [t("not_linked")] + department_expense_link_options(department, answers)
                expense["linked_item"] = st.selectbox(
                    t("linked_to"),
                    link_options,
                    index=select_index(link_options, expense.get("linked_item")),
                    key=f"{department}_{expense_key}_linked_item",
                )


def render_department_workspace_actions(department: str, answers: dict) -> None:
    st.markdown(f"#### {t('actions')}")
    actions = st.columns([1, 1, 4])
    with actions[0]:
        if st.button("Save", type="secondary", use_container_width=True, key=f"save_workspace_{department}"):
            saved_at = save_department_workspace(department, answers, submitted=False)
            st.session_state[f"{department}_workspace_saved_at"] = saved_at
            st.success(f"{t('saved')}: {saved_at}")
    with actions[1]:
        if st.button("Submit", type="primary", use_container_width=True, key=f"submit_workspace_{department}"):
            submitted_at = save_department_workspace(department, answers, submitted=True)
            st.session_state[f"{department}_workspace_submitted_at"] = submitted_at
            # Un nuevo Submit reabre el ciclo de revision: si FP&A lo habia
            # rechazado, ese aviso ya no aplica a esta version actualizada.
            st.session_state.pop(f"fpna_status::{department}", None)
            st.session_state.pop(f"fpna_comment::{department}", None)
            st.success(f"{t('submitted')}: {submitted_at}")
    with actions[2]:
        saved_at = st.session_state.get(f"{department}_workspace_saved_at")
        submitted_at = st.session_state.get(f"{department}_workspace_submitted_at")
        if submitted_at:
            st.caption(f"{t('last_submit')}: {submitted_at}")
        elif saved_at:
            st.caption(f"{t('last_save')}: {saved_at}")


def save_department_workspace(department: str, answers: dict, submitted: bool) -> str:
    SAVED_EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    saved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    payload = {
        "department": department,
        "status": "submitted" if submitted else "draft",
        "saved_at": saved_at,
        "answers": answers,
    }
    department_workspace_path(department).write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return saved_at


def department_workspace_path(department: str) -> Path:
    file_stem = department.lower().replace(" / ", "_").replace(" ", "_").replace("&", "and")
    return SAVED_EVIDENCE_DIR / f"{file_stem}_workspace.json"


def department_slug(department: str) -> str:
    return (
        department.lower()
        .replace(" / ", "_")
        .replace(" & ", "_and_")
        .replace(" ", "_")
        .replace("&", "and")
        .replace("/", "_")
    )


def department_role_options(department: str) -> list[str]:
    return DEPARTMENT_ROLE_OPTIONS.get(department, DEFAULT_ROLE_OPTIONS)


def department_expense_sections(department: str) -> list[dict]:
    if department in DEPARTMENT_EXPENSE_SECTIONS:
        return DEPARTMENT_EXPENSE_SECTIONS[department]
    if department == "Product & Engineering":
        return DEPARTMENT_EXPENSE_SECTIONS["Engineering"]
    return [
        {
            "title": "Software y herramientas",
            "storage_key": "software",
            "item_label": "herramienta",
            "category_label": "Tipo de herramienta",
            "category_options": ["Productividad", "Reporting", "Workflow", "Automation", "Licencia especializada"],
        },
        {
            "title": "Servicios externos",
            "storage_key": "external_services",
            "item_label": "servicio",
            "category_label": "Tipo",
            "category_options": ["Consulting", "Freelance", "Agency", "Implementation", "Advisory"],
        },
        {
            "title": "Others",
            "storage_key": "others",
            "item_label": "gasto",
            "category_label": "Tipo de gasto",
            "category_options": ["Formacion", "Licencias menores", "Suscripciones", "Materiales", "Otro"],
        },
    ]


def department_project_options(answers: dict) -> list[str]:
    department = answers.get("_department", "")
    if not department_needs_clients_projects(department):
        return []
    project_count = int(answers.get("_project_count", 0))
    options = []
    for project_number in range(1, project_count + 1):
        project = answers.get(f"_project_{project_number}", {})
        if not isinstance(project, dict):
            continue
        name = project.get("name") or f"{t('client_project')} {project_number}"
        options.append(f"{project_number}. {name}")
    return options


def department_event_options(department: str, answers: dict) -> list[str]:
    scoped_key = f"{department_slug(department)}_events"
    default_count = default_section_count(department, "events")
    event_count = max(int(answers.get(f"_{scoped_key}_count", 0)), default_count)
    options = []
    for event_number in range(1, event_count + 1):
        event = answers.get(f"_{scoped_key}_{event_number}", {})
        if not isinstance(event, dict):
            event = {}
        name = event.get("name") or f"Evento {event_number}"
        options.append(f"{event_number}. {name}")
    return options


def department_expense_link_options(department: str, answers: dict) -> list[str]:
    project_options = [f"{t('project_prefix')}: {option}" for option in department_project_options(answers)]
    event_options = [f"{t('event_prefix')}: {option}" for option in department_event_options(department, answers)]
    return project_options + event_options


def default_section_count(department: str, storage_key: str) -> int:
    for section in department_expense_sections(department):
        if section["storage_key"] == storage_key:
            return int(section.get("default_count", 0))
    return 0


def department_proposed_budget(department: str, summary: pd.DataFrame, answers: dict) -> float:
    department_summary = summary[summary["Department"] == department]
    baseline = float(department_summary["Suggested Baseline"].sum())
    override_pct = float(answers.get("_override_pct", 0.0))
    algorithmic_proposal = baseline * (1 + override_pct / 100)
    requested_delta = department_requested_plan_total(answers)
    return max(0.0, algorithmic_proposal + requested_delta)


def department_requested_plan_total(answers: dict) -> float:
    department = answers.get("_department", "")
    project_total = department_project_total(answers) if department_needs_clients_projects(department) else 0.0
    return (
        project_total
        + department_headcount_net_total(department, answers)
        + department_configured_expenses_total(department, answers)
    )


def department_project_total(answers: dict) -> float:
    department = answers.get("_department", "")
    if not department_needs_clients_projects(department):
        return 0.0
    project_count = int(answers.get("_project_count", 0))
    project_total = 0.0
    for project_number in range(1, project_count + 1):
        project = answers.get(f"_project_{project_number}", {})
        if not isinstance(project, dict):
            continue
        project_total += float(project.get("cost", 0.0))
    return project_total


def department_headcount_net_total(department: str, answers: dict) -> float:
    if not department:
        return 0.0
    dept_slug = department_slug(department)
    headcount_count = int(answers.get(f"_{dept_slug}_headcount_count", answers.get("_marketing_headcount_count", 0)))
    headcount_total = 0.0
    for role_number in range(1, headcount_count + 1):
        role = answers.get(f"_{dept_slug}_role_{role_number}", answers.get(f"_marketing_role_{role_number}", {}))
        if not isinstance(role, dict):
            continue
        headcount_total += employer_social_security_cost(
            float(role.get("gross_salary", role.get("annual_cost", 0.0))),
            float(role.get("fte", 0.0)),
        )
    reduction_count = int(answers.get(f"_{dept_slug}_headcount_reduction_count", 0))
    headcount_reduction_total = 0.0
    for role_number in range(1, reduction_count + 1):
        role = answers.get(f"_{dept_slug}_reduction_role_{role_number}", {})
        if not isinstance(role, dict):
            continue
        headcount_reduction_total += employer_social_security_cost(
            float(role.get("gross_salary", role.get("annual_cost", 0.0))),
            float(role.get("fte", 0.0)),
        )
    return headcount_total - headcount_reduction_total


def employer_social_security_cost(gross_salary: float, fte: float) -> float:
    return gross_salary * fte * (1 + SPANISH_EMPLOYER_SS_RATE)


def travel_frequency_multiplier(frequency: str | None) -> int:
    if frequency in {"Mensual", "Monthly"}:
        return 12
    if frequency in {"Bimensual", "Bimonthly"}:
        return 6
    if frequency in {"Trimestral", "Quarterly"}:
        return 4
    if frequency in {"Semestral", "Semiannual"}:
        return 2
    return 1


def department_configured_expenses_total(department: str, answers: dict) -> float:
    total = 0.0
    dept_slug = department_slug(department)
    for section in department_expense_sections(department):
        storage_key = section["storage_key"]
        count = int(answers.get(f"_{dept_slug}_{storage_key}_count", 0))
        for expense_number in range(1, count + 1):
            expense = answers.get(f"_{dept_slug}_{storage_key}_{expense_number}", {})
            if not isinstance(expense, dict):
                continue
            total += float(expense.get("cost", 0.0)) * travel_frequency_multiplier(expense.get("frequency"))
    return total


def detected_ytd_month(department: str) -> str:
    actuals = st.session_state.actuals_ytd
    department_actuals = actuals[actuals["Department"] == department]
    if department_actuals.empty:
        return st.session_state.ytd_month
    max_month = int(department_actuals["Month"].max())
    return MONTH_NUMBER_TO_NAME.get(max_month, st.session_state.ytd_month)


def select_index(options: list[str], current: str | None) -> int:
    if current in options:
        return options.index(current)
    return 0


def render_department_workspace_entry(departments: list[str]) -> str | None:
    active = st.session_state.get("active_department_workspace")
    if active in departments:
        top_cols = st.columns([3, 1])
        with top_cols[0]:
            st.caption(f"{t('open_workspace')}: {active}.")
        with top_cols[1]:
            if st.button(t("change_department"), key="change_department", use_container_width=True):
                st.session_state.active_department_workspace = None
                st.rerun()
        return active

    st.write(t("choose_department"))
    for row_start in range(0, len(departments), 3):
        cols = st.columns(3)
        for col, department in zip(cols, departments[row_start : row_start + 3]):
            with col:
                dept_total = department_total(department)
                if st.button(
                    department,
                    key=f"workspace_entry_{department}",
                    use_container_width=True,
                    help=f"{t('open_department_workspace')} {department}. {t('current_approved_budget')}: {money(dept_total)}.",
                ):
                    st.session_state.active_department_workspace = department
                    st.rerun()

    st.info(t("choose_department_continue"))
    return None


def build_department_workspace_list(summary: pd.DataFrame) -> list[str]:
    source_departments = [
        str(department)
        for department in summary["Department"].dropna().unique()
        if str(department).strip()
    ]
    configured_departments = st.session_state.admin_settings.get("departments", STANDARD_DEPARTMENTS)
    ordered = []
    for department in configured_departments + source_departments:
        if department not in ordered:
            ordered.append(department)
    return ordered


def department_total(department: str) -> float:
    if not evidence_loaded():
        return 0.0
    budget = st.session_state.approved_budget
    return float(budget.loc[budget["Department"] == department, "Amount"].sum())


def submitted_department_workspaces(summary: pd.DataFrame) -> dict[str, dict]:
    departments = build_department_workspace_list(summary)
    submitted = {}
    for department in departments:
        session_answers = st.session_state.department_answers.get(department)
        if session_answers and st.session_state.get(f"{department}_workspace_submitted_at"):
            session_answers.setdefault("_department", department)
            submitted[department] = session_answers
            continue

        workspace_path = department_workspace_path(department)
        if not workspace_path.exists():
            continue
        try:
            payload = json.loads(workspace_path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            continue
        if payload.get("status") != "submitted":
            continue
        answers = payload.get("answers", {})
        if isinstance(answers, dict):
            answers.setdefault("_department", department)
            submitted[department] = answers
            st.session_state.department_answers.setdefault(department, answers)
    return submitted


def department_challenge_questions(department: str, summary: pd.DataFrame, answers: dict) -> list[str]:
    department_summary = summary[summary["Department"] == department]
    approved_budget = float(department_summary["Approved Budget"].sum())
    ytd_actuals = float(department_summary["YTD Actuals"].sum())
    run_rate = float(department_summary["Run-rate Forecast"].sum())
    proposed_budget = department_proposed_budget(department, summary, answers)
    requested_plan = department_requested_plan_total(answers)
    project_total = department_project_total(answers)
    expense_total = department_configured_expenses_total(department, answers)
    headcount_net = department_headcount_net_total(department, answers)
    has_clients_projects = department_needs_clients_projects(department)
    # Estas preguntas se generaban solo en español, sin pasar por t() ni por
    # el idioma seleccionado -- por eso se veian en español aunque la UI
    # (Step 1, Step 2...) estuviera en ingles. Cada pregunta lleva ahora su
    # version en los dos idiomas y se elige segun st.session_state.language,
    # igual que el resto de la interfaz.
    is_en = st.session_state.get("language") == "English"
    questions = []

    if proposed_budget > approved_budget * 1.1:
        questions.append(
            f"The proposal is up {money(proposed_budget - approved_budget)} vs the previously approved budget. "
            "Which part is mandatory, which is growth, and which could be phased in?"
            if is_en
            else f"La propuesta sube {money(proposed_budget - approved_budget)} vs el budget aprobado anterior. "
            "Que parte es obligatoria, que parte es crecimiento y que parte se puede fasear?"
        )
    if run_rate > approved_budget * 1.1 and proposed_budget < run_rate * 0.95:
        questions.append(
            f"Annualized YTD actuals point to {money(run_rate)}, above the {money(proposed_budget)} proposal. "
            "What concrete actions will bring that run-rate down?"
            if is_en
            else f"Los actuals YTD anualizados apuntan a {money(run_rate)}, por encima de la propuesta de {money(proposed_budget)}. "
            "Que acciones concretas van a reducir ese run-rate?"
        )
    if proposed_budget > run_rate * 1.1:
        if has_clients_projects:
            questions.append(
                f"The proposal is {money(proposed_budget - run_rate)} above the YTD run-rate. "
                "Which clients/projects, hires or new commitments explain that increase?"
                if is_en
                else f"La propuesta queda {money(proposed_budget - run_rate)} por encima del run-rate YTD. "
                "Que clientes/proyectos, contrataciones o compromisos nuevos explican ese incremento?"
            )
        else:
            questions.append(
                f"The proposal is {money(proposed_budget - run_rate)} above the YTD run-rate. "
                "Which hires, contracts, obligations or operational changes explain that increase?"
                if is_en
                else f"La propuesta queda {money(proposed_budget - run_rate)} por encima del run-rate YTD. "
                "Que contrataciones, contratos, obligaciones o cambios operativos explican ese incremento?"
            )
    if requested_plan and abs(requested_plan) > max(approved_budget * 0.15, 25000):
        questions.append(
            f"The submitted plan adds a net impact of {money(requested_plan)}. "
            "Which items could be prioritized if FP&A asks for a 10% cut?"
            if is_en
            else f"El plan enviado anade un impacto neto de {money(requested_plan)}. "
            "Que elementos se pueden priorizar si FP&A pide recortar un 10%?"
        )
    if has_clients_projects and project_total:
        questions.append(
            f"Clients and projects add up to {money(project_total)}. "
            "Are they tied to revenue, renewals, pipeline or a measurable operational commitment?"
            if is_en
            else f"Clientes y proyectos suman {money(project_total)}. "
            "Estan vinculados a ingresos, renovaciones, pipeline o una obligacion operativa medible?"
        )
    if headcount_net:
        questions.append(
            f"The net headcount impact is {money(headcount_net)} including employer Social Security. "
            "What volume, SLA or capacity shows this headcount is needed?"
            if is_en
            else f"El impacto neto de personal es {money(headcount_net)} incluyendo Seguridad Social patronal. "
            "Que volumen, SLA o capacidad demuestra que esta plantilla es necesaria?"
        )
    if expense_total and (not has_clients_projects or not project_total):
        questions.append(
            f"Departmental expenses add up to {money(expense_total)}. "
            "Which expense should stay as recurring run-rate and which could be approved case by case?"
            if is_en
            else f"Los gastos departamentales suman {money(expense_total)}. "
            "Que gasto debe quedar como run-rate recurrente y cual podria aprobarse caso a caso?"
        )
    if not questions:
        questions.append(
            f"{department} is aligned with the previous approved budget and YTD run-rate. "
            "Confirm there are no risks, signed commitments or scope changes missing from the submission."
            if is_en
            else f"{department} esta alineado con budget anterior y run-rate YTD. "
            "Confirma que no hay riesgos, compromisos firmados o cambios de alcance que falten en el submit."
        )
    if ytd_actuals == 0 and proposed_budget:
        questions.append(
            "No YTD actuals are loaded for this department, but there is a proposal. "
            "Is this a new department, or is actuals mapping missing?"
            if is_en
            else "No hay actuals YTD cargados para este departamento, pero si propuesta. "
            "Se trata de un departamento nuevo o falta mapeo de actuals?"
        )
    # Targets top-down definidos en Settings (ventas, margen, EBITDA,
    # plantilla, cash...) -- si hay alguno con valor, el challenge tambien
    # pregunta si la propuesta de este departamento ayuda o contradice
    # esos objetivos de compania, no solo si cuadra con su propio
    # historico/run-rate.
    targets = [
        target
        for target in st.session_state.admin_settings.get("corporate_targets", [])
        if isinstance(target, dict) and str(target.get("target", "")).strip() and str(target.get("value", "")).strip()
    ]
    if targets:
        targets_text = "; ".join(
            f"{target['target']}: {target['value']} {target.get('unit', '')}".strip()
            for target in targets
        )
        questions.append(
            f"Company-wide targets for this cycle: {targets_text}. "
            "Does this department's proposal support these targets, or does it work against any of them?"
            if is_en
            else f"Objetivos de compania para este ciclo: {targets_text}. "
            "La propuesta de este departamento ayuda a estos objetivos, o va en contra de alguno?"
        )
    return questions


def render_budget_challenge() -> None:
    st.subheader(t("challenge_title"))
    if not require_evidence():
        return

    st.write(t("challenge_intro"))
    summary = build_gap_summary()
    submitted = submitted_department_workspaces(summary)
    if not submitted:
        st.info(t("no_submitted_departments"))
        return

    review_rows = []
    for department, answers in submitted.items():
        department_summary = summary[summary["Department"] == department]
        approved_budget = float(department_summary["Approved Budget"].sum())
        ytd_actuals = float(department_summary["YTD Actuals"].sum())
        run_rate = float(department_summary["Run-rate Forecast"].sum())
        requested_plan = department_requested_plan_total(answers)
        proposed_budget = department_proposed_budget(department, summary, answers)
        review_rows.append(
            {
                "Department": department,
                t("review_status"): t("ready_for_review"),
                t("approved_budget"): approved_budget,
                t("ytd_actuals"): ytd_actuals,
                t("run_rate_forecast"): run_rate,
                t("requested_plan"): requested_plan,
                t("proposed_budget"): proposed_budget,
                t("delta_vs_budget"): proposed_budget - approved_budget,
                t("delta_vs_run_rate"): proposed_budget - run_rate,
            }
        )

    review_df = pd.DataFrame(review_rows)
    amount_columns = [
        t("approved_budget"),
        t("ytd_actuals"),
        t("run_rate_forecast"),
        t("requested_plan"),
        t("proposed_budget"),
        t("delta_vs_budget"),
        t("delta_vs_run_rate"),
    ]
    st.markdown(f"#### {t('submitted_departments')}")
    st.dataframe(format_named_amount_columns(review_df, amount_columns), use_container_width=True, hide_index=True)

    for department, answers in submitted.items():
        with st.expander(department, expanded=True):
            questions = department_challenge_questions(department, summary, answers)
            st.markdown(f"##### {t('challenge_questions')}")
            for idx, question in enumerate(questions):
                # Claves por departamento + indice de pregunta: la respuesta
                # guardada (answer_key) la usa tambien department_proposal_note
                # para construir la nota de "Department's Note" con lo que se
                # ha rellenado aqui en el Step 4.
                answer_key = f"challenge_answer::{department}::{idx}"
                status_key = f"challenge_status::{department}::{idx}"
                show_input_key = f"challenge_show_input::{department}::{idx}"

                st.write(f"- {question}")

                saved_answer = st.session_state.get(answer_key, "")
                status = st.session_state.get(status_key)
                if status == "approved":
                    st.success(t("challenge_approved"))
                elif status == "modify_requested":
                    st.info(t("challenge_needs_changes"))
                if saved_answer:
                    st.caption(f"{t('challenge_answer_label')}: {saved_answer}")

                button_cols = st.columns([1, 1, 1, 3])
                with button_cols[0]:
                    if st.button(t("add_answer"), key=f"{answer_key}_toggle", use_container_width=True):
                        st.session_state[show_input_key] = not st.session_state.get(show_input_key, False)
                with button_cols[1]:
                    if st.button(
                        t("approve"),
                        key=f"{status_key}_approve",
                        type="primary" if status == "approved" else "secondary",
                        use_container_width=True,
                    ):
                        st.session_state[status_key] = "approved"
                        st.rerun()
                with button_cols[2]:
                    # Modificar manda al responsable al Step 3 (Workspaces)
                    # de este departamento, para que ajuste ahi los datos que
                    # explican la variacion, en vez de responder por escrito.
                    if st.button(t("modify"), key=f"{status_key}_modify", use_container_width=True):
                        st.session_state[status_key] = "modify_requested"
                        st.session_state.active_department_workspace = department
                        st.session_state.active_step = "workspaces"
                        st.rerun()

                if st.session_state.get(show_input_key, False):
                    new_value = st.text_area(
                        t("challenge_answer_label"),
                        value=saved_answer,
                        key=f"{answer_key}_input",
                        height=100,
                    )
                    if st.button(t("save_answer"), key=f"{answer_key}_save"):
                        st.session_state[answer_key] = new_value
                        st.session_state[show_input_key] = False
                        st.rerun()
                st.divider()


def render_proposed_budget() -> None:
    st.subheader(t("proposal_title"))
    if not require_evidence():
        return

    proposal = build_proposed_budget()
    monthly_proposal = build_monthly_proposed_budget(proposal)
    summary = proposal.groupby("Department", as_index=False).agg(
        **{
            "Approved Budget": ("Approved Budget", "sum"),
            "Run-rate Forecast": ("Run-rate Forecast", "sum"),
            "Proposed Budget": ("Proposed Budget", "sum"),
            "Variance vs Approved": ("Variance vs Approved", "sum"),
        }
    )
    summary["Variance %"] = safe_divide(summary["Variance vs Approved"], summary["Approved Budget"])
    # "Department's Note" va al final de la tabla (despues de Variance %),
    # no justo antes -- es la misma nota por departamento que ya calculo
    # build_proposed_budget en `proposal`, solo se recupera una vez por
    # departamento.
    notes_by_department = proposal.drop_duplicates("Department").set_index("Department")["Department's Note"]
    summary["Department's Note"] = summary["Department"].map(notes_by_department)

    cols = st.columns(4)
    cols[0].metric(t("approved_budget"), money(summary["Approved Budget"].sum()))
    cols[1].metric(t("run_rate_forecast"), money(summary["Run-rate Forecast"].sum()))
    cols[2].metric(t("proposed_budget"), money(summary["Proposed Budget"].sum()))
    cols[3].metric(t("change_vs_approved"), pct(safe_divide(summary["Variance vs Approved"].sum(), summary["Approved Budget"].sum())))

    chart_cols = st.columns(2)
    with chart_cols[0]:
        treemap = render_proposal_treemap(summary)
        if treemap is not None:
            st.altair_chart(treemap, use_container_width=True)
    with chart_cols[1]:
        st.altair_chart(render_proposal_budget_chart(summary), use_container_width=True)
        st.caption(t("run_rate_forecast_explainer").format(month=st.session_state.ytd_month))
    st.dataframe(format_amount_columns(summary), use_container_width=True, hide_index=True, key="proposal_department_summary")

    st.markdown(f"#### {t('fpna_review_title')}")
    st.caption(t("fpna_review_intro"))
    for _, row in summary.iterrows():
        department = str(row["Department"])
        with st.expander(department, expanded=False):
            st.write(
                f"**{t('proposed_budget')}:** {money(row['Proposed Budget'])}  ·  "
                f"**{t('delta_vs_budget')}:** {money(row['Variance vs Approved'])} ({pct(row['Variance %'])})"
            )
            note = row.get("Department's Note", "")
            if note:
                st.caption(note)

            fpna_status_key = f"fpna_status::{department}"
            fpna_comment_key = f"fpna_comment::{department}"
            status = st.session_state.get(fpna_status_key)
            if status == "approved":
                st.success(t("fpna_approved_caption"))
            elif status == "rejected":
                st.error(t("fpna_rejected_caption"))

            comment_value = st.text_area(
                t("fpna_comment_label"),
                value=st.session_state.get(fpna_comment_key, ""),
                key=f"{fpna_comment_key}_input",
                height=70,
            )

            fpna_action_cols = st.columns([1, 1, 3])
            with fpna_action_cols[0]:
                if st.button(
                    t("fpna_approval"),
                    key=f"fpna_approve_{department}",
                    type="primary" if status == "approved" else "secondary",
                    use_container_width=True,
                ):
                    st.session_state[fpna_status_key] = "approved"
                    st.session_state[fpna_comment_key] = comment_value
                    st.rerun()
            with fpna_action_cols[1]:
                # Rejection manda al departamento de vuelta al Step 3
                # (Workspaces) -- igual que "Modify" en el Step 4 -- y deja
                # una notificacion (fpna_status + fpna_comment) que se
                # muestra ahi arriba del workspace hasta que el
                # departamento vuelva a hacer Submit.
                if st.button(t("fpna_rejection"), key=f"fpna_reject_{department}", use_container_width=True):
                    st.session_state[fpna_status_key] = "rejected"
                    st.session_state[fpna_comment_key] = comment_value
                    st.session_state.active_department_workspace = department
                    st.session_state.active_step = "workspaces"
                    st.rerun()

    with st.expander(t("line_detail"), expanded=False):
        st.dataframe(format_amount_columns(monthly_proposal), use_container_width=True, hide_index=True, key="proposal_line_summary")

    download_cols = st.columns([1, 1, 3])
    with download_cols[0]:
        st.download_button(
            t("download_csv"),
            data=to_csv_bytes(monthly_proposal),
            file_name=f"proposed_budget_{st.session_state.budget_year}.csv",
            mime="text/csv",
            type="primary",
            use_container_width=True,
        )
    with download_cols[1]:
        st.download_button(
            t("download_excel"),
            data=to_excel_bytes(summary, monthly_proposal, budget_year=st.session_state.budget_year),
            file_name=f"proposed_budget_{st.session_state.budget_year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def render_admin_settings() -> None:
    st.subheader(t("settings_title"))
    settings = st.session_state.admin_settings
    st.caption(t("settings_caption"))

    master_cols = st.columns(2)
    with master_cols[0]:
        departments = st.text_area(
            t("visible_departments"),
            value="\n".join(settings.get("departments", [])),
            height=220,
            key="settings_departments",
        )
        cost_centers = st.text_area(
            t("cost_centers"),
            value="\n".join(settings.get("cost_centers", [])),
            height=160,
            key="settings_cost_centers",
        )
    with master_cols[1]:
        accounts = st.text_area(
            t("accounts"),
            value="\n".join(settings.get("accounts", [])),
            height=160,
            key="settings_accounts",
        )
        categories = st.text_area(
            t("categories"),
            value="\n".join(settings.get("categories", [])),
            height=220,
            key="settings_categories",
        )

    st.markdown(f"#### {t('extra_dimensions')}")
    extra_dimensions = settings.get("extra_dimensions", [])
    dimension_names_value = "\n".join(str(item.get("dimension", "")) for item in extra_dimensions if isinstance(item, dict))
    dimension_values_value = "\n".join(str(item.get("values", "")) for item in extra_dimensions if isinstance(item, dict))
    dimension_cols = st.columns(2)
    with dimension_cols[0]:
        dimension_names = st.text_area(
            t("dimension_names"),
            value=dimension_names_value,
            height=160,
            key="settings_dimension_names",
        )
    with dimension_cols[1]:
        dimension_values = st.text_area(
            t("dimension_values"),
            value=dimension_values_value,
            height=160,
            key="settings_dimension_values",
        )

    st.markdown(f"#### {t('smart_mapping')}")
    st.caption(t("smart_mapping_caption"))
    mapping_df = pd.DataFrame(settings.get("smart_mapping", []))
    if mapping_df.empty:
        mapping_df = pd.DataFrame([{"keyword": "", "department": "", "cost_center": "", "account": "", "category": ""}])
    mapping_df = st.data_editor(
        mapping_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="settings_smart_mapping",
    )

    st.markdown(f"#### {t('corporate_targets')}")
    st.caption(t("corporate_targets_caption"))
    targets_df = pd.DataFrame(settings.get("corporate_targets", []))
    if targets_df.empty:
        targets_df = pd.DataFrame([{"target": "", "value": "", "unit": "", "note": ""}])
    targets_df = st.data_editor(
        targets_df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        key="settings_corporate_targets",
    )

    action_cols = st.columns([1, 5])
    with action_cols[0]:
        if st.button(t("save_settings"), type="primary", use_container_width=True):
            updated_settings = {
                "departments": lines_to_list(departments),
                "cost_centers": lines_to_list(cost_centers),
                "accounts": lines_to_list(accounts),
                "categories": lines_to_list(categories),
                "extra_dimensions": dimensions_from_text(dimension_names, dimension_values),
                "smart_mapping": clean_records(mapping_df, ["keyword", "department", "cost_center", "account", "category"]),
                "corporate_targets": clean_records(targets_df, ["target", "value", "unit", "note"]),
            }
            save_admin_settings(updated_settings)
            st.session_state.admin_settings = updated_settings
            st.success(t("settings_saved"))


def read_uploaded_table(uploaded_file) -> pd.DataFrame:
    if uploaded_file.name.lower().endswith(".csv"):
        return pd.read_csv(uploaded_file)
    try:
        return pd.read_excel(uploaded_file)
    except ImportError:
        st.error("XLSX upload requires openpyxl in the Streamlit environment. Please use CSV for this MVP.")
        st.stop()


def render_column_mapping(df: pd.DataFrame, prefix: str) -> ColumnMap:
    options = [""] + list(df.columns)
    st.markdown(f"**{prefix.title()} file**")
    cols = st.columns(6)
    guesses = guess_columns(df.columns)
    with cols[0]:
        department = st.selectbox("Department", options, index=options.index(guesses.department or "") if (guesses.department or "") in options else 0, key=f"{prefix}_department")
    with cols[1]:
        cost_center = st.selectbox("Cost center", options, index=options.index(guesses.cost_center or "") if (guesses.cost_center or "") in options else 0, key=f"{prefix}_cost_center")
    with cols[2]:
        account = st.selectbox("Account", options, index=options.index(guesses.account or "") if (guesses.account or "") in options else 0, key=f"{prefix}_account")
    with cols[3]:
        category = st.selectbox("Category", options, index=options.index(guesses.category or "") if (guesses.category or "") in options else 0, key=f"{prefix}_category")
    with cols[4]:
        month = st.selectbox("Month", options, index=options.index(guesses.month or "") if (guesses.month or "") in options else 0, key=f"{prefix}_month")
    with cols[5]:
        amount = st.selectbox("Amount", options, index=options.index(guesses.amount or "") if (guesses.amount or "") in options else 0, key=f"{prefix}_amount")
    return ColumnMap(department or None, cost_center or None, account or None, category or None, month or None, amount or None)


def render_combined_column_mapping(df: pd.DataFrame) -> CombinedColumnMap:
    options = [""] + list(df.columns)
    guesses = guess_combined_columns(df)
    numeric_columns = [column for column in df.columns if pd.api.types.is_numeric_dtype(df[column])]

    cols = st.columns(4)
    with cols[0]:
        month = st.selectbox("Month/date", options, index=options.index(guesses.month or "") if (guesses.month or "") in options else 0, key="combined_month")
    with cols[1]:
        scenario = st.selectbox("Scenario/type column", options, index=options.index(guesses.scenario or "") if (guesses.scenario or "") in options else 0, key="combined_scenario")
    with cols[2]:
        department = st.selectbox("Department", options, index=options.index(guesses.department or "") if (guesses.department or "") in options else 0, key="combined_department")
    with cols[3]:
        cost_center = st.selectbox("Cost center", options, index=options.index(guesses.cost_center or "") if (guesses.cost_center or "") in options else 0, key="combined_cost_center")

    amount_cols = st.columns(3)
    with amount_cols[0]:
        amount = st.selectbox("Single amount column", options, index=options.index(guesses.amount or "") if (guesses.amount or "") in options else 0, key="combined_amount")
    with amount_cols[1]:
        budget_amount = st.selectbox("Budget amount column", options, index=options.index(guesses.budget_amount or "") if (guesses.budget_amount or "") in options else 0, key="combined_budget_amount")
    with amount_cols[2]:
        actual_amount = st.selectbox("Actual amount column", options, index=options.index(guesses.actual_amount or "") if (guesses.actual_amount or "") in options else 0, key="combined_actual_amount")

    value_columns = st.multiselect(
        "Wide value columns",
        list(df.columns),
        default=list(guesses.value_columns) or numeric_columns,
        help="Use this for files like the Kaggle sample, where each cost type is its own numeric column.",
        key="combined_value_columns",
    )

    return CombinedColumnMap(
        department=department or None,
        cost_center=cost_center or None,
        account=None,
        month=month or None,
        scenario=scenario or None,
        amount=amount or None,
        budget_amount=budget_amount or None,
        actual_amount=actual_amount or None,
        value_columns=tuple(value_columns),
    )


def guess_columns(columns: Iterable[str]) -> ColumnMap:
    normalized = {str(column).lower().replace("_", " ").strip(): column for column in columns}

    def pick(*candidates: str) -> str | None:
        for candidate in candidates:
            for key, original in normalized.items():
                if candidate in key:
                    return original
        return None

    return ColumnMap(
        department=pick("department", "dept", "area", "function"),
        cost_center=pick("cost center", "cost centre", "cc", "centro"),
        account=pick("account", "gl", "cuenta"),
        category=pick("category", "expense type", "line item", "categoria"),
        month=pick("month", "period", "mes"),
        amount=pick("amount", "value", "importe", "actual", "budget"),
    )


def guess_combined_columns(df: pd.DataFrame) -> CombinedColumnMap:
    base = guess_columns(df.columns)
    lowered = {str(column).lower().replace("_", " ").strip(): column for column in df.columns}

    def pick(*candidates: str) -> str | None:
        for candidate in candidates:
            for key, original in lowered.items():
                if candidate in key:
                    return original
        return None

    scenario = pick("scenario", "version", "type", "expenses")
    budget_amount = pick("budget amount", "budget")
    actual_amount = pick("actual amount", "actual")
    value_columns = tuple(
        column
        for column in df.columns
        if pd.api.types.is_numeric_dtype(df[column])
        and column not in {base.month, budget_amount, actual_amount}
    )
    return CombinedColumnMap(
        department=base.department,
        cost_center=base.cost_center,
        account=base.account,
        month=base.month,
        scenario=scenario,
        amount=None if budget_amount or actual_amount else base.amount,
        budget_amount=budget_amount if budget_amount != actual_amount else None,
        actual_amount=actual_amount,
        value_columns=value_columns,
    )


def default_kaggle_combined_map(df: pd.DataFrame) -> CombinedColumnMap:
    guesses = guess_combined_columns(df)
    return CombinedColumnMap(
        department=None,
        cost_center=None,
        account=None,
        month=guesses.month or "Month",
        scenario=guesses.scenario or "Expenses",
        amount=None,
        budget_amount=None,
        actual_amount=None,
        value_columns=guesses.value_columns,
    )


def normalize_table(df: pd.DataFrame, mapping: ColumnMap, annual_source: bool) -> pd.DataFrame:
    missing = []
    if not mapping.department:
        missing.append("Department")
    if not mapping.amount:
        missing.append("Amount")
    if missing:
        st.error(f"Missing required mapping: {', '.join(missing)}")
        st.stop()

    clean = pd.DataFrame()
    clean["Department"] = df[mapping.department].fillna("Unassigned").astype(str)
    clean["Cost Center"] = source_or_default(df, mapping.cost_center, "Unassigned")
    clean["Account"] = source_or_default(df, mapping.account, "Unspecified")
    clean["Category"] = source_or_default(df, mapping.category, "Unspecified")
    clean["Amount"] = pd.to_numeric(df[mapping.amount], errors="coerce").fillna(0.0)

    if mapping.month:
        clean["Month"] = df[mapping.month].map(normalize_month).fillna(1)
    elif annual_source:
        clean["Month"] = 12
    else:
        clean["Month"] = MONTH_TO_NUMBER[st.session_state.ytd_month]

    clean["Type"] = "Approved Budget" if annual_source else "Actuals YTD"
    return apply_smart_mapping(clean)


def normalize_combined_table(df: pd.DataFrame, mapping: CombinedColumnMap) -> tuple[pd.DataFrame, pd.DataFrame]:
    if mapping.value_columns:
        normalized = normalize_combined_wide_table(df, mapping)
    else:
        normalized = normalize_combined_long_table(df, mapping)

    budget = normalized[normalized["Type"] == "Approved Budget"].copy()
    actuals = normalized[normalized["Type"] == "Actuals YTD"].copy()
    ytd_month_number = MONTH_TO_NUMBER[st.session_state.ytd_month]
    actuals = actuals[actuals["Month"] <= ytd_month_number]
    if budget.empty or actuals.empty:
        st.error("The combined file did not produce both Budget and Actual rows. Check scenario mapping or amount columns.")
        st.stop()
    return budget, actuals


def normalize_combined_wide_table(df: pd.DataFrame, mapping: CombinedColumnMap) -> pd.DataFrame:
    if not mapping.month or not mapping.scenario:
        st.error("Combined wide files require Month/date and Scenario/type mapping.")
        st.stop()
    id_columns = [column for column in [mapping.department, mapping.cost_center, mapping.account, mapping.month, mapping.scenario] if column]
    melted = df.melt(id_vars=id_columns, value_vars=list(mapping.value_columns), var_name="Category", value_name="Amount")
    return normalize_combined_rows(
        melted,
        department_col=mapping.department,
        cost_center_col=mapping.cost_center,
        account_col=mapping.account,
        category_col="Category",
        month_col=mapping.month,
        scenario_col=mapping.scenario,
        amount_col="Amount",
    )


def normalize_combined_long_table(df: pd.DataFrame, mapping: CombinedColumnMap) -> pd.DataFrame:
    if mapping.budget_amount and mapping.actual_amount:
        budget = df.copy()
        budget["_Scenario"] = "Budget"
        budget["_Amount"] = pd.to_numeric(budget[mapping.budget_amount], errors="coerce").fillna(0)
        actuals = df.copy()
        actuals["_Scenario"] = "Actual"
        actuals["_Amount"] = pd.to_numeric(actuals[mapping.actual_amount], errors="coerce").fillna(0)
        combined = pd.concat([budget, actuals], ignore_index=True)
        return normalize_combined_rows(
            combined,
            department_col=mapping.department,
            cost_center_col=mapping.cost_center,
            account_col=mapping.account,
            category_col=None,
            month_col=mapping.month,
            scenario_col="_Scenario",
            amount_col="_Amount",
        )
    if not mapping.scenario or not mapping.amount:
        st.error("Combined long files need either Budget/Actual amount columns, or one Scenario column plus one Amount column.")
        st.stop()
    return normalize_combined_rows(
        df,
        department_col=mapping.department,
        cost_center_col=mapping.cost_center,
        account_col=mapping.account,
        category_col=None,
        month_col=mapping.month,
        scenario_col=mapping.scenario,
        amount_col=mapping.amount,
    )


def normalize_combined_rows(
    df: pd.DataFrame,
    department_col: str | None,
    cost_center_col: str | None,
    account_col: str | None,
    category_col: str | None,
    month_col: str | None,
    scenario_col: str,
    amount_col: str,
) -> pd.DataFrame:
    clean = pd.DataFrame()
    clean["Cost Center"] = source_or_default(df, cost_center_col, "Unassigned")
    clean["Account"] = source_or_default(df, account_col, "Unspecified")
    clean["Category"] = source_or_default(df, category_col, "Unspecified")
    if department_col:
        clean["Department"] = source_or_default(df, department_col, "Unassigned")
    else:
        clean["Department"] = clean["Category"].map(infer_department_from_category)
    clean["Amount"] = pd.to_numeric(df[amount_col], errors="coerce").fillna(0.0)
    clean["Month"] = df[month_col].map(normalize_month).fillna(1) if month_col else MONTH_TO_NUMBER[st.session_state.ytd_month]
    clean["Type"] = df[scenario_col].map(normalize_scenario)
    clean = clean[clean["Type"].isin(["Approved Budget", "Actuals YTD"])]
    clean = apply_smart_mapping(clean)
    return clean


def infer_department_from_category(category: str) -> str:
    text = str(category).lower()
    for mapping in st.session_state.admin_settings.get("smart_mapping", []):
        keyword = str(mapping.get("keyword", "")).lower().strip()
        department = str(mapping.get("department", "")).strip()
        if keyword and department and keyword in text:
            return department
    for keyword, department in FALLBACK_DEPARTMENT_BY_KEYWORD.items():
        if keyword in text:
            return department
    return "General & Admin"


def apply_smart_mapping(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    for index, row in clean.iterrows():
        haystack = " ".join(str(row.get(column, "")) for column in ["Department", "Cost Center", "Account", "Category"]).lower()
        for mapping in st.session_state.admin_settings.get("smart_mapping", []):
            keyword = str(mapping.get("keyword", "")).lower().strip()
            if not keyword or keyword not in haystack:
                continue
            mapped_department = str(mapping.get("department", "")).strip()
            mapped_cost_center = str(mapping.get("cost_center", "")).strip()
            mapped_account = str(mapping.get("account", "")).strip()
            mapped_category = str(mapping.get("category", "")).strip()
            if mapped_department:
                clean.at[index, "Department"] = mapped_department
            if mapped_cost_center and row.get("Cost Center") in {"", "Unassigned"}:
                clean.at[index, "Cost Center"] = mapped_cost_center
            if mapped_account and row.get("Account") in {"", "Unspecified"}:
                clean.at[index, "Account"] = mapped_account
            if mapped_category and row.get("Category") in {"", "Unspecified"}:
                clean.at[index, "Category"] = mapped_category
            break
    return clean


def normalize_scenario(value) -> str | None:
    text = str(value).strip().lower()
    if "budget" in text or "plan" in text:
        return "Approved Budget"
    if "actual" in text or "real" in text:
        return "Actuals YTD"
    return None


def source_or_default(df: pd.DataFrame, column: str | None, default: str) -> pd.Series:
    if column and column in df:
        return df[column].fillna(default).astype(str)
    return pd.Series([default] * len(df))


def normalize_month(value) -> int | None:
    if pd.isna(value):
        return None
    if isinstance(value, (int, float)):
        month = int(value)
        return month if 1 <= month <= 12 else None
    text = str(value).strip()[:3].title()
    return MONTH_TO_NUMBER.get(text)


def build_gap_summary() -> pd.DataFrame:
    budget = st.session_state.approved_budget
    actuals = st.session_state.actuals_ytd
    keys = DEFAULT_DIMENSIONS
    budget_summary = budget.groupby(keys, as_index=False)["Amount"].sum().rename(columns={"Amount": "Approved Budget"})
    actuals_summary = actuals.groupby(keys, as_index=False)["Amount"].sum().rename(columns={"Amount": "YTD Actuals"})
    summary = budget_summary.merge(actuals_summary, on=keys, how="outer").fillna({"Approved Budget": 0, "YTD Actuals": 0})

    ytd_month_number = MONTH_TO_NUMBER[st.session_state.ytd_month]
    summary["Run-rate Forecast"] = summary["YTD Actuals"] / ytd_month_number * 12
    summary["Variance vs Approved"] = summary["Run-rate Forecast"] - summary["Approved Budget"]
    summary["Variance %"] = safe_divide(summary["Variance vs Approved"], summary["Approved Budget"])
    summary["Signal"] = summary.apply(classify_signal, axis=1)
    summary["Suggested Adjustment"] = summary.apply(suggest_adjustment, axis=1)
    summary["Suggested Baseline"] = summary["Approved Budget"] + summary["Suggested Adjustment"]
    return summary.sort_values(["Department", "Signal", "Variance vs Approved"], ascending=[True, True, False])


def classify_signal(row: pd.Series) -> str:
    variance_pct = row["Variance %"]
    if row["Approved Budget"] == 0 and row["YTD Actuals"] != 0:
        return "New spend"
    if variance_pct > st.session_state.growth_cap / 100:
        return "Over run-rate"
    if variance_pct < -0.15:
        return "Under run-rate"
    return "Stable"


def suggest_adjustment(row: pd.Series) -> float:
    signal = row["Signal"]
    variance = row["Variance vs Approved"]
    if signal == "New spend":
        return row["Run-rate Forecast"]
    if signal == "Over run-rate":
        return variance * 0.65
    if signal == "Under run-rate":
        return variance * 0.5
    return 0.0


def build_proposed_budget() -> pd.DataFrame:
    proposal = build_gap_summary()
    proposal["Department Override %"] = proposal["Department"].map(
        lambda dept: st.session_state.department_answers.get(dept, {}).get("_override_pct", 0.0)
    )
    proposal["Proposed Budget"] = proposal["Suggested Baseline"] * (1 + proposal["Department Override %"] / 100)
    for department, answers in st.session_state.department_answers.items():
        answers.setdefault("_department", department)
        requested_delta = department_requested_plan_total(answers)
        current_total = float(proposal.loc[proposal["Department"] == department, "Proposed Budget"].sum())
        target_total = max(0.0, current_total + requested_delta)
        if current_total > 0 and target_total != current_total:
            factor = target_total / current_total
            proposal.loc[proposal["Department"] == department, "Proposed Budget"] *= factor
    proposal["Variance vs Approved"] = proposal["Proposed Budget"] - proposal["Approved Budget"]
    # "Department's Note" (antes "FP&A Note") explica la variacion con lo
    # rellenado en el Step 3 (workspace del departamento) y en el Step 4
    # (respuestas guardadas a las preguntas de reto) -- ver
    # department_proposal_note. Se usa .apply en vez de .map porque ahora
    # necesita tambien `proposal` para recalcular esas preguntas.
    proposal["Department's Note"] = proposal["Department"].apply(
        lambda dept: department_proposal_note(dept, proposal)
    )
    return proposal


def build_monthly_proposed_budget(proposal: pd.DataFrame) -> pd.DataFrame:
    # "Line detail" solo muestra el Proposed Budget mes a mes (+ su total
    # anual). Actuals, Approved Budget/Annual Approved Budget (budget del
    # ano anterior), Variance vs Approved y la nota del departamento ya se
    # ven en el resumen agregado de mas arriba y aqui sobraban.
    keys = DEFAULT_DIMENSIONS
    budget_monthly = (
        st.session_state.approved_budget.groupby(keys + ["Month"], as_index=False)["Amount"]
        .sum()
        .rename(columns={"Amount": "Approved Budget"})
    )
    budget_lookup = budget_monthly.set_index(keys + ["Month"])["Approved Budget"].to_dict()

    rows = []
    for _, line in proposal.iterrows():
        line_key = tuple(line[key] for key in keys)
        proposed_annual = float(line.get("Proposed Budget", 0.0))
        monthly_approved = [float(budget_lookup.get((*line_key, month), 0.0)) for month in range(1, 13)]
        monthly_total = sum(monthly_approved)
        if monthly_total:
            weights = [amount / monthly_total for amount in monthly_approved]
        else:
            weights = [1 / 12] * 12

        for month_number, weight in enumerate(weights, start=1):
            proposed_month = proposed_annual * weight
            rows.append(
                {
                    **{key: line[key] for key in keys},
                    "Budget Year": st.session_state.budget_year,
                    "Month": MONTH_NUMBER_TO_NAME[month_number],
                    "Month Number": month_number,
                    "Proposed Budget": proposed_month,
                    "Annual Proposed Budget": proposed_annual,
                }
            )
    return pd.DataFrame(rows)


def department_proposal_note(department: str, summary: pd.DataFrame) -> str:
    """Nota del departamento ("Department's Note" en la tabla de resumen y
    en el Excel). Combina el detalle operativo del Step 3 (plantilla,
    proyectos, gastos) con las respuestas que el responsable de
    departamento haya guardado en el Step 4 (Challenge) a sus preguntas de
    reto -- asi la nota explica la variacion con los datos de los dos
    pasos, no solo con los del Step 3."""
    answers = st.session_state.department_answers.get(department, {})
    position = answers.get("_proposal_position", "")
    has_clients_projects = department_needs_clients_projects(department)
    project_count = int(answers.get("_project_count", 0)) if has_clients_projects else 0
    project_total = department_project_total(answers) if has_clients_projects else 0.0
    dept_slug = department_slug(department)
    current_headcount = float(answers.get(f"_{dept_slug}_current_headcount", 0.0))
    headcount_count = int(answers.get(f"_{dept_slug}_headcount_count", answers.get("_marketing_headcount_count", 0)))
    headcount_total = 0.0
    for role_number in range(1, headcount_count + 1):
        role = answers.get(f"_{dept_slug}_role_{role_number}", answers.get(f"_marketing_role_{role_number}", {}))
        if not isinstance(role, dict):
            continue
        headcount_total += employer_social_security_cost(
            float(role.get("gross_salary", role.get("annual_cost", 0.0))),
            float(role.get("fte", 0.0)),
        )
    reduction_count = int(answers.get(f"_{dept_slug}_headcount_reduction_count", 0))
    headcount_reduction_total = 0.0
    for role_number in range(1, reduction_count + 1):
        role = answers.get(f"_{dept_slug}_reduction_role_{role_number}", {})
        if not isinstance(role, dict):
            continue
        headcount_reduction_total += employer_social_security_cost(
            float(role.get("gross_salary", role.get("annual_cost", 0.0))),
            float(role.get("fte", 0.0)),
        )
    expense_total = department_configured_expenses_total(department, answers)
    net_headcount_total = headcount_total - headcount_reduction_total

    step3_note = ""
    if position and (project_total or headcount_total or headcount_reduction_total or expense_total):
        project_note = (
            f"{t('clients_projects')}: {project_count}; coste clientes/proyectos: {money(project_total)}. "
            if has_clients_projects
            else ""
        )
        step3_note = (
            f"{position}. {project_note}"
            f"Plantilla actual: {current_headcount:g}; altas: {headcount_count} ({money(headcount_total)}); "
            f"bajas: {reduction_count} (-{money(headcount_reduction_total)}); impacto neto personal: {money(net_headcount_total)}. "
            f"Gastos departamentales: {money(expense_total)}."
        )
    elif position:
        step3_note = position

    # Step 4: respuestas que el responsable de departamento haya guardado
    # con el boton "Add answer" en render_budget_challenge, para las mismas
    # preguntas de reto que se le mostraron ahi (mismo indice -> misma
    # pregunta, ver challenge_answer::{department}::{idx}).
    questions = department_challenge_questions(department, summary, answers)
    answered = []
    for idx, question in enumerate(questions):
        response = str(st.session_state.get(f"challenge_answer::{department}::{idx}", "")).strip()
        if response:
            answered.append(f"{question} -> {response}")

    parts = [part for part in [step3_note] if part]
    if answered:
        parts.append(f"{t('challenge_responses_label')} " + " | ".join(answered))
    return " ".join(parts)


def render_gap_chart(summary: pd.DataFrame):
    chart_data = summary.groupby("Department", as_index=False)[["Approved Budget", "Run-rate Forecast"]].sum()
    melted = chart_data.melt("Department", var_name="Measure", value_name="Amount")
    return (
        alt.Chart(melted)
        .mark_bar()
        .encode(
            x=alt.X("Department:N", sort="-y"),
            y=alt.Y("Amount:Q"),
            color=alt.Color("Measure:N", scale=alt.Scale(range=["#2563eb", "#0f766e"])),
            tooltip=["Department", "Measure", alt.Tooltip("Amount:Q", format=",.0f")],
        )
        .properties(height=320)
    )


def render_proposal_budget_chart(summary: pd.DataFrame):
    """Grafico de barras agrupadas (lado a lado, no apiladas) comparando
    por departamento tres cifras: el budget aprobado el ano anterior, el
    run-rate forecast (proyeccion de cierre segun los actuals de este ano)
    y el proposed budget (lo que se pide para el proximo ano)."""
    measures = ["Approved Budget", "Run-rate Forecast", "Proposed Budget"]
    # "Approved Budget" es la misma columna que en el resto de la app (no se
    # renombra ahi -- la usan la tabla Summary, el Excel, etc.), pero aqui
    # se ve junto a dos cifras de este/proximo ano, asi que en ESTE grafico
    # se aclara con una etiqueta que dice de que ano es cada barra.
    display_labels = {
        "Approved Budget": "Approved Budget (prior year)",
        "Run-rate Forecast": "Run-rate Forecast (this year)",
        "Proposed Budget": "Proposed Budget (next year)",
    }
    ordered_labels = [display_labels[measure] for measure in measures]
    melted = summary.melt(
        "Department",
        value_vars=measures,
        var_name="Measure",
        value_name="Amount",
    )
    melted["Measure"] = melted["Measure"].map(display_labels)
    return (
        alt.Chart(melted)
        .mark_bar()
        .encode(
            x=alt.X("Department:N", sort="-y", title=None),
            xOffset=alt.XOffset("Measure:N", sort=ordered_labels),
            y=alt.Y("Amount:Q"),
            # Gris para el budget del ano anterior (referencia/pasado),
            # ambar para el run-rate (proyeccion/aviso) y teal de marca
            # para la propuesta (lo nuevo que se pide) -- tres tonos bien
            # diferenciados entre si.
            color=alt.Color(
                "Measure:N",
                sort=ordered_labels,
                scale=alt.Scale(domain=ordered_labels, range=["#64748b", "#d97706", "#0f766e"]),
                title=None,
            ),
            tooltip=["Department", "Measure", alt.Tooltip("Amount:Q", format=",.0f")],
        )
        .properties(height=320, title="Approved budget vs run-rate forecast vs proposed budget")
    )


# --- Treemap ---------------------------------------------------------------
# Vega-Lite (la base de Altair) no trae un layout de treemap incorporado,
# asi que se calcula aqui mismo con el algoritmo "squarified" clasico
# (Bruls, Huizing, van Wijk 1999) en puro Python -- sin anadir una
# dependencia nueva al proyecto (p.ej. la libreria squarify) -- y despues
# se dibuja con mark_rect + mark_text, dos primitivas que si son nativas
# de Altair.
def _treemap_layout_row(sizes: list[float], x: float, y: float, dx: float, dy: float) -> list[dict]:
    covered = sum(sizes)
    rects = []
    if dx >= dy:
        row_height = covered / dx if dx else 0.0
        cx = x
        for size in sizes:
            width = size / row_height if row_height else 0.0
            rects.append({"x": cx, "y": y, "dx": width, "dy": row_height})
            cx += width
    else:
        row_width = covered / dy if dy else 0.0
        cy = y
        for size in sizes:
            height = size / row_width if row_width else 0.0
            rects.append({"x": x, "y": cy, "dx": row_width, "dy": height})
            cy += height
    return rects


def _treemap_leftover(sizes: list[float], x: float, y: float, dx: float, dy: float) -> tuple[float, float, float, float]:
    covered = sum(sizes)
    if dx >= dy:
        row_height = covered / dx if dx else 0.0
        return x, y + row_height, dx, max(dy - row_height, 0.0)
    row_width = covered / dy if dy else 0.0
    return x + row_width, y, max(dx - row_width, 0.0), dy


def _treemap_worst_ratio(sizes: list[float], x: float, y: float, dx: float, dy: float) -> float:
    rects = _treemap_layout_row(sizes, x, y, dx, dy)
    ratios = [max(r["dx"] / r["dy"], r["dy"] / r["dx"]) for r in rects if r["dx"] > 0 and r["dy"] > 0]
    return max(ratios) if ratios else float("inf")


def _squarified_treemap(sizes: list[float], x: float, y: float, dx: float, dy: float) -> list[dict]:
    sizes = [size for size in sizes if size > 0]
    if not sizes:
        return []
    if len(sizes) == 1 or dx <= 0 or dy <= 0:
        return _treemap_layout_row(sizes, x, y, dx, dy)

    row = sizes[:1]
    remaining = sizes[1:]
    while remaining and _treemap_worst_ratio(row, x, y, dx, dy) >= _treemap_worst_ratio(row + remaining[:1], x, y, dx, dy):
        row.append(remaining.pop(0))

    rects = _treemap_layout_row(row, x, y, dx, dy)
    nx, ny, ndx, ndy = _treemap_leftover(row, x, y, dx, dy)
    return rects + _squarified_treemap(remaining, nx, ny, ndx, ndy)


def render_proposal_treemap(summary: pd.DataFrame):
    """Treemap con un rectangulo por departamento, con area proporcional
    al proposed budget (el presupuesto ya preparado) de ese departamento."""
    data = summary[summary["Proposed Budget"] > 0][["Department", "Proposed Budget"]].sort_values(
        "Proposed Budget", ascending=False
    )
    if data.empty:
        return None

    width, height = 100.0, 100.0
    values = data["Proposed Budget"].astype(float).tolist()
    total = sum(values)
    scaled = [value * (width * height) / total for value in values]
    rects = _squarified_treemap(scaled, 0.0, 0.0, width, height)

    rows = []
    # Paleta cualitativa con tonos repartidos por toda la rueda de color
    # (azul, rojo, ambar, verde, violeta, rosa, cian, gris) para que
    # departamentos vecinos en el treemap no se confundan -- la paleta
    # anterior tenia dos verdes/teales muy parecidos uno al lado del otro.
    palette = ["#2563eb", "#dc2626", "#d97706", "#059669", "#7c3aed", "#db2777", "#0891b2", "#78716c"]
    for i, ((_, record), rect) in enumerate(zip(data.iterrows(), rects)):
        rows.append(
            {
                "Department": record["Department"],
                "Proposed Budget": record["Proposed Budget"],
                "x0": rect["x"],
                "x1": rect["x"] + rect["dx"],
                "y0": rect["y"],
                "y1": rect["y"] + rect["dy"],
                "xc": rect["x"] + rect["dx"] / 2,
                "yc": rect["y"] + rect["dy"] / 2,
                "color": palette[i % len(palette)],
            }
        )
    treemap_df = pd.DataFrame(rows)

    rects_chart = (
        alt.Chart(treemap_df)
        .mark_rect(stroke="#ffffff", strokeWidth=2)
        .encode(
            x=alt.X("x0:Q", axis=None),
            x2="x1:Q",
            y=alt.Y("y0:Q", axis=None),
            y2="y1:Q",
            color=alt.Color("color:N", scale=None, legend=None),
            tooltip=["Department", alt.Tooltip("Proposed Budget:Q", format=",.0f")],
        )
    )
    labels = (
        alt.Chart(treemap_df)
        .mark_text(color="#ffffff", fontWeight="bold", fontSize=12)
        .encode(x="xc:Q", y="yc:Q", text="Department:N")
    )
    return (rects_chart + labels).properties(height=320, title="Proposed budget by department")


def demo_data() -> tuple[pd.DataFrame, pd.DataFrame]:
    rows = [
        ("Sales", "CC100", "6100", "Travel", 24000, 21000),
        ("Sales", "CC100", "6200", "Events", 60000, 72000),
        ("Marketing", "CC200", "6300", "Demand generation", 120000, 88000),
        ("Marketing", "CC200", "6400", "Software", 36000, 41000),
        ("Operations", "CC300", "6500", "Logistics", 180000, 162000),
        ("Operations", "CC310", "6600", "External services", 95000, 104000),
        ("Finance", "CC400", "6700", "Audit and advisory", 70000, 62000),
        ("Finance", "CC400", "6800", "Training", 18000, 7000),
        ("IT", "CC500", "6900", "Cloud infrastructure", 90000, 98000),
        ("IT", "CC500", "6950", "Cyber security", 0, 36000),
    ]
    budget = pd.DataFrame(rows, columns=DEFAULT_DIMENSIONS + ["Amount", "YTD Example"])
    budget = budget.drop(columns=["YTD Example"])
    actuals = pd.DataFrame(rows, columns=DEFAULT_DIMENSIONS + ["Approved Example", "Amount"])
    actuals = actuals.drop(columns=["Approved Example"])
    actuals["Month"] = MONTH_TO_NUMBER[st.session_state.ytd_month]
    budget["Month"] = 12
    budget["Type"] = "Approved Budget"
    actuals["Type"] = "Actuals YTD"
    return budget, actuals


def save_current_evidence() -> str:
    SAVED_EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    saved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    st.session_state.approved_budget.to_csv(SAVED_EVIDENCE_DIR / "approved_budget.csv", index=False)
    st.session_state.actuals_ytd.to_csv(SAVED_EVIDENCE_DIR / "actuals_ytd.csv", index=False)
    manifest = {
        "saved_at": saved_at,
        "budget_rows": int(len(st.session_state.approved_budget)),
        "actual_rows": int(len(st.session_state.actuals_ytd)),
        "departments": int(st.session_state.approved_budget["Department"].nunique()),
    }
    (SAVED_EVIDENCE_DIR / "manifest.json").write_text(json.dumps(manifest, indent=2), encoding="utf-8")
    return saved_at


def load_admin_settings() -> dict:
    if not ADMIN_SETTINGS_PATH.exists():
        return json.loads(json.dumps(DEFAULT_ADMIN_SETTINGS))
    try:
        saved = json.loads(ADMIN_SETTINGS_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return json.loads(json.dumps(DEFAULT_ADMIN_SETTINGS))
    settings = json.loads(json.dumps(DEFAULT_ADMIN_SETTINGS))
    settings.update({key: value for key, value in saved.items() if key in settings})
    return settings


def save_admin_settings(settings: dict) -> None:
    SAVED_EVIDENCE_DIR.mkdir(parents=True, exist_ok=True)
    ADMIN_SETTINGS_PATH.write_text(json.dumps(settings, indent=2), encoding="utf-8")


def lines_to_list(value: str) -> list[str]:
    return [line.strip() for line in value.splitlines() if line.strip()]


def dimensions_from_text(names: str, values: str) -> list[dict[str, str]]:
    name_lines = [line.strip() for line in names.splitlines()]
    value_lines = [line.strip() for line in values.splitlines()]
    length = max(len(name_lines), len(value_lines))
    dimensions = []
    for index in range(length):
        dimension = name_lines[index] if index < len(name_lines) else ""
        allowed_values = value_lines[index] if index < len(value_lines) else ""
        if dimension or allowed_values:
            dimensions.append({"dimension": dimension, "values": allowed_values})
    return dimensions


def clean_records(df: pd.DataFrame, columns: list[str]) -> list[dict[str, str]]:
    if df.empty:
        return []
    clean = df.fillna("").astype(str)
    records = []
    for record in clean.to_dict(orient="records"):
        normalized = {column: record.get(column, "").strip() for column in columns}
        if any(normalized.values()):
            records.append(normalized)
    return records


def load_saved_evidence_if_available() -> None:
    if st.session_state.get("_saved_evidence_checked"):
        return
    st.session_state._saved_evidence_checked = True
    budget_path = SAVED_EVIDENCE_DIR / "approved_budget.csv"
    actuals_path = SAVED_EVIDENCE_DIR / "actuals_ytd.csv"
    if not budget_path.exists() or not actuals_path.exists():
        return
    try:
        st.session_state.approved_budget = pd.read_csv(budget_path)
        st.session_state.actuals_ytd = pd.read_csv(actuals_path)
    except (OSError, pd.errors.ParserError):
        return
    manifest_path = SAVED_EVIDENCE_DIR / "manifest.json"
    if manifest_path.exists():
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            manifest = {}
        st.session_state.last_evidence_saved_at = manifest.get("saved_at")


def evidence_loaded() -> bool:
    return "approved_budget" in st.session_state and "actuals_ytd" in st.session_state


def require_evidence() -> bool:
    if evidence_loaded():
        return True
    st.info("Load approved budget and actuals first, or use the demo finance data.")
    return False


def safe_divide(numerator, denominator):
    if isinstance(denominator, pd.Series):
        return numerator.where(denominator != 0, 0) / denominator.where(denominator != 0, 1)
    return 0 if denominator == 0 else numerator / denominator


def money(value: float) -> str:
    return f"EUR {value:,.0f}"


def pct(value: float) -> str:
    return f"{value:.1%}"


def format_amount_columns(df: pd.DataFrame) -> pd.DataFrame:
    formatted = df.copy()
    for column in formatted.columns:
        if column == "Variance %":
            continue
        if any(token in column for token in ["Budget", "Actuals", "Forecast", "Variance"]):
            formatted[column] = formatted[column].map(money)
    if "Variance %" in formatted:
        formatted["Variance %"] = formatted["Variance %"].map(pct)
    return formatted


def format_named_amount_columns(df: pd.DataFrame, amount_columns: list[str]) -> pd.DataFrame:
    formatted = df.copy()
    for column in amount_columns:
        if column in formatted:
            formatted[column] = formatted[column].map(money)
    return formatted


def to_csv_bytes(df: pd.DataFrame) -> bytes:
    return df.to_csv(index=False).encode("utf-8")


# Paleta de marca de la app (mismos tonos que las variables CSS --finance-*
# usadas en la interfaz), para que el Excel descargable se vea como una
# extension de Budget Intelligence y no como una hoja generica de pandas.
_EXCEL_ACCENT = "#0f766e"
_EXCEL_ACCENT_STRONG = "#115e59"
_EXCEL_SOFT = "#f5f7fb"
_EXCEL_LINE = "#d9e1ec"
_EXCEL_INK = "#1f2937"
_EXCEL_MUTED = "#667085"
_EXCEL_CORAL = "#e94f4f"
_EXCEL_BLUE = "#2563eb"


def _excel_column_kind(column: str) -> str:
    """Mismo criterio que format_amount_columns para decidir si una
    columna es importe, porcentaje o texto, asi la vista en pantalla y el
    Excel descargable quedan siempre coherentes. "Budget Year" y "Month
    Number" se excluyen a mano: contienen la palabra "Budget" pero son un
    año/indice, no un importe (si no, saldrian formateados como "EUR
    2.026")."""
    if column == "Variance %":
        return "pct"
    if column in ("Budget Year", "Month Number"):
        return "text"
    if any(token in column for token in ["Budget", "Actuals", "Forecast", "Variance"]):
        return "money"
    return "text"


def _excel_column_width(column: str, values: pd.Series, kind: str) -> int:
    if kind == "money":
        sample = values.map(lambda v: money(v) if pd.notna(v) else "")
    elif kind == "pct":
        sample = values.map(lambda v: pct(v) if pd.notna(v) else "")
    else:
        sample = values.map(lambda v: "" if pd.isna(v) else str(v))
    longest = int(sample.map(len).max()) if len(sample) else 0
    # Las columnas de texto (p.ej. "Department's Note", que ahora puede
    # incluir tambien las respuestas del Step 4) admiten mas ancho que un
    # importe o un porcentaje antes de recortarse.
    cap = 70 if kind == "text" else 42
    return max(11, min(cap, max(len(str(column)), longest) + 3))


def _write_excel_sheet(
    workbook,
    df: pd.DataFrame,
    sheet_name: str,
    tab_color: str,
    title: str,
    subtitle: str,
    chart_columns: list[str] | None = None,
    pie_column: str | None = None,
    show_title: bool = True,
    row_formulas: dict[str, str] | None = None,
    annual_formula: dict[str, object] | None = None,
) -> None:
    """Escribe `df` como una hoja con la identidad visual de la app:
    banner de titulo (opcional, ver `show_title`), cabecera con color de
    acento, filas alternas, formato de moneda/porcentaje en rojo si es
    negativo, fila de TOTAL, mapa de calor sobre la varianza, autofiltro,
    paneles inmovilizados y, si se piden, un grafico de barras comparando
    `chart_columns` y/o un grafico de tarta con la distribucion de
    `pie_column` por la primera columna de la tabla.

    `row_formulas` y `annual_formula` permiten que, en vez de volcar un
    numero ya calculado en Python, la celda quede como una formula de
    Excel de verdad (para que si alguien edita a mano "Approved Budget" o
    "Proposed Budget", la varianza o el total anual se recalculen solos):

    - `row_formulas`: {columna: plantilla}, una formula por fila que
      combina otras columnas de la MISMA fila. La plantilla usa
      `{col:Nombre columna}` para la letra de esa columna y `{ROW}` para
      la fila de Excel (1-based), p.ej.
      "={col:Proposed Budget}{ROW}-{col:Approved Budget}{ROW}".
    - `annual_formula`: {"target": columna a escribir, "source": columna
      a sumar, "group_by": columnas que identifican una misma "linea"} --
      para cada bloque contiguo de filas que comparten esos valores (p.ej.
      las 12 filas mensuales de una misma linea), escribe un
      =SUM(...) sobre el bloque en vez del importe anual fijo.

    En ambos casos se sigue pasando el valor ya calculado en Python como
    "cached value" de `write_formula`, asi que la celda muestra el numero
    correcto aunque el libro se abra sin recalcular (LibreOffice, algunos
    visores)."""
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.set_tab_color(tab_color)

    n_rows, n_cols = len(df), len(df.columns)
    last_col = max(n_cols - 1, 0)
    header_row = 2 if show_title else 0
    first_data_row = header_row + 1
    last_data_row = first_data_row + n_rows - 1

    title_format = workbook.add_format(
        {
            "bold": True,
            "font_size": 14,
            "font_color": "#ffffff",
            "bg_color": _EXCEL_ACCENT_STRONG,
            "valign": "vcenter",
            "indent": 1,
        }
    )
    subtitle_format = workbook.add_format(
        {
            "italic": True,
            "font_size": 9,
            "font_color": _EXCEL_MUTED,
            "bg_color": _EXCEL_SOFT,
            "valign": "vcenter",
            "indent": 1,
        }
    )
    header_format = workbook.add_format(
        {
            "bold": True,
            "font_color": "#ffffff",
            "bg_color": _EXCEL_ACCENT,
            "border": 1,
            "border_color": _EXCEL_LINE,
            "valign": "vcenter",
            "text_wrap": True,
        }
    )

    # Un formato reutilizable por combinacion (tipo x banda), calculado una
    # sola vez fuera del bucle de filas -- crear un Format nuevo por celda
    # es lento y, con tablas grandes, puede disparar el numero de formatos
    # del libro innecesariamente.
    money_num_format = '"EUR "#,##0;[RED]("EUR "#,##0)'
    pct_num_format = "0.0%;[RED]-0.0%"
    data_formats = {}
    for kind in ("money", "pct", "text"):
        for banded in (False, True):
            opts = {"font_color": _EXCEL_INK}
            if kind == "money":
                opts["num_format"] = money_num_format
            elif kind == "pct":
                opts["num_format"] = pct_num_format
            if banded:
                opts["bg_color"] = _EXCEL_SOFT
            data_formats[(kind, banded)] = workbook.add_format(opts)
    total_formats = {}
    for kind in ("money", "pct", "text"):
        opts = {"font_color": _EXCEL_INK, "bold": True, "top": 2, "top_color": _EXCEL_ACCENT}
        if kind == "money":
            opts["num_format"] = money_num_format
        elif kind == "pct":
            opts["num_format"] = pct_num_format
        total_formats[kind] = workbook.add_format(opts)

    if show_title:
        worksheet.merge_range(0, 0, 0, last_col, title, title_format)
        worksheet.merge_range(1, 0, 1, last_col, subtitle, subtitle_format)
        worksheet.set_row(0, 26)
        worksheet.set_row(1, 18)

    kinds = [_excel_column_kind(column) for column in df.columns]
    for col_idx, column in enumerate(df.columns):
        worksheet.write(header_row, col_idx, column, header_format)
        worksheet.set_column(col_idx, col_idx, _excel_column_width(column, df[column], kinds[col_idx]))
    worksheet.set_row(header_row, 30)

    col_letters = {column: xl_col_to_name(idx) for idx, column in enumerate(df.columns)}

    def resolve_formula(template: str, excel_row_1based: int) -> str:
        resolved = template.replace("{ROW}", str(excel_row_1based))
        for column, letter in col_letters.items():
            resolved = resolved.replace("{col:%s}" % column, letter)
        return resolved

    # Bloques contiguos de filas que forman una misma "linea" (p.ej. las 12
    # filas mensuales de un Department/Cost Center/Account/Category) para
    # poder escribir el importe anual como =SUM(bloque) en vez de un valor
    # fijo repetido en cada fila.
    block_bounds: list[tuple[int, int]] = []
    if annual_formula and n_rows:
        group_cols = annual_formula["group_by"]
        keys = df[group_cols].astype(str).agg("|".join, axis=1).tolist()
        starts = [0]
        for i in range(1, n_rows):
            if keys[i] != keys[i - 1]:
                starts.append(i)
        starts.append(n_rows)
        block_bounds = [None] * n_rows  # type: ignore[list-item]
        for start, end in zip(starts, starts[1:]):
            for i in range(start, end):
                block_bounds[i] = (start, end - 1)

    for row_offset in range(n_rows):
        row = first_data_row + row_offset
        banded = row_offset % 2 == 1
        for col_idx, kind in enumerate(kinds):
            column = df.columns[col_idx]
            value = df.iloc[row_offset, col_idx]
            fmt = data_formats[(kind, banded)]
            if kind in ("money", "pct"):
                cached_value = 0.0 if pd.isna(value) else float(value)
                if annual_formula and column == annual_formula["target"]:
                    source_letter = col_letters[annual_formula["source"]]
                    start, end = block_bounds[row_offset]
                    formula = f"=SUM({source_letter}{first_data_row + start + 1}:{source_letter}{first_data_row + end + 1})"
                    worksheet.write_formula(row, col_idx, formula, fmt, cached_value)
                elif row_formulas and column in row_formulas:
                    formula = resolve_formula(row_formulas[column], row + 1)
                    worksheet.write_formula(row, col_idx, formula, fmt, cached_value)
                else:
                    worksheet.write_number(row, col_idx, cached_value, fmt)
            else:
                worksheet.write(row, col_idx, "" if pd.isna(value) else value, fmt)

    if n_rows:
        total_row = last_data_row + 1
        worksheet.write(total_row, 0, "TOTAL", total_formats["text"])
        for col_idx, kind in enumerate(kinds):
            if col_idx == 0:
                continue
            column = df.columns[col_idx]
            # Las columnas "Annual ..." repiten el mismo importe anual en
            # cada una de las filas mensuales de una misma linea -- sumarlas
            # multiplicaria el total x12 en vez de dar el total real, asi
            # que se dejan en blanco en vez de mostrar un numero enganoso.
            if kind == "money" and not column.startswith("Annual "):
                col_letter = xl_col_to_name(col_idx)
                cached_total = float(df[column].sum())
                worksheet.write_formula(
                    total_row,
                    col_idx,
                    f"=SUM({col_letter}{first_data_row + 1}:{col_letter}{last_data_row + 1})",
                    total_formats["money"],
                    cached_total,
                )
            else:
                worksheet.write(total_row, col_idx, "", total_formats[kind])
        worksheet.autofilter(header_row, 0, last_data_row, last_col)

    worksheet.freeze_panes(first_data_row, 1)

    # Mapa de calor rojo/blanco/verde sobre la varianza: de un vistazo se
    # ve que lineas se disparan al alza (coral) o mejoran (verde), sin
    # tener que leer cada numero.
    for col_idx, column in enumerate(df.columns):
        if column in ("Variance vs Approved", "Variance %") and n_rows:
            worksheet.conditional_format(
                first_data_row,
                col_idx,
                last_data_row,
                col_idx,
                {
                    "type": "3_color_scale",
                    "min_color": _EXCEL_ACCENT,
                    "mid_color": "#ffffff",
                    "max_color": _EXCEL_CORAL,
                },
            )

    if chart_columns and n_rows:
        chart = workbook.add_chart({"type": "column"})
        # Mismos colores y etiquetas que el grafico de barras agrupadas de
        # la web (render_proposal_budget_chart): gris = budget del ano
        # anterior, ambar = proyeccion de este ano, teal de marca = lo que
        # se propone para el proximo ano.
        series_style = {
            "Approved Budget": ("Approved Budget (prior year)", "#64748b"),
            "Run-rate Forecast": ("Run-rate Forecast (this year)", "#d97706"),
            "Proposed Budget": ("Proposed Budget (next year)", "#0f766e"),
        }
        for column in chart_columns:
            if column not in df.columns:
                continue
            col_idx = list(df.columns).index(column)
            label, color = series_style.get(column, (column, _EXCEL_ACCENT))
            chart.add_series(
                {
                    "name": label,
                    "categories": [sheet_name, first_data_row, 0, last_data_row, 0],
                    "values": [sheet_name, first_data_row, col_idx, last_data_row, col_idx],
                    "fill": {"color": color},
                    "gap": 40,
                }
            )
        chart.set_title({"name": "Approved budget vs run-rate forecast vs proposed budget"})
        chart.set_legend({"position": "bottom"})
        chart.set_y_axis({"num_format": '"EUR "#,##0'})
        chart.set_size({"width": 640, "height": 340})
        worksheet.insert_chart(header_row, last_col + 2, chart, {"x_offset": 10})

    if pie_column and pie_column in df.columns and n_rows:
        # Excel/xlsxwriter no tiene un tipo de grafico "treemap" nativo (a
        # diferencia de Vega-Lite en la web) -- un grafico de tarta con un
        # color distinto por departamento cuenta la misma historia (que
        # parte del total representa cada departamento) con un tipo de
        # grafico que Excel si soporta de forma nativa e interactiva.
        pie_col_idx = list(df.columns).index(pie_column)
        palette = ["#2563eb", "#dc2626", "#d97706", "#059669", "#7c3aed", "#db2777", "#0891b2", "#78716c"]
        pie = workbook.add_chart({"type": "pie"})
        pie.add_series(
            {
                "name": pie_column,
                "categories": [sheet_name, first_data_row, 0, last_data_row, 0],
                "values": [sheet_name, first_data_row, pie_col_idx, last_data_row, pie_col_idx],
                "points": [{"fill": {"color": palette[i % len(palette)]}} for i in range(n_rows)],
                "data_labels": {"category": True, "percentage": True, "position": "outside_end"},
            }
        )
        pie.set_title({"name": f"{pie_column} by department"})
        pie.set_legend({"position": "bottom"})
        pie.set_size({"width": 480, "height": 340})
        # Debajo del grafico de barras (18 filas ~ su alto en pixeles) para
        # que no se solapen.
        worksheet.insert_chart(header_row + 18, last_col + 2, pie, {"x_offset": 10})


def to_excel_bytes(summary: pd.DataFrame, proposal: pd.DataFrame, budget_year: int | None = None) -> bytes:
    generated_at = datetime.now().strftime("%d/%m/%Y %H:%M")
    year_label = f" {budget_year}" if budget_year else ""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        workbook = writer.book
        workbook.set_properties(
            {
                "title": f"Proposed Budget{year_label}",
                "subject": "Budget Intelligence",
                "author": "Budget Intelligence",
                "company": "Budget Intelligence",
            }
        )
        _write_excel_sheet(
            workbook,
            summary,
            "Summary",
            _EXCEL_ACCENT,
            title=f"Presupuesto propuesto{year_label}",
            subtitle=f"Generado el {generated_at} · Budget Intelligence",
            chart_columns=["Approved Budget", "Run-rate Forecast", "Proposed Budget"],
            pie_column="Proposed Budget",
            row_formulas={
                # Asi, si alguien ajusta a mano un Approved/Proposed Budget
                # en el Excel, la varianza (importe y %) se recalcula sola
                # en vez de quedarse con el numero que traia de la app.
                "Variance vs Approved": "={col:Proposed Budget}{ROW}-{col:Approved Budget}{ROW}",
                "Variance %": (
                    "=IF({col:Approved Budget}{ROW}=0,0,"
                    "{col:Variance vs Approved}{ROW}/{col:Approved Budget}{ROW})"
                ),
            },
        )
        _write_excel_sheet(
            workbook,
            proposal,
            "Line detail",
            _EXCEL_BLUE,
            title="Detalle mensual por linea",
            subtitle=f"Generado el {generated_at} · Budget Intelligence",
            show_title=False,
            annual_formula={
                # Las 12 filas mensuales de una misma linea son contiguas
                # (ver build_monthly_proposed_budget), asi que el total
                # anual puede ser un =SUM() real sobre ese bloque en vez de
                # repetir el mismo numero fijo en las 12 filas.
                "target": "Annual Proposed Budget",
                "source": "Proposed Budget",
                "group_by": DEFAULT_DIMENSIONS + ["Budget Year"],
            },
        )
    return buffer.getvalue()


if __name__ == "__main__":
    main()
