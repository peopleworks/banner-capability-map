#  PUBLIC CATALOG. The capability definitions and the guidance notes, with every
#  institution-specific finding removed: what to look at and how to read it, never
#  what one school found. port_to_public.py grafts this onto the internal engine.
#  Edit THIS file to change the public catalog; never edit public_draft/bcm_report.py
#  by hand, it is generated.

CATALOG = [
    dict(name="Fixed Assets", buys="asset-inventory system",
         forms="FFA%", owner="FIMSMGR", tbl="FFBMAST", tbls=["FIMSMGR.FF%"],
         note="The asset master (FFBMAST) holds tracked, depreciated assets. Banner auto-flags "
              "asset candidates from purchasing into FFBOTAG; if those are never converted into "
              "FFBMAST records, asset accounting is not being done even though origination tags exist."),
    dict(name="Event Management", buys="EMS / 25Live (events, room scheduling)",
         forms="SLA%", owner="SATURN", tbl="SLBEVNT", tbls=["SATURN.SLB%"],
         note="Banner's own event and room-booking module. If SLBEVNT is empty or last touched "
              "years ago, event booking is handled elsewhere or not at all. Buildings and rooms may "
              "still be defined for class scheduling, so read the SLBEVNT row as the judge."),
    dict(name="HR Recruitment / Applicant Tracking", buys="NEOGOV / PageUp / Cornerstone",
         forms="PAAAPPL", owner="PAYROLL", tbl="PABAPPL", tbls=["PAYROLL.PAB%"],
         note="Applicant, requisition and faculty-applicant tables. If empty, recruiting is done "
              "outside Banner."),
    dict(name="Employee Skills Inventory", buys="skills / competency platform",
         forms="PPA%", owner="PAYROLL", tbl="PPRSKIL", tbls=["PAYROLL.PPRSKIL", "PAYROLL.PPRCERT"],
         note="The employee skills matrix (PPRSKIL). Related tracking (certifications, honors) may "
              "be used even when the skills matrix is empty, so read this one narrowly."),
    dict(name="Effort Certification", buys="effort-certification software",
         forms_desc="%EFFORT%", owner="PAYROLL", tbl="PTRECPD",
         note="Matched by form description, not prefix, so the count is effort-cert forms only. If "
              "your institution has grants, effort reporting may be a compliance requirement; verify "
              "the obligation before calling an empty result a gap."),
    #  Banner Workflow does not fit the form/table shape: it is a separate engine with
    #  its own schema, its own users, and its own runtime tables. Measured specially
    #  (see measure_workflow). Its verdict is computed from the real processes defined.
    dict(name="Banner Workflow", buys="Kuali / DocuSign / Formstack (approvals, e-forms)",
         special="workflow", owner="WORKFLOW", tbl="ENG_WORKFLOW",
         lf_stat="processes built", lf_sec="Business processes ever defined in Banner Workflow",
         note="Banner Workflow is a separate engine that automates approvals and business "
              "processes. Every fresh install ships with one sample process named 'system "
              "verification', the built-in installation test. If that is the only process ever "
              "defined, the engine was stood up and never used for real work. EPAF is a separate "
              "mechanism from this engine."),
    dict(name="Salary Planner (budget planning)", buys="budget-planning software (e.g. Axiom)",
         forms_desc="%SALARY PLANNER%", owner="POSNCTL", tbl="NBREHDR",
         tbls=["POSNCTL.NBREHDR", "POSNCTL.NBREPSA", "POSNCTL.NBREJLD", "POSNCTL.NTRSPEX"],
         note="Banner's native Salary Planner builds salary and position-budget scenarios. If its "
              "extract tables (NBREHDR, NBREPSA) are empty, the native worksheet flow is not used. "
              "Some institutions build a custom table for this instead; if native is empty, check "
              "whether a custom process fills the role before concluding anything."),
    dict(name="Purchasing / e-Procurement", buys="Jaggaer / Unimarket",
         forms="FPA%", owner="FIMSMGR", tbl="FPBREQH", tbls=["FIMSMGR.FP%"],
         note="Requisitions and purchase orders. High volumes here mean e-procurement runs in Banner."),
    dict(name="EPAF Personnel e-Forms", buys="DocuSign / Kuali workflow",
         forms="NOA%", owner="POSNCTL", tbl="NOBTRAN", tbls=["POSNCTL.NOB%"],
         note="Electronic Personnel Action Forms. Activity here means e-forms for personnel actions "
              "already run in Banner."),
    dict(name="Web Time Entry", buys="timekeeping product",
         forms="PHA%", owner="PAYROLL", tbl="PHRHOUR",
         tbls=["PAYROLL.PHRHOUR", "PAYROLL.PHREARN", "PAYROLL.PHRELPR"],
         note="Employee time entry. High volumes mean timekeeping runs in Banner."),
    dict(name="Benefits / Open Enrollment", buys="benefits-administration platform",
         forms="PDA%", owner="PAYROLL", tbl="PDRBCOV", tbls=["PAYROLL.PDR%"],
         note="Benefit coverage and enrollment."),
    dict(name="Grants / Research Accounting", buys="grants-management system",
         forms="FRA%", owner="FIMSMGR", tbl="FRBGRNT",
         tbls=["FIMSMGR.FRBGRNT", "FIMSMGR.FRRGRNT", "FIMSMGR.FRRBUDG"],
         note="The grant ledger for sponsored programs."),
    dict(name="Communication Management (BCM)", buys="Slate / Constant Contact / mass email",
         forms="GCA%", owner="GENERAL", tbl="GCRLENT", tbls=["GENERAL.GC%"],
         note="Banner's built-in mass-communication engine. Large volumes in the GC* tables mean a "
              "mass-comms engine already runs in Banner."),
    dict(name="Population Selection / Letter Gen", buys="mail-merge / CRM segmentation",
         forms="GLA%", owner="GENERAL", tbl="GLBEXTR",
         tbls=["GENERAL.GLBEXTR", "GENERAL.GLRSLCT"],
         note="Audience extracts and letter generation."),
    dict(name="Advancement / Alumni / Fundraising", buys="Raiser's Edge / fundraising CRM",
         forms="APA%", owner="ALUMMGR", tbl="APBCONS",
         note="Alumni and fundraising. If the ALUMMGR schema / APBCONS table is absent, the module "
              "is not installed here, which may be a licensing choice rather than shelfware."),
    dict(name="Housing / Residence Life", buys="housing-management system",
         forms="SLR%", owner="SATURN", tbl="SLRRASG", context="expected-absent",
         tbls=["SATURN.SLR%"],
         note="Room assignments. If your institution has no residence halls, an empty result is "
              "expected here, not a finding. Included so its absence is never mistaken for a gap."),
]
