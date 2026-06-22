import pandas as pd

TICK_VALUES = {"x", "yes", "y", "1", True, "true"}

UNIVERSAL_FORMS = {
    "Adverse Event", "Concomitant Medication",
    "AESI/SAE", "SAE Evaluation",
}

def build_bookmark_plan(df_matrix: pd.DataFrame, df_forms: pd.DataFrame) -> list:
    """
    Takes the matrix DataFrame and the forms DataFrame to build the 3-level dictionary structure.
    """
    forms_data = {}
    for _, row in df_forms.iterrows():
        # Skip forms that are marked as "Not Submitted"
        if row.get("Status") == "Not Submitted":
            continue
        forms_data[row["Form Name"]] = int(row["Start Page"])

    visits = [col for col in df_matrix.columns if col != "Form Name"]
    
    visit_forms = {v: [] for v in visits}
    form_visits = {f: [] for f in df_matrix["Form Name"]}
    
    for _, row in df_matrix.iterrows():
        form_name = row["Form Name"]
        if form_name not in forms_data:
            continue
            
        for visit in visits:
            val = row[visit]
            if val and str(val).strip().lower() in TICK_VALUES:
                visit_forms[visit].append(form_name)
                form_visits[form_name].append(visit)

    visit_list = []
    for visit in visits:
        form_names = visit_forms.get(visit, [])
        if not form_names:
            continue
            
        specific = [f for f in form_names if f not in UNIVERSAL_FORMS and forms_data.get(f, 0) > 0]
        pool     = specific if specific else [f for f in form_names if forms_data.get(f, 0) > 0]
        rep_page = min(forms_data[f] for f in pool) if pool else 1
        
        children = sorted(
            [{"name": f, "page": forms_data[f]} for f in form_names if forms_data.get(f, 0) > 0],
            key=lambda x: x["page"]
        )
        visit_list.append({"name": visit, "page": rep_page, "children": children})

    form_list = []
    for form_name in sorted(forms_data.keys(), key=str.lower):
        page = forms_data[form_name]
        if page <= 0:
            continue
        visit_children = [
            {"name": v, "page": rep_page}
            for v in visits
            if v in form_visits.get(form_name, [])
            for rep_page in [page]
        ]
        form_list.append({"name": form_name, "page": page, "children": visit_children})

    bookmark_plan = [
        {"group": "Visit", "items": visit_list},
        {"group": "Form",  "items": form_list},
    ]
    
    return bookmark_plan
