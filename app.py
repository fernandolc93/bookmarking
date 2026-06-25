import streamlit as st
import pandas as pd
import os
import uuid
from core.extract import extract_forms_from_pdf
from core.plan import build_bookmark_plan
from core.pdf_writer import write_bookmarks_to_pdf
from core.matrix_io import parse_imported_matrix, matrix_to_excel_bytes
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, DataReturnMode, JsCode

# ── Session bootstrap ─────────────────────────────────────────────────────────
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())

st.set_page_config(page_title="aCRF Bookmarker", layout="wide", page_icon="🔖")

# ── Styling ───────────────────────────────────────────────────────────────────

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.stApp {
    background-color: #233142;
    color: #e3e3e3;
}

h1, h2, h3 { color: #e3e3e3; font-weight: 800; }

/* Accent line under the title */
h1 { border-bottom: 3px solid #f95959; padding-bottom: 0.4rem; }

/* Regular buttons — steel blue with red hover */
div.stButton > button:first-child {
    background: #455d7a;
    color: #e3e3e3; border: none; border-radius: 8px;
    padding: 0.5rem 1rem; font-weight: 600;
    transition: transform 0.2s, box-shadow 0.2s, background 0.2s;
}
div.stButton > button:first-child:hover {
    background: #f95959;
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(249, 89, 89, 0.4);
    color: white;
}

/* Download button — red accent */
div.stDownloadButton > button:first-child {
    background: #f95959;
    color: white; border: none; border-radius: 8px;
    padding: 0.5rem 1rem; font-weight: 600;
    transition: transform 0.2s, box-shadow 0.2s;
}
div.stDownloadButton > button:first-child:hover {
    transform: translateY(-2px);
    box-shadow: 0 4px 12px rgba(249, 89, 89, 0.5);
    color: white;
}

/* File uploader border */
[data-testid="stFileUploader"] {
    border: 2px dashed #f95959;
    border-radius: 8px;
    padding: 1rem;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background-color: #455d7a;
}
</style>
""", unsafe_allow_html=True)

# ── Custom single-click checkbox renderer ─────────────────────────────────────
# Pure client-side checkbox: toggles on click but does NOT call
# params.setValue(), so it will NOT trigger a Streamlit rerun.
# The checked state is stored on the DOM element; the "Save Matrix"
# button reads it back via grid_response["data"].
CHECKBOX_RENDERER = JsCode("""
class CheckboxRenderer {
    init(params) {
        this.params = params;

        this.eGui = document.createElement('div');
        this.eGui.style.cssText =
            'display:flex; align-items:center; justify-content:center; height:100%;';

        this.cb = document.createElement('input');
        this.cb.type = 'checkbox';
        this.cb.checked = params.value === true || params.value === 'true';
        this.cb.style.cssText = 'width:15px; height:15px; cursor:pointer; accent-color:#3b82f6;';

        // Update the grid's row data directly — no Streamlit round-trip.
        this.cb.addEventListener('change', () => {
            this.params.node.setDataValue(this.params.colDef.field, this.cb.checked);
        });

        this.eGui.appendChild(this.cb);
    }
    getGui()  { return this.eGui; }
    refresh(params) {
        this.params = params;
        this.cb.checked = params.value === true || params.value === 'true';
        return true;
    }
}
""")

st.title("aCRF Automatic Bookmarking")
st.markdown(
    "Transform your print output into a fully bookmarked PDF. "
    "Upload your document to extract forms, define the schedule of assessments, "
    "map forms to visits, and download the finished product."
)

# ── Step 1: Upload ────────────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Step 1: Upload aCRF")
uploaded_file = st.file_uploader("Upload aCRF (PDF)", type=["pdf"])

if uploaded_file:
    temp_input  = f"temp_input_{st.session_state.session_id}.pdf"
    temp_output = f"temp_output_{st.session_state.session_id}.pdf"

    # Flush all state when a different file is uploaded
    if st.session_state.get("_last_upload") != uploaded_file.name:
        for key in ("edited_forms", "df_visits", "matrix", "_last_upload"):
            st.session_state.pop(key, None)
        st.session_state["_last_upload"] = uploaded_file.name
        st.cache_data.clear()

    with open(temp_input, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # ── Step 2: Extracted forms ───────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 2: Extracted Forms")

    @st.cache_data
    def load_forms(path):
        return extract_forms_from_pdf(path)

    try:
        with st.spinner("Extracting titles..."):
            df_forms_initial = load_forms(temp_input)

        if "edited_forms" not in st.session_state:
            df_forms_initial["Status"] = "Submitted"
            st.session_state.edited_forms = df_forms_initial

        st.markdown("Review and edit the extracted forms, start/end pages, and status:")
        
        # Calculate height to prevent internal scrollbar: 35px per row + 38px for header/padding
        table_height = len(st.session_state.edited_forms) * 35 + 38
        
        edited_forms = st.data_editor(
            st.session_state.edited_forms,
            num_rows="dynamic",
            use_container_width=True,
            height=table_height,
            key="forms_editor",
            column_config={
                "Status": st.column_config.SelectboxColumn(
                    "Status",
                    help="Not Submitted forms are excluded from the matrix.",
                    width="medium",
                    options=["Submitted", "Not Submitted"],
                    required=True,
                )
            },
        )
        if not edited_forms.equals(st.session_state.edited_forms):
            st.session_state.edited_forms = edited_forms
            st.session_state.needs_rerun = True

    except Exception as e:
        st.error(f"Error extracting forms: {e}")
        st.stop()

    # ── Step 3: Schedule of Assessments ──────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 3: Schedule of Assessments")

    if "df_visits" not in st.session_state:
        st.session_state.df_visits = pd.DataFrame(
            {"Visit Name": ["Screening", "Cycle 1 Day 1"]}
        )

    st.markdown("Define the list of visits for your study:")
    edited_visits = st.data_editor(
        st.session_state.df_visits,
        num_rows="dynamic",
        use_container_width=True,
        key="visits_editor",
    )
    if not edited_visits.equals(st.session_state.df_visits):
        st.session_state.df_visits = edited_visits
        st.session_state.needs_rerun = True
    edited_visits = st.session_state.df_visits

    # ── Step 4: Matrix mapping ────────────────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 4: Map Forms to Visits")

    # Auto-generate or auto-update the matrix based on current forms and visits
    visit_cols = [
        str(v).strip() for v in edited_visits["Visit Name"].tolist()
        if pd.notna(v) and str(v).strip() and str(v).strip() != "None"
    ]
    active_forms = edited_forms[edited_forms["Status"] == "Submitted"]["Form Name"].tolist()

    if "matrix" not in st.session_state:
        df_matrix = pd.DataFrame({"Form Name": active_forms})
        for col in visit_cols:
            df_matrix[col] = False
        st.session_state.matrix = df_matrix
    else:
        old_matrix = st.session_state.matrix
        # Strip any AG Grid internal columns (e.g. ::auto_unique_id::) immediately
        _internal = [c for c in old_matrix.columns if isinstance(c, str) and c.startswith("::")]
        if _internal:
            old_matrix = old_matrix.drop(columns=_internal)
            
        new_matrix = pd.DataFrame({"Form Name": active_forms})
        for col in visit_cols:
            if col in old_matrix.columns:
                mapping = old_matrix.set_index("Form Name")[col].to_dict()
                new_matrix[col] = new_matrix["Form Name"].map(mapping).fillna(False).astype(bool)
            else:
                new_matrix[col] = False
        st.session_state.matrix = new_matrix

    matrix     = st.session_state.matrix
    form_names = matrix["Form Name"].tolist()

    # ---- Import / Export ------------------------------------------------
    with st.expander("📥 Import / 📤 Export the matrix as Excel", expanded=False):
        st.caption(
            "For a big grid it's often faster to fill in Excel. Export, "
            "tick cells offline, re-import. Matching is by name; anything "
            "that doesn't match your Step 1-3 lists is skipped and reported."
        )
        exp_col, imp_col = st.columns(2)
        with exp_col:
            st.download_button(
                "📤 Export current matrix",
                data=matrix_to_excel_bytes(matrix),
                file_name="bookmark_matrix.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with imp_col:
            uploaded_matrix = st.file_uploader(
                "Import a filled-in matrix", type=["xlsx"], key="matrix_upload"
            )
        if uploaded_matrix is not None:
            if st.button("Apply imported matrix", use_container_width=True):
                try:
                    imported, warnings = parse_imported_matrix(
                        uploaded_matrix, form_names, visit_cols
                    )
                    for form_name, visit_name, ticked in imported:
                        matrix.loc[matrix["Form Name"] == form_name, visit_name] = ticked
                    st.session_state.matrix = matrix
                    if warnings:
                        st.warning(
                            "Imported with some mismatches:\n\n"
                            + "\n".join(f"- {w}" for w in warnings)
                        )
                    else:
                        st.success(f"Imported cleanly — {len(imported)} cell(s) applied.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Could not read that file: {e}")

    # ---- Bulk tick / clear ----------------------------------------------
    st.markdown(
        "**Bulk actions** — tick or clear many cells at once (handy for forms "
        "like that apply to almost every visit):"
    )
    bulk_form_col, bulk_visit_col, bulk_tick_col, bulk_clear_col = st.columns([3, 3, 1, 1])
    with bulk_form_col:
        bulk_forms = st.multiselect("Form(s)", form_names, key="bulk_forms")
    with bulk_visit_col:
        bulk_visits = st.multiselect(
            "Visit(s)", visit_cols, default=visit_cols, key="bulk_visits"
        )
    with bulk_tick_col:
        st.write("")
        if st.button("✅ Tick", use_container_width=True):
            for f in bulk_forms:
                matrix.loc[matrix["Form Name"] == f, bulk_visits] = True
            st.session_state.matrix = matrix
            st.rerun()
    with bulk_clear_col:
        st.write("")
        if st.button("⬜ Clear", use_container_width=True):
            for f in bulk_forms:
                matrix.loc[matrix["Form Name"] == f, bulk_visits] = False
            st.session_state.matrix = matrix
            st.rerun()

    select_all_col, clear_all_col, _spacer = st.columns([1, 1, 4])
    with select_all_col:
        if st.button("Select all cells"):
            matrix[visit_cols] = True
            st.session_state.matrix = matrix
            st.rerun()
    with clear_all_col:
        if st.button("Clear all cells"):
            matrix[visit_cols] = False
            st.session_state.matrix = matrix
            st.rerun()

    # ---- Coverage summary -----------------------------------------------
    matrix = st.session_state.matrix
    # Ensure visit columns are bool (AG Grid may return strings)
    bool_matrix = matrix[visit_cols].apply(
        lambda c: c.map(lambda v: v is True or str(v).strip().lower() in ('true', '1'))
    ) if visit_cols else pd.DataFrame()
    n_total  = len(form_names) * len(visit_cols) if visit_cols else 0
    n_ticked = int(bool_matrix.to_numpy().sum()) if visit_cols else 0
    unmapped_forms = [
        f for f in form_names
        if visit_cols and not bool_matrix.loc[matrix["Form Name"] == f].any(axis=1).iloc[0]
    ]
    empty_visits = [v for v in visit_cols if not bool_matrix[v].any()]

    st.caption(f"{n_ticked} / {n_total} cells ticked.")
    if unmapped_forms:
        shown = ", ".join(unmapped_forms[:8]) + (" …" if len(unmapped_forms) > 8 else "")
        st.warning(f"{len(unmapped_forms)} form(s) not assigned to any visit yet: {shown}")
    if empty_visits:
        shown = ", ".join(empty_visits[:8]) + (" …" if len(empty_visits) > 8 else "")
        st.warning(f"{len(empty_visits)} visit(s) have no forms ticked yet: {shown}")

    # ---- AG Grid with native single-click checkboxes --------------------
    st.markdown("Tick the checkboxes to assign forms to visits:")
    st.caption("💡 Click checkboxes freely"
               "When you're done, click **Save Matrix** below the grid.")

    # Pass a clean copy to AG Grid (no internal columns)
    _clean_matrix = st.session_state.matrix[["Form Name"] + visit_cols]
    gb = GridOptionsBuilder.from_dataframe(_clean_matrix)
    gb.configure_column(
        "Form Name",
        pinned="left",
        editable=False,
        width=280,
        suppressMovable=True,
    )
    for col in visit_cols:
        gb.configure_column(
            col,
            editable=False,          # renderer handles editing, not the editor
            width=50,
            cellRenderer=CHECKBOX_RENDERER,
            suppressMenu=True,
            headerTooltip=col,
            resizable=False,
        )
    gb.configure_grid_options(
        rowHeight=30,
        headerHeight=120,
        suppressRowClickSelection=True,
        suppressCellFocus=True,      # no blue focus border between clicks
    )

    grid_response = AgGrid(
        _clean_matrix,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,   # no reruns on click
        data_return_mode=DataReturnMode.AS_INPUT,
        allow_unsafe_jscode=True,    # required for JsCode renderer
        fit_columns_on_grid_load=False,
        height=min(60 + 30 * len(form_names), 600),
        theme="alpine",
        key="matrix_grid",
    )

    # Commit grid state only when user clicks Save
    if st.button("💾 Save Matrix", use_container_width=True):
        final_matrix = grid_response["data"]
        # Drop AG Grid internal columns (e.g. ::auto_unique_id::)
        internal_cols = [c for c in final_matrix.columns if c.startswith("::")]
        if internal_cols:
            final_matrix = final_matrix.drop(columns=internal_cols)
        for col in visit_cols:
            if col in final_matrix.columns:
                final_matrix[col] = final_matrix[col].map(
                    lambda v: v is True or str(v).strip().lower() in ('true', '1')
                )
        st.session_state.matrix = final_matrix
        st.rerun()

    final_matrix = st.session_state.matrix

    # ── Step 5: Build bookmarked PDF ──────────────────────────────────────
    st.markdown("---")
    st.subheader("Step 5: Build Bookmarked PDF")
    if st.button("Build Final PDF"):
        with st.spinner("Writing bookmarks..."):
            try:
                # Filter out "Not Submitted" forms before building the plan
                submitted_forms = edited_forms[edited_forms["Status"] == "Submitted"]
                plan = build_bookmark_plan(final_matrix, submitted_forms)
                write_bookmarks_to_pdf(temp_input, temp_output, plan)
                with open(temp_output, "rb") as pdf_file:
                    st.session_state.pdf_bytes = pdf_file.read()
                st.success("Successfully generated!")
                if os.path.exists(temp_output):
                    os.remove(temp_output)
            except Exception as e:
                st.error(f"Error building PDF: {e}")

    # Show download button outside the Build button block so it persists
    if "pdf_bytes" in st.session_state:
        st.download_button(
            label="📥 Download Bookmarked PDF",
            data=st.session_state.pdf_bytes,
            file_name="acrf_bookmarked_auto.pdf",
            mime="application/pdf",
        )

# Deferred rerun execution
if st.session_state.get("needs_rerun", False):
    st.session_state.needs_rerun = False
    st.rerun()
