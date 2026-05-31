import streamlit as st
import io
import os
import glob

try:
    import docx
except ImportError:
    st.error("Please ensure the `python-docx` package is installed. Run: pip install python-docx")

st.set_page_config(
    page_title="Huliot India - Workstation Deployer",
    page_icon="💼",
    layout="wide"
)

def replace_text_in_paragraph(paragraph, old_text, new_text):
    """
    Safely replaces text within a paragraph run-by-run.
    This preserves exact font formatting (bold, color, size, underlines).
    """
    if old_text in paragraph.text:
        # First try to find a single run that contains the target phrase
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                return True
        
        # Fallback split-run search: Join adjacent runs if needed
        full_text = ""
        for run in paragraph.runs:
            full_text += run.text
        
        if old_text in full_text:
            # Safely replace across the entire paragraph text block if isolated runs failed
            paragraph.text = paragraph.text.replace(old_text, new_text)
            return True
    return False

def toggle_checkbox_in_paragraph(paragraph, target_phrase, check=True):
    """
    Finds a target phrase inside a paragraph and toggles standard Word unicode
    checkbox characters from ☐ (unchecked) to ☑ (checked) or vice-versa.
    """
    if target_phrase.lower() in paragraph.text.lower():
        for run in paragraph.runs:
            if "☐" in run.text or "☑" in run.text:
                if check:
                    run.text = run.text.replace("☐", "☑")
                else:
                    run.text = run.text.replace("☑", "☐")

def process_document(template_bytes, data):
    """
    Loads the original Word document template binary, parses all paragraph runs 
    and table cell boundaries, and outputs the filled document stream.
    """
    doc = docx.Document(io.BytesIO(template_bytes))
    
    def process_p(paragraph):
        # 1. Base Text Field Variable Replacements
        if "Project No. / Bitrix deal ID:" in paragraph.text:
            # Support both unicode en-space placeholders and traditional underlines
            replace_text_in_paragraph(paragraph, "     ", data['project_no'])
            replace_text_in_paragraph(paragraph, "_____________", data['project_no'])
            
        if "Date:" in paragraph.text:
            # Safe replacement of placeholder blanks following 'Date:'
            replace_text_in_paragraph(paragraph, "Date:", f"Date: {data['date']}")
            
        if "Huliot Sales Person & No:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Huliot Sales Person & No:", f"Huliot Sales Person & No: {data['sales_person']} {data['sales_phone']}")
            
        if "Project Name & Address:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Project Name & Address:", f"Project Name & Address: {data['project_name']} {data['project_address']}")
            
        if "City:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "City:", f"City: {data['city']}")
            
        if "Costumer / Client Name & No:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Costumer / Client Name & No:", f"Costumer / Client Name & No: {data['client_name']} {data['client_phone']}")
            
        if "Architect Name & No:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Architect Name & No:", f"Architect Name & No: {data['architect_name']} {data['architect_phone']}")
            
        if "MEP Consultant Name & No:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "MEP Consultant Name & No:", f"MEP Consultant Name & No: {data['mep_name']} {data['mep_phone']}")

        if "Estimated date to close the offer" in paragraph.text:
            replace_text_in_paragraph(paragraph, "………………………….", data['offer_date'])
        
        if "Probability ---" in paragraph.text:
            # Replaces the default template value of 50% with user value
            replace_text_in_paragraph(paragraph, "50%", f"{data['probability']}%")

        # 2. Dynamic Explanations / Summary Mapping Updates (Page 3 Bottom)
        if "Toilet is suspended: -" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Toilet is suspended: -", f"Toilet is suspended: - {data['toilet_suspended']}")
        if "Kitchen / Utility is sunken (How much in mm): -" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Kitchen / Utility is sunken (How much in mm): -", f"Kitchen / Utility is sunken (How much in mm): - {data['kitchen_sunken']}")
        if "Kitchen is suspended: -" in paragraph.text:
            replace_text_in_paragraph(paragraph, "Kitchen is suspended: -", f"Kitchen is suspended: - {data['kitchen_suspended']}")

        # 3. Drawing Checkbox Toggles
        toggle_checkbox_in_paragraph(paragraph, "All or Typical Floor drawing", data['drw_typical'])
        toggle_checkbox_in_paragraph(paragraph, "Typical Toilet and kitchen section", data['drw_toilet_kitchen'])
        toggle_checkbox_in_paragraph(paragraph, "Podium/ Diversion/ Basement drawing", data['drw_podium'])
        toggle_checkbox_in_paragraph(paragraph, "Full Building section view", data['drw_building'])
        toggle_checkbox_in_paragraph(paragraph, "Only Architectural layout", data['drw_arch_only'])
        toggle_checkbox_in_paragraph(paragraph, "Any Other drawing attached", data['drw_other'])
        if data['drw_other'] and data['drw_other_text']:
            replace_text_in_paragraph(paragraph, "Any Other drawing attached if any", f"Any Other drawing attached if any: {data['drw_other_text']}")

        # 4. Drainage Spec Toggles
        toggle_checkbox_in_paragraph(paragraph, "Internal Toilets -Ultra Silent", data['drn_int_us'])
        toggle_checkbox_in_paragraph(paragraph, "Internal Toilets -HT Pro", data['drn_int_ht'])
        toggle_checkbox_in_paragraph(paragraph, "Vertical/ External stack -Ultra Silent", data['drn_vert_us'])
        toggle_checkbox_in_paragraph(paragraph, "Vertical/ External stack -HT Pro", data['drn_vert_ht'])
        toggle_checkbox_in_paragraph(paragraph, "Podium/Diversion/ Basement – Ultra Silent", data['drn_pod_us'])
        toggle_checkbox_in_paragraph(paragraph, "Podium/Diversion/ Basement - HT Pro", data['drn_pod_ht'])

        # 5. Water Supply Spec Toggles
        toggle_checkbox_in_paragraph(paragraph, "Internal Pipe  (PERT/AL/PERT)", data['ws_int_pipe_pert'])
        toggle_checkbox_in_paragraph(paragraph, "Internal Pipe  (PPR)", data['ws_int_pipe_ppr'])
        toggle_checkbox_in_paragraph(paragraph, "Internal Fittings   PPSU", data['ws_int_fit_ppsu'])
        toggle_checkbox_in_paragraph(paragraph, "Internal Fittings   BRASS", data['ws_int_fit_brass'])
        toggle_checkbox_in_paragraph(paragraph, "External Pipe  (PERT/AL/PERT)", data['ws_ext_pipe_pert'])
        toggle_checkbox_in_paragraph(paragraph, "External Pipe  (PPR)", data['ws_ext_pipe_ppr'])
        toggle_checkbox_in_paragraph(paragraph, "External Fittings   PPSU", data['ws_ext_fit_ppsu'])
        toggle_checkbox_in_paragraph(paragraph, "External Fittings   BRASS", data['ws_ext_fit_brass'])
        toggle_checkbox_in_paragraph(paragraph, "Terrace looping Pipe  (PERT/AL/PERT)", data['ws_ter_pipe_pert'])
        toggle_checkbox_in_paragraph(paragraph, "Terrace looping Pipe  (PPR)", data['ws_ter_pipe_ppr'])
        toggle_checkbox_in_paragraph(paragraph, "Terrace looping Fittings   PPSU", data['ws_ter_fit_ppsu'])
        toggle_checkbox_in_paragraph(paragraph, "Terrace looping Fittings   BRASS", data['ws_ter_fit_brass'])

        # 6. Priority Toggles
        toggle_checkbox_in_paragraph(paragraph, "A: High", data['priority'] == 'A')
        toggle_checkbox_in_paragraph(paragraph, "B: Medium", data['priority'] == 'B')
        toggle_checkbox_in_paragraph(paragraph, "C: Low", data['priority'] == 'C')

        # 7. Building Configuration Toggles
        for b_type in ["Office / commercial", "Shopping centre", "Hotel / Restaurant", "Industrial", 
                       "Hospital / nursing home", "School", "Social housing", "Sport facility", "Infrastructure", "Residential"]:
            if data['building_type'] == b_type:
                toggle_checkbox_in_paragraph(paragraph, b_type, True)
            else:
                toggle_checkbox_in_paragraph(paragraph, b_type, False)

        # 8. Structural Values
        if "No of Typical floors:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "     ", data['floors'])
        if "No of shafts:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "     ", data['shafts'])
        if "No of rooms:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "     ", data['rooms'])
        if "Room/floor to floor height in meter:" in paragraph.text:
            replace_text_in_paragraph(paragraph, "     ", data['floor_height'])

        # 9. Calculation Routing & Systems
        toggle_checkbox_in_paragraph(paragraph, "Single Stack System – HULIOT design", data['calc_drainage'] == "single-stack")
        toggle_checkbox_in_paragraph(paragraph, "Two Stack System (soil / waste", data['calc_drainage'] == "two-stack")
        toggle_checkbox_in_paragraph(paragraph, "Drainage- Same as per Client/ Consultant", data['calc_drainage'] == "consultant-drawing")

        for c_type in ["PVC", "Cast Iron", "HDPE", "CPVC/UPVC/OTHER"]:
            toggle_checkbox_in_paragraph(paragraph, c_type, c_type in data['boq_compare'])

        # 10. Pipeline Routing Specifications
        toggle_checkbox_in_paragraph(paragraph, "Within false ceiling/ ceiling suspended", data['rte_ceiling'])
        if "Toilet Sunken-" in paragraph.text:
            val = f"{data['rte_toilet_mm']} mm" if data['rte_toilet_sunken'] else "No"
            replace_text_in_paragraph(paragraph, "     ", val)
        if "Kitchen Sunken-" in paragraph.text:
            val = f"{data['rte_kitchen_mm']} mm" if data['rte_kitchen_sunken'] else "No"
            replace_text_in_paragraph(paragraph, "     ", val)
        if "Utility Sunken-" in paragraph.text:
            val = f"{data['rte_utility_mm']} mm" if data['rte_utility_sunken'] else "No"
            replace_text_in_paragraph(paragraph, "     ", val)

        toggle_checkbox_in_paragraph(paragraph, "Within false ceiling & Pipe drop ceiling to wall chasing", data['ws_ceiling_drop'])
        toggle_checkbox_in_paragraph(paragraph, "In wall chasing piping full - Internal toilets", data['ws_wall_case'])

        # 11. Service Floor Elements
        toggle_checkbox_in_paragraph(paragraph, "Service floor available:   Yes", data['service_floor'] == "Yes")
        toggle_checkbox_in_paragraph(paragraph, "No    if yes where:", data['service_floor'] == "No")
        if data['service_floor'] == "Yes" and data['service_floor_where']:
            replace_text_in_paragraph(paragraph, "     ", data['service_floor_where'])

        toggle_checkbox_in_paragraph(paragraph, "one shaft", data['calc_for'] == "one-shaft")
        toggle_checkbox_in_paragraph(paragraph, "all shafts", data['calc_for'] == "all-shafts")
        toggle_checkbox_in_paragraph(paragraph, "including basement", data['calc_for'] == "with-basement")

    # Scan standard body paragraphs
    for p in doc.paragraphs:
        process_p(p)
    
    # Scan cells inside nested template tables
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    process_p(p)
                # Recurse inside nested tables if present
                for nested in c.tables:
                    for nr in nested.rows:
                        for nc in nr.cells:
                            for np in nc.paragraphs:
                                process_p(np)

    out_stream = io.BytesIO()
    doc.save(out_stream)
    return out_stream.getvalue()

st.markdown("""
    <div style='background-color:#0f5132; padding:20px; border-radius:12px; margin-bottom:20px;'>
        <h1 style='color:white; margin:0; font-family:sans-serif;'>💼 Huliot India</h1>
        <p style='color:#d1e7dd; margin:5px 0 0 0; font-size:14px; font-weight:bold;'>Checklist Workstation - High-Fidelity Word Generator</p>
    </div>
""", unsafe_allow_html=True)

# Auto-detect any Checklist docx file in the repository
detected_templates = glob.glob("Checklist*.docx")
template_data = None

if detected_templates:
    detected_file = detected_templates[0]
    with open(detected_file, "rb") as f:
        template_data = f.read()
    st.success(f"✔️ Automatically loaded original template layout: `{detected_file}` from repository.")
else:
    st.warning("⚠️ Original template file was not found in the root directory. Please upload it below to process the checklist.")

uploaded_template = st.file_uploader(
    "Upload original file: 'Checklist -Questinorie for Huliot (Drainage and Water supply)_3.docx'",
    type=["docx"]
)

if uploaded_template is not None:
    template_data = uploaded_template.read()
    st.success("✔️ Custom checklist template uploaded successfully!")

if template_data is not None:
    tab1, tab2, tab3 = st.tabs(["📝 General & Drawings", "💧 Piping & Design", "🏢 Building & Routing"])
    
    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("General Information")
            date = st.date_input("Date", value=None)
            project_no = st.text_input("Project No. / Bitrix deal ID", "HPL-2026-004")
            sales_person = st.text_input("Huliot Sales Representative", "Mehta Patel")
            sales_phone = st.text_input("Representative Number", "+91 98765 43210")
            project_name = st.text_input("Project Name", "Greenway Heights Tower A")
            project_address = st.text_input("Project Address", "Plot 42, GIDC Road")
            city = st.text_input("City", "Vadodara")
            priority = st.selectbox("Priority Class", ["A", "B", "C"], format_func=lambda x: f"{x} - High/Medium/Low")
        
        with col2:
            st.subheader("Stakeholder Contacts")
            client_name = st.text_input("Client Name", "Client Company Ltd")
            client_phone = st.text_input("Client Number", "+91 99999 88888")
            architect_name = st.text_input("Architect Name", "Architect Consultant")
            architect_phone = st.text_input("Architect Number", "+91 99999 77777")
            mep_name = st.text_input("MEP Plumbing Consultant", "Plumbing Design Labs")
            mep_phone = st.text_input("MEP Contact Number", "+91 99999 66666")

            st.subheader("Attached Drawings Checklist")
            drw_typical = st.checkbox("All or Typical Floor drawing attached", True)
            drw_toilet_kitchen = st.checkbox("Typical Toilet and kitchen section details", False)
            drw_podium = st.checkbox("Podium/ Diversion/ Basement drawing", False)
            drw_building = st.checkbox("Full Building section view", False)
            drw_arch_only = st.checkbox("Only Architectural layout", False)
            drw_other = st.checkbox("Any Other drawing attached", False)
            drw_other_text = st.text_input("Describe other drawing details", "")

    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Drainage Specifications")
            drn_int_us = st.checkbox("Internal Toilets - Ultra Silent", True)
            drn_int_ht = st.checkbox("Internal Toilets - HT Pro", False)
            drn_vert_us = st.checkbox("Vertical/ External stack - Ultra Silent", True)
            drn_vert_ht = st.checkbox("Vertical/ External stack - HT Pro", False)
            drn_pod_us = st.checkbox("Podium/Diversion/ Basement – Ultra Silent", False)
            drn_pod_ht = st.checkbox("Podium/Diversion/ Basement - HT Pro", False)
            
        with col2:
            st.subheader("Water Supply Pipes & Fittings")
            ws_int_pipe_pert = st.checkbox("Internal Pipe - PERT/AL/PERT", True)
            ws_int_pipe_ppr = st.checkbox("Internal Pipe - PPR", False)
            ws_int_fit_ppsu = st.checkbox("Internal Fittings - PPSU", True)
            ws_int_fit_brass = st.checkbox("Internal Fittings - Brass", False)
            
            ws_ext_pipe_pert = st.checkbox("External Pipe - PERT/AL/PERT", True)
            ws_ext_pipe_ppr = st.checkbox("External Pipe - PPR", False)
            ws_ext_fit_ppsu = st.checkbox("External Fittings - PPSU", True)
            ws_ext_fit_brass = st.checkbox("External Fittings - Brass", False)
            
            ws_ter_pipe_pert = st.checkbox("Terrace Pipe - PERT/AL/PERT", False)
            ws_ter_pipe_ppr = st.checkbox("Terrace Pipe - PPR", False)
            ws_ter_fit_ppsu = st.checkbox("Terrace Fittings - PPSU", False)
            ws_ter_fit_brass = st.checkbox("Terrace Fittings - Brass", False)

    with tab3:
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Building Specifications")
            building_type = st.selectbox("Kind of Building", [
                "Office / commercial", "Shopping centre", "Hotel / Restaurant", "Industrial", 
                "Hospital / nursing home", "School", "Social housing", "Sport facility", "Infrastructure", "Residential"
            ], index=0)
            floors = st.text_input("No of Typical floors", "12")
            shafts = st.text_input("No of shafts", "4")
            rooms = st.text_input("No of rooms", "84")
            floor_height = st.text_input("Room/floor to floor height in meter", "3.2")
            offer_date = st.text_input("Estimated date to close offer", "30.06.2026")
            probability = st.slider("Probability (%)", 0, 100, 50)

        with col2:
            st.subheader("Calculation Configurations")
            calc_drainage = st.selectbox("Drainage system requirements", [
                ("single-stack", "Single Stack System – HULIOT design drawing, BOQ & Quotation"),
                ("two-stack", "Two Stack System (soil / waste / vent) – HULIOT design"),
                ("consultant-drawing", "Same as per Client/ Consultant drawing - BOQ Quantity only")
            ], format_func=lambda x: x[1])
            
            boq_compare = st.multiselect("BOQ is compared with", ["PVC", "Cast Iron", "HDPE", "CPVC/UPVC/OTHER"], default=["PVC"])
            calc_for = st.selectbox("Calculation design scale", [
                ("one-shaft", "one shaft"),
                ("all-shafts", "all shafts"),
                ("with-basement", "including basement")
            ], format_func=lambda x: x[1])

            st.subheader("Pipe Routing Options")
            rte_ceiling = st.checkbox("Within false ceiling / suspended / underslung system", True)
            rte_toilet_sunken = st.checkbox("Toilet Sunken")
            rte_toilet_mm = st.text_input("Toilet Sunken depth in mm", "250") if rte_toilet_sunken else "0"
            rte_kitchen_sunken = st.checkbox("Kitchen Sunken")
            rte_kitchen_mm = st.text_input("Kitchen Sunken depth in mm", "250") if rte_kitchen_sunken else "0"
            rte_utility_sunken = st.checkbox("Utility Sunken")
            rte_utility_mm = st.text_input("Utility Sunken depth in mm", "250") if rte_utility_sunken else "0"

            ws_ceiling_drop = st.checkbox("Water supply - within false ceiling & pipe drop ceiling")
            ws_wall_case = st.checkbox("Water supply - in wall chasing piping full")
            
            service_floor = st.radio("Service floor available", ["Yes", "No"], index=1)
            service_floor_where = st.text_input("If Yes, specify location", "") if service_floor == "Yes" else ""

    st.write("---")
    
    # Payload collection dictionary mapping directly to Word template variables
    data = {
        'date': str(date) if date else "",
        'project_no': project_no,
        'sales_person': sales_person,
        'sales_phone': sales_phone,
        'project_name': project_name,
        'project_address': project_address,
        'city': city,
        'client_name': client_name,
        'client_phone': client_phone,
        'architect_name': architect_name,
        'architect_phone': architect_phone,
        'mep_name': mep_name,
        'mep_phone': mep_phone,
        'drw_typical': drw_typical,
        'drw_toilet_kitchen': drw_toilet_kitchen,
        'drw_podium': drw_podium,
        'drw_building': drw_building,
        'drw_arch_only': drw_arch_only,
        'drw_other': drw_other,
        'drw_other_text': drw_other_text,
        'drn_int_us': drn_int_us,
        'drn_int_ht': drn_int_ht,
        'drn_vert_us': drn_vert_us,
        'drn_vert_ht': drn_vert_ht,
        'drn_pod_us': drn_pod_us,
        'drn_pod_ht': drn_pod_ht,
        'ws_int_pipe_pert': ws_int_pipe_pert,
        'ws_int_pipe_ppr': ws_int_pipe_ppr,
        'ws_int_fit_ppsu': ws_int_fit_ppsu,
        'ws_int_fit_brass': ws_int_fit_brass,
        'ws_ext_pipe_pert': ws_ext_pipe_pert,
        'ws_ext_pipe_ppr': ws_ext_pipe_ppr,
        'ws_ext_fit_ppsu': ws_ext_fit_ppsu,
        'ws_ext_fit_brass': ws_ext_fit_brass,
        'ws_ter_pipe_pert': ws_ter_pipe_pert,
        'ws_ter_pipe_ppr': ws_ter_pipe_ppr,
        'ws_ter_fit_ppsu': ws_ter_fit_ppsu,
        'ws_ter_fit_brass': ws_ter_fit_brass,
        'priority': priority,
        'building_type': building_type,
        'floors': floors,
        'shafts': shafts,
        'rooms': rooms,
        'floor_height': floor_height,
        'offer_date': offer_date,
        'probability': probability,
        'calc_drainage': calc_drainage[0],
        'boq_compare': boq_compare,
        'calc_for': calc_for[0],
        'rte_ceiling': rte_ceiling,
        'rte_toilet_sunken': rte_toilet_sunken,
        'rte_toilet_mm': rte_toilet_mm,
        'rte_kitchen_sunken': rte_kitchen_sunken,
        'rte_kitchen_mm': rte_kitchen_mm,
        'rte_utility_sunken': rte_utility_sunken,
        'rte_utility_mm': rte_utility_mm,
        'ws_ceiling_drop': ws_ceiling_drop,
        'ws_wall_case': ws_wall_case,
        'service_floor': service_floor,
        'service_floor_where': service_floor_where,
        'toilet_suspended': "Yes" if rte_ceiling else "No",
        'kitchen_sunken': f"Toilet: {rte_toilet_mm}mm, Kitchen: {rte_kitchen_mm}mm, Utility: {rte_utility_mm}mm" if (rte_toilet_sunken or rte_kitchen_sunken or rte_utility_sunken) else "No",
        'kitchen_suspended': "Yes" if rte_ceiling else "No"
    }

    if st.button("🚀 Build Completed Word Document", type="primary"):
        with st.spinner("Injecting values to original document structure..."):
            try:
                processed_data = process_document(template_data, data)
                st.success("🎉 Word document created successfully with no format alterations!")
                st.download_button(
                    label="⬇️ Download Verbatim Filled Document (.docx)",
                    data=processed_data,
                    file_name=f"Checklist_Huliot_{project_no}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"Failed to process original template. Please check file schema. Error details: {e}")
```
_Note: `requirements.txt` remains unchanged as it already lists `streamlit>=1.35.0` and `python-docx>=1.1.0`._

---

### What to Do Next (Deploying to Streamlit):

Your GitHub repository pictured in `image_7bed20.png` is perfectly organized! To deploy your working application, follow these simple steps:

1. **Log in to Streamlit Community Cloud**:
   * Open your web browser and go to [share.streamlit.io](https://share.streamlit.io/).
   * Click the **"Connect GitHub"** or **"Sign In with GitHub"** button and log in with your GitHub account credentials.

2. **Deploy your Repository**:
   * Click the **"New app"** button in the top right of your Streamlit Dashboard.
   * Under **Repository**, select your repository: `plumco/huliot_eze_checklist`.
   * Under **Branch**, select `main`.
   * Under **Main file path**, type: `streamlit_app.py`.

3. **Launch the Workstation**:
   * Click **"Deploy!"** at the bottom of the page.
   * Streamlit will spend 1–2 minutes setting up your server, installing `python-docx` and `streamlit` as specified in `requirements.txt`, and launching your app.

Once deployed, the app will automatically load your uploaded Word template file `Checklist -Questinorie for Huliot (Drainage and Water supply)_3.docx` directly from your GitHub directory and allow you to generate perfectly formatted Word documents!