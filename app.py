"""
Web interface for the AI Data Analyst Agent, built with Streamlit.

Setup:
    pip install streamlit google-genai pandas openpyxl python-docx python-pptx pypdf
    export GEMINI_API_KEY="your-key-here"

Run:
    streamlit run app.py

Theme colors live in .streamlit/config.toml — make sure that folder is
kept alongside this file (some file-transfer methods hide dotfolders).
"""

import os
import tempfile
import streamlit as st

st.set_page_config(page_title="AI Data Analyst Agent", page_icon="◆", layout="wide")

# ----------------------------------------------------------------------
# Custom styling — a "data terminal" look: dark slate background, amber
# accent, monospace headers. Kept as a single injected stylesheet so the
# rest of the app can just use plain Streamlit components.
# ----------------------------------------------------------------------
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;600;800&family=Inter:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

    h1, h2, h3, .app-title { font-family: 'JetBrains Mono', monospace; }

    .app-header {
        display: flex;
        align-items: baseline;
        gap: 0.6rem;
        border-bottom: 1px solid #2A3140;
        padding-bottom: 0.9rem;
        margin-bottom: 0.4rem;
    }
    .app-title {
        font-size: 1.6rem;
        font-weight: 800;
        color: #F0B429;
        letter-spacing: -0.02em;
    }
    .app-subtitle {
        color: #8A93A6;
        font-size: 0.92rem;
    }

    .stButton>button {
        border: 1px solid #2A3140;
        background-color: #1B2129;
        color: #E8E6E1;
        border-radius: 6px;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.82rem;
        transition: border-color 0.15s ease;
    }
    .stButton>button:hover {
        border-color: #F0B429;
        color: #F0B429;
    }

    section[data-testid="stSidebar"] {
        border-right: 1px solid #2A3140;
    }

    .log-line {
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.85rem;
        color: #C9CDD6;
        padding: 2px 0;
    }
    .log-line::before { content: "› "; color: #F0B429; }
    </style>

    <div class="app-header">
        <span class="app-title">◆ AI DATA ANALYST AGENT</span>
        <span class="app-subtitle">upload &middot; clean &middot; query &middot; report</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# Check the API key BEFORE importing agent.py, since agent.py creates the
# Gemini client at import time and would otherwise crash with a raw
# traceback instead of a friendly message.
if not os.environ.get("GEMINI_API_KEY"):
    st.error(
        "GEMINI_API_KEY is not set. Get a free key at "
        "https://aistudio.google.com/apikey and set it as an environment "
        "variable before running this app."
    )
    st.stop()

import agent  # reuses all the logic already built in agent.py

# ----------------------------------------------------------------------
# Session state setup
# ----------------------------------------------------------------------
for key, default in [
    ("df", None), ("filepath", None), ("kind", None),
    ("text_content", None), ("history", []), ("last_file", None),
    ("cleaning_log", []),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def run_action(label, fn, *args):
    with st.spinner(f"Running {label}..."):
        result = fn(*args)
    st.session_state.history.append((label, result))


# ----------------------------------------------------------------------
# Sidebar: file upload, status, download
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("**FILE**")
    uploaded = st.file_uploader(
        "Upload a file",
        type=["csv", "xlsx", "xls", "pdf", "docx", "pptx"],
        label_visibility="collapsed",
    )

    if uploaded is not None and st.session_state.get("uploaded_name") != uploaded.name:
        tmp_dir = tempfile.mkdtemp()
        tmp_path = os.path.join(tmp_dir, uploaded.name)
        with open(tmp_path, "wb") as f:
            f.write(uploaded.getbuffer())

        st.session_state.uploaded_name = uploaded.name
        st.session_state.filepath = tmp_path
        st.session_state.history = []
        st.session_state.last_file = None

        with st.spinner("Loading and cleaning..."):
            kind, content = agent.load_any_file(tmp_path)
            st.session_state.kind = kind
            if kind == "table":
                df, log = agent.auto_clean(content)
                st.session_state.df = df
                st.session_state.cleaning_log = log
            else:
                st.session_state.text_content = content
        st.rerun()

    if st.session_state.get("uploaded_name"):
        st.markdown("**LOADED**")
        st.markdown(f"`{st.session_state.uploaded_name}`")
        if st.session_state.kind == "table" and st.session_state.df is not None:
            r, c = st.session_state.df.shape
            st.caption(f"{r} rows × {c} columns")

    st.divider()

    if st.button("Reset / clear file", use_container_width=True):
        for key in ["df", "filepath", "kind", "text_content", "history",
                    "last_file", "cleaning_log", "uploaded_name"]:
            st.session_state[key] = [] if key in ("history", "cleaning_log") else None
        st.rerun()

    if st.session_state.get("last_file") and os.path.exists(st.session_state.last_file):
        st.divider()
        st.markdown("**LAST GENERATED FILE**")
        with open(st.session_state.last_file, "rb") as f:
            st.download_button(
                "Download",
                f,
                file_name=os.path.basename(st.session_state.last_file),
                use_container_width=True,
            )

# ----------------------------------------------------------------------
# Main area
# ----------------------------------------------------------------------
if st.session_state.kind == "table" and st.session_state.df is not None:
    df = st.session_state.df
    base = os.path.splitext(st.session_state.filepath)[0]

    tab_overview, tab_actions, tab_chat = st.tabs(["Overview", "Actions & Exports", "Chat"])

    # ---- Overview tab ----
    with tab_overview:
        with st.expander("Cleaning log", expanded=True):
            if st.session_state.cleaning_log:
                for line in st.session_state.cleaning_log:
                    st.markdown(f"<div class='log-line'>{line}</div>", unsafe_allow_html=True)
            else:
                st.caption("No log yet.")

        st.markdown("**DATA PREVIEW**")
        st.dataframe(df.head(20), use_container_width=True)

        if st.button("Re-run auto clean"):
            with st.spinner("Cleaning..."):
                new_df, log = agent.auto_clean(df)
            st.session_state.df = new_df
            st.session_state.cleaning_log = log
            st.rerun()

    # ---- Actions & Exports tab ----
    with tab_actions:
        st.markdown("**REPORTS & ANALYSIS**")
        row1 = st.columns(3)
        if row1[0].button("Report", use_container_width=True):
            run_action("Report", agent.generate_report, df)
        if row1[1].button("Save cleaned file", use_container_width=True):
            run_action("Save", agent.save_data, df, base)
        if row1[2].button("SQL export (.db)", use_container_width=True):
            out = base + ".db"
            run_action("SQL export", agent.export_sql_db, df, out)
            st.session_state.last_file = out

        st.markdown("**EXCEL**")
        row2 = st.columns(3)
        if row2[0].button("Excel dashboard", use_container_width=True):
            out = base + "_dashboard.xlsx"
            run_action("Dashboard", agent.generate_excel_dashboard, df, out)
            st.session_state.last_file = out
        if row2[1].button("Edit Excel in place", use_container_width=True):
            out = base + "_edited.xlsx"
            run_action("Edit Excel", agent.edit_excel_in_place, st.session_state.filepath, df, out)
            st.session_state.last_file = out
        row2[2].caption("Use `formula:`, `format:`, `validate:` in Chat for column-level Excel edits.")

        st.markdown("**POWER BI**")
        row3 = st.columns(3)
        if row3[0].button("Power BI export", use_container_width=True):
            out = base + "_powerbi.xlsx"
            run_action("Power BI export", agent.export_for_powerbi, df, out)
            st.session_state.last_file = out
        if row3[1].button("Power BI guide", use_container_width=True):
            out = base + "_powerbi_guide.md"
            run_action("Power BI guide", agent.generate_powerbi_guide, df, out)
            st.session_state.last_file = out

        if st.session_state.history:
            st.divider()
            st.markdown("**RESULTS**")
            for label, result in reversed(st.session_state.history):
                with st.expander(label, expanded=(label == st.session_state.history[-1][0])):
                    st.text(result)

    # ---- Chat tab ----
    with tab_chat:
        st.caption(
            "Ask a question in plain English, or use a command prefix: "
            "`clean:`, `sql:`, `formula:`, `format:`, `validate:`"
        )

        user_input = st.chat_input("Ask a question or give an instruction...")

        if user_input:
            low = user_input.lower()
            if low.startswith("clean:"):
                with st.spinner("Cleaning..."):
                    new_df, msg = agent.custom_clean(df, user_input[6:].strip())
                st.session_state.df = new_df
                st.session_state.history.append((user_input, msg))
            elif low.startswith("sql:"):
                with st.spinner("Running SQL..."):
                    msg = agent.run_sql(df, user_input[4:].strip())
                st.session_state.history.append((user_input, msg))
            elif low.startswith("formula:"):
                out = base + "_formula.xlsx"
                with st.spinner("Adding formula column..."):
                    msg = agent.add_formula_column(df, user_input[8:].strip(), out)
                st.session_state.history.append((user_input, msg))
                st.session_state.last_file = out
            elif low.startswith("format:"):
                out = base + "_formatted.xlsx"
                with st.spinner("Applying conditional formatting..."):
                    msg = agent.add_conditional_formatting(df, user_input[7:].strip(), out)
                st.session_state.history.append((user_input, msg))
                st.session_state.last_file = out
            elif low.startswith("validate:"):
                out = base + "_validated.xlsx"
                with st.spinner("Adding data validation..."):
                    msg = agent.add_data_validation(df, user_input[9:].strip(), out)
                st.session_state.history.append((user_input, msg))
                st.session_state.last_file = out
            else:
                with st.spinner("Thinking..."):
                    msg = agent.ask_question(df, user_input)
                st.session_state.history.append((user_input, msg))
            st.rerun()

        for question, answer in reversed(st.session_state.history):
            with st.chat_message("user"):
                st.write(question)
            with st.chat_message("assistant"):
                st.text(answer)

elif st.session_state.kind == "text" and st.session_state.text_content is not None:
    st.markdown("**DOCUMENT LOADED**")
    with st.expander("Preview", expanded=False):
        st.text(st.session_state.text_content[:3000])

    user_input = st.chat_input("Ask something about the document...")
    if user_input:
        with st.spinner("Thinking..."):
            msg = agent.ask_about_text(st.session_state.text_content, user_input)
        st.session_state.history.append((user_input, msg))
        st.rerun()

    for question, answer in reversed(st.session_state.history):
        with st.chat_message("user"):
            st.write(question)
        with st.chat_message("assistant"):
            st.write(answer)

else:
    st.info("Upload a CSV, Excel, PDF, Word, or PowerPoint file from the sidebar to get started.")