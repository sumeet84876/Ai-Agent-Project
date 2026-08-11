"""
Web interface for the AI Data Analyst Agent, built with Streamlit.

Setup:
    pip install streamlit google-genai pandas openpyxl python-docx python-pptx pypdf
    export GEMINI_API_KEY="your-key-here"

Run:
    streamlit run app.py
"""

import os
import tempfile
import streamlit as st

st.set_page_config(page_title="AI Data Analyst Agent", layout="wide")
st.title("AI Data Analyst Agent")
st.caption("Upload a file, then explore it with buttons or plain-English chat.")

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

# ----------------------------------------------------------------------
# Sidebar: file upload + reset
# ----------------------------------------------------------------------
with st.sidebar:
    st.header("File")
    uploaded = st.file_uploader(
        "Upload a file",
        type=["csv", "xlsx", "xls", "pdf", "docx", "pptx"],
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

    if st.button("Reset / clear file"):
        for key in ["df", "filepath", "kind", "text_content", "history",
                    "last_file", "cleaning_log", "uploaded_name"]:
            st.session_state[key] = None if key != "history" and key != "cleaning_log" else []
        st.rerun()

    if st.session_state.get("last_file") and os.path.exists(st.session_state.last_file):
        with open(st.session_state.last_file, "rb") as f:
            st.download_button(
                "Download last generated file",
                f,
                file_name=os.path.basename(st.session_state.last_file),
            )

# ----------------------------------------------------------------------
# Main area
# ----------------------------------------------------------------------
if st.session_state.kind == "table" and st.session_state.df is not None:
    df = st.session_state.df
    base = os.path.splitext(st.session_state.filepath)[0]

    with st.expander("Cleaning log", expanded=True):
        for line in st.session_state.cleaning_log:
            st.write("- " + line)

    st.subheader("Data preview")
    st.dataframe(df.head(20), use_container_width=True)

    st.divider()
    st.subheader("Quick actions")

    row1 = st.columns(4)
    row2 = st.columns(4)

    def run_action(label, fn, *args):
        with st.spinner(f"Running {label}..."):
            result = fn(*args)
        st.session_state.history.append((label, result))

    if row1[0].button("Report"):
        run_action("Report", agent.generate_report, df)
    if row1[1].button("Save cleaned file"):
        run_action("Save", agent.save_data, df, base)
    if row1[2].button("Excel dashboard"):
        out = base + "_dashboard.xlsx"
        run_action("Dashboard", agent.generate_excel_dashboard, df, out)
        st.session_state.last_file = out
    if row1[3].button("Power BI export"):
        out = base + "_powerbi.xlsx"
        run_action("Power BI export", agent.export_for_powerbi, df, out)
        st.session_state.last_file = out

    if row2[0].button("SQL export (.db)"):
        out = base + ".db"
        run_action("SQL export", agent.export_sql_db, df, out)
        st.session_state.last_file = out
    if row2[1].button("Power BI guide"):
        out = base + "_powerbi_guide.md"
        run_action("Power BI guide", agent.generate_powerbi_guide, df, out)
        st.session_state.last_file = out
    if row2[2].button("Edit Excel in place"):
        out = base + "_edited.xlsx"
        run_action("Edit Excel", agent.edit_excel_in_place, st.session_state.filepath, df, out)
        st.session_state.last_file = out
    if row2[3].button("Re-run auto clean"):
        with st.spinner("Cleaning..."):
            new_df, log = agent.auto_clean(df)
        st.session_state.df = new_df
        st.session_state.cleaning_log = log
        st.rerun()

    st.divider()
    st.subheader("Chat")
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
    st.subheader("Document loaded")
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