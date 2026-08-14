"""
Web interface for the AI Data Analyst Agent, built with Streamlit.

Setup:
    pip install -r requirements.txt
    export GEMINI_API_KEY="your-key-here"   (Windows: set as system env var)

Run:
    streamlit run app.py

Design: minimal chat interface (like talking to Claude/ChatGPT with a file
attached). Upload a file -> agent asks what to do -> give an instruction in
plain language (Hinglish or English) -> agent actually performs the action
and gives a download link. The "+" menu next to the chat box jumps straight
to a specific plugin (Excel Dashboard / SQL Query / Power BI Guide & DAX)
without needing the agent to guess intent.

Theme colors live in .streamlit/config.toml — keep that folder alongside
this file (some file-transfer methods hide dotfolders, check for it).
"""

import os
import uuid
import streamlit as st

st.set_page_config(page_title="AI Data Analyst Agent", page_icon="◆", layout="centered")

# ----------------------------------------------------------------------
# Custom styling — same "data terminal" look as before.
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
    .app-title { font-size: 1.6rem; font-weight: 800; color: #F0B429; letter-spacing: -0.02em; }
    .app-subtitle { color: #8A93A6; font-size: 0.92rem; }

    .stButton>button {
        border: 1px solid #2A3140;
        background-color: #1B2129;
        color: #E8E6E1;
        border-radius: 6px;
        font-family: 'JetBrains Mono', monospace;
        font-size: 0.82rem;
        transition: border-color 0.15s ease;
    }
    .stButton>button:hover { border-color: #F0B429; color: #F0B429; }

    section[data-testid="stSidebar"] { border-right: 1px solid #2A3140; }
    </style>

    <div class="app-header">
        <span class="app-title">◆ AI DATA ANALYST AGENT</span>
        <span class="app-subtitle">upload &middot; chat &middot; done</span>
    </div>
    """,
    unsafe_allow_html=True,
)

# Check the API key BEFORE importing agent.py, since agent.py creates the
# Gemini client at import time and would otherwise crash with a raw
# traceback instead of a friendly message.
if not (os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")):
    st.error(
        "GEMINI_API_KEY is not set. Get a free key at "
        "https://aistudio.google.com/apikey and set it as a system "
        "environment variable, then restart this app."
    )
    st.stop()

import agent          # reuses all the logic already built in agent.py
import memory_store as mem

mem.init_db()

# ----------------------------------------------------------------------
# Session state
# ----------------------------------------------------------------------
if "session_id" not in st.session_state:
    st.session_state.session_id = str(uuid.uuid4())
    mem.create_session(st.session_state.session_id)

for key, default in [
    ("df", None), ("filepath", None), ("kind", None), ("text_content", None),
    ("output_base", None), ("uploaded_name", None), ("forced_mode", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


def new_chat():
    st.session_state.session_id = str(uuid.uuid4())
    mem.create_session(st.session_state.session_id)
    for key in ["df", "filepath", "kind", "text_content", "output_base",
                "uploaded_name", "forced_mode"]:
        st.session_state[key] = None


# ----------------------------------------------------------------------
# Sidebar — chat history (persistent) + active file status
# ----------------------------------------------------------------------
with st.sidebar:
    st.markdown("### Chats")
    if st.button("➕ New chat", use_container_width=True):
        new_chat()
        st.rerun()

    for s in mem.list_sessions():
        label = (s["title"] or "New chat")[:40]
        if st.button(label, key=f"sess_{s['session_id']}", use_container_width=True):
            st.session_state.session_id = s["session_id"]
            st.session_state.df = None
            st.session_state.filepath = None
            st.session_state.kind = None
            st.session_state.text_content = None
            st.session_state.output_base = None
            st.session_state.uploaded_name = None
            st.rerun()

    st.divider()

    if st.session_state.uploaded_name:
        st.markdown("**ACTIVE FILE**")
        st.caption(st.session_state.uploaded_name)
        if st.session_state.kind == "table" and st.session_state.df is not None:
            r, c = st.session_state.df.shape
            st.caption(f"{r} rows x {c} columns")
        st.caption(f"Outputs saved to: `{agent.OUTPUTS_DIR}`")

    st.divider()
    if st.button("🗑️ Clear ALL history", use_container_width=True):
        st.session_state.confirm_clear = True

    if st.session_state.get("confirm_clear"):
        st.warning("Ye permanent hai — saara chat history delete ho jayega.")
        c1, c2 = st.columns(2)
        if c1.button("Haan, delete", use_container_width=True):
            mem.clear_all()
            new_chat()
            st.session_state.confirm_clear = False
            st.rerun()
        if c2.button("Cancel", use_container_width=True):
            st.session_state.confirm_clear = False
            st.rerun()

# ----------------------------------------------------------------------
# Render chat history
# ----------------------------------------------------------------------
history = mem.get_messages(st.session_state.session_id)
for msg in history:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])
        if msg["file_ref"] and os.path.exists(msg["file_ref"]):
            with open(msg["file_ref"], "rb") as f:
                st.download_button(
                    "⬇️ Download file", f, file_name=os.path.basename(msg["file_ref"]),
                    key=f"dl_{msg['created_at']}",
                )

# ----------------------------------------------------------------------
# "+" plugin menu + file uploader
# ----------------------------------------------------------------------
col_plus, col_upload = st.columns([1, 4])
with col_plus:
    with st.popover("➕"):
        st.caption("Plugins")
        if st.button("📊 Excel Dashboard", use_container_width=True):
            st.session_state.forced_mode = "dashboard"
        if st.button("🗄️ SQL Query", use_container_width=True):
            st.session_state.forced_mode = "sql" if st.session_state.kind == "table" else "sql_question"
        if st.button("📈 Power BI Guide & DAX", use_container_width=True):
            st.session_state.forced_mode = "powerbi_guide" if st.session_state.kind == "table" else "powerbi_question"
        if st.session_state.forced_mode:
            st.success(f"Mode set: {st.session_state.forced_mode}")
            if st.button("Clear mode"):
                st.session_state.forced_mode = None

with col_upload:
    uploaded = st.file_uploader(
        "Upload file", type=["csv", "xlsx", "xls", "pdf", "docx", "pptx"],
        label_visibility="collapsed",
    )

if uploaded is not None and st.session_state.uploaded_name != uploaded.name:
    upload_id = uuid.uuid4().hex[:6]
    saved_path = os.path.join(agent.UPLOADS_DIR, f"{upload_id}_{uploaded.name}")
    with open(saved_path, "wb") as f:
        f.write(uploaded.getbuffer())

    st.session_state.uploaded_name = uploaded.name
    st.session_state.filepath = saved_path
    stem = os.path.splitext(uploaded.name)[0]
    st.session_state.output_base = os.path.join(agent.OUTPUTS_DIR, f"{stem}_{upload_id}")

    with st.spinner("Loading..."):
        kind, content = agent.load_any_file(saved_path)
        st.session_state.kind = kind
        if kind == "table":
            df, log = agent.auto_clean(content)
            st.session_state.df = df
            ask_msg = (
                f"'{uploaded.name}' load ho gayi ({df.shape[0]} rows, {df.shape[1]} columns). "
                f"Auto-clean check kiya:\n\n" + "\n".join(f"- {l}" for l in log) +
                "\n\nAb isse kya karna hai? Kuch options:\n"
                "- Data quality **report** chahiye\n"
                "- Kuch **clean/edit** karna hai (batao kya)\n"
                "- Excel **dashboard** banana hai\n"
                "- **SQL** se query chalani hai\n"
                "- **Power BI** guide/DAX chahiye\n\n"
                "Bas message me instruction likh do."
            )
        else:
            st.session_state.text_content = content
            ask_msg = (
                f"'{uploaded.name}' load ho gayi ({len(content)} characters). "
                "Iske baare me kuch bhi pooch sakte ho — bas message likh do."
            )

    mem.add_message(st.session_state.session_id, "assistant", ask_msg)
    with st.chat_message("assistant"):
        st.markdown(ask_msg)

# ----------------------------------------------------------------------
# Chat input + routing
# ----------------------------------------------------------------------
user_input = st.chat_input("Message the agent...")

if user_input:
    mem.add_message(st.session_state.session_id, "user", user_input)
    with st.chat_message("user"):
        st.markdown(user_input)

    existing = mem.get_messages(st.session_state.session_id)
    if len(existing) <= 2:
        mem.rename_session(st.session_state.session_id, user_input[:40])

    has_file = st.session_state.kind is not None

    if st.session_state.forced_mode:
        intent = st.session_state.forced_mode
        st.session_state.forced_mode = None  # one-shot
    else:
        with st.spinner("Understanding..."):
            intent = agent.detect_intent(user_input, has_active_file=has_file)

    reply_text = ""
    file_ref = None
    base = st.session_state.output_base

    with st.chat_message("assistant"):
        with st.spinner("Working..."):
            try:
                needs_table = intent in (
                    "report", "clean", "ask", "sql", "save", "sql_export",
                    "dashboard", "formula", "format", "validate", "edit_excel",
                    "powerbi_export", "powerbi_guide",
                )
                if needs_table and st.session_state.kind != "table":
                    reply_text = "Ye kaam ke liye ek CSV/Excel file upload karo pehle."

                elif intent == "report":
                    reply_text = agent.generate_report(st.session_state.df)

                elif intent == "clean":
                    new_df, msg = agent.custom_clean(st.session_state.df, user_input)
                    st.session_state.df = new_df
                    reply_text = msg

                elif intent == "ask":
                    reply_text = agent.ask_question(st.session_state.df, user_input)

                elif intent == "sql":
                    reply_text = agent.run_sql(st.session_state.df, user_input)

                elif intent == "save":
                    reply_text = agent.save_data(st.session_state.df, base)
                    file_ref = base + "_cleaned.xlsx"

                elif intent == "sql_export":
                    out = base + ".db"
                    reply_text = agent.export_sql_db(st.session_state.df, out)
                    file_ref = out

                elif intent == "dashboard":
                    out = base + "_dashboard.xlsx"
                    reply_text = agent.generate_excel_dashboard(st.session_state.df, out)
                    file_ref = out

                elif intent == "formula":
                    out = base + "_formula.xlsx"
                    reply_text = agent.add_formula_column(st.session_state.df, user_input, out)
                    file_ref = out

                elif intent == "format":
                    out = base + "_formatted.xlsx"
                    reply_text = agent.add_conditional_formatting(st.session_state.df, user_input, out)
                    file_ref = out

                elif intent == "validate":
                    out = base + "_validated.xlsx"
                    reply_text = agent.add_data_validation(st.session_state.df, user_input, out)
                    file_ref = out

                elif intent == "edit_excel":
                    out = base + "_edited.xlsx"
                    reply_text = agent.edit_excel_in_place(st.session_state.filepath, st.session_state.df, out)
                    file_ref = out

                elif intent == "powerbi_export":
                    out = base + "_powerbi.xlsx"
                    reply_text = agent.export_for_powerbi(st.session_state.df, out)
                    file_ref = out

                elif intent == "powerbi_guide":
                    out = base + "_powerbi_guide.md"
                    reply_text = agent.generate_powerbi_guide(st.session_state.df, out)
                    file_ref = out

                elif intent == "excel_question":
                    reply_text = agent.answer_excel_question(user_input)

                elif intent == "sql_question":
                    reply_text = agent.answer_sql_question(user_input)

                elif intent == "powerbi_question":
                    reply_text = agent.answer_powerbi_question(user_input)

                else:  # general_chat
                    if has_file and st.session_state.kind == "text":
                        reply_text = agent.ask_about_text(st.session_state.text_content, user_input)
                    else:
                        reply_text = agent.chat_reply(user_input)

            except Exception as e:
                reply_text = f"Kuch error aaya: `{e}`"

        st.markdown(reply_text)
        if file_ref and os.path.exists(file_ref):
            with open(file_ref, "rb") as f:
                st.download_button("⬇️ Download file", f, file_name=os.path.basename(file_ref))

    mem.add_message(st.session_state.session_id, "assistant", reply_text, file_ref=file_ref)

elif st.session_state.kind is None and not history:
    st.info("Upload a CSV, Excel, PDF, Word, or PowerPoint file to get started.")