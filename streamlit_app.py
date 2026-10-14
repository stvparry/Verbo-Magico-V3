import re
import traceback
from urllib.parse import quote

import requests
import streamlit as st

from generator import build_workbook

st.set_page_config(
    page_title="Verbo Mágico",
    page_icon="🇪🇸",
    layout="centered",
)

st.markdown(
    """
    <style>
    .block-container {
        max-width: 900px;
        padding-top: 2rem;
    }

    .hero {
        padding: 1.5rem;
        border-radius: 12px;
        background: linear-gradient(135deg,#17365D,#2F75B5);
        color: white;
        margin-bottom: 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
        <h1>🇪🇸 Verbo Mágico</h1>
        <p>Generate a Spanish verb workbook from any infinitive.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


@st.cache_data(ttl=86400)
def fetch_conjugation(verb: str):

    url = f"https://verbe.cc/verbecc/conjugate/es/{quote(verb)}"

    response = requests.get(
        url,
        timeout=30,
        headers={"Accept": "application/json"},
    )

    response.raise_for_status()

    return response.json()


st.subheader("Spanish infinitive")

verb = st.text_input(
    "",
    value="tener",
    placeholder="tener, hablar, ser, ir...",
).strip().lower()

valid = (
    bool(re.fullmatch(r"[a-záéíóúüñ]+", verb))
    and verb.endswith(("ar", "er", "ir"))
)

if verb and not valid:
    st.warning(
        "Please enter a Spanish infinitive ending in -ar, -er or -ir."
    )

if st.button(
    "Generate workbook",
    disabled=not valid,
    use_container_width=True,
):

    try:

        with st.spinner(f"Fetching conjugation for '{verb}'..."):

            payload = fetch_conjugation(verb)

            workbook = build_workbook(
                verb,
                payload,
            )

        st.success("Workbook ready")

        st.download_button(
            label="📥 Download workbook",
            data=workbook,
            file_name=f"{verb}_complete_conjugation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception:

        st.error("Workbook generation failed.")

        with st.expander("Technical details"):

            st.code(
                traceback.format_exc(),
                language="text",
            )

st.divider()

st.caption(
    "Version 3 • Streamlit Cloud • No Windows install • Workbook download"
)
