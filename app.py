import io
import json
import streamlit as st
import pandas as pd
import fitz
from PIL import Image
from groq import Groq
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt


MODEL = "openai/gpt-oss-120b"


st.set_page_config(
    page_title="AI Letter OCR Extractor",
    page_icon="📄",
    layout="wide",
)


def image_to_bytes(image):
    buf = io.BytesIO()
    image.save(buf, format="PNG")
    return buf.getvalue()


def pdf_to_images(pdf_bytes):
    """Render every PDF page as a high-resolution PNG."""
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    images = []

    for page in pdf:
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
        images.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))

    pdf.close()
    return images


def extract_text_with_groq(images, api_key):
    """
    Use Groq's vision-capable model to read the uploaded document.
    GPT-OSS-120B is then used separately for structured extraction.
    """
    # Vision OCR model. The extracted text is passed to GPT-OSS-120B.
    # This keeps the requested GPT-OSS-120B model as the extraction/structuring model.
    vision_model = "qwen/qwen3.8-27b"
    client = Groq(api_key=api_key)

    chunks = []

    for index, image in enumerate(images, start=1):
        image_b64 = __import__("base64").b64encode(
            image_to_bytes(image)
        ).decode("utf-8")

        response = client.chat.completions.create(
            model=vision_model,
            temperature=0,
            max_tokens=6000,
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an OCR transcription assistant for official "
                        "government letters. Read the document image carefully. "
                        "Transcribe all visible text as accurately as possible. "
                        "Preserve letter numbers, dates, organization names, "
                        "addresses, and subject lines. Do not summarize or invent text."
                    ),
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": f"Transcribe page {index} exactly. Preserve important line breaks.",
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{image_b64}"
                            },
                        },
                    ],
                },
            ],
        )

        chunks.append(f"--- PAGE {index} ---\n{response.choices[0].message.content}")

    return "\n\n".join(chunks)


def extract_fields_with_gpt_oss(ocr_text, api_key):
    """Convert OCR text into the required structured fields."""
    client = Groq(api_key=api_key)

    system_prompt = """
You are an expert official-letter data extraction system.

Extract information ONLY from the supplied OCR text.

Required fields:
1. letter_no
2. dated
3. address
4. subject

Rules:
- Do not invent or guess information.
- Preserve the official letter number exactly when possible.
- Normalize only obvious OCR mistakes.
- The dated field should contain the letter date, not another date mentioned in the body.
- Address means the issuing organization/header shown on the letter, for example:
  "Government of Khyber Pakhtunkhwa, Right to Public Services Commission".
- Subject must be selected from the letter's Subject/SUBJECT line, not invented from the body.
- Remove obvious OCR noise.
- If a field cannot be reliably identified, return an empty string.
- Return JSON only.
"""

    response = client.chat.completions.create(
        model=MODEL,
        temperature=0,
        messages=[
            {"role": "system", "content": system_prompt},
            {
                "role": "user",
                "content": f"Extract the fields from this OCR text:\n\n{ocr_text}",
            },
        ],
        response_format={
            "type": "json_schema",
            "json_schema": {
                "name": "letter_record",
                "strict": True,
                "schema": {
                    "type": "object",
                    "properties": {
                        "letter_no": {"type": "string"},
                        "dated": {"type": "string"},
                        "address": {"type": "string"},
                        "subject": {"type": "string"},
                    },
                    "required": ["letter_no", "dated", "address", "subject"],
                    "additionalProperties": False,
                },
            },
        },
    )

    return json.loads(response.choices[0].message.content)


def make_docx(rows):
    """Create a Google-Docs-compatible DOCX with the requested table."""
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)
    section.left_margin = Inches(0.5)
    section.right_margin = Inches(0.5)

    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    headers = ["S. no", "Letter no", "Dated", "Address", "Subject"]
    widths = [0.55, 1.25, 0.85, 2.35, 3.25]

    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        cell.width = Inches(widths[i])
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

        for p in cell.paragraphs:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.bold = True
                run.font.size = Pt(9)

    for row in rows:
        cells = table.add_row().cells
        values = [
            str(row["S. no"]),
            row["Letter no"],
            row["Dated"],
            row["Address"],
            row["Subject"],
        ]

        for i, value in enumerate(values):
            cells[i].text = value
            cells[i].width = Inches(widths[i])
            cells[i].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

            for p in cells[i].paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if i >= 1 else WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.size = Pt(8.5)

    # Add a small amount of document metadata.
    props = doc.core_properties
    props.title = "Extracted Letter Records"
    props.subject = "OCR extracted official letter records"

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def main():
    st.title("📄 AI Government Letter OCR Extractor")
    st.caption(
        "Upload PDF/image files, extract letter information with OCR + AI, "
        "review it, and generate a Google Docs-compatible DOCX."
    )

    with st.sidebar:
        st.header("Settings")
        api_key = st.text_input(
            "Groq API Key",
            value=st.secrets.get("GROQ_API_KEY", ""),
            type="password",
            help="For Streamlit Cloud, store this as GROQ_API_KEY in Secrets.",
        )

        st.markdown(
            "**AI model:** `openai/gpt-oss-120b`"
        )
        st.markdown(
            "Vision/OCR: `qwen/qwen3.8-27b`\n\n"
            "Structured extraction: `openai/gpt-oss-120b`."
        )

    uploaded = st.file_uploader(
        "Upload one or more letters",
        type=["pdf", "png", "jpg", "jpeg"],
        accept_multiple_files=True,
    )

    if not uploaded:
        st.info(
            "Upload a scanned PDF or image to begin. You can upload multiple "
            "letters and the app will number them automatically from 1."
        )
        return

    if not api_key:
        st.warning("Enter your Groq API key in the sidebar before processing.")
        return

    if "records" not in st.session_state:
        st.session_state.records = []

    if st.button("🚀 Process Uploaded Files", type="primary", use_container_width=True):
        records = []
        all_ocr = []

        progress = st.progress(0)

        try:
            for file_index, file in enumerate(uploaded):
                file_bytes = file.getvalue()

                if file.type == "application/pdf":
                    images = pdf_to_images(file_bytes)
                else:
                    images = [Image.open(io.BytesIO(file_bytes)).convert("RGB")]

                st.write(f"**Processing:** {file.name}")

                ocr_text = extract_text_with_groq(images, api_key)
                all_ocr.append(f"===== {file.name} =====\n{ocr_text}")

                data = extract_fields_with_gpt_oss(ocr_text, api_key)

                records.append(
                    {
                        "S. no": file_index + 1,
                        "Letter no": data.get("letter_no", ""),
                        "Dated": data.get("dated", ""),
                        "Address": data.get("address", ""),
                        "Subject": data.get("subject", ""),
                    }
                )

                progress.progress((file_index + 1) / len(uploaded))

            st.session_state.records = records
            st.session_state.ocr_text = "\n\n".join(all_ocr)
            st.success(f"Processed {len(records)} file(s) successfully.")

        except Exception as e:
            st.error(f"Processing failed: {e}")
            st.exception(e)

    if not st.session_state.get("records"):
        return

    st.divider()
    st.subheader("✏️ Review and Correct Extracted Information")
    st.caption(
        "Check the extracted fields before generating the final document. "
        "You can manually correct OCR/AI mistakes."
    )

    df = pd.DataFrame(st.session_state.records)

    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "S. no": st.column_config.NumberColumn("S. no", min_value=1, step=1),
            "Letter no": st.column_config.TextColumn("Letter no"),
            "Dated": st.column_config.TextColumn("Dated"),
            "Address": st.column_config.TextColumn("Address", width="large"),
            "Subject": st.column_config.TextColumn("Subject", width="large"),
        },
    )

    # Keep serial numbers sequential.
    edited_df["S. no"] = range(1, len(edited_df) + 1)
    st.session_state.records = edited_df.to_dict("records")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("🔄 Renumber S. no", use_container_width=True):
            st.rerun()

    with col2:
        docx_bytes = make_docx(st.session_state.records)
        st.download_button(
            "📥 Download DOCX for Google Docs",
            data=docx_bytes,
            file_name="letter_records.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with st.expander("🔎 View Raw OCR Text"):
        st.text_area(
            "OCR output",
            st.session_state.get("ocr_text", ""),
            height=400,
        )


if __name__ == "__main__":
    main()
