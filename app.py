import io
import tempfile
import zipfile
from pathlib import Path

import streamlit as st

import po_pdf_to_excel as conv

st.set_page_config(page_title="PO PDF to Excel", page_icon="📄")
st.title("📄 PO PDF → Excel Converter")
st.write("Upload one or more purchase order PDFs and download them as Excel files.")

files = st.file_uploader("Select PDF files", type="pdf", accept_multiple_files=True)

if files and st.button("Convert to Excel", type="primary"):
    results = []
    bar = st.progress(0.0)
    with tempfile.TemporaryDirectory() as tmp:
        for i, f in enumerate(files, 1):
            src = Path(tmp) / f.name
            src.write_bytes(f.getbuffer())
            dest = Path(tmp) / (src.stem + ".xlsx")
            try:
                po = conv.parse_po(src)
                problems, total = conv.validate(po)
                conv.write_excel(po, dest)
                results.append((dest.name, dest.read_bytes()))
                st.success(f"{f.name}: {len(po['items'])} lines, total ${total:,.2f}")
                for p in problems:
                    st.warning(p)
                if not problems:
                    st.caption("All checks passed (line math, item count, subtotal match the PDF).")
            except Exception as e:
                st.error(f"{f.name}: could not convert ({e})")
            bar.progress(i / len(files))

    if len(results) == 1:
        name, data = results[0]
        st.download_button("⬇️ Download Excel", data, file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    elif results:
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w") as z:
            for name, data in results:
                z.writestr(name, data)
        st.download_button("⬇️ Download all (ZIP)", buf.getvalue(), file_name="converted_po.zip",
                           mime="application/zip")
