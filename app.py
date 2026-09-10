import streamlit as st

from marksheet_core import parse_marks_docx, compute_marks, generate_marksheet, MarkSheetError

st.set_page_config(page_title="Mark Sheet Generator", page_icon="📝", layout="centered")

st.title("📝 Mark Sheet Generator")
st.caption(
    "Developed by Dr. Rakesh Chandran S. B., IQAC Coordinator, S. D. College, Alappuzha. "
    "An IQAC SDC Initiative."
)
st.write(
    "Upload the student marks list (.docx with **Sl.No, Name, Marks** columns), "
    "fill in the course details, and generate a formatted mark sheet — out of **25** marks, "
    "with the **SSA component** calculated automatically."
)

st.divider()

# ---- Course details ----
st.subheader("Course details")
col1, col2 = st.columns(2)
with col1:
    department = st.text_input("Department", placeholder="e.g. Department of Computer Science")
    course_code = st.text_input("Course Code", placeholder="e.g. CS1234")
with col2:
    course_name = st.text_input("Course Name", placeholder="e.g. Data Structures")

st.subheader("Marks configuration")
col3, col4 = st.columns(2)
with col3:
    max_test_marks = st.number_input(
        "Maximum Test Marks",
        min_value=0.1,
        value=50.0,
        step=0.5,
        help="The maximum marks the test was conducted for. Used to scale each student's marks to out of 25.",
    )
with col4:
    ssa_component = st.number_input(
        "SSA Component (max marks)",
        min_value=0.0,
        value=5.0,
        step=0.5,
        help="SSA marks are derived as (Marks out of 25 ÷ 25) × this value.",
    )

st.subheader("Student marks list")
uploaded_file = st.file_uploader(
    "Upload the student marks (.docx) file",
    type=["docx"],
    help="The file must contain a table with Sl.No, Name and Marks columns.",
)

st.divider()

if uploaded_file is not None:
    try:
        df = parse_marks_docx(uploaded_file)
        result_df = compute_marks(df, max_test_marks=max_test_marks, ssa_component=ssa_component)
    except MarkSheetError as e:
        st.error(str(e))
    else:
        st.success(f"Loaded {len(result_df)} student(s) from the uploaded file.")

        display_df = result_df.rename(
            columns={
                "Marks_25": "Marks (out of 25)",
                "SSA": f"SSA (out of {ssa_component:g})",
            }
        )
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        if not department or not course_code or not course_name:
            st.warning("Fill in Department, Course Code and Course Name before generating the mark sheet.")
        else:
            docx_buffer = generate_marksheet(
                department=department,
                course_code=course_code,
                course_name=course_name,
                max_test_marks=max_test_marks,
                ssa_component=ssa_component,
                df=result_df,
            )
            file_name = f"MarkSheet_{course_code.strip().replace(' ', '_')}.docx"
            st.download_button(
                label="⬇️ Download Mark Sheet (.docx)",
                data=docx_buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
            )
            st.caption(
                "Tip: to generate a mark sheet for another batch, just upload a new list above — "
                "the previous one is not saved anywhere."
            )
else:
    st.info("Upload a student marks (.docx) file to get started.")
