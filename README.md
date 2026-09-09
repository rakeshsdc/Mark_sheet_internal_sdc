# Mark Sheet Generator

A Streamlit app that generates a formatted college mark sheet (`.docx`) from a
student marks list, scaling marks to **out of 25** and calculating the **SSA
component** automatically.

## What it does

1. You upload a `.docx` file containing a table with three columns:
   **Sl.No, Name, Marks** (see `sample_student_list.docx` for the expected format).
2. You fill in:
   - **Department**, **Course Code**, **Course Name**
   - **Maximum Test Marks** — the marks the test was actually conducted for
   - **SSA Component** — the maximum SSA marks
3. The app calculates, for every student:
   - `Marks (out of 25)` = `Marks ÷ Maximum Test Marks × 25` (rounded to 1 decimal)
   - `SSA` = `Marks (out of 25) ÷ 25 × SSA Component` (rounded to 1 decimal)
4. Click **Download Mark Sheet (.docx)** to get the final formatted mark sheet,
   with the college header, Department/Course Code/Course Name filled in, and a
   5-column table: **Sl.No | Name | Marks | Marks (out of 25) | SSA**.

You can repeat this any number of times with different lists/courses in the same
session — nothing is stored, each run produces a fresh download.

## Project structure

```
marksheet_app/
├── app.py                  # Streamlit UI
├── marksheet_core.py        # Parsing, calculation, and document generation logic
├── requirements.txt
├── assets/
│   └── logo.jpg              # College logo used in the generated mark sheet
└── README.md
```

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then open the URL Streamlit prints (usually `http://localhost:8501`).

## Deploy for free on Streamlit Community Cloud (hosted on GitHub)

1. **Create a GitHub repository** and push this folder's contents to it:

   ```bash
   git init
   git add .
   git commit -m "Mark sheet generator app"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```

2. Go to **[share.streamlit.io](https://share.streamlit.io)** and sign in with
   your GitHub account.
3. Click **"New app"**, select your repository, branch (`main`), and set the
   main file path to `app.py`.
4. Click **Deploy**. Streamlit Cloud will install `requirements.txt`
   automatically and give you a public URL you can share.

Any time you push a new commit to the repo, the hosted app updates automatically.

## Notes on the uploaded student list

- The app looks for columns containing "Sl", "Name", and "Mark" in the header
  row (case-insensitive), so headers like `Sl.no`, `SL NO`, `Marks Obtained`
  are all recognized — you don't need an exact match.
- Marks must be numeric (decimals like `21.5` are fine).
- Only the **first table** found in the uploaded `.docx` is read.

## Customizing the template

- To change the college name or exam title shown in the generated mark sheet,
  edit the `college_name` and `exam_title` defaults in
  `marksheet_core.generate_marksheet(...)`, or wire them up as additional
  text inputs in `app.py` if you want them editable per run.
- To change the logo, replace `assets/logo.jpg`.
