import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Loom Production Planner",
    page_icon="🏭",
    layout="wide"
)


# ============================================================
# DEFAULT LOOM SETTINGS
# ============================================================

DEFAULT_SETTINGS = pd.DataFrame([
    {"Use": True, "Loom": "ALT1", "Weekly Capacity": 1500, "Starting Week": 33},
    {"Use": True, "Loom": "ALT2", "Weekly Capacity": 1200, "Starting Week": 33},
    {"Use": True, "Loom": "ALT4", "Weekly Capacity": 800, "Starting Week": 34},
    {"Use": True, "Loom": "ALT5", "Weekly Capacity": 1300, "Starting Week": 33},
    {"Use": True, "Loom": "ALT6", "Weekly Capacity": 800, "Starting Week": 33},

    {"Use": True, "Loom": "RP1", "Weekly Capacity": 500, "Starting Week": 33},
    {"Use": True, "Loom": "RP2", "Weekly Capacity": 700, "Starting Week": 33},
    {"Use": True, "Loom": "RP3", "Weekly Capacity": 900, "Starting Week": 33},
    {"Use": True, "Loom": "RP4", "Weekly Capacity": 900, "Starting Week": 33},
    {"Use": True, "Loom": "RP5", "Weekly Capacity": 1100, "Starting Week": 33},
    {"Use": True, "Loom": "RP6", "Weekly Capacity": 1200, "Starting Week": 33},
    {"Use": True, "Loom": "RP7", "Weekly Capacity": 1000, "Starting Week": 33},

    {"Use": True, "Loom": "DB4M1", "Weekly Capacity": 300, "Starting Week": 33},
])


# ============================================================
# INITIALIZE SESSION STATE
# ============================================================

if "loom_settings" not in st.session_state:
    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()


# ============================================================
# HELPER FUNCTION
# ============================================================

def clean_loom_name(loom):

    if loom is None:
        return ""

    return (
        str(loom)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


# ============================================================
# TITLE
# ============================================================

st.title("🏭 Loom Production Planner")

st.write(
    "Upload the production Excel file, configure the loom settings, "
    "and automatically assign production weeks."
)


# ============================================================
# STEP 1 — UPLOAD EXCEL
# ============================================================

st.subheader("1️⃣ Upload Production Excel")

uploaded_file = st.file_uploader(
    "Select the Excel file",
    type=["xlsx"]
)


# ============================================================
# STEP 2 — LOOM SETTINGS
# ============================================================

st.subheader("2️⃣ Loom Settings")

st.caption(
    "Edit the weekly capacity and starting week. "
    "Uncheck a loom if it should not be used for this planning."
)


edited_settings = st.data_editor(
    st.session_state.loom_settings,
    use_container_width=True,
    hide_index=True,

    column_config={

        "Use": st.column_config.CheckboxColumn(
            "Use",
            help="Check to include this loom in production planning.",
            default=True
        ),

        "Loom": st.column_config.TextColumn(
            "Loom",
            disabled=True
        ),

        "Weekly Capacity": st.column_config.NumberColumn(
            "Weekly Capacity",
            help="Maximum quantity that can be produced by this loom per week.",
            min_value=1,
            step=50
        ),

        "Starting Week": st.column_config.NumberColumn(
            "Starting Week",
            help="Week number from which this loom should start.",
            min_value=1,
            max_value=52,
            step=1
        )
    },

    key="loom_settings_editor"
)


# ============================================================
# RESET SETTINGS
# ============================================================

if st.button("🔄 Reset to Default Settings"):

    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()

    if "loom_settings_editor" in st.session_state:
        del st.session_state["loom_settings_editor"]

    st.rerun()


# ============================================================
# STEP 3 — RUN PLANNING
# ============================================================

st.divider()

st.subheader("3️⃣ Run Production Planning")

run_button = st.button(
    "🚀 Run Production Planning",
    type="primary",
    use_container_width=True
)


# ============================================================
# MAIN PROCESSING
# ============================================================

if run_button:

    # --------------------------------------------------------
    # Check Excel upload
    # --------------------------------------------------------

    if uploaded_file is None:

        st.error("Please upload the Excel file first.")

        st.stop()


    # --------------------------------------------------------
    # Create loom configuration
    # --------------------------------------------------------

    loom_settings = {}

    for _, row in edited_settings.iterrows():

        loom = clean_loom_name(row["Loom"])

        use_loom = bool(row["Use"])

        capacity = float(row["Weekly Capacity"])

        starting_week = int(row["Starting Week"])

        # Safety check
        if starting_week < 1 or starting_week > 52:

            st.error(
                f"Invalid starting week for {loom}. "
                "Starting week must be between 1 and 52."
            )

            st.stop()

        if capacity <= 0:

            st.error(
                f"Invalid capacity for {loom}. "
                "Capacity must be greater than 0."
            )

            st.stop()

        loom_settings[loom] = {
            "use": use_loom,
            "capacity": capacity,
            "starting_week": starting_week
        }


    # --------------------------------------------------------
    # Load workbook
    # --------------------------------------------------------

    try:

        file_data = uploaded_file.getvalue()

        workbook = load_workbook(
            filename=BytesIO(file_data)
        )

    except Exception as e:

        st.error(
            f"Unable to open the Excel file: {e}"
        )

        st.stop()


    # --------------------------------------------------------
    # Check required sheet
    # --------------------------------------------------------

    SHEET_NAME = "Loom wise Structure"

    if SHEET_NAME not in workbook.sheetnames:

        st.error(
            f"Sheet '{SHEET_NAME}' was not found in the Excel file."
        )

        st.stop()


    worksheet = workbook[SHEET_NAME]


    # ========================================================
    # FIXED COLUMN STRUCTURE
    # ========================================================

    # A = No.
    # B = External Document
    # C = Description
    # D = Roll Width
    # E = Roll Length
    # F = Quantity
    # G = Loom
    # H = Production Week

    QUANTITY_COL = 6
    LOOM_COL = 7
    PRODUCTION_WEEK_COL = 8


    # ========================================================
    # CREATE PRODUCTION WEEK COLUMN
    # ========================================================

    worksheet.cell(
        row=1,
        column=PRODUCTION_WEEK_COL
    ).value = "Production Week"


    # Clear existing Production Week values

    for row in range(2, worksheet.max_row + 1):

        worksheet.cell(
            row=row,
            column=PRODUCTION_WEEK_COL
        ).value = None


    # ========================================================
    # PROCESS EACH LOOM INDEPENDENTLY
    # ========================================================

    summary = []


    for loom, settings in loom_settings.items():

        # ----------------------------------------------------
        # Skip unchecked looms
        # ----------------------------------------------------

        if not settings["use"]:

            continue


        capacity = settings["capacity"]

        current_week = settings["starting_week"]

        current_load = 0

        rows_processed = 0

        weeks_used = set()


        # ----------------------------------------------------
        # Go through Excel rows in ORIGINAL ORDER
        # ----------------------------------------------------

        for excel_row in range(
            2,
            worksheet.max_row + 1
        ):

            excel_loom = worksheet.cell(
                row=excel_row,
                column=LOOM_COL
            ).value


            cleaned_excel_loom = clean_loom_name(
                excel_loom
            )


            # ------------------------------------------------
            # Only process this loom
            # ------------------------------------------------

            if cleaned_excel_loom != loom:

                continue


            # ------------------------------------------------
            # Get quantity
            # ------------------------------------------------

            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


            # Skip blank quantity

            if quantity is None:

                continue


            # ------------------------------------------------
            # Convert quantity to number
            # ------------------------------------------------

            try:

                quantity = float(quantity)

            except (ValueError, TypeError):

                continue


            # ------------------------------------------------
            # Check whether quantity fits in current week
            # ------------------------------------------------

            if current_load + quantity <= capacity:

                # Quantity fits in current week

                assigned_week = current_week

                current_load += quantity


            else:

                # ------------------------------------------------
                # Move entire order to next week
                # ------------------------------------------------

                current_week += 1


                # ------------------------------------------------
                # WEEK 52 → WEEK 1
                # ------------------------------------------------

                if current_week > 52:

                    current_week = 1


                assigned_week = current_week

                current_load = quantity


            # ------------------------------------------------
            # Write Production Week
            # ------------------------------------------------

            worksheet.cell(
                row=excel_row,
                column=PRODUCTION_WEEK_COL
            ).value = f"WK-{assigned_week}"


            rows_processed += 1

            weeks_used.add(assigned_week)


        # ----------------------------------------------------
        # Add loom summary
        # ----------------------------------------------------

        if rows_processed > 0:

            summary.append({

                "Loom": loom,

                "Weekly Capacity": capacity,

                "Starting Week":
                    f"WK-{settings['starting_week']}",

                "Final Week":
                    f"WK-{current_week}",

                "Weeks Used":
                    len(weeks_used),

                "Rows Processed":
                    rows_processed
            })


    # ========================================================
    # SAVE OUTPUT
    # ========================================================

    output = BytesIO()

    workbook.save(output)

    output.seek(0)


    # ========================================================
    # SUCCESS MESSAGE
    # ========================================================

    st.success(
        "✅ Production planning completed successfully!"
    )


    # ========================================================
    # PLANNING SUMMARY
    # ========================================================

    st.subheader("📊 Planning Summary")


    if len(summary) > 0:

        summary_df = pd.DataFrame(summary)

        st.dataframe(
            summary_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
            "No enabled loom data was found in the Excel file."
        )


    # ========================================================
    # DOWNLOAD BUTTON
    # ========================================================

    st.subheader("4️⃣ Download Result")

    st.download_button(

        label="📥 Download Planned Excel",

        data=output,

        file_name="Loom_wise_Production_Planning_Automated.xlsx",

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        use_container_width=True
    )
