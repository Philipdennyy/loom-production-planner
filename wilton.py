import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO
from datetime import date, timedelta


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
# SESSION STATE
# ============================================================

if "loom_settings" not in st.session_state:
    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_loom_name(loom):
    """Make loom names consistent for comparison."""

    if loom is None:
        return ""

    return (
        str(loom)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


def next_week(week):
    """Move to the next production week."""

    if week >= 52:
        return 1

    return week + 1


def week_difference(start_week, target_week):
    """
    Number of weeks between starting week and target week.

    Handles:
    WK-52 -> WK-1
    """

    return (target_week - start_week) % 52


def get_week_start_date(
    planning_start_date,
    starting_week,
    production_week
):
    """
    Convert a production week into its starting date.

    Example:

    Start date = 31-08-2026
    Starting week = 33

    WK-33 = 31-08-2026
    WK-34 = 07-09-2026
    WK-35 = 14-09-2026
    """

    difference = week_difference(
        starting_week,
        production_week
    )

    return (
        planning_start_date
        + timedelta(days=difference * 7)
    )


# ============================================================
# TITLE
# ============================================================

st.title("🏭 Loom Production Planner")

st.write(
    "Automatically assign production weeks and production dates "
    "based on loom capacity and production quantity."
)


# ============================================================
# STEP 1 — UPLOAD FILE
# ============================================================

st.subheader("1️⃣ Upload Production Excel")

uploaded_file = st.file_uploader(
    "Select the Excel file",
    type=["xlsx"]
)


# ============================================================
# STEP 2 — STARTING DATE
# ============================================================

st.subheader("2️⃣ Planning Start Date")

planning_start_date = st.date_input(
    "Select the date from which production planning starts",
    value=date.today()
)

st.caption(
    "The selected date is the first day of the starting production week. "
    "Each production week contains 7 working days."
)


# ============================================================
# STEP 3 — LOOM SETTINGS
# ============================================================

st.subheader("3️⃣ Loom Settings")

st.caption(
    "Select the looms required for planning and edit their "
    "weekly capacity and starting week."
)


edited_settings = st.data_editor(
    st.session_state.loom_settings,
    use_container_width=True,
    hide_index=True,

    column_config={

        "Use": st.column_config.CheckboxColumn(
            "Use",
            help="Check to include this loom.",
            default=True
        ),

        "Loom": st.column_config.TextColumn(
            "Loom",
            disabled=True
        ),

        "Weekly Capacity": st.column_config.NumberColumn(
            "Weekly Capacity",
            help="Maximum quantity that the loom can produce per week.",
            min_value=1,
            step=50
        ),

        "Starting Week": st.column_config.NumberColumn(
            "Starting Week",
            help="Production week from which this loom starts.",
            min_value=1,
            max_value=52,
            step=1
        )
    },

    key="loom_settings_editor"
)


# ============================================================
# RESET BUTTON
# ============================================================

if st.button("🔄 Reset to Default Settings"):

    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()

    if "loom_settings_editor" in st.session_state:
        del st.session_state["loom_settings_editor"]

    st.rerun()


# ============================================================
# RUN BUTTON
# ============================================================

st.divider()

run_button = st.button(
    "🚀 Run Production Planning",
    type="primary",
    use_container_width=True
)


# ============================================================
# MAIN PROCESS
# ============================================================

if run_button:

    # ========================================================
    # CHECK FILE
    # ========================================================

    if uploaded_file is None:

        st.error(
            "Please upload the Excel file first."
        )

        st.stop()


    # ========================================================
    # CREATE LOOM CONFIGURATION
    # ========================================================

    loom_settings = {}

    for _, row in edited_settings.iterrows():

        loom = clean_loom_name(row["Loom"])

        use_loom = bool(row["Use"])

        capacity = float(row["Weekly Capacity"])

        starting_week = int(row["Starting Week"])


        # ----------------------------------------------------
        # VALIDATION
        # ----------------------------------------------------

        if capacity <= 0:

            st.error(
                f"Weekly capacity for {loom} must be greater than 0."
            )

            st.stop()


        if starting_week < 1 or starting_week > 52:

            st.error(
                f"Starting week for {loom} must be between 1 and 52."
            )

            st.stop()


        loom_settings[loom] = {
            "use": use_loom,
            "capacity": capacity,
            "starting_week": starting_week
        }


    # ========================================================
    # LOAD EXCEL
    # ========================================================

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


    # ========================================================
    # SELECT FIRST WORKSHEET
    # ========================================================

    if not workbook.sheetnames:

        st.error(
            "The Excel file contains no worksheets."
        )

        st.stop()


    worksheet = workbook[
        workbook.sheetnames[0]
    ]


    # ========================================================
    # INPUT COLUMNS
    # ========================================================

    # F = Quantity
    # G = Loom
    #
    # New output:
    #
    # H = Production Week
    # I = Production Date
    # J = Day

    QUANTITY_COL = 6
    LOOM_COL = 7


    # ========================================================
    # MAKE SURE OUTPUT COLUMNS EXIST
    # ========================================================

    # --------------------------------------------------------
    # Production Week
    # --------------------------------------------------------

    h_value = worksheet.cell(
        row=1,
        column=8
    ).value


    if (
        h_value is not None
        and
        str(h_value).strip().lower()
        == "production week"
    ):

        production_week_col = 8

    elif all(
        worksheet.cell(
            row=row,
            column=8
        ).value is None

        for row in range(
            1,
            worksheet.max_row + 1
        )
    ):

        production_week_col = 8

    else:

        # H already contains something else.
        # Insert a new H.

        worksheet.insert_cols(8)

        production_week_col = 8


    worksheet.cell(
        row=1,
        column=production_week_col
    ).value = "Production Week"


    # --------------------------------------------------------
    # Production Date
    # --------------------------------------------------------

    production_date_col = production_week_col + 1

    date_header = worksheet.cell(
        row=1,
        column=production_date_col
    ).value


    if not (
        date_header is not None
        and
        str(date_header).strip().lower()
        == "production date"
    ):

        if any(
            worksheet.cell(
                row=row,
                column=production_date_col
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        ):

            worksheet.insert_cols(
                production_date_col
            )


    worksheet.cell(
        row=1,
        column=production_date_col
    ).value = "Production Date"


    # --------------------------------------------------------
    # Day
    # --------------------------------------------------------

    day_col = production_date_col + 1

    day_header = worksheet.cell(
        row=1,
        column=day_col
    ).value


    if not (
        day_header is not None
        and
        str(day_header).strip().lower()
        == "day"
    ):

        if any(
            worksheet.cell(
                row=row,
                column=day_col
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        ):

            worksheet.insert_cols(
                day_col
            )


    worksheet.cell(
        row=1,
        column=day_col
    ).value = "Day"


    # ========================================================
    # CLEAR OLD OUTPUT
    # ========================================================

    for row in range(
        2,
        worksheet.max_row + 1
    ):

        worksheet.cell(
            row=row,
            column=production_week_col
        ).value = None

        worksheet.cell(
            row=row,
            column=production_date_col
        ).value = None

        worksheet.cell(
            row=row,
            column=day_col
        ).value = None


    # ========================================================
    # FIRST PASS
    #
    # Determine Production Week ONLY.
    #
    # Date allocation is NOT involved here.
    # ========================================================

    row_week = {}

    loom_summary = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        capacity = settings["capacity"]

        current_week = settings["starting_week"]

        current_week_load = 0

        rows_processed = 0

        weeks_used = []


        # ----------------------------------------------------
        # Process rows in original Excel order
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


            if cleaned_excel_loom != loom:
                continue


            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


            if quantity is None:
                continue


            try:

                quantity = float(quantity)

            except (ValueError, TypeError):

                continue


            if quantity <= 0:
                continue


            # =================================================
            # WEEKLY CAPACITY LOGIC
            # =================================================

            if (
                current_week_load + quantity
                <= capacity
            ):

                # Product fits in current week.

                assigned_week = current_week

                current_week_load += quantity


            else:

                # Product does not fit.
                #
                # Move the WHOLE product to the next week.
                #
                # The product is NEVER split.

                current_week = next_week(
                    current_week
                )

                assigned_week = current_week

                current_week_load = quantity


            row_week[excel_row] = assigned_week

            rows_processed += 1

            if assigned_week not in weeks_used:

                weeks_used.append(
                    assigned_week
                )


        loom_summary[loom] = {

            "capacity": capacity,

            "starting_week":
                settings["starting_week"],

            "final_week":
                current_week,

            "rows_processed":
                rows_processed,

            "weeks_used":
                weeks_used
        }


    # ========================================================
    # SECOND PASS
    #
    # Assign dates.
    #
    # IMPORTANT:
    #
    # - No daily capacity.
    # - No product splitting.
    # - Every product gets ONE date.
    # - Try to spread products across the 7 days.
    # ========================================================

    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        starting_week = settings["starting_week"]


        # ----------------------------------------------------
        # Group rows belonging to this loom by week
        # ----------------------------------------------------

        week_rows = {}


        for excel_row, assigned_week in row_week.items():

            excel_loom = worksheet.cell(
                row=excel_row,
                column=LOOM_COL
            ).value


            if clean_loom_name(excel_loom) != loom:
                continue


            if assigned_week not in week_rows:

                week_rows[assigned_week] = []


            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


            week_rows[assigned_week].append(
                (
                    excel_row,
                    float(quantity)
                )
            )


        # ----------------------------------------------------
        # Process each production week
        # ----------------------------------------------------

        for production_week, rows in week_rows.items():

            week_start_date = get_week_start_date(
                planning_start_date,
                starting_week,
                production_week
            )


            # =================================================
            # BALANCED DATE DISTRIBUTION
            # =================================================
            #
            # There is NO daily capacity.
            #
            # We simply try to distribute complete products
            # across the 7 days as evenly as possible.
            #
            # A product always stays on ONE date.
            # =================================================

            daily_totals = {

                day_number: 0.0

                for day_number in range(7)
            }


            for excel_row, quantity in rows:

                # ------------------------------------------------
                # Find the day with the lowest current workload.
                # ------------------------------------------------

                selected_day = min(
                    daily_totals,
                    key=daily_totals.get
                )


                assigned_date = (
                    week_start_date
                    + timedelta(
                        days=selected_day
                    )
                )


                # ------------------------------------------------
                # Add the COMPLETE product to that date.
                # ------------------------------------------------

                daily_totals[selected_day] += quantity


                # ------------------------------------------------
                # Write Production Week
                # ------------------------------------------------

                worksheet.cell(
                    row=excel_row,
                    column=production_week_col
                ).value = (
                    f"WK-{production_week}"
                )


                # ------------------------------------------------
                # Write Production Date
                # ------------------------------------------------

                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).value = assigned_date


                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).number_format = "DD-MM-YYYY"


                # ------------------------------------------------
                # Write Day
                # ------------------------------------------------

                worksheet.cell(
                    row=excel_row,
                    column=day_col
                ).value = (
                    assigned_date.strftime("%A")
                )


    # ========================================================
    # SAVE OUTPUT
    # ========================================================

    output = BytesIO()

    workbook.save(output)

    output.seek(0)


    # ========================================================
    # SUCCESS
    # ========================================================

    st.success(
        "✅ Production week and date planning completed!"
    )


    # ========================================================
    # SUMMARY
    # ========================================================

    st.subheader("📊 Planning Summary")


    summary_rows = []


    for loom, info in loom_summary.items():

        if info["rows_processed"] == 0:
            continue


        summary_rows.append({

            "Loom":
                loom,

            "Weekly Capacity":
                info["capacity"],

            "Starting Week":
                f"WK-{info['starting_week']}",

            "Final Week":
                f"WK-{info['final_week']}",

            "Weeks Used":
                len(info["weeks_used"]),

            "Rows Processed":
                info["rows_processed"]
        })


    if summary_rows:

        summary_df = pd.DataFrame(
            summary_rows
        )

        st.dataframe(
            summary_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.warning(
            "No enabled loom data was found."
        )


    # ========================================================
    # DOWNLOAD
    # ========================================================

    st.subheader("4️⃣ Download Result")

    st.download_button(

        label="📥 Download Planned Excel",

        data=output,

        file_name=(
            "Loom_wise_Production_Planning_Automated.xlsx"
        ),

        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),

        use_container_width=True
    )
