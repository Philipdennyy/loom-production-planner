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
# GET DATE FOR PRODUCTION WEEK
# ============================================================

def get_week_start_date(
    production_week,
    starting_week,
    planning_start_date
):
    """
    Finds the starting date of a production week.

    Example:
    Planning start date = 31-Aug-2026
    Starting week = 33

    WK-33 = 31-Aug-2026
    WK-34 = 07-Sep-2026
    WK-35 = 14-Sep-2026

    Handles:
    WK-52 -> WK-1
    """

    week_difference = (
        production_week - starting_week
    ) % 52

    return (
        planning_start_date
        + timedelta(days=week_difference * 7)
    )


# ============================================================
# TITLE
# ============================================================

st.title("🏭 Loom Production Planner")

st.write(
    "Automatically assign production weeks and production dates "
    "based on loom capacity and quantity."
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
# PLANNING DATE
# ============================================================

st.subheader("📅 Planning Start Date")

planning_start_date = st.date_input(
    "Select the date from which the production planning should start",
    value=date.today()
)

st.caption(
    "This date represents the beginning of the selected starting week. "
    "Each following production week covers 7 days."
)


# ============================================================
# STEP 2 — LOOM SETTINGS
# ============================================================

st.subheader("2️⃣ Loom Settings")

st.caption(
    "Edit the weekly capacity and starting week. "
    "Uncheck a loom if it should not be included."
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
            help="Maximum production quantity per week.",
            min_value=1,
            step=50
        ),

        "Starting Week": st.column_config.NumberColumn(
            "Starting Week",
            help="Production week from which this loom should start.",
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
    # Check upload
    # --------------------------------------------------------

    if uploaded_file is None:

        st.error(
            "Please upload the Excel file first."
        )

        st.stop()


    # ========================================================
    # CREATE LOOM SETTINGS
    # ========================================================

    loom_settings = {}

    for _, row in edited_settings.iterrows():

        loom = clean_loom_name(row["Loom"])

        use_loom = bool(row["Use"])

        capacity = float(row["Weekly Capacity"])

        starting_week = int(row["Starting Week"])


        # ----------------------------------------------------
        # Validation
        # ----------------------------------------------------

        if capacity <= 0:

            st.error(
                f"Capacity for {loom} must be greater than 0."
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
    # USE FIRST SHEET
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
    # FIXED INPUT COLUMNS
    # ========================================================

    # F = Quantity
    # G = Loom
    #
    # H = Production Week
    # I = Production Date
    #
    # If H/I are occupied by other data, new columns
    # will be inserted safely.

    QUANTITY_COL = 6
    LOOM_COL = 7

    PRODUCTION_WEEK_COL = 8
    PRODUCTION_DATE_COL = 9


    # ========================================================
    # CHECK COLUMN H
    # ========================================================

    h_header = worksheet.cell(
        row=1,
        column=PRODUCTION_WEEK_COL
    ).value


    # --------------------------------------------------------
    # CASE 1:
    # H already contains Production Week
    # --------------------------------------------------------

    if (
        h_header is not None
        and
        str(h_header).strip().lower()
        == "production week"
    ):

        production_week_col = 8


    # --------------------------------------------------------
    # CASE 2:
    # H is completely empty
    # --------------------------------------------------------

    elif all(
        worksheet.cell(
            row=row,
            column=PRODUCTION_WEEK_COL
        ).value is None

        for row in range(
            1,
            worksheet.max_row + 1
        )
    ):

        production_week_col = 8

        worksheet.cell(
            row=1,
            column=production_week_col
        ).value = "Production Week"


    # --------------------------------------------------------
    # CASE 3:
    # H contains other information
    # --------------------------------------------------------

    else:

        worksheet.insert_cols(8)

        production_week_col = 8

        worksheet.cell(
            row=1,
            column=production_week_col
        ).value = "Production Week"


    # ========================================================
    # PRODUCTION DATE COLUMN
    # ========================================================

    production_date_col = production_week_col + 1


    i_header = worksheet.cell(
        row=1,
        column=production_date_col
    ).value


    # --------------------------------------------------------
    # If I already contains Production Date
    # --------------------------------------------------------

    if (
        i_header is not None
        and
        str(i_header).strip().lower()
        == "production date"
    ):

        pass


    # --------------------------------------------------------
    # If I contains something else
    # insert a new column
    # --------------------------------------------------------

    elif any(
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


    # Set header

    worksheet.cell(
        row=1,
        column=production_date_col
    ).value = "Production Date"


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


    # ========================================================
    # PROCESS EACH LOOM
    # ========================================================

    summary = []


    for loom, settings in loom_settings.items():

        # ----------------------------------------------------
        # Skip unchecked loom
        # ----------------------------------------------------

        if not settings["use"]:

            continue


        capacity = settings["capacity"]

        starting_week = settings["starting_week"]

        current_week = starting_week

        current_load = 0

        rows_processed = 0

        weeks_used = set()

        dates_used = set()


        # ====================================================
        # DAILY CAPACITY
        # ====================================================

        daily_capacity = capacity / 7

        current_day_load = 0

        current_date = planning_start_date


        # ====================================================
        # PROCESS EXCEL ROWS
        # ====================================================

        for excel_row in range(
            2,
            worksheet.max_row + 1
        ):


            # ------------------------------------------------
            # Read loom
            # ------------------------------------------------

            excel_loom = worksheet.cell(
                row=excel_row,
                column=LOOM_COL
            ).value


            cleaned_excel_loom = clean_loom_name(
                excel_loom
            )


            # ------------------------------------------------
            # Only process matching loom
            # ------------------------------------------------

            if cleaned_excel_loom != loom:

                continue


            # ------------------------------------------------
            # Read quantity
            # ------------------------------------------------

            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


            if quantity is None:

                continue


            # ------------------------------------------------
            # Convert quantity
            # ------------------------------------------------

            try:

                quantity = float(quantity)

            except (ValueError, TypeError):

                continue


            # =================================================
            # WEEKLY ALLOCATION
            # =================================================

            if current_load + quantity <= capacity:

                # Order fits current week

                assigned_week = current_week

                current_load += quantity


            else:

                # Move entire order to next week

                current_week += 1


                # WK-52 → WK-1

                if current_week > 52:

                    current_week = 1


                assigned_week = current_week

                current_load = quantity


            # =================================================
            # GET WEEK START DATE
            # =================================================

            week_start_date = get_week_start_date(
                assigned_week,
                starting_week,
                planning_start_date
            )


            # =================================================
            # DAILY DATE ALLOCATION
            # =================================================

            # If the week changed, start from the first
            # available day of that week.

            if assigned_week != starting_week:

                pass


            # -------------------------------------------------
            # Determine how much of this order goes onto
            # each day.
            # -------------------------------------------------

            remaining_quantity = quantity


            # -------------------------------------------------
            # Find the first date for this order.
            #
            # We maintain daily loads separately for each
            # production week.
            # -------------------------------------------------

            if (
                not hasattr(
                    st.session_state,
                    "_daily_loads"
                )
            ):

                st.session_state._daily_loads = {}


            # Key = (loom, production_week, date)

            # Find first available day in this week

            assigned_date = None


            for day_number in range(7):

                possible_date = (
                    week_start_date
                    + timedelta(days=day_number)
                )


                key = (
                    loom,
                    assigned_week,
                    possible_date
                )


                if key not in st.session_state._daily_loads:

                    st.session_state._daily_loads[key] = 0


                available_capacity = (
                    daily_capacity
                    - st.session_state._daily_loads[key]
                )


                if available_capacity > 0:

                    assigned_date = possible_date

                    quantity_for_day = min(
                        remaining_quantity,
                        available_capacity
                    )


                    st.session_state._daily_loads[key] += (
                        quantity_for_day
                    )


                    remaining_quantity -= (
                        quantity_for_day
                    )


                    dates_used.add(
                        possible_date
                    )


                    if remaining_quantity <= 0:

                        break


            # -------------------------------------------------
            # If quantity is larger than the capacity of
            # one week, continue into subsequent weeks.
            # -------------------------------------------------

            while remaining_quantity > 0:

                current_week += 1

                if current_week > 52:

                    current_week = 1


                next_week_start = get_week_start_date(
                    current_week,
                    starting_week,
                    planning_start_date
                )


                for day_number in range(7):

                    possible_date = (
                        next_week_start
                        + timedelta(days=day_number)
                    )


                    key = (
                        loom,
                        current_week,
                        possible_date
                    )


                    if key not in st.session_state._daily_loads:

                        st.session_state._daily_loads[key] = 0


                    available_capacity = (
                        daily_capacity
                        - st.session_state._daily_loads[key]
                    )


                    if available_capacity <= 0:

                        continue


                    quantity_for_day = min(
                        remaining_quantity,
                        available_capacity
                    )


                    st.session_state._daily_loads[key] += (
                        quantity_for_day
                    )


                    remaining_quantity -= (
                        quantity_for_day
                    )


                    dates_used.add(
                        possible_date
                    )


                    if remaining_quantity <= 0:

                        break


            # =================================================
            # WRITE WEEK
            # =================================================

            worksheet.cell(
                row=excel_row,
                column=production_week_col
            ).value = f"WK-{assigned_week}"


            # =================================================
            # WRITE PRODUCTION DATE
            # =================================================

            if assigned_date is not None:

                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).value = assigned_date


                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).number_format = "DD-MM-YYYY"


            rows_processed += 1

            weeks_used.add(
                assigned_week
            )


        # ====================================================
        # SUMMARY
        # ====================================================

        if rows_processed > 0:

            summary.append({

                "Loom": loom,

                "Weekly Capacity": capacity,

                "Daily Capacity":
                    round(daily_capacity, 2),

                "Starting Week":
                    f"WK-{starting_week}",

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
    # SUCCESS
    # ========================================================

    st.success(
        "✅ Production planning completed successfully!"
    )


    # ========================================================
    # SUMMARY
    # ========================================================

    st.subheader("📊 Planning Summary")


    if summary:

        summary_df = pd.DataFrame(
            summary
        )

        st.dataframe(
            summary_df,
            use_container_width=True,
            hide_index=True
        )

    else:

        st.info(
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
