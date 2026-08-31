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
    """Standardize loom names for comparison."""

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
    """Move to next production week. WK-52 -> WK-1."""

    if week >= 52:
        return 1

    return week + 1


def week_difference(start_week, target_week):
    """
    Calculate how many weeks target_week is after start_week.

    Handles WK-52 -> WK-1.

    Example:
    start = 33
    target = 34
    result = 1

    start = 52
    target = 1
    result = 1
    """

    return (target_week - start_week) % 52


def get_week_start_date(
    planning_start_date,
    starting_week,
    production_week
):
    """
    Get the date on which a production week begins.

    Example:

    Start date = 31-Aug-2026
    Starting week = 33

    WK-33 = 31-Aug-2026
    WK-34 = 07-Sep-2026
    WK-35 = 14-Sep-2026
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
    "Automate loom-wise production week and date planning "
    "based on quantity and loom capacity."
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
# STEP 2 — STARTING DATE
# ============================================================

st.subheader("2️⃣ Planning Start Date")

planning_start_date = st.date_input(
    "Select the starting date",
    value=date.today()
)

st.info(
    f"Planning will begin from "
    f"{planning_start_date.strftime('%d-%m-%Y')} "
    f"and each production week will cover 7 days."
)


# ============================================================
# STEP 3 — LOOM SETTINGS
# ============================================================

st.subheader("3️⃣ Loom Settings")

st.caption(
    "Select the looms to use and edit their weekly capacity "
    "and starting production week."
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
            help="Maximum production quantity for this loom per week.",
            min_value=1,
            step=50
        ),

        "Starting Week": st.column_config.NumberColumn(
            "Starting Week",
            help="Week from which this loom starts production.",
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
# STEP 4 — PRODUCTION CALENDAR PREVIEW
# ============================================================

st.divider()

st.subheader("📅 Production Calendar Preview")

calendar_rows = []

for _, row in edited_settings.iterrows():

    if not bool(row["Use"]):
        continue

    loom = clean_loom_name(row["Loom"])

    starting_week = int(row["Starting Week"])

    # Show 10 weeks as preview
    current_week = starting_week

    for _ in range(10):

        week_start = get_week_start_date(
            planning_start_date,
            starting_week,
            current_week
        )

        for day_number in range(7):

            current_date = (
                week_start
                + timedelta(days=day_number)
            )

            calendar_rows.append({

                "Loom": loom,

                "Production Week":
                    f"WK-{current_week}",

                "Date":
                    current_date.strftime("%d-%m-%Y"),

                "Day":
                    current_date.strftime("%A")
            })

        current_week = next_week(current_week)


calendar_df = pd.DataFrame(calendar_rows)


if not calendar_df.empty:

    st.dataframe(
        calendar_df,
        use_container_width=True,
        hide_index=True
    )

    st.caption(
        "Calendar preview shows the first 10 production weeks "
        "for the selected looms."
    )

else:

    st.info(
        "Select at least one loom to display the calendar."
    )


# ============================================================
# RUN BUTTON
# ============================================================

st.divider()

st.subheader("5️⃣ Run Production Planning")

run_button = st.button(
    "🚀 Run Production Planning",
    type="primary",
    use_container_width=True
)


# ============================================================
# MAIN PROCESSING
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
    # BUILD LOOM CONFIGURATION
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
    # LOAD WORKBOOK
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
    # INPUT COLUMNS
    # ========================================================

    # F = Quantity
    # G = Loom
    #
    # We will create:
    #
    # H = Production Week
    # I = Production Date
    # J = Day


    QUANTITY_COL = 6
    LOOM_COL = 7

    PRODUCTION_WEEK_COL = 8
    PRODUCTION_DATE_COL = 9
    DAY_COL = 10


    # ========================================================
    # SAFELY CREATE OUTPUT COLUMNS
    # ========================================================

    # --------------------------------------------------------
    # Production Week
    # --------------------------------------------------------

    h_header = worksheet.cell(
        row=1,
        column=PRODUCTION_WEEK_COL
    ).value


    if (
        h_header is not None
        and
        str(h_header).strip().lower()
        == "production week"
    ):

        production_week_col = 8

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

    else:

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

            worksheet.insert_cols(day_col
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
    # DAILY ALLOCATION STORAGE
    # ========================================================

    # Key:
    #
    # (loom, production_week, date)
    #
    # Value:
    #
    # quantity already allocated on that date


    daily_loads = {}


    # ========================================================
    # SUMMARY
    # ========================================================

    summary = []


    # ========================================================
    # PROCESS EACH LOOM INDEPENDENTLY
    # ========================================================

    for loom, settings in loom_settings.items():

        # ----------------------------------------------------
        # Skip disabled loom
        # ----------------------------------------------------

        if not settings["use"]:
            continue


        capacity = settings["capacity"]

        starting_week = settings["starting_week"]

        current_week = starting_week

        current_week_load = 0

        rows_processed = 0

        weeks_used = set()

        dates_used = set()


        # ----------------------------------------------------
        # Daily capacity
        # ----------------------------------------------------

        daily_capacity = capacity / 7


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


            try:

                quantity = float(quantity)

            except (ValueError, TypeError):

                continue


            if quantity <= 0:
                continue


            # =================================================
            # WEEK ALLOCATION
            # =================================================

            if (
                current_week_load + quantity
                <= capacity
            ):

                assigned_week = current_week

                current_week_load += quantity


            else:

                # Move entire order to next week

                current_week = next_week(
                    current_week
                )

                assigned_week = current_week

                current_week_load = quantity


            # =================================================
            # FIND WEEK START DATE
            # =================================================

            week_start_date = get_week_start_date(
                planning_start_date,
                starting_week,
                assigned_week
            )


            # =================================================
            # DAILY ALLOCATION
            # =================================================

            remaining_quantity = quantity

            assigned_first_date = None


            # -------------------------------------------------
            # Search the 7 days of the assigned week
            # -------------------------------------------------

            for day_number in range(7):

                current_date = (
                    week_start_date
                    + timedelta(days=day_number)
                )


                key = (
                    loom,
                    assigned_week,
                    current_date
                )


                if key not in daily_loads:

                    daily_loads[key] = 0


                available_capacity = (
                    daily_capacity
                    - daily_loads[key]
                )


                if available_capacity <= 0:

                    continue


                # Quantity assigned to this date

                quantity_for_day = min(
                    remaining_quantity,
                    available_capacity
                )


                daily_loads[key] += (
                    quantity_for_day
                )


                remaining_quantity -= (
                    quantity_for_day
                )


                if assigned_first_date is None:

                    assigned_first_date = current_date


                dates_used.add(
                    current_date
                )


                # ------------------------------------------------
                # If the complete order has been allocated
                # ------------------------------------------------

                if remaining_quantity <= 0:

                    break


            # =================================================
            # IF ORDER DOES NOT FIT IN THE WEEK
            # =================================================

            while remaining_quantity > 0:

                current_week = next_week(
                    current_week
                )


                new_week_start = get_week_start_date(
                    planning_start_date,
                    starting_week,
                    current_week
                )


                for day_number in range(7):

                    current_date = (
                        new_week_start
                        + timedelta(days=day_number)
                    )


                    key = (
                        loom,
                        current_week,
                        current_date
                    )


                    if key not in daily_loads:

                        daily_loads[key] = 0


                    available_capacity = (
                        daily_capacity
                        - daily_loads[key]
                    )


                    if available_capacity <= 0:

                        continue


                    quantity_for_day = min(
                        remaining_quantity,
                        available_capacity
                    )


                    daily_loads[key] += (
                        quantity_for_day
                    )


                    remaining_quantity -= (
                        quantity_for_day
                    )


                    dates_used.add(
                        current_date
                    )


                    if remaining_quantity <= 0:

                        break


            # =================================================
            # WRITE OUTPUT
            # =================================================

            worksheet.cell(
                row=excel_row,
                column=production_week_col
            ).value = (
                f"WK-{assigned_week}"
            )


            if assigned_first_date is not None:

                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).value = assigned_first_date


                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).number_format = "DD-MM-YYYY"


                worksheet.cell(
                    row=excel_row,
                    column=day_col
                ).value = (
                    assigned_first_date.strftime("%A")
                )


            rows_processed += 1

            weeks_used.add(
                assigned_week
            )


        # ====================================================
        # ADD SUMMARY
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

                "Dates Used":
                    len(dates_used),

                "Rows Processed":
                    rows_processed
            })


    # ========================================================
    # SAVE WORKBOOK
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

        st.warning(
            "No enabled loom data was found in the Excel file."
        )


    # ========================================================
    # DAILY CAPACITY VIEW
    # ========================================================

    st.subheader("📅 Daily Capacity")


    daily_summary = []


    for key, load in daily_loads.items():

        loom, production_week, production_date = key

        capacity = loom_settings[loom]["capacity"]

        daily_capacity = capacity / 7


        daily_summary.append({

            "Loom": loom,

            "Production Week":
                f"WK-{production_week}",

            "Date":
                production_date.strftime(
                    "%d-%m-%Y"
                ),

            "Day":
                production_date.strftime(
                    "%A"
                ),

            "Daily Capacity":
                round(daily_capacity, 2),

            "Planned Quantity":
                round(load, 2),

            "Remaining Capacity":
                round(
                    daily_capacity - load,
                    2
                )
        })


    if daily_summary:

        daily_summary_df = pd.DataFrame(
            daily_summary
        )

        st.dataframe(
            daily_summary_df,
            use_container_width=True,
            hide_index=True
        )


    # ========================================================
    # DOWNLOAD
    # ========================================================

    st.subheader("6️⃣ Download Result")

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
