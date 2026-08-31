import streamlit as st
import pandas as pd

from openpyxl import load_workbook
from io import BytesIO
from datetime import date, timedelta


# ============================================================
# PAGE CONFIG
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
    {"Use": True, "Loom": "ALT1", "Weekly Capacity": 1500},
    {"Use": True, "Loom": "ALT2", "Weekly Capacity": 1200},
    {"Use": True, "Loom": "ALT4", "Weekly Capacity": 800},
    {"Use": True, "Loom": "ALT5", "Weekly Capacity": 1300},
    {"Use": True, "Loom": "ALT6", "Weekly Capacity": 800},

    {"Use": True, "Loom": "RP1", "Weekly Capacity": 500},
    {"Use": True, "Loom": "RP2", "Weekly Capacity": 700},
    {"Use": True, "Loom": "RP3", "Weekly Capacity": 900},
    {"Use": True, "Loom": "RP4", "Weekly Capacity": 900},
    {"Use": True, "Loom": "RP5", "Weekly Capacity": 1100},
    {"Use": True, "Loom": "RP6", "Weekly Capacity": 1200},
    {"Use": True, "Loom": "RP7", "Weekly Capacity": 1000},

    {"Use": True, "Loom": "DB4M1", "Weekly Capacity": 300},
])


if "loom_settings" not in st.session_state:
    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()


# ============================================================
# BASIC HELPERS
# ============================================================

def clean_loom(value):

    if value is None:
        return ""

    return (
        str(value)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


def get_quantity(value):

    if value is None:
        return None

    try:
        quantity = float(value)
    except (ValueError, TypeError):
        return None

    if quantity <= 0:
        return None

    return quantity


# ============================================================
# COMPANY WEEK CALENDAR
# ============================================================
#
# IMPORTANT:
#
# This is NOT ISO week logic.
#
# Company's calendar:
#
# Week 1:
#     January 1 until the first Sunday
#
# Week 2:
#     First Monday until Sunday
#
# Week 3:
#     Following Monday until Sunday
#
# ...
#
# This produces:
#
# 2026:
#
# WK-32 -> 03-Aug to 09-Aug
# WK-33 -> 10-Aug to 16-Aug
# WK-34 -> 17-Aug to 23-Aug
# WK-35 -> 24-Aug to 30-Aug
# WK-36 -> 31-Aug to 06-Sep
#
# No date is hardcoded.
#
# ============================================================


def get_first_monday(year):

    """
    Find the first Monday on or after January 1.
    """

    jan_1 = date(
        year,
        1,
        1
    )

    days_until_monday = (
        7 - jan_1.weekday()
    ) % 7

    return (
        jan_1
        + timedelta(
            days=days_until_monday
        )
    )


def get_current_production_week(
    target_date
):

    """
    Determine the company's production week
    from an actual calendar date.

    Week 1 starts on January 1.

    After the first Sunday,
    weeks run Monday-Sunday.
    """

    year = target_date.year

    jan_1 = date(
        year,
        1,
        1
    )

    first_monday = get_first_monday(
        year
    )


    # --------------------------------------------------------
    # January 1 until the first Sunday = Week 1
    # --------------------------------------------------------

    if target_date < first_monday:

        return 1


    # --------------------------------------------------------
    # Every following Monday-Sunday = next week
    # --------------------------------------------------------

    days_difference = (
        target_date
        - first_monday
    ).days


    calculated_week = (
        2
        + days_difference // 7
    )


    # --------------------------------------------------------
    # Company has Week 1-52.
    #
    # The remaining end-of-year dates stay in Week 52.
    # --------------------------------------------------------

    return min(
        calculated_week,
        52
    )


def get_week_start_date(
    production_week,
    year
):

    """
    Convert company production week
    into its Monday/start date.

    Week 1 starts on January 1.

    For Week 2 onward, weeks start Monday.
    """

    production_week = int(
        production_week
    )


    if production_week <= 1:

        return date(
            year,
            1,
            1
        )


    first_monday = get_first_monday(
        year
    )


    return (
        first_monday
        + timedelta(
            days=(production_week - 2) * 7
        )
    )


def get_week_end_date(
    production_week,
    year
):

    start = get_week_start_date(
        production_week,
        year
    )


    # Week 1 ends on the first Sunday.
    if production_week == 1:

        days_until_sunday = (
            6 - start.weekday()
        ) % 7

        return (
            start
            + timedelta(
                days=days_until_sunday
            )
        )


    return (
        start
        + timedelta(
            days=6
        )
    )


# ============================================================
# MOVE TO NEXT PRODUCTION WEEK
# ============================================================

def move_to_next_week(
    current_week,
    current_year
):

    """
    Move from one production week to the next.

    WK-52 -> WK-1 of next year.
    """

    if current_week >= 52:

        return 1, current_year + 1

    return (
        current_week + 1,
        current_year
    )


# ============================================================
# BALANCED DAILY DISTRIBUTION
# ============================================================

def distribute_products_across_week(
    rows,
    week_start_date,
    week_end_date
):

    """
    Distribute complete products across the available
    production days.

    Rules:

    - No daily capacity.
    - Product quantity is NEVER split.
    - Excel order is preserved.
    - External Document No. does not affect order.
    - Try to balance total quantity across the week.
    """

    if not rows:
        return {}


    # --------------------------------------------------------
    # Determine number of working days.
    #
    # Currently all seven days are working.
    # Sunday exceptions can be added later.
    # --------------------------------------------------------

    number_of_days = (
        (week_end_date - week_start_date).days
        + 1
    )


    # --------------------------------------------------------
    # Safety.
    # --------------------------------------------------------

    if number_of_days <= 0:

        return {}


    number_of_rows = len(
        rows
    )


    # --------------------------------------------------------
    # If number of products <= number of days,
    # assign sequentially.
    # --------------------------------------------------------

    if number_of_rows <= number_of_days:

        result = {}


        for index, (
            excel_row,
            quantity
        ) in enumerate(rows):

            assigned_date = (
                week_start_date
                + timedelta(
                    days=index
                )
            )


            result[
                excel_row
            ] = assigned_date


        return result


    # ========================================================
    # Total quantity
    # ========================================================

    total_quantity = sum(
        quantity
        for _, quantity in rows
    )


    target = (
        total_quantity
        / number_of_days
    )


    # ========================================================
    # Prefix sums
    # ========================================================

    prefix = [0.0]


    for _, quantity in rows:

        prefix.append(
            prefix[-1] + quantity
        )


    # ========================================================
    # Dynamic programming
    #
    # Divide products into consecutive groups.
    #
    # No product is split.
    # ========================================================

    INF = float("inf")


    dp = [

        [INF] * (
            number_of_rows + 1
        )

        for _ in range(
            number_of_days + 1
        )
    ]


    parent = [

        [None] * (
            number_of_rows + 1
        )

        for _ in range(
            number_of_days + 1
        )
    ]


    dp[0][0] = 0.0


    for day in range(
        1,
        number_of_days + 1
    ):

        for end in range(
            day,
            number_of_rows + 1
        ):

            for start in range(
                day - 1,
                end
            ):

                previous_cost = dp[
                    day - 1
                ][start]


                if previous_cost == INF:

                    continue


                day_quantity = (
                    prefix[end]
                    - prefix[start]
                )


                difference = (
                    day_quantity
                    - target
                )


                cost = (
                    difference
                    * difference
                )


                total_cost = (
                    previous_cost
                    + cost
                )


                if total_cost < dp[
                    day
                ][end]:

                    dp[
                        day
                    ][end] = total_cost


                    parent[
                        day
                    ][end] = start


    # ========================================================
    # Recover groups
    # ========================================================

    boundaries = []

    end = number_of_rows


    for day in range(
        number_of_days,
        0,
        -1
    ):

        start = parent[
            day
        ][end]


        if start is None:

            break


        boundaries.append(
            (
                start,
                end
            )
        )


        end = start


    boundaries.reverse()


    # ========================================================
    # Assign dates
    # ========================================================

    result = {}


    for day_index, (
        start,
        end
    ) in enumerate(
        boundaries
    ):

        assigned_date = (
            week_start_date
            + timedelta(
                days=day_index
            )
        )


        for index in range(
            start,
            end
        ):

            excel_row = rows[
                index
            ][0]


            result[
                excel_row
            ] = assigned_date


    return result


# ============================================================
# TITLE
# ============================================================

st.title(
    "🏭 Loom Production Planner"
)

st.write(
    "Automatically plan production using the current "
    "production week, loom capacity and Excel row order."
)


# ============================================================
# CURRENT DATE / WEEK
# ============================================================

today = date.today()


current_week = (
    get_current_production_week(
        today
    )
)


current_week_start = (
    get_week_start_date(
        current_week,
        today.year
    )
)


current_week_end = (
    get_week_end_date(
        current_week,
        today.year
    )
)


st.info(
    f"📅 Today: "
    f"{today.strftime('%d-%m-%Y')}"
    f"   |   "
    f"Production Week: "
    f"WK-{current_week}"
    f"   |   "
    f"Week: "
    f"{current_week_start.strftime('%d-%m-%Y')}"
    f" → "
    f"{current_week_end.strftime('%d-%m-%Y')}"
)


# ============================================================
# FILE UPLOAD
# ============================================================

st.subheader(
    "1️⃣ Upload Production Excel"
)


uploaded_file = st.file_uploader(
    "Select Excel file",
    type=["xlsx"]
)


# ============================================================
# LOOM SETTINGS
# ============================================================

st.subheader(
    "2️⃣ Loom Settings"
)


st.caption(
    "Starting production week is automatically determined "
    "from today's date."
)


edited_settings = st.data_editor(

    st.session_state.loom_settings,

    use_container_width=True,

    hide_index=True,

    column_config={

        "Use":
            st.column_config.CheckboxColumn(
                "Use"
            ),

        "Loom":
            st.column_config.TextColumn(
                "Loom",
                disabled=True
            ),

        "Weekly Capacity":
            st.column_config.NumberColumn(
                "Weekly Capacity",
                min_value=1,
                step=50
            )
    },

    key="loom_settings_editor"
)


# ============================================================
# RESET
# ============================================================

if st.button(
    "🔄 Reset Loom Settings"
):

    st.session_state.loom_settings = (
        DEFAULT_SETTINGS.copy()
    )

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
    # FILE CHECK
    # ========================================================

    if uploaded_file is None:

        st.error(
            "Please upload the Excel file first."
        )

        st.stop()


    # ========================================================
    # CURRENT PRODUCTION WEEK
    # ========================================================

    planning_today = date.today()

    starting_week = (
        get_current_production_week(
            planning_today
        )
    )

    starting_year = (
        planning_today.year
    )


    # ========================================================
    # READ LOOM SETTINGS
    # ========================================================

    loom_settings = {}


    for _, row in edited_settings.iterrows():

        loom = clean_loom(
            row["Loom"]
        )


        use_loom = bool(
            row["Use"]
        )


        weekly_capacity = float(
            row["Weekly Capacity"]
        )


        if weekly_capacity <= 0:

            st.error(
                f"{loom}: Weekly Capacity must be "
                "greater than 0."
            )

            st.stop()


        loom_settings[
            loom
        ] = {

            "use":
                use_loom,

            "capacity":
                weekly_capacity
        }


    # ========================================================
    # LOAD WORKBOOK
    # ========================================================

    try:

        workbook = load_workbook(
            filename=BytesIO(
                uploaded_file.getvalue()
            )
        )

    except Exception as e:

        st.error(
            f"Could not open Excel file: {e}"
        )

        st.stop()


    if not workbook.sheetnames:

        st.error(
            "No worksheets found in the Excel file."
        )

        st.stop()


    worksheet = workbook[
        workbook.sheetnames[0]
    ]


    # ========================================================
    # COLUMN STRUCTURE
    # ========================================================
    #
    # A = No.
    # B = External Document No.
    # C = Description
    # D = Roll Width
    # E = Roll Length
    # F = Quantity
    # G = Loom
    #
    # H = Production Week
    # I = Production Date
    # J = Day
    #
    # ========================================================

    QUANTITY_COL = 6
    LOOM_COL = 7

    WEEK_COL = 8
    DATE_COL = 9
    DAY_COL = 10


    # ========================================================
    # HEADERS
    # ========================================================

    worksheet.cell(
        row=1,
        column=WEEK_COL
    ).value = "Production Week"


    worksheet.cell(
        row=1,
        column=DATE_COL
    ).value = "Production Date"


    worksheet.cell(
        row=1,
        column=DAY_COL
    ).value = "Day"


    # ========================================================
    # CLEAR PREVIOUS OUTPUT
    # ========================================================

    for row in range(
        2,
        worksheet.max_row + 1
    ):

        worksheet.cell(
            row=row,
            column=WEEK_COL
        ).value = None


        worksheet.cell(
            row=row,
            column=DATE_COL
        ).value = None


        worksheet.cell(
            row=row,
            column=DAY_COL
        ).value = None


    # ========================================================
    # READ LOOM ROWS
    # ========================================================
    #
    # VERY IMPORTANT:
    #
    # We read Excel from top to bottom.
    #
    # There is NO sorting.
    #
    # External Document No. is NOT used.
    #
    # ========================================================

    loom_rows = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        loom_rows[
            loom
        ] = []


    for row in range(
        2,
        worksheet.max_row + 1
    ):

        excel_loom = clean_loom(
            worksheet.cell(
                row=row,
                column=LOOM_COL
            ).value
        )


        if excel_loom not in loom_rows:

            continue


        quantity = get_quantity(
            worksheet.cell(
                row=row,
                column=QUANTITY_COL
            ).value
        )


        if quantity is None:

            continue


        # ----------------------------------------------------
        # EXACT EXCEL ORDER.
        # ----------------------------------------------------

        loom_rows[
            excel_loom
        ].append(
            (
                row,
                quantity
            )
        )


    # ========================================================
    # ASSIGN PRODUCTION WEEKS
    # ========================================================

    row_week = {}

    row_week_year = {}


    loom_summary = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:

            continue


        weekly_capacity = (
            settings["capacity"]
        )


        # ----------------------------------------------------
        # EVERY LOOM STARTS FROM TODAY'S WEEK.
        # ----------------------------------------------------

        current_week = (
            starting_week
        )


        current_year = (
            starting_year
        )


        current_load = 0.0

        rows_processed = 0

        weeks_used = []


        # ----------------------------------------------------
        # PROCESS IN EXACT EXCEL ORDER.
        # ----------------------------------------------------

        for excel_row, quantity in loom_rows.get(
            loom,
            []
        ):


            # ------------------------------------------------
            # Product cannot be split.
            #
            # If it doesn't fit, move the ENTIRE product
            # to the next week.
            # ------------------------------------------------

            if (

                current_load > 0

                and

                current_load
                + quantity
                > weekly_capacity

            ):

                (
                    current_week,
                    current_year
                ) = move_to_next_week(
                    current_week,
                    current_year
                )


                current_load = 0.0


            # ------------------------------------------------
            # Assign production week.
            # ------------------------------------------------

            row_week[
                excel_row
            ] = current_week


            row_week_year[
                excel_row
            ] = current_year


            # ------------------------------------------------
            # Update weekly capacity.
            # ------------------------------------------------

            current_load += quantity


            rows_processed += 1


            week_key = (
                current_year,
                current_week
            )


            if week_key not in weeks_used:

                weeks_used.append(
                    week_key
                )


        # ----------------------------------------------------
        # Summary.
        # ----------------------------------------------------

        loom_summary[
            loom
        ] = {

            "capacity":
                weekly_capacity,

            "starting_week":
                starting_week,

            "starting_year":
                starting_year,

            "final_week":
                current_week,

            "final_year":
                current_year,

            "weeks_used":
                weeks_used,

            "rows_processed":
                rows_processed
        }


    # ========================================================
    # ASSIGN DATES
    # ========================================================

    row_dates = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:

            continue


        # ----------------------------------------------------
        # Group rows by YEAR + WEEK.
        #
        # Original Excel order remains intact.
        # ----------------------------------------------------

        week_rows = {}


        for excel_row, week in row_week.items():

            excel_loom = clean_loom(
                worksheet.cell(
                    row=excel_row,
                    column=LOOM_COL
                ).value
            )


            if excel_loom != loom:

                continue


            quantity = get_quantity(
                worksheet.cell(
                    row=excel_row,
                    column=QUANTITY_COL
                ).value
            )


            if quantity is None:

                continue


            year = row_week_year[
                excel_row
            ]


            week_key = (
                year,
                week
            )


            if week_key not in week_rows:

                week_rows[
                    week_key
                ] = []


            # ------------------------------------------------
            # NO SORTING.
            # ------------------------------------------------

            week_rows[
                week_key
            ].append(
                (
                    excel_row,
                    quantity
                )
            )


        # ====================================================
        # PROCESS EACH WEEK
        # ====================================================

        for (
            week_key,
            rows
        ) in week_rows.items():

            year, production_week = (
                week_key
            )


            # ------------------------------------------------
            # Calculate actual calendar dates.
            # ------------------------------------------------

            week_start = (
                get_week_start_date(
                    production_week,
                    year
                )
            )


            week_end = (
                get_week_end_date(
                    production_week,
                    year
                )
            )


            # ------------------------------------------------
            # Distribute products.
            # ------------------------------------------------

            assignments = (
                distribute_products_across_week(
                    rows,
                    week_start,
                    week_end
                )
            )


            for (
                excel_row,
                assigned_date
            ) in assignments.items():

                row_dates[
                    excel_row
                ] = assigned_date


    # ========================================================
    # WRITE RESULTS
    # ========================================================

    for excel_row, production_week in (
        row_week.items()
    ):

        if excel_row not in row_dates:

            continue


        production_date = (
            row_dates[
                excel_row
            ]
        )


        # ----------------------------------------------------
        # Production Week
        # ----------------------------------------------------

        worksheet.cell(
            row=excel_row,
            column=WEEK_COL
        ).value = (
            f"WK-{production_week}"
        )


        # ----------------------------------------------------
        # Production Date
        # ----------------------------------------------------

        worksheet.cell(
            row=excel_row,
            column=DATE_COL
        ).value = (
            production_date
        )


        worksheet.cell(
            row=excel_row,
            column=DATE_COL
        ).number_format = (
            "DD-MM-YYYY"
        )


        # ----------------------------------------------------
        # Day
        # ----------------------------------------------------

        worksheet.cell(
            row=excel_row,
            column=DAY_COL
        ).value = (
            production_date.strftime(
                "%A"
            )
        )


    # ========================================================
    # IMPORTANT:
    #
    # THERE IS NO SORTING HERE.
    #
    # Excel order stays exactly the same.
    # ========================================================


    # ========================================================
    # SAVE
    # ========================================================

    output = BytesIO()


    workbook.save(
        output
    )


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

    st.subheader(
        "📊 Planning Summary"
    )


    summary_rows = []


    for loom, info in loom_summary.items():

        if info[
            "rows_processed"
        ] <= 0:

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
                len(
                    info["weeks_used"]
                ),

            "Rows Processed":
                info["rows_processed"]
        })


    if summary_rows:

        st.dataframe(

            pd.DataFrame(
                summary_rows
            ),

            use_container_width=True,

            hide_index=True
        )

    else:

        st.warning(
            "No production rows were found for "
            "the selected/enabled looms."
        )


    # ========================================================
    # DOWNLOAD
    # ========================================================

    st.subheader(
        "4️⃣ Download Result"
    )


    st.download_button(

        label=(
            "📥 Download Planned Excel"
        ),

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
