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
# HELPER FUNCTIONS
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


def next_week(week):

    if week >= 52:
        return 1

    return week + 1


# ============================================================
# CURRENT PRODUCTION WEEK
# ============================================================

def get_current_production_week():

    today = date.today()

    return today.isocalendar().week


# ============================================================
# WEEK -> MONDAY
# ============================================================

def get_week_start_date(
    production_week,
    year
):

    return date.fromisocalendar(
        year,
        production_week,
        1
    )


# ============================================================
# DISTRIBUTE PRODUCTS ACROSS WEEK
# ============================================================

def distribute_products_across_week(
    rows,
    week_start_date
):

    """
    Distribute complete products across Monday-Sunday.

    Rules:

    - No daily capacity.
    - Product quantity is never split.
    - Excel order is preserved.
    - External Document No. does NOT affect ordering.
    - Try to balance the quantity across the seven days.
    """

    if not rows:
        return {}

    number_of_rows = len(rows)


    # --------------------------------------------------------
    # Seven or fewer products
    # --------------------------------------------------------

    if number_of_rows <= 7:

        result = {}

        for index, (
            excel_row,
            quantity
        ) in enumerate(rows):

            result[
                excel_row
            ] = (
                week_start_date
                + timedelta(days=index)
            )

        return result


    # ========================================================
    # Total quantity
    # ========================================================

    total_quantity = sum(
        quantity
        for _, quantity in rows
    )


    target = (
        total_quantity / 7
    )


    # ========================================================
    # Prefix sums
    # ========================================================

    prefix = [0]

    for _, quantity in rows:

        prefix.append(
            prefix[-1] + quantity
        )


    # ========================================================
    # Dynamic programming
    # ========================================================

    INF = float("inf")

    dp = [
        [INF] * (number_of_rows + 1)
        for _ in range(8)
    ]

    parent = [
        [None] * (number_of_rows + 1)
        for _ in range(8)
    ]

    dp[0][0] = 0


    for day in range(
        1,
        8
    ):

        for end in range(
            day,
            number_of_rows + 1
        ):

            for start in range(
                day - 1,
                end
            ):

                previous = dp[
                    day - 1
                ][start]


                if previous == INF:
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
                    previous
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
    # Recover boundaries
    # ========================================================

    boundaries = []

    end = number_of_rows


    for day in range(
        7,
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
    ) in enumerate(boundaries):

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
    "Automatically plan loom production using the current "
    "production week, weekly capacity and Excel row order."
)


# ============================================================
# CURRENT WEEK DISPLAY
# ============================================================

today = date.today()

current_week = (
    get_current_production_week()
)

week_start = (
    get_week_start_date(
        current_week,
        today.year
    )
)

week_end = (
    week_start
    + timedelta(days=6)
)


st.info(
    f"📅 Today: {today.strftime('%d-%m-%Y')}  |  "
    f"Current Production Week: WK-{current_week}  |  "
    f"Week: {week_start.strftime('%d-%m-%Y')} → "
    f"{week_end.strftime('%d-%m-%Y')}"
)


# ============================================================
# UPLOAD
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
    "The Starting Week is automatically determined from "
    "today's date."
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
# RUN
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
    # CURRENT WEEK
    # ========================================================

    planning_date = date.today()

    current_week = (
        planning_date.isocalendar().week
    )

    planning_year = (
        planning_date.year
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
                f"{loom}: Weekly Capacity must be greater than 0."
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
    # C = Product / Description
    # D = Roll Width
    # E = Roll Length
    # F = Quantity
    # G = Loom
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
    # OUTPUT HEADERS
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
    # READ ROWS IN EXACT EXCEL ORDER
    # ========================================================

    loom_rows = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        loom_rows[
            loom
        ] = []


    # --------------------------------------------------------
    # IMPORTANT:
    #
    # We loop through the Excel rows directly.
    #
    # There is NO sorting.
    #
    # External Document No. is NOT used.
    # --------------------------------------------------------

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
        # Append exactly as it appears in Excel.
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

    loom_summary = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        weekly_capacity = (
            settings["capacity"]
        )


        current_week = (
            current_week
        )


        current_load = 0.0

        rows_processed = 0

        weeks_used = []


        # ----------------------------------------------------
        # Excel order.
        # ----------------------------------------------------

        for excel_row, quantity in loom_rows.get(
            loom,
            []
        ):


            # ------------------------------------------------
            # Do NOT split product.
            # ------------------------------------------------

            if (

                current_load > 0

                and

                current_load + quantity
                > weekly_capacity

            ):

                current_week = (
                    next_week(
                        current_week
                    )
                )

                current_load = 0.0


            # ------------------------------------------------
            # Assign week.
            # ------------------------------------------------

            row_week[
                excel_row
            ] = current_week


            # ------------------------------------------------
            # Update weekly load.
            # ------------------------------------------------

            current_load += quantity

            rows_processed += 1


            if current_week not in weeks_used:

                weeks_used.append(
                    current_week
                )


        loom_summary[
            loom
        ] = {

            "capacity":
                weekly_capacity,

            "starting_week":
                current_week
                if rows_processed == 0
                else weeks_used[0],

            "final_week":
                current_week,

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
        # Group rows by production week.
        #
        # NO SORTING.
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


            if week not in week_rows:

                week_rows[
                    week
                ] = []


            week_rows[
                week
            ].append(
                (
                    excel_row,
                    quantity
                )
            )


        # ====================================================
        # PROCESS EACH WEEK
        # ====================================================

        for production_week, rows in week_rows.items():

            # ------------------------------------------------
            # DO NOT SORT ROWS.
            # ------------------------------------------------

            week_start = (
                get_week_start_date(
                    production_week,
                    planning_year
                )
            )


            assignments = (
                distribute_products_across_week(
                    rows,
                    week_start
                )
            )


            for excel_row, assigned_date in (
                assignments.items()
            ):

                row_dates[
                    excel_row
                ] = assigned_date


    # ========================================================
    # WRITE RESULTS
    # ========================================================

    for excel_row, week in row_week.items():

        if excel_row not in row_dates:
            continue


        production_date = (
            row_dates[
                excel_row
            ]
        )


        # ----------------------------------------------------
        # Week
        # ----------------------------------------------------

        worksheet.cell(
            row=excel_row,
            column=WEEK_COL
        ).value = (
            f"WK-{week}"
        )


        # ----------------------------------------------------
        # Date
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
    # NO SORTING
    #
    # The Excel order remains exactly as supplied.
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
