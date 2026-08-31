import streamlit as st
import pandas as pd

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from io import BytesIO
from datetime import date, timedelta
from copy import copy


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
    {
        "Use": True,
        "Loom": "ALT1",
        "Weekly Capacity": 1500,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "ALT2",
        "Weekly Capacity": 1200,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "ALT4",
        "Weekly Capacity": 800,
        "Starting Week": 34
    },

    {
        "Use": True,
        "Loom": "ALT5",
        "Weekly Capacity": 1300,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "ALT6",
        "Weekly Capacity": 800,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP1",
        "Weekly Capacity": 500,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP2",
        "Weekly Capacity": 700,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP3",
        "Weekly Capacity": 900,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP4",
        "Weekly Capacity": 900,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP5",
        "Weekly Capacity": 1100,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP6",
        "Weekly Capacity": 1200,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "RP7",
        "Weekly Capacity": 1000,
        "Starting Week": 33
    },

    {
        "Use": True,
        "Loom": "DB4M1",
        "Weekly Capacity": 300,
        "Starting Week": 33
    },
])


# ============================================================
# SESSION STATE
# ============================================================

if "loom_settings" not in st.session_state:

    st.session_state.loom_settings = (
        DEFAULT_SETTINGS.copy()
    )


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_loom_name(loom):
    """
    Standardize loom names.

    ALT 1
    ALT-1
    alt1

    will all become:

    ALT1
    """

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
# NEXT WEEK
# ============================================================

def next_week(week):
    """
    Move to next production week.

    WK-52 -> WK-1
    """

    if week >= 52:
        return 1

    return week + 1


# ============================================================
# WEEK DISTANCE
# ============================================================

def week_difference(
    start_week,
    target_week
):
    """
    Calculate number of weeks between
    starting week and target week.

    Example:

    Starting WK-33

    WK-33 = 0
    WK-34 = 1
    WK-35 = 2

    WK-52 = 19
    WK-1  = 20
    WK-2  = 21
    """

    return (
        target_week - start_week
    ) % 52


# ============================================================
# GET WEEK START DATE
# ============================================================

def get_week_start_date(
    planning_start_date,
    starting_week,
    production_week
):
    """
    Calculate actual date for the beginning
    of a production week.

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
        + timedelta(
            days=difference * 7
        )
    )


# ============================================================
# SORT FINAL EXCEL
# ============================================================

def sort_planned_rows(
    worksheet,
    loom_settings,
    loom_col,
    production_week_col,
    production_date_col
):
    """
    Sort complete Excel rows by:

    1. Loom
    2. Production Week
    3. Production Date
    4. Original order

    Handles merged cells by temporarily unmerging them,
    sorting the rows, and then restoring the merged ranges.
    """

    # ========================================================
    # SAVE MERGED RANGES
    # ========================================================

    merged_ranges = [
        str(rng)
        for rng in worksheet.merged_cells.ranges
    ]


    # ========================================================
    # TEMPORARILY UNMERGE
    # ========================================================

    for rng in list(
        worksheet.merged_cells.ranges
    ):

        worksheet.unmerge_cells(
            str(rng)
        )


    # ========================================================
    # LOOM ORDER
    # ========================================================

    loom_order = {}

    for index, loom in enumerate(
        loom_settings.keys()
    ):

        loom_order[
            clean_loom_name(loom)
        ] = index


    # ========================================================
    # COLLECT ROW INFORMATION
    # ========================================================

    rows = []


    for row_number in range(
        2,
        worksheet.max_row + 1
    ):

        loom_value = worksheet.cell(
            row=row_number,
            column=loom_col
        ).value


        loom = clean_loom_name(
            loom_value
        )


        week_value = worksheet.cell(
            row=row_number,
            column=production_week_col
        ).value


        date_value = worksheet.cell(
            row=row_number,
            column=production_date_col
        ).value


        # ----------------------------------------------------
        # Extract week number
        # ----------------------------------------------------

        week_number = None


        if week_value is not None:

            try:

                week_number = int(
                    str(week_value)
                    .replace("WK-", "")
                    .strip()
                )

            except:

                week_number = None


        # ----------------------------------------------------
        # Starting week
        # ----------------------------------------------------

        starting_week = loom_settings.get(
            loom,
            {}
        ).get(
            "starting_week",
            1
        )


        # ----------------------------------------------------
        # Relative week position
        # ----------------------------------------------------

        if week_number is not None:

            week_rank = (
                week_number
                - starting_week
            ) % 52

        else:

            week_rank = 999


        # ----------------------------------------------------
        # Date
        # ----------------------------------------------------

        if date_value is not None:

            try:

                date_sort_value = pd.to_datetime(
                    date_value
                )

            except:

                date_sort_value = (
                    pd.Timestamp.max
                )

        else:

            date_sort_value = (
                pd.Timestamp.max
            )


        rows.append({

            "original_row":
                row_number,

            "loom":
                loom,

            "loom_rank":
                loom_order.get(
                    loom,
                    999
                ),

            "week_rank":
                week_rank,

            "date":
                date_sort_value
        })


    # ========================================================
    # SORT
    # ========================================================

    rows.sort(
        key=lambda x: (
            x["loom_rank"],
            x["week_rank"],
            x["date"],
            x["original_row"]
        )
    )


    # ========================================================
    # STORE COMPLETE ROW DATA
    # ========================================================

    max_col = worksheet.max_column

    row_data = {}


    for item in rows:

        original_row = (
            item["original_row"]
        )

        row_data[
            original_row
        ] = []


        for col in range(
            1,
            max_col + 1
        ):

            cell = worksheet.cell(
                row=original_row,
                column=col
            )


            row_data[
                original_row
            ].append({

                "value":
                    cell.value,

                "style":
                    copy(cell._style),

                "number_format":
                    cell.number_format,

                "font":
                    copy(cell.font),

                "fill":
                    copy(cell.fill),

                "border":
                    copy(cell.border),

                "alignment":
                    copy(cell.alignment),

                "protection":
                    copy(cell.protection),

                "hyperlink":
                    cell.hyperlink,

                "comment":
                    cell.comment
            })


    # ========================================================
    # REWRITE SORTED ROWS
    # ========================================================

    for new_row, item in enumerate(
        rows,
        start=2
    ):

        original_row = (
            item["original_row"]
        )


        for col in range(
            1,
            max_col + 1
        ):

            source = row_data[
                original_row
            ][col - 1]


            target = worksheet.cell(
                row=new_row,
                column=col
            )


            # Safety check
            if isinstance(
                target,
                MergedCell
            ):

                continue


            target.value = (
                source["value"]
            )

            target._style = copy(
                source["style"]
            )

            target.number_format = (
                source["number_format"]
            )

            target.font = copy(
                source["font"]
            )

            target.fill = copy(
                source["fill"]
            )

            target.border = copy(
                source["border"]
            )

            target.alignment = copy(
                source["alignment"]
            )

            target.protection = copy(
                source["protection"]
            )

            target.hyperlink = (
                source["hyperlink"]
            )

            target.comment = (
                source["comment"]
            )


    # ========================================================
    # RESTORE MERGED RANGES
    # ========================================================

    for rng in merged_ranges:

        try:

            worksheet.merge_cells(
                rng
            )

        except Exception:

            # If a merged range cannot be restored,
            # don't crash the entire application.
            pass


# ============================================================
# TITLE
# ============================================================

st.title(
    "🏭 Loom Production Planner"
)


st.write(
    "Automatically assign production weeks and dates "
    "based on loom capacity and production quantity."
)


# ============================================================
# STEP 1 — UPLOAD EXCEL
# ============================================================

st.subheader(
    "1️⃣ Upload Production Excel"
)


uploaded_file = st.file_uploader(
    "Select the Excel file",
    type=["xlsx"]
)


# ============================================================
# STEP 2 — START DATE
# ============================================================

st.subheader(
    "2️⃣ Planning Start Date"
)


planning_start_date = st.date_input(
    "Select the starting date",
    value=date.today()
)


st.caption(
    "This date represents the first day of the "
    "starting production week. Each week has 7 working days."
)


# ============================================================
# STEP 3 — LOOM SETTINGS
# ============================================================

st.subheader(
    "3️⃣ Loom Settings"
)


st.caption(
    "Enable or disable looms and edit their weekly "
    "capacity and starting production week."
)


edited_settings = st.data_editor(

    st.session_state.loom_settings,

    use_container_width=True,

    hide_index=True,

    column_config={

        "Use":
            st.column_config.CheckboxColumn(
                "Use",
                help="Include this loom in planning.",
                default=True
            ),

        "Loom":
            st.column_config.TextColumn(
                "Loom",
                disabled=True
            ),

        "Weekly Capacity":
            st.column_config.NumberColumn(
                "Weekly Capacity",
                help=(
                    "Maximum production quantity "
                    "for this loom per week."
                ),
                min_value=1,
                step=50
            ),

        "Starting Week":
            st.column_config.NumberColumn(
                "Starting Week",
                help=(
                    "Production week from which "
                    "this loom starts."
                ),
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

if st.button(
    "🔄 Reset to Default Settings"
):

    st.session_state.loom_settings = (
        DEFAULT_SETTINGS.copy()
    )

    if (
        "loom_settings_editor"
        in st.session_state
    ):

        del st.session_state[
            "loom_settings_editor"
        ]

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

        loom = clean_loom_name(
            row["Loom"]
        )


        use_loom = bool(
            row["Use"]
        )


        capacity = float(
            row["Weekly Capacity"]
        )


        starting_week = int(
            row["Starting Week"]
        )


        # ----------------------------------------------------
        # Validate capacity
        # ----------------------------------------------------

        if capacity <= 0:

            st.error(
                f"Weekly capacity for {loom} "
                f"must be greater than 0."
            )

            st.stop()


        # ----------------------------------------------------
        # Validate week
        # ----------------------------------------------------

        if (
            starting_week < 1
            or
            starting_week > 52
        ):

            st.error(
                f"Starting week for {loom} "
                f"must be between 1 and 52."
            )

            st.stop()


        loom_settings[
            loom
        ] = {

            "use":
                use_loom,

            "capacity":
                capacity,

            "starting_week":
                starting_week
        }


    # ========================================================
    # READ EXCEL
    # ========================================================

    try:

        file_data = (
            uploaded_file.getvalue()
        )


        workbook = load_workbook(
            filename=BytesIO(
                file_data
            )
        )

    except Exception as e:

        st.error(
            f"Unable to open the Excel file: {e}"
        )

        st.stop()


    # ========================================================
    # CHECK WORKSHEET
    # ========================================================

    if not workbook.sheetnames:

        st.error(
            "The Excel file contains no worksheets."
        )

        st.stop()


    # ========================================================
    # USE FIRST SHEET
    # ========================================================

    worksheet = workbook[
        workbook.sheetnames[0]
    ]


    # ========================================================
    # FIXED INPUT COLUMNS
    # ========================================================
    #
    # A = No.
    # B = External Document
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


    # ========================================================
    # PRODUCTION WEEK COLUMN
    # ========================================================

    production_week_col = 8


    h_value = worksheet.cell(
        row=1,
        column=8
    ).value


    # --------------------------------------------------------
    # If H is not Production Week and contains data,
    # insert a new column.
    # --------------------------------------------------------

    if (

        h_value is not None

        and

        str(h_value).strip().lower()
        != "production week"

    ):

        # Check whether H actually contains data

        h_has_data = any(

            worksheet.cell(
                row=row,
                column=8
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )


        if h_has_data:

            worksheet.insert_cols(
                8
            )


    worksheet.cell(
        row=1,
        column=8
    ).value = (
        "Production Week"
    )


    production_week_col = 8


    # ========================================================
    # PRODUCTION DATE COLUMN
    # ========================================================

    production_date_col = 9


    i_value = worksheet.cell(
        row=1,
        column=production_date_col
    ).value


    if (

        i_value is not None

        and

        str(i_value).strip().lower()
        != "production date"

    ):

        i_has_data = any(

            worksheet.cell(
                row=row,
                column=production_date_col
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )


        if i_has_data:

            worksheet.insert_cols(
                production_date_col
            )


    worksheet.cell(
        row=1,
        column=production_date_col
    ).value = (
        "Production Date"
    )


    # ========================================================
    # DAY COLUMN
    # ========================================================

    day_col = 10


    j_value = worksheet.cell(
        row=1,
        column=day_col
    ).value


    if (

        j_value is not None

        and

        str(j_value).strip().lower()
        != "day"

    ):

        j_has_data = any(

            worksheet.cell(
                row=row,
                column=day_col
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )


        if j_has_data:

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
    # CALCULATE PRODUCTION WEEK
    # ========================================================

    row_week = {}

    loom_summary = {}


    for loom, settings in loom_settings.items():

        # ----------------------------------------------------
        # Skip disabled loom
        # ----------------------------------------------------

        if not settings["use"]:

            continue


        capacity = (
            settings["capacity"]
        )


        starting_week = (
            settings["starting_week"]
        )


        current_week = (
            starting_week
        )


        current_week_load = 0


        rows_processed = 0


        weeks_used = []


        # ----------------------------------------------------
        # Process rows
        # ----------------------------------------------------

        for excel_row in range(
            2,
            worksheet.max_row + 1
        ):

            excel_loom = worksheet.cell(
                row=excel_row,
                column=LOOM_COL
            ).value


            cleaned_excel_loom = (
                clean_loom_name(
                    excel_loom
                )
            )


            # Only process current loom

            if (
                cleaned_excel_loom
                != loom
            ):

                continue


            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


            # ------------------------------------------------
            # Empty quantity
            # ------------------------------------------------

            if quantity is None:

                continue


            try:

                quantity = float(
                    quantity
                )

            except (
                ValueError,
                TypeError
            ):

                continue


            if quantity <= 0:

                continue


            # =================================================
            # PRODUCT LARGER THAN WEEKLY CAPACITY
            # =================================================

            if quantity > capacity:

                st.warning(
                    f"⚠️ {loom}: Row {excel_row} "
                    f"has quantity {quantity}, which is "
                    f"greater than the weekly capacity "
                    f"of {capacity}. "
                    f"The complete product will be kept "
                    f"together in one week."
                )


            # =================================================
            # WEEKLY CAPACITY LOGIC
            # =================================================

            if (

                current_week_load == 0

                or

                current_week_load
                + quantity
                <= capacity

            ):

                # Product fits into current week

                assigned_week = (
                    current_week
                )


                current_week_load += (
                    quantity
                )


            else:

                # ------------------------------------------------
                # Product does not fit.
                #
                # Move COMPLETE product to next week.
                #
                # NEVER split the product.
                # ------------------------------------------------

                current_week = (
                    next_week(
                        current_week
                    )
                )


                assigned_week = (
                    current_week
                )


                current_week_load = (
                    quantity
                )


            # ------------------------------------------------
            # Store result
            # ------------------------------------------------

            row_week[
                excel_row
            ] = assigned_week


            rows_processed += 1


            if (
                assigned_week
                not in weeks_used
            ):

                weeks_used.append(
                    assigned_week
                )


        # ====================================================
        # SUMMARY
        # ====================================================

        loom_summary[
            loom
        ] = {

            "capacity":
                capacity,

            "starting_week":
                starting_week,

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
    # ASSIGN PRODUCTION DATES
    # ========================================================
    #
    # IMPORTANT:
    #
    # There is NO daily capacity.
    #
    # The daily totals below are ONLY used to spread
    # complete products across the seven days.
    #
    # One product = one date.
    #
    # ========================================================

    for loom, settings in loom_settings.items():

        if not settings["use"]:

            continue


        starting_week = (
            settings["starting_week"]
        )


        # ----------------------------------------------------
        # Group rows by production week
        # ----------------------------------------------------

        week_rows = {}


        for excel_row, assigned_week in (
            row_week.items()
        ):

            excel_loom = worksheet.cell(
                row=excel_row,
                column=LOOM_COL
            ).value


            if (
                clean_loom_name(
                    excel_loom
                )
                != loom
            ):

                continue


            if (
                assigned_week
                not in week_rows
            ):

                week_rows[
                    assigned_week
                ] = []


            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


            week_rows[
                assigned_week
            ].append(
                (
                    excel_row,
                    float(quantity)
                )
            )


        # ----------------------------------------------------
        # Process each week
        # ----------------------------------------------------

        for production_week, rows in (
            week_rows.items()
        ):

            # ------------------------------------------------
            # Calculate beginning date of this week
            # ------------------------------------------------

            week_start_date = (
                get_week_start_date(
                    planning_start_date,
                    starting_week,
                    production_week
                )
            )


            # =================================================
            # DAY LOADS
            # =================================================

            day_totals = {

                day_number: 0.0

                for day_number in range(7)
            }


            # =================================================
            # ASSIGN PRODUCTS
            # =================================================

            for excel_row, quantity in rows:

                # ------------------------------------------------
                # Find currently least-loaded day
                # ------------------------------------------------

                selected_day = min(
                    day_totals,
                    key=day_totals.get
                )


                # ------------------------------------------------
                # Calculate actual date
                # ------------------------------------------------

                assigned_date = (
                    week_start_date
                    + timedelta(
                        days=selected_day
                    )
                )


                # ------------------------------------------------
                # Add COMPLETE PRODUCT
                #
                # Nothing is split.
                # ------------------------------------------------

                day_totals[
                    selected_day
                ] += quantity


                # =================================================
                # WRITE PRODUCTION WEEK
                # =================================================

                worksheet.cell(
                    row=excel_row,
                    column=production_week_col
                ).value = (
                    f"WK-{production_week}"
                )


                # =================================================
                # WRITE PRODUCTION DATE
                # =================================================

                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).value = (
                    assigned_date
                )


                worksheet.cell(
                    row=excel_row,
                    column=production_date_col
                ).number_format = (
                    "DD-MM-YYYY"
                )


                # =================================================
                # WRITE DAY
                # =================================================

                worksheet.cell(
                    row=excel_row,
                    column=day_col
                ).value = (
                    assigned_date.strftime(
                        "%A"
                    )
                )


    # ========================================================
    # SORT FINAL DATA
    # ========================================================

    try:

        sort_planned_rows(

            worksheet=worksheet,

            loom_settings=loom_settings,

            loom_col=LOOM_COL,

            production_week_col=(
                production_week_col
            ),

            production_date_col=(
                production_date_col
            )
        )

    except Exception as e:

        st.warning(
            "⚠️ Production planning was completed, "
            "but the final row sorting could not be "
            f"completed: {e}"
        )


    # ========================================================
    # SAVE FILE
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
        "✅ Production week and date planning "
        "completed successfully!"
    )


    # ========================================================
    # PLANNING SUMMARY
    # ========================================================

    st.subheader(
        "📊 Planning Summary"
    )


    summary_rows = []


    for loom, info in (
        loom_summary.items()
    ):

        if (
            info["rows_processed"]
            == 0
        ):

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
