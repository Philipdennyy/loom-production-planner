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
    }
])


# ============================================================
# SESSION STATE
# ============================================================

if "loom_settings" not in st.session_state:
    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_loom_name(value):

    if value is None:
        return ""

    return (
        str(value)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


def clean_document_number(value):

    if value is None:
        return ""

    return str(value).strip()


def next_week(week):

    if week >= 52:
        return 1

    return week + 1


def week_difference(start_week, target_week):

    return (
        target_week - start_week
    ) % 52


def get_week_start_date(
    planning_start_date,
    starting_week,
    production_week
):

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
# BALANCED DATE ASSIGNMENT
# ============================================================

def assign_balanced_dates(
    rows,
    week_start_date
):
    """
    Distribute complete rows across the 7 days.

    Rules:
    - No row/product is split.
    - Original document sequence is preserved.
    - A document may continue onto another day.
    - The next document may start on the same day.
    - No daily capacity is imposed.
    - Attempts to balance total quantity across 7 days.
    """

    if not rows:
        return {}

    number_of_rows = len(rows)

    quantities = [
        float(item[1])
        for item in rows
    ]

    total_quantity = sum(quantities)

    # --------------------------------------------------------
    # If there are 7 or fewer rows, distribute sequentially.
    # --------------------------------------------------------

    if number_of_rows <= 7:

        result = {}

        for index, item in enumerate(rows):

            excel_row = item[0]

            result[excel_row] = (
                week_start_date
                + timedelta(days=index)
            )

        return result

    # --------------------------------------------------------
    # Ideal quantity per day.
    #
    # This is NOT a daily capacity.
    # It is only a balancing target.
    # --------------------------------------------------------

    ideal_daily_quantity = (
        total_quantity / 7
    )

    # --------------------------------------------------------
    # Prefix sums
    # --------------------------------------------------------

    prefix = [0.0]

    for quantity in quantities:

        prefix.append(
            prefix[-1] + quantity
        )

    # --------------------------------------------------------
    # Dynamic programming.
    #
    # Divide the sequence into 7 consecutive groups.
    # --------------------------------------------------------

    INF = float("inf")

    dp = [
        [INF] * (number_of_rows + 1)
        for _ in range(8)
    ]

    parent = [
        [None] * (number_of_rows + 1)
        for _ in range(8)
    ]

    dp[0][0] = 0.0

    for day_count in range(1, 8):

        for end_row in range(
            day_count,
            number_of_rows + 1
        ):

            for start_row in range(
                day_count - 1,
                end_row
            ):

                previous_cost = dp[
                    day_count - 1
                ][start_row]

                if previous_cost == INF:
                    continue

                day_quantity = (
                    prefix[end_row]
                    - prefix[start_row]
                )

                difference = (
                    day_quantity
                    - ideal_daily_quantity
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
                    day_count
                ][end_row]:

                    dp[
                        day_count
                    ][end_row] = total_cost

                    parent[
                        day_count
                    ][end_row] = start_row

    # --------------------------------------------------------
    # Recover the best 7-day partition.
    # --------------------------------------------------------

    days_to_use = min(
        7,
        number_of_rows
    )

    boundaries = []

    end_row = number_of_rows

    for day_count in range(
        days_to_use,
        0,
        -1
    ):

        start_row = parent[
            day_count
        ][end_row]

        # Safety fallback
        if start_row is None:

            boundaries = []

            rows_per_day = (
                number_of_rows
                // days_to_use
            )

            remainder = (
                number_of_rows
                % days_to_use
            )

            current = 0

            for d in range(days_to_use):

                extra = (
                    1
                    if d < remainder
                    else 0
                )

                start = current

                end = (
                    current
                    + rows_per_day
                    + extra
                )

                boundaries.append(
                    (start, end)
                )

                current = end

            break

        boundaries.append(
            (
                start_row,
                end_row
            )
        )

        end_row = start_row

    boundaries.reverse()

    # --------------------------------------------------------
    # Assign dates.
    # --------------------------------------------------------

    result = {}

    for day_index, (
        start_row,
        end_row
    ) in enumerate(boundaries):

        assigned_date = (
            week_start_date
            + timedelta(
                days=day_index
            )
        )

        for row_index in range(
            start_row,
            end_row
        ):

            excel_row = rows[
                row_index
            ][0]

            result[
                excel_row
            ] = assigned_date

    return result


# ============================================================
# SORT FINAL EXCEL
# ============================================================

def sort_planned_rows(
    worksheet,
    loom_settings,
    loom_col,
    production_week_col,
    production_date_col,
    row_sequence
):

    """
    Sort:

    Loom
      ↓
    Production Week
      ↓
    Production Date
      ↓
    Original production sequence
    """

    # --------------------------------------------------------
    # Save merged ranges.
    # --------------------------------------------------------

    merged_ranges = [
        str(rng)
        for rng in worksheet.merged_cells.ranges
    ]

    # --------------------------------------------------------
    # Temporarily remove merged cells.
    # --------------------------------------------------------

    for rng in list(
        worksheet.merged_cells.ranges
    ):

        try:
            worksheet.unmerge_cells(
                str(rng)
            )
        except Exception:
            pass

    # --------------------------------------------------------
    # Loom order.
    # --------------------------------------------------------

    loom_order = {}

    for index, loom in enumerate(
        loom_settings.keys()
    ):

        loom_order[
            clean_loom_name(loom)
        ] = index

    # --------------------------------------------------------
    # Collect rows.
    # --------------------------------------------------------

    rows = []

    for row_number in range(
        2,
        worksheet.max_row + 1
    ):

        loom = clean_loom_name(
            worksheet.cell(
                row=row_number,
                column=loom_col
            ).value
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
        # Week number.
        # ----------------------------------------------------

        week_number = None

        if week_value is not None:

            try:

                week_number = int(
                    str(week_value)
                    .replace("WK-", "")
                    .strip()
                )

            except Exception:

                week_number = None

        starting_week = loom_settings.get(
            loom,
            {}
        ).get(
            "starting_week",
            1
        )

        if week_number is not None:

            week_rank = (
                week_number
                - starting_week
            ) % 52

        else:

            week_rank = 999

        # ----------------------------------------------------
        # Date.
        # ----------------------------------------------------

        if date_value is not None:

            try:

                date_sort = pd.to_datetime(
                    date_value
                )

            except Exception:

                date_sort = pd.Timestamp.max

        else:

            date_sort = pd.Timestamp.max

        # ----------------------------------------------------
        # Original production sequence.
        # ----------------------------------------------------

        sequence = row_sequence.get(
            row_number,
            999999999
        )

        rows.append({

            "original_row":
                row_number,

            "loom_rank":
                loom_order.get(
                    loom,
                    999
                ),

            "week_rank":
                week_rank,

            "date":
                date_sort,

            "sequence":
                sequence
        })

    # --------------------------------------------------------
    # Sort.
    # --------------------------------------------------------

    rows.sort(
        key=lambda x: (
            x["loom_rank"],
            x["week_rank"],
            x["date"],
            x["sequence"],
            x["original_row"]
        )
    )

    # --------------------------------------------------------
    # Store complete row data.
    # --------------------------------------------------------

    max_col = worksheet.max_column

    row_data = {}

    for item in rows:

        original_row = item[
            "original_row"
        ]

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

    # --------------------------------------------------------
    # Rewrite rows.
    # --------------------------------------------------------

    for new_row, item in enumerate(
        rows,
        start=2
    ):

        original_row = item[
            "original_row"
        ]

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

    # --------------------------------------------------------
    # Restore merged ranges.
    # --------------------------------------------------------

    for rng in merged_ranges:

        try:

            worksheet.merge_cells(
                rng
            )

        except Exception:

            pass


# ============================================================
# APPLICATION TITLE
# ============================================================

st.title(
    "🏭 Loom Production Planner"
)

st.write(
    "Automatically plan loom production using weekly "
    "capacity, external document sequence and production dates."
)


# ============================================================
# FILE UPLOAD
# ============================================================

st.subheader(
    "1️⃣ Upload Production Excel"
)

uploaded_file = st.file_uploader(
    "Select the Excel file",
    type=["xlsx"]
)


# ============================================================
# START DATE
# ============================================================

st.subheader(
    "2️⃣ Planning Start Date"
)

planning_start_date = st.date_input(
    "Select the starting date",
    value=date.today()
)

st.caption(
    "The selected date is the first production day "
    "of the starting production week."
)


# ============================================================
# LOOM SETTINGS
# ============================================================

st.subheader(
    "3️⃣ Loom Settings"
)

st.caption(
    "Enable/disable looms and edit weekly capacity "
    "and starting week."
)

edited_settings = st.data_editor(

    st.session_state.loom_settings,

    use_container_width=True,

    hide_index=True,

    column_config={

        "Use":
            st.column_config.CheckboxColumn(
                "Use",
                help="Include this loom in production planning.",
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
                    "for this loom in one week."
                ),
                min_value=1,
                step=50
            ),

        "Starting Week":
            st.column_config.NumberColumn(
                "Starting Week",
                help=(
                    "Production week in which "
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
# MAIN PLANNING PROCESS
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
    # READ LOOM SETTINGS
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

        if capacity <= 0:

            st.error(
                f"Weekly capacity for {loom} "
                "must be greater than 0."
            )

            st.stop()

        if (
            starting_week < 1
            or
            starting_week > 52
        ):

            st.error(
                f"Starting week for {loom} "
                "must be between 1 and 52."
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
            f"Unable to open Excel file: {e}"
        )

        st.stop()


    # ========================================================
    # CHECK SHEETS
    # ========================================================

    if not workbook.sheetnames:

        st.error(
            "The Excel file contains no worksheets."
        )

        st.stop()


    # ========================================================
    # USE FIRST WORKSHEET
    # ========================================================

    worksheet = workbook[
        workbook.sheetnames[0]
    ]


    # ========================================================
    # FIXED INPUT COLUMNS
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
    # Output:
    #
    # H = Production Week
    # I = Production Date
    # J = Day
    #
    # ========================================================

    DOCUMENT_COL = 2
    QUANTITY_COL = 6
    LOOM_COL = 7

    PRODUCTION_WEEK_COL = 8
    PRODUCTION_DATE_COL = 9
    DAY_COL = 10


    # ========================================================
    # PREPARE COLUMN H
    # ========================================================

    h_header = worksheet.cell(
        row=1,
        column=PRODUCTION_WEEK_COL
    ).value

    if (

        h_header is not None

        and

        str(h_header).strip().lower()
        != "production week"

    ):

        h_has_data = any(

            worksheet.cell(
                row=row,
                column=PRODUCTION_WEEK_COL
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )

        if h_has_data:

            worksheet.insert_cols(
                PRODUCTION_WEEK_COL
            )


    worksheet.cell(
        row=1,
        column=PRODUCTION_WEEK_COL
    ).value = "Production Week"


    # ========================================================
    # PREPARE COLUMN I
    # ========================================================

    i_header = worksheet.cell(
        row=1,
        column=PRODUCTION_DATE_COL
    ).value

    if (

        i_header is not None

        and

        str(i_header).strip().lower()
        != "production date"

    ):

        i_has_data = any(

            worksheet.cell(
                row=row,
                column=PRODUCTION_DATE_COL
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )

        if i_has_data:

            worksheet.insert_cols(
                PRODUCTION_DATE_COL
            )


    worksheet.cell(
        row=1,
        column=PRODUCTION_DATE_COL
    ).value = "Production Date"


    # ========================================================
    # PREPARE COLUMN J
    # ========================================================

    j_header = worksheet.cell(
        row=1,
        column=DAY_COL
    ).value

    if (

        j_header is not None

        and

        str(j_header).strip().lower()
        != "day"

    ):

        j_has_data = any(

            worksheet.cell(
                row=row,
                column=DAY_COL
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )

        if j_has_data:

            worksheet.insert_cols(
                DAY_COL
            )


    worksheet.cell(
        row=1,
        column=DAY_COL
    ).value = "Day"


    # ========================================================
    # CLEAR OLD PLANNING OUTPUT
    # ========================================================

    for row in range(
        2,
        worksheet.max_row + 1
    ):

        worksheet.cell(
            row=row,
            column=PRODUCTION_WEEK_COL
        ).value = None

        worksheet.cell(
            row=row,
            column=PRODUCTION_DATE_COL
        ).value = None

        worksheet.cell(
            row=row,
            column=DAY_COL
        ).value = None


    # ========================================================
    # CREATE DOCUMENT GROUPS
    # ========================================================
    #
    # SAME DOCUMENT NUMBER = SAME PRODUCTION GROUP.
    #
    # The original Excel order is preserved.
    #
    # ========================================================

    loom_documents = {}

    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue

        loom_documents[
            loom
        ] = {}

        for excel_row in range(
            2,
            worksheet.max_row + 1
        ):

            excel_loom = clean_loom_name(
                worksheet.cell(
                    row=excel_row,
                    column=LOOM_COL
                ).value
            )

            if excel_loom != loom:
                continue

            document = clean_document_number(
                worksheet.cell(
                    row=excel_row,
                    column=DOCUMENT_COL
                ).value
            )

            # Blank document numbers are treated
            # as individual rows.

            if document == "":

                document = (
                    f"ROW-{excel_row}"
                )

            if document not in (
                loom_documents[loom]
            ):

                loom_documents[
                    loom
                ][document] = []

            loom_documents[
                loom
            ][document].append(
                excel_row
            )


    # ========================================================
    # PLANNING STORAGE
    # ========================================================

    row_week = {}

    row_date = {}

    row_sequence = {}

    loom_summary = {}


    # ========================================================
    # FIRST PASS:
    # ASSIGN PRODUCTION WEEKS
    # ========================================================

    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue

        weekly_capacity = (
            settings["capacity"]
        )

        starting_week = (
            settings["starting_week"]
        )

        current_week = (
            starting_week
        )

        current_week_load = 0.0

        rows_processed = 0

        weeks_used = []

        sequence_counter = 0

        documents = loom_documents.get(
            loom,
            {}
        )


        # ----------------------------------------------------
        # Documents are processed in their original order.
        # ----------------------------------------------------

        for document, document_rows in (
            documents.items()
        ):

            for excel_row in document_rows:

                quantity = worksheet.cell(
                    row=excel_row,
                    column=QUANTITY_COL
                ).value

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
                # WEEKLY CAPACITY
                # =================================================
                #
                # DO NOT split a product row.
                #
                # If the complete row does not fit in the
                # remaining weekly capacity, move it to the
                # next week.
                #
                # =================================================

                if (

                    current_week_load > 0

                    and

                    current_week_load
                    + quantity
                    > weekly_capacity

                ):

                    current_week = (
                        next_week(
                            current_week
                        )
                    )

                    current_week_load = 0.0


                # ------------------------------------------------
                # Warning if a single row is larger than the
                # entire weekly capacity.
                # ------------------------------------------------

                if quantity > weekly_capacity:

                    st.warning(

                        f"{loom} - External Document "
                        f"{document}: quantity {quantity} "
                        f"is greater than weekly capacity "
                        f"{weekly_capacity}. "
                        "The product will remain unsplit."
                    )


                # ------------------------------------------------
                # Store week.
                # ------------------------------------------------

                row_week[
                    excel_row
                ] = current_week


                # ------------------------------------------------
                # Store original production sequence.
                # ------------------------------------------------

                row_sequence[
                    excel_row
                ] = sequence_counter

                sequence_counter += 1


                # ------------------------------------------------
                # Update weekly quantity.
                # ------------------------------------------------

                current_week_load += quantity


                rows_processed += 1


                if current_week not in weeks_used:

                    weeks_used.append(
                        current_week
                    )


        # ====================================================
        # SAVE LOOM SUMMARY
        #
        # THIS WAS THE MISSING PART IN THE PREVIOUS VERSION.
        # ====================================================

        loom_summary[
            loom
        ] = {

            "capacity":
                weekly_capacity,

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
    # SECOND PASS:
    # ASSIGN DATES
    # ========================================================

    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue

        starting_week = (
            settings["starting_week"]
        )


        # ----------------------------------------------------
        # Group rows by production week.
        # ----------------------------------------------------

        week_rows = {}


        for excel_row, assigned_week in (
            row_week.items()
        ):

            excel_loom = clean_loom_name(
                worksheet.cell(
                    row=excel_row,
                    column=LOOM_COL
                ).value
            )

            if excel_loom != loom:
                continue


            quantity = worksheet.cell(
                row=excel_row,
                column=QUANTITY_COL
            ).value


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


            if assigned_week not in week_rows:

                week_rows[
                    assigned_week
                ] = []


            week_rows[
                assigned_week
            ].append(

                (
                    excel_row,
                    quantity,
                    row_sequence.get(
                        excel_row,
                        999999
                    )
                )
            )


        # ----------------------------------------------------
        # Process each week.
        # ----------------------------------------------------

        for production_week, rows in (
            week_rows.items()
        ):

            # ------------------------------------------------
            # Preserve original document/row sequence.
            # ------------------------------------------------

            rows.sort(
                key=lambda x: x[2]
            )


            # ------------------------------------------------
            # Get week starting date.
            # ------------------------------------------------

            week_start_date = (
                get_week_start_date(
                    planning_start_date,
                    starting_week,
                    production_week
                )
            )


            # ------------------------------------------------
            # Balance production across 7 days.
            # ------------------------------------------------

            date_assignments = (
                assign_balanced_dates(
                    rows,
                    week_start_date
                )
            )


            # ------------------------------------------------
            # Save dates.
            # ------------------------------------------------

            for (
                excel_row,
                quantity,
                sequence
            ) in rows:

                if excel_row not in date_assignments:
                    continue


                row_date[
                    excel_row
                ] = date_assignments[
                    excel_row
                ]


    # ========================================================
    # WRITE OUTPUT TO EXCEL
    # ========================================================

    for excel_row, week in (
        row_week.items()
    ):

        if excel_row not in row_date:
            continue


        assigned_date = (
            row_date[
                excel_row
            ]
        )


        # ----------------------------------------------------
        # Production Week
        # ----------------------------------------------------

        worksheet.cell(
            row=excel_row,
            column=PRODUCTION_WEEK_COL
        ).value = (
            f"WK-{week}"
        )


        # ----------------------------------------------------
        # Production Date
        # ----------------------------------------------------

        worksheet.cell(
            row=excel_row,
            column=PRODUCTION_DATE_COL
        ).value = assigned_date

        worksheet.cell(
            row=excel_row,
            column=PRODUCTION_DATE_COL
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
            assigned_date.strftime(
                "%A"
            )
        )


    # ========================================================
    # SORT FINAL EXCEL
    # ========================================================

    try:

        sort_planned_rows(

            worksheet=worksheet,

            loom_settings=loom_settings,

            loom_col=LOOM_COL,

            production_week_col=(
                PRODUCTION_WEEK_COL
            ),

            production_date_col=(
                PRODUCTION_DATE_COL
            ),

            row_sequence=row_sequence
        )

    except Exception as e:

        st.warning(
            "Planning was completed, but final Excel "
            f"sorting could not be completed: {e}"
        )


    # ========================================================
    # CREATE OUTPUT FILE
    # ========================================================

    output = BytesIO()

    workbook.save(
        output
    )

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

    st.subheader(
        "📊 Planning Summary"
    )


    summary_rows = []


    for loom, info in (
        loom_summary.items()
    ):

        # ----------------------------------------------------
        # Only show enabled looms that actually had data.
        # ----------------------------------------------------

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
            "No production rows were found for the "
            "selected/enabled looms."
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
