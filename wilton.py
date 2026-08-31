import streamlit as st
import pandas as pd

from openpyxl import load_workbook
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
    {"Use": True, "Loom": "ALT1",  "Weekly Capacity": 1500, "Starting Week": 33},
    {"Use": True, "Loom": "ALT2",  "Weekly Capacity": 1200, "Starting Week": 33},
    {"Use": True, "Loom": "ALT4",  "Weekly Capacity": 800,  "Starting Week": 34},
    {"Use": True, "Loom": "ALT5",  "Weekly Capacity": 1300, "Starting Week": 33},
    {"Use": True, "Loom": "ALT6",  "Weekly Capacity": 800,  "Starting Week": 33},

    {"Use": True, "Loom": "RP1",   "Weekly Capacity": 500,  "Starting Week": 33},
    {"Use": True, "Loom": "RP2",   "Weekly Capacity": 700,  "Starting Week": 33},
    {"Use": True, "Loom": "RP3",   "Weekly Capacity": 900,  "Starting Week": 33},
    {"Use": True, "Loom": "RP4",   "Weekly Capacity": 900,  "Starting Week": 33},
    {"Use": True, "Loom": "RP5",   "Weekly Capacity": 1100, "Starting Week": 33},
    {"Use": True, "Loom": "RP6",   "Weekly Capacity": 1200, "Starting Week": 33},
    {"Use": True, "Loom": "RP7",   "Weekly Capacity": 1000, "Starting Week": 33},

    {"Use": True, "Loom": "DB4M1", "Weekly Capacity": 300,  "Starting Week": 33},
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

    if loom is None:
        return ""

    return (
        str(loom)
        .strip()
        .upper()
        .replace(" ", "")
        .replace("-", "")
    )


def clean_document_number(doc):

    if doc is None:
        return ""

    return str(doc).strip()


# ============================================================
# NEXT WEEK
# ============================================================

def next_week(week):

    if week >= 52:
        return 1

    return week + 1


# ============================================================
# WEEK DIFFERENCE
# ============================================================

def week_difference(start_week, target_week):

    return (
        target_week - start_week
    ) % 52


# ============================================================
# GET DATE FOR PRODUCTION WEEK
# ============================================================

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
# SORT FINAL ROWS
# ============================================================

def sort_planned_rows(
    worksheet,
    loom_settings,
    loom_col,
    document_col,
    production_week_col,
    production_date_col
):

    """
    Final sorting order:

    Loom
        ↓
    Production Week
        ↓
    Production Date
        ↓
    External Document No.
        ↓
    Original row order

    Complete Excel rows are moved together.
    """

    # --------------------------------------------------------
    # Save merged ranges
    # --------------------------------------------------------

    merged_ranges = [
        str(rng)
        for rng in worksheet.merged_cells.ranges
    ]

    # --------------------------------------------------------
    # Temporarily remove merged cells
    # --------------------------------------------------------

    for rng in list(
        worksheet.merged_cells.ranges
    ):

        worksheet.unmerge_cells(
            str(rng)
        )

    # --------------------------------------------------------
    # Loom order
    # --------------------------------------------------------

    loom_order = {}

    for index, loom in enumerate(
        loom_settings.keys()
    ):

        loom_order[
            clean_loom_name(loom)
        ] = index

    # --------------------------------------------------------
    # Collect row information
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

        document = clean_document_number(
            worksheet.cell(
                row=row_number,
                column=document_col
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
        # Week number
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
        # Starting week for loom
        # ----------------------------------------------------

        starting_week = loom_settings.get(
            loom,
            {}
        ).get(
            "starting_week",
            1
        )

        # ----------------------------------------------------
        # Week ranking
        # ----------------------------------------------------

        if week_number is not None:

            week_rank = (
                week_number
                - starting_week
            ) % 52

        else:

            week_rank = 999

        # ----------------------------------------------------
        # Date ranking
        # ----------------------------------------------------

        if date_value is not None:

            try:

                date_sort_value = pd.to_datetime(
                    date_value
                )

            except:

                date_sort_value = pd.Timestamp.max

        else:

            date_sort_value = pd.Timestamp.max

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
                date_sort_value,

            "document":
                document
        })

    # --------------------------------------------------------
    # Sort
    # --------------------------------------------------------

    rows.sort(
        key=lambda x: (
            x["loom_rank"],
            x["week_rank"],
            x["date"],
            x["original_row"]
        )
    )

    # --------------------------------------------------------
    # Store complete row data
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
    # Rewrite rows
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

            target.value = source[
                "value"
            ]

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
    # Restore merged ranges
    # --------------------------------------------------------

    for rng in merged_ranges:

        try:

            worksheet.merge_cells(
                rng
            )

        except Exception:

            pass


# ============================================================
# TITLE
# ============================================================

st.title(
    "🏭 Loom Production Planner"
)

st.write(
    "Automatically assign production weeks and dates "
    "based on loom capacity, external document sequence "
    "and production quantity."
)


# ============================================================
# UPLOAD
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
    "The selected date is the first day of the starting "
    "production week."
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
                help="Maximum quantity per week.",
                min_value=1,
                step=50
            ),

        "Starting Week":
            st.column_config.NumberColumn(
                "Starting Week",
                help="Starting production week.",
                min_value=1,
                max_value=52,
                step=1
            )
    },

    key="loom_settings_editor"
)


# ============================================================
# RESET
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
    # CHECK FILE
    # ========================================================

    if uploaded_file is None:

        st.error(
            "Please upload the Excel file first."
        )

        st.stop()

    # ========================================================
    # LOOM SETTINGS
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
    # LOAD EXCEL
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
    # FIRST SHEET
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
    # CREATE / FIND PRODUCTION WEEK COLUMN
    # ========================================================

    h_value = worksheet.cell(
        row=1,
        column=8
    ).value

    if (

        h_value is not None

        and

        str(h_value).strip().lower()
        != "production week"

    ):

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

            worksheet.insert_cols(8)

    worksheet.cell(
        row=1,
        column=8
    ).value = "Production Week"

    # ========================================================
    # CREATE / FIND PRODUCTION DATE COLUMN
    # ========================================================

    i_value = worksheet.cell(
        row=1,
        column=9
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
                column=9
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )

        if i_has_data:

            worksheet.insert_cols(9)

    worksheet.cell(
        row=1,
        column=9
    ).value = "Production Date"

    # ========================================================
    # CREATE / FIND DAY COLUMN
    # ========================================================

    j_value = worksheet.cell(
        row=1,
        column=10
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
                column=10
            ).value is not None

            for row in range(
                1,
                worksheet.max_row + 1
            )
        )

        if j_has_data:

            worksheet.insert_cols(10)

    worksheet.cell(
        row=1,
        column=10
    ).value = "Day"

    # ========================================================
    # CLEAR OLD PLANNING DATA
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
    # For each loom:
    #
    # Document 1 rows
    # Document 2 rows
    # Document 3 rows
    #
    # The order is based on the first appearance
    # of the document in the original Excel.
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

            if document == "":
                document = (
                    f"ROW-{excel_row}"
                )

            if document not in loom_documents[loom]:

                loom_documents[
                    loom
                ][document] = []

            loom_documents[
                loom
            ][document].append(
                excel_row
            )

    # ========================================================
    # STORE PLANNING RESULTS
    # ========================================================

    row_week = {}
    row_date = {}

    loom_summary = {}

    # ========================================================
    # PROCESS EACH LOOM
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

        # ----------------------------------------------------
        # 7-day production schedule
        #
        # day 0 = Monday / starting date
        # day 1 = next day
        # ...
        # day 6 = seventh day
        #
        # ----------------------------------------------------

        current_day_index = 0

        current_day_load = 0.0

        # ====================================================
        # PROCESS DOCUMENTS IN ORIGINAL ORDER
        # ====================================================

        documents = loom_documents.get(
            loom,
            {}
        )

        for document, document_rows in (
            documents.items()
        ):

            # ------------------------------------------------
            # Calculate soft daily target.
            #
            # This is NOT a daily capacity.
            #
            # It is only used to spread production across
            # the seven days.
            # ------------------------------------------------

            daily_target = (
                weekly_capacity / 7
            )

            # =================================================
            # PROCESS EVERY ROW OF DOCUMENT
            # =================================================

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
                # If this complete product doesn't fit into the
                # remaining weekly capacity, move it to next week.
                #
                # The product is NEVER split.
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

                    current_day_index = 0

                    current_day_load = 0.0

                # ------------------------------------------------
                # If product itself is larger than weekly capacity
                # it must still remain complete.
                # ------------------------------------------------

                if quantity > weekly_capacity:

                    st.warning(
                        f"{loom} - External Document "
                        f"{document}: quantity {quantity} "
                        f"is greater than weekly capacity "
                        f"{weekly_capacity}. "
                        f"The product will remain unsplit."
                    )

                # =================================================
                # DATE ALLOCATION
                # =================================================
                #
                # IMPORTANT:
                #
                # There is NO hard daily capacity.
                #
                # daily_target is only used to distribute work.
                #
                # =================================================

                # If current day already has production and adding
                # this complete product would go beyond the soft
                # target, move to next day.
                #
                # But if we are on day 6, we cannot create another
                # day inside the same week.

                if (

                    current_day_load > 0

                    and

                    current_day_load
                    + quantity
                    > daily_target

                    and

                    current_day_index < 6

                ):

                    current_day_index += 1

                    current_day_load = 0.0

                # ------------------------------------------------
                # Calculate date
                # ------------------------------------------------

                week_start_date = (
                    get_week_start_date(
                        planning_start_date,
                        starting_week,
                        current_week
                    )
                )

                assigned_date = (
                    week_start_date
                    + timedelta(
                        days=current_day_index
                    )
                )

                # ------------------------------------------------
                # Store planning
                # ------------------------------------------------

                row_week[
                    excel_row
                ] = current_week

                row_date[
                    excel_row
                ] = assigned_date

                # ------------------------------------------------
                # Update loads
                # ------------------------------------------------

                current_week_load += quantity

                current_day_load += quantity

                rows_processed += 1

                if current_week not in weeks_used:

                    weeks_used.append(
                        current_week
                    )

                # ------------------------------------------------
                # If this product pushed the day beyond the
                # soft target, the NEXT product can move to
                # the next day.
                #
                # This allows:
                #
                # DOC1 → 31 Aug
                # DOC1 → 01 Sep
                # DOC2 → 01 Sep
                #
                # when DOC1 finishes on 01 Sep.
                # ------------------------------------------------

                if (

                    current_day_load
                    >= daily_target

                    and

                    current_day_index < 6

                ):

                    current_day_index += 1

                    current_day_load = 0.0

            # =================================================
            # END OF DOCUMENT
            # =================================================
            #
            # We DO NOT automatically move to the next day.
            #
            # Therefore the next document can start on the
            # same day if there is room according to the
            # balancing logic.
            #
            # =================================================

        # ====================================================
        # SUMMARY
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
    # WRITE RESULTS TO EXCEL
    # ========================================================

    for excel_row, week in row_week.items():

        assigned_date = row_date[
            excel_row
        ]

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
        ).value = (
            assigned_date
        )

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

            document_col=DOCUMENT_COL,

            production_week_col=(
                PRODUCTION_WEEK_COL
            ),

            production_date_col=(
                PRODUCTION_DATE_COL
            )
        )

    except Exception as e:

        st.warning(
            "Production planning was completed, "
            "but final sorting could not be completed: "
            f"{e}"
        )

    # ========================================================
    # SAVE OUTPUT
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

    for loom, info in (
        loom_summary.items()
    ):

        if info[
            "rows_processed"
        ] == 0:

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
            "No production rows were found "
            "for the selected looms."
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
