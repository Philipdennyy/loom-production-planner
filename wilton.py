import streamlit as st
import pandas as pd

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from io import BytesIO
from datetime import date, timedelta
from copy import copy


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


if "loom_settings" not in st.session_state:
    st.session_state.loom_settings = DEFAULT_SETTINGS.copy()


# ============================================================
# HELPERS
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


def clean_document(value):

    if value is None:
        return ""

    return str(value).strip()


def get_quantity(value):

    if value is None:
        return None

    try:
        value = float(value)
    except (ValueError, TypeError):
        return None

    if value <= 0:
        return None

    return value


def next_week(week):

    return 1 if week >= 52 else week + 1


# ============================================================
# PRODUCTION WEEK -> DATE
# ============================================================
#
# Uses the normal calendar week system:
#
# Monday = first day of production week
# Sunday = last day of production week
#
# No specific date is hardcoded.
#
# Python calculates the correct calendar relationship.
#
# ============================================================

def get_week_start_date(
    production_week,
    year
):

    production_week = int(
        production_week
    )

    if production_week < 1:
        production_week = 1

    if production_week > 52:
        production_week = 52

    try:

        # ISO calendar:
        # weekday 1 = Monday
        return date.fromisocalendar(
            year,
            production_week,
            1
        )

    except ValueError:

        # Safety fallback for years where the requested
        # ISO week does not exist.
        #
        # This should normally not be reached for the
        # production weeks being planned.

        first_day = date(
            year,
            1,
            1
        )

        days_to_monday = (
            7 - first_day.weekday()
        ) % 7

        first_monday = (
            first_day
            + timedelta(
                days=days_to_monday
            )
        )

        return (
            first_monday
            + timedelta(
                days=(production_week - 1) * 7
            )
        )


# ============================================================
# DATE -> PRODUCTION WEEK
# ============================================================

def get_production_week(
    production_date
):

    """
    Automatically determine the calendar week.

    Monday-Sunday week.

    No hardcoded dates.
    """

    return production_date.isocalendar().week


# ============================================================
# BALANCED DAILY DISTRIBUTION
# ============================================================

def assign_balanced_dates(
    rows,
    week_start_date
):

    """
    Spread complete product rows across the week.

    Rules:
    - No daily capacity.
    - A product row is NEVER split.
    - External Document order is preserved.
    - Product order is preserved.
    - A document can continue into the next day.
    - Another document can start on the same day.
    - Try to keep daily quantities reasonably balanced.
    """

    if not rows:
        return {}

    count = len(rows)

    quantities = [
        float(item[1])
        for item in rows
    ]

    total_quantity = sum(
        quantities
    )

    # --------------------------------------------------------
    # Seven or fewer products.
    # --------------------------------------------------------

    if count <= 7:

        result = {}

        for index, item in enumerate(rows):

            row = item[0]

            result[row] = (
                week_start_date
                + timedelta(
                    days=index
                )
            )

        return result


    # ========================================================
    # Target quantity per day
    # ========================================================

    target = (
        total_quantity / 7
    )


    # ========================================================
    # Prefix sums
    # ========================================================

    prefix = [0.0]

    for quantity in quantities:

        prefix.append(
            prefix[-1] + quantity
        )


    # ========================================================
    # Dynamic programming
    # ========================================================

    INF = float("inf")

    dp = [
        [INF] * (count + 1)
        for _ in range(8)
    ]

    parent = [
        [None] * (count + 1)
        for _ in range(8)
    ]

    dp[0][0] = 0.0


    for day_count in range(
        1,
        8
    ):

        for end in range(
            day_count,
            count + 1
        ):

            for start in range(
                day_count - 1,
                end
            ):

                previous = dp[
                    day_count - 1
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
                    day_count
                ][end]:

                    dp[
                        day_count
                    ][end] = total_cost

                    parent[
                        day_count
                    ][end] = start


    # ========================================================
    # Recover groups
    # ========================================================

    boundaries = []

    end = count

    for day_count in range(
        7,
        0,
        -1
    ):

        start = parent[
            day_count
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

            row = rows[
                index
            ][0]

            result[row] = (
                assigned_date
            )


    return result


# ============================================================
# SORT OUTPUT
# ============================================================

def sort_planned_rows(
    worksheet,
    loom_settings,
    loom_col,
    week_col,
    date_col,
    row_sequence
):

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
            clean_loom(loom)
        ] = index


    rows = []


    # --------------------------------------------------------
    # Collect rows.
    # --------------------------------------------------------

    for row in range(
        2,
        worksheet.max_row + 1
    ):

        loom = clean_loom(
            worksheet.cell(
                row=row,
                column=loom_col
            ).value
        )


        week_value = worksheet.cell(
            row=row,
            column=week_col
        ).value


        try:

            week = int(
                str(week_value)
                .replace("WK-", "")
                .strip()
            )

        except Exception:

            week = 999


        date_value = worksheet.cell(
            row=row,
            column=date_col
        ).value


        try:

            sort_date = pd.to_datetime(
                date_value
            )

        except Exception:

            sort_date = pd.Timestamp.max


        sequence = row_sequence.get(
            row,
            999999999
        )


        starting_week = loom_settings.get(
            loom,
            {}
        ).get(
            "starting_week",
            1
        )


        # ----------------------------------------------------
        # Handles Week 52 -> Week 1.
        # ----------------------------------------------------

        week_rank = (
            week
            - starting_week
        ) % 52


        rows.append({

            "row":
                row,

            "loom_rank":
                loom_order.get(
                    loom,
                    999
                ),

            "week_rank":
                week_rank,

            "date":
                sort_date,

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
            x["row"]
        )
    )


    # ========================================================
    # Store row data
    # ========================================================

    max_col = worksheet.max_column

    row_data = {}


    for item in rows:

        original_row = item["row"]

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
    # Rewrite rows
    # ========================================================

    for new_row, item in enumerate(
        rows,
        start=2
    ):

        original_row = item["row"]


        for col in range(
            1,
            max_col + 1
        ):

            target = worksheet.cell(
                row=new_row,
                column=col
            )


            if isinstance(
                target,
                MergedCell
            ):

                continue


            source = row_data[
                original_row
            ][col - 1]


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
    # Restore merged cells
    # ========================================================

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
    "Automatically plan production using loom capacity, "
    "External Document No., production week and calendar date."
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
    "Enable/disable looms and edit weekly capacity "
    "and starting production week."
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
            ),

        "Starting Week":
            st.column_config.NumberColumn(
                "Starting Week",
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

        loom = clean_loom(
            row["Loom"]
        )

        capacity = float(
            row["Weekly Capacity"]
        )

        starting_week = int(
            row["Starting Week"]
        )


        if capacity <= 0:

            st.error(
                f"{loom}: Weekly Capacity must be greater than 0."
            )

            st.stop()


        if not (
            1 <= starting_week <= 52
        ):

            st.error(
                f"{loom}: Starting Week must be between 1 and 52."
            )

            st.stop()


        loom_settings[
            loom
        ] = {

            "use":
                bool(row["Use"]),

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

    DOCUMENT_COL = 2
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
    # GROUP BY LOOM + EXTERNAL DOCUMENT
    # ========================================================

    loom_documents = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        loom_documents[
            loom
        ] = {}


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


            if excel_loom != loom:
                continue


            document = clean_document(
                worksheet.cell(
                    row=row,
                    column=DOCUMENT_COL
                ).value
            )


            if document == "":

                document = (
                    f"ROW-{row}"
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
                row
            )


    # ========================================================
    # ASSIGN PRODUCTION WEEKS
    # ========================================================

    row_week = {}

    row_sequence = {}

    loom_summary = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        weekly_capacity = (
            settings["capacity"]
        )

        current_week = (
            settings["starting_week"]
        )

        current_load = 0.0

        sequence = 0

        rows_processed = 0

        weeks_used = []


        documents = loom_documents.get(
            loom,
            {}
        )


        # ----------------------------------------------------
        # External Document No. order is preserved.
        # ----------------------------------------------------

        for document, document_rows in (
            documents.items()
        ):

            for row in document_rows:

                quantity = get_quantity(
                    worksheet.cell(
                        row=row,
                        column=QUANTITY_COL
                    ).value
                )


                if quantity is None:
                    continue


                # ------------------------------------------------
                # Do NOT split a product.
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
                    row
                ] = current_week


                # ------------------------------------------------
                # Preserve order.
                # ------------------------------------------------

                row_sequence[
                    row
                ] = sequence

                sequence += 1


                # ------------------------------------------------
                # Update weekly quantity.
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
                settings["starting_week"],

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

    row_date = {}


    # --------------------------------------------------------
    # Planning year.
    # --------------------------------------------------------

    planning_year = date.today().year


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        week_rows = {}


        # ----------------------------------------------------
        # Group by production week.
        # ----------------------------------------------------

        for row, week in row_week.items():

            excel_loom = clean_loom(
                worksheet.cell(
                    row=row,
                    column=LOOM_COL
                ).value
            )


            if excel_loom != loom:
                continue


            quantity = get_quantity(
                worksheet.cell(
                    row=row,
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
                    row,
                    quantity,
                    row_sequence.get(
                        row,
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

            rows.sort(
                key=lambda x: x[2]
            )


            # ------------------------------------------------
            # Calculate Monday automatically.
            # ------------------------------------------------

            week_start_date = (
                get_week_start_date(
                    production_week,
                    planning_year
                )
            )


            # ------------------------------------------------
            # Spread products across Monday-Sunday.
            # ------------------------------------------------

            assignments = (
                assign_balanced_dates(
                    rows,
                    week_start_date
                )
            )


            for row, assigned_date in (
                assignments.items()
            ):

                row_date[
                    row
                ] = assigned_date


    # ========================================================
    # WRITE RESULTS
    # ========================================================

    for row, week in row_week.items():

        if row not in row_date:
            continue


        production_date = (
            row_date[row]
        )


        worksheet.cell(
            row=row,
            column=WEEK_COL
        ).value = (
            f"WK-{week}"
        )


        worksheet.cell(
            row=row,
            column=DATE_COL
        ).value = (
            production_date
        )


        worksheet.cell(
            row=row,
            column=DATE_COL
        ).number_format = (
            "DD-MM-YYYY"
        )


        worksheet.cell(
            row=row,
            column=DAY_COL
        ).value = (
            production_date.strftime(
                "%A"
            )
        )


    # ========================================================
    # SORT
    # ========================================================

    try:

        sort_planned_rows(

            worksheet=worksheet,

            loom_settings=loom_settings,

            loom_col=LOOM_COL,

            week_col=WEEK_COL,

            date_col=DATE_COL,

            row_sequence=row_sequence
        )

    except Exception as e:

        st.warning(
            "Planning completed, but final sorting "
            f"could not be completed: {e}"
        )


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


    for loom, info in (
        loom_summary.items()
    ):

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
