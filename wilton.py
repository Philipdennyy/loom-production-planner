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

    if week >= 52:
        return 1

    return week + 1


# ============================================================
# PRODUCTION WEEK → DATE
# ============================================================
#
# Company calendar:
#
# Week 1  = January 1 - January 7
# Week 2  = January 8 - January 14
# ...
# Week 52 = December 24 - December 30
#
# No manual date entry.
#
# ============================================================

def get_week_start_date(
    production_week,
    year
):

    production_week = int(
        production_week
    )

    return (
        date(year, 1, 1)
        + timedelta(
            days=(production_week - 1) * 7
        )
    )


# ============================================================
# BALANCE PRODUCTS ACROSS 7 DAYS
# ============================================================

def assign_balanced_dates(
    rows,
    week_start_date
):
    """
    Spread complete product rows over the seven days.

    IMPORTANT:

    - There is NO daily capacity.
    - A product row is NEVER split.
    - Product/document order is preserved.
    - The next product can start on the same day.
    - The distribution tries to avoid very uneven days.
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
    # If 7 or fewer products.
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
    # Dynamic programming
    # ========================================================

    ideal = total_quantity / 7

    prefix = [0.0]

    for quantity in quantities:

        prefix.append(
            prefix[-1] + quantity
        )


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
                    day_quantity - ideal
                )

                cost = (
                    difference
                    * difference
                )

                total_cost = (
                    previous + cost
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


        week_rank = (
            week - starting_week
        ) % 52


        rows.append({

            "row": row,

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
    # Sort
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


    # --------------------------------------------------------
    # Store row data.
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # Rewrite rows.
    # --------------------------------------------------------

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


    # --------------------------------------------------------
    # Restore merged cells.
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
    "Automatically allocate production by loom, "
    "weekly capacity, External Document No. and date."
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
    "Set the weekly capacity and starting production week "
    "for each loom."
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
# RESET SETTINGS
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
# MAIN PLANNING
# ============================================================

if run_button:

    # --------------------------------------------------------
    # File check.
    # --------------------------------------------------------

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
                f"{loom}: weekly capacity must "
                "be greater than 0."
            )

            st.stop()


        if not (
            1 <= starting_week <= 52
        ):

            st.error(
                f"{loom}: starting week must "
                "be between 1 and 52."
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
            f"Could not open Excel file: {e}"
        )

        st.stop()


    if not workbook.sheetnames:

        st.error(
            "The Excel file has no worksheets."
        )

        st.stop()


    worksheet = workbook[
        workbook.sheetnames[0]
    ]


    # ========================================================
    # INPUT COLUMNS
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

    WEEK_COL = 8
    DATE_COL = 9
    DAY_COL = 10


    # ========================================================
    # CREATE OUTPUT COLUMNS
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
    # GROUP ROWS BY LOOM + EXTERNAL DOCUMENT
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
    # PRODUCTION WEEK ASSIGNMENT
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
        # IMPORTANT:
        #
        # External Document No. order is preserved.
        #
        # If DOC-1 has multiple products, DOC-1 is processed
        # completely before DOC-2.
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
                # Weekly capacity.
                #
                # Product quantity is NOT split.
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
                # Preserve production order.
                # ------------------------------------------------

                row_sequence[
                    row
                ] = sequence

                sequence += 1


                # ------------------------------------------------
                # Update weekly load.
                # ------------------------------------------------

                current_load += quantity

                rows_processed += 1


                if current_week not in weeks_used:

                    weeks_used.append(
                        current_week
                    )


        # ----------------------------------------------------
        # Save summary.
        # ----------------------------------------------------

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
    # DATE ASSIGNMENT
    # ========================================================

    row_date = {}


    for loom, settings in loom_settings.items():

        if not settings["use"]:
            continue


        # ----------------------------------------------------
        # Group rows by production week.
        # ----------------------------------------------------

        week_rows = {}


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
            # The production year is determined from the
            # current planning file/year.
            #
            # Change this if your company plans multiple
            # calendar years at once.
            # ------------------------------------------------

            planning_year = date.today().year


            # ------------------------------------------------
            # Week 1 starts Jan 1.
            # ------------------------------------------------

            week_start_date = (
                get_week_start_date(
                    production_week,
                    planning_year
                )
            )


            # ------------------------------------------------
            # Spread products across seven days.
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
    # WRITE OUTPUT
    # ========================================================

    for row, week in row_week.items():

        if row not in row_date:
            continue


        production_date = (
            row_date[row]
        )


        # ----------------------------------------------------
        # Week.
        # ----------------------------------------------------

        worksheet.cell(
            row=row,
            column=WEEK_COL
        ).value = (
            f"WK-{week}"
        )


        # ----------------------------------------------------
        # Date.
        # ----------------------------------------------------

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


        # ----------------------------------------------------
        # Day.
        # ----------------------------------------------------

        worksheet.cell(
            row=row,
            column=DAY_COL
        ).value = (
            production_date.strftime(
                "%A"
            )
        )


    # ========================================================
    # SORT OUTPUT
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
            "Planning was completed, but the final "
            f"sorting could not be completed: {e}"
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
