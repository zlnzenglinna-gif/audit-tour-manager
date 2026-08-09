
from services.kostenerstellung_service import (
    export_kostenerstellung_excel,
)

from services.cost_service import (
    calculate_cost_summary,
    calculate_costs_by_interpreter,
)
from services.excel_service import (
    export_tourplan_excel,
)

from openpyxl.cell.cell import MergedCell

import hashlib
import json
import uuid

from datetime import datetime
from io import BytesIO
from pathlib import Path

import streamlit as st

from services.hotel_service import (
    find_historical_hotels,
    read_hotel_overview,
)

from services.tourplan_service import (
    read_tourplan_timeline,
)


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Audit Tour Manager",
    layout="wide",
)

st.markdown(
    """
    <style>
    .block-container {
        max-width: 1350px;
        padding-top: 1.5rem;
        padding-bottom: 2rem;
    }

    textarea {
        white-space: pre-wrap !important;
        overflow-wrap: break-word !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title(
    "Audit Tour Manager"
)

st.caption(
    "Tourplan bearbeiten, zusätzliche Zeilen ergänzen "
    "und Arbeitsstände speichern."
)


# ============================================================
# SESSION STATE
# ============================================================

if "added_rows" not in st.session_state:
    st.session_state.added_rows = []

if "meal_periods" not in st.session_state:
    st.session_state.meal_periods = []


# ============================================================
# HELPER: CALCULATE WEEKDAY
# ============================================================

def calculate_day(
    date_text,
):

    day_mapping = {
        0: "Mo",
        1: "Di",
        2: "Mi",
        3: "Do",
        4: "Fr",
        5: "Sa",
        6: "So",
    }

    try:

        parsed_date = datetime.strptime(
            str(date_text).strip(),
            "%d.%m.%Y",
        )

        return day_mapping[
            parsed_date.weekday()
        ]

    except (
        ValueError,
        TypeError,
    ):

        return ""


# ============================================================
# HELPER: TEXTAREA AUTO HEIGHT
# ============================================================

def calculate_textarea_height(
    text,
):

    if not text:
        return 100

    text = str(
        text
    )

    lines = text.splitlines()

    estimated_lines = 0

    for line in lines:

        wrapped_lines = max(
            1,
            (len(line) // 90) + 1,
        )

        estimated_lines += (
            wrapped_lines
        )

    height = (
        estimated_lines * 24
        + 50
    )

    return min(
        max(
            height,
            100,
        ),
        750,
    )


# ============================================================
# HELPER: FILE HASH
# ============================================================

def calculate_file_hash(
    file_bytes,
):

    return hashlib.sha256(
        file_bytes
    ).hexdigest()


# ============================================================
# HELPER: ADD MANUAL ROW
# ============================================================

def add_manual_row(
    after_excel_row,
):

    new_row = {
        "id": str(
            uuid.uuid4()
        ),
        "after_excel_row":
            after_excel_row,
    }

    st.session_state.added_rows.append(
        new_row
    )


# ============================================================
# HELPER: DELETE MANUAL ROW
# ============================================================

def delete_manual_row(
    row_id,
):

    st.session_state.added_rows = [
        row
        for row
        in st.session_state.added_rows
        if row["id"] != row_id
    ]


# ============================================================
# HELPER: ADD MEAL PERIOD
# ============================================================

def add_meal_period():

    new_period = {
        "id": str(
            uuid.uuid4()
        )
    }

    st.session_state.meal_periods.append(
        new_period
    )


# ============================================================
# HELPER: DELETE MEAL PERIOD
# ============================================================

def delete_meal_period(
    period_id,
):

    st.session_state.meal_periods = [
        period
        for period
        in st.session_state.meal_periods
        if period["id"] != period_id
    ]


# ============================================================
# HELPER: CLEAR OLD TOUR STATE
# ============================================================

def clear_tour_widget_state():

    keys_to_delete = []

    for key in list(
        st.session_state.keys()
    ):

        if (
            key.startswith("timeline_")
            or key.startswith("manual_")
            or key.startswith("meal_")
        ):

            keys_to_delete.append(
                key
            )

    for key in keys_to_delete:

        del st.session_state[
            key
        ]

    st.session_state.added_rows = []
    st.session_state.meal_periods = []


# ============================================================
# HELPER: PROGRESS FILE PATH
# ============================================================

def get_progress_path(
    tourplan_name,
    tourplan_hash,
):

    data_folder = Path(
        "data"
    )

    data_folder.mkdir(
        exist_ok=True
    )

    tourplan_stem = Path(
        tourplan_name
    ).stem

    short_hash = (
        tourplan_hash[:8]
    )

    progress_path = (
        data_folder
        / (
            f"{tourplan_stem}_"
            f"{short_hash}_progress.json"
        )
    )

    return progress_path


# ============================================================
# HELPER: SAVE PROGRESS
# ============================================================

def save_progress(
    tourplan_name,
    tourplan_hash,
    timeline_df,
):

    progress_data = {

        "tourplan_file":
            tourplan_name,

        "tourplan_hash":
            tourplan_hash,

        "original_rows":
            [],

        "manual_rows":
            [],

        "meal_periods":
            [],
    }


    # ========================================================
    # ORIGINAL EXCEL ROWS
    # ========================================================

    for _, row in (
        timeline_df.iterrows()
    ):

        excel_row = int(
            row["Excel Row"]
        )

        row_key = (
            f"timeline_{excel_row}"
        )

        original_row = {

            "excel_row":
                excel_row,

            "date":
                str(
                    row["Date"]
                ).strip(),

            "day":
                str(
                    row["Day"]
                ).strip(),

            "travel_plan":
                st.session_state.get(
                    f"{row_key}_travel_plan",
                    str(
                        row["Travel plan"]
                    ).strip(),
                ),

            "factory_hotel":
                st.session_state.get(
                    f"{row_key}_factory_hotel",
                    str(
                        row["Factory/Hotel"]
                    ).strip(),
                ),

            "address":
                st.session_state.get(
                    f"{row_key}_address",
                    str(
                        row["Address"]
                    ).strip(),
                ),

            "distance":
                st.session_state.get(
                    f"{row_key}_distance",
                    None,
                ),

            "drive_time":
                st.session_state.get(
                    f"{row_key}_drive_time",
                    None,
                ),

            "hotel_price":
                st.session_state.get(
                    f"{row_key}_hotel_price",
                    None,
                ),

            "payment_status":
                st.session_state.get(
                    f"{row_key}_payment_status",
                    "",
                ),

            "breakfast":
                st.session_state.get(
                    f"{row_key}_breakfast",
                    "",
                ),
        }

        progress_data[
            "original_rows"
        ].append(
            original_row
        )


    # ========================================================
    # MANUAL ROWS
    # ========================================================

    for manual_row in (
        st.session_state.added_rows
    ):

        row_id = (
            manual_row["id"]
        )

        manual_data = {

            "id":
                row_id,

            "after_excel_row":
                manual_row[
                    "after_excel_row"
                ],

            "date":
                st.session_state.get(
                    f"manual_{row_id}_date",
                    "",
                ),

            "travel_plan":
                st.session_state.get(
                    f"manual_{row_id}_travel_plan",
                    "",
                ),

            "factory_hotel":
                st.session_state.get(
                    f"manual_{row_id}_factory_hotel",
                    "",
                ),

            "address":
                st.session_state.get(
                    f"manual_{row_id}_address",
                    "",
                ),

            "distance":
                st.session_state.get(
                    f"manual_{row_id}_distance",
                    None,
                ),

            "drive_time":
                st.session_state.get(
                    f"manual_{row_id}_drive_time",
                    None,
                ),

            "hotel_price":
                st.session_state.get(
                    f"manual_{row_id}_hotel_price",
                    None,
                ),

            "payment_status":
                st.session_state.get(
                    f"manual_{row_id}_payment_status",
                    "",
                ),

            "breakfast":
                st.session_state.get(
                    f"manual_{row_id}_breakfast",
                    "",
                ),
        }

        progress_data[
            "manual_rows"
        ].append(
            manual_data
        )


    # ========================================================
    # MEAL PERIODS
    # ========================================================

    for period in (
        st.session_state.meal_periods
    ):

        period_id = (
            period["id"]
        )

        start_date = (
            st.session_state.get(
                f"meal_{period_id}_start_date"
            )
        )

        end_date = (
            st.session_state.get(
                f"meal_{period_id}_end_date"
            )
        )

        meal_period_data = {

            "id":
                period_id,

            "start_date":
                (
                    start_date.isoformat()
                    if start_date
                    else None
                ),

            "end_date":
                (
                    end_date.isoformat()
                    if end_date
                    else None
                ),

          
            "interpreter_name":
                st.session_state.get(
                    f"meal_{period_id}_interpreter_name",
                    "",
                ),

            "auditor_count":
                st.session_state.get(
                    f"meal_{period_id}_auditor_count",
                    1,
                ),

            "xia_included":
                st.session_state.get(
                    f"meal_{period_id}_xia_included",
                    False,
                ),
        }

        progress_data[
            "meal_periods"
        ].append(
            meal_period_data
        )


    # ========================================================
    # SAVE JSON
    # ========================================================

    save_path = get_progress_path(
        tourplan_name=(
            tourplan_name
        ),
        tourplan_hash=(
            tourplan_hash
        ),
    )

    with open(
        save_path,
        "w",
        encoding="utf-8",
    ) as file:

        json.dump(
            progress_data,
            file,
            ensure_ascii=False,
            indent=4,
        )

    return save_path


# ============================================================
# HELPER: LOAD PROGRESS
# ============================================================

def load_progress(
    tourplan_name,
    tourplan_hash,
):

    progress_path = get_progress_path(
        tourplan_name=(
            tourplan_name
        ),
        tourplan_hash=(
            tourplan_hash
        ),
    )

    if not progress_path.exists():

        return None

    try:

        with open(
            progress_path,
            "r",
            encoding="utf-8",
        ) as file:

            return json.load(
                file
            )

    except (
        json.JSONDecodeError,
        OSError,
    ):

        return None


# ============================================================
# HELPER: RESTORE PROGRESS
# ============================================================

def restore_progress_to_session(
    progress_data,
):

    if not progress_data:
        return


    # ========================================================
    # ORIGINAL ROWS
    # ========================================================

    for saved_row in progress_data.get(
        "original_rows",
        [],
    ):

        excel_row = saved_row.get(
            "excel_row"
        )

        if excel_row is None:
            continue

        row_key = (
            f"timeline_{excel_row}"
        )

        fields = [
            "travel_plan",
            "factory_hotel",
            "address",
            "distance",
            "drive_time",
            "hotel_price",
            "payment_status",
            "breakfast",
        ]

        for field in fields:

            value = saved_row.get(
                field
            )

            if value is not None:

                st.session_state[
                    f"{row_key}_{field}"
                ] = value


    # ========================================================
    # MANUAL ROWS
    # ========================================================

    st.session_state.added_rows = []

    for saved_row in progress_data.get(
        "manual_rows",
        [],
    ):

        row_id = saved_row.get(
            "id"
        )

        if not row_id:

            row_id = str(
                uuid.uuid4()
            )

        after_excel_row = saved_row.get(
            "after_excel_row"
        )

        st.session_state.added_rows.append(
            {
                "id": row_id,
                "after_excel_row": after_excel_row,
            }
        )

        fields = [
            "date",
            "travel_plan",
            "factory_hotel",
            "address",
            "distance",
            "drive_time",
            "hotel_price",
            "payment_status",
            "breakfast",
        ]

        for field in fields:

            value = saved_row.get(
                field
            )

            if value is not None:

                st.session_state[
                    f"manual_{row_id}_{field}"
                ] = value


    # ========================================================
    # KOSTENERSTELLUNG / DOLMETSCHER PERIODS
    # ========================================================

    st.session_state.meal_periods = []

    for saved_period in progress_data.get(
        "meal_periods",
        [],
    ):

        period_id = saved_period.get(
            "id"
        )

        if not period_id:

            period_id = str(
                uuid.uuid4()
            )


        # ----------------------------------------------------
        # PERIOD STRUCTURE
        # ----------------------------------------------------

        st.session_state.meal_periods.append(
            {
                "id": period_id,
            }
        )


        # ----------------------------------------------------
        # INTERPRETER NAME
        # ----------------------------------------------------

        st.session_state[
            f"meal_{period_id}_interpreter_name"
        ] = saved_period.get(
            "interpreter_name",
            "",
        )


        # ----------------------------------------------------
        # START DATE
        # ----------------------------------------------------

        start_date = saved_period.get(
            "start_date"
        )

        if start_date:

            st.session_state[
                f"meal_{period_id}_start_date"
            ] = datetime.fromisoformat(
                start_date
            ).date()


        # ----------------------------------------------------
        # END DATE
        # ----------------------------------------------------

        end_date = saved_period.get(
            "end_date"
        )

        if end_date:

            st.session_state[
                f"meal_{period_id}_end_date"
            ] = datetime.fromisoformat(
                end_date
            ).date()


        # ----------------------------------------------------
        # AUDITORS
        # ----------------------------------------------------

        st.session_state[
            f"meal_{period_id}_auditor_count"
        ] = saved_period.get(
            "auditor_count",
            1,
        )


        # ----------------------------------------------------
        # XIA
        # ----------------------------------------------------

        st.session_state[
            f"meal_{period_id}_xia_included"
        ] = saved_period.get(
            "xia_included",
            False,
        )

# ============================================================
# HELPER: MANUAL ROW UI
# ============================================================

def render_manual_row(
    manual_row,
    default_date="",
):

    row_id = (
        manual_row["id"]
    )

    with st.container(
        border=True
    ):

        # ====================================================
        # DATE / DAY
        # ====================================================

        date_col, day_col, info_col = (
            st.columns(
                [1, 1, 4]
            )
        )

        with date_col:

            date_value = st.text_input(
                "Date",
                value=default_date,
                key=f"manual_{row_id}_date",
                placeholder="DD.MM.YYYY",
            )

        calculated_day = (
            calculate_day(
                date_value
            )
        )

        with day_col:

            st.text_input(
                "Day",
                value=calculated_day,
                disabled=True,
                key=(
                    f"manual_{row_id}_"
                    f"day_{calculated_day}"
                ),
            )

        with info_col:

            st.caption(
                "Manuell hinzugefügte Zeile"
            )


        # ====================================================
        # TRAVEL PLAN
        # ====================================================

        manual_travel_key = (
            f"manual_{row_id}_travel_plan"
        )

        manual_travel_value = (
            st.session_state.get(
                manual_travel_key,
                "",
            )
        )

        st.text_area(
            "Travel plan",
            key=manual_travel_key,
            placeholder=(
                "Travel plan eingeben"
            ),
            height=(
                calculate_textarea_height(
                    manual_travel_value
                )
            ),
        )


        # ====================================================
        # FACTORY / HOTEL
        # ====================================================

        st.text_input(
            "Factory/Hotel",
            key=(
                f"manual_{row_id}_"
                "factory_hotel"
            ),
            placeholder=(
                "Hotel oder andere Information eingeben"
            ),
        )


        # ====================================================
        # ADDRESS
        # ====================================================

        st.text_input(
            "Address",
            key=(
                f"manual_{row_id}_address"
            ),
            placeholder=(
                "Adresse eingeben"
            ),
        )


        # ====================================================
        # DISTANCE / DRIVE TIME
        # ====================================================

        distance_col, time_col = (
            st.columns(
                2
            )
        )

        with distance_col:

            st.number_input(
                "Distance (km)",
                min_value=0.0,
                value=None,
                step=1.0,
                key=(
                    f"manual_{row_id}_"
                    "distance"
                ),
            )

        with time_col:

            st.number_input(
                "Drive time (min)",
                min_value=0,
                value=None,
                step=5,
                key=(
                    f"manual_{row_id}_"
                    "drive_time"
                ),
            )


        # ====================================================
        # PRICE / PAYMENT
        # ====================================================

        price_col, payment_col = (
            st.columns(
                2
            )
        )

        with price_col:

            st.number_input(
                "Preis (EUR)",
                min_value=0.0,
                value=None,
                step=10.0,
                key=(
                    f"manual_{row_id}_"
                    "hotel_price"
                ),
            )

        with payment_col:

            st.selectbox(
                "Zahlungsstatus",
                options=[
                    "",
                    "bezahlt",
                    "vor Ort",
                ],
                key=(
                    f"manual_{row_id}_"
                    "payment_status"
                ),
            )


        # ====================================================
        # BREAKFAST
        # ====================================================

        st.selectbox(
            "Frühstück",
            options=[
                "",
                "inklusive",
                "nicht inklusive",
            ],
            key=(
                f"manual_{row_id}_"
                "breakfast"
            ),
        )


        # ====================================================
        # DELETE
        # ====================================================

        if st.button(
            "🗑 Delete Row",
            key=f"delete_{row_id}",
        ):

            delete_manual_row(
                row_id
            )

            st.rerun()


# ============================================================
# FILE UPLOAD
# ============================================================

upload_col1, upload_col2, upload_col3= (
    st.columns(
        3
    )
)

with upload_col1:

    tourplan_file = (
        st.file_uploader(
            "Tourplan hochladen",
            type=["xlsx"],
        )
    )

with upload_col2:

    hotel_file = (
        st.file_uploader(
            "Hotelübersicht hochladen",
            type=["xlsx"],
        )
    )
with upload_col3:

    kostenerstellung_file = (
        st.file_uploader(
            "Kostenerstellung hochladen",
            type=["xlsx"],
        )
    )


# ============================================================
# CHECK FILES
# ============================================================

if (
    tourplan_file is None
    or hotel_file is None
    or kostenerstellung_file is None
):

    st.info(
        "Bitte Tourplan, Hotelübersicht "
        "und Kostenerstellung hochladen."
    )

    st.stop()


# ============================================================
# MAIN
# ============================================================

try:

    # ========================================================
    # FILE BYTES
    # ========================================================

    tourplan_bytes = (
        tourplan_file.getvalue()
    )

    hotel_bytes = (
        hotel_file.getvalue()
    )

    kostenerstellung_bytes = (
    kostenerstellung_file.getvalue()
    )


    # ========================================================
    # CURRENT EXCEL VERSION
    # ========================================================

    tourplan_hash = (
        calculate_file_hash(
            tourplan_bytes
        )
    )

    current_version = (
        f"{tourplan_file.name}:"
        f"{tourplan_hash}"
    )


    # ========================================================
    # READ CURRENT EXCEL
    # ========================================================

    timeline_df = (
        read_tourplan_timeline(
            BytesIO(
                tourplan_bytes
            )
        )
    )

    hotel_df = (
        read_hotel_overview(
            BytesIO(
                hotel_bytes
            )
        )
    )


    # ========================================================
    # LOAD PROGRESS
    # ========================================================

    progress_data = (
        load_progress(
            tourplan_name=(
                tourplan_file.name
            ),
            tourplan_hash=(
                tourplan_hash
            ),
        )
    )

    previous_version = (
        st.session_state.get(
            "current_tourplan_version"
        )
    )


    # ========================================================
    # NEW EXCEL VERSION
    # ========================================================

    if (
        previous_version
        != current_version
    ):

        clear_tour_widget_state()

        if progress_data:

            restore_progress_to_session(
                progress_data
            )

        st.session_state[
            "current_tourplan_version"
        ] = current_version

        st.session_state[
            "progress_restored"
        ] = bool(
            progress_data
        )

        st.rerun()


    # ========================================================
    # CHECK TIMELINE
    # ========================================================

    if timeline_df.empty:

        st.warning(
            "Keine Tourplan-Zeilen gefunden."
        )

        st.stop()


    st.success(
        f"{len(timeline_df)} Tourplan-Zeilen und "
        f"{len(hotel_df)} Hotel-Einträge wurden geladen."
    )


    if st.session_state.get(
        "progress_restored",
        False,
    ):

        st.info(
            "Gespeicherter Arbeitsstand für diese "
            "Excel-Version wurde wiederhergestellt."
        )


    st.subheader(
        "Tour Timeline"
    )


    # ========================================================
    # TIMELINE
    # ========================================================

    for _, row in (
        timeline_df.iterrows()
    ):

        date_value = str(
            row["Date"]
        ).strip()

        day_value = str(
            row["Day"]
        ).strip()

        travel_plan_value = str(
            row["Travel plan"]
        ).strip()

        factory_hotel_value = str(
            row["Factory/Hotel"]
        ).strip()

        address_value = str(
            row["Address"]
        ).strip()

        is_factory = bool(
            row["Is Factory"]
        )

        excel_row = int(
            row["Excel Row"]
        )

        row_key = (
            f"timeline_{excel_row}"
        )

        current_travel_plan = (
            st.session_state.get(
                f"{row_key}_travel_plan",
                travel_plan_value,
            )
        )


        # ====================================================
        # ORIGINAL CARD
        # ====================================================

        with st.container(
            border=True
        ):

            header_col1, header_col2 = (
                st.columns(
                    [1, 5]
                )
            )

            with header_col1:

                st.markdown(
                    f"### {date_value}"
                )

            with header_col2:

                if day_value:

                    st.caption(
                        day_value
                    )


            # =================================================
            # FACTORY CARD
            # =================================================

            if is_factory:

                left_col, right_col = (
                    st.columns(
                        [1.15, 0.85],
                        gap="medium",
                    )
                )

                with left_col:

                    st.text_area(
                        "Travel plan",
                        value=(
                            travel_plan_value
                        ),
                        key=(
                            f"{row_key}_"
                            "travel_plan"
                        ),
                        height=(
                            calculate_textarea_height(
                                current_travel_plan
                            )
                        ),
                    )

                    current_factory_hotel = (
                        st.text_input(
                            "Factory/Hotel",
                            value=(
                                factory_hotel_value
                            ),
                            key=(
                                f"{row_key}_"
                                "factory_hotel"
                            ),
                        )
                    )

                    st.text_input(
                        "Address",
                        value=(
                            address_value
                        ),
                        key=(
                            f"{row_key}_"
                            "address"
                        ),
                    )

                    distance_col, time_col = (
                        st.columns(
                            2
                        )
                    )

                    with distance_col:

                        st.number_input(
                            "Distance (km)",
                            min_value=0.0,
                            value=None,
                            step=1.0,
                            key=(
                                f"{row_key}_"
                                "distance"
                            ),
                        )

                    with time_col:

                        st.number_input(
                            "Drive time (min)",
                            min_value=0,
                            value=None,
                            step=5,
                            key=(
                                f"{row_key}_"
                                "drive_time"
                            ),
                        )


                # =============================================
                # HISTORICAL HOTELS
                # =============================================

                with right_col:

                    st.markdown(
                        "### Historische Hotels"
                    )

                    matching_hotels = (
                        find_historical_hotels(
                            current_factory_hotel,
                            hotel_df,
                        )
                    )

                    if (
                        matching_hotels.empty
                    ):

                        st.info(
                            "Keine passenden Hotels gefunden."
                        )

                    else:

                        st.caption(
                            f"{len(matching_hotels)} "
                            "mögliche Hotel-Einträge"
                        )

                        display_columns = [
                            "Kunde",
                            "Hotel",
                            "Kommentar",
                        ]

                        available_columns = [
                            column
                            for column
                            in display_columns
                            if column
                            in matching_hotels.columns
                        ]

                        st.dataframe(
                            matching_hotels[
                                available_columns
                            ],
                            width="stretch",
                            hide_index=True,
                        )


            # =================================================
            # NORMAL EVENT
            # =================================================

            else:

                st.text_area(
                    "Travel plan",
                    value=(
                        travel_plan_value
                    ),
                    key=(
                        f"{row_key}_"
                        "travel_plan"
                    ),
                    height=(
                        calculate_textarea_height(
                            current_travel_plan
                        )
                    ),
                )

                st.text_input(
                    "Factory/Hotel",
                    value=(
                        factory_hotel_value
                    ),
                    key=(
                        f"{row_key}_"
                        "factory_hotel"
                    ),
                )

                st.text_input(
                    "Address",
                    value=(
                        address_value
                    ),
                    key=(
                        f"{row_key}_"
                        "address"
                    ),
                )

                distance_col, time_col = (
                    st.columns(
                        2
                    )
                )

                with distance_col:

                    st.number_input(
                        "Distance (km)",
                        min_value=0.0,
                        value=None,
                        step=1.0,
                        key=(
                            f"{row_key}_"
                            "distance"
                        ),
                    )

                with time_col:

                    st.number_input(
                        "Drive time (min)",
                        min_value=0,
                        value=None,
                        step=5,
                        key=(
                            f"{row_key}_"
                            "drive_time"
                        ),
                    )

                price_col, payment_col = (
                    st.columns(
                        2
                    )
                )

                with price_col:

                    st.number_input(
                        "Preis (EUR)",
                        min_value=0.0,
                        value=None,
                        step=10.0,
                        key=(
                            f"{row_key}_"
                            "hotel_price"
                        ),
                    )

                with payment_col:

                    st.selectbox(
                        "Zahlungsstatus",
                        options=[
                            "",
                            "bezahlt",
                            "vor Ort",
                        ],
                        key=(
                            f"{row_key}_"
                            "payment_status"
                        ),
                    )

                st.selectbox(
                    "Frühstück",
                    options=[
                        "",
                        "inklusive",
                        "nicht inklusive",
                    ],
                    key=(
                        f"{row_key}_"
                        "breakfast"
                    ),
                )


        # ====================================================
        # MANUAL ROWS
        # ====================================================

        manual_rows_here = [
            manual_row
            for manual_row
            in st.session_state.added_rows
            if (
                manual_row[
                    "after_excel_row"
                ]
                == excel_row
            )
        ]

        for manual_row in (
            manual_rows_here
        ):

            render_manual_row(
                manual_row=(
                    manual_row
                ),
                default_date=(
                    date_value
                ),
            )


        # ====================================================
        # ADD ROW
        # ====================================================

        if st.button(
            "＋ Add Row",
            key=(
                f"add_after_"
                f"{excel_row}"
            ),
        ):

            add_manual_row(
                after_excel_row=(
                    excel_row
                )
            )

            st.rerun()



    # ========================================================
    # ESSEN
    # ========================================================

    st.markdown(
    "### Kostenerstellung / Dolmetscher-Zeiträume"
)

    st.caption(
    "Für jeden Dolmetscher kann ein eigener Zeitraum "
    "innerhalb der Tour definiert werden."
)

    st.caption(
        "Die Personenkonstellation kann sich "
        "innerhalb einer Tour ändern."
    )


    # ========================================================
    # MEAL PERIOD CARDS
    # ========================================================

    for (
        period_number,
        period,
    ) in enumerate(
        st.session_state.meal_periods,
        start=1,
    ):

        period_id = (
            period["id"]
        )

        with st.container(
            border=True
        ):

            header_col, delete_col = (
                st.columns(
                    [5, 1]
                )
            )

            with header_col:

                st.markdown(
                    f"#### Zeitraum "
                    f"{period_number}"
                )
                st.text_input(
                   "Dolmetscher Name",
                    key=(
                    f"meal_{period_id}_"
                    "interpreter_name"
            ),
                    placeholder=(
                   "Nachname, Vorname"
    ),
)
            with delete_col:

                if st.button(
                    "🗑",
                    key=(
                        f"delete_meal_"
                        f"{period_id}"
                    ),
                    help=(
                        "Zeitraum löschen"
                    ),
                ):

                    delete_meal_period(
                        period_id
                    )

                    st.rerun()


            # =================================================
            # DATE RANGE
            # =================================================

            date_col1, date_col2 = (
                st.columns(
                    2
                )
            )

            with date_col1:

                start_date = (
                    st.date_input(
                        "Von",
                        key=(
                            f"meal_{period_id}_"
                            "start_date"
                        ),
                        format=(
                            "DD.MM.YYYY"
                        ),
                    )
                )

            with date_col2:

                end_date = (
                    st.date_input(
                        "Bis",
                        key=(
                            f"meal_{period_id}_"
                            "end_date"
                        ),
                        format=(
                            "DD.MM.YYYY"
                        ),
                    )
                )


            # =================================================
            # PEOPLE
            # =================================================

            st.number_input(
               "Auditoren",
               min_value=0,
               value=1,
               step=1,
               key=(
                    f"meal_{period_id}_"
                    "auditor_count"
                ),
)


            # =================================================
            # XIA
            # =================================================

            st.checkbox(
                "Ist ein Auditor aus dem eigenen Unternehmen dabei?",
                key=(
                    f"meal_{period_id}_"
                    "xia_included"
                ),
            )


            # =================================================
            # DAYS PREVIEW
            # =================================================

            if (
                start_date
                and end_date
            ):

                if (
                    end_date
                    >= start_date
                ):

                    days = (
                        end_date
                        - start_date
                    ).days + 1

                    st.caption(
                        f"{days} Kalendertage"
                    )

                else:

                    st.warning(
                        "Das Bis-Datum liegt "
                        "vor dem Von-Datum."
                    )

    # ========================================================
    # ADD PERIOD
    # ========================================================

    if st.button(
        "＋ Zeitraum hinzufügen",
        key="add_meal_period",
    ):

        add_meal_period()

        st.rerun()


    # ========================================================
    # SAVE PROGRESS
    # ========================================================

    st.divider()

    st.subheader(
        "Arbeitsstand"
    )

    if st.button(
        "💾 Save Progress"
    ):

        save_path = save_progress(
            tourplan_name=tourplan_file.name,
            tourplan_hash=tourplan_hash,
            timeline_df=timeline_df,
        )

        st.success(
            f"Gespeichert: {save_path}"
        )

        st.caption(
            "Dieser Arbeitsstand wird nur "
            "für exakt diese Excel-Version "
            "wiederhergestellt."
        )


    # ========================================================
    # EXPORT TOURPLAN
    # ========================================================

    st.divider()

    st.subheader(
        "Tourplan Export"
    )

    try:

        exported_tourplan = export_tourplan_excel(
            original_excel_bytes=tourplan_bytes,
            timeline_df=timeline_df,
            session_state=st.session_state,
        )

        export_filename = (
            f"{Path(tourplan_file.name).stem}"
            "_updated.xlsx"
        )

        st.download_button(
            label="📥 Export Tourplan.xlsx",
            data=exported_tourplan.getvalue(),
            file_name=export_filename,
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
        )

    except Exception as export_error:

        st.error(
            "Fehler beim Exportieren des Tourplans: "
            f"{type(export_error).__name__}: "
            f"{export_error}"
        )
 

    # ========================================================
    # KOSTENERSTELLUNG EXPORT
    # ========================================================

    st.divider()

    st.subheader(
        "Kostenerstellung Export"
    )


    # ========================================================
    # CALCULATE COSTS BY INTERPRETER
    # ========================================================

    interpreter_costs = calculate_costs_by_interpreter(
        timeline_df=timeline_df,
        session_state=st.session_state,
        added_rows=st.session_state.added_rows,
        meal_periods=st.session_state.meal_periods,
    )


    if not st.session_state.meal_periods:

        st.info(
            "Bitte zuerst mindestens einen "
            "Dolmetscher-Zeitraum hinzufügen."
        )

    else:

        # ====================================================
        # EVERY PERIOD / INTERPRETER
        # ====================================================

        for period_number, period in enumerate(
            st.session_state.meal_periods,
            start=1,
        ):

            period_id = period["id"]

            interpreter_name = (
                st.session_state.get(
                    f"meal_{period_id}_interpreter_name",
                    "",
                )
                or ""
            ).strip()

            start_date = st.session_state.get(
                f"meal_{period_id}_start_date",
                None,
            )

            end_date = st.session_state.get(
                f"meal_{period_id}_end_date",
                None,
            )


            # ------------------------------------------------
            # VALIDATE
            # ------------------------------------------------

            if not interpreter_name:

                st.warning(
                    f"Zeitraum {period_number}: "
                    "Dolmetscher Name fehlt."
                )

                continue


            if (
                start_date is None
                or end_date is None
            ):

                st.warning(
                    f"Zeitraum {period_number}: "
                    "Von/Bis fehlt."
                )

                continue


            # ------------------------------------------------
            # GET COST DATA
            # ------------------------------------------------

            interpreter_cost = (
                interpreter_costs.get(
                    interpreter_name
                )
            )


            if interpreter_cost is None:

                st.warning(
                    f"Keine Kostendaten für "
                    f"{interpreter_name} gefunden."
                )

                continue


            # ------------------------------------------------
            # EXPORT CARD
            # ------------------------------------------------

            with st.container(
                border=True
            ):

                st.markdown(
                    f"### {interpreter_name}"
                )

                st.caption(
                    f"{start_date.strftime('%d.%m.%Y')} "
                    f"– "
                    f"{end_date.strftime('%d.%m.%Y')}"
                )


                try:

                    kostenerstellung_output = (
                        export_kostenerstellung_excel(
                            template_excel_bytes=(
                                kostenerstellung_bytes
                            ),
                            interpreter_name=(
                                interpreter_name
                            ),
                            start_date=(
                                start_date
                            ),
                            end_date=(
                                end_date
                            ),
                            interpreter_cost=(
                                interpreter_cost
                            ),
                            meal_excel_formula=(
                                interpreter_cost.get(
                                    "meal_excel_formula",
                                    "=0",
                                )
                            ),
                            sonstiges_value=0.0,
                        )
                    )


                    safe_name = (
                        interpreter_name
                        .strip()
                        .replace(" ", "_")
                        .replace("/", "_")
                        .replace("\\", "_")
                    )


                    filename = (
                        f"Kostenerstellung_"
                        f"{safe_name}_"
                        f"{start_date.strftime('%Y%m%d')}_"
                        f"{end_date.strftime('%Y%m%d')}"
                        ".xlsx"
                    )


                    st.download_button(
                        label=(
                            "📥 Kostenerstellung herunterladen"
                        ),
                        data=(
                            kostenerstellung_output.getvalue()
                        ),
                        file_name=(
                            filename
                        ),
                        mime=(
                            "application/"
                            "vnd.openxmlformats-officedocument."
                            "spreadsheetml.sheet"
                        ),
                        key=(
                            f"download_cost_"
                            f"{period_id}"
                        ),
                    )


                except Exception as cost_export_error:

                    st.error(
                        "Fehler bei der Kostenerstellung: "
                        f"{type(cost_export_error).__name__}: "
                        f"{cost_export_error}"
                    )

# ============================================================
# ERROR HANDLING
# ============================================================

except Exception as error:

    st.error(
        f"{type(error).__name__}: "
        f"{error}"
    )