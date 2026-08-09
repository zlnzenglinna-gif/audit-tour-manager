# services/cost_service.py


def safe_float(value):
    """
    Convert Streamlit input values safely to float.
    Empty / None / invalid values become 0.
    """

    if value is None:
        return 0.0

    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def calculate_total_distance(
    timeline_df,
    session_state,
    added_rows,
):
    """
    Add all Distance (km) values:
    - original Tourplan rows
    - manually added rows
    """

    total_distance = 0.0

    # --------------------------------------------------------
    # ORIGINAL TOURPLAN ROWS
    # --------------------------------------------------------

    for _, row in timeline_df.iterrows():

        excel_row = int(
            row["Excel Row"]
        )

        distance = session_state.get(
            f"timeline_{excel_row}_distance",
            None,
        )

        total_distance += safe_float(
            distance
        )

    # --------------------------------------------------------
    # MANUAL ROWS
    # --------------------------------------------------------

    for manual_row in added_rows:

        row_id = manual_row["id"]

        distance = session_state.get(
            f"manual_{row_id}_distance",
            None,
        )

        total_distance += safe_float(
            distance
        )

    return round(
        total_distance,
        2,
    )


def calculate_fuel_cost(
    total_distance,
    consumption_per_100km=7.0,
    fuel_price=1.8,
):
    """
    Fuel formula:

    (total km / 100)
    × 7 L
    × 1.80 EUR
    """

    fuel_cost = (
        total_distance
        / 100
        * consumption_per_100km
        * fuel_price
    )

    return round(
        fuel_cost,
        2,
    )


def collect_hotels(
    timeline_df,
    session_state,
    added_rows,
):
    """
    Collect hotel information from:
    - original Tourplan rows
    - manually added rows
    """

    hotels = []

    # --------------------------------------------------------
    # ORIGINAL ROWS
    # --------------------------------------------------------

    for _, row in timeline_df.iterrows():

        excel_row = int(
            row["Excel Row"]
        )

        row_key = (
            f"timeline_{excel_row}"
        )

        hotel_price = safe_float(
            session_state.get(
                f"{row_key}_hotel_price",
                None,
            )
        )

        payment_status = (
            session_state.get(
                f"{row_key}_payment_status",
                "",
            )
            or ""
        )

        breakfast = (
            session_state.get(
                f"{row_key}_breakfast",
                "",
            )
            or ""
        )

        factory_hotel = (
            session_state.get(
                f"{row_key}_factory_hotel",
                str(
                    row["Factory/Hotel"]
                ).strip(),
            )
            or ""
        )

        # Only consider it a hotel-cost entry
        # when a price has actually been entered.
        if hotel_price > 0:

            hotels.append(
                {
                    "date": str(
                        row["Date"]
                    ).strip(),

                    "hotel": str(
                        factory_hotel
                    ).strip(),

                    "price": hotel_price,

                    "payment_status":
                        payment_status,

                    "breakfast":
                        breakfast,
                }
            )

    # --------------------------------------------------------
    # MANUAL ROWS
    # --------------------------------------------------------

    for manual_row in added_rows:

        row_id = (
            manual_row["id"]
        )

        hotel_price = safe_float(
            session_state.get(
                f"manual_{row_id}_hotel_price",
                None,
            )
        )

        if hotel_price <= 0:
            continue

        hotels.append(
            {
                "date":
                    session_state.get(
                        f"manual_{row_id}_date",
                        "",
                    ),

                "hotel":
                    session_state.get(
                        f"manual_{row_id}_factory_hotel",
                        "",
                    ),

                "price":
                    hotel_price,

                "payment_status":
                    session_state.get(
                        f"manual_{row_id}_payment_status",
                        "",
                    )
                    or "",

                "breakfast":
                    session_state.get(
                        f"manual_{row_id}_breakfast",
                        "",
                    )
                    or "",
            }
        )

    return hotels


def calculate_hotel_totals(
    hotels,
):
    """
    Calculate:
    - all hotel costs
    - vor Ort hotel costs
    - already paid hotel costs
    """

    total_hotels = 0.0
    vor_ort_total = 0.0
    bezahlt_total = 0.0

    for hotel in hotels:

        price = safe_float(
            hotel.get(
                "price"
            )
        )

        status = str(
            hotel.get(
                "payment_status",
                ""
            )
        ).strip().lower()

        total_hotels += price

        if status == "vor ort":

            vor_ort_total += price

        elif status == "bezahlt":

            bezahlt_total += price

    return {
        "total_hotels":
            round(total_hotels, 2),

        "vor_ort_total":
            round(vor_ort_total, 2),

        "bezahlt_total":
            round(bezahlt_total, 2),
    }

def calculate_meal_period_cost(
    start_date,
    end_date,
    interpreter_count,
    auditor_count,
    xia_included,
):
    """
    Calculate meal cost for one date period.

    Current temporary rule:
    - standard daily meal rate = 7 + 15 + 25 = 47 EUR
    - interpreter is always included
    - auditors are added according to auditor_count
    - Xia can be treated separately later if needed

    For now, this function prepares the period structure.
    """

    if (
        start_date is None
        or end_date is None
        or end_date < start_date
    ):
        return {
            "days": 0,
            "cost": 0.0,
        }

    days = (
        end_date - start_date
    ).days + 1

    interpreter_count = int(
        interpreter_count or 0
    )

    auditor_count = int(
        auditor_count or 0
    )

    # Temporary standard daily amount
    daily_rate = (
        7
        + 15
        + 25
    )

    total_people = (
        interpreter_count
        + auditor_count
    )

    meal_cost = (
        days
        * total_people
        * daily_rate
    )

    return {
        "days": days,
        "cost": round(
            meal_cost,
            2,
        ),
        "xia_included": bool(
            xia_included
        ),
    }


def calculate_meal_total(
    session_state,
    meal_periods,
):
    """
    Calculate all Essen periods.
    """

    periods = []
    total_meal_cost = 0.0

    for period in meal_periods:

        period_id = period["id"]

        start_date = session_state.get(
            f"meal_{period_id}_start_date"
        )

        end_date = session_state.get(
            f"meal_{period_id}_end_date"
        )

        interpreter_count = (
            session_state.get(
                f"meal_{period_id}_interpreter_count",
                1,
            )
        )

        auditor_count = (
            session_state.get(
                f"meal_{period_id}_auditor_count",
                1,
            )
        )

        xia_included = (
            session_state.get(
                f"meal_{period_id}_xia_included",
                False,
            )
        )

        result = (
            calculate_meal_period_cost(
                start_date=start_date,
                end_date=end_date,
                interpreter_count=interpreter_count,
                auditor_count=auditor_count,
                xia_included=xia_included,
            )
        )

        periods.append(
            {
                "id": period_id,
                "start_date": start_date,
                "end_date": end_date,
                "interpreter_count":
                    interpreter_count,
                "auditor_count":
                    auditor_count,
                "xia_included":
                    xia_included,
                "days":
                    result["days"],
                "cost":
                    result["cost"],
            }
        )

        total_meal_cost += (
            result["cost"]
        )

    return {
        "periods":
            periods,

        "total_meal_cost":
            round(
                total_meal_cost,
                2,
            ),
    }

def calculate_cost_summary(
    timeline_df,
    session_state,
    added_rows,
    meal_periods=None,
    sonstiges=0.0,
):
    """
    Main function for Kostenübersicht.
    """

    # ========================================================
    # DISTANCE
    # ========================================================

    total_distance = calculate_total_distance(
        timeline_df=timeline_df,
        session_state=session_state,
        added_rows=added_rows,
    )

    # ========================================================
    # FUEL
    # ========================================================

    fuel_cost = calculate_fuel_cost(
        total_distance
    )

    # ========================================================
    # HOTELS
    # ========================================================

    hotels = collect_hotels(
        timeline_df=timeline_df,
        session_state=session_state,
        added_rows=added_rows,
    )

    hotel_totals = calculate_hotel_totals(
        hotels
    )

    # ========================================================
    # MEALS
    # ========================================================

    if meal_periods is None:
        meal_periods = []

    meal_summary = calculate_meal_total(
        session_state=session_state,
        meal_periods=meal_periods,
    )

    # ========================================================
    # OTHER COSTS
    # ========================================================

    sonstiges = safe_float(
        sonstiges
    )

    # ========================================================
    # REQUIRED BUDGET
    #
    # Include:
    # - fuel
    # - hotel "vor Ort"
    # - meals
    # - sonstiges
    #
    # Hotel "bezahlt" is displayed but NOT added again.
    # ========================================================

    required_budget = (
        fuel_cost
        + hotel_totals["vor_ort_total"]
        + meal_summary["total_meal_cost"]
        + sonstiges
    )

    # ========================================================
    # RESULT
    # ========================================================

    return {
        "total_distance":
            total_distance,

        "fuel_cost":
            fuel_cost,

        "hotels":
            hotels,

        "hotel_total":
            hotel_totals["total_hotels"],

        "hotel_vor_ort":
            hotel_totals["vor_ort_total"],

        "hotel_bezahlt":
            hotel_totals["bezahlt_total"],

        "meal_periods":
            meal_summary["periods"],

        "meal_total":
            meal_summary["total_meal_cost"],

        "sonstiges":
            round(
                sonstiges,
                2,
            ),

        "required_budget":
            round(
                required_budget,
                2,
            ),
    }
    # ============================================================
# INTERPRETER COST SPLIT
# ============================================================

from datetime import datetime, date


def parse_tour_date(value):
    """
    Convert different Tourplan date formats into datetime.date.
    """

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    # pandas Timestamp also normally supports .date()
    if hasattr(value, "date"):
        try:
            return value.date()
        except Exception:
            pass

    text = str(value).strip()

    if not text:
        return None

    formats = [
        "%d.%m.%Y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%Y/%m/%d",
    ]

    for date_format in formats:

        try:
            return datetime.strptime(
                text,
                date_format,
            ).date()

        except ValueError:
            continue

    return None


# ============================================================
# READ INTERPRETER PERIODS
# ============================================================

def collect_interpreter_periods(
    session_state,
    meal_periods,
):
    """
    Read all Dolmetscher-Zeiträume from Streamlit session state.

    Same interpreter name can have multiple periods.
    """

    interpreters = {}

    for period in meal_periods:

        period_id = period["id"]

        interpreter_name = str(
            session_state.get(
                f"meal_{period_id}_interpreter_name",
                "",
            )
            or ""
        ).strip()

        start_date = parse_tour_date(
            session_state.get(
                f"meal_{period_id}_start_date"
            )
        )

        end_date = parse_tour_date(
            session_state.get(
                f"meal_{period_id}_end_date"
            )
        )

        auditor_count = int(
            session_state.get(
                f"meal_{period_id}_auditor_count",
                1,
            )
            or 0
        )

        xia_included = bool(
            session_state.get(
                f"meal_{period_id}_xia_included",
                False,
            )
        )

        if not interpreter_name:
            continue

        if (
            start_date is None
            or end_date is None
            or end_date < start_date
        ):
            continue

        if interpreter_name not in interpreters:

            interpreters[
                interpreter_name
            ] = {
                "name": interpreter_name,
                "periods": [],
            }

        interpreters[
            interpreter_name
        ]["periods"].append(
            {
                "period_id": period_id,
                "start_date": start_date,
                "end_date": end_date,
                "auditor_count": auditor_count,
                "xia_included": xia_included,
            }
        )

    return interpreters


# ============================================================
# CHECK DATE IN INTERPRETER PERIOD
# ============================================================

def date_belongs_to_interpreter(
    row_date,
    periods,
):

    row_date = parse_tour_date(
        row_date
    )

    if row_date is None:
        return False

    for period in periods:

        if (
            period["start_date"]
            <= row_date
            <= period["end_date"]
        ):
            return True

    return False


# ============================================================
# COLLECT ALL TOUR COST ROWS
# ============================================================

def collect_tour_cost_rows(
    timeline_df,
    session_state,
    added_rows,
):
    """
    Build one normalized list from:
    - original Tourplan rows
    - manually added rows

    This list is later distributed by date to interpreters.
    """

    cost_rows = []

    # --------------------------------------------------------
    # ORIGINAL TOURPLAN ROWS
    # --------------------------------------------------------

    for _, row in timeline_df.iterrows():

        excel_row = int(
            row["Excel Row"]
        )

        row_key = (
            f"timeline_{excel_row}"
        )

        row_date = parse_tour_date(
            row["Date"]
        )

        distance = safe_float(
            session_state.get(
                f"{row_key}_distance",
                None,
            )
        )

        hotel_price = safe_float(
            session_state.get(
                f"{row_key}_hotel_price",
                None,
            )
        )

        hotel_name = str(
            session_state.get(
                f"{row_key}_factory_hotel",
                row["Factory/Hotel"],
            )
            or ""
        ).strip()

        payment_status = str(
            session_state.get(
                f"{row_key}_payment_status",
                "",
            )
            or ""
        ).strip()

        breakfast = str(
            session_state.get(
                f"{row_key}_breakfast",
                "",
            )
            or ""
        ).strip()

        cost_rows.append(
            {
                "source": "original",
                "excel_row": excel_row,
                "date": row_date,
                "distance": distance,
                "hotel": hotel_name,
                "hotel_price": hotel_price,
                "payment_status": payment_status,
                "breakfast": breakfast,
            }
        )

    # --------------------------------------------------------
    # MANUAL ROWS
    # --------------------------------------------------------

    for manual_row in added_rows:

        row_id = (
            manual_row["id"]
        )

        row_date = parse_tour_date(
            session_state.get(
                f"manual_{row_id}_date",
                "",
            )
        )

        distance = safe_float(
            session_state.get(
                f"manual_{row_id}_distance",
                None,
            )
        )

        hotel_price = safe_float(
            session_state.get(
                f"manual_{row_id}_hotel_price",
                None,
            )
        )

        hotel_name = str(
            session_state.get(
                f"manual_{row_id}_factory_hotel",
                "",
            )
            or ""
        ).strip()

        payment_status = str(
            session_state.get(
                f"manual_{row_id}_payment_status",
                "",
            )
            or ""
        ).strip()

        breakfast = str(
            session_state.get(
                f"manual_{row_id}_breakfast",
                "",
            )
            or ""
        ).strip()

        cost_rows.append(
            {
                "source": "manual",
                "row_id": row_id,
                "date": row_date,
                "distance": distance,
                "hotel": hotel_name,
                "hotel_price": hotel_price,
                "payment_status": payment_status,
                "breakfast": breakfast,
            }
        )

    return cost_rows


# ============================================================
# FUEL FORMULA FOR EXCEL
# ============================================================

def build_fuel_excel_formula(
    total_distance,
):
    """
    Excel formula shown directly in Kostenerstellung.

    Example:
    50 km
    -> =(50/100)*7*1.8
    """

    total_distance = safe_float(
        total_distance
    )

    return (
        f"=({total_distance}/100)*7*1.8"
    )


# ============================================================
# MEAL PERIOD INFORMATION
# ============================================================

def calculate_meal_period_info(
    period,
):
    """
    Prepare the data required for the Essen formula.

    IMPORTANT:
    The final Xia-specific Essen formula is intentionally
    NOT hard-coded here yet.

    One Zeitraum always belongs to one interpreter.
    """

    start_date = period[
        "start_date"
    ]

    end_date = period[
        "end_date"
    ]

    if (
        start_date is None
        or end_date is None
        or end_date < start_date
    ):

        return {
            "days": 0,
            "auditor_count": 0,
            "xia_included": False,
        }

    days = (
        end_date
        - start_date
    ).days + 1

    return {
        "days":
            days,

        "auditor_count":
            int(
                period.get(
                    "auditor_count",
                    0,
                )
                or 0
            ),

        "xia_included":
            bool(
                period.get(
                    "xia_included",
                    False,
                )
            ),
    }


# ============================================================
# BUILD MEAL EXCEL FORMULA
# ============================================================

def build_meal_excel_formula(
    periods,
    hotels,
):
    """
    Build the Excel formula for Essen.

    Rules:
    - 1 interpreter is always included
    - normal auditors are included with 47 EUR/day
    - Xia is excluded from the normal 47 EUR/day count
    - Xia gets +7 EUR breakfast only for days where:
        * Xia is included in the period
        * the relevant hotel has breakfast = "nicht inklusive"

    Example:
    2 days
    1 interpreter
    1 normal auditor
    no Xia

    -> =2*2*(7+15+25)

    Example:
    2 days
    1 interpreter
    2 auditors
    1 of them Xia
    Xia has 1 breakfast not included

    -> =2*2*(7+15+25)+1*7
    """

    formula_parts = []

    # ========================================================
    # NORMAL MEAL COST
    # ========================================================

    for period in periods:

        start_date = period.get(
            "start_date"
        )

        end_date = period.get(
            "end_date"
        )

        if (
            start_date is None
            or end_date is None
            or end_date < start_date
        ):
            continue

        days = (
            end_date
            - start_date
        ).days + 1

        auditor_count = int(
            period.get(
                "auditor_count",
                0,
            )
            or 0
        )

        xia_included = bool(
            period.get(
                "xia_included",
                False,
            )
        )

        # 1 interpreter is always included
        normal_people = (
            1
            + auditor_count
        )

        # Xia is not counted as a normal 47 EUR/day person
        if (
            xia_included
            and auditor_count > 0
        ):

            normal_people -= 1

        if normal_people > 0:

            formula_parts.append(
                f"{days}*{normal_people}*(7+15+25)"
            )


    # ========================================================
    # XIA BREAKFAST EXTRA
    # ========================================================

    xia_breakfast_days = 0

    for period in periods:

        if not period.get(
            "xia_included",
            False,
        ):
            continue

        start_date = period.get(
            "start_date"
        )

        end_date = period.get(
            "end_date"
        )

        if (
            start_date is None
            or end_date is None
        ):
            continue

        for hotel in hotels:

            hotel_date = hotel.get(
                "date"
            )

            breakfast = str(
                hotel.get(
                    "breakfast",
                    ""
                )
                or ""
            ).strip().lower()

            if hotel_date is None:
                continue

            if not (
                start_date
                <= hotel_date
                <= end_date
            ):
                continue

            if breakfast == "nicht inklusive":

                xia_breakfast_days += 1


    if xia_breakfast_days > 0:

        formula_parts.append(
            f"{xia_breakfast_days}*7"
        )


    # ========================================================
    # RESULT
    # ========================================================

    if not formula_parts:

        return "=0"

    return (
        "="
        + "+".join(
            formula_parts
        )
    )

# ============================================================
# CALCULATE COSTS BY INTERPRETER
# ============================================================

def calculate_costs_by_interpreter(
    timeline_df,
    session_state,
    added_rows,
    meal_periods,
):
    """
    Calculate one separate cost dataset per interpreter.

    Included:
    - interpreter periods
    - total distance
    - fuel amount
    - Excel fuel formula
    - hotels
    - hotel vor Ort
    - hotel bezahlt
    - meal calculation parameters
    """

    interpreters = (
        collect_interpreter_periods(
            session_state=session_state,
            meal_periods=meal_periods,
        )
    )
    # ========================================================
    # INTERPRETERS / PERIODS
    # ========================================================

    interpreters = (
        collect_interpreter_periods(
            session_state=session_state,
            meal_periods=meal_periods,
        )
    )


    # ========================================================
    # ALL TOUR COST ROWS
    # ========================================================

    cost_rows = (
        collect_tour_cost_rows(
            timeline_df=timeline_df,
            session_state=session_state,
            added_rows=added_rows,
        )
    )


    results = {}


    # ========================================================
    # EACH INTERPRETER
    # ========================================================

    for (
        interpreter_name,
        interpreter_data,
    ) in interpreters.items():

        periods = (
            interpreter_data[
                "periods"
            ]
        )


        # ====================================================
        # INITIAL VALUES
        # ====================================================

        total_distance = 0.0

        hotels = []

        hotel_total = 0.0
        hotel_vor_ort = 0.0
        hotel_bezahlt = 0.0


        # ====================================================
        # ASSIGN TOUR ROWS BY DATE
        # ====================================================

        for cost_row in cost_rows:

            if not date_belongs_to_interpreter(
                cost_row["date"],
                periods,
            ):
                continue


            # ------------------------------------------------
            # DISTANCE
            # ------------------------------------------------

            total_distance += safe_float(
                cost_row[
                    "distance"
                ]
            )


            # ------------------------------------------------
            # HOTEL
            # ------------------------------------------------

            hotel_price = safe_float(
                cost_row[
                    "hotel_price"
                ]
            )

            if hotel_price <= 0:
                continue


            status = str(
                cost_row[
                    "payment_status"
                ]
            ).strip().lower()


            hotel_data = {

                "date":
                    cost_row[
                        "date"
                    ],

                "hotel":
                    cost_row[
                        "hotel"
                    ],

                "price":
                    hotel_price,

                "payment_status":
                    cost_row[
                        "payment_status"
                    ],

                "breakfast":
                    cost_row[
                        "breakfast"
                    ],
            }


            hotels.append(
                hotel_data
            )


            hotel_total += (
                hotel_price
            )


            if status == "vor ort":

                hotel_vor_ort += (
                    hotel_price
                )


            elif status == "bezahlt":

                hotel_bezahlt += (
                    hotel_price
                )


        # ====================================================
        # DISTANCE
        # ====================================================

        total_distance = round(
            total_distance,
            2,
        )


        # ====================================================
        # FUEL
        # ====================================================

        fuel_cost = (
            calculate_fuel_cost(
                total_distance
            )
        )

        fuel_excel_formula = (
            build_fuel_excel_formula(
                total_distance
            )
        )


        # ====================================================
        # MEAL PERIOD DETAILS
        # ====================================================

        meal_period_details = []

        total_meal_days = 0


        for period in periods:

            meal_info = (
                calculate_meal_period_info(
                    period
                )
            )

            total_meal_days += (
                meal_info[
                    "days"
                ]
            )


            meal_period_details.append(
                {
                    "start_date":
                        period[
                            "start_date"
                        ],

                    "end_date":
                        period[
                            "end_date"
                        ],

                    "days":
                        meal_info[
                            "days"
                        ],

                    "auditor_count":
                        meal_info[
                            "auditor_count"
                        ],

                    "xia_included":
                        meal_info[
                            "xia_included"
                        ],
                }
            )


        # ====================================================
        # ESSEN EXCEL FORMULA
        # ====================================================

        meal_excel_formula = (
            build_meal_excel_formula(
                periods=periods,
                hotels=hotels,
            )
        )


        # ====================================================
        # CURRENT BUDGET
        #
        # 注意：
        # 这里暂时只是 Python 侧的中间值。
        # Essen 最终会在 Excel 里通过公式计算。
        # ====================================================

        current_budget = (
            fuel_cost
            + hotel_vor_ort
        )


        # ====================================================
        # RESULT
        # ====================================================

        results[
            interpreter_name
        ] = {

            "interpreter_name":
                interpreter_name,

            "periods":
                periods,

            "total_distance":
                total_distance,

            "fuel_cost":
                round(
                    fuel_cost,
                    2,
                ),

            "fuel_excel_formula":
                fuel_excel_formula,

            "hotels":
                hotels,

            "hotel_total":
                round(
                    hotel_total,
                    2,
                ),

            "hotel_vor_ort":
                round(
                    hotel_vor_ort,
                    2,
                ),

            "hotel_bezahlt":
                round(
                    hotel_bezahlt,
                    2,
                ),

            "meal_period_details":
                meal_period_details,

            "total_meal_days":
                total_meal_days,

            "meal_cost":
                0.0,

            "meal_excel_formula":
                meal_excel_formula,

            "sonstiges":
                0.0,

            "current_budget":
                round(
                    current_budget,
                    2,
                ),
        }


    return results