INPUT_FILE = "daily-activity.xlsx"
OUTPUT_FILE = "minor_project1_results.txt"

START_DATE = date(2026, 8, 13)
END_DATE = date(2026, 9, 21)

MINUTES_PER_DAY = 24 * 60

# Score ranges described in the project report.
FEELING_MIN, FEELING_MAX = 1, 5
SATISFACTION_MIN, SATISFACTION_MAX = 1, 5
VALID_ENERGY_SCORES = {1, 2, 3}

def average(values):
    """Return the arithmetic mean without using a statistics library."""
    if len(values) == 0:
        return 0
    return sum(values) / len(values)


def is_number(value):
    """Check whether a value is a number (but not True/False)."""
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def pearson_correlation(x_values, y_values):
    """
    Calculate Pearson correlation manually.

    Returns:
        correlation value, or None when correlation cannot be calculated
        because one variable has no variation.
    """
    if len(x_values) != len(y_values) or len(x_values) < 2:
        return None

    x_mean = average(x_values)
    y_mean = average(y_values)

    numerator = 0
    x_squared = 0
    y_squared = 0

    for x, y in zip(x_values, y_values):
        x_difference = x - x_mean
        y_difference = y - y_mean

        numerator += x_difference * y_difference
        x_squared += x_difference ** 2
        y_squared += y_difference ** 2

    denominator = sqrt(x_squared * y_squared)

    if denominator == 0:
        return None

    return numerator / denominator


def correlation_label(value):
    """Give a simple descriptive label to a correlation."""
    if value is None:
        return "Not calculable (one variable has no variation)."

    absolute_value = abs(value)

    if absolute_value < 0.20:
        strength = "very weak"
    elif absolute_value < 0.40:
        strength = "weak"
    elif absolute_value < 0.60:
        strength = "moderate"
    elif absolute_value < 0.80:
        strength = "strong"
    else:
        strength = "very strong"

    direction = "positive" if value > 0 else "negative"
    return f"{strength} {direction} relationship"


def get_value(row, column_number):
    """Read one Excel cell value."""
    return row[column_number - 1].value


def validate_record(record):
    """
    Validate one daily activity record.

    Returns:
        A list of validation errors. Empty list means the record is valid.
    """
    errors = []

    # Time fields must be numeric and non-negative.
    time_fields = [
        "sleep",
        "fitness",
        "study",
        "coding",
        "class_time",
        "other",
        "tracked",
        "free",
    ]

    for field in time_fields:
        value = record[field]
        if not is_number(value):
            errors.append(f"{field} is not numeric")
        elif value < 0:
            errors.append(f"{field} is negative")

    # The Excel totals must agree with the activity columns.
    if not any("not numeric" in error for error in errors):
        calculated_tracked = (
            record["sleep"]
            + record["fitness"]
            + record["study"]
            + record["coding"]
            + record["class_time"]
            + record["other"]
        )

        calculated_free = MINUTES_PER_DAY - calculated_tracked

        if record["tracked"] != calculated_tracked:
            errors.append(
                f"Total Tracked should be {calculated_tracked}, "
                f"but Excel contains {record['tracked']}"
            )

        if record["free"] != calculated_free:
            errors.append(
                f"Free/Unaccounted should be {calculated_free}, "
                f"but Excel contains {record['free']}"
            )

    # Project score validation.
    feeling = record["feeling"]
    satisfaction = record["satisfaction"]
    energy = record["energy"]

    if not is_number(feeling) or not (FEELING_MIN <= feeling <= FEELING_MAX):
        errors.append("Feeling score must be between 1 and 5")

    if not is_number(satisfaction) or not (
        SATISFACTION_MIN <= satisfaction <= SATISFACTION_MAX
    ):
        errors.append("Satisfaction score must be between 1 and 5")

    if not is_number(energy) or energy not in VALID_ENERGY_SCORES:
        errors.append("Energy score must be 1, 2, or 3")

    return errors


def load_activity_data(filename):
    """
    Load the workbook and return records inside the project date period.

    Rows outside the official project period are reported separately.
    """
    try:
        workbook = openpyxl.load_workbook(filename, data_only=True)
    except FileNotFoundError:
        raise FileNotFoundError(
            f"Could not find '{filename}'. Put the Excel file in the same "
            f"folder as this Python program."
        )
    except Exception as error:
        raise RuntimeError(f"Could not open the Excel file: {error}")

    if "daily-activity" not in workbook.sheetnames:
        raise ValueError(
            "The workbook does not contain a sheet named 'daily-activity'."
        )

    sheet = workbook["daily-activity"]

    records = []
    outside_period = []

    # Data starts at row 6 according to the supplied activity sheet.
    for row_number in range(6, sheet.max_row + 1):
        raw_date = get_value(sheet[row_number], 1)

        # Ignore completely blank rows.
        if raw_date is None:
            continue

        if isinstance(raw_date, datetime):
            activity_date = raw_date.date()
        elif isinstance(raw_date, date):
            activity_date = raw_date
        else:
            outside_period.append(
                (row_number, f"Invalid date value: {raw_date}")
            )
            continue

        record = {
            "row": row_number,
            "date": activity_date,
            "sleep": get_value(sheet[row_number], 2),
            "fitness": get_value(sheet[row_number], 3),
            "study": get_value(sheet[row_number], 4),
            "coding": get_value(sheet[row_number], 5),
            "class_time": get_value(sheet[row_number], 6),
            "classes_attended": get_value(sheet[row_number], 7),
            "other": get_value(sheet[row_number], 8),
            "tracked": get_value(sheet[row_number], 9),
            "free": get_value(sheet[row_number], 10),
            "feeling": get_value(sheet[row_number], 11),
            "satisfaction": get_value(sheet[row_number], 12),
            "energy": get_value(sheet[row_number], 13),
            "notes": get_value(sheet[row_number], 14),
        }

        if START_DATE <= activity_date <= END_DATE:
            records.append(record)
        else:
            outside_period.append(
                (row_number, f"Date {activity_date} is outside project period")
            )

    return records, outside_period

def calculate_indices(valid_records):
    """Calculate all project indices from valid recorded days."""
    tpi = average([record["coding"] for record in valid_records])

    aai = average([
        record["study"] + record["class_time"]
        for record in valid_records
    ])

    phai = average([record["fitness"] for record in valid_records])

    sri = average([record["sleep"] for record in valid_records])

    abi = average([record["free"] for record in valid_records])

    tui = average([record["tracked"] for record in valid_records])

    ei = average([
        (
            record["feeling"]
            + record["satisfaction"]
            + record["energy"]
        ) / 3
        for record in valid_records
    ])

    expected_days = (END_DATE - START_DATE).days + 1
    dci = (len(valid_records) / expected_days) * 100

    # PAI formula exactly as supplied in the project report.
    pai = (
        0.15 * tpi
        + 0.20 * aai
        + 0.15 * phai
        + 0.20 * sri
        + 0.15 * tui
        + 0.10 * ei
        + 0.05 * dci
    )

    return {
        "TPI": tpi,
        "AAI": aai,
        "PhAI": phai,
        "SRI": sri,
        "ABI": abi,
        "TUI": tui,
        "EI": ei,
        "DCI": dci,
        "PAI": pai,

        def calculate_relationships(valid_records):
    """Calculate the requested relationships from valid records."""
    relationships = {}

    relationships["Sleep <-> Energy"] = pearson_correlation(
        [r["sleep"] for r in valid_records],
        [r["energy"] for r in valid_records],
    )

    relationships["Study <-> Satisfaction"] = pearson_correlation(
        [r["study"] for r in valid_records],
        [r["satisfaction"] for r in valid_records],
    )

    relationships["Coding <-> Energy"] = pearson_correlation(
        [r["coding"] for r in valid_records],
        [r["energy"] for r in valid_records],
    )

    # Additional relationship useful when Satisfaction has no variation.
    relationships["Study <-> Feeling"] = pearson_correlation(
        [r["study"] for r in valid_records],
        [r["feeling"] for r in valid_records],
    )

    return relationships

ef print_record_table(valid_records, invalid_records):
    """Print valid and invalid records for transparency."""
    print("\nVALID RECORDS")
    print("-" * 70)

    for record in valid_records:
        print(
            f"{record['date']} | "
            f"Sleep={record['sleep']:>4} | "
            f"Study={record['study']:>3} | "
            f"Coding={record['coding']:>3} | "
            f"Class={record['class_time']:>3} | "
            f"Energy={record['energy']}"
        )

    print("\nINVALID / EXCLUDED RECORDS")
    print("-" * 70)

    if not invalid_records:
        print("None")
    else:
        for record, errors in invalid_records:
            print(f"{record['date']} -> {'; '.join(errors)}")


def build_report(
    valid_records,
    invalid_records,
    outside_period,
    indices,
    relationships,
):
    """Create a text report using only calculated program results."""

    expected_days = (END_DATE - START_DATE).days + 1

    lines = []

    lines.append("CAP776 - MINOR PROJECT #1")
    lines.append("MY DATA, MY STORY")
    lines.append("=" * 70)
    lines.append("")

    lines.append("1. ACTIVITY DATA SUMMARY")
    lines.append("-" * 70)
    lines.append(f"Project period: {START_DATE} to {END_DATE}")
    lines.append(f"Expected number of days: {expected_days}")
    lines.append(f"Valid days recorded: {len(valid_records)}")
    lines.append(f"Missing days: {expected_days - len(valid_records)}")
    lines.append(f"Invalid / excluded records: {len(invalid_records)}")
    lines.append("")

    if valid_records:
        lines.append(f"Average Sleep/day: {indices['SRI']:.2f} min")
        lines.append(
            f"Average Fitness/day: {indices['PhAI']:.2f} min"
        )
        lines.append(
            f"Average Study/day: "
            f"{average([r['study'] for r in valid_records]):.2f} min"
        )
        lines.append(f"Average Coding/day: {indices['TPI']:.2f} min")
        lines.append(
            f"Average Class/day: "
            f"{average([r['class_time'] for r in valid_records]):.2f} min"
        )
        lines.append(
            f"Average Other Activities/day: "
            f"{average([r['other'] for r in valid_records]):.2f} min"
        )
        lines.append(
            f"Average Free / Unaccounted Time/day: "
            f"{indices['ABI']:.2f} min"
        )

    lines.append("")
    lines.append("2. INDEX VALUES")
    lines.append("-" * 70)
    lines.append(
        f"TPI  - Tech Productivity:       {indices['TPI']:.2f} min/day"
    )
    lines.append(
        f"AAI  - Academic Activity:       {indices['AAI']:.2f} min/day"
    )
    lines.append(
        f"PhAI - Physical Activity:       {indices['PhAI']:.2f} min/day"
    )
    lines.append(
        f"SRI  - Sleep & Recovery:        {indices['SRI']:.2f} min/day"
    )
    lines.append(
        f"ABI  - Activity Balance:        {indices['ABI']:.2f} min/day"
    )
    lines.append(
        f"TUI  - Time Utilization:        {indices['TUI']:.2f} min/day"
    )
    lines.append(
        f"EI   - Experience Index:        {indices['EI']:.2f} / 5"
    )
    lines.append(
        f"DCI  - Data Continuity Index:   {indices['DCI']:.2f} %"
    )
    lines.append(
        f"PAI  - Personal Activity Index: {indices['PAI']:.2f}"
    )

    lines.append("")
    lines.append("3. RELATIONSHIP ANALYSIS")
    lines.append("-" * 70)

    for name, value in relationships.items():
        if value is None:
            lines.append(f"{name}: Not calculable.")
            lines.append(
                "  Reason: one variable has no variation in the valid data."
            )
        else:
            lines.append(
                f"{name}: r = {value:.3f} "
                f"({correlation_label(value)})"
            )

    lines.append("")
    lines.append("4. VALIDATION NOTES")
    lines.append("-" * 70)

    if invalid_records:
        for record, errors in invalid_records:
            lines.append(
                f"{record['date']}: {'; '.join(errors)}"
            )
    else:
        lines.append("No invalid records found.")

    if outside_period:
        lines.append("")
        lines.append("Rows outside the official project period:")
        for row_number, reason in outside_period:
            lines.append(f"Row {row_number}: {reason}")

    lines.append("")
    lines.append("5. IMPORTANT")
    lines.append("-" * 70)
    lines.append(
        "The program does not fabricate missing days or invalid scores. "
        "Only valid records inside the official project period are used "
        "for the index calculations."
    )
    lines.append(
        "Review the generated values and write the personal findings "
        "and improvement section yourself using evidence from your data."
    )

    return "\n".join(lines)


def main():
    print("CAP776 Minor Project #1 - My Data, My Story")
    print("=" * 70)

    try:
        records, outside_period = load_activity_data(INPUT_FILE)
    except Exception as error:
        print(f"\nERROR: {error}")
        return

    valid_records = []
    invalid_records = []

    for record in records:
        errors = validate_record(record)

        if errors:
            invalid_records.append((record, errors))
        else:
            valid_records.append(record)

    if len(valid_records) == 0:
        print("\nNo valid records were found.")
        print("Check the Excel data and the validation rules.")
        return

    indices = calculate_indices(valid_records)
    relationships = calculate_relationships(valid_records)

    print_record_table(valid_records, invalid_records)

    report = build_report(
        valid_records,
        invalid_records,
        outside_period,
        indices,
        relationships,
    )

    print("\n" + report)

    try:
        with open(OUTPUT_FILE, "w", encoding="utf-8") as file:
            file.write(report)

        print("\n" + "=" * 70)
        print(f"Results saved to: {OUTPUT_FILE}")

    except OSError as error:
        print(f"\nCould not save the output file: {error}")


if __name__ == "__main__":
    main()

