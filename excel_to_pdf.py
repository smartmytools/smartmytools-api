import io
from datetime import date, datetime, time

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape, portrait
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    PageBreak,
)


# ============================================================
# COLOR HELPERS
# ============================================================

def excel_color_to_reportlab(color):
    """
    Convert an openpyxl color into a ReportLab color.
    Supports RGB colors and common indexed colors.
    """

    if color is None:
        return None

    try:
        if color.type == "rgb" and color.rgb:
            rgb = color.rgb

            if len(rgb) == 8:
                rgb = rgb[2:]

            if len(rgb) == 6:
                return colors.HexColor("#" + rgb)

        if color.type == "indexed" and color.indexed is not None:
            indexed = {
                0: "#000000",
                1: "#FFFFFF",
                2: "#FF0000",
                3: "#00FF00",
                4: "#0000FF",
                5: "#FFFF00",
                6: "#FF00FF",
                7: "#00FFFF",
                8: "#000000",
                9: "#FFFFFF",
            }

            if color.indexed in indexed:
                return colors.HexColor(
                    indexed[color.indexed]
                )

    except Exception:
        pass

    return None


# ============================================================
# BORDER HELPERS
# ============================================================

def border_style_to_width(style):
    """
    Convert Excel border styles into approximate PDF widths.
    """

    widths = {
        "hair": 0.25,
        "dotted": 0.5,
        "dashed": 0.75,
        "thin": 0.75,
        "medium": 1.25,
        "thick": 2.0,
        "double": 1.5,
    }

    return widths.get(style, 0.5)


def add_cell_borders(table_commands, cell, row, column):
    """
    Add Excel cell borders to ReportLab TableStyle commands.
    """

    border = cell.border

    sides = [
        ("LEFT", border.left),
        ("RIGHT", border.right),
        ("TOP", border.top),
        ("BOTTOM", border.bottom),
    ]

    for side_name, side in sides:

        if side is None:
            continue

        if not side.style:
            continue

        border_color = excel_color_to_reportlab(
            side.color
        )

        if border_color is None:
            border_color = colors.black

        width = border_style_to_width(
            side.style
        )

        table_commands.append(
            (
                f"LINE{side_name}",
                (column, row),
                (column, row),
                width,
                border_color,
            )
        )


# ============================================================
# FONT HELPERS
# ============================================================

def get_font_name(cell):
    """
    Choose an appropriate ReportLab font.
    """

    font = cell.font

    if font.bold and font.italic:
        return "Helvetica-BoldOblique"

    if font.bold:
        return "Helvetica-Bold"

    if font.italic:
        return "Helvetica-Oblique"

    return "Helvetica"


def get_font_size(cell):
    """
    Return Excel font size or a sensible default.
    """

    if cell.font.sz:
        return max(6, min(float(cell.font.sz), 24))

    return 10


# ============================================================
# ALIGNMENT HELPERS
# ============================================================

def get_horizontal_alignment(cell):
    """
    Convert Excel horizontal alignment to ReportLab alignment.
    """

    alignment = cell.alignment.horizontal

    # Respect explicit Excel alignment first.
    if alignment == "center":
        return TA_CENTER

    if alignment == "right":
        return TA_RIGHT

    if alignment == "left":
        return TA_LEFT

    # If Excel has no explicit alignment, automatically
    # right-align numeric values.
    value = cell.value

    if (
        isinstance(value, (int, float))
        and not isinstance(value, bool)
    ):
        return TA_RIGHT

    # Text remains left-aligned.
    return TA_LEFT


def get_vertical_alignment(cell):
    """
    Convert Excel vertical alignment to ReportLab vertical alignment.
    """

    alignment = cell.alignment.vertical

    if alignment == "top":
        return "TOP"

    if alignment == "bottom":
        return "BOTTOM"

    return "MIDDLE"


# ============================================================
# VALUE FORMATTING
# ============================================================

def format_cell_value(value, number_format=None):
    """
    Convert Excel values into readable PDF text.
    """

    if value is None:
        return ""

    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")

    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    if isinstance(value, time):
        return value.strftime("%H:%M:%S")

    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"

    if isinstance(value, float):

        if number_format:
            try:

                if "%" in number_format:
                    return f"{value * 100:.2f}%"

                if (
                    "0.00" in number_format
                    or "#,##0.00" in number_format
                ):
                    return f"{value:,.2f}"

                if (
                    "0.0" in number_format
                    or "#,##0.0" in number_format
                ):
                    return f"{value:,.1f}"

                if (
                    "#,##0" in number_format
                    or "0" in number_format
                ):
                    return f"{value:,.0f}"

            except Exception:
                pass

    if isinstance(value, int):
        return str(value)

    return str(value)


# ============================================================
# COLUMN WIDTH
# ============================================================

def excel_width_to_points(width):
    """
    Convert Excel column-width units into PDF points.

    The conversion keeps the original Excel proportions
    without unnecessarily stretching narrow worksheets.
    """

    if width is None:
        width = 8.43

    width = max(
        2,
        float(width)
    )

    # Approximate Excel character width in points.
    return width * 5.5

# ============================================================
# WORKSHEET DIMENSIONS
# ============================================================

def get_used_dimensions(worksheet):
    """
    Determine the useful worksheet dimensions.

    Trailing completely empty columns and rows are removed,
    while meaningful merged-cell ranges are preserved.
    """

    max_row = worksheet.max_row
    max_column = worksheet.max_column

    # --------------------------------------------------------
    # Find last row containing actual content.
    # --------------------------------------------------------

    while max_row > 1:

        has_content = False

        for column in range(1, max_column + 1):

            cell = worksheet.cell(
                row=max_row,
                column=column
            )

            if cell.value is not None:
                has_content = True
                break

        if has_content:
            break

        max_row -= 1

    # --------------------------------------------------------
    # Find last column containing actual content.
    # --------------------------------------------------------

    content_max_column = 1

    for row in range(1, max_row + 1):

        for column in range(1, max_column + 1):

            cell = worksheet.cell(
                row=row,
                column=column
            )

            if cell.value is not None:
                content_max_column = max(
                    content_max_column,
                    column
                )

    max_column = content_max_column

    # --------------------------------------------------------
    # Preserve merged ranges that contain actual content.
    #
    # We only extend the used area when the top-left cell of
    # the merged range contains a value.
    # --------------------------------------------------------

    for merged_range in worksheet.merged_cells.ranges:

        min_row = merged_range.min_row
        min_column = merged_range.min_col

        if (
            min_row <= max_row
            and min_column <= max_column
        ):

            top_left_cell = worksheet.cell(
                row=min_row,
                column=min_column
            )

            if top_left_cell.value is not None:

                max_row = max(
                    max_row,
                    merged_range.max_row
                )

                max_column = max(
                    max_column,
                    merged_range.max_col
                )

    # --------------------------------------------------------
    # Remove trailing columns that are still completely empty.
    #
    # This prevents unnecessary blank columns caused only by
    # worksheet formatting.
    # --------------------------------------------------------

    while max_column > 1:

        has_content = False

        for row in range(1, max_row + 1):

            cell = worksheet.cell(
                row=row,
                column=max_column
            )

            if cell.value is not None:
                has_content = True
                break

        if has_content:
            break

        # Check whether this column is required by a
        # meaningful merged range.
        required_by_merge = False

        for merged_range in worksheet.merged_cells.ranges:

            if (
                merged_range.min_col
                <= max_column
                <= merged_range.max_col
            ):

                top_left_cell = worksheet.cell(
                    row=merged_range.min_row,
                    column=merged_range.min_col
                )

                if top_left_cell.value is not None:
                    required_by_merge = True
                    break

        if required_by_merge:
            break

        max_column -= 1

    return max_row, max_column

    # ============================================================
# MERGED CELL LOOKUP
# ============================================================

def build_merged_cell_map(worksheet):
    """
    Create a lookup table identifying merged cells.
    """

    merged_map = {}

    for merged_range in worksheet.merged_cells.ranges:

        min_column = merged_range.min_col
        max_column = merged_range.max_col

        min_row = merged_range.min_row
        max_row = merged_range.max_row

        for row in range(
            min_row,
            max_row + 1
        ):

            for column in range(
                min_column,
                max_column + 1
            ):

                merged_map[
                    (row, column)
                ] = (
                    min_row,
                    min_column,
                    max_row,
                    max_column
                )

    return merged_map


# ============================================================
# CELL PARAGRAPH
# ============================================================

def make_cell_paragraph(cell, value):
    """
    Create a ReportLab Paragraph using Excel formatting.
    """

    text = format_cell_value(
        value,
        cell.number_format
    )

    if not text:
        text = " "

    text = (
        str(text)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace("\n", "<br/>")
    )

    font_color = excel_color_to_reportlab(
        cell.font.color
    )

    if font_color is None:
        font_color = colors.black

    style = ParagraphStyle(
        name="ExcelCell",
        fontName=get_font_name(cell),
        fontSize=get_font_size(cell),
        leading=max(
            get_font_size(cell) + 2,
            10
        ),
        textColor=font_color,
        alignment=get_horizontal_alignment(cell),
        spaceBefore=0,
        spaceAfter=0,
        leftIndent=0,
        rightIndent=0,
        wordWrap="CJK",
    )

    return Paragraph(
        text,
        style
    )

def create_worksheet_table(
    worksheet,
    value_worksheet,
    max_row,
    max_column,
    available_width,
    column_group=None,
    render_from_row=None
):
    """
    Convert one Excel worksheet into a ReportLab Table.

    Preserves:
    - merged cells
    - column widths
    - row heights
    - cell backgrounds
    - font formatting
    - alignment
    - borders
    """

    data = []

    merged_map = build_merged_cell_map(
        worksheet
    )

    # --------------------------------------------------------
    # Detect the actual table header row automatically.
    #
    # A header row normally contains multiple non-empty
    # text cells. This avoids hard-coding a specific Excel
    # row number such as row 6.
    # --------------------------------------------------------

    header_row = 1

    for candidate_row in range(
        1,
        max_row + 1
    ):

        non_empty_values = []

        for candidate_column in range(
            1,
            max_column + 1
        ):

            value = worksheet.cell(
                row=candidate_row,
                column=candidate_column
            ).value

            if (
                value is not None
                and str(value).strip() != ""
            ):

                non_empty_values.append(
                    value
                )

        text_value_count = sum(
            1
            for value in non_empty_values
            if isinstance(value, str)
        )

        if (
            len(non_empty_values) >= 2
            and text_value_count >= 2
        ):

            header_row = candidate_row
            break


    # --------------------------------------------------------
    # Determine the first row that should be rendered.
    #
    # The first horizontal group keeps the complete worksheet.
    # Later horizontal groups start directly at the actual
    # detected header row.
    # --------------------------------------------------------

    if (
        column_group is not None
        and 1 not in column_group
    ):
        first_render_row = header_row
    else:
        first_render_row = 1


    # --------------------------------------------------------
    # Build cell matrix
    # --------------------------------------------------------

    # --------------------------------------------------------
    # Horizontal groups after the first group should not carry
    # the worksheet title/blank rows.
    # --------------------------------------------------------

    for row in range(
        first_render_row,
        max_row + 1
    ):

        print(
            "DEBUG:",
            "first_render_row =", first_render_row,
            "header_row =", header_row,
            "column_group =", column_group,
            "max_row =", max_row
        )
        row_data = []

        if column_group is None:

            columns_to_render = list(
                range(1, max_column + 1)
            )

        else:

            columns_to_render = [
                column_index + 1
                for column_index in column_group
            ]

        for column in columns_to_render:

            cell = worksheet.cell(
                row=row,
                column=column
            )

            value_cell = value_worksheet.cell(
                row=row,
                column=column
            )

            value = value_cell.value

            # Keep formula when there is no cached value
            if (
                value is None
                and isinstance(cell.value, str)
                and cell.value.startswith("=")
            ):
                value = cell.value

            # Only top-left cell of a merged range
            # contains the displayed value
            merged_info = merged_map.get(
                (row, column)
            )

            if merged_info:

                min_row, min_column, _, _ = (
                    merged_info
                )

                if (
                    row != min_row
                    or column != min_column
                ):
                    value = ""

            row_data.append(
                make_cell_paragraph(
                    cell,
                    value
                )
            )

        data.append(row_data)

        # --------------------------------------------------------
    # Excel column widths
    # --------------------------------------------------------

    column_widths = []

    for column in columns_to_render:

        letter = get_column_letter(
            column
        )

        width = worksheet.column_dimensions[
            letter
        ].width

        if width is None:
            width = 8.43

        column_widths.append(
            excel_width_to_points(width)
        )

    # --------------------------------------------------------
    # Scale this horizontal group to use the full
    # available PDF page width.
    # --------------------------------------------------------

    total_width = sum(
        column_widths
    )

    if total_width > 0:

        if (
            column_group is not None
            and total_width < available_width
        ):

            group_scale = (
                available_width
                / total_width
            )

            column_widths = [
                width * group_scale
                for width in column_widths
            ]

        elif (
            column_group is None
            and total_width > available_width
        ):

            scale = (
                available_width
                / total_width
            )

            column_widths = [
                width * scale
                for width in column_widths
            ]

    # --------------------------------------------------------
    # Excel row heights
    # --------------------------------------------------------

    row_heights = []

    # Calculate the same PDF column widths that will be
    # used by the ReportLab table.
    measurement_column_widths = []

    for column in columns_to_render:

        letter = get_column_letter(
            column
        )

        width = worksheet.column_dimensions[
            letter
        ].width

        if width is None:
            width = 8.43

        measurement_column_widths.append(
            excel_width_to_points(width)
        )

    measurement_total_width = sum(
        measurement_column_widths
    )

    if measurement_total_width > 0:

        if (
            column_group is not None
            and measurement_total_width < available_width
        ):

            measurement_scale = (
                available_width
                / measurement_total_width
            )

            measurement_column_widths = [
                width * measurement_scale
                for width in measurement_column_widths
            ]

        elif (
            column_group is None
            and measurement_total_width > available_width
        ):

            measurement_scale = (
                available_width
                / measurement_total_width
            )

            measurement_column_widths = [
                width * measurement_scale
                for width in measurement_column_widths
            ]

            
    for row in range(
        first_render_row,
        max_row + 1
    ):

        # -------------------------------------------------
        # Hide worksheet title/blank rows in later
        # horizontal groups.
        #
        # Keep the rows in the table so data and
        # row_heights always have the same length.
        # -------------------------------------------------

        if (
            column_group is not None
            and 1 not in column_group
            and row < header_row
        ):

            row_heights.append(0)
            continue

        height = worksheet.row_dimensions[
            row
        ].height

        required_height = 14

        # ----------------------------------------------------
        # Inventory title rows
        # ----------------------------------------------------

        if row in (2, 4):

            row_heights.append(28)

            continue

        

        # ----------------------------------------------------
        # Measure the actual ReportLab Paragraph height.
        # This detects automatic wrapping as well as
        # explicit line breaks.
        # ----------------------------------------------------

            if column_group is None:

                columns_to_measure = list(
                    range(1, max_column + 1)
                )

            else:

                columns_to_measure = [
                    column_index + 1
                    for column_index in column_group
                ]

            for column in columns_to_measure:

                if column_group is None:

                    data_index = column - 1
                    width_index = column - 1

                else:

                    data_index = columns_to_measure.index(
                        column
                    )
                    width_index = data_index

                paragraph = data[
                    row - 1
                ][
                    data_index
                ]

                if paragraph is None:
                    continue

                column_width = section_column_widths[
                    column_index
                ]

            try:

                _, paragraph_height = paragraph.wrap(
                    column_width - 10,
                    1000
                )

                required_height = max(
                    required_height,
                    paragraph_height + 8
                )

            except Exception:

                pass

        # ----------------------------------------------------
        # Keep a minimum height and preserve any larger
        # height explicitly set in Excel.
        # ----------------------------------------------------

        if height is None:

            row_heights.append(
                max(
                    14,
                    required_height
                )
            )

        else:

            row_heights.append(
                max(
                    14,
                    float(height),
                    required_height
                )
            )
    # --------------------------------------------------------
    # Create ReportLab table
    # --------------------------------------------------------

    repeat_rows = 0

    # Repeat the actual table header on continuation pages.
    if column_group is not None:

        # First horizontal group:
        # repeat only the actual Excel header row.
        if 1 in column_group:

            if header_row is not None:

                repeat_row_index = (
                    header_row
                    - first_render_row
                )

                repeat_rows = (
                    repeat_row_index,
                )

            else:

                repeat_rows = 0

        else:

            # Other horizontal groups start
            # directly with their table header.
            repeat_rows = 1
            
    table = Table(
        data,
        colWidths=column_widths,
        rowHeights=row_heights,
        hAlign="LEFT",
        splitByRow=1,
        repeatRows=repeat_rows,
    )

    commands = []
    # --------------------------------------------------------
    # Spreadsheet grid
    # --------------------------------------------------------

    commands.append(
        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.5,
            colors.black
        )
    )

    # --------------------------------------------------------
    # Cell formatting
    # --------------------------------------------------------

    for row in range(
        first_render_row,
        max_row + 1
    ):

        if column_group is None:

            columns_to_format = list(
                range(1, max_column + 1)
            )

        else:

            columns_to_format = [
                column_index + 1
                for column_index in column_group
            ]

        for column in columns_to_format:

            cell = worksheet.cell(
                row=row,
                column=column
            )

            pdf_row = row - first_render_row

            if column_group is None:

                pdf_column = column - 1

            else:

                pdf_column = column_group.index(
                    column - 1
                )

            # ------------------------------------------------
            # Background
            # ------------------------------------------------

            fill_color = None

            try:

                if cell.fill.fill_type:

                    fill_color = (
                        excel_color_to_reportlab(
                            cell.fill.fgColor
                        )
                    )

            except Exception:

                fill_color = None

            if fill_color is not None:

                commands.append(
                    (
                        "BACKGROUND",
                        (
                            pdf_column,
                            pdf_row
                        ),
                        (
                            pdf_column,
                            pdf_row
                        ),
                        fill_color
                    )
                )

            # ------------------------------------------------
            # Horizontal alignment
            # ------------------------------------------------

            horizontal = (
                cell.alignment.horizontal
            )

            # Center Inventory title rows across
            # the full displayed table width.
            if (
                row in (2, 4)
                and column == 2
            ):

                commands.append(
                    (
                        "ALIGN",
                        (1, pdf_row),
                        (max_column - 1, pdf_row),
                        "CENTER"
                    )
                )

            elif horizontal == "center":

                commands.append(
                    (
                        "ALIGN",
                        (
                            pdf_column,
                            pdf_row
                        ),
                        (
                            pdf_column,
                            pdf_row
                        ),
                        "CENTER"
                    )
                )

            elif horizontal == "right":

                commands.append(
                    (
                        "ALIGN",
                        (
                            pdf_column,
                            pdf_row
                        ),
                        (
                            pdf_column,
                            pdf_row
                        ),
                        "RIGHT"
                    )
                )

            else:

                commands.append(
                    (
                        "ALIGN",
                        (
                            pdf_column,
                            pdf_row
                        ),
                        (
                            pdf_column,
                            pdf_row
                        ),
                        "LEFT"
                    )
                )

            # ------------------------------------------------
            # Vertical alignment
            # ------------------------------------------------
            commands.append(
                (
                    "VALIGN",
                    (
                        pdf_column,
                        pdf_row
                    ),
                    (
                        pdf_column,
                        pdf_row
                    ),
                    get_vertical_alignment(cell)
                )
            )

    # --------------------------------------------------------
    # Merged cells
    #
    # If a merged range extends beyond the useful worksheet
    # area, clip it to the last displayed column/row.
    # --------------------------------------------------------

    for merged_range in worksheet.merged_cells.ranges:

        # Original Excel column positions
        excel_min_column = (
            merged_range.min_col - 1
        )

        excel_max_column = (
            merged_range.max_col - 1
        )

        min_row = (
            merged_range.min_row - 1
        )

        max_row_range = (
            merged_range.max_row - 1
        )

        # -------------------------------------------------
        # When rendering a horizontal group, only create
        # a span for the portion inside this group.
        # -------------------------------------------------

        if column_group is not None:

            group_columns = column_group

            group_min_excel_column = (
                group_columns[0]
            )

            group_max_excel_column = (
                group_columns[-1]
            )

            if (
                excel_max_column
                < group_min_excel_column
                or
                excel_min_column
                > group_max_excel_column
            ):
                continue

            clipped_min_column = max(
                excel_min_column,
                group_min_excel_column
            )

            clipped_max_column = min(
                excel_max_column,
                group_max_excel_column
            )

            if (
                clipped_min_column
                not in group_columns
                or
                clipped_max_column
                not in group_columns
            ):
                continue

            min_column = group_columns.index(
                clipped_min_column
            )

            max_column_range = group_columns.index(
                clipped_max_column
            )

        else:

            min_column = excel_min_column

            max_column_range = min(
                excel_max_column,
                max_column - 1
            )

        # -------------------------------------------------
        # Row safety
        # -------------------------------------------------

        if min_row >= max_row:
            continue

        max_row_range = min(
            max_row_range,
            max_row - 1
        )

        if max_row_range < min_row:
            continue

        # A one-cell span is unnecessary.
        if (
            min_column == max_column_range
            and
            min_row == max_row_range
        ):
            continue

        commands.append(
            (
                "SPAN",
                (
                    min_column,
                    min_row
                ),
                (
                    max_column_range,
                    max_row_range
                )
            )
        )


       # --------------------------------------------------------
    # Excel title rows
    #
    # Treat a row as a title only when the cell in column B
    # is actually part of a merged range.
    #
    # This prevents normal data rows from being treated as
    # title rows.
    # --------------------------------------------------------

    for row in range(
        1,
        max_row + 1
    ):

        title_is_merged = False

        for merged_range in worksheet.merged_cells.ranges:

            if (
                merged_range.min_row == row
                and merged_range.min_col == 2
                and merged_range.max_col > 2
            ):

                title_is_merged = True
                break

        if not title_is_merged:
            continue

        if column_group is None:

            commands.append(
                (
                    "SPAN",
                    (0, row - first_render_row),
                    (max_column - 1, row - first_render_row)
                )
            )

        else:

            if 1 in column_group:

                title_start = (
                    column_group.index(1)
                )

                title_end = (
                    len(column_group) - 1
                )

                if title_end > title_start:

                    commands.append(
                        (
                            "SPAN",
                            (
                                title_start,
                                row - 1
                            ),
                            (
                                title_end,
                                row - 1
                            )
                        )
                    )
                  
    # --------------------------------------------------------
    # Center actual Excel title rows
    #
    # Only center rows that are actually merged title rows.
    # Normal data rows must keep their original alignment.
    # --------------------------------------------------------

    for row in range(
        1,
        max_row + 1
    ):

        title_is_merged = False

        for merged_range in worksheet.merged_cells.ranges:

            if (
                merged_range.min_row == row
                and merged_range.min_col == 2
                and merged_range.max_col > 2
            ):

                title_is_merged = True
                break

        if not title_is_merged:
            continue

        commands.append(
            (
                "ALIGN",
                (0, row - first_render_row),
                (max_column - 1, row - first_render_row),
                "CENTER"
            )
        )

        commands.append(
            (
                "VALIGN",
                (0, row - first_render_row),
                (max_column - 1, row - first_render_row),
                "MIDDLE"
            )
        )
    # --------------------------------------------------------
    # Cell padding
    # --------------------------------------------------------

    commands.extend(
        [
            (
                "LEFTPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "RIGHTPADDING",
                (0, 0),
                (-1, -1),
                5
            ),
            (
                "TOPPADDING",
                (0, 0),
                (-1, -1),
                4
            ),
            (
                "BOTTOMPADDING",
                (0, 0),
                (-1, -1),
                4
            ),
        ]
    )

    # --------------------------------------------------------
    # Apply table style
    # --------------------------------------------------------

    table.setStyle(
        TableStyle(commands)
    )

    return table

    
# ============================================================
# MAIN CONVERTER
# ============================================================

def convert_excel_to_pdf(excel_bytes):
    """
    Convert an Excel workbook to PDF.

    Each logical table section is rendered as a separate
    ReportLab table so that a table heading does not get
    separated from its table content when the table fits
    on the next page.
    """

    workbook = load_workbook(
        io.BytesIO(excel_bytes),
        data_only=False
    )

    value_workbook = load_workbook(
        io.BytesIO(excel_bytes),
        data_only=True
    )

    output = io.BytesIO()

    page_width, page_height = landscape(A4)

    left_margin = 24
    right_margin = 24
    top_margin = 24
    bottom_margin = 24

    document = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=left_margin,
        rightMargin=right_margin,
        topMargin=top_margin,
        bottomMargin=bottom_margin,
    )

    available_width = (
        page_width
        - left_margin
        - right_margin
    )

    # -------------------------------------------------
    # Determine automatic horizontal column groups.
    # -------------------------------------------------

    
    story = []

    for worksheet in workbook.worksheets:

        value_worksheet = value_workbook[worksheet.title]

        max_row, max_column = get_used_dimensions(
            worksheet
        )

        # -------------------------------------------------
        # Find logical table heading rows.
        #
        # In this workbook the table headings are in
        # column A and begin with "Table ".
        # -------------------------------------------------

        heading_rows = []

        for row in range(1, max_row + 1):
            value = worksheet.cell(
                row=row,
                column=1
            ).value

            if (
                isinstance(value, str)
                and value.strip().startswith("Table ")
            ):
                heading_rows.append(row)

        # -------------------------------------------------
        # Rows before the first table heading.
        # Usually this contains the worksheet title.
        # -------------------------------------------------

        if heading_rows:

            first_heading = heading_rows[0]

            if first_heading > 1:

                pre_table = create_worksheet_table(
                    worksheet,
                    value_worksheet,
                    first_heading - 1,
                    max_column,
                    available_width
                )

                story.append(pre_table)
                story.append(Spacer(1, 8))

        else:

                    
            # -------------------------------------------------
            # No table headings were found.
            #
            # Build automatic horizontal column groups.
            # -------------------------------------------------

            column_widths = []

            for column in range(
                1,
                max_column + 1
            ):

                letter = get_column_letter(
                    column
                )

                width = worksheet.column_dimensions[
                    letter
                ].width

                if width is None:

                    width = 8.43

                column_widths.append(
                    excel_width_to_points(
                        width
                    )
                )

            # -------------------------------------------------
            # Maximum 5 Excel columns per horizontal group.
            # Also break earlier if the group becomes too wide.
            # -------------------------------------------------

            preferred_group_width = (
                available_width * 0.70
            )

            maximum_columns_per_group = 5

            column_groups = []

            current_group = []
            current_width = 0

            for column_index, width in enumerate(
                column_widths
            ):

                if not current_group:

                    current_group = [
                        column_index
                    ]

                    current_width = width

                    continue

                if (
                    len(current_group)
                    >= maximum_columns_per_group
                    or
                    (
                        current_width + width
                        > preferred_group_width
                    )
                ):

                    column_groups.append(
                        current_group
                    )

                    current_group = [
                        column_index
                    ]

                    current_width = width

                else:

                    current_group.append(
                        column_index
                    )

                    current_width += width

            if current_group:

                column_groups.append(
                    current_group
                )

            print(
                "HORIZONTAL GROUPING ACTIVE:",
                max_column,
                column_groups
            )

            # -------------------------------------------------
            # Remove completely empty worksheet columns before
            # rendering horizontal groups.
            # -------------------------------------------------

            non_empty_columns = []

            for column in range(
                1,
                max_column + 1
            ):

                has_value = False

                for row in range(
                    1,
                    max_row + 1
                ):

                    value = worksheet.cell(
                        row=row,
                        column=column
                    ).value

                    if (
                        value is not None
                        and str(value).strip() != ""
                    ):

                        has_value = True
                        break

                if has_value:

                    non_empty_columns.append(
                        column - 1
                    )


            filtered_column_groups = []

            for group in column_groups:

                filtered_group = [
                    column_index
                    for column_index in group
                    if column_index in non_empty_columns
                ]

                if filtered_group:

                    filtered_column_groups.append(
                        filtered_group
                    )


            # -------------------------------------------------
            # Render each horizontal group.
            # The same worksheet rows are used for every group.
            # -------------------------------------------------

            for group_number, column_group in enumerate(
                filtered_column_groups
            ):

                # -------------------------------------------------
                # Create the horizontal group table.
                # -------------------------------------------------

                group_table = create_worksheet_table(
                    worksheet,
                    value_worksheet,
                    max_row,
                    max_column,
                    available_width,
                    column_group=column_group
                )

                # -------------------------------------------------
                # Start every horizontal group after the first
                # group on a fresh PDF page.
                # -------------------------------------------------

                if group_number > 0:

                    story.append(
                        PageBreak()
                    )

                story.append(
                    group_table
                )

                if (
                    group_number
                    < len(filtered_column_groups) - 1
                ):

                    story.append(
                        Spacer(1, 8)
                    )
            continue

        # -------------------------------------------------
        # Render every logical table as its own flowable.
        #
        # Because each table is separate, ReportLab will
        # move the whole table to the next page when the
        # complete table fits there.
        # -------------------------------------------------

        for index, heading_row in enumerate(heading_rows):

            if index + 1 < len(heading_rows):

                next_heading_row = heading_rows[index + 1]

                section_end_row = (
                    next_heading_row - 1
                )

            else:

                section_end_row = max_row

            # -------------------------------------------------
            # Determine the useful columns for this table only.
            #
            # The heading row itself is ignored when deciding
            # the last useful column. This prevents a heading
            # merged farther right than the actual table data
            # from creating an unnecessary blank column.
            # -------------------------------------------------

            section_max_column = 1

            for row in range(
                heading_row + 1,
                section_end_row + 1
            ):

                for column in range(
                    1,
                    max_column + 1
                ):

                    cell = worksheet.cell(
                        row=row,
                        column=column
                    )

                    if cell.value is not None:
                        section_max_column = max(
                            section_max_column,
                            column
                        )

            # Preserve merged cells belonging to the table,
            # but ignore the heading row's oversized merge.

            for merged_range in worksheet.merged_cells.ranges:

                if (
                    merged_range.max_row < heading_row + 1
                    or merged_range.min_row > section_end_row
                ):
                    continue

                top_left_cell = worksheet.cell(
                    row=merged_range.min_row,
                    column=merged_range.min_col
                )

                if (
                    isinstance(top_left_cell.value, str)
                    and top_left_cell.value.lstrip().startswith("(")
                ):
                    continue

                section_max_column = max(
                    section_max_column,
                    merged_range.max_col
                )

            section_max_column = min(
                section_max_column,
                max_column
            )

           
            

            # -------------------------------------------------
            # Important:
            #
            # create_worksheet_table() normally starts from
            # row 1. We need a table containing only this
            # section.
            #
            # Therefore this version temporarily uses the
            # worksheet rows belonging to the section.
            # -------------------------------------------------

            section_rows = []

            for row in range(
                heading_row,
                section_end_row + 1
            ):

                row_data = []

                for column in range(
                    1,
                    section_max_column + 1
                ):

                    cell = worksheet.cell(
                        row=row,
                        column=column
                    )

                    value_cell = value_worksheet.cell(
                        row=row,
                        column=column
                    )

                    value = value_cell.value

                    if (
                        value is None
                        and isinstance(cell.value, str)
                        and cell.value.startswith("=")
                    ):
                        value = cell.value

                    merged_info = None

                    for merged_range in worksheet.merged_cells.ranges:

                        if (
                            merged_range.min_row
                            <= row
                            <= merged_range.max_row
                            and
                            merged_range.min_col
                            <= column
                            <= merged_range.max_col
                        ):
                            merged_info = (
                                merged_range.min_row,
                                merged_range.min_col,
                                merged_range.max_row,
                                merged_range.max_col
                            )
                            break

                    if merged_info:

                        min_row, min_column, _, _ = merged_info

                        if (
                            row != min_row
                            or column != min_column
                        ):
                            value = ""

                    # Create the paragraph normally.
                    paragraph = make_cell_paragraph(
                        cell,
                        value
                    )

                    # --------------------------------------------------------
                    # Center Inventory title text inside the full title span.
                    # --------------------------------------------------------

                    if (
                        row in (2, 4)
                        and column == 2
                        and value is not None
                    ):

                        title_style = ParagraphStyle(
                            name=f"InventoryTitle_{row}",
                            fontName=get_font_name(cell),
                            fontSize=get_font_size(cell),
                            leading=get_font_size(cell) + 2,
                            textColor=(
                                excel_color_to_reportlab(
                                    cell.font.color
                                ) or colors.black
                            ),
                            alignment=TA_CENTER,
                            spaceBefore=0,
                            spaceAfter=0,
                            leftIndent=0,
                            rightIndent=0,
                            firstLineIndent=0,
                        )

                        paragraph = Paragraph(
                            str(value),
                            title_style
                        )

                    row_data.append(
                        paragraph
                    )

                section_rows.append(
                    row_data
                )

                        # -------------------------------------------------
            # Build automatic horizontal column groups.
            #
            # Each group contains the Excel columns that fit
            # naturally within the available PDF width.
            #
            # We do NOT scale all columns down to one page.
            # Instead, wide worksheets are split horizontally.
            # -------------------------------------------------

            section_column_widths = []

            # -------------------------------------------------
            # Ignore completely empty leading columns.
            # These should not consume PDF page width.
            # -------------------------------------------------

            first_display_column = 1

            while (
                first_display_column <= section_max_column
            ):

                has_value = False

                for row in range(
                    heading_row,
                    section_end_row + 1
                ):

                    value = worksheet.cell(
                        row=row,
                        column=first_display_column
                    ).value

                    if (
                        value is not None
                        and str(value).strip() != ""
                    ):

                        has_value = True
                        break

                if has_value:
                    break

                first_display_column += 1


            for column in range(
                first_display_column,
                section_max_column + 1
            ):

                letter = get_column_letter(
                    column
                )

                width = worksheet.column_dimensions[
                    letter
                ].width

                if width is None:

                    width = 8.43

                section_column_widths.append(
                    excel_width_to_points(
                        width
                    )
                )

                                   # -------------------------------------------------
            # Build horizontal column groups.
            #
            # A group can contain up to 5 columns.
            #
            # The actual Excel column widths are also checked,
            # so unusually wide columns can create an earlier
            # group break.
            # -------------------------------------------------

            preferred_group_width = (
                available_width * 0.70
            )

            maximum_columns_per_group = 5

            

            section_column_groups = []

            current_group = []
            current_width = 0

            for column_index, width in enumerate(
                section_column_widths
            ):

                if not current_group:

                    current_group = [
                        column_index
                    ]

                    current_width = width

                    continue

                # -------------------------------------------------
                # Start a new group when either:
                #
                # 1. The group reaches 5 columns, OR
                # 2. Adding another column would make it too wide.
                # -------------------------------------------------

                if (
                    len(current_group)
                    >= maximum_columns_per_group
                    or
                    (
                        current_width + width
                        > preferred_group_width
                    )
                ):

                    section_column_groups.append(
                        current_group
                    )

                    current_group = [
                        column_index
                    ]

                    current_width = width

                else:

                    current_group.append(
                        column_index
                    )

                    current_width += width

            if current_group:

                section_column_groups.append(
                    current_group
                )

                print(
                    "HORIZONTAL GROUPING ACTIVE:",
                    section_max_column,
                    section_column_groups
                )

            # -------------------------------------------------
            # Safety fallback.
            #
            # There should always be at least one group when
            # the section contains columns.
            # -------------------------------------------------

            if not section_column_groups:

                section_column_groups = [
                    list(
                        range(
                            section_max_column
                        )
                    )
                ]

            # -------------------------------------------------
            # Calculate row heights using ALL column groups.
            #
            # The same row height will be used by every
            # horizontal group so the rows remain synchronized
            # between pages.
            # -------------------------------------------------

            section_row_heights = []

            for local_row, source_row in enumerate(
                range(
                    heading_row,
                    section_end_row + 1
                )
            ):

                height = worksheet.row_dimensions[
                    source_row
                ].height

                required_height = 14

                # -------------------------------------------------
                # Check every column in every horizontal group.
                #
                # This uses the actual Paragraph wrapping and
                # the actual width of that Excel column.
                # -------------------------------------------------

                for group in section_column_groups:

                    for column_index in group:

                        paragraph = section_rows[
                            local_row
                        ][
                            column_index
                        ]

                        if paragraph is None:

                            continue

                        column_width = (
                            measurement_column_widths[
                                column_index
                            ]
                        )

                        try:

                            _, paragraph_height = (
                                paragraph.wrap(
                                    column_width - 10,
                                    1000
                                )
                            )

                            required_height = max(
                                required_height,
                                paragraph_height + 8
                            )

                        except Exception:

                            pass

                # -------------------------------------------------
                # Preserve an explicit Excel row height when it
                # is larger than the automatically required height.
                # -------------------------------------------------

                if height is None:

                    section_row_heights.append(
                        max(
                            14,
                            required_height
                        )
                    )

                else:

                    section_row_heights.append(
                        max(
                            14,
                            float(height),
                            required_height
                        )
                    )
                        # -------------------------------------------------
            # Create one PDF table for each horizontal
            # column group.
            #
            # The same Excel rows are used in every group.
            # This creates the horizontal page flow used by
            # PDF converters such as iLovePDF.
            # -------------------------------------------------

            for group_number, column_group in enumerate(
                section_column_groups
            ):

                # -------------------------------------------------
                # Build the rows for this column group only.
                # -------------------------------------------------

                group_rows = []

                for local_row in range(
                    len(section_rows)
                ):

                    group_row = []

                    for column_index in column_group:

                        group_row.append(
                            section_rows[
                                local_row
                            ][
                                column_index
                            ]
                        )

                    group_rows.append(
                        group_row
                    )

                # -------------------------------------------------
                # Get the widths belonging to this group.
                # -------------------------------------------------

                group_column_widths = [
                    section_column_widths[
                        column_index
                    ]
                    for column_index in column_group
                ]

                # -------------------------------------------------
                # Scale this horizontal group to use the full
                # available PDF page width.
                # -------------------------------------------------

                group_total_width = sum(
                    group_column_widths
                )

                if (
                    group_total_width > 0
                    and group_total_width < available_width
                ):

                    group_scale = (
                        available_width
                        / group_total_width
                    )

                    group_column_widths = [
                        width * group_scale
                        for width in group_column_widths
                    ]

                # -------------------------------------------------
                # Create the PDF table for this group.
                # -------------------------------------------------

                section_pdf_table = Table(
                    group_rows,
                    colWidths=group_column_widths,
                    rowHeights=section_row_heights,
                    hAlign="LEFT",
                    splitByRow=1,
                )

                commands = []

                # -------------------------------------------------
                # Grid
                # -------------------------------------------------

                commands.append(
                    (
                        "GRID",
                        (0, 0),
                        (-1, -1),
                        0.5,
                        colors.black
                    )
                )

                # -------------------------------------------------
                # Cell formatting
                # -------------------------------------------------

                for local_row, source_row in enumerate(
                    range(
                        heading_row,
                        section_end_row + 1
                    )
                ):

                    for group_column, column_index in enumerate(
                        column_group
                    ):

                        cell = worksheet.cell(
                            row=source_row,
                            column=column_index + 1
                        )

                        pdf_row = local_row
                        pdf_column = group_column

                        # -------------------------------------------------
                        # Background
                        # -------------------------------------------------

                        fill_color = None

                        try:

                            if cell.fill.fill_type:

                                fill_color = (
                                    excel_color_to_reportlab(
                                        cell.fill.fgColor
                                    )
                                )

                        except Exception:

                            fill_color = None

                        if fill_color is not None:

                            commands.append(
                                (
                                    "BACKGROUND",
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    fill_color
                                )
                            )

                        # -------------------------------------------------
                        # Horizontal alignment
                        # -------------------------------------------------

                        horizontal = (
                            cell.alignment.horizontal
                        )

                        if horizontal == "center":

                            commands.append(
                                (
                                    "ALIGN",
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    "CENTER"
                                )
                            )

                        elif horizontal == "right":

                            commands.append(
                                (
                                    "ALIGN",
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    "RIGHT"
                                )
                            )

                        else:

                            commands.append(
                                (
                                    "ALIGN",
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    (
                                        pdf_column,
                                        pdf_row
                                    ),
                                    "LEFT"
                                )
                            )

                        # -------------------------------------------------
                        # Vertical alignment
                        # -------------------------------------------------

                        commands.append(
                            (
                                "VALIGN",
                                (
                                    pdf_column,
                                    pdf_row
                                ),
                                (
                                    pdf_column,
                                    pdf_row
                                ),
                                get_vertical_alignment(cell)
                            )
                        )

                # -------------------------------------------------
                # Merged cells
                #
                # A merged cell can cross a horizontal group boundary.
                # We only span the portion that exists inside the
                # current group.
                # -------------------------------------------------

                actual_group_columns = len(
                    column_group
                )

                for merged_range in worksheet.merged_cells.ranges:

                    # Ignore merges completely outside this section.
                    if (
                        merged_range.max_row < heading_row
                        or merged_range.min_row > section_end_row
                    ):
                        continue

                    merge_min_column = (
                        merged_range.min_col - 1
                    )

                    merge_max_column = (
                        merged_range.max_col - 1
                    )

                    # -------------------------------------------------
                    # Find which columns from the Excel merged range
                    # are present in this PDF group.
                    # -------------------------------------------------

                    group_positions = []

                    for position, column_index in enumerate(
                        column_group
                    ):

                        if (
                            merge_min_column
                            <= column_index
                            <= merge_max_column
                        ):

                            group_positions.append(
                                position
                            )

                    if not group_positions:
                        continue

                    # -------------------------------------------------
                    # Row coordinates inside this section table.
                    # -------------------------------------------------

                    min_row = (
                        max(
                            merged_range.min_row,
                            heading_row
                        )
                        - heading_row
                    )

                    max_row_range = (
                        min(
                            merged_range.max_row,
                            section_end_row
                        )
                        - heading_row
                    )

                    if min_row > max_row_range:
                        continue

                    group_min_column = min(
                        group_positions
                    )

                    group_max_column = max(
                        group_positions
                    )

                    # -------------------------------------------------
                    # A one-cell span is unnecessary.
                    # -------------------------------------------------

                    if (
                        group_min_column
                        == group_max_column
                        and
                        min_row
                        == max_row_range
                    ):
                        continue

                    commands.append(
                        (
                            "SPAN",
                            (
                                group_min_column,
                                min_row
                            ),
                            (
                                group_max_column,
                                max_row_range
                            )
                        )
                    )

                # -------------------------------------------------
                # Center Inventory title inside the current
                # horizontal group.
                #
                # The title is repeated for each group so every
                # horizontal page section remains understandable.
                # -------------------------------------------------

                for title_row in (2, 4):

                    if not (
                        heading_row
                        <= title_row
                        <= section_end_row
                    ):
                        continue

                    title_excel_column = 2

                    title_group_position = None

                    for position, column_index in enumerate(
                        column_group
                    ):

                        if (
                            column_index
                            == title_excel_column - 1
                        ):

                            title_group_position = position

                            break

                    if title_group_position is None:
                        continue

                    commands.append(
                        (
                            "ALIGN",
                            (
                                0,
                                title_row - heading_row
                            ),
                            (
                                actual_group_columns - 1,
                                title_row - heading_row
                            ),
                            "CENTER"
                        )
                    )

                    commands.append(
                        (
                            "VALIGN",
                            (
                                0,
                                title_row - heading_row
                            ),
                            (
                                actual_group_columns - 1,
                                title_row - heading_row
                            ),
                            "MIDDLE"
                        )
                    )

                # -------------------------------------------------
                # Padding
                # -------------------------------------------------

                commands.extend(
                    [
                        (
                            "LEFTPADDING",
                            (0, 0),
                            (-1, -1),
                            5
                        ),
                        (
                            "RIGHTPADDING",
                            (0, 0),
                            (-1, -1),
                            5
                        ),
                        (
                            "TOPPADDING",
                            (0, 0),
                            (-1, -1),
                            4
                        ),
                        (
                            "BOTTOMPADDING",
                            (0, 0),
                            (-1, -1),
                            4
                        ),
                    ]
                )

                section_pdf_table.setStyle(
                    TableStyle(commands)
                )

                # -------------------------------------------------
                # Start every horizontal group after the first
                # group on a fresh PDF page.
                # -------------------------------------------------

                if group_number > 0:

                    story.append(
                        PageBreak()
                    )

                story.append(
                    section_pdf_table
                )

                # -------------------------------------------------
                # Separate horizontal groups with a little space.
                # -------------------------------------------------

                if (
                    group_number
                    < len(section_column_groups) - 1
                ):

                    story.append(
                        Spacer(1, 8)
                    )
    document.build(story)

    return output.getvalue()