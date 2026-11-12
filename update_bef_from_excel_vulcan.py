from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


# -----------------------------------------------------------------------------
# 1. BEF / Excel processing functions
# -----------------------------------------------------------------------------

# capture the group from the header, ie the text after BEGIN$DEF in bef file
BEGIN_DEF_RE = re.compile(r"^BEGIN\$DEF\s+(\S+)\s*$")
# capture the parameter name in the correct format eg 'search_radius   =   25.000000000000'
PARAM_RE = re.compile(r"^(\s*)([A-Za-z0-9_]+)(\s*=\s*)(.*?)(\r?\n)?$")


def read_parameters(
    excel_path: Path,
    sheet_name: str = "Parameters",
) -> dict[int, dict[str, Any]]:
    """Read BEF_Parameter_Name plus Run1/Run2/Run3 values from Excel.
    
    The returned object looks conceptually as follows. Outer dict key is run number,
    inner dict is parameters mapped to values from Excel. 
    
            {
                1: {"parameter1": parameter1_value, "parameter2": parameter2_value},
                2: {"parameter1": parameter1_value, "parameter2": parameter2_value},
                3: {"parameter1": parameter1_value, "parameter2": parameter2_value},
            }
        """
    # import Excel.
    # data_only=True returns the result of a formula, not the formula itself
    # read_only=True opens in memory efficient mode, as not modifying the Excel file. 
    workbook = load_workbook(excel_path, data_only=True, read_only=True)

    # Check worksheet exists. If not, list available sheets.
    if sheet_name not in workbook.sheetnames:
        raise ValueError(
            f"Worksheet {sheet_name!r} was not found. "
            f"Available sheets: {workbook.sheetnames}"
        )
    sheet = workbook[sheet_name]

    # Read the first row and build a dictionary that maps each column heading
    # to its Excel column number.
    headers = {
        cell.value: index
        for index, cell in enumerate(next(sheet.iter_rows()), start=1)
    }
    # check no essential columns are missing
    required = ["BEF_Parameter_Name", "Run1", "Run2", "Run3"]
    missing = [name for name in required if name not in headers]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")

    # Create an empty dictionary for each estimation run.
    # outer key = integer run number
    # inner key = string parameter name, inner value = any type (number, text, boolean, etc.)
    parameters: dict[int, dict[str, Any]] = {1: {}, 2: {}, 3: {}}

    # Loop over all rows AFTER the header row.
    # values_only=True to read values rather than openpyxl Cell objects.
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # -1 to change Excel counting (from 1) to Python counting (from 0)
        parameter_name = row[headers["BEF_Parameter_Name"] - 1]

        # Ignore any blank rows in the spreadsheet.
        if parameter_name is None or str(parameter_name).strip() == "":
            continue
        # Convert parameter name to text and strip whitespace.
        parameter_name = str(parameter_name).strip()

        # Store the value for each estimation run, only if Excel value not blank
        # ie if empty, then do not update parameter for this run 
        # TODO consider whether best to have this and not need to update all values in Excel, or to build bef from scratch
        for run in (1, 2, 3):
            value = row[headers[f"Run{run}"] - 1]
            if value is not None and str(value).strip() != "":
                parameters[run][parameter_name] = value

    workbook.close()
    return parameters


def format_value(value: Any, original_rhs: str) -> str:
    """Format an Excel value to resemble the existing fomat in the BEF."""
    # strip whitespace
    original_rhs = original_rhs.strip()

    # If the original BEF value is enclosed in single OR double quotes, preserve that quoting style.
    if (original_rhs.startswith("'") and original_rhs.endswith("'")) or (
        original_rhs.startswith('"') and original_rhs.endswith('"')
    ):
        # extract the quote type eg ' or "
        quote = original_rhs[0]
        # convert Excel value to string
        text = str(value)

        # if single quote, this can be escaped by doubling it
        # so clean the text and put the original quote type back in if so
        if quote == "'":
            text = text.replace("'", "''")
        return f"{quote}{text}{quote}"

    # check if bool type then convert to int with 12 dp in keepinh with numeric values in bef
    if isinstance(value, bool):
        return "1.000000000000" if value else "0.000000000000"

    # check if int, then ensure 12 dp in keeping with bef convention
    if isinstance(value, (int, float)):
        return f"{float(value):.12f}"
    
    # check if string. Numbers can be stored as strings by Excel. 
    # If so try to convert to number with 12 dp in keeping with bef convention
    try:
        return f"{float(str(value).strip()):.12f}"
    # If Python cannot convert (ie it's really a string), then keep as string
    except ValueError:
        return str(value)


def update_bef(
    input_bef: Path,
    output_bef: Path,
    parameters: dict[int, dict[str, Any]],
    variable: str = "fibre",
) -> dict[tuple[int, str], int]:
    """
    Read the original BEF, replace matching parameters for Runs 1-3,
    validate that every requested parameter matched exactly once,
    and write a NEW BEF file.

    The source BEF is deliberately left untouched."""

    # Safety check: never replace the source BEF. 
    # Path.resolve() converts each path to its absolute location.
    if input_bef.resolve() == output_bef.resolve():
        raise ValueError("Output BEF must be different from the input BEF.")

    # Read bef. 
    # splitlines(keepends=True) makes a list where each item is one line AND
    # retains the original newline characters. This makes it easier
    # to reproduce the original file formatting.
    lines = input_bef.read_text(encoding="utf-8").splitlines(keepends=True)

    # current_run records which estimation run we are currently in.
    # The syntax "int | None" is a type hint meaning either integer or None.
    # None is where we are outside of any run
    current_run: int | None = None

    # Initiate dictionary to count how many times every parameter is found.
    # Dictionary keys are tuples such as: (run_number, parameter) eg (1, "search_radius")
    # Dictionary values are match counts: 0 = not found (initial value), 1 = exactly right, 2+ = unexpectedly found more than once
    match_counts: dict[tuple[int, str], int] = {
        (run, name): 0
        for run, run_params in parameters.items()
        for name in run_params
    }

    # Initiate empty list of outputs
    output_lines: list[str] = []

    # for each line in the bef:
    for line in lines:
        # start trying to match each line from Excel to the bef, starting from the header BEGIN$DEF
        begin_match = BEGIN_DEF_RE.match(line.rstrip("\r\n"))

        # If there IS a match...
        if begin_match:
            # ...then return the first match as the section name
            # eg for BEGIN$DEF fibre_c1_type, section_name = fibre_c1_type
            section_name = begin_match.group(1)

            # Now work out from the section name whether it belongs to Run 1, 2 or 3.
            # rf matches both raw and formatted string from section_name
            # re.escape(variable) makes the variable text safe to insert inside a regular expression.
            # ([123]) captures exactly 1, 2 or 3 for the run
            # (?:_|$) is a non-capturing group: it groups the rest without creating another numbered group
            run_match = re.match(
                rf"^{re.escape(variable)}_c([123])(?:_|$)",
                section_name,
            )
            # if run_match exists, convert to int, else skip as it's not a relevant section
            # allows all lines in a real bed to be preserved, only updating relevant parameters. 
            current_run = int(run_match.group(1)) if run_match else None

            # append the current line to the output_lines variable and moved to the next line
            output_lines.append(line)
            continue

        # at end of section, desctivate the current run
        if line.startswith("END$DEF"):
            current_run = None
            output_lines.append(line)
            continue

        # Only match parameters while run is not empty 
        # and parameter and run are in the parameters dict from Excel
        if current_run is not None and current_run in parameters:
            param_match = PARAM_RE.match(line)

            # .groups() extracts values from PARAM_RE as a tuple.
            # eg search_radius = 25.000000000000\n
            # indent="    ", name="search_radius", separator="=", rhs="25.000000000000", newline="\n"
            if param_match:
                indent, name, separator, rhs, newline = param_match.groups()

                # if the extracted parameter is in the current run...
                if name in parameters[current_run]:
                    # ...count the number of matches...    
                    match_counts[(current_run, name)] += 1
                    # ...format the value from Excel to correct format for output bef
                    new_rhs = format_value(parameters[current_run][name], rhs)
                    # Rebuild the line using original indentation, parameter name etc etc, but assign new value
                    # "newline or ''" handles final line in file with no newline character.
                    line = f"{indent}{name}{separator}{new_rhs}{newline or ''}"
        # add original untouched line, or newly updated line
        output_lines.append(line)

    # Abort if Excel parameter was not found, or appeared more than once in one estimation run.
    missing = [key for key, count in match_counts.items() if count == 0]
    duplicates = [key for key, count in match_counts.items() if count > 1]
    if missing or duplicates:
        problems: list[str] = []
        if missing:
            problems.append(
                "not found: "
                + ", ".join(f"Run{run}:{name}" for run, name in missing)
            )
        if duplicates:
            problems.append(
                "matched more than once: "
                + ", ".join(
                    f"Run{run}:{name} ({match_counts[(run, name)]} matches)"
                    for run, name in duplicates
                )
            )
        raise ValueError(
            "BEF update aborted because parameter matching was not one-to-one; "
            + "; ".join(problems)
        )

    # Do not write anything until all validation has passed.
    # parents=True creates missing parent directories, exists=True then pass
    output_bef.parent.mkdir(parents=True, exist_ok=True)
    # Join all lines back into single string and save as UTF-8 text.
    output_bef.write_text("".join(output_lines), encoding="utf-8")

    # return match_counts for the GUI
    return match_counts


def create_bef_from_excel(
    excel_path: Path,
    input_bef: Path,
    output_bef: Path,
    sheet_name: str = "Parameters",
    variable: str = "fibre",
) -> dict[tuple[int, str], int]:
    """Convenience function joining the Excel-reading and BEF-writing steps."""

    #TODO - remove default/hard coding of sheet name and vaariable fibre

    # runner function for triggering the read_parameters and update_bef functions
    # First convert the Excel worksheet into a nested dict of parameters
    parameters = read_parameters(excel_path, sheet_name)
    # Then use that dictionary to make the new BEF.
    # "return" passes update_bef's match-count dictionary back to the caller (shortcut)
    return update_bef(input_bef, output_bef, parameters, variable)

# -----------------------------------------------------------------------------
# 2. Vulcan GUI
#    Only this function needs maptek.vulcan_gui.
#    Kept separate for debugging in VSCode
# -----------------------------------------------------------------------------

def run_vulcan_gui() -> None:
    """Show a Maptek Vulcan panel, collect the filenames from the user,
        run the BEF update, and display a summary."""

    # Import Maptek here so processing functions can still be debugged outside Vulcan.
    from maptek import vulcan_gui

    # Open a new Vulcan panel/window
    pan = vulcan_gui.panel("Create BEF from Excel Parameters")

    # Add explanatory text and a dividing line.
    pan.statictext("Select the input files and enter a name for the new BEF.")
    pan.staticline()

    # Add a file-selection field for the source/template BEF, limit to only bef files. 
    pan.fileselector(
        tag="input_bef",
        prompt="Template BEF file",
        filter="*.bef",
        width=50,
    )

    # Add a file-selection field for the source/template BEF, limit to only bef files. 
    pan.fileselector(
        tag="excel_file",
        prompt="Excel parameter file",
        filter="*.xlsx",
        width=50,
    )

    # Add field for output file name, with default provided.
    pan.editbox(
        tag="output_bef",
        prompt="Output BEF name",
        defaultVal="Updated_Estimation.bef",
        width=50,
    )

    # display panel
    button = pan.display()
    if not button:
        print("Script cancelled by user.")
        return

    # extract the values supplied to the panel (get values extracts a dictionary)
    values = pan.get_values()
    # Convert everything to strings and strip accidental spaces.
    input_text = str(values["input_bef"]).strip()
    excel_text = str(values["excel_file"]).strip()
    output_text = str(values["output_bef"]).strip()

    # provide error messages if required
    if not input_text:
        vulcan_gui.error("No input BEF file was selected.")
        return
    if not excel_text:
        vulcan_gui.error("No Excel parameter file was selected.")
        return
    if not output_text:
        vulcan_gui.error("No output BEF name was entered.")
        return

    # Convert panel values into pathlib Path objects.
    input_bef = Path(input_text)
    excel_path = Path(excel_text)
    output_path = Path(output_text)

    # If the user typed only a filename, put the result in the same folder as
    # the selected input BEF. A full path can also be entered if desired.
    if output_path.parent == Path("."):
        output_path = input_bef.parent / output_path
    # Add .bef automatically if the user omitted it.
    if output_path.suffix.lower() != ".bef":
        output_path = output_path.with_suffix(".bef")

    # try triggering the workflow with the 'runner' function
    try:
        match_counts = create_bef_from_excel(
            excel_path=excel_path,
            input_bef=input_bef,
            output_bef=output_path,
        )

    # Else show errors in Vulcan rather than losing them in the console.
    except Exception as exc:
        vulcan_gui.error(f"BEF was not created.\n\n{exc}")
        raise

    # display success message
    summary_lines = [f"BEF created successfully:\n{output_path}\n"]

    # Count successful bef updates separately for each run
    # "number of parameters successfully updated exactly once for this run".
    for run in (1, 2, 3):
        updated = sum(
            1
            for (r, _name), count in match_counts.items()
            if r == run and count == 1
        )
        # append number of successfully updated parameters per run into success message
        summary_lines.append(f"Run {run}: {updated} parameter(s) updated")

    # Get Vulcan to display success message
    vulcan_gui.message("\n".join(summary_lines))
    print("\n".join(summary_lines))

def main() -> None:
    run_vulcan_gui()

if __name__ == "__main__":
    main()
