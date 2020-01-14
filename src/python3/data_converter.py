import glob
import os

from pathlib import Path
from typing import Tuple


def read_input(in_file: Path) -> list:
    """Reads a file into a list."""
    read_list = [line.strip() for line in open(in_file, 'r')]
    return read_list


def write_output(out_data, out_file_path):
    """Write a list into a file."""
    with open(out_file_path, 'w') as f:
        for line in out_data:
            f.write("%s\n" % line)


def get_key_values(data_line):
    """Gets a data line and returns the key and the values."""
    key = data_line[:8].rstrip()
    values = data_line[8:].rstrip().split(",")

    return key, values


def convert_data(in_data):
    """Gets an input list, converts data and returns a list."""
    
    # These are the constant keys.
    IGNORED_KEYS = ["JOB", "INST", "UNITS", "XYZ", "HV"]
    USABLE_KEYS = ["STN", "ST", "BS", "SS", "SD"]

    out_data = []
    in_line_index = 0

    while True:
        # Checks if the current loop is the last one.
        if in_line_index > len(in_data) - 2:
            break

        # Extracts the key and the values from the current line.
        current_line = in_data[in_line_index]
        key, values = get_key_values(current_line)

        # Ignore certain keys.
        if key in IGNORED_KEYS:
            in_line_index += 1

        # The STN key enables the 1st phase where we read A and B values.
        elif key == USABLE_KEYS[0]:
            # Reads the two values of the STN.
            A1 = values[0]
            A2 = values[1]

            while True:
                # Moving to the next line on the input file.
                in_line_index += 1

                # Extracts the key and the values from the current line.
                current_line = in_data[in_line_index]
                key1, values1 = get_key_values(current_line)

                # If the key is BS then we write ST and move to 2nd phase.
                if key1 == USABLE_KEYS[2]:
                    B1 = values1[0]
                    out_data.append(",".join([USABLE_KEYS[1], A1, A2, B1]))
                    break

        # The BS or SS keys enable the 2nd phase where we read K and L values.
        elif (key == USABLE_KEYS[2]) or (key == USABLE_KEYS[3]):
            K1 = values[0]
            K2 = values[1]
            K3 = values[2]

            while True:
                # Moving to the next line on the input file.
                in_line_index += 1

                # Extracts the key and the values from the current line.
                current_line = in_data[in_line_index]
                key2, values2 = get_key_values(current_line)

                # If the key is SD then we write SS and move back to 1st phase.
                if key2 == USABLE_KEYS[4]:
                    L1 = values2[0]
                    L2 = values2[1]
                    L3 = values2[2]
                    out_data.append(",".join([USABLE_KEYS[3], K1, L1, L2, L3, K2, K3]))
                    in_line_index += 1
                    break
        else:
            # Moving to the next line on the input file.
            in_line_index += 1

    return out_data


def locate_in_out_dirs(script_file: str) -> Tuple[str, str]:
    """Locates the input and output directories relatively of the script path."""
    
    script_dir = Path(os.path.dirname(os.path.realpath(script_file)))
    in_dir = Path(script_dir / ".." / ".." / "input").resolve()
    out_dir = Path(script_dir / ".." / ".." / "output").resolve()
    if not in_dir.exists():
        print("Input directory not found: " + in_dir)
        exit(1)

    if not out_dir.exists():
        print("Output directory not found: " + out_dir)
        exit(2)

    return in_dir, out_dir


def main():
    # Locates the input and output files.
    in_dir, out_dir = locate_in_out_dirs(__file__)
    print("Reading input files from: " + str(in_dir))

    # Searches for input files within the input directory.
    search_mask = in_dir / "*.dat"
    in_files = map(Path, glob.glob(str(search_mask)))

    # Loops over each input file.
    for in_file in in_files:
        print("Converting file: " + str(in_file))

        # Reads the input data into a list.
        in_data = read_input(in_file)

        # Converts the data.
        out_data = convert_data(in_data)

        # Writes to the output file.
        out_file_path = Path(out_dir / str(in_file.stem + ".csv"))
        write_output(out_data, out_file_path)


if __name__ == '__main__':
    main()
