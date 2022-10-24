"""
Segregate images and videos based on date of capture from the metadata into folders of Year -> month --> file_name_dd_time
"""
import pathlib
import sys
import logging
import win32com.client

# TODO: Add GUI for media folder selection


def get_file_metadata(path: str, filename: str, metadata: list[str]) -> dict[str: str]:
    """Returns dictionary containing all file metadata."""
    sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
    ns = sh.NameSpace(path)

    # Enumeration is necessary because ns.GetDetailsOf only accepts an integer as 2nd argument
    file_metadata = dict()
    item = ns.ParseName(str(filename))
    for ind, attribute in enumerate(metadata):
        attr_value = ns.GetDetailsOf(item, ind)
        if attr_value:
            file_metadata[attribute] = attr_value

    return file_metadata


def segregate(data: str, file: pathlib.Path, years: list=[], months: list=[]):
    date = [data.split(" ")[0].split("-"), data.split(" ")[1].split(":")]
    day, month, year = date[0]

    # Create folder of years if it doesnt exist
    if year not in years:
        years.append(year)
        pathlib.Path(rf"{file.resolve().parent}\{year}").mkdir(parents=True, exist_ok=True)

    # Create folder of months if it doesnt exist
    if month not in months:
        months.append(month)
        pathlib.Path(rf"{file.resolve().parent}\{year}\{month}").mkdir(parents=True, exist_ok=True)

    # Moving and renaming the file
    file.rename(rf"{file.resolve().parent}\{year}\{month}\{file.stem}_{day}_{date[1][0]}{date[1][1]}{file.suffix}")


def main(path, files):
    for file in files:
        if file.is_file() and str(file.parent) == path:
            meta = ['Name', 'Size', 'Item type', 'Date modified', 'Date created']
            data = get_file_metadata(str(file.parent), str(file.name), meta)["Date created"]
            segregate(data, file)


if __name__ == '__main__':
    try:
        folder_path = sys.argv[1]
    except IndexError as e:
        logging.exception("Enter absolute filepath in quotes")
        exit(1)

    file_iter = pathlib.Path(folder_path).glob("**/*")
    main(folder_path, file_iter)
