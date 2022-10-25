"""
Segregate images and videos based on date of capture from the metadata into folders of Year -> month --> file_name_dd_time
"""
import os
import pathlib
import tkinter
import ctypes
import win32com.client
from tkinter import filedialog

ctypes.windll.shcore.SetProcessDpiAwareness(1)


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


def main(path: pathlib.Path, files) -> None:
    for file in files:
        if file.is_file() and file.parent == path:
            meta = ['Name', 'Size', 'Item type', 'Date modified', 'Date created']
            data = get_file_metadata(str(file.parent), str(file.name), meta)["Date created"]
            segregate(data, file)


if __name__ == '__main__':
    root = tkinter.Tk()
    root.withdraw()

    folder_path = filedialog.askdirectory(title="Open folder containing your media", initialdir=os.getcwd(), mustexist=True)
    abs_folder_path = pathlib.Path(folder_path).resolve()

    file_iter = pathlib.Path(str(folder_path)).glob("**/*")

    main(abs_folder_path, file_iter)
