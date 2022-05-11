import json
from typing import List, Optional
import time

import pandas as pd
import xlwings as xw
from pydantic import BaseModel

class LoggingError(Exception):
    """Custom error that is raised if error dict exists."""
    def __init__(self, title: str, message: str) -> None:
        self.title = title
        self.message = message
        super().__init__(message)

class Unit(BaseModel):
    """Represents data from a row in 'Bokningsblad'."""

    BOOKING: Optional[str]
    MLO: Optional[str]
    POL: Optional[str]
    TOL: Optional[str]
    CONTAINER: Optional[str]
    ISO_TYPE: Optional[str]
    NET_WEIGHT: Optional[float]
    POD_STATUS: Optional[str]
    LOAD_STATUS: Optional[str]
    VGM: Optional[float]
    OOG: Optional[str]
    REMARK: Optional[str]
    IMDG: Optional[float]
    UNNR: Optional[int]
    CHEM_REF: Optional[str]
    MRN: Optional[str]
    TEMP: Optional[str]
    PO_NUMBER: Optional[str]
    CUSTOMS_STATUS: Optional[str]
    PACKAGES: Optional[int]
    GOODS_DESCRIPTION: Optional[str]
    OCEAN_VESSEL: Optional[str]
    VOYAGE: Optional[str]
    ETA: Optional[str]
    FINAL_POD: Optional[str]


def main2() -> None:
    """Main2 function."""

    with open(r"C:\Users\SWV224\BOLLORE\XPF - Documents\MAINTENANCE\PYTHON\ELL-creation\src\ell-data.json") as fp:
        data = json.load(fp)
        print(data)
        rader: List[Unit] = [Unit(**item) for item in data]
        #print(rader)

def main() -> None:
    """Main function."""

    tic = time.perf_counter()

    #path = r"C:\Users\SWV224\BOLLORE\XPF - Documents\MAINTENANCE\Test files\0115_Bokningsblad_data.xlsb"

    #with xw.App(visible=False) as app:
    #    wb = app.books.open(path)
    #    rng = wb.sheets["INFO"].range('A4').expand()
    #    df = wb.sheets["INFO"].range(rng).options(pd.DataFrame, index=False, header=True).value

    wb = xw.Book.caller()
    rng = wb.sheets["INFO"].range('A4').expand()
    df = wb.sheets["INFO"].range(rng).options(pd.DataFrame, index=False, header=True).value

    result = df.to_json(orient='records')
    parsed = json.loads(result)

    rader: List[Unit] = [Unit(**item) for item in parsed]
    print(rader[0])
    toc = time.perf_counter()
    print(f'{toc - tic:0.3f} seconds')

if __name__ == "__main__":
    xw.Book(r"C:\Users\SWV224\BOLLORE\XPF - Documents\MAINTENANCE\Test files\0115_Bokningsblad_data.xlsb").set_mock_caller()
    main()