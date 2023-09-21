import tkinter as tk
from dataclasses import dataclass
from tkinter import filedialog
from typing import Optional

import polars
import polars as pl

window = tk.Tk()
window.title('Bookings Data App')

window.minsize(300, 200)
window.maxsize(300, 200)


@dataclass
class BookingsData:
    data: Optional[polars.DataFrame]


def load_data():
    file_path = filedialog.askopenfilename()
    df = pl.read_csv(file_path, separator='\t')
    button2['state'] = "normal"

    out = df.select(
        pl.col("Date Time").str.strptime(pl.Date, "%m/%d/%Y %l:%M %p", strict=False).alias("Date"),
        pl.col("Booking Id"),
        pl.col("Staff"),
        pl.col("Customer Name"),
        pl.col("Customer Email"),
        pl.col("Customer Phone"),
        pl.col(" Custom Fields").str.extract(r"^\{.*In which term would you like to start\?\": \"(\w+)\"",
                                             group_index=1).alias("Term"),
        pl.col(" Custom Fields").str.extract(r"^\{.*\"How did you hear about us\?\": \"(.+)\"",
                                             group_index=1).alias("How did you hear about us"),
        pl.col(" Custom Fields").str.extract(
            r"\"Please share anything that will help prepare for our meeting\.\"\: \"((?:\\\"|[^\"])+)",
            group_index=1).alias(
            "Please share anything that will help prepare for our meeting"),
    )

    booking_data.data = out


def save_data(df):
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx')
    if file_path:
        df.write_excel(file_path)
        button2['state'] = "disabled"


if __name__ == '__main__':
    booking_data = BookingsData(data=None)

    button1 = tk.Button(window, text='Load Data', command=load_data)
    button1.pack()

    button2 = tk.Button(window, text='Save Data', command=lambda: save_data(booking_data.data))
    button2.pack()
    button2['state'] = "disabled"

    window.mainloop()
