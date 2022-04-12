import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
from pathlib import Path
from PIL import Image, ImageTk
import regex as re

column_names = ['PG/IG', 'From Name', 'Related Office', 'Who From?',
                    'First Name', 'Last Name', 'Email', 'Job Title', 'Company',
                    'Business Country', 'Business Address 1', 'Business Address 2',
                    'Business Address 3', 'Zip Code', 'Phone Number'
                    ]



class Window(tk.Frame):
    def __init__(self, master = None):
        tk.Frame.__init__(self, master)
        self.master = master
        self.init_window()


    def init_window(self):

        #logo
        logo = Image.open('logo.png')
        logo = ImageTk.PhotoImage(logo)
        logo_label = tk.Label(image = logo)
        logo_label.image = logo
        logo_label.place(x=158, y=20)

        #instructions
        instructions = tk.Label(form, text="Select an appropriate '.xlsx' file for cleaning.")
        instructions.place(x=180, y=280)

        self.master.title("The Mailing List Cleaner by Izak Jessen")
        self.pack(fill = 'both', expand = 1)

        self.filepath = tk.StringVar()


        convertButton = tk.Button(self, text = 'Convert',
                               command = self.convert, bg="#ad6bff", fg="white", height="2", width="15")
        convertButton.place(x = 242, y = 200)

    def convert(self):

        def title_case(self):
            return self.str.title()

        def country_convert(self):
            if self.isupper():
                return self.upper()
            elif self.islower():
                return self.upper()
            else:
                return self.title()

        def name_convert(self):
            for i in str(self):
                d = list(self.split())
                if len(d) >= 2:
                    return d[1] + ', ' + d[0]
                break



        def lowercase(self):
            return self.lower()

        def find_email(self):
            email = re.findall(r'[\w\.-]+@[\w\.-]+',str(self))
            return ",".join(email)

        src_file = app.show_file_browser()

        df2 = pd.read_excel(src_file, header=8, usecols='A:O')

        df2.columns = column_names

        df2[['From Name', 'First Name', 'Last Name', 'Job Title']] = df2[
            ['From Name', 'First Name', 'Last Name', 'Job Title']].apply(title_case)

        df2['From Name'] = df2['From Name'].astype(str).apply(name_convert)


        df2.dropna(subset = ['Email'], inplace=True)
        df2.drop_duplicates(['Email'], inplace=True)
        df2['Email'] = df2['Email'].apply(lowercase)
        df2['Email'] = df2['Email'].apply(lambda x: find_email(x))


        df2.applymap(lambda x: x.strip() if isinstance(x, str) else x)

        df2 = df2.replace(r'\n',' ', regex=True)

        df2 = df2.replace(r'  ',' ', regex=True)

        df2['Business Country'] = df2['Business Country'].astype(str).apply(country_convert)

        c_file = Path.cwd() /  r'Country-List.xlsx'

        country_mapping = pd.read_excel(c_file)
        country_mapping.astype(str)

        maps = dict(zip(country_mapping['Country Name'], country_mapping['Country Code']))

        df2['Business Country'] = df2['Business Country'].map(maps)

        clean_data = df2.reindex(columns=['PG/IG', 'Related Office', 'From Name', 'First Name',
                         'Last Name', 'Email', 'Job Title', 'Company', 'Who From?',
                        'Business Country', 'Phone Number', 'Business Address 1', 'Business Address 2',
                        'Business Address 3',  'Zip Code'
                        ])

        clean_data.rename(columns = {'Who From?': 'Business City', 'PG/IG': 'SPB Contact',
        'Zip Code': 'Business Zip or Post Code'}, inplace=True)
        clean_data['Business State'] = np.nan
        clean_data['SPB Contact'] = np.nan
        clean_data['From Name'].dropna(inplace=True)

        clean_data.to_csv('Cleaned Mailing List.csv', encoding="utf-8-sig", index=0)

    def show_file_browser(self):
        self.filename = filedialog.askopenfilename()
        return self.filename

    def first_browser(self):
        file = self.show_file_browser()
        self.filepath.set(file)



form = tk.Tk()
form.geometry("600x400")
form.resizable(0, 0)


app = Window(form)

form.mainloop()
