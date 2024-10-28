from kivy.app import App
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.textinput import TextInput
from kivy.uix.spinner import Spinner
from kivy.uix.scrollview import ScrollView
from datetime import datetime
import sqlite3
import pandas as pd
import os


class Database:
    def __init__(self):
        self.conn = sqlite3.connect('color_quality.db')
        self.create_table()

    def create_table(self):
        self.conn.execute('''
            CREATE TABLE IF NOT EXISTS quality_data (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                color_code TEXT,
                date TEXT,
                tank TEXT,
                batch TEXT,
                tsc REAL,
                viscosity REAL,
                dl REAL,
                da REAL,
                db REAL,
                strength REAL,
                disposition TEXT,
                remarks TEXT,
                batch_pigment TEXT,
                adjustments TEXT,
                color_group TEXT
            )
        ''')
        self.conn.commit()

    def insert_data(self, data):
        query = '''INSERT INTO quality_data 
                  (color_code, date, tank, batch, tsc, viscosity, dl, da, db, 
                   strength, disposition, remarks, batch_pigment, adjustments, color_group)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)'''
        self.conn.execute(query, data)
        self.conn.commit()

    def update_data(self, id, data):
        query = '''UPDATE quality_data SET
                  color_code=?, date=?, tank=?, batch=?, tsc=?, viscosity=?, dl=?, da=?, db=?,
                  strength=?, disposition=?, remarks=?, batch_pigment=?, adjustments=?, color_group=?
                  WHERE id=?'''
        self.conn.execute(query, data + (id,))
        self.conn.commit()

    def delete_data(self, id):
        self.conn.execute('DELETE FROM quality_data WHERE id=?', (id,))
        self.conn.commit()

    def get_all_data(self):
        cursor = self.conn.execute('SELECT * FROM quality_data')
        return cursor.fetchall()

    def search_data(self, criteria):
        query = 'SELECT * FROM quality_data WHERE ' + criteria
        cursor = self.conn.execute(query)
        return cursor.fetchall()


class DataEntryForm(GridLayout):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.cols = 2
        self.padding = 10
        self.spacing = 10

        # Add form fields
        self.add_widget(Label(text='Color Code:'))
        self.color_code = Spinner(
            text='Black',
            values=('Black', 'Blue', 'Violet', 'Orange', 'Yellow', 'Green')
        )
        self.add_widget(self.color_code)

        self.add_widget(Label(text='Date:'))
        self.date = TextInput(text=datetime.now().strftime('%Y-%m-%d'))
        self.add_widget(self.date)

        # Add other form fields similarly
        fields = ['Tank', 'Batch', 'TSC-%', 'Viscosity-CPS', 'dL', 'da', 'db',
                  '% Strength SUM']
        self.inputs = {}
        for field in fields:
            self.add_widget(Label(text=field))
            self.inputs[field] = TextInput()
            self.add_widget(self.inputs[field])

        self.add_widget(Label(text='Disposition:'))
        self.disposition = Spinner(text='Passed', values=('Passed', 'Rework'))
        self.add_widget(self.disposition)

        self.add_widget(Label(text='Remarks:'))
        self.remarks = TextInput(multiline=True)
        self.add_widget(self.remarks)

        self.add_widget(Label(text='Batch Pigment:'))
        self.batch_pigment = TextInput(multiline=True)
        self.add_widget(self.batch_pigment)

        self.add_widget(Label(text='Adjustments:'))
        self.adjustments = TextInput(multiline=True)
        self.add_widget(self.adjustments)


class MainScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.db = Database()

        layout = BoxLayout(orientation='vertical', padding=10, spacing=10)

        # Add form
        self.form = DataEntryForm()
        layout.add_widget(self.form)

        # Add buttons
        buttons = BoxLayout(size_hint_y=0.2)
        buttons.add_widget(Button(text='Save', on_press=self.save_data))
        buttons.add_widget(Button(text='Search', on_press=self.show_search))
        buttons.add_widget(Button(text='Export', on_press=self.export_to_excel))
        layout.add_widget(buttons)

        self.add_widget(layout)

    def save_data(self, instance):
        data = (
            self.form.color_code.text,
            self.form.date.text,
            self.form.inputs['Tank'].text,
            self.form.inputs['Batch'].text,
            float(self.form.inputs['TSC-%'].text or 0),
            float(self.form.inputs['Viscosity-CPS'].text or 0),
            float(self.form.inputs['dL'].text or 0),
            float(self.form.inputs['da'].text or 0),
            float(self.form.inputs['db'].text or 0),
            float(self.form.inputs['% Strength SUM'].text or 0),
            self.form.disposition.text,
            self.form.remarks.text,
            self.form.batch_pigment.text,
            self.form.adjustments.text,
            self.form.color_code.text  # Using color code as color group
        )
        self.db.insert_data(data)

    def export_to_excel(self, instance):
        data = self.db.get_all_data()
        df = pd.DataFrame(data, columns=[
            'ID', 'Color Code', 'Date', 'Tank', 'Batch', 'TSC', 'Viscosity',
            'dL', 'da', 'db', 'Strength', 'Disposition', 'Remarks',
            'Batch Pigment', 'Adjustments', 'Color Group'
        ])

        # Create Excel writer
        writer = pd.ExcelWriter('color_quality_report.xlsx', engine='xlsxwriter')

        # Split data by color and save to different sheets
        for color in ['Black', 'Blue', 'Violet', 'Orange', 'Yellow', 'Green']:
            color_data = df[df['Color Code'] == color]
            if not color_data.empty:
                color_data.to_excel(writer, sheet_name=color, index=False)

        writer.save()


class ColorQualityApp(App):
    def build(self):
        sm = ScreenManager()
        sm.add_widget(MainScreen(name='main'))
        return sm


if __name__ == '__main__':
    ColorQualityApp().run()