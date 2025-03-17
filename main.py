from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.uix.button import Button
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.image import Image
import pandas as pd
import openpyxl
import matplotlib.pyplot as plt
import os

class FinancialApp(BoxLayout):
    def __init__(self, **kwargs):
        super().__init__(orientation='vertical', **kwargs)

        self.file_chooser = FileChooserListView()
        self.add_widget(self.file_chooser)

        self.upload_button = Button(text="Upload & Process Excel", size_hint=(1, 0.2))
        self.upload_button.bind(on_press=self.process_excel)
        self.add_widget(self.upload_button)

        self.result_label = Label(text="Select an Excel file to process", size_hint=(1, 0.2))
        self.add_widget(self.result_label)

        self.chart_image = Image(size_hint=(1, 0.5))
        self.add_widget(self.chart_image)

    def process_excel(self, instance):
        file_path = self.file_chooser.selection and self.file_chooser.selection[0]

        if file_path:
            try:
                df = pd.read_excel(file_path)

                df['Total'] = df['Opening Balance'] + df['Additional Funding'] + df['Settlement Amt']
                df['Net Balance'] = df['Closing Balance'] - df['Dispense']

                processed_file = "processed_data.xlsx"
                df.to_excel(processed_file, index=False, engine='openpyxl')

                self.generate_chart(df)

                self.result_label.text = f"File processed successfully!\nSaved as {processed_file}"
            except Exception as e:
                self.result_label.text = f"Error processing file: {str(e)}"
        else:
            self.result_label.text = "No file selected!"

    def generate_chart(self, df):
        plt.figure(figsize=(6, 4))
        plt.plot(df["Date"], df["Closing Balance"], marker="o", linestyle="-", color="b", label="Closing Balance")
        plt.plot(df["Date"], df["Net Balance"], marker="s", linestyle="--", color="r", label="Net Balance")
        plt.xlabel("Date")
        plt.ylabel("Amount")
        plt.title("Financial Trend")
        plt.legend()
        plt.grid(True)

        chart_file = "chart.png"
        plt.savefig(chart_file)
        self.chart_image.source = chart_file  

class MyApp(App):
    def build(self):
        return FinancialApp()

if __name__ == '__main__':
    MyApp().run()
