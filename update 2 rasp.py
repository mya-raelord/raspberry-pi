from kivy.app import App
from kivy.graphics import Color, Rectangle
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.floatlayout import FloatLayout
from kivy.uix.image import AsyncImage
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.widget import Widget
from kivy.properties import ListProperty
import xlsxwriter
from xlsxwriter import*
import os.path
import string
import tkinter as tk
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import animation
import math
import time
import random

def cadence():         #return cadence
    c = 5                #cadence value
    return c

def power():           #return power
    p = random.randint(0,10)               #power value
    return p

def excel():
    #def returnCadence (num):

    save_path = 'D:/EDP Programming'
    #Variable used to name file number
    n=1

    #Searches for exixting file of same name
    match = os.path.exists("D:/EDP Programming/Cycling Session " + str(n)+".xlsx")

    while match==True:
        n = n + 1
        match = os.path.exists("D:/EDP Programming/Cycling Session " + str(n)+".xlsx")

    #Names File
    name_of_file = "Cycling Session " + str(n)
    completeName = os.path.join(save_path, name_of_file+".xlsx")

    # Workbook is created
    wb = xlsxwriter.Workbook(completeName)

    # add_sheet is used to create sheet.
    sheet1 = wb.add_worksheet('Cycling_Data')

    #sheet1.add_table('A1:B12')

    sheet1.write(0, 0, 'Time (s)')
    sheet1.write(0, 1, 'Cadence (rpm)')
    sheet1.write(0, 2, 'Power (W)')

    i = 1                #time iterator

    #take cadence and power values initially

    #write values in sheet
    #while cadence()!=0:
    sheet1.write(i, 0, i)
    sheet1.write(i, 1, cadence())
    sheet1.write(i, 2, power())
    i = i + 1
        #create loop here to update cadence and power


    #cadence graph
    cadenceChart = wb.add_chart({'type': 'line'})
    sheet1.insert_chart('E2', cadenceChart)

    cadenceChart.add_series({
        'categories': '=Cycling_Data!$A$2:$A$11',
        'values':     '=Cycling_Data!$B$2:$B$11',
        'line':       {'color': 'blue'},
    })

    cadenceChart.set_title({ 'name': 'Cadence (time)'})
    cadenceChart.set_x_axis({'name': 'Time (s)'})
    cadenceChart.set_y_axis({'name': 'Cadence (rpm)'})
    cadenceChart.set_legend({'none': True})

    #power graph
    powerChart = wb.add_chart({'type': 'line'})
    sheet1.insert_chart('M2', powerChart)

    powerChart.add_series({
        'categories': '=Cycling_Data!$A$2:$A$11',
        'values':     '=Cycling_Data!$C$2:$C$11',
        'line':       {'color': 'red'},
    })

    powerChart.set_title({ 'name': 'Power (time)'})
    powerChart.set_x_axis({'name': 'Time (s)'})
    powerChart.set_y_axis({'name': 'Power (W)'})
    powerChart.set_legend({'none': True})

    wb.close()

class start(FloatLayout):

    def __init__(self, **kwargs):
        # make sure we aren't overriding any important functionality
        super(start, self).__init__(**kwargs)
        self.add_widget(
            AsyncImage(
                source="D:/EDP Programming/carbon.jpg",
                size_hint= (2, 2),
                pos_hint={'center_x':.5, 'center_y':.5}))
        self.startbtn = Button(
                text="START",
                background_color=(0,1,0,1),
                size_hint=(.3, .3),
                pos_hint={'center_x': .5, 'center_y': .7})
        self.startbtn.bind(on_press=self.btn_pressedstart)
        self.add_widget(self.startbtn)

        self.quitbtn = Button(
                text="SAVE AND QUIT",
                background_color=(1,0,0,1),
                size_hint=(.3, .3),
                pos_hint={'center_x': .5, 'center_y': .3})
        self.quitbtn.bind(on_press=self.btn_pressedquit)
        self.add_widget(self.quitbtn)


        with self.canvas.before:
            Color(0, 0, 0, 0)  # green; colors range from 0-1 instead of 0-255
            self.rect = Rectangle(size=self.size, pos=self.pos)

        self.bind(size=self._update_rect, pos=self._update_rect)

    def _update_rect(self, instance, value):
        self.rect.pos = instance.pos
        self.rect.size = instance.size

    def btn_pressedstart(self, instance):
        T = 60/cadence()               #time per revolution
        a = 6.283185/T       #angle for each updated power input
        # V = [x,y]
        V = np.array([[1,1],[-2,2],[4,-7]])
        origin = [0], [0] # origin point
        fig = plt.figure()
        plt.quiver(*origin, V[:,0], V[:,1], color=['r'], scale=21)
        plt.show()
        excel()

    def btn_pressedquit(self, instance):
        self.destroy


class MainApp(App):

    def build(self):
        root = start()
        return root

if __name__ == '__main__':
    MainApp().run()
