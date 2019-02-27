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

def cadence():         #return cadence
    c=5                #cadence value
    return c

def power(j):           #return power
    p=10               #power value
    return p

def excel():
    #def returnCadence (num):

    save_path = 'D:/EDP Programming'
    #Variable used to name file number
    n=1

    #Searches for exixting file of same name
    match = os.path.exists("D:/EDP Programming/Cycling Session " + str(n)+".xlsx")

    while match==True:
        n=n+1
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

    i=1                #time iterator

    #take cadence and power values initially

    #write values in sheet
    #while cadence()!=0:
    sheet1.write(i, 0, i)
    sheet1.write(i, 1, cadence())
    sheet1.write(i, 2, power())
    i=i+1
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

    #returnCadence (time)

    #   return cadence


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.hi_there = tk.Button(self)
        self.hi_there["text"] = "START"
        self.hi_there["command"] = self.say_hi
        self.hi_there.pack(side="top")

        self.quit = tk.Button(self, text="QUIT AND SAVE", fg="red",
                              command=self.master.destroy)
        self.quit.pack(side="bottom")

    def say_hi(self):
        #while power()!=0:
        T = 60/cadence()               #time per revolution
        a = 6.283185/T       #angle for each updated power input
        #V = [x,y]
        #V = np.array([[1,1],[-2,2],[4,-7]])

        fig = plt.figure()
        origin = [0], [0] # origin point
        ax = plt.axes(xlim=(-1, 1), ylim=(-1,1))
        line, = ax.plot([], [], lw=2)
        #plt.quiver(*origin, x, y, color=['r'], scale=30)

        def init():
            line.set_data([],[])
            return line,

        def animate(j):
            x = math.sin(a)*power(j)           #x-component of vector
            y = math.cos(a)*power(j)           #y-component of vector
            line.set_data(x,y)
            return line,

        anim = animation.FuncAnimation(fig, animate, init_func=init, frames =200, interval=20, blit=True)

        plt.show()
            #time.sleep(2)
            #plt.close()
            #excel()

root = tk.Tk()
app = Application(master=root)
app.mainloop()
