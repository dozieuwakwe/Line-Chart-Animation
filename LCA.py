import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import ctypes
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
import matplotlib
from matplotlib.animation import FuncAnimation

#Data input
def fileimport():
    global bcrdf, importlabel, b1
    while True:
        filename = filedialog.askopenfilename()
        if 'xls' not in filename:
            ctypes.windll.user32.MessageBoxW(0, 'Please import an Excel file', "Import Error", 16)
        else:
            break
    bcrdf = pd.read_excel(filename)
    importlabel.configure(text='File imported!')
    b1.configure(text='Change Data File')

intro='Welcome to the Line Chart Animator! Please make sure your input data is in the form of an Excel file containing a table in which the rows represent each unit of time and the columns represent each category tracked).'
ctypes.windll.user32.MessageBoxW(0, intro, "Welcome!", 64)

top = Tk(className=' Line Chart Animator')
importlabel=Label(top, text='')
importlabel.grid(row=1, column=1)
Label(top, text='Please import a data file and fill out the information below').grid(row=0,columnspan=2)
Label(top, text='Chart Title: ').grid(row=2)
Label(top, text='X-Axis Label: ').grid(row=3)
Label(top, text='Y-Axis Label: ').grid(row=4)
Label(top, text='Time Unit: ').grid(row=5)
Label(top, text='Timesteps: ').grid(row=6)

e1=Entry(top)
e2=Entry(top)
e3=Entry(top)
e4=Entry(top)
e5=Entry(top)

e1.grid(row=2, column=1)
e2.grid(row=3, column=1)
e3.grid(row=4, column=1)
e4.grid(row=5, column=1)
e5.grid(row=6, column=1)

play=False
bcrdf=pd.DataFrame()

def callback(self):
    global title, xlabel, ylabel, units, timesteps, play
    if bcrdf.empty:
        ctypes.windll.user32.MessageBoxW(0, 'Please import a data file', "Welcome!", 16)
    else:
        title=e1.get() 
        xlabel=e2.get()
        ylabel=e3.get()
        units=e4.get()
        if e5.get()=='':
            timesteps=1
        else:
            timesteps=int(e5.get())
        top.destroy()
        play=True

top.bind('<Return>', callback)

def quit():
    top.destroy()

b1 = Button(top, text = "Import Data File", width=15, command=fileimport)
b1.grid(row=1,column=0)
b2 = Button(top, text = "Enter Information", width=15)
b2.bind('<Button-1>', callback)
b2.grid(row=7,columnspan=2)

mainloop()

#Line chart animation
def prepare_data(df, steps):
    index=df.columns[0]
    df.index = df.index * steps
    last_idx = df.index[-1] + 1
    df_expanded = df.reindex(range(last_idx))
    df_expanded = df_expanded.interpolate()
    df_expanded = df_expanded.set_index(index)
    return df_expanded

def update(i):
    colors = plt.cm.tab20(range(20))
    y = bcrdf_expanded.iloc[:i]
    x = bcrdf_expanded.index[:i]
    ax.plot(x,y)
    ax.grid(True, axis='y', color='grey')
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.set_title(title, fontsize=20)
    
    if i>1:
        ax.set_ybound(lower=max(y.values[-1])/2,upper=max(y.values[-1]+10))
        ax.set_xbound(lower=max(x)/2,upper=max(x)*1.1)
        
        for j in ax.get_children():   
            if type(j)==matplotlib.text.Annotation:
                j.remove()
    
        for i in range(len(bcrdf_expanded.columns)):
            ax.annotate(bcrdf_expanded.columns[i],(x[-1],y.iloc[-1][i]))
    
if play==True:
    bcrdf_expanded = prepare_data(bcrdf,timesteps)
    fig = plt.Figure(figsize=(12, 8))
    ax = fig.add_subplot()
    anim = FuncAnimation(fig=fig, func=update, frames=len(bcrdf_expanded), interval=100, repeat=False)
    anim.save('Line Chart Animation.mp4')
    ctypes.windll.user32.MessageBoxW(0, 'The Line Chart animation has been successfully completed and saved!', "Success!", 64)

    #Playback/Output
    class Video(object):
        def __init__(self,path):
            self.path = path

        def play(self):
            from os import startfile
            startfile(self.path)

    class Movie_MP4(Video):
        type = "MP4"

    movie = Movie_MP4("Line Chart Animation.mp4")
    movie.play()
