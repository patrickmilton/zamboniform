
import sys

try:
	from Tkinter import *
except ImportError:
	from tkinter import *

try:
	import ttk
	py3 = 0
except ImportError:
	import tkinter.ttk as ttk
	py3 = 1

import time
import xlsxwriter
import zamtest_support
import os.path

from PIL import Image, ImageTk

global banner
global bannerPhoto

def vp_start_gui():

	global val, w, root
	root = Tk()
	zamtest_support.set_Tk_var()
	top = Zam_Form (root)
	zamtest_support.init(root, top)
	root.mainloop()

w = None
def create_Zam_Form(root, *args, **kwargs):

	global w, w_win, rt
	rt = root
	w = Toplevel (root)
	zamtest_support.set_Tk_var()
	top = Zam_Form (w)
	zamtest_support.init(w, top, *args, **kwargs)
	return (w, top)

def destroy_Zam_Form():

	global w
	w.destroy()
	w = None




class Zam_Form:

	def __init__(self, top=None):

		_bgcolor = '#d9d9d9'  # X11 color: 'gray85'
		_fgcolor = '#000000'  # X11 color: 'black'
		_compcolor = '#d9d9d9' # X11 color: 'gray85'
		_ana1color = '#d9d9d9' # X11 color: 'gray85'
		_ana2color = '#d9d9d9' # X11 color: 'gray85'
		self.style = ttk.Style()
		if sys.platform == "win32":
			self.style.theme_use('winnative')
		self.style.configure('.',background=_bgcolor)
		self.style.configure('.',foreground=_fgcolor)
		self.style.map('.',background=
			[('selected', _compcolor), ('active',_ana2color)])
		top.geometry("819x659+573+178")
		top.title("Resurface Log")
		top.configure(background="#d9d9d9")

		self.editing = False
		##Ice Complex Banner
		script_dir = os.path.abspath(os.path.dirname(sys.argv[0]))
		banner = Image.open(os.path.join(script_dir, 'banner.png'))
		bannerPhoto = ImageTk.PhotoImage(banner)
		self.label1 = Label(top, image=bannerPhoto)
		self.label1.image = bannerPhoto
		self.label1.place(relx=0.01, rely=0.02, height=60, width=300)
		self.label1.configure(activebackground="#f9f9f9")
		self.label1.configure(activeforeground="black")
		self.label1.configure(anchor=CENTER)
		self.label1.configure(background="#d9d9d9")
		self.label1.configure(disabledforeground="#a3a3a3")
		self.label1.configure(foreground="#000000")
		self.label1.configure(highlightbackground="#d9d9d9")
		self.label1.configure(highlightcolor="black")
		self.label1.configure(width=154)
		#-----------------------------

		#Flood Checkbutton
		self.Flood = Checkbutton(top)
		self.Flood.place(relx=0.78, rely=0.3, relheight=0.04
				, relwidth=0.21)
		self.Flood.configure(activebackground="#d9d9d9")
		self.Flood.configure(activeforeground="#000000")
		self.Flood.configure(anchor=W)
		self.Flood.configure(background="#d9d9d9")
		self.Flood.configure(disabledforeground="#a3a3a3")
		self.Flood.configure(foreground="#000000")
		self.Flood.configure(highlightbackground="#d9d9d9")
		self.Flood.configure(highlightcolor="black")
		self.Flood.configure(justify=LEFT)
		self.Flood.configure(text='''Flood''')
		self.Flood.configure(variable=zamtest_support.flood)
		self.Flood.configure(width=171)
		self.Flood.deselect()
		#---------------------------

		#Dry Checkbutton
		self.Dry = Checkbutton(top)
		self.Dry.place(relx=0.78, rely=0.38, relheight=0.04
				, relwidth=0.22)
		self.Dry.configure(activebackground="#d9d9d9")
		self.Dry.configure(activeforeground="#000000")
		self.Dry.configure(anchor=W)
		self.Dry.configure(background="#d9d9d9")
		self.Dry.configure(disabledforeground="#a3a3a3")
		self.Dry.configure(foreground="#000000")
		self.Dry.configure(highlightbackground="#d9d9d9")
		self.Dry.configure(highlightcolor="black")
		self.Dry.configure(justify=LEFT)
		self.Dry.configure(text='''Dry''')
		self.Dry.configure(variable=zamtest_support.dry)
		self.Dry.configure(width=181)
		self.Dry.deselect()
		#---------------------------

		#Edge Checkbutton
		self.Edge = Checkbutton(top)
		self.Edge.place(relx=0.78, rely=0.46, relheight=0.04
				, relwidth=0.21)
		self.Edge.configure(activebackground="#d9d9d9")
		self.Edge.configure(activeforeground="#000000")
		self.Edge.configure(anchor=W)
		self.Edge.configure(background="#d9d9d9")
		self.Edge.configure(disabledforeground="#a3a3a3")
		self.Edge.configure(foreground="#000000")
		self.Edge.configure(highlightbackground="#d9d9d9")
		self.Edge.configure(highlightcolor="black")
		self.Edge.configure(justify=LEFT)
		self.Edge.configure(text='''Edge''')
		self.Edge.configure(variable=zamtest_support.edge)
		self.Edge.configure(width=176)
		self.Edge.deselect()
		#---------------------------

		#3 Lap Checkbutton
		self.ThreeLap = Checkbutton(top)
		self.ThreeLap.place(relx=0.78, rely=0.34, relheight=0.04
				, relwidth=0.21)
		self.ThreeLap.configure(activebackground="#d9d9d9")
		self.ThreeLap.configure(activeforeground="#000000")
		self.ThreeLap.configure(anchor=W)
		self.ThreeLap.configure(background="#d9d9d9")
		self.ThreeLap.configure(disabledforeground="#a3a3a3")
		self.ThreeLap.configure(foreground="#000000")
		self.ThreeLap.configure(highlightbackground="#d9d9d9")
		self.ThreeLap.configure(highlightcolor="black")
		self.ThreeLap.configure(justify=LEFT)
		self.ThreeLap.configure(text='''3-lap''')
		self.ThreeLap.configure(variable=zamtest_support.threeLap)
		self.ThreeLap.configure(width=171)
		self.ThreeLap.deselect()
		#---------------------------

		#Recent Resurfaces Checkbutton
		self.Scrolledlistbox1 = ScrolledListBox(top)
		self.Scrolledlistbox1.place(relx=0.12, rely=0.68, relheight=0.3
				, relwidth=0.76)
		self.Scrolledlistbox1.configure(background="white")
		self.Scrolledlistbox1.configure(disabledforeground="#a3a3a3")
		self.Scrolledlistbox1.configure(font="TkFixedFont")
		self.Scrolledlistbox1.configure(foreground="black")
		self.Scrolledlistbox1.configure(highlightbackground="#d9d9d9")
		self.Scrolledlistbox1.configure(highlightcolor="#d9d9d9")
		self.Scrolledlistbox1.configure(selectbackground="#c4c4c4")
		self.Scrolledlistbox1.configure(selectforeground="black")
		self.Scrolledlistbox1.configure(width=10)
		self.Scrolledlistbox1.configure(listvariable=zamtest_support.Recent_Resurfaces)
		self.Scrolledlistbox1.bind('<FocusOut>', lambda e: self.Scrolledlistbox1.selection_clear(0,END))
		#---------------------------

		#Board Brush Checkbutton
		self.Brush = Checkbutton(top)
		self.Brush.place(relx=0.78, rely=0.18, relheight=0.04
				, relwidth=0.21)
		self.Brush.configure(activebackground="#d9d9d9")
		self.Brush.configure(activeforeground="#000000")
		self.Brush.configure(anchor=W)
		self.Brush.configure(background="#d9d9d9")
		self.Brush.configure(disabledforeground="#a3a3a3")
		self.Brush.configure(foreground="#000000")
		self.Brush.configure(highlightbackground="#d9d9d9")
		self.Brush.configure(highlightcolor="black")
		self.Brush.configure(justify=LEFT)
		self.Brush.configure(text='''Board Brush''')
		self.Brush.configure(variable=zamtest_support.boardBrush)
		self.Brush.configure(width=171)
		self.Brush.deselect()
		#---------------------------------

		#Wash water Checkbutton
		self.Flood0 = Checkbutton(top)
		self.Flood0.place(relx=0.78, rely=0.22, relheight=0.04
				, relwidth=0.21)
		self.Flood0.configure(activebackground="#d9d9d9")
		self.Flood0.configure(activeforeground="#000000")
		self.Flood0.configure(anchor=W)
		self.Flood0.configure(background="#d9d9d9")
		self.Flood0.configure(disabledforeground="#a3a3a3")
		self.Flood0.configure(foreground="#000000")
		self.Flood0.configure(highlightbackground="#d9d9d9")
		self.Flood0.configure(highlightcolor="black")
		self.Flood0.configure(justify=LEFT)
		self.Flood0.configure(text='''Wash Water''')
		self.Flood0.configure(variable=zamtest_support.wash)
		self.Flood0.configure(width=171)
		self.Flood0.deselect()
		#---------------------------------

		#Wet Checkbutton
		self.Wet = Checkbutton(top)
		self.Wet.place(relx=0.78, rely=0.26, relheight=0.04
				, relwidth=0.21)
		self.Wet.configure(activebackground="#d9d9d9")
		self.Wet.configure(activeforeground="#000000")
		self.Wet.configure(anchor=W)
		self.Wet.configure(background="#d9d9d9")
		self.Wet.configure(disabledforeground="#a3a3a3")
		self.Wet.configure(foreground="#000000")
		self.Wet.configure(highlightbackground="#d9d9d9")
		self.Wet.configure(highlightcolor="black")
		self.Wet.configure(justify=LEFT)
		self.Wet.configure(text='''Wet''')
		self.Wet.configure(variable=zamtest_support.wet)
		self.Wet.configure(width=171)
		self.Wet.deselect()
		#---------------------------

		#Center Flood Checkbutton
		self.CenterFlood = Checkbutton(top)
		self.CenterFlood.place(relx=0.78, rely=0.42, relheight=0.04
				, relwidth=0.21)
		self.CenterFlood.configure(activebackground="#d9d9d9")
		self.CenterFlood.configure(activeforeground="#000000")
		self.CenterFlood.configure(anchor=W)
		self.CenterFlood.configure(background="#d9d9d9")
		self.CenterFlood.configure(disabledforeground="#a3a3a3")
		self.CenterFlood.configure(foreground="#000000")
		self.CenterFlood.configure(highlightbackground="#d9d9d9")
		self.CenterFlood.configure(highlightcolor="black")
		self.CenterFlood.configure(justify=LEFT)
		self.CenterFlood.configure(text='''Center Flood''')
		self.CenterFlood.configure(variable=zamtest_support.centerFlood)
		self.CenterFlood.configure(width=171)
		self.CenterFlood.deselect()
		#---------------------------

		#Dump tank Label
		self.Label3 = Label(top)
		self.Label3.place(relx=0.0, rely=0.2, height=21, width=154)
		self.Label3.configure(anchor=E)
		self.Label3.configure(background="#d9d9d9")
		self.Label3.configure(disabledforeground="#a3a3a3")
		self.Label3.configure(foreground="#000000")
		self.Label3.configure(text='''Dump Tank Level''')
		self.Label3.configure(width=154)
		#---------------------------

		#Temp Label
		self.Label4 = Label(top)
		self.Label4.place(relx=0.0, rely=0.26, height=21, width=154)
		self.Label4.configure(activebackground="#f9f9f9")
		self.Label4.configure(activeforeground="black")
		self.Label4.configure(anchor=E)
		self.Label4.configure(background="#d9d9d9")
		self.Label4.configure(disabledforeground="#a3a3a3")
		self.Label4.configure(foreground="#000000")
		self.Label4.configure(highlightbackground="#d9d9d9")
		self.Label4.configure(highlightcolor="black")
		self.Label4.configure(text='''Temperature''')
		self.Label4.configure(width=154)
		#---------------------------

		#Temp/Humid Label
		self.Label5 = Label(top)
		self.Label5.place(relx=0.0, rely=0.32, height=21, width=154)
		self.Label5.configure(activebackground="#f9f9f9")
		self.Label5.configure(activeforeground="black")
		self.Label5.configure(anchor=E)
		self.Label5.configure(background="#d9d9d9")
		self.Label5.configure(disabledforeground="#a3a3a3")
		self.Label5.configure(foreground="#000000")
		self.Label5.configure(highlightbackground="#d9d9d9")
		self.Label5.configure(highlightcolor="black")
		self.Label5.configure(text='''Temp / Humidity''')
		self.Label5.configure(width=154)
		#---------------------------

		#Comment Label
		self.Label6 = Label(top)
		self.Label6.place(relx=0.0, rely=0.44, height=21, width=154)
		self.Label6.configure(activebackground="#f9f9f9")
		self.Label6.configure(activeforeground="black")
		self.Label6.configure(anchor=E)
		self.Label6.configure(background="#d9d9d9")
		self.Label6.configure(disabledforeground="#a3a3a3")
		self.Label6.configure(foreground="#000000")
		self.Label6.configure(highlightbackground="#d9d9d9")
		self.Label6.configure(highlightcolor="black")
		self.Label6.configure(text='''Comment''')
		self.Label6.configure(width=154)
		#---------------------------

		#Initials Label
		self.Label7 = Label(top)
		self.Label7.place(relx=0.0, rely=0.38, height=21, width=154)
		self.Label7.configure(activebackground="#f9f9f9")
		self.Label7.configure(activeforeground="black")
		self.Label7.configure(anchor=E)
		self.Label7.configure(background="#d9d9d9")
		self.Label7.configure(disabledforeground="#a3a3a3")
		self.Label7.configure(foreground="#000000")
		self.Label7.configure(highlightbackground="#d9d9d9")
		self.Label7.configure(highlightcolor="black")
		self.Label7.configure(text='''Zam Driver Initials''')
		self.Label7.configure(width=154)
		#---------------------------

		#Dump Tank Entry
		self.Entry3 = Entry(top)
		self.Entry3.place(relx=0.21, rely=0.2, relheight=0.03, relwidth=0.2)
		self.Entry3.configure(background="white")
		self.Entry3.configure(disabledforeground="#a3a3a3")
		self.Entry3.configure(font="TkFixedFont")
		self.Entry3.configure(foreground="#000000")
		self.Entry3.configure(highlightbackground="#d9d9d9")
		self.Entry3.configure(highlightcolor="black")
		self.Entry3.configure(insertbackground="black")
		self.Entry3.configure(selectbackground="#c4c4c4")
		self.Entry3.configure(selectforeground="black")
		self.Entry3.configure(width=164)
		self.Entry3.configure(textvariable=zamtest_support.dumpStr)
		#---------------------------

		#Temp Entry
		self.Entry4 = Entry(top)
		self.Entry4.place(relx=0.21, rely=0.26, relheight=0.03, relwidth=0.2)
		self.Entry4.configure(background="white")
		self.Entry4.configure(disabledforeground="#a3a3a3")
		self.Entry4.configure(font="TkFixedFont")
		self.Entry4.configure(foreground="#000000")
		self.Entry4.configure(highlightbackground="#d9d9d9")
		self.Entry4.configure(highlightcolor="black")
		self.Entry4.configure(insertbackground="black")
		self.Entry4.configure(selectbackground="#c4c4c4")
		self.Entry4.configure(selectforeground="black")
		self.Entry4.configure(width=164)
		self.Entry4.configure(textvariable=zamtest_support.tempStr)
		#---------------------------

		#Temp/Humid Entry
		self.Entry5 = Entry(top)
		self.Entry5.place(relx=0.21, rely=0.32, relheight=0.03, relwidth=0.2)
		self.Entry5.configure(background="white")
		self.Entry5.configure(disabledforeground="#a3a3a3")
		self.Entry5.configure(font="TkFixedFont")
		self.Entry5.configure(foreground="#000000")
		self.Entry5.configure(highlightbackground="#d9d9d9")
		self.Entry5.configure(highlightcolor="black")
		self.Entry5.configure(insertbackground="black")
		self.Entry5.configure(selectbackground="#c4c4c4")
		self.Entry5.configure(selectforeground="black")
		self.Entry5.configure(width=164)
		self.Entry5.configure(textvariable=zamtest_support.humidStr)
		#---------------------------

		#Initials Entry
		self.Entry7 = Entry(top)
		self.Entry7.place(relx=0.21, rely=0.38, relheight=0.03, relwidth=0.2)
		self.Entry7.configure(background="white")
		self.Entry7.configure(disabledforeground="#a3a3a3")
		self.Entry7.configure(font="TkFixedFont")
		self.Entry7.configure(foreground="#000000")
		self.Entry7.configure(highlightbackground="#d9d9d9")
		self.Entry7.configure(highlightcolor="black")
		self.Entry7.configure(insertbackground="black")
		self.Entry7.configure(selectbackground="#c4c4c4")
		self.Entry7.configure(selectforeground="black")
		self.Entry7.configure(width=164)
		self.Entry7.configure(textvariable=zamtest_support.initStr)
		#---------------------------

		#Comment Entry
		self.Entry6 = Entry(top)
		self.Entry6.place(relx=0.21, rely=0.44, relheight=0.03, relwidth=0.2)
		self.Entry6.configure(background="white")
		self.Entry6.configure(disabledforeground="#a3a3a3")
		self.Entry6.configure(font="TkFixedFont")
		self.Entry6.configure(foreground="#000000")
		self.Entry6.configure(highlightbackground="#d9d9d9")
		self.Entry6.configure(highlightcolor="black")
		self.Entry6.configure(insertbackground="black")
		self.Entry6.configure(selectbackground="#c4c4c4")
		self.Entry6.configure(selectforeground="black")
		self.Entry6.configure(width=164)
		self.Entry6.configure(textvariable=zamtest_support.commStr)
		#---------------------------

		#Submit Button
		self.Button1 = Button(top)
		self.Button1.place(relx=0.67, rely=0.56, height=34, width=157)
		self.Button1.configure(activebackground="#d9d9d9")
		self.Button1.configure(activeforeground="#000000")
		self.Button1.configure(background="#d9d9d9")
		self.Button1.configure(disabledforeground="#a3a3a3")
		self.Button1.configure(foreground="#000000")
		self.Button1.configure(highlightbackground="#d9d9d9")
		self.Button1.configure(highlightcolor="black")
		self.Button1.configure(pady="0")
		self.Button1.configure(text='''Submit''')
		self.Button1.configure(width=157)
		self.Button1.configure(command=self.writeResurface)
		#---------------------------


		#Export Button
		#self.ExportButton = Button(top)
		#self.ExportButton.place(relx=0.40, rely=0.56, height=34, width=157)
		#self.ExportButton.configure(activebackground="#d9d9d9")
		#self.ExportButton.configure(activeforeground="#000000")
		#self.ExportButton.configure(background="#d9d9d9")
		#self.ExportButton.configure(disabledforeground="#a3a3a3")
		#self.ExportButton.configure(foreground="#000000")
		#self.ExportButton.configure(highlightbackground="#d9d9d9")
		#self.ExportButton.configure(highlightcolor="black")
		#self.ExportButton.configure(pady="0")
		#self.ExportButton.configure(text="Export")
		#self.ExportButton.configure(width=157)
		#self.ExportButton.configure(command =lambda : self.getExportData(self.Scrolledlistbox1.get(0,END)))
		#---------------------------

		#Recent Resurface label
		self.Label8 = Label(top)
		self.Label8.place(relx=0.12, rely=0.64, height=21, width=174)
		self.Label8.configure(background="#d9d9d9")
		self.Label8.configure(disabledforeground="#a3a3a3")
		self.Label8.configure(foreground="#000000")
		self.Label8.configure(text='''Recent Resurfaces''')
		self.Label8.configure(width=174)

		#Rink1 Checkbutton
		self.Rink1 = Checkbutton(top)
		self.Rink1.place(relx=0.1, rely=0.13, relheight=0.04
				, relwidth=0.07)
		self.Rink1.configure(activebackground="#d9d9d9")
		self.Rink1.configure(activeforeground="#000000")
		self.Rink1.configure(background="#d9d9d9")
		self.Rink1.configure(disabledforeground="#a3a3a3")
		self.Rink1.configure(foreground="#000000")
		self.Rink1.configure(highlightbackground="#d9d9d9")
		self.Rink1.configure(highlightcolor="black")
		self.Rink1.configure(justify=LEFT)
		self.Rink1.configure(text='''Rink 1''')
		self.Rink1.configure(variable=zamtest_support.rink1)
		self.Rink1.deselect()
		#---------------------------

		#Rink2 Checkbutton
		self.Rink2 = Checkbutton(top)
		self.Rink2.place(relx=0.26, rely=0.13, relheight=0.04
				, relwidth=0.07)
		self.Rink2.configure(activebackground="#d9d9d9")
		self.Rink2.configure(activeforeground="#000000")
		self.Rink2.configure(background="#d9d9d9")
		self.Rink2.configure(disabledforeground="#a3a3a3")
		self.Rink2.configure(foreground="#000000")
		self.Rink2.configure(highlightbackground="#d9d9d9")
		#self.Rink2.configure(highlightcolor="black")
		self.Rink2.configure(justify=LEFT)
		#self.Rink2.configure(state="n)
		self.Rink2.configure(text='''Rink 2''')
		self.Rink2.configure(variable=zamtest_support.rink2)
		self.Rink2.deselect()
		#---------------------------

		#Edit Button
		self.Edit = Button(top)
		self.Edit.place(relx=0.90, rely=0.68, height=34, width=67)
		self.Edit.configure(activebackground="#d9d9d9")
		self.Edit.configure(activeforeground="#000000")
		self.Edit.configure(background="#d9d9d9")
		self.Edit.configure(disabledforeground="#a3a3a3")
		self.Edit.configure(foreground="#000000")
		self.Edit.configure(highlightbackground="#d9d9d9")
		self.Edit.configure(highlightcolor="black")
		self.Edit.configure(pady="0")
		self.Edit.configure(text='''Edit''')
		self.Edit.configure(width=67)
		self.Edit.configure(command = lambda: self.editSelect())
		#---------------------------


		#Replace Button
		self.Replace = Button(top)
		self.Replace.place(relx=0.90, rely=0.76, height=34, width=67)
		self.Replace.configure(activebackground="#d9d9d9")
		self.Replace.configure(activeforeground="#000000")
		self.Replace.configure(background="#d9d9d9")
		self.Replace.configure(disabledforeground="#a3a3a3")
		self.Replace.configure(foreground="#000000")
		self.Replace.configure(highlightbackground="#d9d9d9")
		self.Replace.configure(highlightcolor="black")
		self.Replace.configure(pady="0")
		self.Replace.configure(text='''Replace''')
		self.Replace.configure(width=67)
		self.Replace.configure(command = lambda: self.replace())
		#---------------------------


		#Delete Button
		self.Delete = Button(top)
		self.Delete.place(relx=0.90, rely=0.83, height=34, width=67)
		self.Delete.configure(activebackground="#d9d9d9")
		self.Delete.configure(activeforeground="#000000")
		self.Delete.configure(background="#d9d9d9")
		self.Delete.configure(disabledforeground="#a3a3a3")
		self.Delete.configure(foreground="#000000")
		self.Delete.configure(highlightbackground="#d9d9d9")
		self.Delete.configure(highlightcolor="black")
		self.Delete.configure(pady="0")
		self.Delete.configure(text='''Delete''')
		self.Delete.configure(width=67)
		self.Delete.configure(command=lambda: self.deletCurr() )
		#---------------------------


		#Date label
		self.label1 = Label(top)
		self.label1.place(relx=0.6, rely=0.02, height=21, width=50)
		self.label1.configure(background="#d9d9d9")
		self.label1.configure(disabledforeground="#a3a3a3")
		self.label1.configure(foreground="#000000")
		self.label1.configure(text='''Date''')
		self.label1.configure(width=50)
		#---------------------------

		#Date Entry
		self.date = Entry(top)
		self.date.place(relx=0.66, rely=0.02, relheight=0.03, relwidth=0.1)
		self.date.configure(background="white")
		self.date.configure(disabledforeground="#595959")
		self.date.configure(font="TkFixedFont")
		self.date.configure(foreground="#000000")
		self.date.configure(insertbackground="black")
		self.date.configure(state=DISABLED)
		self.date.configure(width=150)
		self.date.configure(textvariable = zamtest_support.dateStr)
		#---------------------------

		#Time Label
		self.Label2 = Label(top)
		self.Label2.place(relx=0.77, rely=0.02, height=21, width=33)
		self.Label2.configure(background="#d9d9d9")
		self.Label2.configure(disabledforeground="#a3a3a3")
		self.Label2.configure(foreground="#000000")
		self.Label2.configure(text='''Time''')
		#---------------------------

		#Time Entry
		self.timeEnt = Entry(top)
		self.timeEnt.place(relx=0.82, rely=0.02, relheight=0.03, relwidth=0.08)
		self.timeEnt.configure(background="white")
		self.timeEnt.configure(disabledforeground="#595959")
		self.timeEnt.configure(font="TkFixedFont")
		self.timeEnt.configure(foreground="#000000")
		self.timeEnt.configure(insertbackground="black")
		self.timeEnt.configure(state=DISABLED)
		self.timeEnt.configure(width=100)
		self.timeEnt.configure(textvariable = zamtest_support.timeStr)
		#---------------------------


		self.getTime()

		#callbacks for modifying time and date see : editSelect()
		zamtest_support.dateStr.trace("w", self.getTime)
		zamtest_support.timeStr.trace("w", self.getTime)

		self.menubar = Menu(top,font="TkMenuFont",bg=_bgcolor,fg=_fgcolor)
		top.configure(menu = self.menubar)





	#function EDIT:
	def editSelect(self):
		#----------------
		## Description: Edits the selected line and
		## re checks options based on line in listbox
		#----------------

		self.editing = True

		self.date.configure(state = NORMAL)
		zamtest_support.dateStr.set("")

		self.timeEnt.configure(state = NORMAL)
		zamtest_support.timeStr.set("")

		try:
			#get selected entry
			selection = self.Scrolledlistbox1.curselection()

			#parse entry into array
			value = self.Scrolledlistbox1.get(selection[0])
			line = self.parseResurface(value)

		#set label1 and time
			zamtest_support.dateStr.set(line[0])
			zamtest_support.timeStr.set(line[1])

		#Clearing all checkboxes
			self.Flood.deselect()
			self.Dry.deselect()
			self.Edge.deselect()
			self.ThreeLap.deselect()
			self.Wet.deselect()
			self.Flood0.deselect()
			self.CenterFlood.deselect()
			self.Rink1.deselect()
			self.Rink2.deselect()
			self.Brush.deselect()
		#if element in array reselect checkbutton /
			c = 0
			if line[2] == "Rink1":
				zamtest_support.rink1.set(True)
				#print("blah")
			elif line[2] == "Rink2":
				zamtest_support.rink2.set(True)
			if line[3] == "Brush":
				zamtest_support.boardBrush.set(True)
			if line[4] == "Wash":
				zamtest_support.wash.set(True)
			if line[5] == "Wet":
				zamtest_support.wet.set(True)
			if line[6] == "Dry":
				zamtest_support.dry.set(True)
			if line[7] == "Edged":
				zamtest_support.edge.set(True)
			if line[8] == "Three Lap":
				zamtest_support.threeLap.set(True)
			if line[9] == "Flood":
				zamtest_support.flood.set(True)
			if line[10] == "Center Flood":
				zamtest_support.centerFlood.set(True)

			# re enter element into entry box
			zamtest_support.dumpStr.set(line[11])
			zamtest_support.tempStr.set(line[12])
			zamtest_support.humidStr.set(line[13])
			zamtest_support.initStr.set(line[14])
			zamtest_support.commStr.set(line[15])

		except IndexError:
			print("Nope")

	def replace(self):
		#----------------
		## Description: Replaces current selected list
		## with current options
		#----------------

		try:

			#get selected list
			selection = self.Scrolledlistbox1.curselection()
			self.Scrolledlistbox1.delete(selection[0])

			#Add date and time from entry
			resurfaceText = ""
			resurfaceText = resurfaceText + zamtest_support.dateStr.get() + " | "
			resurfaceText = resurfaceText + zamtest_support.timeStr.get() + " | "

			#check all check buttons and format
			if int(zamtest_support.rink1.get()) == 1:
				resurfaceText = resurfaceText + "Rink1 | "
			elif int(zamtest_support.rink2.get()) == 1:
				resurfaceText = resurfaceText + "Rink2 | "
			else:
				resurfaceText = resurfaceText + "0 | "
			if int(zamtest_support.boardBrush.get()) == 1:
				resurfaceText = resurfaceText + "Brush | "
			else:
				resurfaceText = resurfaceText + "0 | "
			if int(zamtest_support.wash.get()) == 1:
				resurfaceText = resurfaceText + "Wash | "
			else:
				resurfaceText = resurfaceText + "0 | "
			if int(zamtest_support.wet.get()) == 1:
				resurfaceText = resurfaceText + "Wet | 0 | 0 | 0 | "
				self.Dry.deselect()
			elif int(zamtest_support.dry.get()) == 1:
				resurfaceText = resurfaceText + "0 | Dry | "
				if int(zamtest_support.edge.get()) == 1:
					resurfaceText = resurfaceText + "Edged | "
				else:
					resurfaceText = resurfaceText + "0 | "
				if int(zamtest_support.threeLap.get()) == 1:
					resurfaceText = resurfaceText + "Three Lap | "
				else:
					resurfaceText = resurfaceText + "0 | "
			else:
				resurfaceText = resurfaceText + "0 | 0 | 0 | 0 | "
			if int(zamtest_support.flood.get()) == 1:
				resurfaceText = resurfaceText + "Flood | "
			else:
				resurfaceText = resurfaceText + "0 | "
			if int(zamtest_support.centerFlood.get()) == 1:
				resurfaceText = resurfaceText + "Center Flood | "
			else:
				resurfaceText = resurfaceText + "0 | "

			#Add entries and format
			resurfaceText = resurfaceText + self.Entry3.get() + " | "
			resurfaceText = resurfaceText + self.Entry4.get() + " | "
			resurfaceText = resurfaceText + self.Entry5.get() + " | "
			resurfaceText = resurfaceText + self.Entry7.get() + " | "
			resurfaceText = resurfaceText + self.Entry6.get() + " | "

			#Replace listbox string
			self.Scrolledlistbox1.insert(selection[0],resurfaceText)
			self.getTime()



		except IndexError:
			print("Nope")

	#funcion DeleteCurr:
	def deletCurr(self):
		#----------------
		## Description: Deletes current selected member of listbox
		#----------------

		#get slected list from scrolled listbox
		selection = self.Scrolledlistbox1.curselection()
		self.Scrolledlistbox1.delete(selection[0])


	def parseResurface(self,txt):
		#----------------
		## Description: Accepts string of resurface text and returns
		## list of indvidual elements
		#----------------

		line =  txt.split(" | ")

		return (line)

	def exportXls(self, exp):
		#---------------------------
		##Description: Accepts 2d list of all Recent Resurfaces
		## and writes them to xls document based on position in list
		#---------------------------
		timm = time.localtime()
		date = str(timm[1]) + "-" + str(timm[2]) + "-" + str(timm[0])
		workbook = xlsxwriter.Workbook(date + ".xlsx")
		worksheet = workbook.add_worksheet()
		#print("export xls")
		row = 0
		col = 0

		for i in exp:
			for l in i:
				worksheet.write(row,col, l)
				col += 1
			col = 0
			row += 1
		workbook.close()


	def getExportData(self,arr):
		#---------------------------
		##Description: Accepts list of strings from Recent resurfaces(scrolled listbox)
		## adds formats adds heading and
		#---------------------------
		resurfaceText =  arr

		heading = ["label1", "Time", "Rink", "Board Brush", "Wash Water", "Wet Cut", "Dry Cut","Edged",
				   "Three Lap", "Flood","Center Flood" , "Dump Tank", "HoneyWells", "Room Temp/Humidity", "Initials", "Comment"]

		#2d array for exporting
		exp = [[] for i in range(len(arr) + 1)]

		#iterator
		count = 1

	##        #set heading for xls
		exp[0] = heading

		#loop through arr
		for c in arr:
			line = self.parseResurface(c)
			exp[count] = line
			count += 1

		#export
		self.exportXls(exp)


	def writeTime(self):
		#---------------------------
		##Description: Gets and writes time (formatted)
		#---------------------------

		timm = time.localtime()
		label1 = str(timm[1]) + "/" + str(timm[2]) + "/" + str(timm[0])

		if timm[3] > 12:
			pmam = "PM"
			hour = timm[3] - 12
		else:
			pmam = "AM"
			hour = timm[3]
		if timm[4] < 10:
			minit = "0" + str(timm[4])
		else:
			minit = str(timm[4])

		timofday =  str(hour) + ":" +  minit
		label1andtime = label1 + " | " + timofday

		return(label1andtime + pmam + " | ")


	def getTime(self, *args):
		#---------------------------
		##Description: Gets and updates date and time to date and time entries
		## While editing boolean is false
		#---------------------------

		if self.editing == False:
			timm = time.localtime()
			date = str(timm[1]) + "/" + str(timm[2]) + "/" + str(timm[0])

			if timm[3] > 12:
				pmam = "PM"
				hour = timm[3] - 12
			else:
				pmam = "AM"
				hour = timm[3]
			if timm[4] < 10:
				minit = "0" + str(timm[4])
			else:
				minit = str(timm[4])

			timofday =  str(hour) + ":" +  minit + pmam

			self.date.configure(state = NORMAL)
			self.date.delete(0,END)
			self.date.insert(0,date)
			self.date.configure(state = DISABLED)


			self.timeEnt.configure(state = NORMAL)
			self.timeEnt.delete(0,END)
			self.timeEnt.insert(0,timofday)
			self.timeEnt.configure(state = DISABLED)

		elif self.editing == True:
			return

		else:
			print("wtf")

	#this commen
	def writeResurface(self):
		#---------------------------
		##Description: Gets resurface data from the entries and check boxes
		## formats, and adds them to the listbox
		#---------------------------

		resurfaceText=self.writeTime()

		#checks
		if int(zamtest_support.rink1.get()) == 1:
			resurfaceText = resurfaceText + "Rink1 | "
		elif int(zamtest_support.rink2.get()) == 1:
			resurfaceText = resurfaceText + "Rink2 | "
		else:
			resurfaceText = resurfaceText + "0 | "
		if int(zamtest_support.boardBrush.get()) == 1:
			resurfaceText = resurfaceText + "Brush | "
		else:
			resurfaceText = resurfaceText + "0 | "
		if int(zamtest_support.wash.get()) == 1:
			resurfaceText = resurfaceText + "Wash | "
		else:
			resurfaceText = resurfaceText + "0 | "
		if int(zamtest_support.wet.get()) == 1:
			resurfaceText = resurfaceText + "Wet | 0 | 0 | 0 | "
			self.Dry.deselect()
		elif int(zamtest_support.dry.get()) == 1:
			resurfaceText = resurfaceText + "0 | Dry | "
			if int(zamtest_support.edge.get()) == 1:
				resurfaceText = resurfaceText + "Edged | "
			else:
				resurfaceText = resurfaceText + "0 | "
			if int(zamtest_support.threeLap.get()) == 1:
				resurfaceText = resurfaceText + "Three Lap | "
			else:
				resurfaceText = resurfaceText + "0 | "
		else:
			resurfaceText = resurfaceText + "0 | 0 | 0 | 0 | "
		if int(zamtest_support.flood.get()) == 1:
			resurfaceText = resurfaceText + "Flood | "
		else:
			resurfaceText = resurfaceText + "0 | "
		if int(zamtest_support.centerFlood.get()) == 1:
			resurfaceText = resurfaceText + "Center Flood | "
		else:
			resurfaceText = resurfaceText + "0 | "



		#entries
		resurfaceText = resurfaceText + self.Entry3.get() + " | "
		resurfaceText = resurfaceText + self.Entry4.get() + " | "
		resurfaceText = resurfaceText + self.Entry5.get() + " | "
		resurfaceText = resurfaceText + self.Entry7.get() + " | "
		resurfaceText = resurfaceText + self.Entry6.get() + " | "


		#set resurface text += " %checks and %entries"
		self.Scrolledlistbox1.insert(END, resurfaceText)

		#clearing all the entries
		self.Entry3.delete(0,END)
		self.Entry4.delete(0,END)
		self.Entry5.delete(0,END)
		self.Entry6.delete(0,END)
		self.Entry7.delete(0,END)

		#deselecting all the buttons
		self.Flood.deselect()
		self.Dry.deselect()
		self.Edge.deselect()
		self.ThreeLap.deselect()
		self.Wet.deselect()
		self.Flood0.deselect()
		self.CenterFlood.deselect()
		self.Rink1.deselect()
		self.Rink2.deselect()
		self.Brush.deselect()

		self.editing = False
		self.getTime()
		self.exportXls(self.Scrolledlistbox1.get(0,END))


class AutoScroll(object):
	'''Configure the scrollbars for a widget.'''

	def __init__(self, master):
		#  Rozen. Added the try-except clauses so that this class
		#  could be used for scrolled entry widget for which vertical
		#  scrolling is not supported. 5/7/14.
		try:
			vsb = ttk.Scrollbar(master, orient='vertical', command=self.yview)
		except:
			pass
		hsb = ttk.Scrollbar(master, orient='horizontal', command=self.xview)

		#self.configure(yscrollcommand=_autoscroll(vsb),
		#    xscrollcommand=_autoscroll(hsb))
		try:
			self.configure(yscrollcommand=self._autoscroll(vsb))
		except:
			pass
		self.configure(xscrollcommand=self._autoscroll(hsb))

		self.grid(column=0, row=0, sticky='nsew')
		try:
			vsb.grid(column=1, row=0, sticky='ns')
		except:
			pass
		hsb.grid(column=0, row=1, sticky='ew')

		master.grid_columnconfigure(0, weight=1)
		master.grid_rowconfigure(0, weight=1)

		# Copy geometry methods of master  (taken from ScrolledText.py)
		if py3:
			methods = Pack.__dict__.keys() | Grid.__dict__.keys() \
				  | Place.__dict__.keys()
		else:
			methods = Pack.__dict__.keys() + Grid.__dict__.keys() \
				  + Place.__dict__.keys()

		for meth in methods:
			if meth[0] != '_' and meth not in ('config', 'configure'):
				setattr(self, meth, getattr(master, meth))

	@staticmethod
	def _autoscroll(sbar):
		'''Hide and show scrollbar as needed.'''
		def wrapped(first, last):
			first, last = float(first), float(last)
			if first <= 0 and last >= 1:
				sbar.grid_remove()
			else:
				sbar.grid()
			sbar.set(first, last)
		return wrapped

	def __str__(self):
		return str(self.master)

def _create_container(func):
	'''Creates a ttk Frame with a given master, and use this new frame to
	place the scrollbars and the widget.'''
	def wrapped(cls, master, **kw):
		container = ttk.Frame(master)
		return func(cls, container, **kw)
	return wrapped

class ScrolledListBox(AutoScroll, Listbox):
	'''A standard Tkinter Text widget with scrollbars that will
	automatically show/hide as needed.'''
	@_create_container
	def __init__(self, master, **kw):
		Listbox.__init__(self, master, **kw)
		AutoScroll.__init__(self, master)

if __name__ == '__main__':
	vp_start_gui()
	if zamtest_support.dry.get() == 1:
		print("1 Test")
