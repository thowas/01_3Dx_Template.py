import subprocess as sb
import win32com
import win32com.client
from pathlib import Path
import pythoncom
from win32com.client import VARIANT
from win32com.client import Dispatch
from pythoncom import *
import datetime
from datetime import date
import tkinter as tk
from tkinter import messagebox
from tkinter.messagebox import showinfo
from tkinter import filedialog as fd
from tkinter import ttk
import os
import psutil
from os import walk
import glob
import sys
import xml.etree.cElementTree as ET
import logging
import untangle
import shutil
from threading import Thread
import threading
import queue
import multiprocessing as mp
import time

try:
	
	date = datetime.datetime.now()
	y = date.strftime("%G")
	m = date.strftime("%m")
	d = date.strftime("%d")
	h = date.strftime("%H")
	mi = date.strftime("%M")
	proclist = []
	cDate = y + ("-") + m + ("-") + d
	expirationDate = "2023-12-30"
	# FilePath = ""
	swApp = None
	swModel = None
	version = "2.0.1"
	mode = 0o666
	xmlConfig = "configXML.xml"
	logUser = os.getlogin()
	deskPath = 'C:/Users/' + logUser + '/Desktop/'
	tempFolder = 'C:/Users/' + logUser + '/Desktop/SheetConfig'
	xmlFolder = 'C:/Users/' + logUser + '/Desktop/SheetConfig/XML'
	logPath = "C:/Users/" + logUser + "/Desktop/SheetConfig/LOG"
	assetPath = "C:/Users/" + logUser + "/Desktop/SheetConfig/icon"
	swxVersion = ""
	swxType = ""
	swxPID = False
	drt_path = ""
	str_Message = ""
	sheetconfig = []
	xml_path = None
	logFilename = m + "_" + y + "_" + h + "_" + mi + "_" + os.getlogin() + ".log"
	
	if not os.path.exists(deskPath):
		deskPath = 'C:/Temp/'
		tempFolder = 'C:/Temp/SheetConfig'
		xmlFolder = 'C:/Temp/SheetConfig/XML'
		logPath = 'C:/Temp/SheetConfig/LOG'
		assetPath = 'C:/Temp/SheetConfig/icon'
		if not os.path.exists(deskPath):
			os.mkdir('C:/Temp')
		else:
			print("Desktop Path: --- {s}".format(s=deskPath))
	else:
		print("Desktop Path: --- {s}".format(s=deskPath))
	
	if not os.path.exists(tempFolder):
		os.mkdir(tempFolder)
	elif not os.path.exists(logPath):
		os.mkdir(logPath)
		logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)-8s %(message)s',
		                    datefmt='%a, %d %b %Y %H:%M:%S', filename=logPath + '/' + logFilename,
		                    filemode='w')
	else:
		logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)-8s %(message)s',
		                    datefmt='%a, %d %b %Y %H:%M:%S', filename=logPath + '/' + logFilename,
		                    filemode='w')

except OSError as e:
	logging.error("Error: %s : %s" % (e, e.strerror))
	print("Error: %s : %s" % (e, e.strerror))
	sys.exit()

logging.info("##################### 12_Change_Sheet_Title Block ############################" + "\r\n" +
             "SWX Change_Sheet_Title Block Tool for: " + "\r\n" +
             "changes the sheetformat = the ""paper"" of your drawing for  " + "\r\n" +
             "all sheets of the active drawing. You have to adjust the path and  " + "\r\n" +
             "the file names to the new sheet formats. After successfully changing  " + "\r\n" +
             "the sheetformat the drawing is saved with its current name." + "\r\n"
                                                                             "" "\n")

logging.info("###################### 12_Change_Sheet_Title Block ###########################" + "\r\n")

logging.info("Log-USER: --- {s}".format(s=logUser))
logging.info("RUN_TIME_END: --- {s}".format(s=expirationDate))
print("Log-USER: --- {s}".format(s=logUser))
print("RUN_TIME_END: --- {s}".format(s=expirationDate))

if os.path.exists(deskPath):
	logging.info("Desktop Path: --- {s}".format(s=deskPath))
	logging.info("_______________________________________________________________________________")
	print("Desktop Path: --- {s}".format(s=deskPath))
	print("______________________________________________________________________________________")
else:
	logging.info("Desktop Path missing: --- {s}".format(s=deskPath))
	logging.info("_______________________________________________________________________________")
	print("Desktop Path missing: --- {s}".format(s=deskPath))
	print("______________________________________________________________________________________")
	sys.exit()

if cDate >= expirationDate:
	messagebox.showinfo("  Attention ", " The Test Time has expired ! ")
	logging.info(" The Test Time has expired ! = {s}: ".format(s=expirationDate))
	logging.info("______________________________________________________________________________")
	print(" The Test Time has expired ! = {s}: ".format(s=expirationDate))
	print("_____________________________________________________________________________________")
	sys.exit()


class AsyncSWX(Thread):
	def __init__(self, str_Message, target):
		super().__init__()
		
		self.drwFile_box = None
		self.drtFile_box = None
		self.str_Message = str_Message
		self.target = target
	
	def run(self):
		
		if self.target == "drw":
			self.drwFile_box = self.str_Message
		elif self.target == "drt":
			self.drtFile_box = self.str_Message


def close():
	sys.exit()


class MyGUI(tk.Tk):
	DRW_File_Path_List = []
	SLSDRT_File_Path_List = []
	FilePath = None
	proclist = []
	src_path = None
	src_path2 = None
	xml_path = None
	folder_path = None
	swxVersion = None
	swxType = None
	drt_path = None
	swApp = None
	
	def resource_path(self):
		base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
		return os.path.join(base_path, self)
	
	src_path = resource_path("SheetConfig\XML")
	if not os.path.exists(xmlFolder):
		shutil.copytree(src_path, xmlFolder)
		logging.debug("Create Folder:  {s}".format(s=xmlFolder))
		print("Create Folder:  {s}".format(s=xmlFolder))
	
	src_path2 = resource_path("SheetConfig\icon")
	if not os.path.exists(assetPath):
		shutil.copytree(src_path2, assetPath)
	if not os.path.isfile(assetPath + "/view-refresh.ico") or not os.path.isfile(assetPath + "/openfolder.png"):
		shutil.rmtree(assetPath)
		shutil.copytree(src_path2, assetPath)
		logging.debug("Create Folder:  {s}".format(s=assetPath))
		print("Create Folder:  {s}".format(s=assetPath))
	
	xml_path = resource_path(xmlFolder + "/configXML.xml")
	
	if not os.path.isfile(xml_path):
		shutil.rmtree(xmlFolder)
		shutil.copytree(src_path, xmlFolder)
	
	x = untangle.parse(xml_path)
	folder_path = x.data.datamodel.path1['filepath']
	swxVersion = x.data.datamodel.version['release']
	drt_path = x.data.datamodel.path2['slddrtpath']
	swxType = x.data.datamodel.version['type']
	logging.debug("configXML type: = {s}".format(s=swxType))
	print("configXML type: = {s}".format(s=swxType))
	
	if folder_path == "":
		logging.debug("configXML filepath: = {s}".format(s=folder_path))
		print("configXML filepath: = {s}".format(s=folder_path))
		messagebox.showinfo("  Attention ", "configXML filepath:" + "\r\n" "filepath --- no value !")
		sys.exit()
	elif swxVersion == "":
		logging.info("configXML release: = {s}".format(s=swxVersion))
		print("configXML release: = {s}".format(s=swxVersion))
		messagebox.showinfo("  Attention ", "configXML release:" + "\r\n" "release --- no value !")
		sys.exit()
	elif drt_path == "":
		logging.info("configXML slddrtpath: = {s}".format(s=drt_path))
		print("configXML slddrtpath: = {s}".format(s=drt_path))
		messagebox.showinfo("  Attention ", "configXML slddrtpath:" + "\r\n" "slddrtpath --- no value !")
		sys.exit()
	
	strTmp = folder_path[1:]
	drive = folder_path.split(strTmp)
	drive = drive[0]
	
	def IsDriveExists(strDrive):
		return os.path.exists(strDrive + ':\\')
	
	if not IsDriveExists(drive):
		logging.info("Folder Path not available = {s}: ".format(s=folder_path))
		messagebox.showinfo("  Attention ",
		                    "configXML Drive:" + "\r\n\n" + drive + "\r\n\n" + " --- Not available ! ---")
		sys.exit()
	if not os.path.exists(folder_path):
		logging.info("Folder Path not available = {s}: ".format(s=folder_path))
		messagebox.showinfo("  Attention ",
		                    "configXML filepath:" + "\r\n\n" + folder_path + "\r\n\n" + " --- Not available ! ---")
		sys.exit()
	
	# API Version
	apiVersion = int(swxVersion) - 1992
	readSLDWORKS = Dispatch('Scripting.FileSystemObject')
	info = readSLDWORKS.GetFileVersion(folder_path)
	split_string = info.split(".", 1)
	substring = split_string[0]
	
	if substring != str(apiVersion):
		apiVersion = substring
		logging.info("SWX Version = {s}: ".format(s=swxVersion))
		logging.info("API Version changed to = {s}: ".format(s=apiVersion))
		print("SWX Version = {s}: ".format(s=swxVersion))
		print("API Version changed to = {s}: ".format(s=apiVersion))
	else:
		print("-------------------------------------------------------")
	
	def __init__(self):
		global FilePath
		super().__init__()
		self.swModel = None
		self.SW_PROCESS_NAME = None
		self.swApp = None
		self.processName = None
		self.folder_path = None
		self.str_Message = None
		self.PaperEsize = None
		self.PaperDsize = None
		self.PaperCsize = None
		self.PaperBsize = None
		self.PaperAsize = None
		self.PaperUserDefined = None
		self.PaperA0size = None
		self.PaperA1size = None
		self.PaperA2size = None
		self.PaperA3size = None
		self.PaperA4sizeVertical = None
		self.PaperA4size = None
		self.drt_path = None
		self.x = None
		self.sheet_config = []
		self.xml_path = None
		self.drt_type = None
		self.dir_Name = None
		self.drw_type = None
		self.file1 = None
		self.swx_thread = None
		# self.FilePath = File_Path
		self.geometry('700x500')
		self.resizable(0, 0)
		self.iconbitmap(assetPath + '/view-refresh.ico')
		self.title('11_Change_SWX_Properties                                                ' + 'Version: ' + version)
		self.attributes('-alpha', 0.9)
		self.config(bg='#b4b4b4')
		self.attributes('-topmost', 1)
		
		self.window_width = 700
		self.window_height = 510
		
		# get the screen dimension
		self.screen_width = self.winfo_screenwidth()
		self.screen_height = self.winfo_screenheight()
		
		# find the center point
		self.center_x = int(self.screen_width / 2 - self.window_width / 2)
		self.center_y = int(self.screen_height / 2 - self.window_height / 2)
		
		# set the position of the window to the center of the screen
		self.geometry(f'{self.window_width}x{self.window_height}+{self.center_x}+{self.center_y}')
		
		# UI options
		border_1 = {"flat": tk.FLAT, "sunken": tk.SUNKEN, "raised": tk.RAISED, "groove": tk.GROOVE,
		            "ridge": tk.RIDGE, }
		
		border_2 = {"sunken": tk.SUNKEN, }
		
		# configure style
		self.style = ttk.Style(self)
		self.style.theme_use('clam')
		self.style.configure('Frame1.TFrame', font=('Veranda', 10), **border_1, background='#b4b4b4')
		self.style.configure('TLabel', font=('Veranda', 10), **border_1, background='#b4b4b4', lightcolor='grey')
		self.style.map('C.TButton', background=[('active', '#FF4C4C')])
		self.style.map('N.TButton', background=[('active', '#9ba9db')])
		self.style.map('R.TButton', background=[('active', 'lime green')])
		self.style.configure("blue.Horizontal.TProgressbar", background='SteelBlue4', darkcolor='lime green',
		                     lightcolor='lime green', )
		
		# heading style
		self.style.configure('Heading.TLabel', font=('Veranda', 10), **border_1)
		
		self.columnconfigure(0, weight=1)
		self.columnconfigure(1, weight=1)
		self.columnconfigure(2, weight=1)
		self.columnconfigure(3, weight=1)
		self.columnconfigure(4, weight=1)
		
		self.header = ttk.Frame(self, style='Frame1.TFrame')
		self.label_drw = ttk.Label(self, text="SW Folder: ", style='Heading.TLabel')
		self.label_drw.grid(column=0, row=0, columnspan=4, sticky=tk.NW, padx=15, pady=10)
		
		self.file_drwInput = tk.Text(self, width=133, height=1.2)
		self.file_drwInput.grid(column=0, row=1, sticky=tk.W, columnspan=4, padx=15, pady=5)
		self.file_drwInput.config(bg='#989898')
		self.FilePath_icon1 = tk.PhotoImage(file=assetPath + '/openfolder.png')
		self.FilePath_button1 = ttk.Button(self, image=self.FilePath_icon1, style='N.TButton',
		                                   command=self.filePathDialog,
		                                   state=tk.NORMAL)
		self.FilePath_button1.grid(column=4, row=1, sticky=tk.W, columnspan=1, padx=5, pady=5)
		
		self.drtFile_box = tk.Text(self, height=20, width=80, wrap='word', state="disabled")
		self.drtFile_box.config(bg='black', foreground='orange')
		self.drtFile_box.grid(column=0, row=2, columnspan=2, rowspan=2, sticky=tk.NE, padx=10, pady=5)
		self.drtFile_box.config(state="normal")
		self.drtFile_box.insert(tk.END, "--Sheet-Formate--" + "\n" + "\n")
		self.drtFile_box.config(state="disabled")
		
		self.drwFile_box = tk.Text(self, height=20, width=80, wrap='word', state="disabled")
		self.drwFile_box.config(bg='black', foreground='orange')
		self.drwFile_box.grid(column=3, row=2, columnspan=2, rowspan=2, sticky=tk.NW, padx=10, pady=5)
		self.drwFile_box.config(state="normal")
		self.drwFile_box.insert(tk.END, "--SW-File--" + "\n" + "\n")
		self.drwFile_box.config(state="disabled")
		
		self.scrollbar = ttk.Scrollbar(self, orient='vertical', command=self.drwFile_box.yview)
		self.scrollbar.grid(column=5, row=2, columnspan=2, rowspan=2, sticky=tk.NS, padx=5, pady=5)
		self.drwFile_box['yscrollcommand'] = self.scrollbar.set
		
		self.footer = ttk.Frame(self, style='Heading.TLabel')
		self.cancelButton = ttk.Button(self, text="Cancel", style='C.TButton', command=self.close)
		self.cancelButton.grid(column=3, row=4, sticky=tk.E, columnspan=1, padx=10, pady=5)
		self.runButton = ttk.Button(self, text="Run", width=15, style='R.TButton', state="disabled",
		                            command=self.background_process)
		self.runButton.grid(column=0, row=4, sticky=tk.E, padx=15, pady=5)
		
		# progressbar
		self.progressbar = ttk.Progressbar(self, style="blue.Horizontal.TProgressbar", orient='horizontal',
		                                   mode='determinate', length=420)
		# place the progressbar
		self.progressbar.grid(column=0, row=5, sticky=tk.W, columnspan=3, padx=10, pady=5)
		
		# label
		self.value_label = ttk.Label(self, text='Waiting for "work"')
		self.value_label.grid(column=3, row=5, sticky=tk.W, columnspan=2, padx=5, pady=5, ipadx=5, ipady=0)
		
		self.queue = mp.Queue()
		self.process = None
		self.run = self.check_IfProcessRunning()
		
		if self.swxType == "UES":
			self.drtFile()
			self.check_IfProcessRunning()
			self.run_SWX()
			self.startSW()
		
		elif self.swxType == "default" and self.run == True:
			self.drtFile()
			print("SWX Process running = {s}: ".format(s=True))
			logging.info("SWX Process running = {s}: ".format(s=True))
		elif self.swxType == "hidden" and self.run == None:
			self.drtFile()
		elif self.swxType == "hidden" and self.run == True:
			self.drtFile()
			logging.info("SWX Process running = {s}: ".format(s="hidden"))
			print("SWX Process running = {s}: ".format(s="hidden"))
	
	def filePathDialog(self):
		global FilePath, drw_type, DRW_File_Path_List, xml_path
		self.DRW_File_Path_List.clear()
		self.drw_type = "*.SLDDRW"
		self.file_drwInput.delete('1.0', tk.END)
		self.dir_Name = fd.askdirectory(initialdir="/", title='Please select a directory')
		self.file_drwInput.insert('1.0', self.dir_Name)
		log = str(self.dir_Name)
		FilePath = log + "/"
		logging.info("Selected Folder Path: = {s}. ".format(s=FilePath))
		base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
		xml_path = os.path.join(base_path, xmlFolder + "/configXML.xml")
		tree = ET.parse(xml_path)
		root = tree.getroot()
		for temp in root.iter('path3'):
			temp.set('temppath', FilePath)
			tree.write(xml_path)
		self.value_label.configure(text='Waiting for "work"')
		self.progressbar.configure(value=0)
		self.drwFile_box.config(state="normal")
		self.drwFile_box.delete('1.0', tk.END)
		self.drwFile_box.insert(tk.END, "--SW-File--" + "\n" + "\n")
		for self.file in glob.glob(FilePath + "*.SLDDRW"):
			if self.drw_type == "*.SLDDRW":
				self.file1 = self.file.replace("\\", "/")
				if not os.path.basename(self.file1)[0] == "~":
					self.str_Message = os.path.basename(self.file1)
					self.DRW_File_Path_List.append(self.file1)
					self.swx_thread = AsyncSWX(self.str_Message, "drw")
					self.swx_thread.start()
					self.monitor(self.swx_thread)
		self.runButton.config(state="normal")
		return FilePath
	
	def drtFile(self):
		global SLSDRT_File_Path_List
		self.drt_type = "*.slddrt"
		self.x = untangle.parse(MyGUI.xml_path)
		self.PaperA4size = self.x.data.datamodel.PaperA4size['filename']
		self.PaperA4sizeVertical = self.x.data.datamodel.PaperA4sizeVertical['filename']
		self.PaperA3size = self.x.data.datamodel.PaperA3size['filename']
		self.PaperA2size = self.x.data.datamodel.PaperA2size['filename']
		self.PaperA1size = self.x.data.datamodel.PaperA1size['filename']
		self.PaperA0size = self.x.data.datamodel.PaperA0size['filename']
		self.PaperUserDefined = self.x.data.datamodel.PaperUserDefined['filename']
		self.PaperAsize = self.x.data.datamodel.PaperAsize['filename']
		self.PaperBsize = self.x.data.datamodel.PaperBsize['filename']
		self.PaperCsize = self.x.data.datamodel.PaperCsize['filename']
		self.PaperDsize = self.x.data.datamodel.PaperDsize['filename']
		self.PaperEsize = self.x.data.datamodel.PaperEsize['filename']
		
		self.sheet_config.append('PaperA4size_H -- ' + self.PaperA4size)
		self.sheet_config.append('PaperA4size_V -- ' + self.PaperA4sizeVertical)
		self.sheet_config.append('PaperA3size -- ' + self.PaperA3size)
		self.sheet_config.append('PaperA2size -- ' + self.PaperA2size)
		self.sheet_config.append('PaperA1size -- ' + self.PaperA1size)
		self.sheet_config.append('PaperA0size -- ' + self.PaperA0size)
		self.sheet_config.append('PaperUserDefined -- ' + self.PaperUserDefined)
		self.sheet_config.append('PaperAsize -- ' + self.PaperAsize)
		self.sheet_config.append('PaperBsize -- ' + self.PaperBsize)
		self.sheet_config.append('PaperCsize -- ' + self.PaperCsize)
		self.sheet_config.append('PaperDsize -- ' + self.PaperDsize)
		self.sheet_config.append('PaperEsize -- ' + self.PaperEsize)
		
		lastChar = MyGUI.drt_path[-1]
		if lastChar != "\\":
			for self.file in glob.glob(MyGUI.drt_path + "\\" + self.drt_type):
				self.file1 = self.file.replace("\\", "/")
				if not os.path.basename(self.file1)[0] == "~":
					self.SLSDRT_File_Path_List.append(self.file1)
		else:
			for self.file in glob.glob(MyGUI.drt_path + self.drt_type):
				self.file1 = self.file.replace("\\", "/")
				if not os.path.basename(self.file1)[0] == "~":
					self.SLSDRT_File_Path_List.append(self.file1)
		
		for self.conf in self.sheet_config:
			self.str_Message = self.conf
			swx_thread = AsyncSWX(self.str_Message, "drt")
			swx_thread.start()
			self.monitor(swx_thread)
	
	def close(self):
		if self.swxType == "hidden":
			self.check_IfProcessRunning()
			self.run_SWX()
			self.destroy()
		elif self.swxType == "default":
			self.destroy()
		elif self.swxType == "UES":
			self.check_IfProcessRunning()
			# self.run_SWX()
			self.destroy()
	
	def run_SWX(self):
		if len(self.proclist) >= 0:
			for id in self.proclist:
				sb.call('Taskkill /pid ' + str(id) + ' /F')
				self.proclist.remove(id)
				logging.info("SolidWorks process was running PID: = {s}. ".format(s=id))
				print("SolidWorks process was running PID: = {s}. ".format(s=id))
		else:
			logging.info("No SolidWorks process was running")
			print("No SolidWorks process was running")
	
	def check_IfProcessRunning(self):
		self.processName = 'SLDWORKS.exe'
		for proc in psutil.process_iter():
			try:
				if self.processName.lower() in proc.name().lower():
					self.proclist.append(proc.pid)
					return True
			except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
				pass
	
	def startSW(self):
		self.SW_PROCESS_NAME = MyGUI.folder_path
		sb.Popen(self.SW_PROCESS_NAME)
	
	def connectToSW(self):
		global swApp
		try:
			self.swApp = win32com.client.dynamic.Dispatch("SldWorks.Application." + str(MyGUI.apiVersion))
			logging.info(" connect To Solidworks API Version: {a}. ".format(a=MyGUI.apiVersion))
			print(" connect To Solidworks API Version: {a}. ".format(a=MyGUI.apiVersion))
			return self.swApp
		except TypeError as er:
			logging.error("Error: %s : %s" % (er, er.strerror))
			print("Error: %s : %s" % (er, er.strerror))
			sys.exit()
	
	def background_process(self):
		global str_Message
		self.process = mp.Process(target=SLD_CAD.work, args=(self.queue,))
		self.process.start()
		self.periodic_call()
	
	def update_drwbox(self):
		global FilePath
		self.drw_type = "*.SLDDRW"
		base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
		xml_path = os.path.join(base_path, xmlFolder + "/configXML.xml")
		tree = ET.parse(xml_path)
		root = tree.getroot()
		for temp in root.iter('path3'):
			temp.set('temppath', FilePath)
			tree.write(xml_path)
		self.drwFile_box.config(state="normal")
		self.drwFile_box.tag_config("done", foreground='lime green')
		self.drwFile_box.insert(tk.END, "" + "\n" + "\n")
		self.drwFile_box.insert(tk.END, "--Done--" + "\n" + "\n", "done")
		
		for self.file in glob.glob(FilePath + "*.SLDDRW"):
			if self.drw_type == "*.SLDDRW":
				self.file1 = self.file.replace("\\", "/")
				if not os.path.basename(self.file1)[0] == "~":
					self.str_Message = os.path.basename(self.file1)
					self.swx_thread = AsyncSWX(self.str_Message, "done")
					self.swx_thread.start()
					self.monitor(self.swx_thread)
	
	def start_work(self):
		self.process = mp.Process(target=SLD_CAD.work, args=(self.queue,))
		self.process.start()
		self.periodic_call()
	
	def periodic_call(self):
		self.check_queue()
		if self.process.exitcode is None:
			self.after(100, self.periodic_call)
		else:
			self.process.join()
			self.runButton.configure(state='normal')
			self.update_drwbox()
	
	def check_queue(self):
		while self.queue.qsize():
			try:
				self.value_label.configure(text=self.queue.get(0))
				i = len(MyGUI.DRW_File_Path_List)
				self.progressbar.configure(value=self.progressbar['value'] + (100 / i))
			except queue.Empty:
				pass
	
	def monitor(self, thread):
		if thread.is_alive():
			self.after(100, lambda: self.monitor(thread))
		
		else:
			if thread.target == "drw":
				self.drwFile_box.config(state="normal")
				self.drwFile_box.insert(tk.END, thread.str_Message + "\n")
				self.drwFile_box.config(state="disabled")
			elif thread.target == "drt":
				self.drtFile_box.config(state="normal")
				# self.drtFile_box.delete('1.0', tk.END)
				self.drtFile_box.insert(tk.END, thread.str_Message + "\n")
				self.drtFile_box.config(state="disabled")
			elif thread.target == "done":
				self.drwFile_box.config(state="normal")
				self.drwFile_box.tag_config("done", foreground='lime green')
				self.drwFile_box.insert(tk.END, thread.str_Message + "\n", "done")
				self.drwFile_box.config(state="disabled")
				self.runButton.config(state="disabled")
			else:
				self.showinfo(message='The progress completed!')
				self.runButton.config(state="disabled")


class SLD_CAD():
	FilePath = None
	folder_path = None
	swApp = None
	swModel = None
	swxVersion = None
	drt_path = None
	swxType = None
	DRW_FilePath_List = []
	SLDDRT_FilePath_List = []
	PaperA4size = ""
	PaperA4sizeVertical = ""
	PaperA3size = ""
	PaperA2size = ""
	PaperA1size = ""
	PaperA0size = ""
	PaperUserDefined = ""
	PaperAsize = ""
	PaperBsize = ""
	PaperCsize = ""
	PaperDsize = ""
	PaperEsize = ""
	
	def __init__(self):
		""" Initalisieren über Eltern-Klasse """
		super().__init__()
	
	def resource_path(self):
		base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
		return os.path.join(base_path, self)
	
	src_path = resource_path("SheetConfig\XML")
	
	if not os.path.exists(xmlFolder):
		shutil.copytree(src_path, xmlFolder)
		logging.info("Create Folder:  {s}".format(s=xmlFolder))
		print("Create Folder:  {s}".format(s=xmlFolder))
	
	src_path2 = resource_path("SheetConfig\icon")
	if not os.path.exists(assetPath):
		shutil.copytree(src_path2, assetPath)
	if not os.path.isfile(assetPath + "/view-refresh.ico") or not os.path.isfile(assetPath + "/openfolder.png"):
		shutil.rmtree(assetPath)
		shutil.copytree(src_path2, assetPath)
		logging.info("Create Folder:  {s}".format(s=assetPath))
		print("Create Folder:  {s}".format(s=assetPath))
	
	xml_path = resource_path(xmlFolder + "/configXML.xml")
	
	if not os.path.isfile(xml_path):
		shutil.rmtree(xmlFolder)
		shutil.copytree(src_path, xmlFolder)
	
	x = untangle.parse(xml_path)
	# logging.info(x.data.datamodel.path1['filepath'])
	print(x.data.datamodel.path1['filepath'])
	folder_path = x.data.datamodel.path1['filepath']
	swxVersion = x.data.datamodel.version['release']
	drt_path = x.data.datamodel.path2['slddrtpath']
	FilePath = x.data.datamodel.path3['temppath']
	swxType = x.data.datamodel.version['type']
	
	if folder_path == "":
		logging.info("configXML filepath: = {s}".format(s=folder_path))
		print("configXML filepath: = {s}".format(s=folder_path))
		messagebox.showinfo("  Attention ", xml_path + "\r\n" "filepath --- no value !")
		sys.exit()
	elif swxVersion == "":
		logging.info("configXML release: = {s}".format(s=swxVersion))
		print("configXML release: = {s}".format(s=swxVersion))
		messagebox.showinfo("  Attention ", xml_path + "\r\n" "release --- no value !")
		sys.exit()
	elif drt_path == "":
		logging.info("configXML slddrtpath: = {s}".format(s=drt_path))
		print("configXML slddrtpath: = {s}".format(s=drt_path))
		messagebox.showinfo("  Attention ", xml_path + "\r\n" "slddrtpath --- no value !")
		sys.exit()
	
	strTmp = folder_path[1:]
	drive = folder_path.split(strTmp)
	drive = drive[0]
	
	def IsDriveExists(strDrive):
		return os.path.exists(strDrive + ':\\')
	
	if not IsDriveExists(drive):
		logging.info("Folder Path not available = {s}: ".format(s=folder_path))
		print("Folder Path not available = {s}: ".format(s=folder_path))
		sys.exit()
	if not os.path.exists(folder_path):
		logging.info("Folder Path not available = {s}: ".format(s=folder_path))
		print("Folder Path not available = {s}: ".format(s=folder_path))
		sys.exit()
	
	# API Version
	apiVersion = int(swxVersion) - 1992
	readSLDWORKS = Dispatch('Scripting.FileSystemObject')
	info = readSLDWORKS.GetFileVersion(folder_path)
	split_string = info.split(".", 1)
	substring = split_string[0]
	#
	if substring != str(apiVersion):
		apiVersion = substring
		logging.info("SWX Version = {s}: ".format(s=swxVersion))
		logging.info("API Version changed to = {s}: ".format(s=apiVersion))
	else:
		logging.info("SWX Version = {s}: ".format(s=swxVersion))
		logging.info("API Version = {s}: ".format(s=apiVersion))
	
	def Folder_Data_List(path, strtype):
		global str_Message
		for file in glob.glob(path + strtype):
			if strtype == "*.SLDDRW":
				file1 = file.replace("\\", "/")
				if not os.path.basename(file1)[0] == "~":
					SLD_CAD.DRW_FilePath_List.append(file1)
			if strtype == "*.SLDDRT":
				file2 = file.replace("\\", "/")
				if not os.path.basename(file1)[0] == "~":
					SLD_CAD.SLSDRT_FilePath_List.append(file2)
			else:
				print(strtype)
	
	@staticmethod
	def connectToSW():
		global swApp
		try:
			swApp = win32com.client.dynamic.Dispatch("SldWorks.Application." + str(SLD_CAD.apiVersion))
			logging.info(" connect To Solidworks API Version: {a}. ".format(a=SLD_CAD.apiVersion))
			print(" connect To Solidworks API Version: {a}. ".format(a=SLD_CAD.apiVersion))
			return swApp
		except TypeError as er:
			logging.error("Error: %s : %s" % (er, er.strerror))
			print("Error: %s : %s" % (er, er.strerror))
			sys.exit()
	
	def openFile(self, sPath):
		global swModel, str_Message
		f = self.getopendocspec(sPath)
		swModel = self.opendoc7(f)
		logging.info("openFile = {s}. ".format(s=sPath))
		print("openFile = {s}. ".format(s=sPath))
		return swModel
	
	@staticmethod
	def GetSheetSizeFromPaperSize(SheetWidth, SheetHeight):
		swDwgPaperA4size = 0
		swDwgPaperA4sizeVertical = 1
		swDwgPaperA3size = 2
		swDwgPaperA2size = 3
		swDwgPaperA1size = 4
		swDwgPaperA0size = 5
		swDwgPapersUserDefined = 6
		swDwgPaperAsize = 7
		swDwgPaperAsizeVertical = 8
		swDwgPaperBsize = 9
		swDwgPaperCsize = 10
		swDwgPaperDsize = 11
		swDwgPaperEsize = 12
		
		if round(SheetWidth, 4) == 0.2794 and round(SheetHeight, 4) == 0.2159:
			GetSheetSizeFromPaperSize = swDwgPaperAsize
		elif round(SheetWidth, 4) == 0.2159 and round(SheetHeight, 4) == 0.2794:
			GetSheetSizeFromPaperSize = swDwgPaperAsizeVertical
		elif round(SheetWidth, 4) == 0.4318 and round(SheetHeight, 4) == 0.2794:
			GetSheetSizeFromPaperSize = swDwgPaperBsize
		elif round(SheetWidth, 4) == 0.5588 and round(SheetHeight, 4) == 0.4318:
			GetSheetSizeFromPaperSize = swDwgPaperCsize
		elif round(SheetWidth, 4) == 0.8636 and round(SheetHeight, 4) == 0.5588:
			GetSheetSizeFromPaperSize = swDwgPaperDsize
		elif round(SheetWidth, 4) == 1.1176 and round(SheetHeight, 4) == 0.8636:
			GetSheetSizeFromPaperSize = swDwgPaperEsize
		elif round(SheetWidth, 4) == 0.297 and round(SheetHeight, 4) == 0.21:
			GetSheetSizeFromPaperSize = swDwgPaperA4size
		elif round(SheetWidth, 4) == 0.21 and round(SheetHeight, 4) == 0.297:
			GetSheetSizeFromPaperSize = swDwgPaperA4sizeVertical
		elif round(SheetWidth, 4) == 0.42 and round(SheetHeight, 4) == 0.297:
			GetSheetSizeFromPaperSize = swDwgPaperA3size
		elif round(SheetWidth, 4) == 0.594 and round(SheetHeight, 4) == 0.42:
			GetSheetSizeFromPaperSize = swDwgPaperA2size
		elif round(SheetWidth, 4) == 0.841 and round(SheetHeight, 4) == 0.594:
			GetSheetSizeFromPaperSize = swDwgPaperA1size
		elif round(SheetWidth, 4) == 1.189 and round(SheetHeight, 4) == 0.841:
			GetSheetSizeFromPaperSize = swDwgPaperA0size
		else:
			GetSheetSizeFromPaperSize = swDwgPapersUserDefined
		return GetSheetSizeFromPaperSize
	
	@staticmethod
	def open_DRW(swModel):
		global templateName
		drt_type = "*.slddrt"
		SLDDRT_FilePath_List = []
		allowedSheetFormat = True
		swDocDRAWING = 3
		swDwgTemplateCustom = 12
		swDwgTemplateNone = 13
		
		lastChar = SLD_CAD.drt_path[-1]
		if lastChar != "\\":
			SLD_CAD.drt_path = SLD_CAD.drt_path + "\\"
		
		sheetformatpath = []
		sheetformatpath.append(SLD_CAD.drt_path + PaperA4size)
		sheetformatpath.append(SLD_CAD.drt_path + PaperA4sizeVertical)
		sheetformatpath.append(SLD_CAD.drt_path + PaperA3size)
		sheetformatpath.append(SLD_CAD.drt_path + PaperA2size)
		sheetformatpath.append(SLD_CAD.drt_path + PaperA1size)
		sheetformatpath.append(SLD_CAD.drt_path + PaperA0size)
		sheetformatpath.append(SLD_CAD.drt_path + PaperUserDefined)
		sheetformatpath.append(SLD_CAD.drt_path + PaperAsize)
		sheetformatpath.append(SLD_CAD.drt_path + PaperBsize)
		sheetformatpath.append(SLD_CAD.drt_path + PaperCsize)
		sheetformatpath.append(SLD_CAD.drt_path + PaperDsize)
		sheetformatpath.append(SLD_CAD.drt_path + PaperEsize)
		
		counter = 0
		for counter in range(0, len(sheetformatpath), 1):
			file = sheetformatpath[counter]
			file1 = file.replace("\\", "/")
			SLDDRT_FilePath_List.append(file1)
		
		AnzahlBl = swModel.GetSheetCount
		SheetNames = swModel.GetSheetNames
		
		i = 0
		for i in range(0, AnzahlBl, 1):
			if swModel.ActivateSheet(SheetNames[i]) == True:
				Sheet = swModel.GetCurrentSheet
				SheetProperties = Sheet.GetProperties
				Name = Sheet.GetName
				paperSize = SheetProperties[0]
				templateIn = swDwgTemplateNone
				scale1 = SheetProperties[2]
				scale2 = SheetProperties[3]
				firstAngle = SheetProperties[4]
				templateName = ""
				Width = SheetProperties[5]
				Height = SheetProperties[6]
				propertyViewName = Sheet.CustomPropertyView
				
				retval = swModel.SetupSheet4(Name, paperSize, templateIn, scale1, scale2, firstAngle, templateName,
				                             Width, Height, propertyViewName)
				if retval == False:
					print('retval = False')
				else:
					templateIn = swDwgTemplateCustom
					paperSize = SLD_CAD.GetSheetSizeFromPaperSize(Width, Height)
					templateName = SLDDRT_FilePath_List[paperSize]
					retval = swModel.SetupSheet4(Name, paperSize, templateIn, scale1, scale2, firstAngle, templateName,
					                             Width, Height, propertyViewName)
					if retval == False:
						allowedSheetFormat = False
					else:
						Sheet.SheetFormatVisible = True
			else:
				print(templateName)
	
	def work(working_queue):
		global folder_path, FilePath, str_FN, listcount, swxVersion, drt_path, PaperA4size, PaperA4sizeVertical, PaperA3size, PaperA2size, PaperA1size, PaperA0size, PaperUserDefined, PaperAsize, PaperBsize, PaperCsize, PaperDsize, PaperEsize
		
		xml_path = SLD_CAD.resource_path(xmlFolder + "/configXML.xml")
		x = untangle.parse(xml_path)
		# logging.info(x.data.datamodel.path1['filepath'])
		logging.info("--------------------------------------------")
		folder_path = x.data.datamodel.path1['filepath']
		swxVersion = x.data.datamodel.version['release']
		drt_path = x.data.datamodel.path2['slddrtpath']
		FilePath = x.data.datamodel.path3['temppath']
		logging.info("--------------------------------------------")
		logging.info("Selected Folder: = {s}  ".format(s=FilePath))
		PaperA4size = x.data.datamodel.PaperA4size['filename']
		logging.info("PaperA4size: = {s}  ".format(s=PaperA4size))
		PaperA4sizeVertical = x.data.datamodel.PaperA4sizeVertical['filename']
		logging.info("PaperA4sizeVertical: = {s}  ".format(s=PaperA4sizeVertical))
		PaperA3size = x.data.datamodel.PaperA3size['filename']
		logging.info("PaperA3size: = {s}  ".format(s=PaperA3size))
		PaperA2size = x.data.datamodel.PaperA2size['filename']
		logging.info("PaperA2size: = {s}  ".format(s=PaperA2size))
		PaperA1size = x.data.datamodel.PaperA1size['filename']
		logging.info("PaperA1size: = {s}  ".format(s=PaperA1size))
		PaperA0size = x.data.datamodel.PaperA0size['filename']
		logging.info("PaperA0size: = {s}  ".format(s=PaperA0size))
		PaperUserDefined = x.data.datamodel.PaperUserDefined['filename']
		logging.info("PaperUserDefined: = {s}  ".format(s=PaperUserDefined))
		PaperAsize = x.data.datamodel.PaperAsize['filename']
		logging.info("PaperAsize: = {s}  ".format(s=PaperAsize))
		PaperBsize = x.data.datamodel.PaperBsize['filename']
		logging.info("PaperBsize: = {s}  ".format(s=PaperBsize))
		PaperCsize = x.data.datamodel.PaperCsize['filename']
		logging.info("PaperCsize: = {s}  ".format(s=PaperCsize))
		PaperDsize = x.data.datamodel.PaperDsize['filename']
		logging.info("PaperDsize: = {s}  ".format(s=PaperDsize))
		PaperEsize = x.data.datamodel.PaperEsize['filename']
		logging.info("PaperEsize: = {s}  ".format(s=PaperEsize))
		logging.info("--------------------------------------------")
		SLD_CAD.Folder_Data_List(FilePath, "*.SLDDRT")
		SLD_CAD.Folder_Data_List(FilePath, "*.SLDDRW")
		
		SLD_CAD.swApp = SLD_CAD.connectToSW()
		i = 1
		for file in SLD_CAD.DRW_FilePath_List:
			SLD_CAD.swModel = SLD_CAD.openFile(SLD_CAD.swApp, file)
			if SLD_CAD.swModel != None:
				SLD_CAD.open_DRW(SLD_CAD.swModel)
				SLD_CAD.swModel.Save2(True)
				str_FN = os.path.basename(file)
				logging.info("Status: = {s} Done ".format(s="Save"))
				print("Status: = {s} Done ".format(s="Save"))
				working_queue.put(str(i) + " / " + str(len(SLD_CAD.DRW_FilePath_List)) + "  " + str_FN)
				SLD_CAD.swApp.CloseAllDocuments(True)
				logging.info("Status: = {s} Done ".format(s="Close"))
				print("Status: = {s} Done ".format(s="Close"))
				time.sleep(1.5)
				i = i + 1


if __name__ == '__main__':
	app = MyGUI()
	app.mainloop()
