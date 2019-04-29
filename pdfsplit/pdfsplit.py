import PyPDF2
import os.path
import wx
from wx import adv
import re

#~ Find the current directory of the python file
base_dir = os.path.dirname(os.path.abspath(__file__))

#~ Build the frame class that is called in the main program loop
class MasterFrame(wx.Frame):
	#~ Define actions to be taken on initially loading the MasterFrame class
	def __init__(self):
		#~ create the initial frame and give it a title
		wx.Frame.__init__(self,None,wx.ID_ANY, title='PDF Splitter')
		
		#~ place a panel that holds all other items (by default takes up the whole frame)
		#~ panels change their appearance based on the operating system being used (makes things look cleaner)
		self.panel = wx.Panel(self,wx.ID_ANY)
		
		#~ create a dropdown to hold possible preset commands (this program only has quit)
		filemenu = wx.Menu()
		
		#~ add a quit button to the filemenu
		fquit = filemenu.Append(wx.ID_EXIT,"Exit","Terminate the Program")
		
		#~ create a menu bar to hold individual menu items
		menubar = wx.MenuBar()
		#~ add filemenu to menubar
		menubar.Append(filemenu,"File")
		
		#~ add menubar to panel
		self.SetMenuBar(menubar)
		
		#~ bind a function to the filemenu quit button
		self.Bind(wx.EVT_MENU,self.onQuit,fquit)
		
		#~ create a sizer that will hold all top level items
		self.mainbagsizer = wx.GridBagSizer(5,5)
		
		#~ create a second sizer that will hold individual widgets
		self.controls_gb = wx.GridBagSizer(5,5)
		#~ create a staticbox item (puts a thin border around all it contains)
		self.controls_box = wx.StaticBox(self.panel,wx.ID_ANY)
		#~ set a static box sizer to the static box (allows you to add items to the static box)
		self.controls_sz = wx.StaticBoxSizer(self.controls_box,wx.VERTICAL)
		
		#~ create a file picker which has default actions to select files from directories
		#~ automatically opens file dialog, holds filename and path, etc
		self.filepick = wx.FilePickerCtrl(self.panel,id=wx.ID_ANY,message="Select a PDF",wildcard="*.pdf",style=wx.FLP_OPEN|wx.FLP_FILE_MUST_EXIST,name="mainFile")
		#~ set default directory of the file picker
		self.filepick.SetInitialDirectory(base_dir)
		#~ add file picker to the base-level sizer
		self.controls_gb.Add(self.filepick,pos=(0,0),span=(1,5),flag=wx.EXPAND|wx.ALL,border=8)
		
		#~ add a button that will call a specific action when pressed
		self.printbtn = wx.Button(self.panel,label="Get File",id=wx.ID_ANY)
		#~ add button to base-level sizer
		self.controls_gb.Add(self.printbtn,pos=(1,0),span=(1,1),flag=wx.ALL,border=8)
		
		#~ create array that holds text strings for our radio buttons
		actions = ["Extract selected pages to single file","Split all pages into separate files"]
		#~ create radio buttons (radio box automatically places them in a box) using our "choices" array
		self.radioActions = wx.RadioBox(self.panel,id=wx.ID_ANY,label="Actions",choices=actions,style=wx.RA_SPECIFY_ROWS,name="radioActions")
		#~ add radio buttons box to base-level sizer
		self.controls_gb.Add(self.radioActions,pos=(2,0),span=(1,2),flag=wx.EXPAND|wx.ALL,border=8)
		
		#~ add multi-line text box to list pdf pages for extraction
		self.filePages = wx.TextCtrl(self.panel,id=wx.ID_ANY,style=wx.TE_MULTILINE,name="Extract Page Selection")
		#~ add text box to base-level sizer
		self.controls_gb.Add(self.filePages,pos=(2,2),span=(1,3),flag=wx.EXPAND|wx.ALL,border=8)
		
		#~ add base-level sizer to static box
		self.controls_sz.Add(self.controls_gb)
		#~ add static box to top level sizer
		self.mainbagsizer.Add(self.controls_sz,pos=(0,0),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		#~ set the panel to use the top level sizer as its base
		self.panel.SetSizer(self.mainbagsizer)
		#~ stretch/shrink the sizer to fit the widgets
		self.mainbagsizer.Fit(self)
		#~ stretch/shrink the window to fit the sizers
		self.panel.Layout()
		
		#~ bind clicking the "printbtn" button to execute the fileClick function
		self.Bind(wx.EVT_BUTTON,self.fileClick,self.printbtn)
		#~ bind changing the radio button to the radioSelect function
		self.Bind(wx.EVT_RADIOBOX,self.radioSelect,self.radioActions)
		
		
	#~ function that is called when clicking quit in the file menu
	def onQuit(self,event):
		#~ closes the program by exiting the main loop function
		self.Destroy()
	
	#~ function that fires whenever the radio button selection changes
	def radioSelect(self,event):
		#~ get the radio button currently selected (displayed as an integer depending on the order they were originally listed)
		rSelect = self.radioActions.GetSelection()
		#~ if the first radio button is selected (which is default) do this action
		if rSelect == 0:
			#~ show the page selection text box
			self.controls_gb.Show(self.filePages,True)
		#~ when any other radio button is selected, do this action
		else:
			#~ hide the page selection text box
			self.controls_gb.Hide(self.filePages,True)
		
	#~ function that fires whenever Get File button is clicked
	def fileClick(self,event):
		#~ get radio button selection
		returnAction = self.radioActions.GetSelection()
		
		#~ get file selection
		returnFile = self.filepick.GetPath()
		
		#~ only do these actions if a file is selected
		if returnFile != '':
			
			#~ if the first radio button is selected, do this path
			if returnAction == 0:
				#~ get the number of lines entered in the text box
				lineCt = self.filePages.GetNumberOfLines()
				#~ create a blank variable to hold string as we build it
				textRtn = ""
				#~ create an array to hold page numbers when they have been split out
				iterReturn = []
				#~ loop through the lines in the text box
				for i in range(lineCt):
					#~ add each line of the text box to the blank string (if a string continues to a second line, it is added without any spaces to the first line
					textRtn += self.filePages.GetLineText(i)
				#~ take the complete string from the text box and split it into an array, using commas as the split point
				pages = textRtn.split(",")
				#~ loop through the items in the pages array
				for i in range(len(pages)):
					#~ if the current item in the pages array contains a hyphen, do these actions
					if "-" in pages[i]:
						#~ find the text before the hyphen and convert it to an integer
						rangeBegin = int(pages[i][:pages[i].find("-")])
						#~ find the text after the hyphen and convert it to an integer
						rangeEnd = int(pages[i][pages[i].find("-")+1:])
						#~ find the different between the number before and after the hyphen
						rangeRange = rangeEnd-rangeBegin
						#~ if the second number is less than or equal to the first, throw an error
						if (rangeRange)<=0:
							#~ create error dialog
							dlg = wx.MessageDialog(self.panel,"The following page range is not valid. \n"+pages[i],"Error",wx.OK|wx.ICON_ERROR)
							#~ show error dialog
							dlg.ShowModal()
							#~ destroy error dialog
							dlg.Destroy()
						#~ if the second number is greater than the first, loop through the numbers including and between them and add them to the final array
						else:
							#~ since we need to include the outer edges of the range, perform the difference+1 number of loops (for example 1-4 would need 4 total loops to get 1,2,3, and 4)
							for j in range(rangeRange+1):
								#~ add the current value of rangeBegin to the final array
								iterReturn.append(rangeBegin)
								#~ increase the rangeBegin by 1 for the next loop
								rangeBegin += 1
					#~ if the current pages item does not contain a hyphen, do this
					else:
						# add the current pages item to the final array
						iterReturn.append(int(pages[i]))
				#~ after building the final array, create a save dialog box to decide on a file path and name
				savedlg = wx.FileDialog(self,message="Create Merged File",defaultDir=base_dir,defaultFile=returnFile[-(returnFile[::-1].find("/")):],wildcard="*.pdf",style=wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT)
				#~ show save file dialog
				rec = savedlg.ShowModal()
				#~ if the save file dialog returns a confirmation, do this
				if rec == wx.ID_OK:
					#~ programmatically open the selected file to execute the next items
					with open(returnFile,'rb') as inputfile:
						#~ create an instance of the PyPDF2 pdf reader
						reader = PyPDF2.PdfFileReader(inputfile)
						#~ create a instance of the PyPDF2 pdf writer
						writer = PyPDF2.PdfFileWriter()
						#~ for each value in the final array, perform a loop
						for i in range(len(iterReturn)):
							#~ add the current page item to the writer (the minus 1 is because the pages are addressed starting with 
							#~ 0, just like and array, this will make it easier for readability when entering 
							#~ page numbers in the text box)
							writer.addPage(reader.getPage(iterReturn[i]-1))
							#~ write the selected file name and path to a variable
							outputfile = savedlg.GetPath()
							#~ programmatically write the completed file out
							with open(outputfile,'wb') as out_pdf: writer.write(out_pdf)
					#~ create a message box to display on success
					dlg = wx.MessageDialog(self.panel,"Files Successfully Split.","Success",wx.OK)
					#~ show message box
					dlg.ShowModal()
					#~ destroy message box
					dlg.Destroy()
					
						
			#~ if the second radio button is selected, do these actions
			elif returnAction == 1:
				#~ programmatically open the selected file
				with open(returnFile,'rb') as inputfile:
					#~ create an instance of the PyPDF2 pdf reader
					reader = PyPDF2.PdfFileReader(inputfile)
					#~ find the number of pages in the selected file and save to variable
					pagenum = reader.getNumPages()
					#~ loop through all pages in the selected file
					for x in range(pagenum):
						#~ create an instance of the PyPDF2 pdf writer
						writer = PyPDF2.PdfFileWriter()
						#~ add the current page to the current writer
						writer.addPage(reader.getPage(x))
						#~ create an output file name by adding a page number to the selected file's name
						outputfile = returnFile[0:returnFile.find(".")]+"_"+str(x+1)+".pdf"
						#~ programmatically write the file out
						with open(outputfile,'wb') as out_pdf: writer.write(out_pdf)
				#~ create message box to display on success
				dlg = wx.MessageDialog(self.panel,"Files Successfully Split.","Success",wx.OK)
				#~ show message box
				dlg.ShowModal()
				#~ destroy message box
				dlg.Destroy()

				
			else:
				pass
		#~ if no file is selected, do these actions
		else:
			#~ create error dialog box
			dlg = wx.MessageDialog(self.panel,"Please Select a File.","Error",wx.OK)
			#~ show error dialog box
			dlg.ShowModal()
			#~ destry error dialog box
			dlg.Destroy()

#~ create main program loop
if __name__ == '__main__':
	app = wx.App(False)
	#~ display MasterFrame item
	frame = MasterFrame().Show()
	app.MainLoop()
