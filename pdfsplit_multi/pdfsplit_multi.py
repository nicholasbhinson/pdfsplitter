import PyPDF2
import os.path
import wx
from wx import adv
import re


base_dir = os.path.dirname(os.path.abspath(__file__))

#~ with open(os.path.join(base_dir,"Test.pdf"),'rb') as inputfile:
		#~ reader = PyPDF2.PdfFileReader(inputfile)
		#~ pagenum = reader.getNumPages()
		#~ print(str(pagenum))
		#~ for x in range(pagenum):
			#~ writer = PyPDF2.PdfFileWriter()
			#~ writer.addPage(reader.getPage(x))
			#~ outputfile = os.path.join(base_dir,"Test_"+str(x+1)+".pdf")
			#~ with open(outputfile,'wb') as out_pdf: writer.write(out_pdf)



class MasterFrame(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self,None,wx.ID_ANY, title='PDF Splitter')
		
		self.panel = wx.Panel(self,wx.ID_ANY)
		
		filemenu = wx.Menu()
		
		fquit = filemenu.Append(wx.ID_EXIT,"Exit","Terminate the Program")
		
		menubar = wx.MenuBar()
		menubar.Append(filemenu,"File")
		
		self.SetMenuBar(menubar)
		
		self.Bind(wx.EVT_MENU,self.onQuit,fquit)
		
		self.mainbagsizer = wx.GridBagSizer(5,5)
		
		self.controls_gb = wx.GridBagSizer(5,5)
		self.controls_box = wx.StaticBox(self.panel,wx.ID_ANY,)
		self.controls_sz = wx.StaticBoxSizer(self.controls_box,wx.VERTICAL)
		
		
		self.file_label = wx.StaticText(self.panel,id=wx.ID_ANY,label="Pick File:")
		self.controls_gb.Add(self.file_label,pos=(0,0),span=(1,1),flag=wx.ALL|wx.ALIGN_CENTER_VERTICAL,border=8)
		
		#~ self.filepick = wx.FilePickerCtrl(self.panel,id=wx.ID_ANY,message="Select a PDF",wildcard="*.pdf",style=wx.FLP_OPEN|wx.FLP_FILE_MUST_EXIST,name="mainFile")
		#~ self.filepick.SetInitialDirectory(base_dir)
		#~ self.controls_gb.Add(self.filepick,pos=(0,1),span=(1,4),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.filepath = wx.ListBox(self.panel,id=wx.ID_ANY,choices=[],style=wx.LB_MULTIPLE|wx.LB_EXTENDED|wx.LB_NEEDED_SB,name="selectedFiles")
		self.controls_gb.Add(self.filepath,pos=(0,1),span=(4,4),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.filebrowse = wx.Button(self.panel,label="Add Files",id=wx.ID_ANY)
		self.controls_gb.Add(self.filebrowse,pos=(0,5),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.fileremove = wx.Button(self.panel,label="Remove Selected",id=wx.ID_ANY)
		self.controls_gb.Add(self.fileremove,pos=(1,5),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.fileclear = wx.Button(self.panel,label="Clear Files",id=wx.ID_ANY)
		self.controls_gb.Add(self.fileclear,pos=(2,5),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.selectedFiles = []
		
		self.file_label = wx.StaticText(self.panel,id=wx.ID_ANY,label="               ")
		self.controls_gb.Add(self.file_label,pos=(4,1),span=(1,1),flag=wx.ALL,border=8)
		
		self.printbtn = wx.Button(self.panel,label="Get File",id=wx.ID_ANY)
		self.controls_gb.Add(self.printbtn,pos=(4,2),span=(1,1),flag=wx.ALL,border=8)
		
		
		actions = ["Extract selected pages to single file","Split all pages into separate files","Merge Files"]
		self.radioActions = wx.RadioBox(self.panel,id=wx.ID_ANY,label="Actions",choices=actions,style=wx.RA_SPECIFY_ROWS,name="radioActions")
		self.controls_gb.Add(self.radioActions,pos=(5,0),span=(1,3),flag=wx.EXPAND|wx.ALL,border=8)
		#~ self.radioActions.SetSelection(2)
		
		self.filePages = wx.TextCtrl(self.panel,id=wx.ID_ANY,style=wx.TE_MULTILINE,name="Extract Page Selection")
		self.controls_gb.Add(self.filePages,pos=(5,3),span=(1,3),flag=wx.EXPAND|wx.ALL,border=8)
		
		
		self.controls_sz.Add(self.controls_gb)
		self.mainbagsizer.Add(self.controls_sz,pos=(0,0),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.panel.SetSizer(self.mainbagsizer)
		self.mainbagsizer.Fit(self)
		self.panel.Layout()
		
		self.Bind(wx.EVT_BUTTON,self.fileClick,self.printbtn)
		self.Bind(wx.EVT_BUTTON,self.browseClick,self.filebrowse)
		self.Bind(wx.EVT_BUTTON,self.removeClick,self.fileremove)
		self.Bind(wx.EVT_BUTTON,self.clearClick,self.fileclear)
		
		self.Bind(wx.EVT_RADIOBOX,self.radioSelect,self.radioActions)
		
		
		
	def onQuit(self,event):
		self.Destroy()
		
	def browseClick(self,event):
		dlg = wx.FileDialog(self.panel,defaultDir=base_dir,defaultFile="",wildcard="*.pdf",style=wx.FD_OPEN|wx.FD_MULTIPLE|wx.FD_CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			paths = dlg.GetPaths()
			#~ self.filepath.Clear()
			for path in paths:
				self.filepath.Append(path)
			self.selectedFiles = paths
		dlg.Destroy()
		
	def removeClick(self,event):
		files = self.filepath.GetSelections()
		print(files)
		for i in reversed(range(len(files))):
			print(i)
			self.filepath.Delete(files[i])
		
	def clearClick(self,event):
		self.filepath.Clear()
		
	def radioSelect(self,event):
		rSelect = self.radioActions.GetSelection()
		if rSelect == 0:
			self.controls_gb.Show(self.filePages,True)
		else:
			self.controls_gb.Hide(self.filePages,True)
		
	def fileClick(self,event):
		returnAction = self.radioActions.GetSelection()
		
		returnFile = self.selectedFiles
		
		if returnAction == 0:
			lineCt = self.filePages.GetNumberOfLines()
			textRtn = ""
			iterReturn = []
			for i in range(lineCt):
				textRtn += self.filePages.GetLineText(i)
			files = re.split(r',\s*(?![^()]*\))',textRtn)
			for i in range(len(files)):
				filepos = files[i]
				filepos1 = filepos[:filepos.find("(")]
				#~ print(filepos1)
				filepos2 = filepos[filepos.find("("):].replace("(","").replace(")","")
				filepos2a = filepos2.split(",")
				iterReturn.append([filepos1,filepos2a])
			
			#~ print(iterReturn)
			#~ print(returnFile)
			#~ print("++++++++++")
			writer = PyPDF2.PdfFileWriter()
			outputfile = os.path.join(base_dir,"Test_Extract.pdf")
			for i in range(len(iterReturn)):
				#~ try:
				with open(returnFile[(int(iterReturn[i-1][0])-1)],'rb') as inputfile:
					print(returnFile[(int(iterReturn[i-1][0])-1)])
					#~ print(iterReturn[i-1][1])
					#~ print(len(iterReturn[i-1][1]))
					#~ print("==========")
					reader = PyPDF2.PdfFileReader(inputfile)
					for j in range(len(iterReturn[i-1][1])):
						#~ print("----------")
						print(int(iterReturn[i-1][1][j-1]))
						writer.addPage(reader.getPage(int(iterReturn[i-1][1][j-1])-1))
						with open(outputfile,'wb') as out_pdf:writer.write(out_pdf)
				#~ except (Exception,ArithmeticError) as e:
					#~ errorTemplate = "An invalid file number was selected.\nThis most likely occured because you tried to select a file number higher than the number of files selected."
					#~ message = errorTemplate.format(type(e).__name__,e.args)
					#~ dlg = wx.MessageDialog(self.panel,message,"ERROR",wx.OK|wx.ICON_ERROR)
					#~ dlg.ShowModal()
					#~ dlg.Destroy()
					
			
		elif returnAction == 1:
			for i in range(len(returnFile)):
				with open(returnFile[i],'rb') as inputfile:
					reader = PyPDF2.PdfFileReader(inputfile)
					pagenum = reader.getNumPages()
					for x in range(pagenum):
						writer = PyPDF2.PdfFileWriter()
						writer.addPage(reader.getPage(x))
						outputfile = returnFile[i][0:returnFile[i].find(".")]+"_"+str(x+1)+".pdf"
						with open(outputfile,'wb') as out_pdf: writer.write(out_pdf)
			dlg = wx.MessageDialog(self.panel,"Files Successfully Split.","Success",wx.OK)
			dlg.ShowModal()
			dlg.Destroy()
			
			
		elif returnAction == 2:
			writer = PyPDF2.PdfFileWriter()
			outputfile = os.path.join(base_dir,"Test_merge.pdf")
			for i in range(len(returnFile)):
				with open(returnFile[i],'rb') as inputfile:
					reader = PyPDF2.PdfFileReader(inputfile)
					pagenum = reader.getNumPages()
					for x in range(pagenum):
						writer.addPage(reader.getPage(x))
					with open(outputfile,'ab') as out_pdf: writer.write(out_pdf)
			dlg = wx.MessageDialog(self.panel,"Files Successfully Merged.","Success",wx.OK)
			dlg.ShowModal()
			dlg.Destroy()
			
		else:
			pass


if __name__ == '__main__':
	app = wx.App(False)
	frame = MasterFrame().Show()
	app.MainLoop()
