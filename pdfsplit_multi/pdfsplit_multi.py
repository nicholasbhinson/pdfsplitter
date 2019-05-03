import PyPDF2
import os.path
import wx
from wx import adv
import re


base_dir = os.path.dirname(os.path.abspath(__file__))

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
		
		self.filepath = wx.ListBox(self.panel,id=wx.ID_ANY,choices=[],style=wx.LB_EXTENDED|wx.LB_NEEDED_SB|wx.LB_HSCROLL,name="selectedFiles")
		self.controls_gb.Add(self.filepath,pos=(0,1),span=(4,7),flag=wx.EXPAND|wx.ALL,border=8)
		
		
		self.filebrowse = wx.Button(self.panel,label="Add Files",id=wx.ID_ANY)
		self.controls_gb.Add(self.filebrowse,pos=(0,8),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.fileremove = wx.Button(self.panel,label="Remove Selected",id=wx.ID_ANY)
		self.controls_gb.Add(self.fileremove,pos=(1,8),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		self.fileclear = wx.Button(self.panel,label="Clear Files",id=wx.ID_ANY)
		self.controls_gb.Add(self.fileclear,pos=(2,8),span=(1,1),flag=wx.EXPAND|wx.ALL,border=8)
		
		
		self.filePathNew_label = wx.StaticText(self.panel,id=wx.ID_ANY,label="Select Save Location")
		self.controls_gb.Add(self.filePathNew_label,pos=(4,0),span=(1,1),flag=wx.ALL|wx.ALIGN_CENTER_VERTICAL,border=8)
		
		self.filePathNew = wx.DirPickerCtrl(self.panel,id=wx.ID_ANY,path=base_dir,message="Select Save Location",style=wx.DIRP_DEFAULT_STYLE,name="newFilePath")
		self.controls_gb.Add(self.filePathNew,pos=(4,1),span=(1,8),flag=wx.ALL|wx.EXPAND,border=8)
		
		
		self.newFileName_label = wx.StaticText(self.panel,id=wx.ID_ANY,label="New File Name")
		self.controls_gb.Add(self.newFileName_label,pos=(5,0),span=(1,1),flag=wx.ALL|wx.ALIGN_CENTER_VERTICAL,border=8)
		
		self.newFileName = wx.TextCtrl(self.panel,id=wx.ID_ANY,style=wx.TE_PROCESS_ENTER,name="newFileName")
		self.controls_gb.Add(self.newFileName,pos=(5,1),span=(1,5),flag=wx.ALL|wx.EXPAND,border=8)
		
		self.fileType = wx.StaticText(self.panel,id=wx.ID_ANY,label=".pdf")
		self.controls_gb.Add(self.fileType,pos=(5,6),span=(1,1),flag=wx.ALL|wx.ALIGN_CENTER_VERTICAL,border=0)
		
		self.printbtn = wx.Button(self.panel,label="Get File(s)",id=wx.ID_ANY)
		self.controls_gb.Add(self.printbtn,pos=(5,7),span=(1,1),flag=wx.ALL,border=8)
		
		
		actions = ["Extract selected pages to single file","Split all pages into separate files","Merge Files"]
		self.radioActions = wx.RadioBox(self.panel,id=wx.ID_ANY,label="Actions",choices=actions,style=wx.RA_SPECIFY_ROWS,name="radioActions")
		self.controls_gb.Add(self.radioActions,pos=(6,0),span=(1,3),flag=wx.EXPAND|wx.ALL,border=8)
		#~ self.radioActions.SetSelection(2)
		
		self.filePages = wx.TextCtrl(self.panel,id=wx.ID_ANY,style=wx.TE_MULTILINE,name="Extract Page Selection")
		self.controls_gb.Add(self.filePages,pos=(6,3),span=(1,6),flag=wx.EXPAND|wx.ALL,border=8)
		
		
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
			returnFiles = []
			fileCount = self.filepath.GetCount()
			for file in range(fileCount):
				returnFiles.append(self.filepath.GetString(file))
			paths = dlg.GetPaths()
			for path in paths:
				if path not in returnFiles:
					self.filepath.Append(path)
		dlg.Destroy()
		
	def removeClick(self,event):
		files = self.filepath.GetSelections()
		for i in reversed(range(len(files))):
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
		
		returnFiles = []
		try:
			fileCount = self.filepath.GetCount()
			for file in range(fileCount):
				returnFiles.append(self.filepath.GetString(file))
		except (Exception) as e:
			errorTemplate = "No files were selected."
			message = errorTemplate.format(type(e).__name__,e.args)
			dlg = wx.MessageDialog(self.panel,message,"ERROR",wx.OK|wx.ICON_ERROR)
			dlg.ShowModal()
			dlg.Destroy()
			return
		newFile = self.newFileName.GetLineText(0)
		if not newFile:
			newFile = "Newpdf"
		newFile += ".pdf"
		
		# Create new file with selected file page(s)
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
				filepos2 = filepos[filepos.find("("):].replace("(","").replace(")","")
				filepos2a = filepos2.split(",")
				iterReturnB = []
				for j in range(len(filepos2a)):
					if "-" in filepos2a[j]:
						rangeBegin = int(filepos2a[j][:filepos2a[j].find("-")])
						rangeEnd = int(filepos2a[j][filepos2a[j].find("-")+1:])
						rangeRange = rangeEnd-rangeBegin
						# if the second number is less than or equal to the first, throw an error
						if (rangeRange)<=0:
							dlg = wx.MessageDialog(self.panel,"The following page range is not valid. \n" + filepos2a[j],"Error",wx.OK|wx.ICON_ERROR)
							dlg.ShowModal()
							dlg.Destroy()
							return
						#~ if the second number is greater than the first, loop through the numbers including and between them and add them to the final array
						else:
							for k in range(rangeRange+1):
								iterReturnB.append(str(rangeBegin))
								rangeBegin += 1
					else:
						iterReturnB.append(filepos2a[j])
				iterReturn.append([filepos1,iterReturnB])
			# print (iterReturn)
			
			writer = PyPDF2.PdfFileWriter()
			outputfile = os.path.join(base_dir,newFile)
			for i in range(len(iterReturn)):
				try:
					with open(returnFiles[(int(iterReturn[i-1][0])-1)],'rb') as inputfile:
						# print(returnFiles[(int(iterReturn[i-1][0])-1)])
						reader = PyPDF2.PdfFileReader(inputfile)
						for j in range(len(iterReturn[i-1][1])):
							writer.addPage(reader.getPage(int(iterReturn[i-1][1][j-1])-1))
							# with open(outputfile,'wb') as out_pdf:writer.write(out_pdf)
				except (Exception,ArithmeticError) as e:
					errorTemplate = "An invalid file number was selected.\nThis likely occured by trying to select a file or page number that does not exist."
					message = errorTemplate.format(type(e).__name__,e.args)
					dlg = wx.MessageDialog(self.panel,message,"ERROR",wx.OK|wx.ICON_ERROR)
					dlg.ShowModal()
					dlg.Destroy()
					return
			dlg = wx.MessageDialog(self.panel,"Pages Successfully Split.","Success",wx.OK)
			dlg.ShowModal()
			dlg.Destroy()
					
		# Split all selected files' pages into separate files
		elif returnAction == 1:
			for i in range(len(returnFiles)):
				with open(returnFiles[i],'rb') as inputfile:
					reader = PyPDF2.PdfFileReader(inputfile)
					pagenum = reader.getNumPages()
					for x in range(pagenum):
						writer = PyPDF2.PdfFileWriter()
						writer.addPage(reader.getPage(x))
						outputfile = returnFiles[i][0:returnFiles[i].find(".")]+"_"+str(x+1)+".pdf"
						with open(outputfile,'wb') as out_pdf: writer.write(out_pdf)
			dlg = wx.MessageDialog(self.panel,"Files Successfully Split.","Success",wx.OK)
			dlg.ShowModal()
			dlg.Destroy()
			
		# Merge all selected files' pages into new file
		elif returnAction == 2:
			writer = PyPDF2.PdfFileWriter()
			outputfile = os.path.join(base_dir,newFile)
			for i in range(len(returnFiles)):
				with open(returnFiles[i],'rb') as inputfile:
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
