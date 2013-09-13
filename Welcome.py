import wx 
import threading
import time

class Worker(threading.Thread):
	def __init__(self,parent):
		threading.Thread.__init__(self)
		self.parent = parent

	def run(self):
		time.sleep(5)
		self.parent.Destroy()

class WelcomeScreen(wx.Dialog):
	def __init__(self, *args, **kwds):
		wx.Dialog.__init__(self, *args, size = (550,400))
		WelcomePanel = wx.Panel(self)
		
		Wsizer = wx.BoxSizer(wx.VERTICAL)
		hbox = wx.BoxSizer(wx.HORIZONTAL)
		
		WelcomeFont = wx.Font(24,wx.DEFAULT,wx.NORMAL,wx.BOLD)
		WelcomeText = "Patient Mapping Software"
		WelcomeStaticText = wx.StaticText(WelcomePanel,label = WelcomeText)
		WelcomeStaticText.SetFont(WelcomeFont)
		PatmapsLogo = wx.StaticBitmap(WelcomePanel)
		PatmapsLogo.SetBitmap(wx.Bitmap("FullLogo.png"))
		CardiffLogo = wx.StaticBitmap(WelcomePanel)
		CardiffLogo.SetBitmap(wx.Bitmap("Logo.ico"))
		
		hbox.Add(WelcomeStaticText,flag = wx.ALIGN_CENTRE)
		hbox.Add((20,0))
		hbox.Add(CardiffLogo,flag = wx.ALIGN_CENTRE)
		
		Wsizer.Add((0,20))
		Wsizer.Add(PatmapsLogo,flag = wx.ALIGN_CENTRE)
		Wsizer.Add((0,20))
		Wsizer.Add(hbox,flag = wx.ALIGN_CENTRE)
		
		
		WelcomePanel.SetSizer(Wsizer)
		
		self.Centre()
		Worker2 = Worker(self)
		Worker2.start()
		