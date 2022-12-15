import wx
import EventApp

if __name__ == '__main__':
    Informart = wx.App(False)
    Cea = EventApp.CEventApp
    frame = Cea(None)
    frame.Show(True)
    Informart.MainLoop()
    