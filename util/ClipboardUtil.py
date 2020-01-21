import win32clipboard as w
import win32con

class Clipboard(object ):
    @staticmethod
    def getText():
        w.OpenClipboard()
        d=w.GetClipboardData(win32con .CF_TEXT )
        w.CloseClipboard()
        return d

    @staticmethod
    def setText(aString):
        w.OpenClipboard()
        w.EmptyClipboard()
        w.SetClipboardData(win32con.CF_UNICODETEXT ,aString )
        w.CloseClipboard()
        