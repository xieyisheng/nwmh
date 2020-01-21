import win32con
import win32api

class keyBoardKeys(object):
    VK_CODE={'enter':0x0D,'ctrl':0x11,'v':0x56}

    @staticmethod
    def keyDown(keyName):
        win32api.keybd_event(keyBoardKeys.VK_CODE [keyName],0,0,0)

    @staticmethod
    def keyUp(keyName):
        win32api.keybd_event(keyBoardKeys.VK_CODE [keyName],0,win32con.KEYEVENTF_KEYUP ,0)

    @staticmethod
    def oneKey(key):
        keyBoardKeys .keyDown(key)
        keyBoardKeys .keyUp(key)

    @staticmethod
    def twoKeys(key1,key2):
        keyBoardKeys .keyDown(key1)
        keyBoardKeys.keyDown(key2)
        keyBoardKeys .keyUp(key1)
        keyBoardKeys.keyUp(key2)

        