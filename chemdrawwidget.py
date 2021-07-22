from PyQt5 import QAxContainer
from PyQt5.QtCore import QSettings
from PyQt5.QtWidgets import QApplication
import time
import winreg


class ChemDrawWidget(QAxContainer.QAxWidget):
    """" Class to create an activeX container of ChemDraw
        function: getSmiles(), getPNG(), getName
    """
    def __init__(self, parent=None, *args, **kwargs):
        super(ChemDrawWidget, self).__init__(parent, *args, **kwargs)
        self.parent = parent
        self.settings = QSettings("Syncom", "Inventory")
        if not self.setControl(self.settings.value("chemdrawctl")):
            self.setCLSID()
        self.state = not(self.isNull())
        if self.state:
            self.resize(250, 250)
            self.setProperty("ShowCrosshair", False)
            self.setProperty("ShowRulers", False)
            self.setProperty("BorderVisible", True)
            self.setMaximumSize(600, 600)

    def mousePressEvent(self, event):
        """To avoid weird behaviour where the widget lose focus
            on the first click and need a second click to get focus"""
        if not self.hasFocus():
            self.setFocus()

    def getSmiles(self):
        params = ["chemical/daylight-smiles"]  ##Got Smiles
        return self.dynamicCall('Data(QVariant)', params)

    def getPng(self):
        params = ["image/png"]  ##Got PNG
        return self.dynamicCall('Data(QVariant)', params)

    def getName(self):
        params = ["chemical/x-name"]  ##Got Name
        return self.dynamicCall('Data(QVariant)', params)

    def isEmpty(self):
        params = ["chemical/daylight-smiles"]
        if len (self.dynamicCall('Data(QVariant)', params)) ==0:
            return True
        else:
            return False

    def clearStructure(self):
        params = ["chemical/daylight-smiles", ""]
        self.dynamicCall('SetData(QVariant,QVariant)',params)

    def setStructure(self, molData, molFormat):
        """  "chemical/x-name", "chemical/daylight-smiles" """
        self.clearStructure()
        params = [molFormat, molData]
        self.dynamicCall('setData(QVariant, QVariant)', params)

    def setCLSID(self):
        for CLSID in get_clsid():
            if self.setControl(CLSID) and not self.isNull():
                self.state = True
                self.settings.setValue("chemdrawctl", CLSID)
                break


def get_clsid():
    """Function used to look through the registry to find the CLSID corresponding
        to the latest installed version of ChemDraw. """
    access_reg = winreg.ConnectRegistry(None, winreg.HKEY_CLASSES_ROOT)
    CLSID = []
    for keyName in get_cdx_ctl_key(access_reg):
        try:
            access_key = winreg.OpenKey(access_reg, r"{}\\".format(keyName), 0, winreg.KEY_READ)
            date = winreg.QueryInfoKey(access_key)[2]
            key = winreg.QueryValue(access_key, "CLSID")
            if (key, date) not in CLSID:
                CLSID.append((key, date))
        except:
            continue
    CLSID.sort(key=lambda info: info[1], reverse=True)  # in order to sort by modified date  - latest first
    CLSID = [clsid[0] for clsid in CLSID]
    return CLSID


def get_cdx_ctl_key(reg_path):
    """Looking throught the registry key to fine the right one with the chemdrawcontrol CLSID."""
    result = []
    for n in range(winreg.QueryInfoKey(reg_path)[0]):
        try:
            x = winreg.EnumKey(reg_path, n)
            if "chemdrawcontrol" in x.lower():
                result.append(x)
        except:
            continue
    return result


if __name__ == '__main__':
    app = QApplication([])
    table = ChemDrawWidget()
    table.show()
    app.exec_()