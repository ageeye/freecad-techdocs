# -*- coding: utf-8 -*-

class TechDocsWB ( Workbench ):

    import TechDocsTools
    MenuText = 'TechDocs'
    ToolTip  = 'A simple wb to export technical stuff to a presentation.'
    Icon     = TechDocsTools.Settings.icon('Presentation.svg')

    def Initialize(self):
        import TechDocsCmds
        self.list = ['makePresentationPlane']
        self.appendToolbar('Presentation',self.list)

    def Activated(self):
        return

    def Deactivated(self):
        return

    def ContextMenu(self, recipient):
        return

    def GetClassName(self):
        return 'Gui::PythonWorkbench'

Gui.addWorkbench(TechDocsWB)