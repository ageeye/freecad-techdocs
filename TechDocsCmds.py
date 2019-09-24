# -*- coding: utf-8 -*-

import FreeCAD, FreeCADGui

class cmdPresentationPlane():
    'Create a plane for presentation.'

    def GetResources(self):
        import TechDocsTools
        return {'Pixmap'  : TechDocsTools.Settings.icon('PresentationPlane.svg'),
                'MenuText': 'PresentationPlane',
                'ToolTip' : 'Create a PresentationPlane.'}

    def cmdType(self):
        return 'PresentationPlane'

    def Activated(self):
        self.proceed()

    def proceed(self):
        FreeCADGui.addModule('TechDocs')
        doc = FreeCAD.ActiveDocument
        doc.openTransaction('PresentationPlane')
        FreeCADGui.doCommand('TechDocs.makePresentationPlane()')
        doc.commitTransaction()
        doc.recompute()
        FreeCADGui.Selection.clearSelection()


FreeCADGui.addCommand('makePresentationPlane', cmdPresentationPlane())
