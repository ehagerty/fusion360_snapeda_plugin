
import adsk.core
import adsk.fusion
import adsk.cam

import traceback
import subprocess
import os
import platform
import threading
import json
import random
import signal
import json
import logging
import math


# Configure logger.
# Remove all handlers associated with the root logger object.
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

logging.basicConfig(filename=os.path.join(os.path.dirname(__file__), 'snapeda.log'), level=logging.DEBUG, filemode = 'w')


# PID of SnapEDA plugin
process_id = []
lib_path = None
commandDefinition = None
toolbarControl = None
handlers = []


class OpenFromWebExecutedEventHandler(adsk.core.CommandEventHandler):

    def __init__(self):
        super().__init__()

    def notify(self, args):

        ui = None
        try:
            app = adsk.core.Application.get()
            ui = app.userInterface

            filename = "SnapEDA for Fusion 360" if platform.system() == "Windows" else "SnapEDA for Fusion 360.app/Contents/MacOS/SnapEDA for Fusion 360"
            path = os.path.join(os.path.dirname(__file__), filename)

            # set PID and start SnapEDA plugin
            global process_id
            process_id.append(subprocess.Popen([path, ]).pid)
                
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


class OpenFromWebCreatedEventHandler(adsk.core.CommandCreatedEventHandler):

    def __init__(self):
        super().__init__()

    def notify(self, args):
        ui = None
        try:
            app = adsk.core.Application.get()
            ui = app.userInterface

            # connect to the command executed event.
            cmd = args.command
            onExecute = OpenFromWebExecutedEventHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
        except:
            if ui:
                ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def save_document(filename='SnapEDA Library'):
    """
    Initially save document.
    """
    logging.info('save_document::filename=' + str(filename))
    app = adsk.core.Application.get()
    document = app.activeDocument
    data = app.data
    project = data.dataProjects.item(0)
    cfolder = project.rootFolder
    if not document.isSaved:
        returnValue = document.saveAs(filename, cfolder, '', '')
    document.save('Imported from SnapEDA')
    logging.info('save_document::document saved')
    return document

def getTargetLibraryURN():
    """
    Retrieves the saved target library URN.
    """
    logging.info('getTargetLibraryURN::')
    urn_text_file = 'team_urn.txt'
    urn_text_file_path = os.path.join(os.path.dirname(__file__), urn_text_file)
    target_library_urn = None

    if (os.path.exists(urn_text_file_path)):
        urn_text_file_handler = open(urn_text_file_path, 'r')
        target_library_urn = urn_text_file_handler.readline().replace('\n', '')

    try:
        # check if URN really exists
        app = adsk.core.Application.get()
        data_file = app.data.findFileById(target_library_urn)
    except Exception as e:
        logging.info(e)
        target_library_urn = None



    if not target_library_urn:
        try:
            logging.info('no target_library_urn')
            # check for any existing SnapEDA Library
            document = app.activeDocument
            data = app.data
            project = data.dataProjects.item(0)
            cfolder = project.rootFolder
            dataFiles = cfolder.dataFiles
            target_library_urn = None
            for i in range(dataFiles.count):
                dataFile = dataFiles.item(i)
                if(dataFile.name == 'SnapEDA Library'):
                    logging.info('SnapEDA Library already exists.')
                    logging.info(dataFile.name)
                    logging.info(dataFile.id)
                    target_library_urn = dataFile.id
                    # write URN
                    file = open(os.path.join(os.path.dirname(__file__), urn_text_file), 'w')
                    file.write(str(target_library_urn) + '\n')
                    file.close()
                    break
        except Exception as e:
            logging.info('Error retrieving URN')
            logging.info(e)

    return target_library_urn

def open_target_library(target_urn: str):
    """
    Opens the target library.
    """
    logging.info('open_target_library::')
    app = adsk.core.Application.get()

    # close opened document
    app.activeDocument.close(False)

    data_file = app.data.findFileById(target_urn)
    library_document = app.documents.open(data_file)
    return library_document

def get_samples_dir():
    """
    Get SnapEDALibrary directory.
    """
    # source_file_directory_name = 'SnapEDALibrary'
    # root_directory = os.path.expanduser("~")
    global lib_path
    samples_dir = os.path.join(lib_path, '')
    return samples_dir

def import_ecad_library(source_name, package_name, package_type):
    """
    Import CAD Library for use.
    """
    logging.info('import_ecad_library::')

    if package_type == 'devicesets':
        format_type = 'dev'
    if package_type == 'packages':
        format_type = 'pac'
    if package_type == 'symbols':
        format_type = 'sym'


    app = adsk.core.Application.get()
    source_path = os.path.join(get_samples_dir(), source_name)
    logging.info('source_path=' + str(source_path))

    c1 = f'Use "{source_path}"'
    c2 = f'Copy "{package_name}.{format_type}@{source_name}"'
    c3 = f'Use "-{source_name}"'

    text_cmd = f'Electron.run "{c1}"'
    app.executeTextCommand(text_cmd)
    logging.info(text_cmd)

    text_cmd = f'Electron.run "{c2}"'
    app.executeTextCommand(text_cmd)
    logging.info(text_cmd)

    text_cmd = f'Electron.run "{c3}"'
    app.executeTextCommand(text_cmd)
    logging.info(text_cmd)

def extractPackageName(source_path, package_type):
    """
    Extracts the package name from the XML.
    """
    logging.info('extractPackageName::')
    import xml.etree.ElementTree as ET
    package_name = None
    try:
        tree = ET.parse(source_path)
        root = tree.getroot()
        package_name = root[0].find('library').find(package_type)[0].attrib['name']
    except Exception as e:
        logging.info(e)

    return package_name

def create_3d_package(package_lib_name):
    """
    Creates a 3D package.
    """
    logging.info('create_3d_package::')
    app = adsk.core.Application.get()
    package_lib = os.path.join(get_samples_dir(), package_lib_name)
    text_cmd = f'Electron.Create3DPackage "{package_lib}"'
    app.executeTextCommand(text_cmd)
    logging.info('3D package created')

def save_as_document(target_project_id):
    """
    Save current document.
    """

    app = adsk.core.Application.get()
    # target_project_id = app.activeDocument.dataFile.parentProject.id
    data = app.data
    target_project = data.dataProjects.item(0)
    # target_project = get_project_from_id(target_project_id)
    target_folder = target_project.rootFolder
    package_document = app.activeDocument

    package_document.saveAs('', target_folder, 'Empty SnapEDA 3D Package', '')

    return package_document

def import_3d_model(source_3d_name):
    """
    Import the 3D model.
    """
    logging.info('import_3d_model::')
    app = adsk.core.Application.get()
    ui = app.userInterface

    # Get active design
    product = app.activeProduct
    design = adsk.fusion.Design.cast(product)

    # Get import manager
    import_manager = app.importManager

    # Get step import options
    stp_file = os.path.join(get_samples_dir(), source_3d_name)
    stp_options = import_manager.createSTEPImportOptions(stp_file)

    # Import step file to root component
    result = import_manager.importToTarget2(stp_options, design.rootComponent)

    return result


def finish_up(document: adsk.core.Document, close: bool):
    logging.info('finish_up::')
    document.activate()
    document.save('Imported from SnapEDA')

    # This is optional
    # It may be desirable to not close the document if the user still needs to align the geometry
    # if close:
    #     document.close(False)

def rotate_step(result):
    app = adsk.core.Application.get()
    ui = app.userInterface

    # Get the resulting occurrence.
    occ = adsk.fusion.Occurrence.cast(result.item(0))

    # Rotate the occurrence.
    angle = math.pi * 0.5
    mat = adsk.core.Matrix3D.create()
    mat.setWithCoordinateSystem(adsk.core.Point3D.create(0,0,0),
                                adsk.core.Vector3D.create(1,0,0),
                                adsk.core.Vector3D.create(0, math.cos(angle), math.sin(angle)),
                                adsk.core.Vector3D.create(0, math.cos(angle + (math.pi/2)), math.sin(angle + (math.pi/2))))
    occ.transform = mat

    if ui:
        ui.messageBox('Please make sure to orient the 3D model properly in the package editor.')

class MyOpenedFromURLHandler(adsk.core.WebRequestEventHandler):

    def __init__(self):
        super().__init__()

    def notify(self, args):
        app = adsk.core.Application.get()
        ui = app.userInterface

        logging.info('Document opened from URL.')

        if not isinstance(args.privateInfo, dict):
            privateInfo = json.loads(args.privateInfo)
        global lib_path
        lib_path = os.path.normpath(privateInfo.get('lib_path', ''))
        target_library = os.path.normpath(privateInfo.get('target_library', ''))

        logging.info(str(privateInfo))
        logging.info('lib_path=' + str(lib_path))
        logging.info('target_library=' + str(target_library))

        if '.lbr' in target_library:
            _, tail = os.path.split(target_library)
            filename = os.path.splitext(tail)[0]
            filename = filename if platform.system() == "Windows" else filename.split('/')[-1]
            logging.info('filename=' + str(filename))
            

            target_library_urn = getTargetLibraryURN()
            logging.info(target_library_urn)
            logging.info('opening library document')
            library_document = open_target_library(target_library_urn) if target_library_urn else save_document()

            source_name = filename + '.lbr'
            package_type = 'devicesets'
            package_name = extractPackageName(os.path.join(get_samples_dir(), source_name), package_type)
            if not package_name:
                package_type = 'packages'
                package_name = extractPackageName(os.path.join(get_samples_dir(), source_name), package_type)
            logging.info('source_name=' + str(source_name))
            logging.info('package_type=' + str(package_type))
            logging.info('package_name=' + str(package_name))
            # if not package_name:
            #     package_type = 'symbols'
            #     package_name = extractPackageName(os.path.join(get_samples_dir(), source_name), package_type)

            logging.info(package_name + '::' + package_type)

            if package_name:
                import_ecad_library(source_name, package_name, package_type)
                logging.info('imported ecad library')
                logging.info(source_name)

                if os.path.exists(os.path.join(get_samples_dir(), filename + '.step')):
                    create_3d_package(source_name)
                    package_document = save_as_document('')
                    result = import_3d_model(filename + '.step')
                    rotate_step(result)
                else:
                    logging.info('no 3D model')

                # before finishing try to get URN
                try:
                    document = app.activeDocument
                    data = app.data
                    project = data.dataProjects.item(0)
                    cfolder = project.rootFolder
                    dataFiles = cfolder.dataFiles
                    target_library_urn = None
                    for i in range(dataFiles.count):
                        dataFile = dataFiles.item(i)
                        if(dataFile.name == 'SnapEDA Library'):
                            logging.info(dataFile.name)
                            logging.info(dataFile.id)
                            target_library_urn = dataFile.id
                            # write URN
                            file = open(os.path.join(os.path.dirname(__file__), 'team_urn.txt'), 'w')
                            file.write(str(target_library_urn) + '\n')
                            file.close()
                            break
                except Exception as e:
                    logging.info('Error retrieving URN')
                    logging.info(e)

            # finish_up(package_document)
            # finish_up(library_document)


def run(context):
    ui = None

    try:
        commandId = 'OpenHtmlPageCommandIdPy'
        commandName = 'SnapEDA Plugin'
        commandDescription = 'Open the SnapEDA Plugin to login and download components.'
        panelId = 'SolidScriptsAddinsPanel'

        app = adsk.core.Application.get()
        ui = app.userInterface

        filename = ""
        if platform.system() == "Darwin":
            filename = "snapeda-fusion-plugin.app/Contents/MacOS/snapeda-fusion-plugin"
        if platform.system() == "Windows":
            try:
                updated_plugin_path = os.path.join(os.path.dirname(
                    __file__), "new-snapeda-fusion-plugin.exe")
                current_plugin_path = os.path.join(
                    os.path.dirname(__file__), "snapeda-fusion-plugin.exe")
                if(os.path.exists(updated_plugin_path)):
                    os.remove(current_plugin_path)
                    os.rename(updated_plugin_path, current_plugin_path)
            except:
                pass
            filename = "snapeda-fusion-plugin"

        # path = os.path.join(os.path.dirname(__file__), filename)
        # subprocess.Popen([path,])

        # create the command definition
        commandDefinition = ui.commandDefinitions.itemById(commandId)
        # delete any existing command definition, and just recreate it
        if commandDefinition:
            commandDefinition.deleteMe()
        commandDefinition = ui.commandDefinitions.addButtonDefinition(
            commandId, commandName, commandDescription, './/Resources//SnapEDA')
        onCommandCreated = OpenFromWebCreatedEventHandler()
        commandDefinition.commandCreated.add(onCommandCreated)

        # keep the handler referenced beyond this function
        handlers.append(onCommandCreated)

        # insert the command into the model:add-ins toolbar
        toolbarControls = ui.allToolbarPanels.itemById(panelId).controls
        # delete any existing control, and just recreate it
        global toolbarControl
        toolbarControl = toolbarControls.itemById(commandId)
        if toolbarControl:
            toolbarControl.deleteMe()
        # toolbarControl = toolbarControls.addCommand(commandDefinition)
        # toolbarControl.isPromotedByDefault = True

        workSpace = ui.workspaces.itemById('FusionSolidEnvironment')
        tbPanels = workSpace.toolbarPanels
        snapEDAPanel = tbPanels.itemById('SnapEDAPanel')
        if snapEDAPanel:
            snapEDAPanel.deleteMe()
        snapEDAPanel = tbPanels.add('SnapEDAPanel', 'SnapEDA')
        snapEDAPanelControl = tbPanels.itemById(
            'SnapEDAPanel').controls.itemById(commandId)
        if snapEDAPanelControl:
            snapEDAPanelControl.deleteMe()
        snapEDAPanelControl = tbPanels.itemById(
            'SnapEDAPanel').controls.addCommand(commandDefinition)
        snapEDAPanelControl.isPromotedByDefault = True

        # insert the command into the toolbar
        boardLayoutEnvironmentToolbarPanels = ui.workspaces.itemById(
            'BoardLayoutEnvironement').toolbarPanels

        # for i in range(boardLayoutEnvironmentToolbarPanels.count):
        #     logging.info(boardLayoutEnvironmentToolbarPanels.item(i).id)

        boardLayoutToolbarControls = boardLayoutEnvironmentToolbarPanels.itemById(
            'PcbPlacePanel').controls
        # delete any existing control, and just recreate it
        boardLayoutToolbarControl = boardLayoutToolbarControls.itemById(
            commandId)
        if boardLayoutToolbarControl:
            boardLayoutToolbarControl.deleteMe()
        # boardLayoutToolbarControl = boardLayoutToolbarControls.addCommand(commandDefinition)

        # for i in range(boardLayoutToolbarControls.count):
        #     toolbarCtrl = boardLayoutToolbarControls.item(i)
        #     toolbarCtrl.isPromoted = False
        #     toolbarCtrl.isPromotedByDefault = False

        # boardLayoutToolbarControl.isPromoted  = True
        # boardLayoutToolbarControl.isPromotedByDefault = True

        anotherWorkspace = ui.workspaces.itemById('BoardLayoutEnvironement')
        snapTbPanels = anotherWorkspace.toolbarPanels
        snapPanel = snapTbPanels.itemById('SnapEDAPanel')
        if snapPanel:
            snapPanel.deleteMe()

        snapPanel = snapTbPanels.add(
            'SnapEDAPanel', 'SnapEDA', 'PcbPlacePanel', False)
        snapPanelControl = snapTbPanels.itemById(
            'SnapEDAPanel').controls.itemById(commandId)
        if snapPanelControl:
            snapPanelControl.deleteMe()
        snapPanelControl = snapTbPanels.itemById(
            'SnapEDAPanel').controls.addCommand(commandDefinition)
        snapPanelControl.isPromotedByDefault = True

        # "application_var" is a variable referencing an Application object.
        onOpenedFromURL = MyOpenedFromURLHandler()
        app.openedFromURL.add(onOpenedFromURL)
        handlers.append(onOpenedFromURL)

    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))


def stop(context):
    ui = None
    try:
        app = adsk.core.Application.get()
        ui = app.userInterface
        # stop the SnapEDA plugin
        try:
            global process_id
            if process_id:
                for process in process_id:
                    if platform.system() == "Windows":
                        # kill if invoked during installation
                        subprocess.Popen('taskkill /F /PID ' +
                                        str(process) + ' /T')
                        subprocess.Popen(
                            'taskkill /F /IM "SnapEDA for Autodesk Fusion 360.exe" /T')
                    elif platform.system() == "Darwin":
                        os.kill(process, signal.SIGKILL)
        except Exception as e:
            ui.messageBox(e)
    except:
        if ui:
            ui.messageBox('Failed:\n{}'.format(traceback.format_exc()))
