import sys
import os
import datetime
import pathlib

winID = "com.kelokorpi.toolkit"
version = 'v1.2.0'

# Settings
supportedMediaFiles = ['.MXF', '.MP4', '.MOV']
clipColors = ['Orange', 'Blue', 'Pink', 'Green', 'Yellow', 'Teal', 'Violet', 'Brown']

ui = fusion.UIManager
dispatcher = bmd.UIDispatcher(ui)

win = ui.FindWindow(winID)
if win:
    win.Show()
    win.Raise()
    exit()

buttons = [
    ui.Button({ 'ID': 'ImportMultiMCCurrent',  'Text': "Import media folders into current folder" }),
    ui.Button({ 'ID': 'ImportMultiMC',  'Text': "Import media folders into subfolders" }),
    ui.Button({ 'ID': 'ImportMC',  'Text': "Import Canon RAW + PROXY folder" }),

    ui.Label(),

    ui.Button({ 'ID': 'CopyAssets',  'Text': "Consolidate assets to new folder" }),
    
    ui.Button({ 'ID': 'ColorClips',  'Text': "Color and number clips based on TC and Duration" }),
    # ui.Button({ 'ID': 'ClearShots',  'Text': "Clear shot numbers" }),

    ui.Label(),

    ui.Button({ 'ID': 'UpdateToolkit',  'Text': "Update toolkit" }),
]

win = dispatcher.AddWindow({
        'ID': winID,
        'Geometry': [ 300, 300, 400, 300 ],
        'WindowTitle': "Resolve Toolkit " + version,
    },
    ui.VGroup(buttons))

projectManager = resolve.GetProjectManager()
project = projectManager.GetCurrentProject()
mediaPool = project.GetMediaPool()
startFolder = mediaPool.GetCurrentFolder()

def FindProxy(fileName, mediaItem, path):
    fileName = fileName.replace('.CRM', '_P.MXF')
    for (root, dirs, file) in os.walk(path):
        for f in file:
            if f == fileName:
                print('PROXY FOUND!')
                mediaItem.LinkProxyMedia(root + '/' + f)

def ImportClips(path):
    DisableAllButtons()
    print('Working...')
    for (root, dirs, file) in os.walk(path):
        for f in file:
            if '.CRM' in f:
                print(root + '/' + f)
                clips = mediaPool.ImportMedia(root + '/' + f)
                FindProxy(f, clips[0], path)
    print('DONE')
    EnableAllButtons()

def ImportClipsToFolder(path, folder):
    rootFolder = mediaPool.GetRootFolder()

    subFolderCreated = False

    if folder == None:
        subFolderCreated = True

    clipsToImport = []

    for (root, dirs, file) in os.walk(path):
        for f in file:
            fileExt = pathlib.Path(f).suffix
            if fileExt.upper() in supportedMediaFiles:
                if subFolderCreated == False:
                    newFolder = mediaPool.AddSubFolder(startFolder, folder)
                    mediaPool.SetCurrentFolder(newFolder)
                    subFolderCreated = True

                # clips = mediaPool.ImportMedia(root + '/' + f)
                clipsToImport.append(root + '/' + f)
                # FindProxy(f, clips[0], path)

    mediaPool.ImportMedia(clipsToImport)

def OnClose(ev):
    dispatcher.ExitLoop()

def ImportMultiClips(path):
    DisableAllButtons()
    print('Working...')
    for (root, dirs, files) in os.walk(path):
        # print(dirs)
        for curDir in dirs:
            # print(curDir)
            if curDir == 'CacheClip':
                continue
            if curDir == 'ProxyMedia':
                continue
            ImportClipsToFolder(root + '/' + curDir, curDir)
        break
    print('DONE')
    EnableAllButtons()

def OnImportMultiMC(ev):
    selectedPath = fusion.RequestDir()
    if selectedPath:
        ImportMultiClips(selectedPath)

def ImportMultiClipsToCurrentFolder(path):
    DisableAllButtons()
    print('Working...')
    for (root, dirs, files) in os.walk(path):
        for curDir in dirs:
            if curDir == 'CacheClip':
                continue
            if curDir == 'ProxyMedia':
                continue
            ImportClipsToFolder(root + '/' + curDir, None)
        break
    print('DONE')
    EnableAllButtons()

def OnImportMultiMCCurrent(ev):
    selectedPath = fusion.RequestDir()
    if selectedPath:
        ImportMultiClipsToCurrentFolder(selectedPath)

def OnExec(ev):
    selectedPath = fusion.RequestDir()
    print(selectedPath)
    if selectedPath:
        ImportClips(selectedPath)

coloredClips = []
currentColor = 0
curShotNumber = 1

def TimecodeToSeconds(timecode):
    t = datetime.datetime.strptime(timecode, "%H:%M:%S:%f")
    seconds = t.second + t.minute * 60 + t.hour * 3600
    return seconds

def SearchSimilar(targetClip, clips):
    global currentColor
    global curShotNumber
    global coloredClips
    targetStart = TimecodeToSeconds(targetClip.GetClipProperty('Start TC'))
    targetDuration = TimecodeToSeconds(targetClip.GetClipProperty('Duration'))

    similarFound = False
    currentCamera = 2

    for clip in clips:
        clipType = clip.GetClipProperty('Type')
        if clip in coloredClips:
            continue
        if clip == targetClip:
            continue
        if clipType != 'Video + Audio':
            continue

        clipStart = TimecodeToSeconds(clip.GetClipProperty('Start TC'))
        clipDuration = TimecodeToSeconds(clip.GetClipProperty('Duration'))

        if abs(targetStart - clipStart) <= 60 and abs(targetDuration - clipDuration) <= 60:
            # print("MATCH FOUND!")
            clip.SetClipColor(clipColors[currentColor])
            clip.SetMetadata('Camera ID', str(currentCamera))
            clip.SetClipProperty('Shot', str(curShotNumber))
            currentCamera = currentCamera + 1
            coloredClips.append(clip)
            similarFound = True

    if similarFound == True:
        targetClip.SetClipColor(clipColors[currentColor])
        targetClip.SetMetadata('Camera ID', '1')
        targetClip.SetClipProperty('Shot', str(curShotNumber))
        coloredClips.append(targetClip)
        currentColor = currentColor + 1
        if currentColor + 1 == len(clipColors):
            currentColor = 0
        curShotNumber = curShotNumber + 1


def OnColorClips(ev):
    DisableAllButtons()
    print('Working...')
    global coloredClips
    global curShotNumber
    curShotNumber = 1
    coloredClips = []
    currentFolder = mediaPool.GetCurrentFolder()
    clips = currentFolder.GetClipList()
    for clip in clips:
        startTC = clip.GetClipProperty('Start TC')
        clipType = clip.GetClipProperty('Type')
        if clipType != 'Video + Audio':
            continue
        if clip in coloredClips:
            continue
        SearchSimilar(clip, clips)
    print('DONE')
    EnableAllButtons()

def OnClearShots(ev):
    currentFolder = mediaPool.GetCurrentFolder()
    clips = currentFolder.GetClipList()
    for clip in clips:
        clipType = clip.GetClipProperty('Type')
        if clipType != 'Video + Audio':
            continue
        clip.SetClipProperty('Shot', '')

def OnShowConsole(ev):
    fusion.ShowConsole()

import inspect
import urllib.request

def OnUpdateToolkit(ev):
    DisableAllButtons()
    print('Working...')
    filename = inspect.getframeinfo(inspect.currentframe()).filename
    url = 'https://raw.githubusercontent.com/ekelokorpi/resolve-toolkit/main/Toolkit.py'
    a,b = urllib.request.urlretrieve(url, filename)
    print(b)
    print('DONE')
    EnableAllButtons()

import shutil

def OnCopyAssets(ev):
    selectedPath = fusion.RequestDir()
    print(selectedPath)
    if selectedPath == None:
        return
    DisableAllButtons()
    print('Working...')
    currentFolder = mediaPool.GetCurrentFolder()
    files = currentFolder.GetClipList()
    for file in files:
        fileType = file.GetClipProperty('Type')
        if fileType == 'Still' or fileType == 'Audio':
            filePath = file.GetClipProperty('File Path')
            newFile = shutil.copy(filePath, selectedPath)
            file.ReplaceClip(newFile)
    print('DONE')
    EnableAllButtons()

def DisableAllButtons():
    for button in buttons:
        button.Enabled = False

def EnableAllButtons():
    for button in buttons:
        button.Enabled = True

win.On[winID].Close = OnClose
win.On['ImportMC'].Clicked = OnExec
win.On['ColorClips'].Clicked = OnColorClips
win.On['ClearShots'].Clicked = OnClearShots
win.On['ImportMultiMC'].Clicked = OnImportMultiMC
win.On['ShowConsole'].Clicked = OnShowConsole
win.On['ImportMultiMCCurrent'].Clicked = OnImportMultiMCCurrent
win.On['UpdateToolkit'].Clicked = OnUpdateToolkit
win.On['CopyAssets'].Clicked = OnCopyAssets

win.Show()
dispatcher.RunLoop()
