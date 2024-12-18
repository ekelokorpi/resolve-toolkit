import sys
import os
import datetime
import pathlib
import avb
from reportlab.pdfgen.canvas import Canvas

winID = "com.kelokorpi.toolkit"
version = 'v1.5.2'

# Settings
supportedMediaFiles = ['.MXF', '.MP4', '.MOV', '.WAV', '.CRM']
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
    ui.Button({ 'ID': 'ImportAvid',  'Text': "Import Avid bins" }),

    ui.Label(),

    ui.Button({ 'ID': 'CopyAssets',  'Text': "Consolidate assets to new folder" }),
    ui.Button({ 'ID': 'ColorClips',  'Text': "Color and number clips based on TC and Duration" }),
    ui.Button({ 'ID': 'Musatiedot',  'Text': "Generate musatiedot file" }),
    ui.Button({ 'ID': 'CalculateClips',  'Text': "Calculate total size of timeline clips" }),
    ui.Button({ 'ID': 'CopyRelinkClips',  'Text': "Copy and relink clips from timeline to new location" }),
    # ui.Button({ 'ID': 'ClearShots',  'Text': "Clear shot numbers" }),

    ui.Label(),

    ui.Button({ 'ID': 'UpdateToolkit',  'Text': "Update toolkit" }),
]

win = dispatcher.AddWindow({
        'ID': winID,
        'Geometry': [ 300, 300, 400, 500 ],
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

def ImportMultiAudioToCurrentFolder(path):
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

def OnImportMultiAudioCurrent(ev):
    selectedPath = fusion.RequestDir()
    if selectedPath:
        ImportMultiAudioToCurrentFolder(selectedPath)

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


class MusicInfo:
  def __init__(self, artist, song, duration):
    self.artist = artist
    self.song = song
    self.duration = duration

def OnMusatiedot(ev):
    DisableAllButtons()
    print('Working...')

    # Works with Epidemic Sound and Artlist

    musicInfos = []

    project = projectManager.GetCurrentProject()

    timeline = project.GetCurrentTimeline()
    frameRate = timeline.GetSetting('timelineFrameRate')
    timelineName = timeline.GetName()
    projectName = project.GetName()
    audioTrackCount = timeline.GetTrackCount("audio")
    print('test')
    for i in range(audioTrackCount):
        trackName = timeline.GetTrackName("audio", i + 1)
        if not "Music" in trackName:
            continue
        items = timeline.GetItemListInTrack("audio", i + 1)
        print('test2')
        for item in items:
            name = item.GetName()
            enabled = item.GetClipEnabled()
            if enabled is False:
                continue
            if ".wav" in name and " - " in name and not "SFX" in name:
                isEpidemic = False
                if "ES_" in name:
                    isEpidemic = True
                name = name.replace("ES_", "")
                name = name.replace(".wav", "")
                name = name.split(" - ")
                itemDuration = item.GetDuration()
                print('test3')
                print(itemDuration)
                print(frameRate)
                durationInSeconds = int(itemDuration) / int(frameRate)
                print('test4')

                if isEpidemic == True:
                    artistName = name[1]
                    songName = name[0]
                else:
                    artistName = name[0]
                    songName = name[1]

                found = False
                for musicInfo in musicInfos:
                    if musicInfo.artist == artistName and musicInfo.song == songName:
                        musicInfo.duration = musicInfo.duration + durationInSeconds
                        found = True
                        break

                if found is False:
                    musicInfos.append(MusicInfo(artistName, songName, durationInSeconds))

    
    if len(musicInfos) > 0:
        selectedPath = fusion.RequestDir()
        if selectedPath:
            fileName = projectName.replace(" ", "_").lower() + "_musatiedot.pdf"

            canvas = Canvas(selectedPath + fileName)
            y = 800
            lineHeight = 20
            canvas.drawString(72, y, projectName)
            y = y - lineHeight;
            canvas.drawString(72, y, "Musiikkitiedot")
            y = y - lineHeight * 2;
            for musicInfo in musicInfos:
                canvas.drawString(72, y, musicInfo.artist + ' - ' + musicInfo.song)
                y = y - lineHeight;
                canvas.drawString(72, y, str(datetime.timedelta(seconds=round(musicInfo.duration))))
                y = y - lineHeight * 2;
            canvas.save()
                
    print('MUSATIEDOT:')
    print("")
    print(projectName)
    print("")
    for musicInfo in musicInfos:
        print(musicInfo.artist + ' - ' + musicInfo.song)
        print(str(datetime.timedelta(seconds=round(musicInfo.duration))))
        print("")
    print('DONE')
    EnableAllButtons()

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

def convert_bytes(num):
    for x in ['bytes', 'KB', 'MB', 'GB', 'TB']:
        if num < 1024.0:
            return "%3.1f %s" % (num, x)
        num /= 1024.0

def OnCalculateClips(ev):
    DisableAllButtons()
    print('Working...')

    totalSize = 0

    calculatedClips = []

    timeline = project.GetCurrentTimeline()
    videoTrackCount = timeline.GetTrackCount("video")
    for i in range(videoTrackCount):
        items = timeline.GetItemListInTrack("video", i + 1)
        for item in items:
            name = item.GetName()
            mediaItem = item.GetMediaPoolItem()
            enabled = item.GetClipEnabled()
            if mediaItem != None and enabled == True:
                path = mediaItem.GetClipProperty('File Path')
                baseName = os.path.basename(path)
                if path not in calculatedClips:
                    calculatedClips.append(path)
                    file_stats = os.stat(path)
                    totalSize = totalSize + file_stats.st_size

    totalSizeFormatted = convert_bytes(totalSize)
    print('Total size of timeline clips is ' + totalSizeFormatted)
    print('DONE')
    EnableAllButtons()

def OnCopyRelinkClips(ev):
    selectedPath = fusion.RequestDir()
    print(selectedPath)
    if selectedPath == None:
        return
    DisableAllButtons()
    print('Working...')

    timeline = project.GetCurrentTimeline()
    videoTrackCount = timeline.GetTrackCount("video")
    for i in range(videoTrackCount):
        items = timeline.GetItemListInTrack("video", i + 1)
        for item in items:
            name = item.GetName()
            mediaItem = item.GetMediaPoolItem()
            enabled = item.GetClipEnabled()
            if mediaItem != None and enabled == True:
                path = mediaItem.GetClipProperty('File Path')
                baseName = os.path.basename(path)
                newFileDest = os.path.join(selectedPath, baseName)
                isFile = os.path.isfile(newFileDest)

                print("Copying file " + path)
                if isFile == False:
                    newFile = shutil.copy(path, selectedPath)
                    print("Copy done. Replacing clip...")
                    mediaItem.ReplaceClip(newFile)
                else:
                    print("Skipping file. Already exists.")
                    mediaItem.ReplaceClip(newFileDest)

    print('DONE')
    EnableAllButtons()

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

def OnImportAvid(ev):
    currentFolder = mediaPool.GetCurrentFolder()

    selectedPath = fusion.RequestDir()
    print(selectedPath)
    if selectedPath == None:
        return
    DisableAllButtons()
    print('Working...')

    for (root, dirs, files) in os.walk(selectedPath):
        for f in files:
            fileExt = pathlib.Path(f).suffix
            if fileExt.upper() == '.AVB':
                print(f)

                folderName = f.split('.', 1)[0]

                folders = currentFolder.GetSubFolders()
                folderFound = False
                for i in folders:
                  if folders[i].GetName() == folderName:
                    newFolder = folders[i]
                    folderFound = True
                    break

                if folderFound == False:
                  newFolder = mediaPool.AddSubFolder(currentFolder, folderName)

                with avb.open(root + '/' + f) as avbfile:

                  for mob in avbfile.content.mobs:

                    if mob.mob_type_id == 1 or mob.mob_type_id == 2:
                      clipName = mob.name.split('.', 1)[0]
                      clipName = clipName.split('_P', 1)[0]
                      print(clipName)

                      mediaPool.SetCurrentFolder(currentFolder)
                      clips = currentFolder.GetClipList()
                      for clip in clips:
                          filename = clip.GetName()
                          filename = filename.split('.', 1)[0]
                          if filename == clipName:
                            print('MATCH FOUND!')
                            mediaPool.MoveClips([clip], newFolder)

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
win.On['Musatiedot'].Clicked = OnMusatiedot
win.On['CopyRelinkClips'].Clicked = OnCopyRelinkClips
win.On['CalculateClips'].Clicked = OnCalculateClips
win.On['ClearShots'].Clicked = OnClearShots
win.On['ImportMultiMC'].Clicked = OnImportMultiMC
win.On['ShowConsole'].Clicked = OnShowConsole
win.On['ImportMultiMCCurrent'].Clicked = OnImportMultiMCCurrent
win.On['ImportMultiAudioCurrent'].Clicked = OnImportMultiAudioCurrent
win.On['UpdateToolkit'].Clicked = OnUpdateToolkit
win.On['CopyAssets'].Clicked = OnCopyAssets
win.On['ImportAvid'].Clicked = OnImportAvid

win.Show()
dispatcher.RunLoop()
