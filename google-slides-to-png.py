from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import yaml
import urllib3
import re
import uuid
import argparse

import sys

from hashlib import sha256

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/presentations', 'https://www.googleapis.com/auth/drive']
STEP = 10000

class SuperSlide:
    def __init__(self, slider, data):
        self.data = data
        self.slider = slider
        self.isNoteHadContent = False
        self.pageId = self.data['objectId']
        self.number = self._parseNumber()
        self.next = None
        self.requestsList = []
        self.thumbnailUrl = None
        self.hash = self._calculateHash()
        self.specialLabel = self._parseSpecialLabel()

    def _parseNumber(self):
        notesShapeId = self.data['slideProperties']['notesPage']['notesProperties']['speakerNotesObjectId']
        notesShapeData = [x for x in self.data['slideProperties']['notesPage']['pageElements'] if x['objectId'] == notesShapeId][0]
        if 'text' in notesShapeData['shape']:
            self.isNoteHadContent = True
            textElementWithText = [x for x in notesShapeData['shape']['text']['textElements'] if 'textRun' in x][0]
            content = textElementWithText['textRun']['content']
            if content.strip().isnumeric():
                return int(content)
            else:
                return None
        else:
            return None

    def _parseSpecialLabel(self):
        layout_id = self.getLayoutId()
        layout_data = [x for x in self.slider.presentation['layouts'] if x['objectId'] == layout_id][0]
        display_name = layout_data['layoutProperties']['displayName']
        m = re.match(r"^.+ – .+ – (\S+)$", display_name)
        if m:
            return m.group(1).lower()
        else:
            return None

    def _calculateHash(self):
        content = yaml.dump(self.data)
        content = re.sub(r'contentUrl:.*', '', content)
        return sha256(content.strip().encode('utf-8')).hexdigest()

    def getNumber(self):
        return self.number

    def getHash(self):
        return self.hash

    def setNext(self, super_slider):
        self.next = super_slider

    def getRequests(self):
        return self.requestsList

    def getLayoutId(self):
        return self.data['slideProperties']['layoutObjectId']

    def getSpecialLabel(self):
        return self.specialLabel

    def renderPNGName(self):
        name = "{:010d}".format(self.getNumber())
        if self.specialLabel:
            name = name + "_" + self.specialLabel
        name = name + ".png"
        return name

    def downloadPNG(self, dst):
        url = 'https://docs.google.com/presentation/d/' + self.slider.presentationId + '/export/png?id=' + self.slider.presentationId + '&pageid=' + self.pageId
        headers = {'Authorization': 'Bearer ' + self.slider.creds.token}
        http = urllib3.PoolManager()
        r = http.request(method='GET',url=url,headers=headers)
        if r.status == 200:
            with open(dst, "wb") as f:
                f.write(r.data)
        else:
            print("Can't download " + dst + ", API HTTP response code is " + str(r.status))

    # recursion
    def enumerate(self, start_point, iteration):
        if self.number and self.next:
            self.next.enumerate(self.number, 0)
        elif self.next:
            next_number = self.next.enumerate(start_point, iteration + 1)
            self.number = int(round(start_point + ((next_number - start_point) / (iteration + 2)) * (iteration + 1)))
        else:
            self.number = start_point + STEP * (iteration + 1)
        return self.number

    def setTransporentBackgroundAsync(self):
        self.requestsList.append(
            {
                'updatePageProperties': {
                    'objectId': self.pageId,
                    'fields': "pageBackgroundFill",
                    'pageProperties': {
                        'pageBackgroundFill': {
                            'solidFill': {
                                'color': {
                                    'rgbColor': {
                                        'red': 0.0,
                                        'green': 0.0,
                                        'blue': 0.0
                                    }
                                },
                                'alpha': 0.0
                            }
                        }
                    }
                }
            }
        )

    def uploadNumberAsync(self):
        if self.isNoteHadContent:
            self.requestsList.append(
                {
                    'deleteText': {
                        'objectId': self.data['slideProperties']['notesPage']['notesProperties'][
                            'speakerNotesObjectId'],
                        'textRange': {
                            'type': 'ALL'
                        }
                    }
                }
            )

        self.requestsList.append(
            {
                'insertText': {
                    'objectId': self.data['slideProperties']['notesPage']['notesProperties']['speakerNotesObjectId'],
                    'insertionIndex': 0,
                    'text': str(self.number)
                }
            }
        )


class Slider:
    def __init__(self, presentation_id, store_path = None, credentials_path = None):
        self.presentationId = presentation_id
        self.creds = None
        # The file token.pickle stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                self.creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    credentials_path, SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle', 'wb') as token:
                pickle.dump(self.creds, token)

        self.service = build('slides', 'v1', credentials=self.creds)
        self.driveService = build('drive', 'v3', credentials=self.creds)

        self.presentation = self.service.presentations().get(presentationId=self.presentationId).execute()
        self.superSlides = []
        self.requests = []
        self.storePath = store_path
        self.storePathNew = None
        self.storePathCurrent = None
        self.statusPath = None
        self.status = {}

        if self.storePath:
            self.statusPath = self.storePath + '/status.yaml'
            self.storePathNew = self.storePath + '/new/'
            self.storePathCurrent = self.storePath + '/current/'

            if not os.path.isdir(self.storePathNew):
                os.makedirs(self.storePathNew)
            if not os.path.isdir(self.storePathCurrent):
                os.makedirs(self.storePathCurrent)

        if self.statusPath:
            if os.path.exists(self.statusPath):
                with open(self.statusPath, 'r') as f:
                    self.status = yaml.safe_load(f)

    def initSuperSlides(self):
        self.superSlides = []
        for i, slide in enumerate(self.presentation['slides']):
            super_slide = SuperSlide(self, slide)
            self.superSlides.append(super_slide)

            if i > 0:
                self.superSlides[i - 1].setNext(self.superSlides[i])

    def enumerateSlides(self):
        self.superSlides[0].enumerate(0, 0)

    def uploadNumbers(self):
        for super_slide in self.superSlides:
            super_slide.uploadNumberAsync()

    def setTransporentBackgrounds(self):
        for super_slide in self.superSlides:
            super_slide.setTransporentBackgroundAsync()

    def wipeLayouts(self):
        layout_ids = []
        for super_slide in self.superSlides:
            layout_ids.append(super_slide.getLayoutId())
        layout_ids = list(set(layout_ids))

        object_ids_to_delete = []
        for layout_id in layout_ids:
            layout = [x for x in self.presentation['layouts'] if x['objectId'] == layout_id][0]
            if 'pageElements' in layout:
                for element in layout['pageElements']:
                    object_ids_to_delete.append(element['objectId'])
        object_ids_to_delete = list(set(object_ids_to_delete))

        for object_id in object_ids_to_delete:
            self.requests.append(
                {
                    'deleteObject': {
                        'objectId': object_id,
                    }
                }
            )

    def updateStatus(self):
        self.status = {}
        for super_slide in self.superSlides:
            self.status[super_slide.pageId] = {'hash': super_slide.getHash(), 'number': super_slide.getNumber()}

    def saveStatus(self):
        if self.statusPath:
            with open(self.statusPath, 'w') as f:
                yaml.dump(self.status, f)

    def downloadFreshPNGs(self):
        for super_slide in self.superSlides:
            if super_slide.pageId in self.status and self.status[super_slide.pageId]['hash'] == super_slide.hash:
                continue

            png_path_new = self.storePathNew + '/' + super_slide.renderPNGName()
            png_path_current = self.storePathCurrent + '/' + super_slide.renderPNGName()
            png_path = ""

            if os.path.exists(png_path_current):
                png_path = png_path_current
            else:
                png_path = png_path_new

            print("  " + png_path)
            super_slide.downloadPNG(png_path)

    def deleteStalePNGs(self):
        actual_numbers = [value['number'] for key, value in self.status.items()]
        for filename in os.listdir(self.storePathCurrent):
            m = re.match(r"([0-9]+).*\.png", filename)
            if m:
                number = int(m.group(1))
                if number not in actual_numbers:
                    os.unlink(self.storePathCurrent + "/" + filename)
                    print("  " + filename)

    def batchUpdateAllRequests(self):
        for super_slide in self.superSlides:
            self.requests = self.requests + super_slide.getRequests()
        # Execute the request.
        body = {
            'requests': self.requests
        }
        return self.service.presentations().batchUpdate(presentationId=self.presentationId, body=body).execute()

    def copyPresentation(self):
        file = self.driveService.files().get(fileId=self.presentationId).execute()
        name = file['name']
        new_name = 'Slider::' + str(uuid.uuid4()) + ' ' + name
        new_file_request_data = {'name': new_name}
        new_presentation = self.driveService.files().copy(
            fileId=self.presentationId, body=new_file_request_data).execute()
        return new_presentation['id']

    def deletePresentation(self):
        self.driveService.files().delete(fileId=self.presentationId).execute()

def main():
    argparser = argparse.ArgumentParser()
    argparser.add_argument("--presentation-id", type=str, action='store', help="Google Presentation ID to download")
    argparser.add_argument("--store-dir", type=str, action='store', help="Dir to store rendered slides (<dir>/new, <dir>/current) and status file (<dir>/status.yaml)")
    argparser.add_argument("--credentials", type=str, action='store', help="path to credentials.json")

    args = argparser.parse_args()

    sliderOriginal = Slider(presentation_id=args.presentation_id, credentials_path=args.credentials)

    print("Initing original presentation: " + args.presentation_id)
    sliderOriginal.initSuperSlides()
    print("Enumerating slides...")
    sliderOriginal.enumerateSlides()
    sliderOriginal.uploadNumbers()
    print("POSTing slide numbers to Google...")
    sliderOriginal.batchUpdateAllRequests()

    print("Creating temporary presentation...")
    tmp_presentation_id = sliderOriginal.copyPresentation()
    print("  " + tmp_presentation_id)
    sliderTmp = Slider(presentation_id=tmp_presentation_id, store_path=args.store_dir, credentials_path=args.credentials)
    print("Initing temporary presentation...")
    sliderTmp.initSuperSlides()
    print("Wiping Dima's face from temporary presentation and setting transporent background...")
    sliderTmp.wipeLayouts()
    sliderTmp.setTransporentBackgrounds()
    sliderTmp.batchUpdateAllRequests()
    print("Downloading fresh PNGs...")
    sliderTmp.downloadFreshPNGs()
    print("Deleting temporary presentation...")
    sliderTmp.deletePresentation()
    print("Updating status registry...")
    sliderTmp.updateStatus()
    print("Deleting stale PNGs...")
    sliderTmp.deleteStalePNGs()
    print("Saving status registry...")
    sliderTmp.saveStatus()

if __name__ == '__main__':
    main()