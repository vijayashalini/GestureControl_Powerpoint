import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import numpy as np
Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open("C:/Users/vignesh/OneDrive/Documents/EmotionalIntelligence-es.pptx")
print(Presentation.Name)
Presentation.SlideShowSettings.Run()

width, height = 360, 360
gestureThreshold = 300

cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)

detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

imgList = []
delay = 30
buttonPressed = False
counter = 0
imgNumber = 20
annotations = [[]]
annotationNumber = -1
annotationStart = False
while True:

    success, img = cap.read()

    hands, img = detectorHand.findHands(img)
    if hands and buttonPressed is False:
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]
        fingers = detectorHand.fingersUp(hand)
        xVal = int(np.interp(lmList[8][0], [1280 // 2, 1280], [0, 1280]))
        yVal = int(np.interp(lmList[8][1], [150, 720-150], [0, 720]))
        indexFinger = xVal, yVal

        if cy <= gestureThreshold:
            if fingers == [1, 1, 1, 1, 1]:
                #Next
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Next()
                    imgNumber -= 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            if fingers == [1, 1, 0, 0, 0]:
                #Previous
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.Previous()
                    imgNumber += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            if fingers == [1, 0, 0, 0, 0]:
                #Home
                buttonPressed = True
                if imgNumber > 0:
                    Presentation.SlideShowWindow.View.First()
                    imgNumber += 1
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            if fingers == [0, 0, 0, 0, 1]:
                #End
                buttonPressed = True
                if imgNumber >= 0:
                    Presentation.SlideShowWindow.View.Last()
                    annotations = [[]]
                    annotationNumber = -1
                    annotationStart = False
            if fingers == [0, 1, 0, 0, 0]:
                slide = Presentation.SlideShowWindow.View.Slide
                shape = slide.Shapes.AddShape(9, int(indexFinger[0]), int(indexFinger[1]), 24, 24)
                shape.Fill.ForeColor.RGB = 255
                delay=60
                annotationNumber += 1
                shape.Line.Visible = False

            if fingers == [0, 1, 1, 1, 0]:
                if annotations:
                    annotations.pop(-1)
                    annotationNumber -= 1
                    buttonPressed = True
                    slide = Presentation.SlideShowWindow.View.Slide
                    for shape in slide.Shapes:
                        if shape.Fill.ForeColor.RGB == 255:
                            shape.Delete()
            if fingers == [0, 1, 1, 0, 0]:
                # Start playing the video on the current slide
                Presentation.SlideShowWindow.View.GotoClick(1)

    else:
        annotationStart = False

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    for i, annotation in enumerate(annotations):
        for j in range(len(annotation)):
            if j != 0:
                cv2.line(img, annotation[j - 1], annotation[j], (0, 0, 200), 12)

    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break