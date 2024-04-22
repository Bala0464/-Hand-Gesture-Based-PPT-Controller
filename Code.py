import win32com.client
from cvzone.HandTrackingModule import HandDetector
import cv2
import pyautogui
import os
import numpy as np
import math
import aspose.slides as slides
import aspose.pydrawing as drawing

Application = win32com.client.Dispatch("PowerPoint.Application")
Presentation = Application.Presentations.Open(r'"G:\Sample PPT Presentation.pptx"') # paste the ppt address here
print(Presentation.Name)
# Presentation.SlideShowSettings.Run()


width, height = 900, 720
cap = cv2.VideoCapture(0)
cap.set(3, width)
cap.set(4, height)
detectorHand = HandDetector(detectionCon=0.8, maxHands=1)

gestureThreshold = 300
delay = 30
buttonPressed = False
counter = 0
imgNumber = 20
annotations = [[]]
lmList = []
annotationNumber = -1
annotationStart = False
drawingMode = False

slide = Presentation.Slides(1)
slideWidth = slide.Master.Width
slideHeight = slide.Master.Height

if Application.Presentations.Count > 0:
    Presentation = Application.Presentations(1)  # Assuming there is only one presentation open

# Retrieve the active slide
if Presentation.SlideShowWindow.View.State == 1:
    active_slide = Presentation.SlideShowWindow.View.Slide
else:
    active_slide = None


def findDistance(p1, p2, img, draw=True, r=15, t=3):
    x1, y1 = lmList[p1][1:]
    x2, y2 = lmList[p2][1:]
    cx, cy = (x1 + x2) // 2, (y1 + y2) // 2

    if draw:
        cv2.line(img, (x1, y1), (x2, y2), (255, 0, 255), t)
        cv2.circle(img, (x1, y1), r, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (x2, y2), r, (255, 0, 255), cv2.FILLED)
        cv2.circle(img, (cx, cy), r, (0, 0, 255), cv2.FILLED)
    length = math.hypot(x2 - x1, y2 - y1)

    return length, img, (x1, y1, x2, y2, cx, cy)


while True:
    success, img = cap.read()
    hands, img = detectorHand.findHands(img)

    cv2.line(img, (0, gestureThreshold), (width, gestureThreshold), (0, 255, 0), 10)

    if hands and buttonPressed is False:
        hand = hands[0]
        cx, cy = hand["center"]
        lmList = hand["lmList"]
        fingers = detectorHand.fingersUp(hand)

        xVal = int(np.interp(lmList[8][0], [width // 2, width], [0, width]))
        yVal = int(np.interp(lmList[8][1], [150, height - 150], [0, height]))
        indexFinger = xVal, yVal

        if cy <= gestureThreshold and fingers == [1, 1, 1, 1, 1]:
            print("Next")
            buttonPressed = True
            if imgNumber > 0:
                try:
                    if Presentation.SlideShowWindow.View.State == 1:
                        Presentation.SlideShowWindow.View.Next()
                        imgNumber -= 1
                        annotations = [[]]
                        annotationNumber = -1
                        annotationStart = False
                except Exception as e:
                    print(f"Error advancing slide: {e}")

        if cy <= gestureThreshold and fingers == [1, 1, 0, 1, 0]:
            print("Previous")
            buttonPressed = True
            if imgNumber > 0:
                try:
                    if Presentation.SlideShowWindow.View.State == 1:
                        Presentation.SlideShowWindow.View.Previous()
                        imgNumber += 1
                        annotations = [[]]
                        annotationNumber = -1
                        annotationStart = False
                except Exception as e:
                    print(f"Error going to previous slide: {e}")

        if cy <= gestureThreshold and fingers == [1, 1, 1, 0, 0]:
            print("home")
            buttonPressed = True
            if imgNumber > 0:
                try:
                    if Presentation.SlideShowWindow.View.State == 1:
                        Presentation.SlideShowWindow.View.GotoSlide(1)
                        # imgNumber += 1
                        annotations = [[]]
                        annotationNumber = -1
                        annotationStart = False
                except Exception as e:
                    print(f"Error going to previous slide: {e}")

        if cy <= gestureThreshold and fingers == [0, 1, 1, 1, 1]:
            print("Go to End Slide")
            buttonPressed = True
            if imgNumber > 0:
                try:
                    if Presentation.SlideShowWindow.View.State == 1:
                        totalSlides = Presentation.Slides.Count
                        Presentation.SlideShowWindow.View.GotoSlide(totalSlides)
                        annotations = [[]]
                        annotationNumber = -1
                        annotationStart = False
                except Exception as e:
                    print(f"Error going to previous slide: {e}")

        thumb_tip = lmList[4]
        index_tip = lmList[8]
        distance, _, _ = findDistance(4, 8, img, draw=False)

        if cy <= gestureThreshold and fingers == [0, 1, 1, 0, 0]:
            print("Zoom In")
            buttonPressed = True
            try:
                pyautogui.hotkey('ctrl', '+')
            except Exception as e:
                print(f"Error zooming in: {e}")
            # Add your zoom-in logic here

        if cy <= gestureThreshold and fingers == [0, 1, 0, 0, 0]:
            print("Pen Mode")
            if not annotationStart:
                annotationStart = True
                annotationNumber += 1
                annotations.append([])

            if drawingMode:
                annotations[annotationNumber].append(indexFinger)

                if len(annotations[annotationNumber]) > 1:
                    for i in range(len(annotations[annotationNumber]) - 1):
                        p1 = annotations[annotationNumber][i]
                        p2 = annotations[annotationNumber][i + 1]
                        x1, y1 = p1
                        x2, y2 = p2
                        active_slide.Shapes.AddLine(x1, y1, x2, y2).Line.ForeColor.RGB = 255

        if cy <= gestureThreshold and fingers == [0, 1, 1, 1, 0]:
            print("Zoom Out")
            buttonPressed = True
            print("Zoom Out")
            buttonPressed = True
            try:
                pyautogui.hotkey('ctrl', '-')
            except Exception as e:
                print(f"Error zooming out: {e}")

    else:
        annotationStart = False

    if buttonPressed:
        counter += 1
        if counter > delay:
            counter = 0
            buttonPressed = False

    cv2.imshow("Image", img)

    key = cv2.waitKey(1)
    if key == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
