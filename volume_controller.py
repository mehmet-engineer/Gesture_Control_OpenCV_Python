import cv2
import time
import math
import mediapipe as mp

from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))
#volume.GetMute()
#volume.GetMasterVolumeLevel()
section = volume.GetVolumeRange()  #-63.5'den 0'a kadar
lower_bound = section[0]  # -63.5
upper_bound = section[1]  # 0

def volume_level(length):
    value = (length * 63.5) / 200
    if value > 63.5:
        value = 63.5
    if value < 6:
        value = 0
    value = value - 63.5
    percent = (length * 100) / 260
    if percent > 100:
        percent = 100
    if percent < 5:
        percent = 0
    percent = int(percent)
    return (value,percent)

cam = cv2.VideoCapture(0)

mpHands = mp.solutions.hands
hands = mpHands.Hands(False,1,0.5,0.5)
mpDraw = mp.solutions.drawing_utils

previousTime = 0
currentTime = 0

while True:
    _,img = cam.read()
    imgRGB = cv2.cvtColor(img,cv2.COLOR_BGR2RGB)
    results = hands.process(imgRGB)
    multi = results.multi_hand_landmarks

    if multi is not None:
        for handlms in multi:
            mpDraw.draw_landmarks(img,handlms,mpHands.HAND_CONNECTIONS)
            height,width,channel = img.shape

            x1 = int(handlms.landmark[8].x * width)
            y1 = int(handlms.landmark[8].y * height)

            x2 = int(handlms.landmark[4].x * width)
            y2 = int(handlms.landmark[4].y * height)

            a = x2 - x1
            b = y2 - y1

            length = math.sqrt(a**2 + b**2)
            length = int(length)
            
            if length > 180:
                cv2.putText(img,"length: {}".format(str(length)),(50,90),cv2.FONT_HERSHEY_SIMPLEX,0.8,[255,0,0],2)
                cv2.line(img,(x1,y1),(x2,y2),[255,0,0],3)
                cv2.circle(img,(x1,y1),11,[255,0,0],cv2.FILLED)
                cv2.circle(img,(x2,y2),11,[255,0,0],cv2.FILLED)
            if length > 100 and length < 180:
                cv2.putText(img,"length: {}".format(str(length)),(50,90),cv2.FONT_HERSHEY_SIMPLEX,0.8,[0,127,255],2)
                cv2.line(img,(x1,y1),(x2,y2),[0,127,255],3)
                cv2.circle(img,(x1,y1),11,[0,127,255],cv2.FILLED)
                cv2.circle(img,(x2,y2),11,[0,127,255],cv2.FILLED)
            if length < 100:
                cv2.putText(img,"length: {}".format(str(length)),(50,90),cv2.FONT_HERSHEY_SIMPLEX,0.8,[10,10,255],2)
                cv2.line(img,(x1,y1),(x2,y2),[10,10,255],3)
                cv2.circle(img,(x1,y1),11,[10,10,255],cv2.FILLED)
                cv2.circle(img,(x2,y2),11,[10,10,255],cv2.FILLED)

            value,percent = volume_level(length)
            volume.SetMasterVolumeLevel(value,None)
            cv2.putText(img,"volume: %{}".format(str(percent)),(50,130),cv2.FONT_HERSHEY_SIMPLEX,0.8,[250,10,10],2)

    currentTime = time.time()
    fps = 1 / (currentTime - previousTime)
    previousTime = currentTime

    fps = int(fps)
    if fps > 30:
        cv2.putText(img,"FPS: {}".format(str(fps)),(50,50),cv2.FONT_HERSHEY_SIMPLEX,0.8,[255,0,0],2)
    else:
        cv2.putText(img,"FPS: {}".format(str(fps)),(50,50),cv2.FONT_HERSHEY_SIMPLEX,0.8,[0,0,255],2)
    
    cv2.imshow("img",img)

    if cv2.waitKey(1) == 27:
        break

cam.release()
cv2.destroyAllWindows()