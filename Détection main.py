# Imports
import numpy as np
import cv2

import math
import pyautogui #permet d'avoir des outputs clavier

import win32com.client

# Open Camera
capture = cv2.VideoCapture(0)

# Initialization of the voice
voix=0#voix désactiver par défaut
compteur=0#délais avant la voix
bBackground=0#booléen
bkgrnd = 0#arrière plan par défaut
speaker = win32com.client.Dispatch("SAPI.SpVoice")



#Traitement capture
while capture.isOpened():

    # Capture frames from the camera
    ret, frame = capture.read()

    # Get hand data from the rectangle sub window   
    cv2.rectangle(frame,(100,100),(300,300),(0,255,0),0)
    crop_image = frame[100:300, 100:300]
    
    #Prétraitement de l'image
    diff = cv2.absdiff(crop_image, bkgrnd)#marche très bien sur le fond est sombre
    if bBackground==1:
        _, diff = cv2.threshold(diff, 25, 255, cv2.THRESH_BINARY)
        diff = cv2.GaussianBlur(diff, (3,3), 5)
        bBackground=0

    # HSV values
    low_range = np.array([0, 50, 80])
    upper_range = np.array([30, 200, 255])

    # Change color-space from BGR -> HSV
    hsv = cv2.cvtColor(diff, cv2.COLOR_BGR2HSV) #projection sur une autre base
    
    # Create a binary image with where white will be skin colors and rest is black
    mask = cv2.inRange(hsv, low_range, upper_range)

    # Kernel for morphological transformation    
    kernel = np.ones((5,5))

    # Apply morphological transformations to filter out the background noise #On fait une fermeture
    dilation = cv2.dilate(mask, kernel, iterations = 1)
    erosion = cv2.erode(dilation, kernel, iterations = 1)    

    # Apply Gaussian Blur and Threshold #Filtrage pour avoir un meilleur résultat
    filtered = cv2.GaussianBlur(erosion, (3,3), 0)
    ret,thresh = cv2.threshold(filtered, 127, 255, 0)

    # Find contours # Traitement pour la main
    # check OpenCV version
    major = cv2.__version__.split('.')[0]
    if major == '3':
        image, contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE )
    else:
        contours, hierarchy = cv2.findContours(thresh, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE )

    try:
        # Find contour with maximum area
        contour = max(contours, key = lambda x: cv2.contourArea(x))

        # Create bounding rectangle around the contour
        x,y,w,h = cv2.boundingRect(contour)
        cv2.rectangle(crop_image,(x,y),(x+w,y+h),(0,0,255),0)

        # Find convex hull
        hull = cv2.convexHull(contour)

        # Draw contour
        drawing = np.zeros(crop_image.shape,np.uint8)
        cv2.drawContours(drawing,[contour],-1,(0,255,0),0)
        cv2.drawContours(drawing,[hull],-1,(0,0,255),0)

        # Fi convexity defects
        hull = cv2.convexHull(contour, returnPoints=False)
        defects = cv2.convexityDefects(contour,hull)

        # Use cosine rule to find angle of the far point from the start and end point i.e. the convex points (the finger 
        # tips) for all defects
        count_defects = 0

        for i in range(defects.shape[0]):
            s,e,f,d = defects[i,0]
            start = tuple(contour[s][0])
            end = tuple(contour[e][0])
            far = tuple(contour[f][0])

            a = math.sqrt((end[0] - start[0])**2 + (end[1] - start[1])**2)
            b = math.sqrt((far[0] - start[0])**2 + (far[1] - start[1])**2)
            c = math.sqrt((end[0] - far[0])**2 + (end[1] - far[1])**2)
            angle = (math.acos((b**2 + c**2 - a**2)/(2*b*c))*180)/3.14

            # if angle >= 90 draw a circle at the far point
            if angle <= 90:
                count_defects += 1
                cv2.circle(crop_image,far,1,[0,0,255],-1)

            cv2.line(crop_image,start,end,[0,255,0],2)


        #Condition sur les doigts
        font = cv2.FONT_HERSHEY_SIMPLEX
        fontScale = 2
        thickness = 2
        linetype = cv2.LINE_AA
        
        if count_defects <0:#Cas si erreur
            cv2.putText(frame,"Close", (50,45), font, fontScale, (0,0,0), thickness,linetype)
        elif count_defects == 0:#Attention pas assez précis
            cv2.putText(frame,"Un", (50,45), font, fontScale,(0,0,0), thickness, linetype)     
        elif count_defects == 1:#Attention pas assez précis
            pyautogui.press('space')#A chaque détection appuye sur la touche espace permettrait d'automatiser des taches
            cv2.putText(frame,"Deux", (50,45), font, fontScale,(0,0,0), thickness, linetype) 
        elif count_defects == 2:
            cv2.putText(frame, "Trois", (50,45), font, fontScale,(255,0,0), thickness, linetype)

        elif count_defects == 3:
            cv2.putText(frame,"Quatre", (50,45), font, fontScale,(0,255,0), thickness, linetype)
            if voix>3:
                compteur=compteur+1
                if compteur>10:
                    speaker.Speak("Quatre")
                    voix=0
                    compteur=0
        elif count_defects == 4:
            cv2.putText(frame,"Cinq", (50,45), font, fontScale,(0,0,255), thickness, linetype) 
            compteur = compteur + 1
            if compteur> 10:
                voix=voix+1
                compteur=0
                
        elif count_defects >= 4:#cas si erreur
            cv2.putText(frame,"Ouvert", (50,45), font, fontScale,(0,0,255), thickness, 2)
            
    except:#interruption
        pass
    # Show required images
    cv2.imshow("Gesture", frame)
    
    # Close the camera if 'q' is pressed
    if cv2.waitKey(1) == ord('q'):
        break     

    elif cv2.waitKey(1) == ord('b'):#Fonctionnalité non pertinante en interruption, erreur
         bBackground=1
         bkgrnd = crop_image
         print('Background soustrait')
         
    # elif cv2.waitKey(1) == ord('d'):
    #      bBackground=0
    #      bkgrnd = 0
    #      print('Background par défaut')
 
capture.release()
cv2.destroyAllWindows()