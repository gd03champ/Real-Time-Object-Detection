try:
    import cv2
    import numpy as np
    import imutils
    import os
    import time
    import argparse
    import win32com.client as wincl

    parser = argparse.ArgumentParser()
    parser.add_argument('--model_cfg', type = str, default = 'yolov3.cfg',
                        help = 'yolov3.cfg')
    parser.add_argument('--model_weights', type=str,
                        default='yolov3.weights',
                        help='yolov3.weights')
    parser.add_argument('--video', type=str, default='Webcam Stream',
                        help='path to video file')
    parser.add_argument('--src', type=int, default=0,
                        help='source of the camera')
    parser.add_argument('--output_dir', type=str, default='outputs/',
                        help='path to the output directory')
    args = parser.parse_args()
    speak=wincl.Dispatch("Sapi.SpVoice")

    def say(c):
        speak.speak(c)

    # print the arguments
    print('----- info -----')
    print('[i] The config file: ', args.model_cfg)
    print('[i] The weights of model file: ', args.model_weights)
    print('[i] Path to video file: ', args.video)
    print('###########################################################\n')
    say("welcome joel!")
    say("this is a GD industries product")
    
    frameWidth= 850
    frameHeight = 620

    # Load YOLO
    print('[i] Loading network...')
    say("loading the network of memory")
    net = cv2.dnn.readNet(args.model_weights, args.model_cfg)
    classes = []
    with open("coco.names", "r") as f:
        classes = [line.strip() for line in f.readlines()] # we put the names in to an array

    layers_names = net.getLayerNames()
    output_layers = [layers_names[i[0] -1] for i in net.getUnconnectedOutLayers()]
    colors = np.random.uniform(0, 255, size = (len(classes), 3))

    video = cv2.VideoCapture(0)
    #video = cv2.VideoCapture(args.video)
    font = cv2.FONT_HERSHEY_PLAIN
    frame_id = 0
    print("[i] Initiating the process..")
    say("initiating the process..")
    say("come to this black window and press control plus c to terminate the service")
    starting_time = time.time()
    _=0
    try:
        while True:
            _+=1
            has_frame, frame = video.read()
            frame = cv2.resize(frame, (frameWidth, frameHeight), None)


            frame_id += 1
            height, width, channels = frame.shape
            # Detect image
            blob = cv2.dnn.blobFromImage(frame, 0.00392, (320, 320), (0,0,0), swapRB = True, crop = False)
            net.setInput(blob)
            start = time.time()
            outs = net.forward(output_layers)

            # Showing informations on the screen
            class_ids = []
            confidences = []
            boxes = []
            for out in outs:
                for detection in out:
                    scores = detection[5:]
                    class_id = np.argmax(scores)
                    confidence = scores[class_id]
                    if confidence > 0.5:
                        #Object detected
                        center_x = int(detection[0] * width)
                        center_y = int(detection[1] * height)
                        w = int(detection[2] * width)
                        h = int(detection[3] * height)

                        # Rectangle coordinates
                        x = int(center_x - w / 2)
                        y = int(center_y -h / 2)
                        #cv2.rectangle(img, (x,y), (x+w, y+h), (0, 255, 0))

                        boxes.append([x, y, w, h])
                        confidences.append(float(confidence))
                        # Name of the object
                        class_ids.append(class_id)

            indexes = cv2.dnn.NMSBoxes(boxes, confidences, 0.5, 0.3)

            for i in range(len(boxes)):
                if i in indexes:
                    x, y, w, h = boxes[i]
                    label = "{}: {:.2f}%".format(classes[class_ids[i]], confidences[i]*100)
                    color = colors[i]
                    cv2.rectangle(frame, (x,y), (x+w, y+h), color, 2)
                    cv2.putText(frame, label, (x,y+10), font, 1, color, 1)

            elapsed_time = time.time() - starting_time
            fps = frame_id / elapsed_time
            cv2.putText(frame, "FPS:" + str(fps), (10,30), font, 1, (0, 1, 0), 1)
            cv2.putText(frame,'GD Industries', (20,60),font,1,(0,0,0),2)

            
            cv2.imshow("Object Detection", frame)
            key = cv2.waitKey(1)
            if key == 's':
                break

    except KeyboardInterrupt:
        frrr=('[i]',_,'frames has be loaded..')
        print(frrr)
        say(frrr)
        print('[i] Rolling back the initation to avoid blasting of your webcam...')
        say("Rolling back the initation to avoid blasting of your webcam, bye")
        time.sleep(1)

        video.release()
        cv2.destroyAllWindows()

except Exception as e:
    print('I have captured the following bullshit')
    print(e)
    say("I have captured the following shit")
    time.sleep(10)