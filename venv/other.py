import cv2
names = (1007, 1120, 1879, 2275, 2327, 3034, 3052, 3055, 3068, 3087, 3172, 3240, 3280)
for name in names:
    # Load image, convert to grayscale, and find edges
    image = cv2.imread(rf'C:\Users\ulito\Desktop\training\patterns\ex_d\schemes\\{name}.jpg')
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    thresh = cv2.threshold(gray, 10, 255, cv2.THRESH_OTSU + cv2.THRESH_BINARY_INV)[1]

    # Find contour and sort by contour area
    cnts = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    cnts = sorted(cnts, key=cv2.contourArea, reverse=True)
    x_=[]
    y_=[]
    # Find bounding box and extract ROI
    for c in cnts:
        x,y,w,h = cv2.boundingRect(c)
        x_.append(x)
        y_.append(y)
    min_x = min(x_)
    min_y = min(y_)
    max_x = max(x_)
    max_y = max(y_)
    ROI = image[min_y-50:max_y+50, min_x-50:max_x+50]

    cv2.imwrite(fr'C:\Users\ulito\Desktop\training\patterns\ex_d\schemes\\{name}_1.jpg',ROI)

