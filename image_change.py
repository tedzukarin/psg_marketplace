import numpy as np
import cv2



def image_show(image_file, name, art):
    image = np.asarray(bytearray(image_file))
    img = cv2.imdecode(image, 0)
    height, width = img.shape[:2]

    new_height = height + 90
    new_width = width
    new_image = np.ones((new_height, new_width), np.uint8) * 255
    y_offset = 65
    x_offset = 0
    new_image[y_offset:y_offset + height, x_offset:x_offset + width] = img

    n = 20
    if name:
        name = name + (' ' * (len(name) % 26))
    else:
        name = ''
    for i in range(0, len(name), 26):
        font = cv2.FONT_HERSHEY_COMPLEX
        bottomLeftCornerOfText = (10, n)
        fontScale = 0.7
        fontColor = (0, 0, 0)
        thickness = 2
        lineType = 1

        cv2.putText(new_image, name[i: i+26],
                    bottomLeftCornerOfText,
                    font,
                    fontScale,
                    fontColor,
                    thickness,
                    lineType)

        n += 20
    cv2.putText(new_image, art, org=(10, 385), fontFace=cv2.FONT_HERSHEY_COMPLEX, fontScale=0.7, color=(0,0,0), thickness=2)
    return new_image

