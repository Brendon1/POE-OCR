import Levenshtein as lv
import numpy as np
import cv2 as cv
import pytesseract
import win32api
import win32con
import time
from playsound import playsound
from mss import mss

def getSearchStrings(fileName = "config/search_strings.txt"):
    """
    Imports the search strings configuration file.

    The search string configuration file will dictate which phrases the
    OCR application will look for.
    """
    file = open(fileName, 'r')
    lines = file.read().splitlines()
    return lines

def getHSVBounds(mode = "equipment"):
    """
    Gives the upper and lower HSV bounds for the mode specified.

    The HSV bounds are used in isolating features in an image based off
    of their color. If no mode is specified the upper and lower ranges
    will include all color values.
    """
    lower = None
    upper = None

    match mode:
        case "sextants":
            lower = np.array([102,57,180])
            upper = np.array([103,61,255])
        case "equipment":
            lower = np.array([120,119,165])
            upper = np.array([121,123,255])
        case _:
            lower = np.array([0,0,0])
            upper = np.array(255,255,255)

    return lower, upper

def getScreenshot(monitor = 1):
    """
    Captures a screenshot of the specified monitor.

    Following the implementation details defined in the 'mss' library
    monitor indexing is as follows:
        0 = all monitors
        1 = primary monitor
        2 = secondary monitor
        etc.
    """
    with mss() as sct:
        screenshot = np.array(sct.grab(sct.monitors[monitor]))
        return screenshot
    
def getIsolatedText(img):
    """
    Returns a black and white version of the given image where the text
    has been isolated.

    Text in the starting image is isolated using HSV bounds determined
    by the specified mode. The desired text will be in black and everything
    else will be white for better OCR performance.
    """
    lower, upper = getHSVBounds()
    hsv = cv.cvtColor(img, cv.COLOR_BGR2HSV)
    mask = cv.inRange(hsv, lower, upper)
    return cv.bitwise_not(mask)

def getPhraseROIs(img):
    """
    Returns a list of black and white images containing the phrases found in
    the given image that match the given mode.
    """
    mask = getIsolatedText(img)
    blurredImg = cv.GaussianBlur(mask, (7,7), 0)
    thresh = cv.threshold(blurredImg, 0, 255, cv.THRESH_BINARY_INV + cv.THRESH_OTSU)[1]
    kernel = cv.getStructuringElement(cv.MORPH_RECT, (12, 4))
    dilate = cv.dilate(thresh, kernel, iterations=2)
    cnts = cv.findContours(dilate, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
    cnts = cnts[0] if len(cnts) == 2 else cnts[1]
    cnts = sorted(cnts, key= lambda x: cv.boundingRect(x)[0])

    # Filters out the contours that are too small to likely be of use
    rois = []
    for c in cnts:
        x, y, w, h = cv.boundingRect(c)
        if h > 20 and w > 200:
            rois.append(mask[y:y+h, x:x+w])

    return rois

def getPhraseStrings(img):
    """
    Returns a list of strings containing all the phrases found in the
    given image.

    Which phrases are found will be influenced by the current mode.
    """
    phrases = []
    for roi in getPhraseROIs(img):
        phrases.append(pytesseract.image_to_string(roi))

    return phrases

def hasMatchingPhrase(phrases, searchStrings):
    """
    Checks if any of the phrases match any of the search strings.

    A match is determined by fuzzy string matching above a score
    cutoff rather than by an exact match.
    """
    # TODO: may want to expose the score cutoff for easier tuning
    for phrase in phrases:
        for string in searchStrings:
            if lv.ratio(phrase.lower(), string, score_cutoff=0.90):
                return True
    return False

def runOCR(searchStrings):
    img = getScreenshot(2)
    phrases = getPhraseStrings(img)

    if hasMatchingPhrase(phrases, searchStrings):
        playsound('quiet_alert.mp3', block=False)


filename = "config/search_strings.txt"
searchStrings = getSearchStrings(filename)

while True:
    exitState = win32api.GetAsyncKeyState(win32con.VK_BACK) & 0x8000

    if exitState:
        break

    LShiftState = win32api.GetAsyncKeyState(win32con.VK_LSHIFT) & 0x8000
    LMouseState = win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000

    if LShiftState and LMouseState:
        time.sleep(0.25)
        runOCR(searchStrings)
    else:
        time.sleep(0.02)