import os

from PptToDocx import pptToDocx

OUTPUT_DIR = './output/'
INPUT_DIR = 'E:/Books/Resources/CSE422/CAT 1/'

for inputFile in os.listdir(INPUT_DIR):
    if inputFile[-5:] == '.pptx':
        pptToDocx(INPUT_DIR+inputFile,OUTPUT_DIR)
