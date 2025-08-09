from tkinter import TOP, LEFT, RIGHT
import customtkinter as ctk
from gui.tk import Tk

ASPECT_RATIO = 9 / 16
RES_X = 720
RES_Y = round(RES_X * ASPECT_RATIO)
RESOLUTION = f"{RES_X}x{RES_Y}"

if __name__ == "__main__":
    root = Tk()
    root.geometry(RESOLUTION)

