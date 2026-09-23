"""Готовий рис · Україна — глибокий розбір ринку по кожному SKU.

Кожен бренд отримує власний розворот: усі його позиції з пакшотом, масою,
ціною за упаковку та ціною за 100 грамів. Без наших позицій.

Фірмовий стиль Morskyi Dim: #303A5D / #F9A50B, хвильовий патерн і логотип.
"""
import statistics as st

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from PIL import Image

W, H = 13.3333, 7.5
M = 0.66
HDR = 1.16

NAVY   = RGBColor(0x30, 0x3A, 0x5D)
NAVY_L = RGBColor(0x46, 0x53, 0x82)
ORANGE = RGBColor(0xF9, 0xA5, 0x0B)
AMBER  = RGBColor(0xFF, 0xC9, 0x5C)
TEAL   = RGBColor(0x1E, 0x9E, 0xA6)
GREEN  = RGBColor(0x37, 0xA1, 0x69)
PLUM   = RGBColor(0x7B, 0x5E, 0xA7)
ROSE   = RGBColor(0xC9, 0x4F, 0x7C)
SLATE  = RGBColor(0x5B, 0x6B, 0x8C)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
MIST   = RGBColor(0xF1, 0xF4, 0xFA)
MIST_D = RGBColor(0xE2, 0xE8, 0xF3)
INK    = RGBColor(0x1C, 0x22, 0x35)
GREY   = RGBColor(0x71, 0x78, 0x90)

FONT = "Segoe UI"
BG_TITLE = "brand/md_bg_title.jpg"
LOGO = "brand/md_logo_white.png"

prs = Presentation()
prs.slide_width, prs.slide_height = Inches(W), Inches(H)
BLANK = prs.slide_layouts[6]
PG = {"n": 0}

