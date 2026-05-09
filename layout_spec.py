"""Shared physical layout specs for badge rendering and print validation."""

from __future__ import annotations

from dataclasses import dataclass


A4_WIDTH_MM = 210.0
A4_HEIGHT_MM = 297.0

CARD_WIDTH_MM = 96.0
CARD_HEIGHT_MM = 55.0
PHOTO_WIDTH_MM = 25.0
PHOTO_HEIGHT_MM = 35.0

COLS = 2
ROWS = 5
CARDS_PER_PAGE = COLS * ROWS

PAGE_MARGIN_LEFT_MM = 7.0
PAGE_MARGIN_RIGHT_MM = 7.0
PAGE_MARGIN_TOP_MM = 5.0
PAGE_MARGIN_BOTTOM_MM = 5.0
CARD_GAP_X_MM = 4.0
CARD_GAP_Y_MM = 2.0


@dataclass(frozen=True)
class LayoutValidation:
    card_width_mm: float
    card_height_mm: float
    photo_width_mm: float
    photo_height_mm: float
    sheet_width_mm: float
    sheet_height_mm: float
    printable_width_mm: float
    printable_height_mm: float
    fits_a4: bool


def page_content_width_mm() -> float:
    return COLS * CARD_WIDTH_MM + (COLS - 1) * CARD_GAP_X_MM


def page_content_height_mm() -> float:
    return ROWS * CARD_HEIGHT_MM + (ROWS - 1) * CARD_GAP_Y_MM


def validate_layout() -> LayoutValidation:
    sheet_w = page_content_width_mm()
    sheet_h = page_content_height_mm()
    printable_w = A4_WIDTH_MM - PAGE_MARGIN_LEFT_MM - PAGE_MARGIN_RIGHT_MM
    printable_h = A4_HEIGHT_MM - PAGE_MARGIN_TOP_MM - PAGE_MARGIN_BOTTOM_MM
    return LayoutValidation(
        card_width_mm=CARD_WIDTH_MM,
        card_height_mm=CARD_HEIGHT_MM,
        photo_width_mm=PHOTO_WIDTH_MM,
        photo_height_mm=PHOTO_HEIGHT_MM,
        sheet_width_mm=sheet_w,
        sheet_height_mm=sheet_h,
        printable_width_mm=printable_w,
        printable_height_mm=printable_h,
        fits_a4=sheet_w <= printable_w and sheet_h <= printable_h,
    )
