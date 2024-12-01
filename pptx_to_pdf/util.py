from math import log10, ceil


def emu_to_pixels(emu_value):
    return None if not emu_value else int(emu_value * 1 / 9525)


def emu_to_pt(emu_value):
    return None if not emu_value else int(emu_value * 1 / 12700)

def pt_to_emu(pt_value):
    return None if not pt_value else int(pt_value * 12700)


def create_slide_number_string(slide_number: int, max_slide_number: int = 9999):
    if slide_number < 0:
        raise ValueError("slide_number must be positive")
    if slide_number > max_slide_number:
        raise ValueError("slide_number exceeds max_slide_number")

    max_num_digits = len(str(max_slide_number))
    current_num_digits = len(str(slide_number))

    zeros_prefix = (max_num_digits - current_num_digits) * "0"

    return zeros_prefix + str(slide_number)
